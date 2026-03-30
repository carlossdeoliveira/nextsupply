import re
import sys
import zipfile
import tempfile
from pathlib import Path

import pandas as pd
import pdfplumber


# ==============================
# UTILITÁRIOS
# ==============================

def norm(text):
    text = (text or "").replace("\u00ad", "")
    text = re.sub(r"\bPág:\s*\d+/\d+\b", "", text, flags=re.I)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\s*\n\s*", "\n", text)
    return text.strip()


def oneline(text):
    text = norm(text).replace("\n", " ")
    text = re.sub(r"\(NEXTSUPPLY\d+\)", "", text, flags=re.I)
    text = re.sub(r"\s{2,}", " ", text).strip()
    return text


def patt(t):
    parts = re.split(r"\s+", str(t).strip().lower())
    return r"\b" + r"[\s\-]?".join(map(re.escape, parts)) + r"\b"


# ==============================
# EXTRAÇÃO
# ==============================

def extract_header(raw, fallback=""):
    m = re.search(r"Número da Oportunidade.*?(\d{10})", raw, flags=re.S)
    if m:
        numero = m.group(1)
    else:
        m = re.search(r"(\d{10})", str(fallback))
        numero = m.group(1) if m else ""

    tipo = re.search(r"Tipo de Oportunidade\s*([^\n]+)", raw)
    tipo = tipo.group(1).strip() if tipo else ""

    crit = re.search(r"Critério de Julgamento\s*([^\n]+)", raw)
    crit = crit.group(1).strip() if crit else ""

    fim = re.search(
        r"Fim do período de cotação\s*([0-9]{2}\.[0-9]{2}\.[0-9]{4}\s*/\s*[0-9]{2}:[0-9]{2}:[0-9]{2})",
        raw,
    )
    fim = re.sub(r"\s*/\s*", " / ", fim.group(1)).strip() if fim else ""

    ml = re.search(r"Local de Entrega\s*(.*?)\nInformações do Comprador", raw, flags=re.S)
    local = "; ".join(x.strip() for x in ml.group(1).split("\n") if x.strip()) if ml else ""

    return numero, tipo, crit, fim, local


def extract_fab(longd):
    if not longd:
        return ""
    m = re.search(r"(?is)(?:Tp:|Type:|----------)\s*(.*?)(?=Tag|Dados|Número|Declara|$)", longd)
    return re.sub(r"\s{2,}", " ", m.group(1)).strip() if m else ""


# ==============================
# RESPONSÁVEL
# ==============================

def assign(df):
    if df.empty:
        df["Responsável"] = pd.Series(dtype="object")
        return df

    text = (
        df["Descrição de Item"].fillna("")
        + " "
        + df["Descrição longa do item"].fillna("")
        + " "
        + df["Fabricante/PN"].fillna("")
    ).str.lower()

    df["Responsável"] = pd.NA

    # PRIORIDADE: Inaplica
    df.loc[
        df["Tipo de Oportunidade"].fillna("").str.contains("Inaplica.", na=False, regex=False),
        "Responsável",
    ] = "Viviana"

    # HELIO (TRAVADO)
    helio_terms = ["abb", "schneider", "siemens", "rittal", "phoenix", "weidmuller", "rockwell"]
    for t in helio_terms:
        df.loc[
            df["Responsável"].isna() & text.str.contains(patt(t), regex=True, na=False),
            "Responsável",
        ] = "Hélio"

    # MAYARA (TRAVADO)
    for t in ["skf", "emerson"]:
        df.loc[
            df["Responsável"].isna() & text.str.contains(patt(t), regex=True, na=False),
            "Responsável",
        ] = "Mayara"

    # VIVIANA
    for t in ["kongsberg", "yamada", "dnh", "evac", "steyr"]:
        df.loc[
            df["Responsável"].isna() & text.str.contains(patt(t), regex=True, na=False),
            "Responsável",
        ] = "Viviana"

    # GARANTIA: ITENS IGUAIS JUNTOS
    df["_chave"] = (
        df["Descrição de Item"].fillna("").str.strip().str.lower() + "|"
        + df["Descrição longa do item"].fillna("").str.strip().str.lower() + "|"
        + df["Fabricante/PN"].fillna("").str.strip().str.lower()
    )

    total = len(df)
    targets = {
        "Viviana": round(total * 0.33),
        "Hélio": round(total * 0.33),
    }
    targets["Mayara"] = total - sum(targets.values())

    counts = df["Responsável"].value_counts(dropna=True).to_dict()
    for k in targets:
        counts.setdefault(k, 0)

    pendentes = df[df["Responsável"].isna()]

    for _, grp in pendentes.groupby("_chave", sort=False):
        escolhido = max(targets, key=lambda k: targets[k] - counts[k])
        idxs = grp.index.tolist()
        df.loc[idxs, "Responsável"] = escolhido
        counts[escolhido] += len(idxs)

    df.drop(columns=["_chave"], inplace=True)

    return df


# ==============================
# PROCESSAMENTO
# ==============================

def process(zip_path, output):
    zip_path = Path(zip_path)
    output = Path(output)

    if not zip_path.exists():
        raise FileNotFoundError(f"Arquivo ZIP não encontrado: {zip_path}")

    expected_columns = [
        "Numero da Oportunidade",
        "Tipo de Oportunidade",
        "Critério de Julgamento",
        "Fim do período de cotação",
        "Local de Entrega",
        "Item",
        "Quantidade",
        "Unidade de medida",
        "Descrição de Item",
        "Descrição longa do item",
        "Fabricante/PN",
    ]

    rows = []

    with tempfile.TemporaryDirectory() as temp_dir:
        work = Path(temp_dir) / "temp_ops"
        work.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(work)

        for pdf in work.rglob("*.pdf"):
            with pdfplumber.open(str(pdf)) as p:
                raw = norm("\n".join((pg.extract_text() or "") for pg in p.pages))

            num, tipo, crit, fim, local = extract_header(raw, pdf.stem)

            for b in re.split(r"(?i)Dados do Item", raw)[1:]:
                m = re.search(r"\n\s*(\d+)\s", b)
                if not m:
                    continue

                iid = m.group(1)

                desc_match = re.search(r"Descrição de Item\s*(.*?)(?=Descrição longa|$)", b, re.S)
                desc = oneline(desc_match.group(1)) if desc_match else ""

                long_match = re.search(r"Descrição longa.*?(.*)$", b, re.S)
                longd = oneline(long_match.group(1)) if long_match else ""

                rows.append(
                    {
                        "Numero da Oportunidade": num,
                        "Tipo de Oportunidade": tipo,
                        "Critério de Julgamento": crit,
                        "Fim do período de cotação": fim,
                        "Local de Entrega": local,
                        "Item": iid,
                        "Quantidade": "",
                        "Unidade de medida": "",
                        "Descrição de Item": desc,
                        "Descrição longa do item": longd,
                        "Fabricante/PN": extract_fab(longd),
                    }
                )

    if rows:
        df = pd.DataFrame(rows)
        df = df.drop_duplicates(subset=["Numero da Oportunidade", "Item"], keep="first")
    else:
        df = pd.DataFrame(columns=expected_columns)

    df = assign(df)

    with pd.ExcelWriter(output, engine="openpyxl") as w:
        if not df.empty:
            for r in ["Viviana", "Hélio", "Mayara"]:
                df[df["Responsável"] == r].to_excel(w, sheet_name=r, index=False)
            df.to_excel(w, sheet_name="Consolidado", index=False)

            resumo = pd.DataFrame({"Responsável": ["Viviana", "Hélio", "Mayara"]})
            resumo["Atual"] = resumo["Responsável"].map(lambda x: int((df["Responsável"] == x).sum()))
            resumo["%"] = (resumo["Atual"] / len(df) * 100).round(1)
            resumo.to_excel(w, sheet_name="Resumo", index=False)
        else:
            pd.DataFrame({"Aviso": ["Nenhum item encontrado nos PDFs"]}).to_excel(
                w,
                sheet_name="Consolidado",
                index=False,
            )

    return output


# ==============================
# EXECUÇÃO
# ==============================

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python script.py arquivo.zip [arquivo_saida.xlsx]")
        sys.exit(1)

    zip_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else "saida.xlsx"

    try:
        result = process(zip_file, output_file)
        print(f"Arquivo gerado com sucesso: {result}")
    except Exception as e:
        print(f"Erro ao processar: {e}")
        sys.exit(1)
