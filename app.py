import re
import zipfile
import tempfile
from pathlib import Path

import pandas as pd
import pdfplumber
import streamlit as st

# ==============================
# UTILITÁRIOS
# ==============================

def norm(text: str) -> str:
    text = (text or "").replace("\u00ad", "")
    text = re.sub(r"\bPág:\s*\d+/\d+\b", "", text, flags=re.I)
    text = re.sub(r"(?i)Resumo extra[ií]do por.*?(?=\n|$)", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\s*\n\s*", "\n", text)
    return text.strip()

def oneline(text: str) -> str:
    text = norm(text).replace("\n", " ")
    text = re.sub(r"\(NEXTSUPPLY\d+\)", "", text, flags=re.I)
    text = re.sub(r"\s{2,}", " ", text).strip()
    return text

def make_term_pattern(term: str) -> str:
    parts = re.split(r"\s+", term.lower().strip())
    return r"\b" + r"[\s\-]?".join(map(re.escape, parts)) + r"\b"

# ==============================
# EXTRAÇÃO
# ==============================

def extract_header(raw: str, fallback: str = ""):
    numero = re.search(r"\d{10}", raw)
    numero = numero.group(0) if numero else ""

    tipo = re.search(r"Tipo de Oportunidade\s*([^\n]+)", raw)
    tipo = tipo.group(1).strip() if tipo else ""

    criterio = re.search(r"Critério de Julgamento\s*([^\n]+)", raw)
    criterio = criterio.group(1).strip() if criterio else ""

    fim = re.search(r"([0-9]{2}\.[0-9]{2}\.[0-9]{4}\s*/\s*[0-9]{2}:[0-9]{2}:[0-9]{2})", raw)
    fim = fim.group(1) if fim else ""

    ml = re.search(r"Local de Entrega\s*(.*?)\nInformações do Comprador", raw, flags=re.S)
    local = "; ".join([x.strip() for x in ml.group(1).split("\n") if x.strip()]) if ml else ""

    return numero, tipo, criterio, fim, local

def extract_item_id(block: str) -> str:
    m = re.search(r"\n\s*(\d+)\s", block)
    return m.group(1) if m else ""

def extract_qty_unit(block: str, item_id: str):
    m = re.search(rf"{item_id}.*?Material\s+([\d\.,]+)\s+([A-Za-z]+)", block, re.S)
    return (m.group(1), m.group(2)) if m else ("", "")

def extract_desc(block: str) -> str:
    m = re.search(r"Descrição de Item\s*(.*?)(?=Descrição longa|$)", block, re.S)
    return oneline(m.group(1)) if m else ""

def extract_long(block: str) -> str:
    block2 = re.split(
        r"(?i)Declarações envolvidas na oportunidade|Declaracoes envolvidas na oportunidade",
        block,
        maxsplit=1,
    )[0]

    m = re.search(
        r"(?is)(?:Descrição longa.*?)(.*)$",
        block2,
    )
    if not m:
        return ""

    text = m.group(1)

    text = re.split(
        r"(?i)\b(?:em\s+)?\d{2}\.\d{2}\.\d{4}\s+às\s+\d{2}:\d{2}\b",
        text,
        maxsplit=1,
    )[0]

    return oneline(text)

def extract_fab(longd: str) -> str:
    if not longd:
        return ""

    m = re.search(
        r"(?is)(?:Tp:|Type:|----------)\s*(.*?)(?=Tag|Dados|Número|Declara|\b(?:em\s+)?\d{2}\.\d{2}\.\d{4}\s+às\s+\d{2}:\d{2}\b|$)",
        longd,
    )
    return re.sub(r"\s{2,}", " ", m.group(1)).strip() if m else ""

# ==============================
# RESPONSÁVEL
# ==============================

# Mapeamento de fabricantes por responsável
FABRICANTES = {
    "Hélio": [
        "siemens", "schneider", "phoenix", "abb", "weidmuller", "conexel",
        "wago", "lovato", "rockwell", "waukesha", "turck", "yongsung",
        "denora", "3x engineering", "topsafe", "mcfarland", "pepperl",
    ],
    "Mayara": [
        "evac", "steyr", "sailor", "emerson", "damcos", "warom",
        "liebherr", "freudenberg", "siebenhaar", "bukh",
    ],
    "Viviana": [
        "danfoss", "orga", "dnh", "hirschmann", "vega", "kimray",
        "zenitel", "grundfos", "george fischer", "george fis", "arca",
        "amot", "flebu", "skf", "parker", "esbee", "aplisens",
        "bitzer", "hilti", "oglaned",
    ],
    "Gabriele": [
        "eaton", "cooper", "fhf", "medc", "bussmann", "deno",
        "westlock", "kidde", "caterpillar", "exheat", "stahl",
        "seatrax", "schweitzer",
    ],
}

# Percentuais alvo de distribuição
PERCENTUAIS = {
    "Hélio":    0.3333,
    "Viviana":  0.3333,
    "Mayara":   0.20,
    "Gabriele": 0.1333,
}

def assign(df):
    text = (
        df["Descrição de Item"].fillna("") +
        " " +
        df["Descrição longa do item"].fillna("") +
        " " +
        df["Fabricante/PN"].fillna("")
    ).str.lower()

    df["Responsável"] = None

    # Inaplica → sempre Viviana
    df.loc[
        df["Tipo de Oportunidade"].str.contains("Inaplica.", na=False),
        "Responsável"
    ] = "Viviana"

    # Atribuição por fabricante, na ordem de prioridade declarada
    for responsavel, fabricantes in FABRICANTES.items():
        for fab in fabricantes:
            pattern = make_term_pattern(fab)
            mask = df["Responsável"].isna() & text.str.contains(pattern, regex=True)
            df.loc[mask, "Responsável"] = responsavel

    # Distribuição proporcional para os itens sem responsável
    total = len(df)
    targets = {k: round(total * v) for k, v in PERCENTUAIS.items()}

    # Ajuste para garantir que a soma bata exatamente com o total
    diff = total - sum(targets.values())
    if diff != 0:
        # Aplica o ajuste no responsável com maior percentual (Hélio ou Viviana)
        adjust_key = max(PERCENTUAIS, key=PERCENTUAIS.get)
        targets[adjust_key] += diff

    counts = df["Responsável"].value_counts().to_dict()
    for k in targets:
        counts.setdefault(k, 0)

    # Agrupa itens sem responsável pela chave descritiva para manter consistência
    df["_key"] = (
        df["Descrição de Item"].str.lower().fillna("") + "|" +
        df["Descrição longa do item"].str.lower().fillna("") + "|" +
        df["Fabricante/PN"].str.lower().fillna("")
    )

    for _, grp in df[df["Responsável"].isna()].groupby("_key"):
        # Escolhe quem está mais longe do seu alvo
        r = max(targets, key=lambda k: targets[k] - counts[k])
        df.loc[grp.index, "Responsável"] = r
        counts[r] += len(grp)

    df.drop(columns="_key", inplace=True)
    return df

# ==============================
# PROCESSAMENTO
# ==============================

def process(zip_path, output):
    work = Path(tempfile.mkdtemp())

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(work)

    rows = []

    for pdf in work.rglob("*.pdf"):
        with pdfplumber.open(str(pdf)) as p:
            raw = norm("\n".join((pg.extract_text() or "") for pg in p.pages))

        num, tipo, crit, fim, local = extract_header(raw)

        for b in re.split(r"Dados do Item", raw)[1:]:
            iid = extract_item_id(b)
            if not iid:
                continue

            qtd, und = extract_qty_unit(b, iid)
            desc = extract_desc(b)
            longd = extract_long(b)

            rows.append({
                "Numero da Oportunidade": num,
                "Tipo de Oportunidade": tipo,
                "Critério de Julgamento": crit,
                "Fim do período de cotação": fim,
                "Local de Entrega": local,
                "Item": iid,
                "Quantidade": qtd,
                "Unidade de medida": und,
                "Descrição de Item": desc,
                "Descrição longa do item": longd,
                "Fabricante/PN": extract_fab(longd)
            })

    df = pd.DataFrame(rows)

    if len(df):
        df = df.drop_duplicates(["Numero da Oportunidade", "Item"])
        df = assign(df)

    with pd.ExcelWriter(output) as w:
        df.to_excel(w, sheet_name="Consolidado", index=False)

# ==============================
# STREAMLIT
# ==============================

st.title("Separador de OPS")

uploaded = st.file_uploader("Envie o ZIP", type=["zip"])

if uploaded:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp:
        tmp.write(uploaded.read())
        tmp_path = tmp.name

    output = tmp_path.replace(".zip", ".xlsx")

    with st.spinner("Processando..."):
        process(tmp_path, output)

    df = pd.read_excel(output)

    st.success("Processado com sucesso")
    st.dataframe(df)

    with open(output, "rb") as f:
        st.download_button("Baixar Excel", f, file_name="Separar_OPS.xlsx")
