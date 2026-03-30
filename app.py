import re

import sys

import zipfile

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

    parts = re.split(r"\s+", t.lower())

    return r"\b" + r"[\s\-]?".join(parts) + r"\b"
 
# ==============================

# EXTRAÇÃO

# ==============================
 
def extract_header(raw, fallback=""):

    m = re.search(r"Número da Oportunidade.*?(\d{10})", raw, flags=re.S)

    numero = m.group(1) if m else ""
 
    tipo = re.search(r"Tipo de Oportunidade\s*([^\n]+)", raw)

    tipo = tipo.group(1).strip() if tipo else ""
 
    crit = re.search(r"Critério de Julgamento\s*([^\n]+)", raw)

    crit = crit.group(1).strip() if crit else ""
 
    fim = re.search(r"Fim do período de cotação\s*([0-9\.]+\s*/\s*[0-9:]+)", raw)

    fim = fim.group(1) if fim else ""
 
    ml = re.search(r"Local de Entrega\s*(.*?)\nInformações do Comprador", raw, flags=re.S)

    local = "; ".join([x.strip() for x in ml.group(1).split("\n") if x.strip()]) if ml else ""
 
    return numero, tipo, crit, fim, local
 
def extract_fab(longd):

    if not longd:

        return ""

    m = re.search(r"(?is)(?:Tp:|Type:|----------)\s*(.*?)(?=Tag|Dados|Número|Declara|$)", longd)

    return re.sub(r"\s{2,}", " ", m.group(1)).strip() if m else ""
 
# ==============================

# RESPONSÁVEL (VERSÃO FINAL)

# ==============================
 
def assign(df):

    text = (df["Descrição de Item"] + " " + df["Descrição longa do item"] + " " + df["Fabricante/PN"]).str.lower()
 
    df["Responsável"] = None
 
    # PRIORIDADE: Inaplica

    df.loc[df["Tipo de Oportunidade"].str.contains("Inaplica.", na=False), "Responsável"] = "Viviana"
 
    # HELIO (TRAVADO)

    helio_terms = ["abb","schneider","siemens","rittal","phoenix","weidmuller","rockwell"]

    for t in helio_terms:

        df.loc[df["Responsável"].isna() & text.str.contains(patt(t), regex=True), "Responsável"] = "Hélio"
 
    # MAYARA (TRAVADO)

    for t in ["skf","emerson"]:

        df.loc[df["Responsável"].isna() & text.str.contains(patt(t), regex=True), "Responsável"] = "Mayara"
 
    # VIVIANA

    for t in ["kongsberg","yamada","dnh","evac","steyr"]:

        df.loc[df["Responsável"].isna() & text.str.contains(patt(t), regex=True), "Responsável"] = "Viviana"
 
    # ==============================

    # GARANTIA: ITENS IGUAIS JUNTOS

    # ==============================
 
    df["_chave"] = (

        df["Descrição de Item"].fillna("").str.strip().str.lower() + "|" +

        df["Descrição longa do item"].fillna("").str.strip().str.lower() + "|" +

        df["Fabricante/PN"].fillna("").str.strip().str.lower()

    )
 
    total = len(df)

    targets = {

        "Viviana": round(total * 0.33),

        "Hélio": round(total * 0.33)

    }

    targets["Mayara"] = total - sum(targets.values())
 
    counts = df["Responsável"].value_counts().to_dict()

    for k in targets:

        counts.setdefault(k, 0)
 
    pendentes = df[df["Responsável"].isna()]
 
    for _, grp in pendentes.groupby("_chave", sort=False):

        r = max(targets, key=lambda k: targets[k] - counts[k])

        idxs = grp.index.tolist()

        df.loc[idxs, "Responsável"] = r

        counts[r] += len(idxs)
 
    df.drop(columns=["_chave"], inplace=True)
 
    return df
 
# ==============================

# PROCESSAMENTO

# ==============================
 
def process(zip_path, output):

    work = Path("temp_ops")

    work.mkdir(exist_ok=True)
 
    with zipfile.ZipFile(zip_path, "r") as z:

        z.extractall(work)
 
    rows = []
 
    for pdf in work.rglob("*.pdf"):

        with pdfplumber.open(str(pdf)) as p:

            raw = norm("\n".join((pg.extract_text() or "") for pg in p.pages))
 
        num, tipo, crit, fim, local = extract_header(raw)
 
        for b in re.split(r"Dados do Item", raw)[1:]:

            m = re.search(r"\n\s*(\d+)\s", b)

            if not m:

                continue
 
            iid = m.group(1)
 
            desc = oneline(re.search(r"Descrição de Item\s*(.*?)(?=Descrição longa|$)", b, re.S).group(1) if re.search(r"Descrição de Item", b) else "")

            longd = oneline(re.search(r"Descrição longa.*?(.*)$", b, re.S).group(1) if re.search(r"Descrição longa", b) else "")
 
            rows.append({

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

                "Fabricante/PN": extract_fab(longd)

            })
 
    df = pd.DataFrame(rows).drop_duplicates(["Numero da Oportunidade","Item"])
 
    df = assign(df)
 
    with pd.ExcelWriter(output) as w:

        for r in ["Viviana","Hélio","Mayara"]:

            df[df["Responsável"] == r].to_excel(w, sheet_name=r, index=False)

        df.to_excel(w, sheet_name="Consolidado", index=False)
 
# ==============================

# EXECUÇÃO

# ==============================
 
if __name__ == "__main__":

    process(sys.argv[1], "saida.xlsx")
 
