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


def patt(term):
    parts = re.split(r"\s+", str(term).strip().lower())
    return r"\b" + r"[\s\-]?".join(map(re.escape, parts)) + r"\b"


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

    fim = re.search(
        r"Fim do período de cotação\s*([0-9]{2}\.[0-9]{2}\.[0-9]{4}\s*/\s*[0-9]{2}:[0-9]{2}:[0-9]{2})",
        raw,
    )
    fim = re.sub(r"\s*/\s*", " / ", fim.group(1)).strip() if fim else ""

    ml = re.search(r"Local de Entrega\s*(.*?)\nInformações do Comprador", raw, flags=re.S)
    local = "; ".join([x.strip() for x in ml.group(1).split("\n") if x.strip()]) if ml else ""

    return numero, tipo, crit, fim, local


def extract_item_id(block):
    m = re.search(r"\n\s*(\d+)\s", block)
    return m.group(1) if m else ""


def extract_qty_unit(block, item_id):
    if not item_id:
        return "", ""

    patterns = [
        rf"(?im){item_id}.*?Material\s+([0-9\.,]+)\s+([A-Za-zÀ-ÿ]+)",
        rf"(?im){item_id}.*?([0-9\.,]+)\s+([A-Za-zÀ-ÿ]+)",
        rf"(?is)Quantidade\s*([0-9\.,]+)\s+([A-Za-zÀ-ÿ]+)",
    ]

    for p in patterns:
        m = re.search(p, block)
        if m:
            return m.group(1).strip(), m.group(2).strip()

    return "", ""


def extract_desc(block):
    m = re.search(r"Descrição de Item\s*(.*?)(?=Descrição longa|$)", block, re.S)
    return oneline(m.group(1)) if m else ""


def extract_long(block):
    m = re.search(r"Descrição longa.*?(.*)$", block, re.S)
    return oneline(m.group(1)) if m else ""


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

    df.loc[df["Tipo de Oportunidade"].fillna("").str.contains("Inaplica.", regex=False), "Responsável"] = "Viviana"

    for t in ["abb","schneider","siemens","rittal","phoenix","weidmuller","rockwell"]:
        df.loc[df["Responsável"].isna() & text.str.contains(patt(t), na=False), "Responsável"] = "Hélio"

    for t in ["skf","emerson"]:
        df.loc[df["Responsável"].isna() & text.str.contains(patt(t), na=False), "Responsável"] = "Mayara"

    for t in ["kongsberg","yamada","dnh","evac","steyr"]:
        df.loc[df["Responsável"].isna() & text.str.contains(patt(t), na=False), "Responsável"] = "Viviana"

    # balanceamento
    df["_chave"] = (
        df["Descrição de Item"].fillna("") + "|" +
        df["Descrição longa do item"].fillna("") + "|" +
        df["Fabricante/PN"].fillna("")
    ).str.lower()

    total = len(df)
    targets = {"Viviana": round(total*0.33), "Hélio": round(total*0.33)}
    targets["Mayara"] = total - sum(targets.values())

    counts = df["Responsável"].value_counts().to_dict()
    for k in targets:
        counts.setdefault(k, 0)

    pendentes = df[df["Responsável"].isna()]

    for _, grp in pendentes.groupby("_chave"):
        escolhido = max(targets, key=lambda k: targets[k] - counts[k])
        df.loc[grp.index, "Responsável"] = escolhido
        counts[escolhido] += len(grp)

    df.drop(columns="_chave", inplace=True)

    return df


# ==============================
# PROCESSAMENTO
# ==============================

def process(zip_path, output_path):
    rows = []

    with tempfile.TemporaryDirectory() as temp_dir:
        work = Path(temp_dir)
        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(work)

        for pdf in work.rglob("*.pdf"):
            with pdfplumber.open(str(pdf)) as p:
                raw = norm("\n".join((pg.extract_text() or "") for pg in p.pages))

            num, tipo, crit, fim, local = extract_header(raw)

            for bloco in re.split(r"(?i)Dados do Item", raw)[1:]:
                item = extract_item_id(bloco)
                if not item:
                    continue

                qtd, und = extract_qty_unit(bloco, item)
                desc = extract_desc(bloco)
                longd = extract_long(bloco)

                rows.append({
                    "Numero da Oportunidade": num,
                    "Tipo de Oportunidade": tipo,
                    "Critério de Julgamento": crit,
                    "Fim do período de cotação": fim,
                    "Local de Entrega": local,
                    "Item": item,
                    "Quantidade": qtd,
                    "Unidade de medida": und,
                    "Descrição de Item": desc,
                    "Descrição longa do item": longd,
                    "Fabricante/PN": extract_fab(longd),
                })

    df = pd.DataFrame(rows)

    if not df.empty:
        df = df.drop_duplicates(["Numero da Oportunidade", "Item"])

        # =============================
        # CORREÇÃO AQUI (DATA E HORA)
        # =============================
        dt = pd.to_datetime(
            df["Fim do período de cotação"].astype(str).str.replace(" / ", " ", regex=False),
            format="%d.%m.%Y %H:%M:%S",
            errors="coerce"
        )

        df["Data (cotação)"] = dt.dt.strftime("%d/%m/%Y")
        df["Hora (cotação)"] = dt.dt.strftime("%H:%M:%S")

        df = assign(df)

    with pd.ExcelWriter(output_path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Consolidado", index=False)

    return output_path


# ==============================
# STREAMLIT
# ==============================

st.set_page_config(page_title="Separador de OPS", layout="wide")

st.title("Separador de OPS")

uploaded_file = st.file_uploader("Envie o ZIP", type=["zip"])

if uploaded_file:
    try:
        with tempfile.TemporaryDirectory() as tmp:
            zip_path = Path(tmp) / uploaded_file.name
            output_path = Path(tmp) / "resultado.xlsx"

            with open(zip_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            with st.spinner("Processando..."):
                result = process(zip_path, output_path)

            df = pd.read_excel(result)

            with open(result, "rb") as f:
                data = f.read()

        st.success("Processado com sucesso")
        st.dataframe(df, use_container_width=True)

        st.download_button(
            "Baixar Excel",
            data,
            file_name="resultado.xlsx"
        )

    except Exception as e:
        st.error(f"Erro: {e}")
