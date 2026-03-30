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

    # Remove paginação
    text = re.sub(r"\bPág:\s*\d+/\d+\b", "", text, flags=re.I)

    # Remove blocos de resumo extraído do portal
    text = re.sub(
        r"(?is)Resumo extra[ií]do por.*?(?=\n|$)",
        "",
        text,
    )

    # Remove linha "Resumo da Oportunidade" quando vier colada nesse ruído
    text = re.sub(
        r"(?im)^\s*Resumo da Oportunidade\s*$",
        "",
        text,
    )

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
    local = "; ".join([x.strip() for x in ml.group(1).split("\n") if x.strip()]) if ml else ""

    return numero, tipo, crit, fim, local


def extract_item_id(block):
    patterns = [
        (r"(?m)^\s*(\d{1,6})\s+\S", 0),
        (r"(?i)Número\s+Descrição.*?\n(\d+)\s", re.S),
        (r"(?i)Número\s*\n\s*do item\s*(?:\n|\s)+([0-9A-Za-z\.\-]+)", 0),
    ]
    for pattern, flags in patterns:
        m = re.search(pattern, block, flags=flags)
        if m:
            return re.sub(r"\D", "", m.group(1))
    return ""


def extract_qty_unit(block, item_id):
    if not item_id:
        return "", ""

    patterns = [
        rf"(?im)^\s*{re.escape(item_id)}\s+.*?\bMaterial\b\s+([0-9]+(?:\.[0-9]{{3}})*(?:,[0-9]+)?)\s+([A-Za-zÀ-ÿ]+)\s+\d{{2}}\.\d{{2}}\.\d{{4}}\b",
        rf"(?im)^\s*{re.escape(item_id)}\s+.*?\b([0-9]+(?:\.[0-9]{{3}})*(?:,[0-9]+)?)\s+([A-Za-zÀ-ÿ]+)\s+\d{{2}}\.\d{{2}}\.\d{{4}}\b",
        rf"(?is)\b{re.escape(item_id)}\b.*?\bMaterial\b\s+([0-9]+(?:\.[0-9]{{3}})*(?:,[0-9]+)?)\s+([A-Za-zÀ-ÿ]+)\b",
        rf"(?is)\bQuantidade\b\s*([0-9]+(?:\.[0-9]{{3}})*(?:,[0-9]+)?)\s+\b([A-Za-zÀ-ÿ]+)\b",
    ]

    for pattern in patterns:
        m = re.search(pattern, block)
        if m:
            return m.group(1).strip(), m.group(2).strip()

    return "", ""


def extract_desc(block):
    m = re.search(
        r"(?is)(?:Descrição de Item|Descricao de Item)\s*(.*?)(?:(?:\n(?:Descrição longa do item|Descricao longa do item|Descrição longa|Descricao longa))|\Z)",
        block,
    )
    return oneline(m.group(1)) if m else ""


def extract_long(block):
    block2 = re.split(
        r"(?i)Declarações envolvidas na oportunidade|Declaracoes envolvidas na oportunidade",
        block,
        maxsplit=1,
    )[0]
    m = re.search(
        r"(?is)(?:Descrição longa do item|Descricao longa do item|Descrição longa|Descricao longa)\s*(.*)$",
        block2,
    )
    return oneline(m.group(1)) if m else ""


def extract_fab(longd):
    if not longd:
        return ""

    # Corta ruídos conhecidos antes da extração
    cleaned = re.split(
        r"(?is)Resumo extra[ií]do por|Resumo da Oportunidade",
        longd,
        maxsplit=1,
    )[0]

    m = re.search(
        r"(?is)(?:Tp:|Type:|----------)\s*(.*?)(?=Tag|Dados|Número|Declara|Resumo|$)",
        cleaned,
    )
    if not m:
        return ""

    fab = re.sub(r"\s{2,}", " ", m.group(1)).strip()

    # Remove resíduos textuais conhecidos
    fab = re.sub(r"(?is)Resumo extra[ií]do por.*$", "", fab).strip()
    fab = re.sub(r"(?is)Resumo da Oportunidade.*$", "", fab).strip()

    # Remove data/hora de extração no final:
    # exemplos:
    # 06.02.2026 às 09:48:04
    # 06.02.2026 as 09:48:04
    # 06.02.2026 09:48:04
    fab = re.sub(
        r"\s*\b\d{2}\.\d{2}\.\d{4}(?:\s*(?:às|as))?\s*\d{2}:\d{2}:\d{2}\b.*$",
        "",
        fab,
        flags=re.I,
    ).strip()

    # Remove eventual data solta no final
    fab = re.sub(r"\s*\b\d{2}\.\d{2}\.\d{4}\b.*$", "", fab).strip()

    # Remove barras ou hífens sobrando no fim
    fab = re.sub(r"[\s/\-|:;,]+$", "", fab).strip()

    return fab


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

    df.loc[
        df["Tipo de Oportunidade"].fillna("").str.contains("Inaplica.", na=False, regex=False),
        "Responsável",
    ] = "Viviana"

    helio_terms = ["abb", "schneider", "siemens", "rittal", "phoenix", "weidmuller", "rockwell"]
    for t in helio_terms:
        df.loc[
            df["Responsável"].isna() & text.str.contains(patt(t), regex=True, na=False),
            "Responsável",
        ] = "Hélio"

    for t in ["skf", "emerson"]:
        df.loc[
            df["Responsável"].isna() & text.str.contains(patt(t), regex=True, na=False),
            "Responsável",
        ] = "Mayara"

    for t in ["kongsberg", "yamada", "dnh", "evac", "steyr"]:
        df.loc[
            df["Responsável"].isna() & text.str.contains(patt(t), regex=True, na=False),
            "Responsável",
        ] = "Viviana"

    df["_chave"] = (
        df["Descrição de Item"].fillna("").str.strip().str.lower()
        + "|"
        + df["Descrição longa do item"].fillna("").str.strip().str.lower()
        + "|"
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

def process(zip_path, output_path):
    zip_path = Path(zip_path)
    output_path = Path(output_path)

    if not zip_path.exists():
        raise FileNotFoundError(f"Arquivo ZIP não encontrado: {zip_path}")

    expected_columns = [
        "Numero da Oportunidade",
        "Tipo de Oportunidade",
        "Critério de Julgamento",
        "Fim do período de cotação",
        "Data (cotação)",
        "Hora (cotação)",
        "Local de Entrega",
        "Item",
        "Quantidade",
        "Unidade de medida",
        "Descrição de Item",
        "Descrição longa do item",
        "Fabricante/PN",
        "Responsável",
    ]

    rows = []

    with tempfile.TemporaryDirectory() as temp_dir:
        work = Path(temp_dir) / "temp_ops"
        work.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(work)

        pdfs = list(work.rglob("*.pdf"))
        if not pdfs:
            df = pd.DataFrame(columns=expected_columns)
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                pd.DataFrame({"Aviso": ["Nenhum PDF encontrado no arquivo ZIP"]}).to_excel(
                    writer, sheet_name="Consolidado", index=False
                )
            return output_path

        for pdf in pdfs:
            with pdfplumber.open(str(pdf)) as p:
                raw = norm("\n".join((pg.extract_text() or "") for pg in p.pages))

            num, tipo, crit, fim, local = extract_header(raw, pdf.stem)

            for bloco in re.split(r"(?i)\bDados do Item\b", raw)[1:]:
                item_id = extract_item_id(bloco)
                if not item_id:
                    continue

                quantidade, unidade = extract_qty_unit(bloco, item_id)
                desc = extract_desc(bloco)
                longd = extract_long(bloco)

                rows.append(
                    {
                        "Numero da Oportunidade": num,
                        "Tipo de Oportunidade": tipo,
                        "Critério de Julgamento": crit,
                        "Fim do período de cotação": fim,
                        "Local de Entrega": local,
                        "Item": item_id,
                        "Quantidade": quantidade,
                        "Unidade de medida": unidade,
                        "Descrição de Item": desc,
                        "Descrição longa do item": longd,
                        "Fabricante/PN": extract_fab(longd),
                    }
                )

    if rows:
        df = pd.DataFrame(rows)
        df = df.drop_duplicates(subset=["Numero da Oportunidade", "Item"], keep="first")

        dt = pd.to_datetime(
            df["Fim do período de cotação"].astype(str).str.replace(" / ", " ", regex=False),
            format="%d.%m.%Y %H:%M:%S",
            errors="coerce",
        )

        df["Data (cotação)"] = dt.dt.strftime("%d/%m/%Y")
        df["Hora (cotação)"] = dt.dt.strftime("%H:%M:%S")

        df = assign(df)

        df = df[
            [
                "Numero da Oportunidade",
                "Tipo de Oportunidade",
                "Critério de Julgamento",
                "Fim do período de cotação",
                "Data (cotação)",
                "Hora (cotação)",
                "Local de Entrega",
                "Item",
                "Quantidade",
                "Unidade de medida",
                "Descrição de Item",
                "Descrição longa do item",
                "Fabricante/PN",
                "Responsável",
            ]
        ]
    else:
        df = pd.DataFrame(columns=expected_columns)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        if not df.empty:
            for responsavel in ["Viviana", "Hélio", "Mayara"]:
                df[df["Responsável"] == responsavel].to_excel(
                    writer,
                    sheet_name=responsavel,
                    index=False,
                )

            df.to_excel(writer, sheet_name="Consolidado", index=False)

            resumo = pd.DataFrame({"Responsável": ["Viviana", "Hélio", "Mayara"]})
            resumo["Atual"] = resumo["Responsável"].map(
                lambda x: int((df["Responsável"] == x).sum())
            )
            resumo["%"] = (resumo["Atual"] / len(df) * 100).round(1)
            resumo.to_excel(writer, sheet_name="Resumo", index=False)
        else:
            pd.DataFrame({"Aviso": ["Nenhum item encontrado nos PDFs"]}).to_excel(
                writer,
                sheet_name="Consolidado",
                index=False,
            )

    return output_path


# ==============================
# UI / STREAMLIT
# ==============================

st.set_page_config(page_title="Separador de OPS", layout="wide")

st.markdown(
    """
    <style>
        :root {
            --bg: #02053d;
            --panel: #060a53;
            --panel-2: #08106a;
            --panel-soft: rgba(255, 255, 255, 0.04);
            --border: rgba(63, 93, 255, 0.35);
            --border-soft: rgba(255, 255, 255, 0.08);
            --text: #f5f7ff;
            --muted: #a9b4ea;
            --primary: #1237ff;
            --primary-2: #3f5dff;
            --shadow: 0 18px 60px rgba(0, 0, 0, 0.28);
            --radius-xl: 28px;
            --radius-lg: 22px;
        }

        .stApp {
            background:
                radial-gradient(circle at top left, rgba(63, 93, 255, 0.20), transparent 28%),
                radial-gradient(circle at top right, rgba(18, 55, 255, 0.16), transparent 24%),
                linear-gradient(180deg, #010332 0%, var(--bg) 100%);
            color: var(--text);
        }

        .block-container {
            max-width: 1360px;
            padding-top: 1.6rem;
            padding-bottom: 2.5rem;
        }

        [data-testid="stHeader"] {
            background: transparent;
        }

        h1, h2, h3, h4, h5, h6, p, label, div, span {
            color: var(--text);
        }

        .topbar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 24px;
            background: rgba(4, 8, 70, 0.80);
            border: 1px solid var(--border-soft);
            border-radius: 24px;
            padding: 18px 24px;
            box-shadow: var(--shadow);
            margin-bottom: 22px;
        }

        .brand-wrap {
            display: flex;
            align-items: center;
            gap: 16px;
        }

        .brand-mark {
            width: 58px;
            height: 58px;
            border-radius: 14px;
            background: linear-gradient(135deg, var(--primary-2), var(--primary));
            display: flex;
            align-items: center;
            justify-content: center;
            color: white !important;
            font-size: 1.8rem;
            font-weight: 800;
            box-shadow: 0 10px 30px rgba(18, 55, 255, 0.35);
        }

        .brand-title {
            font-size: 1.5rem;
            font-weight: 800;
            line-height: 1;
            margin: 0;
        }

        .brand-subtitle {
            margin: 4px 0 0 0;
            color: var(--muted) !important;
            font-size: 0.95rem;
        }

        .nav-links {
            display: flex;
            gap: 28px;
            align-items: center;
            color: var(--muted);
            font-weight: 600;
        }

        .nav-links span.active {
            color: var(--text);
            position: relative;
        }

        .nav-links span.active::after {
            content: "";
            position: absolute;
            left: 0;
            right: 0;
            bottom: -10px;
            height: 2px;
            background: var(--primary-2);
            border-radius: 999px;
        }

        .hero {
            background: linear-gradient(180deg, rgba(6, 10, 83, 0.92) 0%, rgba(4, 7, 58, 0.92) 100%);
            border: 1px solid var(--border);
            border-radius: var(--radius-xl);
            padding: 30px;
            box-shadow: var(--shadow);
            margin-bottom: 20px;
        }

        .hero-title {
            font-size: 2.4rem;
            font-weight: 800;
            margin-bottom: 0.45rem;
        }

        .hero-subtitle {
            color: var(--muted) !important;
            font-size: 1.02rem;
            margin-bottom: 1.4rem;
            max-width: 920px;
        }

        .metric-grid {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 16px;
        }

        .metric-card {
            background: rgba(255, 255, 255, 0.04);
            border: 1px solid var(--border-soft);
            border-radius: var(--radius-lg);
            padding: 18px 20px;
            backdrop-filter: blur(8px);
        }

        .metric-label {
            color: var(--muted) !important;
            font-size: 0.9rem;
            margin-bottom: 8px;
        }

        .metric-value {
            font-size: 1.7rem;
            font-weight: 800;
            line-height: 1.1;
            color: var(--text) !important;
        }

        div[data-testid="stFileUploader"] section {
            background: rgba(255,255,255,0.03) !important;
            border: 1px dashed rgba(88, 117, 255, 0.45) !important;
            border-radius: 18px !important;
            padding: 8px !important;
        }

        div[data-testid="stFileUploaderDropzone"] {
            background: rgba(255,255,255,0.02) !important;
            border: none !important;
            border-radius: 16px !important;
        }

        div[data-testid="stDataFrame"] {
            border: 1px solid var(--border-soft) !important;
            border-radius: 18px !important;
            overflow: hidden !important;
            background: rgba(255,255,255,0.02) !important;
        }

        .stDownloadButton > button,
        .stButton > button {
            border-radius: 14px !important;
            min-height: 46px !important;
            padding: 0.72rem 1.15rem !important;
            font-weight: 800 !important;
            border: 1px solid transparent !important;
            color: white !important;
            background: linear-gradient(135deg, var(--primary), var(--primary-2)) !important;
            box-shadow: 0 10px 24px rgba(18, 55, 255, 0.28) !important;
        }

        .stAlert {
            border-radius: 16px !important;
        }

        @media (max-width: 1100px) {
            .metric-grid {
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }
            .nav-links {
                display: none;
            }
        }

        @media (max-width: 700px) {
            .metric-grid {
                grid-template-columns: 1fr;
            }
            .hero-title {
                font-size: 1.8rem;
            }
            .topbar {
                padding: 16px 18px;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)


def render_topbar():
    st.markdown(
        """
        <div class="topbar">
            <div class="brand-wrap">
                <div class="brand-mark">N</div>
                <div>
                    <p class="brand-title">NEXT SUPPLY</p>
                    <p class="brand-subtitle">Separador de OPS</p>
                </div>
            </div>
            <div class="nav-links">
                <span class="active">Dashboard</span>
                <span>Histórico de Uploads</span>
                <span>Exportações</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_hero(preview_df=None):
    has_preview = preview_df is not None and not preview_df.empty
    total_registros = len(preview_df) if preview_df is not None else 0
    total_arquivos = 1 if preview_df is not None else 0
    total_responsaveis = (
        int(preview_df["Responsável"].nunique())
        if has_preview and "Responsável" in preview_df.columns
        else 0
    )
    status_text = "Pronto" if has_preview else "Aguardando arquivo"

    st.markdown(
        f"""
        <div class="hero">
            <div class="hero-title">Separador de OPS</div>
            <div class="hero-subtitle">
                Faça upload de um arquivo ZIP contendo os PDFs das oportunidades e visualize o consolidado em um painel alinhado à identidade visual da Next Supply.
            </div>
            <div class="metric-grid">
                <div class="metric-card">
                    <div class="metric-label">Arquivos enviados</div>
                    <div class="metric-value">{total_arquivos}</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Registros consolidados</div>
                    <div class="metric-value">{total_registros}</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Responsáveis</div>
                    <div class="metric-value">{total_responsaveis}</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Status</div>
                    <div class="metric-value">{status_text}</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ==============================
# APP
# ==============================

render_topbar()

preview_df = None
excel_bytes = None
result_name = None

st.title("Upload de arquivo ZIP")
st.caption("Selecione o arquivo compactado com os PDFs das oportunidades.")

uploaded_file = st.file_uploader(
    "Selecione o arquivo ZIP",
    type=["zip"],
)

if uploaded_file is not None:
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dir = Path(temp_dir)

            zip_path = temp_dir / uploaded_file.name
            output_path = temp_dir / f"{Path(uploaded_file.name).stem}_saida.xlsx"

            with open(zip_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            with st.spinner("Processando PDFs e gerando Excel..."):
                result_path = process(zip_path, output_path)

            try:
                preview_df = pd.read_excel(result_path, sheet_name="Consolidado")
            except Exception:
                preview_df = pd.DataFrame()

            with open(result_path, "rb") as f:
                excel_bytes = f.read()

            result_name = result_path.name

        st.success("Arquivo processado com sucesso.")

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")

render_hero(preview_df)

if preview_df is not None:
    st.subheader("Prévia do Consolidado")
    st.caption("Visualização do arquivo consolidado gerado a partir dos PDFs processados.")
    st.dataframe(preview_df, use_container_width=True)

if excel_bytes is not None and result_name is not None:
    st.download_button(
        label="Baixar Excel gerado",
        data=excel_bytes,
        file_name=result_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
