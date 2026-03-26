import re
import zipfile
import tempfile
from io import BytesIO
from pathlib import Path

import pandas as pd
import pdfplumber
import streamlit as st


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


def extract_header(raw: str, fallback: str = "") -> tuple[str, str, str, str, str]:
    m = re.search(r"Número da Oportunidade.*?(\d{10})", raw, flags=re.S)
    if m:
        numero = m.group(1)
    else:
        m = re.search(r"(\d{10})", fallback)
        numero = m.group(1) if m else ""

    m = re.search(r"Tipo de Oportunidade\s*([^\n]+)", raw)
    tipo = m.group(1).strip() if m else ""

    m = re.search(r"Critério de Julgamento\s*([^\n]+)", raw)
    criterio = m.group(1).strip() if m else ""

    m = re.search(
        r"Fim do período de cotação\s*([0-9]{2}\.[0-9]{2}\.[0-9]{4}\s*/\s*[0-9]{2}:[0-9]{2}:[0-9]{2})",
        raw,
    )
    fim = re.sub(r"\s*/\s*", " / ", m.group(1)).strip() if m else ""

    local = ""
    ml = re.search(r"Local de Entrega\s*(.*?)\nInformações do Comprador", raw, flags=re.S)
    if ml:
        local = "; ".join([x.strip() for x in ml.group(1).split("\n") if x.strip()])

    return numero, tipo, criterio, fim, local


def extract_item_id(block: str) -> str:
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


def extract_qty_unit(block: str, item_id: str) -> tuple[str, str]:
    patterns = [
        rf"(?im)^\s*{re.escape(item_id)}\s+.*?\bMaterial\b\s+([0-9]+(?:\.[0-9]{{3}})*(?:,[0-9]+)?)\s+([A-Za-zÀ-ÿ]+)\s+\d{{2}}\.\d{{2}}\.\d{{4}}\b",
        rf"(?im)^\s*{re.escape(item_id)}\s+.*?\b([0-9]+(?:\.[0-9]{{3}})*(?:,[0-9]+)?)\s+([A-Za-zÀ-ÿ]+)\s+\d{{2}}\.\d{{2}}\.\d{{4}}\b",
        rf"(?is)\b{re.escape(item_id)}\b.*?\bMaterial\b\s+([0-9]+(?:\.[0-9]{{3}})*(?:,[0-9]+)?)\s+([A-Za-zÀ-ÿ]+)\b",
    ]
    for pattern in patterns:
        m = re.search(pattern, block)
        if m:
            return m.group(1).strip(), m.group(2).strip()
    return "", ""


def extract_desc(block: str) -> str:
    m = re.search(
        r"(?is)(?:Descrição de Item|Descricao de Item)\s*(.*?)(?:(?:\n(?:Descrição longa do item|Descricao longa do item|Descrição longa|Descricao longa))|\Z)",
        block,
    )
    return oneline(m.group(1)) if m else ""


def extract_long(block: str) -> str:
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


def extract_fab(longd: str) -> str:
    if not longd:
        return ""
    m = re.search(r"(?is)(?:Tp:|Type:|----------)\s*(.*?)(?=Tag|Dados|Número|Declara|$)", longd)
    if not m:
        return ""
    return re.sub(r"\s{2,}", " ", m.group(1)).strip()


def make_term_pattern(term: str) -> str:
    parts = re.split(r"\s+", term.lower().strip())
    return r"\b" + r"[\s\-]?".join(map(re.escape, parts)) + r"\b"


def assign_responsavel(df: pd.DataFrame) -> pd.DataFrame:
    text = (
        df["Descrição de Item"].fillna("")
        + " "
        + df["Descrição longa do item"].fillna("")
        + " "
        + df["Fabricante/PN"].fillna("")
    ).str.lower()

    tipo = df["Tipo de Oportunidade"].fillna("")
    mask_total = tipo.str.contains("Dispensa Total", case=False, regex=False) | tipo.str.contains(
        "Inaplica. total", case=False, regex=False
    )
    mask_inap = tipo.str.contains("Inaplica.", case=False, regex=False) & ~tipo.str.contains(
        "Inaplica. total", case=False, regex=False
    )

    df["Resp_pre"] = pd.NA
    df.loc[mask_inap, "Resp_pre"] = "Viviana"

    rules = {
        "Hélio": ["schneider", "siemens"],
        "Mayara": ["rittal", "skf", "emerson"],
        "Viviana": ["kongsberg", "yamada", "george", "ston", "dnh", "evac", "steyr"],
        "Ana Beatriz": ["phoenix", "danfoss", "ex heat"],
    }

    for responsavel, terms in rules.items():
        regex = "|".join(make_term_pattern(t) for t in terms)
        df.loc[df["Resp_pre"].isna() & text.str.contains(regex, regex=True), "Resp_pre"] = responsavel

    df["Responsável"] = pd.NA
    df.loc[mask_inap, "Responsável"] = "Viviana"

    total = len(df)
    targets = {
        "Ana Beatriz": round(total * 0.30),
        "Viviana": round(total * 0.30),
        "Hélio": round(total * 0.20),
    }
    targets["Mayara"] = total - sum(targets.values())

    counts = df["Responsável"].value_counts(dropna=False).to_dict()
    for key in targets:
        counts.setdefault(key, 0)

    for _, grp in df[mask_total].groupby("Numero da Oportunidade"):
        idxs = grp.index.tolist()
        pre = df.loc[idxs, "Resp_pre"].dropna()
        chosen = (
            pre.value_counts().idxmax()
            if len(pre)
            else max(targets, key=lambda k: targets[k] - counts.get(k, 0))
        )
        df.loc[idxs, "Responsável"] = chosen
        counts[chosen] += len(idxs)

    for idx in df[(~mask_total) & df["Responsável"].isna() & df["Resp_pre"].notna()].index:
        chosen = df.at[idx, "Resp_pre"]
        df.at[idx, "Responsável"] = chosen
        counts[chosen] += 1

    for idx in df[df["Responsável"].isna()].index:
        chosen = max(targets, key=lambda k: targets[k] - counts.get(k, 0))
        df.at[idx, "Responsável"] = chosen
        counts[chosen] += 1

    df.drop(columns=["Resp_pre"], inplace=True)

    max_unique = int(df.groupby(["Numero da Oportunidade", "Item"])["Responsável"].nunique().max())
    if max_unique > 1:
        raise ValueError("Conflito detectado: a mesma linha foi atribuída a mais de um responsável.")

    return df


def process_zip(zip_path: str | Path, output_path: str | Path) -> Path:
    zip_path = Path(zip_path)
    output_path = Path(output_path)

    work = output_path.parent / f"_tmp_{output_path.stem}"
    work.mkdir(parents=True, exist_ok=True)

    for p in sorted(work.rglob("*"), reverse=True):
        try:
            if p.is_file():
                p.unlink()
            else:
                p.rmdir()
        except Exception:
            pass

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(work)

    pdfs = sorted(work.rglob("*.pdf"))
    rows = []

    for pdf_path in pdfs:
        with pdfplumber.open(str(pdf_path)) as pdf:
            raw = norm("\n".join((pg.extract_text() or "") for pg in pdf.pages))

        numero, tipo, criterio, fim, local = extract_header(raw, pdf_path.stem)

        for block in re.split(r"(?i)\bDados do Item\b", raw)[1:]:
            item = extract_item_id(block)
            if not item:
                continue

            quantidade, unidade = extract_qty_unit(block, item)
            desc = extract_desc(block)
            longd = extract_long(block)
            fab = extract_fab(longd)

            rows.append(
                {
                    "Numero da Oportunidade": numero,
                    "Tipo de Oportunidade": tipo,
                    "Critério de Julgamento": criterio,
                    "Fim do período de cotação": fim,
                    "Local de Entrega": local,
                    "Item": item,
                    "Quantidade": quantidade,
                    "Unidade de medida": unidade,
                    "Descrição de Item": desc,
                    "Descrição longa do item": longd,
                    "Fabricante/PN": fab,
                }
            )

    df = pd.DataFrame(rows)
    if len(df):
        df = df.drop_duplicates(subset=["Numero da Oportunidade", "Item"], keep="first")

        dt = pd.to_datetime(
            df["Fim do período de cotação"].astype(str).str.replace(" / ", " ", regex=False),
            format="%d.%m.%Y %H:%M:%S",
            errors="coerce",
        )
        df["Data (cotação)"] = dt.dt.date
        df["Hora (cotação)"] = dt.dt.time
        df["Status"] = ""

        df = assign_responsavel(df)

        cols = [
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
            "Status",
        ]
        df = df[cols]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        if len(df):
            for responsavel in ["Ana Beatriz", "Viviana", "Hélio", "Mayara"]:
                df[df["Responsável"] == responsavel].to_excel(
                    writer, sheet_name=responsavel, index=False
                )
            df.to_excel(writer, sheet_name="Consolidado", index=False)

            resumo = pd.DataFrame({"Responsável": ["Ana Beatriz", "Viviana", "Hélio", "Mayara"]})
            resumo["Atual"] = resumo["Responsável"].map(lambda x: int((df["Responsável"] == x).sum()))
            resumo["%"] = (resumo["Atual"] / len(df) * 100).round(1)
            resumo.to_excel(writer, sheet_name="Resumo", index=False)
        else:
            pd.DataFrame({"Aviso": ["Nenhum item encontrado nos PDFs"]}).to_excel(
                writer, sheet_name="Consolidado", index=False
            )

    return output_path


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
            --success-bg: rgba(12, 194, 126, 0.12);
            --success-border: rgba(12, 194, 126, 0.35);
            --success-text: #9df0ca;
            --danger-bg: rgba(255, 82, 82, 0.10);
            --danger-border: rgba(255, 82, 82, 0.30);
            --danger-text: #ffb3b3;
            --shadow: 0 18px 60px rgba(0, 0, 0, 0.28);
            --radius-xl: 28px;
            --radius-lg: 22px;
            --radius-md: 16px;
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

        .main * {
            color: var(--text);
        }

        [data-testid="stHeader"] {
            background: transparent;
        }

        h1, h2, h3 {
            color: var(--text) !important;
            letter-spacing: -0.02em;
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
            color: white;
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
            color: var(--muted);
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
            color: var(--muted);
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
            color: var(--muted);
            font-size: 0.9rem;
            margin-bottom: 8px;
        }

        .metric-value {
            font-size: 1.7rem;
            font-weight: 800;
            line-height: 1.1;
        }

        .section-card {
            background: linear-gradient(180deg, rgba(5, 10, 74, 0.92) 0%, rgba(4, 7, 58, 0.95) 100%);
            border: 1px solid var(--border-soft);
            border-radius: var(--radius-xl);
            padding: 24px;
            box-shadow: var(--shadow);
            margin-bottom: 18px;
        }

        .section-title {
            font-size: 1.7rem;
            font-weight: 800;
            margin-bottom: 0.3rem;
        }

        .section-subtitle {
            color: var(--muted);
            font-size: 0.95rem;
            margin-bottom: 1rem;
        }

        .upload-shell {
            border: 1.5px dashed rgba(88, 117, 255, 0.45);
            border-radius: 22px;
            background: linear-gradient(180deg, rgba(255,255,255,0.03) 0%, rgba(255,255,255,0.02) 100%);
            padding: 10px;
        }

        div[data-testid="stFileUploader"] {
            background: transparent;
        }

        div[data-testid="stFileUploader"] section {
            border: none !important;
            background: transparent !important;
            padding: 0 !important;
        }

        div[data-testid="stFileUploaderDropzone"] {
            background: rgba(255,255,255,0.02) !important;
            border: none !important;
            padding: 18px !important;
            border-radius: 18px !important;
        }

        div[data-testid="stFileUploaderDropzone"] * {
            color: var(--text) !important;
        }

        .status-success {
            background: var(--success-bg);
            border: 1px solid var(--success-border);
            color: var(--success-text);
            border-radius: 16px;
            padding: 0.95rem 1rem;
            font-weight: 600;
            margin-top: 14px;
        }

        .status-error {
            background: var(--danger-bg);
            border: 1px solid var(--danger-border);
            color: var(--danger-text);
            border-radius: 16px;
            padding: 0.95rem 1rem;
            font-weight: 600;
            margin-top: 14px;
        }

        div[data-testid="stSpinner"] > div {
            border-top-color: var(--primary-2) !important;
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

        .stDownloadButton > button:hover,
        .stButton > button:hover {
            filter: brightness(1.05);
            transform: translateY(-1px);
        }

        .toolbar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 12px;
            margin-bottom: 12px;
        }

        .pill {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            background: rgba(63, 93, 255, 0.12);
            color: #dce4ff;
            border: 1px solid rgba(63, 93, 255, 0.24);
            border-radius: 999px;
            padding: 8px 12px;
            font-size: 0.84rem;
            font-weight: 700;
        }

        label, .stMarkdown, p {
            color: var(--text) !important;
        }

        div[data-baseweb="input"] > div,
        div[data-baseweb="select"] > div,
        div[data-testid="stTextInputRootElement"] > div,
        div[data-testid="stNumberInputRootElement"] > div {
            background: rgba(255, 255, 255, 0.04) !important;
            border: 1px solid var(--border-soft) !important;
            border-radius: 14px !important;
            color: var(--text) !important;
        }

        input, textarea {
            color: var(--text) !important;
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


def render_topbar() -> None:
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


def render_hero(has_preview: bool, preview_df: pd.DataFrame | None = None) -> None:
    total_registros = len(preview_df) if has_preview and preview_df is not None else 0
    total_arquivos = 1 if has_preview else 0
    total_responsaveis = (
        int(preview_df["Responsável"].nunique())
        if has_preview and preview_df is not None and "Responsável" in preview_df.columns
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


render_topbar()

uploaded_file = None
error_message = None
preview_df = None
excel_bytes = None
result_name = None

st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.markdown("<div class='section-title'>Upload de arquivo ZIP</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='section-subtitle'>Selecione o arquivo compactado com os PDFs das oportunidades.</div>",
    unsafe_allow_html=True,
)
st.markdown("<div class='upload-shell'>", unsafe_allow_html=True)
uploaded_file = st.file_uploader("Selecione o arquivo ZIP", type=["zip"], label_visibility="collapsed")
st.markdown("</div>", unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

if uploaded_file is not None:
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dir = Path(temp_dir)

            zip_path = temp_dir / uploaded_file.name
            output_path = temp_dir / f"{Path(uploaded_file.name).stem}_Separar_OPS.xlsx"

            with open(zip_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            with st.spinner("Processando PDFs e gerando Excel..."):
                result_path = process_zip(zip_path, output_path)

            preview_df = pd.read_excel(result_path, sheet_name="Consolidado")

            with open(result_path, "rb") as f:
                excel_bytes = f.read()

            result_name = result_path.name

        st.markdown(
            "<div class='status-success'>✅ Arquivo processado com sucesso.</div>",
            unsafe_allow_html=True,
        )

    except Exception as e:
        error_message = f"Erro ao processar o arquivo: {e}"
        st.markdown(
            f"<div class='status-error'>❌ {error_message}</div>",
            unsafe_allow_html=True,
        )

render_hero(preview_df is not None, preview_df)

if preview_df is not None and excel_bytes is not None and result_name is not None:
    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.markdown(
        """
        <div class='toolbar'>
            <div>
                <div class='section-title'>Prévia do Consolidado</div>
                <div class='section-subtitle'>Visualização do arquivo consolidado gerado a partir dos PDFs processados.</div>
            </div>
            <div class='pill'>Consolidado pronto</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.dataframe(preview_df, use_container_width=True)

    st.download_button(
        label="Baixar Excel gerado",
        data=excel_bytes,
        file_name=result_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("</div>", unsafe_allow_html=True)
