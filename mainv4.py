# app.py — Sistema de Monitoramento de Produção do PPG (v7.0 - Website Moderno)
# Streamlit + Google Sheets + E-mails
# =========================================================

import os, time, base64, uuid, hashlib, hmac, smtplib, re
from email.message import EmailMessage
from datetime import datetime
from collections import Counter
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError, WorksheetNotFound, SpreadsheetNotFound
from PIL import Image
from wordcloud import WordCloud
import matplotlib.pyplot as plt

# ---------------------------------------------------------
# CONFIG
# ---------------------------------------------------------
st.set_page_config(
    page_title="PPG PNBC — Monitor de Produção",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"  # Sidebar colapsada por padrão
)

try:
    banner = Image.open("banner_ppg.png")
    st.image(banner, use_container_width=True)
except Exception:
    pass

SPREADSHEET_ID = st.secrets.get("GSHEET_SPREADSHEET_ID", "")
if not SPREADSHEET_ID:
    st.error("Configure GSHEET_SPREADSHEET_ID nos Secrets.")
    st.stop()

SHEET_USERS   = "users"
SHEET_CAD     = "cadastro_requests"
SHEET_PROD    = "producoes"
SHEET_PART    = "participacoes"
SHEET_VINC    = "vinculos_discentes"
SHEET_ORIENT  = "orientacoes"
SHEET_ENSINO  = "atividades_ensino"
SHEET_IMPACTO = "atividades_impacto"

HEADERS_USERS   = ["username", "name", "email", "role", "orientador", "password_hash", "created_at"]
HEADERS_CAD     = ["id", "name", "username", "email", "role", "orientador",
                   "password_hash", "status", "created_at", "reviewed_at",
                   "reviewed_by", "review_reason"]
HEADERS_PROD    = ["id", "docente_username", "titulo", "tipo", "ano",
                   "veiculo", "autores", "doi", "descricao", "co_autores",
                   "discente_primeiro_autor", "docente_ultimo_autor", "created_at"]
HEADERS_PART    = ["id", "producao_id", "tipo_participacao",
                   "nome_participante", "vinculo", "created_at"]
HEADERS_VINC    = ["id", "discente_username", "orientador_username",
                   "producao_id", "created_at"]
HEADERS_ORIENT  = ["id", "docente_username", "discente_nome", "tipo", 
                   "ano_inicio", "ano_conclusao", "status", "created_at"]
HEADERS_ENSINO  = ["id", "docente_username", "tipo", "titulo", "periodo", "ano",
                   "carga_horaria", "nivel", "descricao", "created_at"]
HEADERS_IMPACTO = ["id", "docente_username", "tipo", "titulo", "descricao",
                   "data", "publico_alvo", "local", "created_at"]

ANOS = ["2025", "2026", "2027", "2028"]

TIPOS_PRODUCAO = [
    "Artigo em periódico", "Livro", "Capítulo de livro", "Trabalho em evento",
    "Orientação de TCC", "Orientação de Mestrado", "Orientação de Doutorado",
    "Patente", "Produto técnico/tecnológico", "Outro",
]

TIPOS_PARTICIPACAO = [
    "Discente do PPG", "Discente UFPA externo ao PPG",
    "Pesquisador estrangeiro", "Pesquisador nacional/institucional",
]

TIPOS_ORIENTACAO = ["Mestrado", "Doutorado"]
STATUS_ORIENTACAO = ["Em andamento", "Concluída", "Trancada", "Abandonada"]

TIPOS_ENSINO = [
    "Disciplina na Graduação", "Disciplina na Pós-Graduação",
    "Curso de extensão", "Oficina/Workshop", "Minicurso",
    "Coordenação de curso", "Elaboração de material didático",
]

NIVEIS_ENSINO = ["Graduação", "Pós-Graduação", "Extensão", "Técnico", "Outro"]

TIPOS_IMPACTO = [
    "Palestra/Conferência", "Consultoria técnica", "Projeto de extensão",
    "Matéria em mídia (jornal, TV, rádio)", "Parecer técnico",
    "Participação em comitê/conselho", "Organização de evento",
    "Depoimento em audiência pública", "Produto de divulgação científica",
    "Outro",
]

# ---------------------------------------------------------
# STOPWORDS
# ---------------------------------------------------------
STOPWORDS = {
    "de", "a", "o", "que", "e", "do", "da", "em", "um", "para", "com", "no",
    "uma", "os", "por", "mais", "as", "dos", "como", "mas", "foi", "ao", "ele",
    "das", "tem", "à", "seu", "sua", "ou", "ser", "quando", "muito", "há",
    "nos", "já", "está", "eu", "também", "só", "pelo", "pela", "até", "isso",
    "ela", "entre", "era", "depois", "sem", "mesmo", "aos", "ter", "seus",
    "quem", "nas", "me", "esse", "eles", "estão", "você", "tinha", "foram",
    "essa", "num", "nem", "suas", "meu", "às", "minha", "têm", "numa", "pelos",
    "elas", "havia", "seja", "qual", "será", "nós", "tenho", "lhe", "deles",
    "essas", "esses", "pelas", "este", "fosse", "dele", "tu", "te", "vocês",
    "vos", "lhes", "meus", "minhas", "teu", "tua", "teus", "tuas", "nosso",
    "nossa", "nossos", "nossas", "dela", "delas", "esta", "estes", "estas",
    "aquele", "aquela", "aqueles", "aquelas", "isto", "aquilo", "estou",
    "estamos", "estive", "esteve", "estivemos", "estiveram", "estava",
    "estávamos", "estavam", "estivera", "estivéramos", "estejamos", "estejam",
    "estivesse", "estivéssemos", "estivessem", "estiver", "estivermos",
    "estiverem", "hei", "havemos", "hão", "houve", "houvemos", "houveram",
    "houvera", "houvéramos", "haja", "hajamos", "hajam", "houvesse",
    "houvéssemos", "houvessem", "houver", "houvermos", "houverem", "houverei",
    "houverá", "houveremos", "houverão", "houveria", "houveríamos", "houveriam",
    "sou", "somos", "são", "éramos", "eram", "fui", "fomos", "fora", "fôramos",
    "sejamos", "fôssemos", "formos", "forem", "serei", "seremos", "serão",
    "seria", "seríamos", "seriam", "temos", "tém", "tínhamos", "tinham",
    "tive", "teve", "tivemos", "tiveram", "tivera", "tivéramos", "tenha",
    "tenhamos", "tenham", "tivesse", "tivéssemos", "tivessem", "tiver",
    "tivermos", "tiverem", "terei", "terá", "teremos", "terão", "teria",
    "teríamos", "teriam", "sobre", "sob", "após", "apois", "através", "segundo",
    "conforme", "mediante", "perante", "ante", "contra", "versus", "vs",
    "the", "and", "of", "to", "in", "for", "is", "on", "that", "by", "this",
    "with", "from", "are", "be", "an", "it", "at", "or", "as", "was", "have",
    "has", "had", "not", "but", "all", "can", "her", "were", "there", "their",
    "which", "one", "would", "will", "each", "about", "how", "up", "out",
    "many", "then", "them", "these", "so", "some", "other", "than", "into",
    "its", "your", "also", "new", "may", "day", "just", "after", "before",
    "between", "under", "above", "during", "through", "while", "both", "same",
    "another", "such", "only", "own", "because", "being", "using", "used",
    "based", "approach", "method", "study", "analysis", "results", "data",
    "using", "different", "new", "novel", "proposed", "performance", "system",
    "model", "application", "evaluation", "effect", "impact", "role",
}

# ---------------------------------------------------------
# CSS MODERNO (ESTILO WEBSITE)
# ---------------------------------------------------------
st.markdown("""
<style>
    /* Esconder sidebar completamente */
    [data-testid="stSidebar"] {
        display: none !important;
    }
    
    /* Header moderno */
    .header-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px 0;
        margin-bottom: 30px;
        border-radius: 0 0 20px 20px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.1);
    }
    
    .header-title {
        color: white;
        font-size: 2.5rem;
        font-weight: 700;
        text-align: center;
        margin: 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
    }
    
    .header-subtitle {
        color: rgba(255,255,255,0.9);
        font-size: 1.1rem;
        text-align: center;
        margin-top: 10px;
    }
    
    /* Navegação moderna */
    .nav-container {
        background: white;
        padding: 15px 0;
        margin-bottom: 30px;
        border-radius: 12px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.08);
    }
    
    .nav-button {
        background: transparent;
        border: 2px solid #667eea;
        color: #667eea;
        padding: 10px 20px;
        border-radius: 8px;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.3s ease;
        margin: 0 5px;
    }
    
    .nav-button:hover {
        background: #667eea;
        color: white;
        transform: translateY(-2px);
    }
    
    .nav-button.active {
        background: #667eea;
        color: white;
    }
    
    /* Cards modernos */
    .metric-card {
        background: white;
        border-radius: 16px;
        padding: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        text-align: center;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
        border: 1px solid #f0f0f0;
    }
    
    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 25px rgba(0,0,0,0.12);
    }
    
    .metric-value {
        font-size: 3rem;
        font-weight: 700;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }
    
    .metric-label {
        color: #666;
        font-size: 0.95rem;
        margin-top: 10px;
        font-weight: 500;
    }
    
    /* Seções */
    .section-container {
        background: white;
        border-radius: 16px;
        padding: 30px;
        margin-bottom: 30px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
    }
    
    .section-title {
        font-size: 1.8rem;
        font-weight: 700;
        color: #333;
        margin-bottom: 20px;
        padding-bottom: 10px;
        border-bottom: 3px solid #667eea;
    }
    
    /* Vídeo container */
    .video-section {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 20px;
        padding: 40px;
        margin-bottom: 30px;
        text-align: center;
        box-shadow: 0 8px 30px rgba(102, 126, 234, 0.3);
    }
    
    .video-title {
        color: white;
        font-size: 2rem;
        font-weight: 700;
        margin-bottom: 20px;
    }
    
    /* Footer */
    .footer {
        background: #2d3748;
        color: white;
        padding: 40px 0;
        margin-top: 50px;
        border-radius: 20px 20px 0 0;
        text-align: center;
    }
    
    .footer-text {
        color: rgba(255,255,255,0.8);
        font-size: 0.9rem;
    }
    
    /* Login form */
    .login-container {
        max-width: 500px;
        margin: 50px auto;
        background: white;
        padding: 40px;
        border-radius: 20px;
        box-shadow: 0 10px 40px rgba(0,0,0,0.1);
    }
    
    .login-title {
        font-size: 2rem;
        font-weight: 700;
        color: #333;
        text-align: center;
        margin-bottom: 30px;
    }
    
    /* Badges */
    .badge {
        display: inline-block;
        padding: 5px 12px;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        margin: 2px;
    }
    
    .badge-primary { background: #e3f2fd; color: #1976d2; }
    .badge-success { background: #e8f5e9; color: #388e3c; }
    .badge-warning { background: #fff3e0; color: #f57c00; }
    
    /* Hero section */
    .hero {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 60px 40px;
        border-radius: 20px;
        margin-bottom: 40px;
        text-align: center;
    }
    
    .hero-title {
        font-size: 3rem;
        font-weight: 800;
        margin-bottom: 20px;
    }
    
    .hero-subtitle {
        font-size: 1.3rem;
        opacity: 0.95;
        max-width: 800px;
        margin: 0 auto;
        line-height: 1.6;
    }
    
    /* Features grid */
    .feature-card {
        background: white;
        padding: 30px;
        border-radius: 16px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        text-align: center;
        transition: all 0.3s ease;
        border: 2px solid transparent;
    }
    
    .feature-card:hover {
        border-color: #667eea;
        transform: translateY(-5px);
    }
    
    .feature-icon {
        font-size: 3rem;
        margin-bottom: 15px;
    }
    
    .feature-title {
        font-size: 1.3rem;
        font-weight: 700;
        color: #333;
        margin-bottom: 10px;
    }
    
    .feature-desc {
        color: #666;
        font-size: 0.95rem;
        line-height: 1.6;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# FUNÇÕES AUXILIARES
# ---------------------------------------------------------
def create_metric_card(value, label, icon="📊"):
    return f"""
    <div class="metric-card">
        <div style="font-size: 2.5rem; margin-bottom: 10px;">{icon}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-label">{label}</div>
    </div>
    """

def create_nav_button(label, page, current_page):
    active_class = "active" if page == current_page else ""
    return f"""
    <button class="nav-button {active_class}" onclick="document.getElementById('page-{page}').scrollIntoView({{behavior: 'smooth'}})">
        {label}
    </button>
    """

# ---------------------------------------------------------
# GOOGLE SHEETS
# ---------------------------------------------------------
@st.cache_resource
def gclient():
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["GSERVICE"], scopes=scopes)
    return gspread.authorize(creds)

def _extract_api_error_info(e):
    status = getattr(getattr(e, "response", None), "status_code", None)
    text = getattr(getattr(e, "response", None), "text", "")[:400]
    return status, text

def _retryable(status): return status in (429, 500, 503)

@st.cache_resource
def spreadsheet():
    last_err = None
    for attempt in range(5):
        try:
            return gclient().open_by_key(SPREADSHEET_ID)
        except APIError as e:
            last_err = e
            status, _ = _extract_api_error_info(e)
            if _retryable(status):
                time.sleep(2 ** (attempt + 1))
                continue
            raise
        except SpreadsheetNotFound:
            raise
    raise last_err

def clear_cache(): st.cache_data.clear()

def sheets_health_check_or_stop():
    try: return spreadsheet()
    except Exception as e:
        st.error(f"Falha ao acessar a planilha: {e}"); st.stop()

def get_worksheets_map(sh):
    return {w.title: w for w in sh.worksheets()}

def ensure_header(ws_obj, headers):
    try:
        vals = ws_obj.get_all_values()
        if not vals:
            ws_obj.append_row(headers)
            return
        cabecalho_atual = vals[0]
        if cabecalho_atual == headers:
            return
        colunas_faltantes = [h for h in headers if h not in cabecalho_atual]
        if colunas_faltantes:
            ultima_coluna = len(cabecalho_atual)
            for i, col in enumerate(colunas_faltantes):
                ws_obj.update_cell(1, ultima_coluna + i + 1, col)
        else:
            ws_obj.update("1:1", [headers])
    except APIError:
        pass

@st.cache_resource(show_spinner=False)
def ensure_worksheets(_sh):
    wmap = get_worksheets_map(_sh)
    targets = [
        (SHEET_USERS, HEADERS_USERS), (SHEET_CAD, HEADERS_CAD),
        (SHEET_PROD, HEADERS_PROD),   (SHEET_PART, HEADERS_PART),
        (SHEET_VINC, HEADERS_VINC),
        (SHEET_ORIENT, HEADERS_ORIENT),
        (SHEET_ENSINO, HEADERS_ENSINO),
        (SHEET_IMPACTO, HEADERS_IMPACTO),
    ]
    for title, headers in targets:
        if title not in wmap:
            ws_obj = _sh.add_worksheet(title=title, rows=2000, cols=max(12, len(headers)))
            ws_obj.append_row(headers)
            wmap[title] = ws_obj
        else:
            ensure_header(wmap[title], headers)

_sh = sheets_health_check_or_stop()
ensure_worksheets(_sh)

def ws(sheet_name):
    try: return spreadsheet().worksheet(sheet_name)
    except WorksheetNotFound:
        ensure_worksheets(spreadsheet())
        return spreadsheet().worksheet(sheet_name)

@st.cache_data(ttl=60, show_spinner=False)
def read_df(sheet_name):
    last_err = None
    for attempt in range(5):
        try:
            w = ws(sheet_name)
            values = w.get_all_values()
            if not values: return pd.DataFrame()
            return pd.DataFrame(values[1:], columns=values[0]).fillna("")
        except APIError as e:
            last_err = e
            status, _ = _extract_api_error_info(e)
            if _retryable(status):
                time.sleep(2 ** (attempt + 1))
                continue
            break
    return pd.DataFrame()

# ---------------------------------------------------------
# AUTH / USERS
# ---------------------------------------------------------
def users_get(username):
    df = read_df(SHEET_USERS)
    if df.empty or "username" not in df.columns: return None
    m = df["username"].str.lower() == username.lower()
    return df[m].iloc[0].to_dict() if m.any() else None

def authenticate(username, password):
    u = users_get(username)
    if not u: return False, "Usuário inválido."
    if verify_password(password, str(u.get("password_hash", ""))):
        return True, u
    return False, "Senha inválida."

def role_of(user): 
    role = str(user.get("role", "")).strip().lower()
    if role in ["admin", "administrador", "coordenador"]:
        return "admin"
    elif role in ["docente", "professor"]:
        return "docente"
    elif role in ["discente", "aluno", "estudante"]:
        return "discente"
    return role

def listar_docentes():
    df = read_df(SHEET_USERS)
    if df.empty: return []
    return sorted(df[df["role"].str.lower().isin(["docente", "professor"])]["name"].tolist())

def get_docente_username_by_name(nome):
    df = read_df(SHEET_USERS)
    if df.empty: return None
    m = df["name"].str.lower() == nome.lower()
    if m.any():
        return df[m].iloc[0]["username"]
    return None

def hash_password(password, salt=None, iterations=200_000):
    if salt is None: salt = os.urandom(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, iterations)
    return (f"pbkdf2_sha256${iterations}$"
            f"{base64.b64encode(salt).decode()}${base64.b64encode(dk).decode()}")

def verify_password(password, stored):
    try:
        algo, it, sb, hb = stored.split("$")
        if algo != "pbkdf2_sha256": return False
        salt = base64.b64decode(sb)
        exp  = base64.b64decode(hb)
        dk = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, int(it))
        return hmac.compare_digest(dk, exp)
    except Exception:
        return False

# ---------------------------------------------------------
# WORDCLOUD
# ---------------------------------------------------------
def extrair_texto_titulos(df_prod):
    if df_prod.empty or "titulo" not in df_prod.columns:
        return ""
    
    todas_palavras = []
    for titulo in df_prod["titulo"].dropna():
        texto = str(titulo).lower()
        texto = re.sub(r'[^\w\s]', ' ', texto)
        palavras = texto.split()
        for p in palavras:
            p = p.strip()
            if len(p) >= 3 and p not in STOPWORDS and not p.isdigit():
                todas_palavras.append(p)
    
    return " ".join(todas_palavras)

def gerar_wordcloud(texto, max_words=100):
    if not texto:
        return None
    
    wordcloud = WordCloud(
        width=1200,
        height=600,
        background_color='white',
        max_words=max_words,
        colormap='viridis',
        relative_scaling=0.5,
        random_state=42
    ).generate(texto)
    
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.imshow(wordcloud, interpolation='bilinear')
    ax.axis('off')
    plt.tight_layout(pad=0)
    
    return fig

# ---------------------------------------------------------
# BIBLIOMETRIA
# ---------------------------------------------------------
SCOPUS_COLUMNS = {
    "Authors": "autores",
    "Title": "titulo",
    "Year": "ano",
    "Source title": "periodico",
    "Cited by": "citacoes",
    "DOI": "doi",
    "Link": "link",
    "Author Keywords": "palavras_chave",
    "Index Keywords": "palavras_indexadas",
    "Document Type": "tipo_documento",
    "Open Access": "acesso_aberto",
    "Source": "fonte",
    "EID": "eid",
}

def _safe_int_series(series):
    return pd.to_numeric(series, errors="coerce").fillna(0).astype(int)

def calcular_h_index(citacoes):
    valores = sorted([int(c) for c in citacoes if pd.notna(c)], reverse=True)
    h_index = 0
    for i, citacao in enumerate(valores, start=1):
        if citacao >= i:
            h_index = i
        else:
            break
    return h_index

def extrair_metricas_relatorio_autor(texto):
    metricas = {}
    autor = re.search(r"Author:\s*([^,\n\r]+(?:,\s*[^,\n\r]+)?)", texto, flags=re.I)
    h_index = re.search(r"h-index\s*=\s*(\d+)", texto, flags=re.I)
    documentos_h = re.search(r"Of the\s+(\d+)\s+documents considered", texto, flags=re.I)
    
    if autor:
        metricas["Autor"] = autor.group(1).strip()
    if h_index:
        metricas["h-index informado"] = int(h_index.group(1))
    if documentos_h:
        metricas["Documentos considerados no h-index"] = int(documentos_h.group(1))
    
    return metricas

def ler_csv_bibliometrico(uploaded_file):
    raw = uploaded_file.getvalue()
    texto = raw.decode("utf-8-sig", errors="ignore")
    uploaded_file.seek(0)
    
    relatorio_autor = extrair_metricas_relatorio_autor(texto)
    
    try:
        df = pd.read_csv(uploaded_file)
    except Exception:
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file, sep=";")
    
    df = df.dropna(how="all")
    if "Title" not in df.columns and "Year" not in df.columns:
        return pd.DataFrame(), relatorio_autor
    
    colunas_presentes = {c: SCOPUS_COLUMNS[c] for c in SCOPUS_COLUMNS if c in df.columns}
    df = df.rename(columns=colunas_presentes)
    
    for col in ["autores", "titulo", "ano", "periodico", "citacoes", "doi", "link",
                "palavras_chave", "palavras_indexadas", "tipo_documento",
                "acesso_aberto", "fonte", "eid"]:
        if col not in df.columns:
            df[col] = ""
    
    df = df[["autores", "titulo", "ano", "periodico", "citacoes", "doi", "link",
             "palavras_chave", "palavras_indexadas", "tipo_documento",
             "acesso_aberto", "fonte", "eid"]].copy()
    df["ano"] = _safe_int_series(df["ano"])
    df["citacoes"] = _safe_int_series(df["citacoes"])
    df["arquivo_origem"] = uploaded_file.name
    
    return df, relatorio_autor

def consolidar_arquivos_bibliometricos(uploaded_files):
    bases = []
    metricas_relatorio = {}
    
    for arquivo in uploaded_files:
        df, metricas = ler_csv_bibliometrico(arquivo)
        if not df.empty:
            bases.append(df)
        if metricas:
            metricas_relatorio[arquivo.name] = metricas
    
    if not bases:
        return pd.DataFrame(), metricas_relatorio
    
    df_final = pd.concat(bases, ignore_index=True)
    if "eid" in df_final.columns:
        df_final = df_final.drop_duplicates(subset=["eid"], keep="first")
    else:
        df_final = df_final.drop_duplicates(subset=["titulo", "ano"], keep="first")
    
    return df_final, metricas_relatorio

def texto_bibliometrico_para_wordcloud(df):
    if df.empty:
        return ""
    partes = []
    for col in ["titulo", "palavras_chave", "palavras_indexadas"]:
        if col in df.columns:
            partes.extend(df[col].dropna().astype(str).tolist())
    texto = " ".join(partes).lower()
    texto = re.sub(r"[^\w\s]", " ", texto)
    palavras = [
        p for p in texto.split()
        if len(p) >= 3 and p not in STOPWORDS and not p.isdigit()
    ]
    return " ".join(palavras)

# ---------------------------------------------------------
# ESTATÍSTICAS
# ---------------------------------------------------------
def tem_participacao_ppg(producao_id):
    df_part = read_df(SHEET_PART)
    if df_part.empty: return False
    parts = df_part[df_part["producao_id"] == producao_id]
    return any(parts["tipo_participacao"] == "Discente do PPG")

def tem_pesquisador_estrangeiro(producao_id):
    df_part = read_df(SHEET_PART)
    if df_part.empty: return False
    parts = df_part[df_part["producao_id"] == producao_id]
    return any(parts["tipo_participacao"] == "Pesquisador estrangeiro")

def get_estatisticas_avancadas():
    df_prod = read_df(SHEET_PROD)
    df_orient = read_df(SHEET_ORIENT)
    df_ensino = read_df(SHEET_ENSINO)
    df_impacto = read_df(SHEET_IMPACTO)
    
    total_producoes = len(df_prod)
    producoes_por_ano = {}
    producoes_com_ppg_por_ano = {}
    total_com_ppg = 0
    total_com_estrangeiros = 0
    producoes_com_estrangeiros_por_ano = {}
    artigos_com_discente_primeiro = 0
    artigos_com_docente_ultimo = 0
    
    artigos_periodico = df_prod[df_prod["tipo"] == "Artigo em periódico"] if not df_prod.empty else pd.DataFrame()
    periodicos_unicos = artigos_periodico["veiculo"].dropna().unique() if not artigos_periodico.empty else []
    periodicos_unicos = [p for p in periodicos_unicos if str(p).strip()]
    total_periodicos_unicos = len(periodicos_unicos)
    
    if not artigos_periodico.empty:
        top_periodicos = artigos_periodico["veiculo"].value_counts().head(10).to_dict()
    else:
        top_periodicos = {}
    
    for ano in ANOS:
        subset = df_prod[df_prod["ano"].astype(str).str.strip() == ano] if not df_prod.empty else pd.DataFrame()
        producoes_por_ano[ano] = len(subset)
        
        com_ppg = 0
        com_estrangeiros = 0
        
        for _, row in subset.iterrows():
            prod_id = row["id"]
            if tem_participacao_ppg(prod_id):
                com_ppg += 1
            if tem_pesquisador_estrangeiro(prod_id):
                com_estrangeiros += 1
            
            if row.get("tipo") == "Artigo em periódico":
                discente_primeiro = str(row.get("discente_primeiro_autor", "")).strip().lower()
                docente_ultimo = str(row.get("docente_ultimo_autor", "")).strip().lower()
                if discente_primeiro == "sim":
                    artigos_com_discente_primeiro += 1
                if docente_ultimo == "sim":
                    artigos_com_docente_ultimo += 1
        
        producoes_com_ppg_por_ano[ano] = com_ppg
        producoes_com_estrangeiros_por_ano[ano] = com_estrangeiros
        total_com_ppg += com_ppg
        total_com_estrangeiros += com_estrangeiros
    
    total_orientacoes = len(df_orient)
    total_ensino = len(df_ensino)
    total_impacto = len(df_impacto)
    
    orient_por_ano = {}
    for ano in ANOS:
        if not df_orient.empty:
            subset = df_orient[df_orient["ano_inicio"].astype(str).str.strip() == ano]
            orient_por_ano[ano] = len(subset)
        else:
            orient_por_ano[ano] = 0
    
    return {
        "total_producoes": total_producoes,
        "producoes_por_ano": producoes_por_ano,
        "producoes_com_ppg_por_ano": producoes_com_ppg_por_ano,
        "total_com_ppg": total_com_ppg,
        "total_com_estrangeiros": total_com_estrangeiros,
        "producoes_com_estrangeiros_por_ano": producoes_com_estrangeiros_por_ano,
        "total_periodicos_unicos": total_periodicos_unicos,
        "top_periodicos": top_periodicos,
        "artigos_com_discente_primeiro": artigos_com_discente_primeiro,
        "artigos_com_docente_ultimo": artigos_com_docente_ultimo,
        "total_orientacoes": total_orientacoes,
        "total_ensino": total_ensino,
        "total_impacto": total_impacto,
        "orient_por_ano": orient_por_ano,
    }

# ---------------------------------------------------------
# HEADER E NAVEGAÇÃO
# ---------------------------------------------------------
def render_header():
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">🎓 PPG PNBC</h1>
        <p class="header-subtitle">Programa de Pós-Graduação em Ciências Farmacêuticas</p>
    </div>
    """, unsafe_allow_html=True)

def render_nav(current_page="home"):
    nav_items = [
        ("home", "🏠 Início"),
        ("dashboard", "📊 Dashboard"),
        ("producoes", "📚 Produções"),
        ("bibliometria", "📈 Bibliometria"),
        ("orientacoes", "🎓 Orientações"),
        ("ensino", "📖 Ensino"),
        ("impacto", "🌍 Impacto"),
    ]
    
    if st.session_state.get("logged"):
        nav_items.append(("perfil", "👤 Meu Perfil"))
    
    cols = st.columns(len(nav_items))
    for i, (page, label) in enumerate(nav_items):
        with cols[i]:
            if st.button(label, key=f"nav_{page}", use_container_width=True):
                st.session_state["current_page"] = page
                st.rerun()

def render_footer():
    st.markdown("""
    <div class="footer">
        <p class="footer-text">
            © 2026 PPG PNBC — Programa de Pós-Graduação em Ciências Farmacêuticas<br>
            Universidade Federal do Pará (UFPA)
        </p>
    </div>
    """, unsafe_allow_html=True)

# ---------------------------------------------------------
# PÁGINAS
# ---------------------------------------------------------
def page_home():
    st.markdown("""
    <div class="hero">
        <h1 class="hero-title">Bem-vindo ao PPG PNBC</h1>
        <p class="hero-subtitle">
            Acompanhe as produções científicas, orientações, atividades de ensino e impacto social 
            do nosso programa de pós-graduação.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Vídeo
    video_file = Path("videoEntrada.mp4")
    if video_file.exists():
        st.markdown('<div class="video-section">', unsafe_allow_html=True)
        st.markdown('<h2 class="video-title">🎬 Conheça o PPG</h2>', unsafe_allow_html=True)
        col_esq, col_video, col_dir = st.columns([1, 4, 1])
        with col_video:
            st.video(str(video_file), autoplay=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Features
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="section-title">O que você pode encontrar aqui</h2>', unsafe_allow_html=True)
    
    features = [
        ("📊", "Dashboard Interativo", "Visualize estatísticas completas sobre produções, orientações e atividades do programa."),
        ("📚", "Produções Científicas", "Acesse artigos, livros, capítulos e trabalhos publicados pelos docentes e discentes."),
        ("📈", "Avaliação Bibliométrica", "Importe arquivos da Scopus e acompanhe citações, h-index, periódicos e colaboração."),
        ("🎓", "Orientações", "Acompanhe orientações de mestrado e doutorado em andamento e concluídas."),
        ("📖", "Atividades de Ensino", "Consulte disciplinas ministradas, cursos de extensão e workshops."),
        ("🌍", "Impacto Social", "Descubra como o PPG contribui para a sociedade através de projetos e parcerias."),
    ]
    
    cols = st.columns(3)
    for i, (icon, title, desc) in enumerate(features):
        with cols[i % 3]:
            st.markdown(f"""
            <div class="feature-card">
                <div class="feature-icon">{icon}</div>
                <div class="feature-title">{title}</div>
                <div class="feature-desc">{desc}</div>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

def page_dashboard():
    stats = get_estatisticas_avancadas()
    
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="section-title">📊 Dashboard de Produções</h2>', unsafe_allow_html=True)
    
    # Métricas principais
    cols = st.columns(4)
    with cols[0]:
        st.markdown(create_metric_card(stats['total_producoes'], "Total de Produções", "📚"), unsafe_allow_html=True)
    with cols[1]:
        st.markdown(create_metric_card(stats['total_com_ppg'], "Com Discentes PPG", "🎓"), unsafe_allow_html=True)
    with cols[2]:
        st.markdown(create_metric_card(stats['total_com_estrangeiros'], "Cooperação Internacional", "🌍"), unsafe_allow_html=True)
    with cols[3]:
        percent = (stats['total_com_ppg'] / stats['total_producoes'] * 100) if stats['total_producoes'] > 0 else 0
        st.markdown(create_metric_card(f"{percent:.1f}%", "Participação Discente", "📈"), unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Nuvem de palavras
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="section-title">☁️ Nuvem de Palavras</h2>', unsafe_allow_html=True)
    
    df_prod_cloud = read_df(SHEET_PROD)
    texto_titulos = extrair_texto_titulos(df_prod_cloud)
    
    if texto_titulos:
        fig_wordcloud = gerar_wordcloud(texto_titulos, max_words=100)
        if fig_wordcloud:
            st.pyplot(fig_wordcloud)
    else:
        st.info("Ainda não há títulos cadastrados.")
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Gráficos
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="section-title">📈 Evolução Temporal</h2>', unsafe_allow_html=True)
    
    df_total_ano = pd.DataFrame({
        'Ano': list(stats['producoes_por_ano'].keys()),
        'Produções': list(stats['producoes_por_ano'].values())
    })
    st.bar_chart(df_total_ano.set_index('Ano'))
    
    st.markdown('</div>', unsafe_allow_html=True)

def page_producoes():
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="section-title">📚 Produções Científicas</h2>', unsafe_allow_html=True)
    
    df_prod = read_df(SHEET_PROD)
    df_part = read_df(SHEET_PART)
    
    if df_prod.empty:
        st.info("Nenhuma produção cadastrada ainda.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            filtro_ano = st.multiselect("Filtrar por ano", ANOS, default=ANOS, key="filtro_ano_public")
        with col2:
            filtro_tipo = st.multiselect("Filtrar por tipo", TIPOS_PRODUCAO, key="filtro_tipo_public")
        
        df_filtrado = df_prod.copy()
        if filtro_ano:
            df_filtrado = df_filtrado[df_filtrado["ano"].astype(str).str.strip().isin(filtro_ano)]
        if filtro_tipo:
            df_filtrado = df_filtrado[df_filtrado["tipo"].isin(filtro_tipo)]
        
        st.write(f"**Total:** {len(df_filtrado)} produções encontradas")
        
        for ano in ANOS:
            subset = df_filtrado[df_filtrado["ano"].astype(str).str.strip() == ano]
            if not subset.empty:
                st.markdown(f"### 📅 {ano}")
                for _, row in subset.iterrows():
                    with st.expander(f"**{row['titulo']}** — {row['tipo']}"):
                        st.write(f"**Veículo:** {row['veiculo']}")
                        st.write(f"**Autores:** {row['autores']}")
                        st.write(f"**DOI:** {row['doi'] or '—'}")
                        
                        parts = df_part[df_part["producao_id"] == row["id"]] if not df_part.empty else pd.DataFrame()
                        if not parts.empty:
                            st.write("**Participações:**")
                            st.dataframe(parts[["tipo_participacao","nome_participante","vinculo"]], use_container_width=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

def page_bibliometria():
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="section-title">📈 Avaliação Bibliométrica do Docente</h2>', unsafe_allow_html=True)
    
    st.write(
        "Importe arquivos CSV exportados da Scopus. A página também reconhece relatórios de autor "
        "como o arquivo com h-index e usa essas informações como indicadores complementares."
    )
    
    col_docente, col_periodo = st.columns([2, 1])
    with col_docente:
        docente_nome = st.text_input("Docente avaliado", placeholder="Ex.: Givago Silva Souza")
    with col_periodo:
        periodo_quadrienio = st.multiselect(
            "Quadriênio CAPES",
            ANOS,
            default=ANOS,
            help="Use este filtro para destacar a produção do ciclo 2025-2028."
        )
    
    arquivos = st.file_uploader(
        "Arquivos bibliométricos (.csv)",
        type=["csv"],
        accept_multiple_files=True,
        help="Você pode enviar a exportação de documentos da Scopus e o relatório de autor no mesmo campo."
    )
    
    if not arquivos:
        st.info("Envie um ou mais arquivos CSV para iniciar a avaliação bibliométrica.")
        st.markdown('</div>', unsafe_allow_html=True)
        return
    
    try:
        df_biblio, metricas_relatorio = consolidar_arquivos_bibliometricos(arquivos)
    except Exception as e:
        st.error(f"Não foi possível ler os arquivos enviados: {e}")
        st.markdown('</div>', unsafe_allow_html=True)
        return
    
    if df_biblio.empty:
        st.warning("Os arquivos enviados não contêm uma tabela de documentos reconhecida. Se houver relatório de autor, veja os indicadores abaixo.")
        if metricas_relatorio:
            st.json(metricas_relatorio)
        st.markdown('</div>', unsafe_allow_html=True)
        return
    
    anos_disponiveis = sorted([int(a) for a in df_biblio["ano"].dropna().unique() if int(a) > 0], reverse=True)
    tipos_disponiveis = sorted([t for t in df_biblio["tipo_documento"].dropna().unique() if str(t).strip()])
    periodicos_disponiveis = sorted([p for p in df_biblio["periodico"].dropna().unique() if str(p).strip()])
    
    with st.expander("Filtros avançados", expanded=False):
        col1, col2, col3 = st.columns(3)
        with col1:
            filtro_anos = st.multiselect("Ano", anos_disponiveis, default=anos_disponiveis)
        with col2:
            filtro_tipos = st.multiselect("Tipo de documento", tipos_disponiveis)
        with col3:
            filtro_periodicos = st.multiselect("Periódico", periodicos_disponiveis)
    
    df_filtrado = df_biblio.copy()
    if filtro_anos:
        df_filtrado = df_filtrado[df_filtrado["ano"].isin(filtro_anos)]
    if filtro_tipos:
        df_filtrado = df_filtrado[df_filtrado["tipo_documento"].isin(filtro_tipos)]
    if filtro_periodicos:
        df_filtrado = df_filtrado[df_filtrado["periodico"].isin(filtro_periodicos)]
    
    df_quadrienio = df_biblio[df_biblio["ano"].astype(str).isin(periodo_quadrienio)] if periodo_quadrienio else pd.DataFrame()
    total_docs = len(df_filtrado)
    total_citacoes = int(df_filtrado["citacoes"].sum()) if total_docs else 0
    h_index_calculado = calcular_h_index(df_filtrado["citacoes"]) if total_docs else 0
    media_citacoes = (total_citacoes / total_docs) if total_docs else 0
    docs_quadrienio = len(df_quadrienio)
    taxa_oa = (
        df_filtrado["acesso_aberto"].astype(str).str.strip().ne("").mean() * 100
        if total_docs else 0
    )
    
    st.markdown("### Indicadores principais")
    cols = st.columns(6)
    with cols[0]:
        st.markdown(create_metric_card(total_docs, "Documentos", "📄"), unsafe_allow_html=True)
    with cols[1]:
        st.markdown(create_metric_card(total_citacoes, "Citações", "🔎"), unsafe_allow_html=True)
    with cols[2]:
        st.markdown(create_metric_card(h_index_calculado, "h-index calculado", "📊"), unsafe_allow_html=True)
    with cols[3]:
        st.markdown(create_metric_card(f"{media_citacoes:.1f}", "Citações/doc.", "📈"), unsafe_allow_html=True)
    with cols[4]:
        st.markdown(create_metric_card(docs_quadrienio, "Docs. 2025-2028", "🗓️"), unsafe_allow_html=True)
    with cols[5]:
        st.markdown(create_metric_card(f"{taxa_oa:.0f}%", "Open Access", "🔓"), unsafe_allow_html=True)
    
    if metricas_relatorio:
        with st.expander("Indicadores extraídos de relatório de autor", expanded=True):
            linhas = []
            for arquivo, metricas in metricas_relatorio.items():
                for indicador, valor in metricas.items():
                    linhas.append({"arquivo": arquivo, "indicador": indicador, "valor": valor})
            st.dataframe(pd.DataFrame(linhas), use_container_width=True, hide_index=True)
    
    st.markdown("### Evolução e impacto")
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        docs_por_ano = df_filtrado.groupby("ano").size().sort_index()
        st.caption("Documentos por ano")
        st.bar_chart(docs_por_ano)
    with col_g2:
        citacoes_por_ano = df_filtrado.groupby("ano")["citacoes"].sum().sort_index()
        st.caption("Citações acumuladas por ano de publicação")
        st.line_chart(citacoes_por_ano)
    
    col_g3, col_g4 = st.columns(2)
    with col_g3:
        top_periodicos = df_filtrado["periodico"].value_counts().head(10)
        st.caption("Periódicos mais frequentes")
        st.bar_chart(top_periodicos)
    with col_g4:
        tipos = df_filtrado["tipo_documento"].replace("", "Não informado").value_counts()
        st.caption("Tipos de documento")
        st.bar_chart(tipos)
    
    texto_cloud = texto_bibliometrico_para_wordcloud(df_filtrado)
    if texto_cloud:
        st.markdown("### Temas recorrentes")
        fig = gerar_wordcloud(texto_cloud, max_words=120)
        if fig:
            st.pyplot(fig)
    
    st.markdown("### Documentos")
    col_busca, col_download = st.columns([2, 1])
    with col_busca:
        termo = st.text_input("Buscar em título, autores, periódico ou DOI")
    if termo:
        termo_norm = termo.lower().strip()
        mask = (
            df_filtrado["titulo"].astype(str).str.lower().str.contains(termo_norm, na=False) |
            df_filtrado["autores"].astype(str).str.lower().str.contains(termo_norm, na=False) |
            df_filtrado["periodico"].astype(str).str.lower().str.contains(termo_norm, na=False) |
            df_filtrado["doi"].astype(str).str.lower().str.contains(termo_norm, na=False)
        )
        df_filtrado = df_filtrado[mask]
    
    colunas_tabela = ["ano", "titulo", "periodico", "citacoes", "tipo_documento", "doi", "autores"]
    st.dataframe(
        df_filtrado[colunas_tabela].sort_values(["ano", "citacoes"], ascending=[False, False]),
        use_container_width=True,
        hide_index=True
    )
    
    nome_base = re.sub(r"[^a-zA-Z0-9_-]+", "_", docente_nome.strip() or "docente").strip("_").lower()
    with col_download:
        st.download_button(
            "Baixar tabela filtrada",
            data=df_filtrado[colunas_tabela].to_csv(index=False).encode("utf-8-sig"),
            file_name=f"bibliometria_{nome_base}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    st.markdown('</div>', unsafe_allow_html=True)

def page_orientacoes():
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="section-title">🎓 Orientações Acadêmicas</h2>', unsafe_allow_html=True)
    
    df_orient = read_df(SHEET_ORIENT)
    df_users = read_df(SHEET_USERS)
    
    if df_orient.empty:
        st.info("Nenhuma orientação cadastrada ainda.")
    else:
        df_orient = df_orient.sort_values(by="discente_nome", key=lambda x: x.str.lower())
        
        for _, row in df_orient.iterrows():
            docente_user = users_get(row["docente_username"]) if not df_users.empty else None
            docente_nome = docente_user["name"] if docente_user else row["docente_username"]
            
            with st.expander(f"**{row['discente_nome']}** — {row['tipo']} ({row['ano_inicio']} → {row['ano_conclusao'] or 'em andamento'})"):
                st.write(f"**Orientador(a):** {docente_nome}")
                st.write(f"**Status:** {row['status']}")
    
    st.markdown('</div>', unsafe_allow_html=True)

def page_ensino():
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="section-title">📖 Atividades de Ensino</h2>', unsafe_allow_html=True)
    
    df_ensino = read_df(SHEET_ENSINO)
    df_users = read_df(SHEET_USERS)
    
    if df_ensino.empty:
        st.info("Nenhuma atividade de ensino cadastrada ainda.")
    else:
        for _, row in df_ensino.iterrows():
            docente_user = users_get(row["docente_username"]) if not df_users.empty else None
            docente_nome = docente_user["name"] if docente_user else row["docente_username"]
            
            with st.expander(f"**{row['titulo']}** — {row['tipo']} ({row['ano']})"):
                st.write(f"**Docente:** {docente_nome}")
                st.write(f"**Período:** {row['periodo'] or '—'}")
                st.write(f"**Nível:** {row['nivel']}")
                st.write(f"**Carga horária:** {row['carga_horaria'] or '—'} h")
    
    st.markdown('</div>', unsafe_allow_html=True)

def page_impacto():
    st.markdown('<div class="section-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="section-title">🌍 Atividades de Impacto na Sociedade</h2>', unsafe_allow_html=True)
    
    df_impacto = read_df(SHEET_IMPACTO)
    df_users = read_df(SHEET_USERS)
    
    if df_impacto.empty:
        st.info("Nenhuma atividade de impacto cadastrada ainda.")
    else:
        for _, row in df_impacto.iterrows():
            docente_user = users_get(row["docente_username"]) if not df_users.empty else None
            docente_nome = docente_user["name"] if docente_user else row["docente_username"]
            
            with st.expander(f"**{row['titulo']}** — {row['tipo']}"):
                st.write(f"**Docente:** {docente_nome}")
                st.write(f"**Data:** {row['data'] or '—'}")
                st.write(f"**Local:** {row['local'] or '—'}")
                st.write(f"**Público-alvo:** {row['publico_alvo'] or '—'}")
    
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
if "current_page" not in st.session_state:
    st.session_state["current_page"] = "home"

render_header()
render_nav(st.session_state["current_page"])

# Roteamento de páginas
if st.session_state["current_page"] == "home":
    page_home()
elif st.session_state["current_page"] == "dashboard":
    page_dashboard()
elif st.session_state["current_page"] == "producoes":
    page_producoes()
elif st.session_state["current_page"] == "bibliometria":
    page_bibliometria()
elif st.session_state["current_page"] == "orientacoes":
    page_orientacoes()
elif st.session_state["current_page"] == "ensino":
    page_ensino()
elif st.session_state["current_page"] == "impacto":
    page_impacto()

render_footer()
