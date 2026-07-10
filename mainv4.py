# app.py — Sistema de Monitoramento de Produção do PPG (v6.9 - Vídeo Centralizado)
# Streamlit + Google Sheets + E-mails
# =========================================================

import os, time, base64, uuid, hashlib, hmac, smtplib, re, io, csv
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
    page_title="PPG — Monitor de Produção",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

try:
    banner = Image.open("banner_ppg.png")
    st.image(banner, use_container_width=True)
except Exception:
    pass

st.title("🦠Sistema de Gerenciamento da Produção do PPG em Neurociências e Biologia Celular da UFPA 🧠")

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
# STOPWORDS (Português)
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
    "model", "application", "evaluation", "effect", "impact", "role", "role",
}

# ---------------------------------------------------------
# CSS
# ---------------------------------------------------------
st.markdown("""
<style>
.block-card { background:#f7f7f9; border:1px solid #e7e7ee; padding:14px; border-radius:10px; margin-bottom:10px;}
.descricao-box { background:#fff8e1; border-left:4px solid #ffb300; padding:10px 14px;
                 border-radius:6px; margin:10px 0; font-size:0.95rem; white-space:pre-wrap;}
.coautor-badge { display:inline-block; background:#e3f2fd; color:#1976d2; 
                 padding:4px 10px; border-radius:12px; font-size:0.85rem; margin:2px;}
.autor-principal-badge { display:inline-block; background:#c8e6c9; color:#388e3c; 
                         padding:4px 10px; border-radius:12px; font-size:0.85rem; margin:2px;}
.main-author-tag { background:#f3e5f5; color:#7b1fa2; padding:6px 12px; border-radius:8px; 
                   font-size:0.9rem; margin-bottom:8px; display:inline-block; font-weight:600;}
.metric-card { background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
               color:white; padding:20px; border-radius:12px; text-align:center; margin:10px 0;}
.metric-card-blue { background:linear-gradient(135deg, #2196f3 0%, #21cbf3 100%); 
                    color:white; padding:20px; border-radius:12px; text-align:center; margin:10px 0;}
.metric-card-green { background:linear-gradient(135deg, #11998e 0%, #38ef7d 100%); 
                     color:white; padding:20px; border-radius:12px; text-align:center; margin:10px 0;}
.metric-card-orange { background:linear-gradient(135deg, #f46b45 0%, #eea849 100%); 
                      color:white; padding:20px; border-radius:12px; text-align:center; margin:10px 0;}
.metric-card-pink { background:linear-gradient(135deg, #ee0979 0%, #ff6a00 100%); 
                    color:white; padding:20px; border-radius:12px; text-align:center; margin:10px 0;}
.metric-value { font-size:2.5rem; font-weight:bold; }
.metric-label { font-size:0.9rem; opacity:0.95; }
.public-notice { background:#e3f2fd; border-left:4px solid #2196f3; padding:12px; 
                 border-radius:6px; margin:15px 0; }
.highlight-box { background:#f5f5f5; border:1px solid #ddd; padding:15px; border-radius:8px; margin:10px 0;}
.badge-discente-1 { display:inline-block; background:#e1f5fe; color:#0277bd; 
                    padding:4px 10px; border-radius:12px; font-size:0.8rem; margin:2px; font-weight:600;}
.badge-docente-last { display:inline-block; background:#e8f5e9; color:#2e7d32; 
                      padding:4px 10px; border-radius:12px; font-size:0.8rem; margin:2px; font-weight:600;}
.autoria-section { background:#fff3e0; border:1px solid #ffe0b2; padding:12px; border-radius:8px; margin:10px 0;}
.status-andamento { background:#fff3cd; color:#856404; padding:3px 8px; border-radius:10px; font-size:0.8rem;}
.status-concluida { background:#d4edda; color:#155724; padding:3px 8px; border-radius:10px; font-size:0.8rem;}
.status-outro { background:#f8d7da; color:#721c24; padding:3px 8px; border-radius:10px; font-size:0.8rem;}

/* 🎬 Vídeo de entrada - LARGURA CONTROLADA E CENTRALIZADO */
.video-container {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    border-radius: 16px;
    padding: 20px;
    margin: 20px auto;
    max-width: 800px;
    box-shadow: 0 8px 16px rgba(0,0,0,0.1);
}
.video-wrapper {
    position: relative;
    padding-bottom: 56.25%;
    height: 0;
    overflow: hidden;
    border-radius: 12px;
    background: #000;
    max-width: 100%;
}
.video-wrapper video,
.video-wrapper iframe {
    position: absolute;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    border-radius: 12px;
}
.video-title {
    color: white;
    text-align: center;
    margin-bottom: 15px;
    font-size: 1.5rem;
    font-weight: 600;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 🎬 FUNÇÃO PARA EXIBIR VÍDEO (CENTRALIZADO)
# ---------------------------------------------------------
def exibir_video_entrada():
    """
    Exibe o vídeo com autoplay e loop contínuo.
    """
    video_file = Path("videoEntrada.mp4")
    onedrive_url = "https://1drv.ms/v/c/58f7c307dd0b40d5/IQA3PnOTq7oOSaSa-iZ9QFrpAWog0XjOwi8u-qlM0lf5IuE?e=rOp0G2"
    
    col_esq, col_video, col_dir = st.columns([1, 8, 1])
    
    with col_video:        
        if video_file.exists():
            try:
                video_bytes = video_file.read_bytes()
                b64_video = base64.b64encode(video_bytes).decode()
                
                video_html = f"""
                <video autoplay loop muted playsinline controls
                       style="width: 100%; border-radius: 12px; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
                    <source src="data:video/mp4;base64,{b64_video}" type="video/mp4">
                    Seu navegador não suporta o elemento de vídeo.
                </video>
                """
                components.html(video_html, height=450, scrolling=False)
            except Exception as e:
                st.warning(f"Erro ao carregar vídeo: {e}")
                st.video(str(video_file), autoplay=True)
        else:
            iframe_html = f"""
            <iframe src="{onedrive_url}" 
                    width="100%" 
                    height="450" 
                    frameborder="0" 
                    allowfullscreen
                    style="border-radius: 12px; border: none;">
            </iframe>
            """
            components.html(iframe_html, height=450, scrolling=False)

# ---------------------------------------------------------
# E-MAIL & PASSWORD HASH
# ---------------------------------------------------------
def send_email(to_email, subject, body):
    if "EMAIL" not in st.secrets: return False
    try:
        cfg = st.secrets["EMAIL"]
        msg = EmailMessage()
        msg["Subject"], msg["From"], msg["To"] = subject, cfg["FROM"], to_email
        msg.set_content(body)
        with smtplib.SMTP(cfg["SMTP_HOST"], int(cfg["SMTP_PORT"]), timeout=20) as s:
            s.ehlo(); s.starttls(); s.login(cfg["SMTP_USER"], cfg["SMTP_PASS"])
            s.send_message(msg)
        return True
    except Exception as e:
        st.warning(f"Falha no e-mail: {e}")
        return False

def send_admin_email(subject, body):
    if "EMAIL" not in st.secrets: return False
    return send_email(st.secrets["EMAIL"]["ADMIN_TO"], subject, body)

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

# ---------------------------------------------------------
# ☁️ NUVEM DE PALAVRAS (WordCloud)
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
# 📈 BIBLIOMETRIA (Scopus / relatório de autor)
# ---------------------------------------------------------
SCOPUS_COLUMNS = {
    "Authors": "autores",
    "Author full names": "autores_completos",
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

DOCENTES_BIBLIOMETRIA = {
    "Givago da Silva Souza": [
        "work/GSouza.csv",
        "work/scopus.csv",
    ],
}

class LocalCsvFile(io.BytesIO):
    def __init__(self, path):
        self.path = Path(path)
        super().__init__(self.path.read_bytes())
        self.name = self.path.name

def resolver_caminho_bibliometrico(caminho):
    caminho = Path(caminho)
    candidatos = [
        caminho,
        Path.cwd() / caminho,
        Path(__file__).resolve().parent / caminho,
        Path(__file__).resolve().parent.parent / caminho,
    ]
    for candidato in candidatos:
        if candidato.exists():
            return candidato
    return None

def carregar_arquivos_docente_bibliometria(nome_docente):
    arquivos = []
    for caminho in DOCENTES_BIBLIOMETRIA.get(nome_docente, []):
        encontrado = resolver_caminho_bibliometrico(caminho)
        if encontrado:
            arquivos.append(LocalCsvFile(encontrado))
    return arquivos

def configurar_eixos_pretos(ax):
    ax.tick_params(axis="both", colors="black", labelsize=10)
    ax.xaxis.label.set_color("black")
    ax.yaxis.label.set_color("black")
    ax.title.set_color("black")
    for spine in ax.spines.values():
        spine.set_color("black")

def grafico_evolucao_scopus(df, citacoes_anuais=None):
    producoes_ano = df[df["ano"] > 0].groupby("ano").size().sort_index()
    citacoes_anuais = citacoes_anuais if citacoes_anuais is not None else pd.Series(dtype=int)
    
    if producoes_ano.empty and citacoes_anuais.empty:
        return None
    
    anos = sorted(set(producoes_ano.index.astype(int).tolist()) | set(citacoes_anuais.index.astype(int).tolist()))
    producoes = pd.Series(0, index=anos)
    citacoes = pd.Series(0, index=anos)
    
    if not producoes_ano.empty:
        producoes.loc[producoes_ano.index.astype(int)] = producoes_ano.astype(int).values
    if not citacoes_anuais.empty:
        citacoes.loc[citacoes_anuais.index.astype(int)] = citacoes_anuais.astype(int).values
    
    fig, ax1 = plt.subplots(figsize=(12, 5))
    ax2 = ax1.twinx()
    
    eixo_x = [str(a) for a in anos]
    ax1.bar(eixo_x, producoes.values, color="#4f7ccf", alpha=0.88, label="Produções")
    ax2.plot(
        eixo_x,
        citacoes.values,
        color="#f28e2b",
        marker="o",
        linewidth=2.5,
        label="Citações recebidas"
    )
    
    ax1.set_title("Evolução anual de produções e citações", fontsize=14, fontweight="bold", color="black")
    ax1.set_xlabel("Ano", color="black")
    ax1.set_ylabel("Produções", color="black")
    ax2.set_ylabel("Citações recebidas no ano", color="black")
    ax1.grid(axis="y", alpha=0.25)
    ax1.tick_params(axis="x", rotation=45)
    
    configurar_eixos_pretos(ax1)
    configurar_eixos_pretos(ax2)
    ax2.tick_params(axis="y", colors="black", labelsize=10)
    
    linhas1, labels1 = ax1.get_legend_handles_labels()
    linhas2, labels2 = ax2.get_legend_handles_labels()
    legenda = ax1.legend(linhas1 + linhas2, labels1 + labels2, loc="upper left", frameon=False)
    for texto in legenda.get_texts():
        texto.set_color("black")
    
    fig.tight_layout()
    return fig

def grafico_barras_horizontal(series, titulo, xlabel):
    dados = series.dropna()
    if dados.empty:
        return None
    dados = dados.sort_values(ascending=True)
    
    fig, ax = plt.subplots(figsize=(10, max(4, len(dados) * 0.42)))
    ax.barh(dados.index.astype(str), dados.values, color="#4f7ccf", alpha=0.9)
    ax.set_title(titulo, fontsize=13, fontweight="bold", color="black")
    ax.set_xlabel(xlabel, color="black")
    ax.grid(axis="x", alpha=0.25)
    configurar_eixos_pretos(ax)
    fig.tight_layout()
    return fig

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

def extrair_citacoes_relatorio_scopus(texto):
    rows = list(csv.reader(io.StringIO(texto)))
    header_idx = None
    for i, row in enumerate(rows):
        if len(row) > 7 and row[0].strip().lower() == "publication year" and row[1].strip().lower() == "document title":
            header_idx = i
            break
    
    if header_idx is None or header_idx == 0:
        return pd.Series(dtype=int)
    
    header = rows[header_idx - 1]
    resumo = rows[header_idx]
    citacoes_por_ano = {}
    
    for idx, label in enumerate(header):
        label = str(label).strip()
        if not re.fullmatch(r"\d{4}", label):
            continue
        if idx >= len(resumo):
            continue
        valor = pd.to_numeric(str(resumo[idx]).strip(), errors="coerce")
        if pd.notna(valor):
            citacoes_por_ano[int(label)] = int(valor)
    
    if not citacoes_por_ano:
        return pd.Series(dtype=int)
    
    return pd.Series(citacoes_por_ano, name="citacoes").sort_index()

def obter_citacoes_anuais_de_relatorios(arquivos):
    series = []
    for arquivo in arquivos:
        try:
            arquivo.seek(0)
            texto = arquivo.getvalue().decode("utf-8-sig", errors="ignore")
            s = extrair_citacoes_relatorio_scopus(texto)
            if not s.empty:
                series.append(s)
        except Exception:
            continue
    
    for arquivo in arquivos:
        try:
            arquivo.seek(0)
        except Exception:
            pass
    
    if not series:
        return pd.Series(dtype=int)
    
    return pd.concat(series, axis=1).fillna(0).sum(axis=1).astype(int).sort_index()

def ler_csv_bibliometrico(uploaded_file):
    raw = uploaded_file.getvalue()
    texto = raw.decode("utf-8-sig", errors="ignore")
    metricas_relatorio = extrair_metricas_relatorio_autor(texto)
    
    uploaded_file.seek(0)
    try:
        df = pd.read_csv(uploaded_file)
    except Exception:
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file, sep=";")
    
    df = df.dropna(how="all")
    if "Title" not in df.columns and "Year" not in df.columns:
        return pd.DataFrame(), metricas_relatorio
    
    colunas_presentes = {c: SCOPUS_COLUMNS[c] for c in SCOPUS_COLUMNS if c in df.columns}
    df = df.rename(columns=colunas_presentes)
    
    for col in ["autores", "autores_completos", "titulo", "ano", "periodico", "citacoes",
                "doi", "link", "palavras_chave", "palavras_indexadas", "tipo_documento",
                "acesso_aberto", "fonte", "eid"]:
        if col not in df.columns:
            df[col] = ""
    
    colunas = ["autores", "autores_completos", "titulo", "ano", "periodico", "citacoes",
               "doi", "link", "palavras_chave", "palavras_indexadas", "tipo_documento",
               "acesso_aberto", "fonte", "eid"]
    df = df[colunas].copy()
    df["ano"] = _safe_int_series(df["ano"])
    df["citacoes"] = _safe_int_series(df["citacoes"])
    df["arquivo_origem"] = uploaded_file.name
    
    return df, metricas_relatorio

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
    if "eid" in df_final.columns and df_final["eid"].astype(str).str.strip().any():
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

def render_metric_card(valor, rotulo, classe="metric-card"):
    st.markdown(f"""
    <div class="{classe}">
        <div class="metric-value">{valor}</div>
        <div class="metric-label">{rotulo}</div>
    </div>""", unsafe_allow_html=True)

def render_bibliometria_docente():
    st.subheader("📈 Produtividade bibliométrica do docente")
    
    docentes = list(DOCENTES_BIBLIOMETRIA.keys())
    docente_nome = st.selectbox("Docente", docentes, key="biblio_docente_select")
    arquivos = carregar_arquivos_docente_bibliometria(docente_nome)
    
    with st.expander("📎 Arquivos associados", expanded=True):
        if arquivos:
            st.write(f"**{docente_nome}**")
            st.caption("Arquivos carregados automaticamente para este docente:")
            for arquivo in arquivos:
                st.write(f"• {arquivo.name}")
        else:
            st.warning("Nenhum arquivo associado foi encontrado para este docente.")
        
        arquivos_extra = st.file_uploader(
            "Adicionar CSVs temporários para visualização",
            type=["csv"],
            accept_multiple_files=True,
            key="biblio_upload_extra",
            help="Use apenas para testes. Para deixar permanente, adicione o caminho no dicionário DOCENTES_BIBLIOMETRIA."
        )
        if arquivos_extra:
            arquivos.extend(arquivos_extra)
    
    if not arquivos:
        st.info("Associe arquivos CSV da Scopus ao docente para visualizar a produtividade bibliométrica.")
        return
    
    try:
        citacoes_anuais_relatorio = obter_citacoes_anuais_de_relatorios(arquivos)
        df_biblio, metricas_relatorio = consolidar_arquivos_bibliometricos(arquivos)
    except Exception as e:
        st.error(f"Não foi possível ler os arquivos associados: {e}")
        return
    
    if df_biblio.empty:
        st.warning("Os arquivos associados não contêm uma tabela de documentos reconhecida.")
        if metricas_relatorio:
            st.write("Indicadores extraídos:")
            st.json(metricas_relatorio)
        return
    
    anos_disponiveis = sorted([int(a) for a in df_biblio["ano"].dropna().unique() if int(a) > 0], reverse=True)
    tipos_disponiveis = sorted([t for t in df_biblio["tipo_documento"].dropna().unique() if str(t).strip()])
    
    with st.expander("🔎 Filtros de visualização", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            filtro_anos = st.multiselect("Ano de publicação", anos_disponiveis, default=anos_disponiveis, key="biblio_f_anos")
        with col2:
            filtro_tipos = st.multiselect("Tipo de documento", tipos_disponiveis, key="biblio_f_tipos")
    
    df_filtrado = df_biblio.copy()
    if filtro_anos:
        df_filtrado = df_filtrado[df_filtrado["ano"].isin(filtro_anos)]
    if filtro_tipos:
        df_filtrado = df_filtrado[df_filtrado["tipo_documento"].isin(filtro_tipos)]
    
    df_quadrienio = df_biblio[df_biblio["ano"].astype(str).isin(ANOS)]
    total_docs = len(df_filtrado)
    citacoes_usam_relatorio = not citacoes_anuais_relatorio.empty
    total_citacoes = int(citacoes_anuais_relatorio.sum()) if citacoes_usam_relatorio else (int(df_filtrado["citacoes"].sum()) if total_docs else 0)
    h_index_calculado = calcular_h_index(df_filtrado["citacoes"]) if total_docs else 0
    media_citacoes = (total_citacoes / total_docs) if total_docs else 0
    docs_quadrienio = len(df_quadrienio)
    taxa_oa = (
        df_filtrado["acesso_aberto"].astype(str).str.strip().ne("").mean() * 100
        if total_docs else 0
    )
    
    st.markdown(f"### {docente_nome}")
    st.caption(
        "Painel de visualização da produtividade do docente. Quando houver relatório do tipo GSouza.csv, "
        "a evolução das citações usa as citações recebidas em cada ano desse relatório; a produção anual usa a exportação de documentos da Scopus."
    )
    
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1: render_metric_card(total_docs, "Documentos")
    with c2: render_metric_card(total_citacoes, "Citações", "metric-card-blue")
    with c3: render_metric_card(h_index_calculado, "h-index calculado", "metric-card-green")
    with c4: render_metric_card(f"{media_citacoes:.1f}", "Citações/doc.", "metric-card-orange")
    with c5: render_metric_card(docs_quadrienio, "Docs. 2025-2028", "metric-card-pink")
    with c6: render_metric_card(f"{taxa_oa:.0f}%", "Open Access", "metric-card-blue")
    
    if citacoes_usam_relatorio:
        st.success("Fonte das citações anuais: relatório de citações da Scopus associado ao docente.")
    else:
        st.warning("Relatório anual de citações não encontrado; usando a coluna 'Cited by' da exportação de documentos como alternativa.")
    
    st.divider()
    st.markdown("### Evolução no estilo Scopus")
    fig_evolucao = grafico_evolucao_scopus(df_filtrado, citacoes_anuais_relatorio)
    if fig_evolucao:
        st.pyplot(fig_evolucao)
        plt.close(fig_evolucao)
    else:
        st.info("Não há anos válidos para gerar a evolução bibliométrica.")
    
    st.divider()
    top_periodicos = df_filtrado["periodico"].replace("", "Não informado").value_counts().head(10)
    fig_periodicos = grafico_barras_horizontal(top_periodicos, "Top 10 periódicos", "Documentos")
    if fig_periodicos:
        st.pyplot(fig_periodicos)
        plt.close(fig_periodicos)
    
    texto_cloud = texto_bibliometrico_para_wordcloud(df_filtrado)
    if texto_cloud:
        st.divider()
        st.markdown("### ☁️ Temas recorrentes")
        fig = gerar_wordcloud(texto_cloud, max_words=120)
        if fig:
            st.pyplot(fig)
            plt.close(fig)
    
    st.divider()
    st.markdown("### Documentos do docente")
    termo = st.text_input("Buscar em título, autores, periódico ou DOI", key="biblio_busca")
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
    df_tabela = df_filtrado[colunas_tabela].sort_values(["ano", "citacoes"], ascending=[False, False])
    st.dataframe(df_tabela, use_container_width=True, hide_index=True)

# ---------------------------------------------------------
# FUNÇÕES DE CO-AUTORIA
# ---------------------------------------------------------
def verificar_duplicacao(doi, titulo):
    df = read_df(SHEET_PROD)
    if df.empty: return None
    if doi and doi.strip():
        doi_clean = doi.strip().lower()
        match = df[df["doi"].str.strip().str.lower() == doi_clean]
        if not match.empty: return match.iloc[0]
    if titulo and titulo.strip():
        titulo_clean = titulo.strip().lower()
        df["titulo_norm"] = df["titulo"].str.lower().str.strip()
        match = df[df["titulo_norm"] == titulo_clean]
        if not match.empty: return match.iloc[0]
    return None

def adicionar_co_autor(producao_id, username):
    df = read_df(SHEET_PROD)
    idx = df.index[df["id"] == producao_id]
    if len(idx) == 0: return False, "Produção não encontrada"
    i = int(idx[0])
    if "co_autores" not in df.columns:
        return False, "Coluna co_autores não existe na planilha."
    co_autores_atuais = str(df.loc[i, "co_autores"]).strip() if pd.notna(df.loc[i, "co_autores"]) else ""
    if username in co_autores_atuais.split(","):
        return False, "Você já é co-autor desta produção"
    novos_co_autores = f"{co_autores_atuais},{username}" if co_autores_atuais else username
    w = ws(SHEET_PROD)
    row_number = i + 2
    col_map = {h: (j + 1) for j, h in enumerate(df.columns)}
    if "co_autores" in col_map:
        w.update_cell(row_number, col_map["co_autores"], novos_co_autores)
        clear_cache()
        return True, "✅ Co-autoria adicionada com sucesso!"
    return False, "Erro ao adicionar co-autor"

def get_minhas_producoes(username):
    df = read_df(SHEET_PROD)
    if df.empty: return pd.DataFrame()
    principal = df[df["docente_username"] == username].copy()
    if not principal.empty: principal["tipo_autoria"] = "principal"
    coautor = pd.DataFrame()
    if "co_autores" in df.columns:
        username_clean = username.strip().lower()
        df["co_autores_clean"] = df["co_autores"].fillna("").str.lower().str.strip()
        pattern = rf"(^|,)\s*{re.escape(username_clean)}\s*(,|$)"
        coautor = df[df["co_autores_clean"].str.contains(pattern, regex=True, na=False)].copy()
        if not coautor.empty: coautor["tipo_autoria"] = "coautor"
    if not principal.empty and not coautor.empty:
        todas = pd.concat([principal, coautor]).drop_duplicates(subset=["id"])
    elif not principal.empty: todas = principal
    elif not coautor.empty: todas = coautor
    else: todas = pd.DataFrame()
    return todas

def get_nome_autor_principal(username):
    user_data = users_get(username)
    return user_data["name"] if user_data else username

# ---------------------------------------------------------
# CRUD DE ORIENTAÇÕES
# ---------------------------------------------------------
def orientacao_submit(docente_username, discente_nome, tipo, ano_inicio, ano_conclusao, status):
    orient_id = str(uuid.uuid4())
    ws(SHEET_ORIENT).append_row([
        orient_id, docente_username, discente_nome.strip(), tipo,
        str(ano_inicio).strip(), str(ano_conclusao).strip(), status,
        datetime.utcnow().isoformat(timespec="seconds")
    ])
    clear_cache()
    return orient_id

def orientacao_delete(orient_id):
    df = read_df(SHEET_ORIENT)
    idx = df.index[df["id"] == orient_id]
    if len(idx) == 0: return False, "Orientação não encontrada."
    i = int(idx[0])
    w = ws(SHEET_ORIENT)
    try:
        w.delete_rows(i + 2)
        clear_cache()
        return True, "Orientação excluída!"
    except Exception as e:
        return False, f"Erro ao excluir: {e}"

# ---------------------------------------------------------
# CRUD DE ATIVIDADES DE ENSINO
# ---------------------------------------------------------
def ensino_submit(docente_username, tipo, titulo, periodo, ano, 
                  carga_horaria, nivel, descricao=""):
    ensino_id = str(uuid.uuid4())
    ws(SHEET_ENSINO).append_row([
        ensino_id, docente_username, tipo, titulo.strip(), periodo.strip(),
        str(ano), str(carga_horaria), nivel, descricao.strip(),
        datetime.utcnow().isoformat(timespec="seconds")
    ])
    clear_cache()
    return ensino_id

def ensino_delete(ensino_id):
    df = read_df(SHEET_ENSINO)
    idx = df.index[df["id"] == ensino_id]
    if len(idx) == 0: return False, "Atividade não encontrada."
    i = int(idx[0])
    w = ws(SHEET_ENSINO)
    try:
        w.delete_rows(i + 2)
        clear_cache()
        return True, "Atividade excluída!"
    except Exception as e:
        return False, f"Erro ao excluir: {e}"

# ---------------------------------------------------------
# CRUD DE ATIVIDADES DE IMPACTO
# ---------------------------------------------------------
def impacto_submit(docente_username, tipo, titulo, descricao, data, 
                   publico_alvo, local):
    impacto_id = str(uuid.uuid4())
    ws(SHEET_IMPACTO).append_row([
        impacto_id, docente_username, tipo, titulo.strip(), descricao.strip(),
        data, publico_alvo.strip(), local.strip(),
        datetime.utcnow().isoformat(timespec="seconds")
    ])
    clear_cache()
    return impacto_id

def impacto_delete(impacto_id):
    df = read_df(SHEET_IMPACTO)
    idx = df.index[df["id"] == impacto_id]
    if len(idx) == 0: return False, "Atividade não encontrada."
    i = int(idx[0])
    w = ws(SHEET_IMPACTO)
    try:
        w.delete_rows(i + 2)
        clear_cache()
        return True, "Atividade excluída!"
    except Exception as e:
        return False, f"Erro ao excluir: {e}"

# ---------------------------------------------------------
# ANÁLISE DE AUTORIA (PRODUÇÕES)
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
# CADASTRO & PRODUÇÕES (CRUD)
# ---------------------------------------------------------
def cadastro_submit(name, username, email, password, role, orientador=""):
    if users_get(username): return False, "Username já existe."
    df_cad = read_df(SHEET_CAD)
    if not df_cad.empty:
        m = (df_cad["username"].str.lower() == username.lower()) & (df_cad["status"] == "Pendente")
        if m.any(): return False, "Já existe cadastro pendente."
    req_id = str(uuid.uuid4())
    created = datetime.utcnow().isoformat(timespec="seconds")
    ws(SHEET_CAD).append_row([req_id, name.strip(), username.strip(), email.strip(),
        role, orientador.strip(), hash_password(password), "Pendente", created, "", "", ""])
    clear_cache()
    send_admin_email(subject=f"Novo cadastro PPG ({role})",
        body=f"Nome: {name}\nUsername: {username}\nEmail: {email}\nPerfil: {role}\nOrientador: {orientador or '—'}\nID: {req_id}")
    return True, "Cadastro enviado para aprovação."

def cadastro_review(req_id, action, admin_username, reason=""):
    w = ws(SHEET_CAD)
    vals = w.get_all_values()
    if len(vals) < 2: return False, "Sem solicitações."
    df = pd.DataFrame(vals[1:], columns=vals[0]).fillna("")
    idx = df.index[df["id"] == req_id]
    if len(idx) == 0: return False, "Não encontrada."
    i = int(idx[0])
    status = "Aprovado" if action == "Aprovar" else "Rejeitado"
    reviewed = datetime.utcnow().isoformat(timespec="seconds")
    df.loc[i, "status"] = status; df.loc[i, "reviewed_at"] = reviewed
    df.loc[i, "reviewed_by"] = admin_username; df.loc[i, "review_reason"] = reason
    if action == "Aprovar":
        ws(SHEET_USERS).append_row([df.loc[i, "username"], df.loc[i, "name"], df.loc[i, "email"],
            df.loc[i, "role"], df.loc[i, "orientador"], df.loc[i, "password_hash"],
            datetime.utcnow().isoformat(timespec="seconds")])
    row_number = i + 2
    col_map = {h: (j + 1) for j, h in enumerate(vals[0])}
    for col in ["status", "reviewed_at", "reviewed_by", "review_reason"]:
        if col in col_map: w.update_cell(row_number, col_map[col], str(df.loc[i, col]))
    clear_cache()
    email_user = str(df.loc[i, "email"]).strip()
    if email_user:
        send_email(email_user, f"Cadastro {status} — PPG",
                   f"Olá, {df.loc[i,'name']}!\n\nSeu cadastro foi: {status}.\nMotivo: {reason or '—'}")
    return True, f"Solicitação {status.lower()}."

def producao_submit(docente_username, titulo, tipo, ano, veiculo, autores, doi, 
                    descricao="", co_autores="", discente_primeiro="Não", docente_ultimo="Não"):
    prod_id = str(uuid.uuid4())
    ws(SHEET_PROD).append_row([
        prod_id, docente_username, titulo.strip(), tipo, str(ano),
        veiculo.strip(), autores.strip(), doi.strip(), descricao.strip(), 
        co_autores.strip(), discente_primeiro, docente_ultimo,
        datetime.utcnow().isoformat(timespec="seconds")
    ])
    clear_cache()
    return prod_id

def producao_update(producao_id, titulo, tipo, ano, veiculo, autores, doi, 
                    descricao="", co_autores="", discente_primeiro="Não", docente_ultimo="Não"):
    df = read_df(SHEET_PROD)
    idx = df.index[df["id"] == producao_id]
    if len(idx) == 0: return False, "Produção não encontrada."
    i = int(idx[0])
    w = ws(SHEET_PROD)
    row_number = i + 2
    range_name = f"C{row_number}:L{row_number}"
    values = [[titulo.strip(), tipo, str(ano), veiculo.strip(), autores.strip(), 
               doi.strip(), descricao.strip(), co_autores.strip(), 
               discente_primeiro, docente_ultimo]]
    w.update(range_name, values)
    clear_cache()
    return True, "Produção atualizada com sucesso!"

def producao_delete(producao_id):
    df_prod = read_df(SHEET_PROD)
    idx_prod = df_prod.index[df_prod["id"] == producao_id]
    if len(idx_prod) == 0: return False, "Produção não encontrada."
    df_part = read_df(SHEET_PART)
    if not df_part.empty:
        idx_parts = df_part.index[df_part["producao_id"] == producao_id]
        if len(idx_parts) > 0:
            w_part = ws(SHEET_PART)
            for i in sorted(idx_parts, reverse=True):
                row_number = int(i) + 2
                w_part.delete_rows(row_number)
    i = int(idx_prod[0])
    w_prod = ws(SHEET_PROD)
    row_number = i + 2
    try:
        w_prod.delete_rows(row_number)
        clear_cache()
        return True, "Produção e participações vinculadas excluídas!"
    except Exception as e:
        return False, f"Erro ao excluir: {e}"

def participacao_submit(producao_id, tipo, nome, vinculo=""):
    ws(SHEET_PART).append_row([str(uuid.uuid4()), producao_id, tipo, nome.strip(), vinculo.strip(),
        datetime.utcnow().isoformat(timespec="seconds")])
    clear_cache()

def vinculo_submit(discente_username, orientador_username, producao_id):
    ws(SHEET_VINC).append_row([str(uuid.uuid4()), discente_username, orientador_username, producao_id,
        datetime.utcnow().isoformat(timespec="seconds")])
    clear_cache()

# ---------------------------------------------------------
# CHECKBOXES DE AUTORIA
# ---------------------------------------------------------
def renderizar_checkboxes_autoria(prefixo, discente_primeiro_default=False, docente_ultimo_default=False):
    st.markdown("""
    <div class="autoria-section">
        <strong>✍️ Análise de Autoria</strong><br>
        <small>Marque as opções que se aplicam a esta produção para enriquecer as estatísticas do programa.</small>
    </div>
    """, unsafe_allow_html=True)
    
    col_check1, col_check2 = st.columns(2)
    with col_check1:
        discente_primeiro = st.checkbox(
            "🎓 Discente do PPG é o primeiro/último autor",
            value=discente_primeiro_default,
            key=f"{prefixo}_discente_primeiro",
            help="Marque se o primeiro/último nome na lista de autores é um discente do PPG"
        )
    with col_check2:
        docente_ultimo = st.checkbox(
            "👨‍🏫 Docente do PPG é o primeiro/último autor",
            value=docente_ultimo_default,
            key=f"{prefixo}_docente_ultimo",
            help="Marque se o primeiro/último nome na lista de autores é um docente do PPG"
        )
    
    return ("Sim" if discente_primeiro else "Não", 
            "Sim" if docente_ultimo else "Não")

# ---------------------------------------------------------
# BADGE DE STATUS
# ---------------------------------------------------------
def badge_status(status):
    status_lower = str(status).strip().lower()
    if status_lower in ["em andamento"]:
        return f'<span class="status-andamento">⏳ {status}</span>'
    elif status_lower in ["concluída", "concluida"]:
        return f'<span class="status-concluida">✅ {status}</span>'
    else:
        return f'<span class="status-outro">⚠️ {status}</span>'

# ---------------------------------------------------------
# SESSION
# ---------------------------------------------------------
if "logged" not in st.session_state: st.session_state.logged = False
if "user"   not in st.session_state: st.session_state.user   = {}
if "page"   not in st.session_state: st.session_state.page   = "public"

# ---------------------------------------------------------
# NAVEGAÇÃO (SIDEBAR)
# ---------------------------------------------------------
with st.sidebar:
    st.title("🧭 Navegação")
    
    if st.session_state.logged:
        st.success(f"👤 {st.session_state.user.get('name', '')}")
        st.caption(f"Perfil: {role_of(st.session_state.user)}")
        
        if st.button("🌐 Página Pública", use_container_width=True, key="btn_public_sidebar"):
            st.session_state.page = "public"; st.rerun()
        if st.button("🔒 Área Restrita", use_container_width=True, key="btn_private_sidebar"):
            st.session_state.page = "private"; st.rerun()
        
        st.divider()
        if st.button("🚪 Sair", use_container_width=True, key="btn_logout_sidebar"):
            st.session_state.logged = False
            st.session_state.user = {}
            st.session_state.page = "public"
            st.rerun()
    else:
        st.info("🔓 Área Pública")
        st.caption("Visualize as produções e estatísticas")
        if st.button("🔑 Login", use_container_width=True, key="btn_login_sidebar"):
            st.session_state.page = "login"; st.rerun()

# =========================================================
# PÁGINA PÚBLICA (SEM LOGIN)
# =========================================================
if st.session_state.page == "public":
    st.markdown("""
    <div class="public-notice">
    <h3>📊 Portal Público de Produção Científica</h3>
    <p>Acompanhe as produções científicas do PPG, o envolvimento dos discentes e a diversidade da pesquisa.</p>
    </div>
    """, unsafe_allow_html=True)
    
    exibir_video_entrada()
    
    tab_dashboard, tab_producoes, tab_biblio_pub, tab_orient_pub, tab_ensino_pub, tab_impacto_pub = st.tabs([
        "📊 Dashboard", "📚 Produções Científicas", 
        "📈 Bibliometria", "🎓 Orientações", "📖 Ensino", "🌍 Impacto na Sociedade"
    ])
    
    with tab_dashboard:
        stats = get_estatisticas_avancadas()
        
        st.subheader("📈 Visão Geral")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{stats['total_producoes']}</div>
                <div class="metric-label">Total de Produções</div>
            </div>""", unsafe_allow_html=True)
        with c2:
            st.markdown(f"""
            <div class="metric-card-blue">
                <div class="metric-value">{stats['total_com_ppg']}</div>
                <div class="metric-label">Com Discentes do PPG</div>
            </div>""", unsafe_allow_html=True)
        with c3:
            st.markdown(f"""
            <div class="metric-card-green">
                <div class="metric-value">{stats['total_com_estrangeiros']}</div>
                <div class="metric-label">Com Pesquisadores Estrangeiros</div>
            </div>""", unsafe_allow_html=True)
        with c4:
            percent = (stats['total_com_ppg'] / stats['total_producoes'] * 100) if stats['total_producoes'] > 0 else 0
            st.markdown(f"""
            <div class="metric-card-orange">
                <div class="metric-value">{percent:.1f}%</div>
                <div class="metric-label">Participação Discente</div>
            </div>""", unsafe_allow_html=True)
        
        st.divider()       
        
        st.subheader("🎓 Atividades Acadêmicas Complementares")
        
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"""
            <div class="metric-card-pink">
                <div class="metric-value">{stats['total_orientacoes']}</div>
                <div class="metric-label">🎓 Orientações</div>
            </div>""", unsafe_allow_html=True)
        with c2:
            st.markdown(f"""
            <div class="metric-card-blue">
                <div class="metric-value">{stats['total_ensino']}</div>
                <div class="metric-label">📖 Atividades de Ensino</div>
            </div>""", unsafe_allow_html=True)
        with c3:
            st.markdown(f"""
            <div class="metric-card-green">
                <div class="metric-value">{stats['total_impacto']}</div>
                <div class="metric-label">🌍 Atividades de Impacto</div>
            </div>""", unsafe_allow_html=True)
        
        st.divider()
        
        st.subheader("📰 Análise de Periódicos Científicos")
        col1, col2 = st.columns([1, 2])
        with col1:
            st.markdown(f"""
            <div class="highlight-box">
                <h3 style="text-align:center; color:#667eea;">📚 {stats['total_periodicos_unicos']}</h3>
                <p style="text-align:center; margin:0;"><strong>Periódicos diferentes</strong><br>onde o PPG publicou artigos</p>
            </div>""", unsafe_allow_html=True)
        with col2:
            if stats['top_periodicos']:
                st.markdown("#### 🏆 Top 10 Periódicos Mais Frequentes")
                df_top = pd.DataFrame({
                    'Periódico': list(stats['top_periodicos'].keys()),
                    'Publicações': list(stats['top_periodicos'].values())
                })
                st.bar_chart(df_top.set_index('Periódico'))
            else:
                st.info("Nenhum artigo em periódico cadastrado ainda.")
        
        st.divider()
        
        st.subheader("🌍 Cooperação Internacional")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### 📊 Produções com Pesquisadores Estrangeiros por Ano")
            df_estrangeiros = pd.DataFrame({
                'Ano': list(stats['producoes_com_estrangeiros_por_ano'].keys()),
                'Com Estrangeiros': list(stats['producoes_com_estrangeiros_por_ano'].values())
            })
            st.bar_chart(df_estrangeiros.set_index('Ano'))
        with col2:
            st.markdown(f"""
            <div class="highlight-box">
                <h3 style="text-align:center; color:#11998e;">🌐 {stats['total_com_estrangeiros']}</h3>
                <p style="text-align:center; margin:0;"><strong>Produções com cooperação internacional</strong><br>envolvendo pesquisadores estrangeiros</p>
            </div>""", unsafe_allow_html=True)
            if stats['total_producoes'] > 0:
                perc_estrangeiros = (stats['total_com_estrangeiros'] / stats['total_producoes'] * 100)
                st.markdown(f"""
                <div class="highlight-box" style="background:#e8f5e9;">
                    <p style="text-align:center; margin:0; font-size:1.1rem;">
                        <strong>{perc_estrangeiros:.1f}%</strong> das produções têm<br>colaboração internacional
                    </p>
                </div>""", unsafe_allow_html=True)
        
        st.divider()
        
        st.subheader("✍️ Análise de Autoria em Artigos")
        st.markdown("""
        <div class="highlight-box" style="background:#fff3e0;">
        <p><strong>📌 Como funciona esta análise:</strong> Os docentes marcam <strong>explicitamente</strong> 
        durante o cadastro/edição da produção se um discente do PPG é o primeiro/último autor e/ou se um docente 
        do PPG é o primeiro/último autor. As estatísticas refletem <strong>apenas as marcações explícitas</strong>, 
        garantindo 100% de confiabilidade nos dados.</p>
        </div>""", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"""
            <div class="metric-card-blue">
                <div class="metric-value">{stats['artigos_com_discente_primeiro']}</div>
                <div class="metric-label">🎓 Artigos com Discente do PPG<br>como <strong>Primeiro/Último Autor</strong></div>
            </div>""", unsafe_allow_html=True)
            st.info("💡 Indica destaque discente")
        with col2:
            st.markdown(f"""
            <div class="metric-card-green">
                <div class="metric-value">{stats['artigos_com_docente_ultimo']}</div>
                <div class="metric-label">👨‍🏫 Artigos com Docente do PPG<br>como <strong>Primeiro/Último Autor</strong></div>
            </div>""", unsafe_allow_html=True)
            st.info("💡 Indica destaque docente")
        
        st.divider()
        
        st.subheader("📅 Evolução Temporal")
        st.markdown("#### Total de Produções por Ano")
        df_total_ano = pd.DataFrame({
            'Ano': list(stats['producoes_por_ano'].keys()),
            'Produções': list(stats['producoes_por_ano'].values())
        })
        st.bar_chart(df_total_ano.set_index('Ano'))
        
        st.divider()
        
        st.subheader("📋 Resumo Detalhado por Ano")
        df_resumo = pd.DataFrame({
            'Ano': ANOS,
            'Total de Produções': [stats['producoes_por_ano'].get(ano, 0) for ano in ANOS],
            'Com Discentes PPG': [stats['producoes_com_ppg_por_ano'].get(ano, 0) for ano in ANOS],
            'Com Estrangeiros': [stats['producoes_com_estrangeiros_por_ano'].get(ano, 0) for ano in ANOS],
            'Orientações iniciadas': [stats['orient_por_ano'].get(ano, 0) for ano in ANOS],
        })
        st.dataframe(df_resumo, use_container_width=True)
        st.divider()

        st.subheader("☁️ Nuvem de Palavras dos Títulos")
        st.markdown("""
        <div class="highlight-box">
        <p><strong>📌 O que é isto?</strong> Visualização das palavras mais frequentes nos títulos das produções 
        científicas do PPG. Palavras maiores aparecem com mais frequência.</p>
        </div>""", unsafe_allow_html=True)
        
        df_prod_cloud = read_df(SHEET_PROD)
        texto_titulos = extrair_texto_titulos(df_prod_cloud)
        
        if texto_titulos:
            fig_wordcloud = gerar_wordcloud(texto_titulos, max_words=100)
            if fig_wordcloud:
                st.pyplot(fig_wordcloud)
        else:
            st.info("Ainda não há títulos cadastrados para gerar a nuvem de palavras.")
        
        st.divider()
    with tab_producoes:
        st.subheader("📚 Produções Científicas")
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
                        discente_primeiro = str(row.get("discente_primeiro_autor", "")).strip().lower()
                        tem_destaque = discente_primeiro == "sim"
                        tem_discente_ppg = tem_participacao_ppg(row["id"])
                        
                        titulo_display = row['titulo']
                        if tem_destaque:
                            titulo_display = f"🌟 {titulo_display}"
                        if tem_discente_ppg:
                            titulo_display = f"🚀 {titulo_display}"
                        
                        with st.expander(f"**{titulo_display}** — {row['tipo']}"):
                            autor_principal_nome = get_nome_autor_principal(row["docente_username"])
                            st.markdown(f'<span class="main-author-tag">👤 Autor Principal: {autor_principal_nome}</span>', 
                                       unsafe_allow_html=True)
                            
                            badges = []
                            docente_ultimo = str(row.get("docente_ultimo_autor", "")).strip().lower()
                            if discente_primeiro == "sim":
                                badges.append('<span class="badge-discente-1">🎓 Discente é 1º autor</span>')
                            if docente_ultimo == "sim":
                                badges.append('<span class="badge-docente-last">👨‍🏫 Docente é último autor</span>')
                            if badges:
                                st.markdown(" ".join(badges), unsafe_allow_html=True)
                            
                            st.write(f"**Veículo:** {row['veiculo']}")
                            st.write(f"**Autores:** {row['autores']}")
                            st.write(f"**DOI:** {row['doi'] or '—'}")
                            descricao_text = str(row.get('descricao', '')).strip()
                            if descricao_text:
                                st.markdown(f'<div class="descricao-box"><b>📝 Descrição:</b><br>{descricao_text}</div>', 
                                           unsafe_allow_html=True)
                            parts = df_part[df_part["producao_id"] == row["id"]] if not df_part.empty else pd.DataFrame()
                            if not parts.empty:
                                st.write("**Participações:**")
                                st.dataframe(parts[["tipo_participacao","nome_participante","vinculo"]], use_container_width=True)
                                if any(parts["tipo_participacao"] == "Discente do PPG"):
                                    st.success("✅ Conta com participação de discente(s) do PPG")
                                if any(parts["tipo_participacao"] == "Pesquisador estrangeiro"):
                                    st.info("🌍 Conta com participação de pesquisador(es) estrangeiro(s)")
                            else:
                                st.info("Nenhuma participação registrada.")
            
            if df_filtrado.empty:
                st.info("Nenhuma produção encontrada com os filtros selecionados.")
    
    with tab_biblio_pub:
        render_bibliometria_docente()
    
    with tab_orient_pub:
        st.subheader("🎓 Orientações Acadêmicas")
        df_orient = read_df(SHEET_ORIENT)
        df_users = read_df(SHEET_USERS)
        
        if df_orient.empty:
            st.info("Nenhuma orientação cadastrada ainda.")
        else:
            col1, col2 = st.columns(2)
            with col1:
                filtro_tipo_ori = st.multiselect("Filtrar por tipo", TIPOS_ORIENTACAO, 
                                                 key="filtro_tipo_ori_public")
            with col2:
                filtro_status_ori = st.multiselect("Filtrar por status", STATUS_ORIENTACAO,
                                                   key="filtro_status_ori_public")
            
            df_ori_filt = df_orient.copy()
            if filtro_tipo_ori:
                df_ori_filt = df_ori_filt[df_ori_filt["tipo"].isin(filtro_tipo_ori)]
            if filtro_status_ori:
                df_ori_filt = df_ori_filt[df_ori_filt["status"].isin(filtro_status_ori)]
            
            df_ori_filt = df_ori_filt.sort_values(by="discente_nome", key=lambda x: x.str.lower())
            
            st.write(f"**Total:** {len(df_ori_filt)} orientações encontradas")
            
            for _, row in df_ori_filt.iterrows():
                docente_user = users_get(row["docente_username"]) if not df_users.empty else None
                docente_nome = docente_user["name"] if docente_user else row["docente_username"]
                ano_conc_display = str(row.get('ano_conclusao', '')).strip() if str(row.get('ano_conclusao', '')).strip() else "em andamento"
                
                with st.expander(f"**{row['discente_nome']}** — {row['tipo']} ({row['ano_inicio']} → {ano_conc_display})"):
                    st.markdown(badge_status(row['status']), unsafe_allow_html=True)
                    st.write(f"**Orientador(a):** {docente_nome}")
                    st.write(f"**Ano de entrada:** {row['ano_inicio']}")
                    st.write(f"**Ano de saída:** {row['ano_conclusao'] or '—'}")
    
    with tab_ensino_pub:
        st.subheader("📖 Atividades de Ensino")
        df_ensino = read_df(SHEET_ENSINO)
        df_users = read_df(SHEET_USERS)
        
        if df_ensino.empty:
            st.info("Nenhuma atividade de ensino cadastrada ainda.")
        else:
            col1, col2 = st.columns(2)
            with col1:
                filtro_tipo_ens = st.multiselect("Filtrar por tipo", TIPOS_ENSINO, 
                                                 key="filtro_tipo_ens_public")
            with col2:
                filtro_nivel_ens = st.multiselect("Filtrar por nível", NIVEIS_ENSINO,
                                                  key="filtro_nivel_ens_public")
            
            df_ens_filt = df_ensino.copy()
            if filtro_tipo_ens:
                df_ens_filt = df_ens_filt[df_ens_filt["tipo"].isin(filtro_tipo_ens)]
            if filtro_nivel_ens:
                df_ens_filt = df_ens_filt[df_ens_filt["nivel"].isin(filtro_nivel_ens)]
            
            st.write(f"**Total:** {len(df_ens_filt)} atividades encontradas")
            
            for ano in ANOS:
                subset = df_ens_filt[df_ens_filt["ano"].astype(str).str.strip() == ano]
                if not subset.empty:
                    st.markdown(f"### 📅 {ano}")
                    for _, row in subset.iterrows():
                        docente_user = users_get(row["docente_username"]) if not df_users.empty else None
                        docente_nome = docente_user["name"] if docente_user else row["docente_username"]
                        
                        with st.expander(f"**{row['titulo']}** — {row['tipo']}"):
                            st.write(f"**Docente:** {docente_nome}")
                            st.write(f"**Período:** {row['periodo'] or '—'}")
                            st.write(f"**Nível:** {row['nivel']}")
                            st.write(f"**Carga horária:** {row['carga_horaria'] or '—'} h")
                            desc = str(row.get('descricao', '')).strip()
                            if desc:
                                st.markdown(f'<div class="descricao-box"><b>📝 Descrição:</b><br>{desc}</div>', 
                                           unsafe_allow_html=True)
    
    with tab_impacto_pub:
        st.subheader("🌍 Atividades de Impacto na Sociedade")
        df_impacto = read_df(SHEET_IMPACTO)
        df_users = read_df(SHEET_USERS)
        
        if df_impacto.empty:
            st.info("Nenhuma atividade de impacto cadastrada ainda.")
        else:
            filtro_tipo_imp = st.multiselect("Filtrar por tipo", TIPOS_IMPACTO, 
                                             key="filtro_tipo_imp_public")
            
            df_imp_filt = df_impacto.copy()
            if filtro_tipo_imp:
                df_imp_filt = df_imp_filt[df_imp_filt["tipo"].isin(filtro_tipo_imp)]
            
            st.write(f"**Total:** {len(df_imp_filt)} atividades encontradas")
            
            for _, row in df_imp_filt.iterrows():
                docente_user = users_get(row["docente_username"]) if not df_users.empty else None
                docente_nome = docente_user["name"] if docente_user else row["docente_username"]
                
                with st.expander(f"**{row['titulo']}** — {row['tipo']}"):
                    st.write(f"**Docente:** {docente_nome}")
                    st.write(f"**Data:** {row['data'] or '—'}")
                    st.write(f"**Local:** {row['local'] or '—'}")
                    st.write(f"**Público-alvo:** {row['publico_alvo'] or '—'}")
                    desc = str(row.get('descricao', '')).strip()
                    if desc:
                        st.markdown(f'<div class="descricao-box"><b>📝 Descrição:</b><br>{desc}</div>', 
                                   unsafe_allow_html=True)

# =========================================================
# LOGIN / CADASTRO
# =========================================================
elif st.session_state.page == "login":
    st.subheader("🔐 Acesso Restrito")
    tab_login, tab_cad = st.tabs(["🔑 Login", "📝 Cadastro"])
    
    with tab_login:
        st.markdown("<div class='block-card'>", unsafe_allow_html=True)
        u = st.text_input("Usuário", key="login_user")
        p = st.text_input("Senha", type="password", key="login_pass")
        if st.button("Entrar", use_container_width=True, key="btn_login_main"):
            ok, res = authenticate(u, p)
            if ok:
                st.session_state.logged = True
                st.session_state.user = res
                st.session_state.page = "private"
                st.rerun()
            else: st.error(res)
        st.markdown("</div>", unsafe_allow_html=True)
        if st.button("← Voltar para página pública", key="btn_voltar_public_login"):
            st.session_state.page = "public"; st.rerun()

    with tab_cad:
        st.markdown("<div class='block-card'>", unsafe_allow_html=True)
        name = st.text_input("Nome completo", key="cad_name")
        username = st.text_input("Username (login)", key="cad_username")
        email = st.text_input("Email", key="cad_email")
        role = st.selectbox("Perfil", ["docente", "discente"], key="cad_role")
        orientador = ""
        if role == "discente":
            docentes = listar_docentes()
            if docentes: orientador = st.selectbox("Orientador (no PPG)", docentes, key="cad_orientador_sel")
            else: orientador = st.text_input("Nome do orientador", key="cad_orientador_txt")
        pw1 = st.text_input("Senha", type="password", key="cad_pw1")
        pw2 = st.text_input("Confirmar senha", type="password", key="cad_pw2")
        if st.button("Solicitar cadastro", use_container_width=True, key="btn_cadastro"):
            if not all([name, username, email]): st.error("Preencha nome, username e email.")
            elif len(pw1) < 6: st.error("Senha com no mínimo 6 caracteres.")
            elif pw1 != pw2: st.error("Senhas não coincidem.")
            else:
                ok, msg = cadastro_submit(name, username, email, pw1, role, orientador)
                (st.success if ok else st.error)(msg)
        st.info("O coordenador precisa aprovar seu cadastro.")
        st.markdown("</div>", unsafe_allow_html=True)
        if st.button("← Voltar para página pública", key="btn_voltar_public_cad"):
            st.session_state.page = "public"; st.rerun()

# =========================================================
# ÁREA RESTRITA (LOGADO)
# =========================================================
elif st.session_state.page == "private":
    user = st.session_state.user
    user_role = role_of(user)
    user_username = user.get("username", "")
    
    st.success(f"Logado como **{user.get('name','')}**  | perfil: **{user_role}**")
    
    if user_role == "admin":
        st.subheader("🛠️ Painel do Coordenador")
        df_users = read_df(SHEET_USERS)
        df_cad = read_df(SHEET_CAD)
        df_prod = read_df(SHEET_PROD)
        df_part = read_df(SHEET_PART)
        df_vinc = read_df(SHEET_VINC)
        df_orient = read_df(SHEET_ORIENT)
        df_ensino = read_df(SHEET_ENSINO)
        df_impacto = read_df(SHEET_IMPACTO)

        t1, t2, t3, t4, t5, t6, t7, t8, t9 = st.tabs([
            "👥 Usuários", "👤 Cadastros pendentes", "📚 Produções (geral)", 
            "📊 Resumo por ano", "⚙️ Histórico", 
            "➕ Cadastrar Produção", 
            "🎓 Orientações",
            "📖 Ensino",
            "🌍 Impacto"
        ])

        with t1:
            if df_users.empty: st.info("Sem usuários.")
            else:
                cols = [c for c in ["username","name","email","role","orientador","created_at"] if c in df_users.columns]
                st.dataframe(df_users[cols], use_container_width=True)

        with t2:
            pend = df_cad[df_cad.get("status","") == "Pendente"] if not df_cad.empty else pd.DataFrame()
            if pend.empty: st.info("Sem cadastros pendentes.")
            else:
                cols = [c for c in ["id","name","username","email","role","orientador","created_at"] if c in pend.columns]
                st.dataframe(pend[cols], use_container_width=True)
                sel_id = st.selectbox("Selecione um cadastro", pend["id"].tolist(), key="admin_sel_cad")
                reason = st.text_input("Motivo (opcional)", key="admin_reason_cad")
                cA, cR = st.columns(2)
                with cA:
                    if st.button("✅ Aprovar", use_container_width=True, key="btn_aprovar_cad"):
                        ok, msg = cadastro_review(sel_id, "Aprovar", user["username"], reason)
                        (st.success if ok else st.error)(msg); st.rerun()
                with cR:
                    if st.button("❌ Rejeitar", use_container_width=True, key="btn_rejeitar_cad"):
                        ok, msg = cadastro_review(sel_id, "Rejeitar", user["username"], reason)
                        (st.success if ok else st.error)(msg); st.rerun()

        with t3:
            if df_prod.empty: st.info("Sem produções cadastradas.")
            else:
                st.dataframe(df_prod, use_container_width=True)
                st.caption("Total de produções: " + str(len(df_prod)))
                if not df_part.empty:
                    st.write("**Participações registradas:**"); st.dataframe(df_part, use_container_width=True)

        with t4:
            if df_prod.empty: st.info("Sem produções.")
            else:
                for ano in ANOS:
                    st.markdown(f"### 📅 {ano}")
                    subset = df_prod[df_prod["ano"].astype(str).str.strip() == ano]
                    if subset.empty: st.write("— Nenhuma produção registrada —")
                    else:
                        st.write(f"Total: {len(subset)}")
                        if not df_part.empty:
                            ids_ano = subset["id"].tolist()
                            parts_ano = df_part[df_part["producao_id"].isin(ids_ano)]
                            if not parts_ano.empty: st.dataframe(parts_ano["tipo_participacao"].value_counts(), use_container_width=True)

        with t5:
            hist = df_cad[df_cad.get("status","").isin(["Aprovado","Rejeitado"])] if not df_cad.empty else pd.DataFrame()
            if hist.empty: st.info("Sem histórico.")
            else: st.dataframe(hist, use_container_width=True)

        with t6:
            st.subheader("➕ Cadastrar nova produção para um docente")
            docentes_df = df_users[df_users["role"].str.lower().isin(["docente", "professor"])] if not df_users.empty else pd.DataFrame()
            if docentes_df.empty:
                st.info("Nenhum docente cadastrado no sistema ainda.")
            else:
                docente_options = {f"{row['name']} ({row['username']})": row['username'] for _, row in docentes_df.iterrows()}
                selected_label = st.selectbox("Selecione o Docente", list(docente_options.keys()))
                selected_username = docente_options[selected_label]
                st.divider()
                with st.form("admin_form_prod"):
                    c1, c2 = st.columns(2)
                    with c1:
                        titulo = st.text_input("Título", key="admin_prod_titulo")
                        tipo = st.selectbox("Tipo", TIPOS_PRODUCAO, key="admin_prod_tipo")
                        ano = st.selectbox("Ano", ANOS, key="admin_prod_ano")
                    with c2:
                        veiculo = st.text_input("Veículo/Periódico", key="admin_prod_veiculo")
                        autores = st.text_input("Autores (separados por vírgula, na ordem)", key="admin_prod_autores")
                        doi = st.text_input("DOI (opcional)", key="admin_prod_doi")
                    descricao = st.text_area("📝 Descrição qualitativa (opcional)",
                        placeholder="Descreva o contexto, relevância, impacto...", height=100, key="admin_prod_descricao")
                    st.markdown("**👥 Co-autores do PPG (opcional)**")
                    docentes_list = listar_docentes()
                    docentes_list = [d for d in docentes_list if d != selected_label.split(" (")[0]]
                    co_autores_selecionados = st.multiselect(
                        "Selecione outros docentes do PPG que são co-autores:",
                        docentes_list, key="admin_prod_coautores")
                    co_autores_usernames = ",".join([
                        get_docente_username_by_name(nome) for nome in co_autores_selecionados if get_docente_username_by_name(nome)])
                    
                    discente_primeiro_str, docente_ultimo_str = renderizar_checkboxes_autoria("admin_cad")
                    
                    submitted = st.form_submit_button("💾 Cadastrar", use_container_width=True)
                    if submitted:
                        if not titulo.strip(): st.error("Título é obrigatório.")
                        else:
                            producao_submit(selected_username, titulo, tipo, ano, veiculo, 
                                           autores, doi, descricao, co_autores_usernames,
                                           discente_primeiro_str, docente_ultimo_str)
                            st.success(f"Produção cadastrada para {selected_label.split(' (')[0]}!")
                            st.rerun()
        
        with t7:
            st.subheader("🎓 Orientações Acadêmicas")
            docentes_df = df_users[df_users["role"].str.lower().isin(["docente", "professor"])] if not df_users.empty else pd.DataFrame()
            
            if docentes_df.empty:
                st.info("Cadastre docentes primeiro.")
            else:
                docente_options = {f"{row['name']} ({row['username']})": row['username'] for _, row in docentes_df.iterrows()}
                selected_label = st.selectbox("Selecione o Docente orientador", list(docente_options.keys()), 
                                              key="sel_doc_orient")
                selected_username = docente_options[selected_label]
                st.divider()
                
                with st.form("form_orient"):
                    c1, c2 = st.columns(2)
                    with c1:
                        discente_nome = st.text_input("Nome do orientando", key="ori_discente")
                        tipo = st.selectbox("Tipo de orientação", TIPOS_ORIENTACAO, key="ori_tipo")
                        ano_inicio = st.text_input("Ano de entrada", placeholder="Ex: 2023", key="ori_ano_ini")
                    with c2:
                        status = st.selectbox("Status", STATUS_ORIENTACAO, key="ori_status")
                        ano_conclusao = st.text_input("Ano de saída (se aplicável)", 
                                                       placeholder="Ex: 2025 ou vazio", key="ori_ano_conc")
                    
                    submitted = st.form_submit_button("💾 Cadastrar orientação", use_container_width=True)
                    if submitted:
                        if not discente_nome.strip() or not ano_inicio.strip():
                            st.error("Nome do orientando e ano de entrada são obrigatórios.")
                        else:
                            orientacao_submit(selected_username, discente_nome, tipo,
                                             ano_inicio, ano_conclusao, status)
                            st.success(f"Orientação cadastrada para {selected_label.split(' (')[0]}!")
                            st.rerun()
                
                st.divider()
                st.markdown("### 📋 Orientações cadastradas")
                if df_orient.empty:
                    st.info("Nenhuma orientação cadastrada.")
                else:
                    ori_docente = df_orient[df_orient["docente_username"] == selected_username]
                    if ori_docente.empty:
                        st.info(f"Sem orientações para {selected_label.split(' (')[0]}.")
                    else:
                        ori_docente = ori_docente.sort_values(by="discente_nome", key=lambda x: x.str.lower())
                        
                        for _, row in ori_docente.iterrows():
                            ano_conc_display = str(row.get('ano_conclusao', '')).strip() if str(row.get('ano_conclusao', '')).strip() else "em andamento"
                            with st.expander(f"**{row['discente_nome']}** — {row['tipo']} ({row['ano_inicio']} → {ano_conc_display})"):
                                st.markdown(badge_status(row['status']), unsafe_allow_html=True)
                                st.write(f"**Ano de entrada:** {row['ano_inicio']}")
                                st.write(f"**Ano de saída:** {row['ano_conclusao'] or '—'}")
                                if st.button("🗑️ Excluir", key=f"del_ori_{row['id']}", use_container_width=True):
                                    ok, msg = orientacao_delete(row['id'])
                                    if ok: st.success(msg); st.rerun()
                                    else: st.error(msg)
        
        with t8:
            st.subheader("📖 Atividades de Ensino")
            docentes_df = df_users[df_users["role"].str.lower().isin(["docente", "professor"])] if not df_users.empty else pd.DataFrame()
            
            if docentes_df.empty:
                st.info("Cadastre docentes primeiro.")
            else:
                docente_options = {f"{row['name']} ({row['username']})": row['username'] for _, row in docentes_df.iterrows()}
                selected_label = st.selectbox("Selecione o Docente", list(docente_options.keys()), 
                                              key="sel_doc_ensino")
                selected_username = docente_options[selected_label]
                st.divider()
                
                with st.form("form_ensino"):
                    c1, c2 = st.columns(2)
                    with c1:
                        tipo = st.selectbox("Tipo de atividade", TIPOS_ENSINO, key="ens_tipo")
                        titulo = st.text_input("Título da disciplina/atividade", key="ens_titulo")
                        ano = st.selectbox("Ano", ANOS, key="ens_ano")
                    with c2:
                        periodo = st.text_input("Período (ex: 2026.1)", key="ens_periodo")
                        carga_horaria = st.number_input("Carga horária (horas)", min_value=0, step=1, key="ens_ch")
                        nivel = st.selectbox("Nível", NIVEIS_ENSINO, key="ens_nivel")
                    descricao = st.text_area("Descrição (opcional)", height=80, key="ens_desc")
                    
                    submitted = st.form_submit_button("💾 Cadastrar atividade", use_container_width=True)
                    if submitted:
                        if not titulo.strip():
                            st.error("Título é obrigatório.")
                        else:
                            ensino_submit(selected_username, tipo, titulo, periodo, ano,
                                         carga_horaria, nivel, descricao)
                            st.success(f"Atividade cadastrada para {selected_label.split(' (')[0]}!")
                            st.rerun()
                
                st.divider()
                st.markdown("### 📋 Atividades cadastradas")
                if df_ensino.empty:
                    st.info("Nenhuma atividade cadastrada.")
                else:
                    ens_docente = df_ensino[df_ensino["docente_username"] == selected_username]
                    if ens_docente.empty:
                        st.info(f"Sem atividades para {selected_label.split(' (')[0]}.")
                    else:
                        for ano in ANOS:
                            subset = ens_docente[ens_docente["ano"].astype(str).str.strip() == ano]
                            if not subset.empty:
                                st.markdown(f"#### 📅 {ano}")
                                for _, row in subset.iterrows():
                                    with st.expander(f"**{row['titulo']}** — {row['tipo']}"):
                                        st.write(f"**Período:** {row['periodo'] or '—'}")
                                        st.write(f"**Nível:** {row['nivel']}")
                                        st.write(f"**Carga horária:** {row['carga_horaria'] or '—'} h")
                                        desc = str(row.get('descricao', '')).strip()
                                        if desc:
                                            st.markdown(f'<div class="descricao-box"><b>📝 Descrição:</b><br>{desc}</div>', 
                                                       unsafe_allow_html=True)
                                        if st.button("🗑️ Excluir", key=f"del_ens_{row['id']}", use_container_width=True):
                                            ok, msg = ensino_delete(row['id'])
                                            if ok: st.success(msg); st.rerun()
                                            else: st.error(msg)
        
        with t9:
            st.subheader("🌍 Atividades de Impacto na Sociedade")
            docentes_df = df_users[df_users["role"].str.lower().isin(["docente", "professor"])] if not df_users.empty else pd.DataFrame()
            
            if docentes_df.empty:
                st.info("Cadastre docentes primeiro.")
            else:
                docente_options = {f"{row['name']} ({row['username']})": row['username'] for _, row in docentes_df.iterrows()}
                selected_label = st.selectbox("Selecione o Docente", list(docente_options.keys()), 
                                              key="sel_doc_impacto")
                selected_username = docente_options[selected_label]
                st.divider()
                
                with st.form("form_impacto"):
                    c1, c2 = st.columns(2)
                    with c1:
                        tipo = st.selectbox("Tipo de atividade", TIPOS_IMPACTO, key="imp_tipo")
                        titulo = st.text_input("Título da atividade", key="imp_titulo")
                        data = st.text_input("Data (ex: 2026-03-15)", key="imp_data")
                    with c2:
                        local = st.text_input("Local/instituição", key="imp_local")
                        publico_alvo = st.text_input("Público-alvo", key="imp_publico")
                    descricao = st.text_area("Descrição detalhada", height=100, key="imp_desc")
                    
                    submitted = st.form_submit_button("💾 Cadastrar atividade", use_container_width=True)
                    if submitted:
                        if not titulo.strip():
                            st.error("Título é obrigatório.")
                        else:
                            impacto_submit(selected_username, tipo, titulo, descricao, 
                                          data, publico_alvo, local)
                            st.success(f"Atividade cadastrada para {selected_label.split(' (')[0]}!")
                            st.rerun()
                
                st.divider()
                st.markdown("### 📋 Atividades cadastradas")
                if df_impacto.empty:
                    st.info("Nenhuma atividade cadastrada.")
                else:
                    imp_docente = df_impacto[df_impacto["docente_username"] == selected_username]
                    if imp_docente.empty:
                        st.info(f"Sem atividades para {selected_label.split(' (')[0]}.")
                    else:
                        for _, row in imp_docente.iterrows():
                            with st.expander(f"**{row['titulo']}** — {row['tipo']}"):
                                st.write(f"**Data:** {row['data'] or '—'}")
                                st.write(f"**Local:** {row['local'] or '—'}")
                                st.write(f"**Público-alvo:** {row['publico_alvo'] or '—'}")
                                desc = str(row.get('descricao', '')).strip()
                                if desc:
                                    st.markdown(f'<div class="descricao-box"><b>📝 Descrição:</b><br>{desc}</div>', 
                                               unsafe_allow_html=True)
                                if st.button("🗑️ Excluir", key=f"del_imp_{row['id']}", use_container_width=True):
                                    ok, msg = impacto_delete(row['id'])
                                    if ok: st.success(msg); st.rerun()
                                    else: st.error(msg)
        
        st.divider()

    elif user_role == "docente":
        st.subheader(f"📚 Minhas produções — {user.get('name','')}")
        df_prod = read_df(SHEET_PROD)
        df_part = read_df(SHEET_PART)
        todas_producoes = get_minhas_producoes(user_username)
        
        st.markdown("### 📋 Minhas produções")
        if todas_producoes.empty:
            st.info("Nenhuma produção cadastrada.")
        else:
            for ano in ANOS:
                subset = todas_producoes[todas_producoes["ano"].astype(str).str.strip() == ano]
                if not subset.empty:
                    st.markdown(f"#### 📅 {ano}")
                    for _, row in subset.iterrows():
                        eh_principal = row.get("tipo_autoria", "") == "principal"
                        
                        discente_primeiro = str(row.get("discente_primeiro_autor", "")).strip().lower()
                        tem_destaque = discente_primeiro == "sim"
                        tem_discente_ppg = tem_participacao_ppg(row["id"])
                        
                        titulo_display = row['titulo']
                        if tem_destaque:
                            titulo_display = f"🌟 {titulo_display}"
                        if tem_discente_ppg:
                            titulo_display = f"🚀 {titulo_display}"
                        
                        with st.expander(f"**{titulo_display}** — {row['tipo']}"):
                            autor_principal_nome = get_nome_autor_principal(row["docente_username"])
                            st.markdown(f'<span class="main-author-tag">👤 Autor Principal: {autor_principal_nome}</span>', 
                                       unsafe_allow_html=True)
                            if eh_principal:
                                st.markdown('<span class="autor-principal-badge">📝 Você é o responsável pelo cadastro</span>', 
                                           unsafe_allow_html=True)
                            else:
                                st.markdown('<span class="coautor-badge">👥 Você participa como co-autor</span>', 
                                           unsafe_allow_html=True)
                            
                            badges = []
                            docente_ultimo = str(row.get("docente_ultimo_autor", "")).strip().lower()
                            if discente_primeiro == "sim":
                                badges.append('<span class="badge-discente-1">🎓 Discente é 1º autor</span>')
                            if docente_ultimo == "sim":
                                badges.append('<span class="badge-docente-last">👨‍🏫 Docente é último autor</span>')
                            if badges:
                                st.markdown(" ".join(badges), unsafe_allow_html=True)
                            
                            st.write(f"**Veículo:** {row['veiculo']}")
                            st.write(f"**Autores:** {row['autores']}")
                            st.write(f"**DOI:** {row['doi'] or '—'}")
                            descricao_text = str(row.get('descricao', '')).strip()
                            if descricao_text:
                                st.markdown(f'<div class="descricao-box"><b>📝 Descrição:</b><br>{descricao_text}</div>', 
                                           unsafe_allow_html=True)
                            parts = df_part[df_part["producao_id"] == row["id"]] if not df_part.empty else pd.DataFrame()
                            if not parts.empty:
                                st.write("**Participações:**")
                                st.dataframe(parts[["tipo_participacao","nome_participante","vinculo"]], use_container_width=True)
                            if eh_principal:
                                col1, col2 = st.columns(2)
                                with col1:
                                    if st.button("✏️ Editar", key=f"edit_{row['id']}", use_container_width=True):
                                        st.session_state['editing_prod_id'] = row['id']; st.rerun()
                                with col2:
                                    if st.button("🗑️ Excluir", key=f"del_{row['id']}", use_container_width=True):
                                        st.session_state['deleting_prod_id'] = row['id']; st.rerun()
                            else:
                                st.info("💡 Co-autores podem visualizar, mas não editar/excluir esta produção.")
        
        st.divider()
        
        
        
        if 'editing_prod_id' in st.session_state:
            pid = st.session_state['editing_prod_id']
            prod_filtered = df_prod[df_prod["id"] == pid] if not df_prod.empty else pd.DataFrame()
            prod_data = prod_filtered.iloc[0] if not prod_filtered.empty else None
            if prod_data is not None:
                st.markdown("### ✏️ Editar produção")
                with st.form(f"form_edit_{pid}"):
                    c1, c2 = st.columns(2)
                    with c1:
                        titulo = st.text_input("Título", value=prod_data['titulo'], key=f"edit_titulo_{pid}")
                        tipo_idx = TIPOS_PRODUCAO.index(prod_data['tipo']) if prod_data['tipo'] in TIPOS_PRODUCAO else 0
                        tipo = st.selectbox("Tipo", TIPOS_PRODUCAO, index=tipo_idx, key=f"edit_tipo_{pid}")
                        ano_clean = str(prod_data['ano']).strip()
                        ano_idx = ANOS.index(ano_clean) if ano_clean in ANOS else 0
                        ano = st.selectbox("Ano", ANOS, index=ano_idx, key=f"edit_ano_{pid}")
                    with c2:
                        veiculo = st.text_input("Veículo", value=prod_data['veiculo'], key=f"edit_veiculo_{pid}")
                        autores = st.text_input("Autores (separados por vírgula, na ordem)", value=prod_data['autores'], key=f"edit_autores_{pid}")
                        doi = st.text_input("DOI", value=prod_data['doi'], key=f"edit_doi_{pid}")
                    descricao_atual = str(prod_data.get('descricao', '')).strip()
                    descricao = st.text_area("📝 Descrição qualitativa", value=descricao_atual,
                        placeholder="Descreva o contexto, relevância, impacto...", height=100, key=f"edit_descricao_{pid}")
                    
                    st.markdown("**👥 Co-autores do PPG**")
                    docentes_list = listar_docentes()
                    docentes_list = [d for d in docentes_list if d != user.get('name', '')]
                    
                    co_autores_atuais_str = str(prod_data.get('co_autores', '')).strip() if "co_autores" in prod_data.index else ""
                    co_autores_atuais_list = [u.strip() for u in co_autores_atuais_str.split(",") if u.strip()] if co_autores_atuais_str else []
                    
                    co_autores_nomes = []
                    for username in co_autores_atuais_list:
                        user_data = users_get(username)
                        nome = user_data["name"] if user_data else username
                        if nome in docentes_list:
                            co_autores_nomes.append(nome)
                    
                    co_autores_selecionados = st.multiselect(
                        "Selecione co-autores:", docentes_list, 
                        default=co_autores_nomes, key=f"edit_coautores_{pid}")
                    co_autores_usernames = ",".join([
                        get_docente_username_by_name(nome) for nome in co_autores_selecionados if get_docente_username_by_name(nome)])
                    
                    discente_primeiro_atual = str(prod_data.get('discente_primeiro_autor', '')).strip().lower() == "sim"
                    docente_ultimo_atual = str(prod_data.get('docente_ultimo_autor', '')).strip().lower() == "sim"
                    discente_primeiro_str, docente_ultimo_str = renderizar_checkboxes_autoria(
                        f"edit_{pid}", discente_primeiro_atual, docente_ultimo_atual)
                    
                    st.divider()
                    st.subheader("👥 Participações")
                    parts_current = df_part[df_part["producao_id"] == pid] if not df_part.empty else pd.DataFrame()
                    if not parts_current.empty:
                        st.dataframe(parts_current[["tipo_participacao","nome_participante","vinculo"]], use_container_width=True)
                    c3, c4 = st.columns(2)
                    with c3: tipo_p = st.selectbox("Tipo", TIPOS_PARTICIPACAO, key=f"part_tipo_{pid}")
                    with c4: nome_p = st.text_input("Nome", key=f"part_nome_{pid}")
                    vinc = st.text_input("Vínculo (opcional)", key=f"part_vinc_{pid}")
                    
                    c_save1, c_save2, c_cancel = st.columns([1, 1, 1])
                    with c_save1:
                        if st.form_submit_button("➕ Add participação", use_container_width=True):
                            if nome_p.strip(): participacao_submit(pid, tipo_p, nome_p, vinc); st.success("Adicionado!"); st.rerun()
                    with c_save2:
                        if st.form_submit_button("💾 Salvar", use_container_width=True):
                            if titulo.strip():
                                ok, msg = producao_update(pid, titulo, tipo, ano, veiculo, 
                                                         autores, doi, descricao, co_autores_usernames,
                                                         discente_primeiro_str, docente_ultimo_str)
                                if ok: st.session_state.pop('editing_prod_id', None); st.success(msg); st.rerun()
                    with c_cancel:
                        if st.form_submit_button("❌ Cancelar", use_container_width=True):
                            st.session_state.pop('editing_prod_id', None); st.rerun()
            st.divider()

        if 'deleting_prod_id' in st.session_state:
            pid = st.session_state['deleting_prod_id']
            prod_filtered = df_prod[df_prod["id"] == pid] if not df_prod.empty else pd.DataFrame()
            prod_data = prod_filtered.iloc[0] if not prod_filtered.empty else None
            if prod_data is not None:
                st.error(f"🗑️ Excluir: {prod_data['titulo']}")
                st.warning("⚠️ Esta ação não pode ser desfeita!")
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("Sim, excluir", type="primary", use_container_width=True, key=f"btn_confirm_del_{pid}"):
                        ok, msg = producao_delete(pid)
                        if ok: st.session_state.pop('deleting_prod_id', None); st.success(msg); st.rerun()
                with c2:
                    if st.button("Cancelar", use_container_width=True, key=f"btn_cancel_del_{pid}"):
                        st.session_state.pop('deleting_prod_id', None); st.rerun()
            st.divider()

    elif user_role == "discente":
        st.subheader(f"🎓 Minha trajetória — {user.get('name','')}")
        orientador_nome = user.get("orientador", "")
        st.info(f"Orientador: **{orientador_nome or 'não informado'}**")
        df_prod = read_df(SHEET_PROD)
        df_vinc = read_df(SHEET_VINC)
        orientador_username = ""
        df_users = read_df(SHEET_USERS)
        if not df_users.empty and orientador_nome:
            m = df_users["name"].str.lower() == orientador_nome.lower()
            if m.any(): orientador_username = df_users[m].iloc[0]["username"]
        if orientador_username:
            prods_ori = df_prod[df_prod["docente_username"] == orientador_username] if not df_prod.empty else pd.DataFrame()
            if prods_ori.empty: st.info("Orientador sem produções.")
            else:
                st.markdown("### 📚 Produções do orientador")
                meus_vinc = df_vinc[df_vinc["discente_username"] == user["username"]] if not df_vinc.empty else pd.DataFrame()
                ja_participei = set() if meus_vinc.empty else set(meus_vinc["producao_id"].tolist())
                for _, row in prods_ori.iterrows():
                    pid = row["id"]; ja = pid in ja_participei
                    with st.expander(f"{'✅' if ja else '⬜'} [{row['ano']}] {row['titulo']}"):
                        st.write(f"**Tipo:** {row['tipo']} | **Veículo:** {row['veiculo']}")
                        descricao_text = str(row.get('descricao', '')).strip()
                        if descricao_text:
                            st.markdown(f'<div class="descricao-box"><b>📝 Descrição:</b><br>{descricao_text}</div>', 
                                       unsafe_allow_html=True)
                        if ja: st.success("Você já registrou participação.")
                        else:
                            if st.button("Registrar participação", key=f"vinc_{pid}"):
                                vinculo_submit(user["username"], orientador_username, pid)
                                st.success("Registrado!"); st.rerun()
        st.divider()
        st.markdown("### 📋 Minhas participações")
        if not df_vinc.empty:
            mine = df_vinc[df_vinc["discente_username"] == user["username"]]
            if not mine.empty:
                linhas = []
                for _, v in mine.iterrows():
                    prod = df_prod[df_prod["id"] == v["producao_id"]] if not df_prod.empty else pd.DataFrame()
                    if not prod.empty:
                        p = prod.iloc[0]
                        linhas.append({"Ano": p["ano"], "Tipo": p["tipo"], "Título": p["titulo"]})
                if linhas: st.dataframe(pd.DataFrame(linhas), use_container_width=True)
        else: st.info("Sem participações.")
