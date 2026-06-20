# app.py — Sistema de Monitoramento de Produção do PPG (v4.4 - COMPLETO)
# Streamlit + Google Sheets + E-mails
# =========================================================

import os, time, base64, uuid, hashlib, hmac, smtplib
from email.message import EmailMessage
from datetime import datetime

import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError, WorksheetNotFound, SpreadsheetNotFound
from PIL import Image

# ---------------------------------------------------------
# CONFIG
# ---------------------------------------------------------
st.set_page_config(page_title="PPG — Monitor de Produção", layout="wide")

try:
    banner = Image.open("banner_ppg.png")
    st.image(banner, use_container_width=True)
except Exception:
    pass

st.title("🎓 Sistema de Monitoramento de Produção do PPG")

SPREADSHEET_ID = st.secrets.get("GSHEET_SPREADSHEET_ID", "")
if not SPREADSHEET_ID:
    st.error("Configure GSHEET_SPREADSHEET_ID nos Secrets.")
    st.stop()

SHEET_USERS = "users"
SHEET_CAD   = "cadastro_requests"
SHEET_PROD  = "producoes"
SHEET_PART  = "participacoes"
SHEET_VINC  = "vinculos_discentes"

HEADERS_USERS = ["username", "name", "email", "role", "orientador", "password_hash", "created_at"]
HEADERS_CAD   = ["id", "name", "username", "email", "role", "orientador",
                 "password_hash", "status", "created_at", "reviewed_at",
                 "reviewed_by", "review_reason"]
HEADERS_PROD  = ["id", "docente_username", "titulo", "tipo", "ano",
                 "veiculo", "autores", "doi", "created_at"]
HEADERS_PART  = ["id", "producao_id", "tipo_participacao",
                 "nome_participante", "vinculo", "created_at"]
HEADERS_VINC  = ["id", "discente_username", "orientador_username",
                 "producao_id", "created_at"]

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

# ---------------------------------------------------------
# CSS
# ---------------------------------------------------------
st.markdown("""
<style>
.block-card { background:#f7f7f9; border:1px solid #e7e7ee; padding:14px; border-radius:10px; margin-bottom:10px;}
.big-status-ok  { background:#b9f6ca; padding:12px; font-size:18px; border-radius:10px; text-align:center; }
.big-status-bad { background:#ff8a80; padding:12px; font-size:18px; border-radius:10px; text-align:center; }
.small-muted { color:#666; font-size:0.9rem; }
</style>
""", unsafe_allow_html=True)

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
    for attempt in range(4):
        try:
            return gclient().open_by_key(SPREADSHEET_ID)
        except APIError as e:
            last_err = e
            status, _ = _extract_api_error_info(e)
            if _retryable(status):
                time.sleep(1.5 * (attempt + 1)); continue
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
    """Garante que o cabeçalho está correto SEM APAGAR DADOS."""
    try:
        vals = ws_obj.get_all_values()
        if not vals:
            ws_obj.append_row(headers)
            return
        if len(vals) >= 1 and vals[0] != headers:
            ws_obj.update("1:1", [headers])
    except APIError as e:
        st.warning(f"⚠️ Erro ao verificar aba '{ws_obj.title}': {e}")

def ensure_worksheets(sh):
    wmap = get_worksheets_map(sh)
    targets = [
        (SHEET_USERS, HEADERS_USERS), (SHEET_CAD, HEADERS_CAD),
        (SHEET_PROD, HEADERS_PROD),   (SHEET_PART, HEADERS_PART),
        (SHEET_VINC, HEADERS_VINC),
    ]
    for title, headers in targets:
        if title not in wmap:
            ws_obj = sh.add_worksheet(title=title, rows=2000, cols=max(12, len(headers)))
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

@st.cache_data(ttl=15, show_spinner=False)
def read_df(sheet_name):
    last_err = None
    for attempt in range(4):
        try:
            w = ws(sheet_name)
            values = w.get_all_values()
            if not values: return pd.DataFrame()
            return pd.DataFrame(values[1:], columns=values[0]).fillna("")
        except APIError as e:
            last_err = e
            if _retryable(_extract_api_error_info(e)[0]):
                time.sleep(1.5 * (attempt + 1)); continue
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

def role_of(user): return str(user.get("role", "")).strip().lower()

# ---------------------------------------------------------
# CADASTRO & PRODUÇÕES (CRUD)
# ---------------------------------------------------------
def listar_docentes():
    df = read_df(SHEET_USERS)
    if df.empty: return []
    return sorted(df[df["role"].str.lower() == "docente"]["name"].tolist())

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

def producao_submit(docente_username, titulo, tipo, ano, veiculo, autores, doi):
    prod_id = str(uuid.uuid4())
    ws(SHEET_PROD).append_row([prod_id, docente_username, titulo.strip(), tipo, str(ano),
        veiculo.strip(), autores.strip(), doi.strip(), datetime.utcnow().isoformat(timespec="seconds")])
    clear_cache()
    return prod_id

def producao_update(producao_id, titulo, tipo, ano, veiculo, autores, doi):
    df = read_df(SHEET_PROD)
    idx = df.index[df["id"] == producao_id]
    if len(idx) == 0: return False, "Produção não encontrada."
    i = int(idx[0])
    w = ws(SHEET_PROD)
    row_number = i + 2
    range_name = f"C{row_number}:H{row_number}"
    values = [[titulo.strip(), tipo, str(ano), veiculo.strip(), autores.strip(), doi.strip()]]
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
# SESSION
# ---------------------------------------------------------
if "logged" not in st.session_state: st.session_state.logged = False
if "user"   not in st.session_state: st.session_state.user   = {}

# ---------------------------------------------------------
# LOGIN / CADASTRO
# ---------------------------------------------------------
if not st.session_state.logged:
    tab_login, tab_cad = st.tabs(["🔑 Login", "📝 Cadastro"])
    with tab_login:
        st.markdown("<div class='block-card'>", unsafe_allow_html=True)
        u = st.text_input("Usuário", key="login_user")
        p = st.text_input("Senha", type="password", key="login_pass")
        if st.button("Entrar", use_container_width=True, key="btn_login"):
            ok, res = authenticate(u, p)
            if ok:
                st.session_state.logged = True; st.session_state.user = res; st.rerun()
            else: st.error(res)
        st.markdown("</div>", unsafe_allow_html=True)

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
    st.stop()

# ---------------------------------------------------------
# ÁREA LOGADA
# ---------------------------------------------------------
user = st.session_state.user
st.success(f"Logado como **{user.get('name','')}**  | perfil: **{role_of(user)}**")

# =========================================================
# PAINEL ADMIN (COORDENADOR)
# =========================================================
if role_of(user) == "admin":
    st.subheader("🛠️ Painel do Coordenador")
    
    # Carregar dados
    df_users = read_df(SHEET_USERS)
    df_cad = read_df(SHEET_CAD)
    df_prod = read_df(SHEET_PROD)
    df_part = read_df(SHEET_PART)
    df_vinc = read_df(SHEET_VINC)

    # Criar abas
    t1, t2, t3, t4, t5, t6 = st.tabs([
        "👥 Usuários", "👤 Cadastros pendentes", "📚 Produções (geral)", 
        "📊 Resumo por ano", "⚙️ Histórico", "➕ Cadastrar Produção"
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
        docentes_df = df_users[df_users["role"].str.lower() == "docente"] if not df_users.empty else pd.DataFrame()
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
                    autores = st.text_input("Autores", key="admin_prod_autores")
                    doi = st.text_input("DOI (opcional)", key="admin_prod_doi")
                submitted = st.form_submit_button("💾 Cadastrar", use_container_width=True)
                if submitted:
                    if not titulo.strip(): st.error("Título é obrigatório.")
                    else:
                        producao_submit(selected_username, titulo, tipo, ano, veiculo, autores, doi)
                        st.success(f"Produção cadastrada para {selected_label.split(' (')[0]}!")
                        st.rerun()
    st.divider()

# =========================================================
# PAINEL DOCENTE
# =========================================================
if role_of(user) == "docente":
    st.subheader(f"📚 Minhas produções — {user.get('name','')}")
    df_prod = read_df(SHEET_PROD)
    df_part = read_df(SHEET_PART)
    mine = df_prod[df_prod["docente_username"] == user["username"]] if not df_prod.empty else pd.DataFrame()

    st.markdown("### 📋 Minhas produções")
    tem_producoes = False
    for ano in ANOS:
        subset = mine[mine["ano"].astype(str).str.strip() == ano] if not mine.empty else pd.DataFrame()
        if not subset.empty:
            tem_producoes = True
            st.markdown(f"#### 📅 {ano}")
            for _, row in subset.iterrows():
                with st.expander(f"**{row['titulo']}** — {row['tipo']}"):
                    st.write(f"**Veículo:** {row['veiculo']}")
                    st.write(f"**Autores:** {row['autores']}")
                    st.write(f"**DOI:** {row['doi'] or '—'}")
                    parts = df_part[df_part["producao_id"] == row["id"]] if not df_part.empty else pd.DataFrame()
                    if not parts.empty:
                        st.write("**Participações:**")
                        st.dataframe(parts[["tipo_participacao","nome_participante","vinculo"]], use_container_width=True)
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button("✏️ Editar", key=f"edit_{row['id']}", use_container_width=True):
                            st.session_state['editing_prod_id'] = row['id']; st.rerun()
                    with col2:
                        if st.button("🗑️ Excluir", key=f"del_{row['id']}", use_container_width=True):
                            st.session_state['deleting_prod_id'] = row['id']; st.rerun()
    
    if not tem_producoes: st.info("Nenhuma produção cadastrada.")
    st.divider()

    # Edição
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
                    autores = st.text_input("Autores", value=prod_data['autores'], key=f"edit_autores_{pid}")
                    doi = st.text_input("DOI", value=prod_data['doi'], key=f"edit_doi_{pid}")
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
                            ok, msg = producao_update(pid, titulo, tipo, ano, veiculo, autores, doi)
                            if ok: st.session_state.pop('editing_prod_id', None); st.success(msg); st.rerun()
                with c_cancel:
                    if st.form_submit_button("❌ Cancelar", use_container_width=True):
                        st.session_state.pop('editing_prod_id', None); st.rerun()
        st.divider()

    # Exclusão
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

# =========================================================
# PAINEL DISCENTE
# =========================================================
if role_of(user) == "discente":
    st.subheader(f"🎓 Minha trajetória — {user.get('name','')}")
    orientador_nome = user.get("orientador", "")
    st.info(f"Orientador: **{orientador_nome or 'não informado'}**")
    df_prod = read_df(SHEET_PROD); df_vinc = read_df(SHEET_VINC)
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

# =========================================================
# LOGOUT
# =========================================================
st.divider()
if st.button("🚪 Sair", key="btn_logout"):
    st.session_state.logged = False; st.session_state.user = {}; st.rerun()
