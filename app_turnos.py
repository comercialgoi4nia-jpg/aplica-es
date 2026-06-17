import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore, auth as fb_auth
import json
import re
from datetime import datetime, date
import os

# ─── Configuração da Página ────────────────────────────────────────────────
st.set_page_config(
    page_title="Controle de Turnos",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── Constantes ────────────────────────────────────────────────────────────
PREFIXOS_ALVO = ["GOOC", "GOOH", "GOOL", "GOOK"]
HIERARQUIA_EQUIPES = {
    "GOOC": "Equipe Alfa - Operações Críticas",
    "GOOH": "Equipe Beta - Suporte Corporativo",
    "GOOL": "Equipe Gama - Logística de Turno",
    "GOOK": "Equipe Delta - Auditoria e Qualidade",
}

# ─── CSS Customizado ────────────────────────────────────────────────────────
st.markdown("""
<style>
  :root {
    --brand: #2E4057;
    --accent: #1D9E75;
    --accent-light: #E1F5EE;
    --danger: #E24B4A;
    --success: #639922;
    --success-light: #EAF3DE;
    --gray-100: #F1EFE8;
    --gray-200: #D3D1C7;
    --gray-400: #888780;
    --gray-600: #5F5E5A;
    --gray-800: #444441;
  }
  .topbar {
    background: #2E4057;
    color: white;
    padding: 12px 24px;
    border-radius: 10px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 20px;
  }
  .topbar h1 { margin: 0; font-size: 18px; font-weight: 500; }
  .topbar small { opacity: 0.75; font-size: 13px; }
  .metric-card {
    background: white;
    border: 1px solid #D3D1C7;
    border-radius: 10px;
    padding: 16px;
    text-align: center;
  }
  .metric-label { font-size: 11px; font-weight: 500; color: #888780; text-transform: uppercase; }
  .metric-value { font-size: 28px; font-weight: 700; color: #2E4057; }
  .badge-success {
    background: #EAF3DE; color: #3B6D11;
    padding: 3px 10px; border-radius: 20px;
    font-size: 12px; font-weight: 600;
  }
  div[data-testid="stDataFrame"] thead th { background: #2E4057 !important; color: white !important; }
  .panel-title { font-size: 14px; font-weight: 600; color: #444441; margin-bottom: 8px; }
  .stTabs [data-baseweb="tab"] { font-size: 13px; }
</style>
""", unsafe_allow_html=True)

# ─── Firebase ───────────────────────────────────────────────────────────────
@st.cache_resource
def init_firebase():
    """Inicializa o Firebase a partir dos secrets do Streamlit ou variável de ambiente."""
    if firebase_admin._apps:
        return firestore.client()
    try:
        if "firebase" in st.secrets:
            cfg = dict(st.secrets["firebase"])
            cred = credentials.Certificate(cfg)
        elif "FIREBASE_SERVICE_ACCOUNT" in os.environ:
            cfg = json.loads(os.environ["FIREBASE_SERVICE_ACCOUNT"])
            cred = credentials.Certificate(cfg)
        else:
            st.error("⚠️ Credenciais do Firebase não encontradas. Configure em `.streamlit/secrets.toml`.")
            st.stop()
        firebase_admin.initialize_app(cred)
        return firestore.client()
    except Exception as e:
        st.error(f"Erro ao inicializar Firebase: {e}")
        st.stop()

# ─── Funções Auxiliares ─────────────────────────────────────────────────────
def converter_para_minutos(horario: str) -> int:
    if not horario or not isinstance(horario, str):
        return 0
    partes = horario.split(":")
    try:
        return int(partes[0]) * 60 + int(partes[1])
    except (ValueError, IndexError):
        return 0

def formatar_minutos(minutos: int) -> str:
    if not minutos or minutos <= 0:
        return "—"
    h = minutos // 60
    m = round(minutos % 60)
    return f"{h:02d}:{m:02d}"

def calcular_diferenca_intervalo(inicio: str, fim: str) -> int:
    if not inicio or not fim:
        return 0
    min_inicio = converter_para_minutos(inicio)
    min_fim = converter_para_minutos(fim)
    if min_fim < min_inicio:
        min_fim += 1440  # virada de meia-noite
    return min_fim - min_inicio

def extrair_hora(val) -> str:
    if not val:
        return ""
    m = re.search(r"(\d{1,2}:\d{2}(?::\d{2})?)", str(val).strip())
    return m.group(1)[:5] if m else ""  # retorna só HH:MM

def extrair_data(val) -> str:
    if not val:
        return date.today().isoformat()
    m = re.search(r"(\d{4}-\d{2}-\d{2})", str(val).strip())
    return m.group(1) if m else date.today().isoformat()

def is_supervisor(email: str) -> bool:
    return "supervisor" in email.lower() or "admin" in email.lower()

# ─── Autenticação (simples via Firebase Admin) ──────────────────────────────
def tela_login(db):
    st.markdown("""
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:20px">
      <div style="background:#2E4057;border-radius:10px;width:44px;height:44px;display:flex;align-items:center;justify-content:center;font-size:22px">🔧</div>
      <div><h2 style="margin:0;color:#2E4057">Controle de Turnos</h2>
      <small style="color:#888">Painel do Usuário</small></div>
    </div>
    """, unsafe_allow_html=True)

    with st.form("login_form"):
        email = st.text_input("E-mail", placeholder="usuario@empresa.com")
        senha = st.text_input("Senha", type="password", placeholder="••••••••")
        entrar = st.form_submit_button("Entrar", use_container_width=True, type="primary")

    if entrar:
        if not email or not senha:
            st.error("Preencha e-mail e senha.")
            return
        try:
            # Verifica se o usuário existe no Firebase Auth
            user = fb_auth.get_user_by_email(email)
            # Para verificar a senha precisamos da REST API do Firebase
            # Usamos o SDK REST diretamente
            import requests
            api_key = st.secrets.get("firebase", {}).get("api_key") or os.getenv("FIREBASE_API_KEY", "")
            if not api_key:
                st.error("API Key do Firebase não configurada em secrets.")
                return
            r = requests.post(
                f"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={api_key}",
                json={"email": email, "password": senha, "returnSecureToken": True},
            )
            if r.status_code == 200:
                st.session_state["user_email"] = email
                st.session_state["logged_in"] = True
                st.rerun()
            else:
                err = r.json().get("error", {}).get("message", "Erro desconhecido")
                st.error(f"Dados inválidos: {err}")
        except Exception as e:
            st.error(f"Erro de autenticação: {e}")

# ─── Abas ───────────────────────────────────────────────────────────────────
def aba_upload(db):
    st.markdown('<div class="panel-title">📂 Importar arquivo do dia</div>', unsafe_allow_html=True)

    arquivo = st.file_uploader(
        "Selecione a planilha (.xlsx, .xls)",
        type=["xlsx", "xls"],
        label_visibility="collapsed",
    )

    if arquivo:
        try:
            df = pd.read_excel(arquivo)
        except Exception as e:
            st.error(f"Erro ao ler o arquivo: {e}")
            return

        # Filtra prefixos alvo
        col_prefixo = next((c for c in df.columns if "Prefixo" in c or "prefixo" in c), None)
        col_inicio = next((c for c in df.columns if "Início" in c or "Inicio" in c), None)
        col_intervalo = next((c for c in df.columns if "Intervalo" in c or "Fim" in c), None)

        if not col_prefixo:
            st.error("Coluna 'Prefixo' não encontrada no arquivo.")
            return

        mask = df[col_prefixo].astype(str).apply(
            lambda x: any(x.strip().startswith(p) for p in PREFIXOS_ALVO)
        )
        filtrado = df[mask].copy()

        if filtrado.empty:
            st.warning("Nenhum registro com os prefixos alvo encontrado.")
            return

        registros = []
        for _, row in filtrado.iterrows():
            inicio = extrair_hora(row.get(col_inicio, "")) if col_inicio else ""
            fim = extrair_hora(row.get(col_intervalo, "")) if col_intervalo else ""
            registros.append({
                "prefixo": str(row[col_prefixo]).strip(),
                "dataRef": extrair_data(row.get(col_inicio, "")) if col_inicio else date.today().isoformat(),
                "inicioTurno": inicio,
                "fimIntervalo": fim,
                "tempoMinutos": calcular_diferenca_intervalo(inicio, fim),
                "motivo": "",
            })

        st.session_state["dados_do_dia"] = registros

        # Exibe tabela
        exibir = []
        for r in registros:
            exibir.append({
                "Prefixo": r["prefixo"],
                "Data": r["dataRef"],
                "Início Turno": r["inicioTurno"] or "—",
                "Fim Intervalo": r["fimIntervalo"] or "—",
                "Tempo Calculado": formatar_minutos(r["tempoMinutos"]),
            })

        st.success(f"{len(registros)} registros encontrados.")
        st.dataframe(pd.DataFrame(exibir), use_container_width=True, hide_index=True)

        if st.button("☁️ Salvar no Firebase", type="primary"):
            try:
                batch = db.batch()
                for r in registros:
                    ref = db.collection("turnos").document(f"{r['dataRef']}_{r['prefixo']}")
                    batch.set(ref, {**r, "salvoEm": firestore.SERVER_TIMESTAMP})
                batch.commit()
                st.success("✅ Dados salvos com sucesso!")
            except Exception as e:
                st.error(f"Erro ao salvar: {e}")


def aba_edicao(db):
    st.markdown('<div class="panel-title">✏️ Editar Registros</div>', unsafe_allow_html=True)

    snap = db.collection("turnos").stream()
    todos = [doc.to_dict() | {"_id": doc.id} for doc in snap]

    if not todos:
        st.info("Nenhum registro encontrado.")
        return

    datas = sorted({r["dataRef"] for r in todos}, reverse=True)
    prefixos = sorted({r["prefixo"] for r in todos})

    col1, col2 = st.columns(2)
    with col1:
        filtro_data = st.selectbox("Data", ["Todas"] + datas)
    with col2:
        filtro_prefixo = st.selectbox("Prefixo", ["Todos"] + prefixos)

    filtrado = [
        r for r in todos
        if (filtro_data == "Todas" or r["dataRef"] == filtro_data)
        and (filtro_prefixo == "Todos" or r["prefixo"] == filtro_prefixo)
    ]

    if not filtrado:
        st.warning("Nenhum registro para os filtros selecionados.")
        return

    # Exibe como tabela selecionável
    df_edit = pd.DataFrame([{
        "ID": r["_id"],
        "Data": r["dataRef"],
        "Prefixo": r["prefixo"],
        "Início": r.get("inicioTurno", ""),
        "Fim Intervalo": r.get("fimIntervalo", ""),
        "Tempo": formatar_minutos(r.get("tempoMinutos", 0)),
        "Motivo": r.get("motivo", "") or "—",
    } for r in filtrado])

    st.dataframe(df_edit.drop(columns=["ID"]), use_container_width=True, hide_index=True)

    # Seleção para editar
    st.markdown("---")
    st.markdown("**Selecionar registro para editar:**")
    opcoes = {r["_id"]: f"{r['dataRef']} — {r['prefixo']}" for r in filtrado}
    selecionado_id = st.selectbox("Registro", list(opcoes.keys()), format_func=lambda x: opcoes[x])

    reg = next(r for r in filtrado if r["_id"] == selecionado_id)

    with st.form("form_edicao"):
        novo_inicio = st.text_input("Início do Turno (HH:MM)", value=reg.get("inicioTurno", ""))
        novo_fim = st.text_input("Fim do Intervalo (HH:MM)", value=reg.get("fimIntervalo", ""))
        novo_motivo = st.text_input("Motivo", value=reg.get("motivo", ""))
        salvar = st.form_submit_button("💾 Salvar Alteração", type="primary")

    if salvar:
        try:
            novo_tempo = calcular_diferenca_intervalo(novo_inicio, novo_fim)
            db.collection("turnos").document(selecionado_id).update({
                "inicioTurno": novo_inicio,
                "fimIntervalo": novo_fim,
                "tempoMinutos": novo_tempo,
                "motivo": novo_motivo,
            })
            st.success("✅ Registro atualizado com sucesso!")
            st.rerun()
        except Exception as e:
            st.error(f"Erro ao atualizar: {e}")


def aba_historico(db):
    st.markdown('<div class="panel-title">🏆 Relatório de Médias e Hierarquia de Equipes</div>', unsafe_allow_html=True)

    snap = db.collection("turnos").stream()
    todos = [doc.to_dict() for doc in snap]

    metr = {p: {"soma": 0, "qtd": 0} for p in PREFIXOS_ALVO}
    for r in todos:
        p = r.get("prefixo", "")
        t = r.get("tempoMinutos", 0)
        if t > 0 and p in metr:
            metr[p]["soma"] += t
            metr[p]["qtd"] += 1

    # Métricas resumidas
    total_turnos = sum(m["qtd"] for m in metr.values())
    total_tempo = sum(m["soma"] for m in metr.values())
    media_geral = total_tempo // total_turnos if total_turnos else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Total de Registros", total_turnos)
    c2.metric("Tempo Total", formatar_minutos(total_tempo))
    c3.metric("Média Geral", formatar_minutos(media_geral))

    st.markdown("---")

    # Tabela de hierarquia
    tabela = []
    for p in PREFIXOS_ALVO:
        d = metr[p]
        media = d["soma"] // d["qtd"] if d["qtd"] else 0
        tabela.append({
            "Prefixo": p,
            "Nome da Equipe": HIERARQUIA_EQUIPES[p],
            "Turnos": d["qtd"],
            "Média de Intervalo": formatar_minutos(media),
        })

    st.dataframe(pd.DataFrame(tabela), use_container_width=True, hide_index=True)


def aba_config():
    st.markdown('<div class="panel-title">⚙️ Configurações</div>', unsafe_allow_html=True)
    st.info("Sistema operando com credenciais seguras via `st.secrets` do Streamlit.")
    st.markdown("""
    **Como configurar:**
    
    Crie o arquivo `.streamlit/secrets.toml` com:
    ```toml
    [firebase]
    type = "service_account"
    project_id = "SEU_PROJECT_ID"
    private_key_id = "..."
    private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
    client_email = "..."
    client_id = "..."
    token_uri = "https://oauth2.googleapis.com/token"
    api_key = "SUA_WEB_API_KEY"
    ```
    """)


# ─── App Principal ──────────────────────────────────────────────────────────
def main():
    db = init_firebase()

    # Estado de sessão
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
    if "user_email" not in st.session_state:
        st.session_state["user_email"] = ""

    # Login
    if not st.session_state["logged_in"]:
        col_c = st.columns([1, 1.2, 1])[1]
        with col_c:
            tela_login(db)
        return

    email = st.session_state["user_email"]
    supervisor = is_supervisor(email)

    # Topbar
    col_a, col_b = st.columns([5, 1])
    with col_a:
        st.markdown(f"""
        <div class="topbar">
          <div style="display:flex;align-items:center;gap:10px">
            <span style="font-size:20px">🔧</span>
            <h1>Controle de Turnos</h1>
          </div>
          <small>{email}</small>
        </div>
        """, unsafe_allow_html=True)
    with col_b:
        if st.button("Sair", use_container_width=True):
            st.session_state.clear()
            st.rerun()

    # Abas conforme permissão
    if supervisor:
        tabs = st.tabs(["📋 Turnos do Dia", "✏️ Editar Registros", "📈 Histórico e Equipes", "⚙️ Configurações"])
        with tabs[0]:
            aba_upload(db)
        with tabs[1]:
            aba_edicao(db)
        with tabs[2]:
            aba_historico(db)
        with tabs[3]:
            aba_config()
    else:
        tabs = st.tabs(["📈 Histórico e Equipes"])
        with tabs[0]:
            aba_historico(db)


if __name__ == "__main__":
    main()
