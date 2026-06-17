import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore, auth as fb_auth
import json
import re
from datetime import datetime, date, timedelta
import os

# ─── Configuração da Página ────────────────────────────────────────────────
st.set_page_config(
    page_title="Controle de Turnos",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── Mapeamento completo de Prefixos → Equipe ──────────────────────────────
MAPA_PREFIXO_EQUIPE = {}

for i in ["001","002","003","004","005","006","007","008","009","010",
          "011","012","013","014","015","016","017","018","019","020",
          "021","022","023","025"]:
    MAPA_PREFIXO_EQUIPE[f"GOOL{i}M"] = "Ligação Nova (GOOL)"

for i in ["001","002","003","004","005","006","007"]:
    MAPA_PREFIXO_EQUIPE[f"GOOC{i}M"] = "Corte (GOOC)"

for i in ["001","002","003","004","010","011","012","013"]:
    MAPA_PREFIXO_EQUIPE[f"GOOK{i}M"] = "Geração Distribuída (GOOK)"

for i in ["001","002","003","004","005","006","007","008","009","010",
          "011","012","013","014","015","016","017","018"]:
    MAPA_PREFIXO_EQUIPE[f"GOOH{i}M"] = "Corte Moto (GOOH)"

for i in ["001","002","003","004","005","006","007","008","009","010",
          "011","012","013","014","015","016","017","018","019","020",
          "021","022","023","024","025"]:
    MAPA_PREFIXO_EQUIPE[f"GOOE{i}M"] = "Emergência (GOOE)"

for i in ["025","026","027","028","029","030","031","032","033","034"]:
    MAPA_PREFIXO_EQUIPE[f"GOOE{i}T"] = "Emergência (GOOE)"

for i in ["035","036","037","038","039","040","041","042","043","044"]:
    MAPA_PREFIXO_EQUIPE[f"GOOE{i}N"] = "Emergência (GOOE)"

PREFIXOS_ALVO = ["GOOC", "GOOH", "GOOL", "GOOK", "GOOE"]

def get_equipe(prefixo: str) -> str:
    return MAPA_PREFIXO_EQUIPE.get(prefixo.strip(), "Outros")

def get_grupo(prefixo: str) -> str:
    """Retorna o grupo/categoria do prefixo."""
    p = prefixo.strip().upper()
    if p.startswith("GOOL"): return "GOOL"
    if p.startswith("GOOC"): return "GOOC"
    if p.startswith("GOOK"): return "GOOK"
    if p.startswith("GOOH"): return "GOOH"
    if p.startswith("GOOE"): return "GOOE"
    return "Outros"

# ─── CSS Customizado (tema escuro compatível) ───────────────────────────────
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

  /* ── Inputs e selects no tema escuro ── */
  .stTextInput > div > div > input,
  .stSelectbox > div > div > div,
  .stTextArea > div > div > textarea,
  .stDateInput > div > div > input,
  .stTimeInput > div > div > input {
    background-color: #1e1e2e !important;
    color: #e0e0e0 !important;
    border: 1px solid #444 !important;
    border-radius: 6px !important;
  }

  /* Dropdowns do selectbox */
  div[data-baseweb="select"] > div {
    background-color: #1e1e2e !important;
    color: #e0e0e0 !important;
    border: 1px solid #444 !important;
  }
  div[data-baseweb="popover"] ul {
    background-color: #1e1e2e !important;
    color: #e0e0e0 !important;
  }
  div[data-baseweb="popover"] li:hover {
    background-color: #2a2a3e !important;
  }

  /* ── Cards e topbar ── */
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
    background: #1e1e2e;
    border: 1px solid #333;
    border-radius: 10px;
    padding: 16px;
    text-align: center;
  }
  .metric-label { font-size: 11px; font-weight: 500; color: #888780; text-transform: uppercase; }
  .metric-value { font-size: 28px; font-weight: 700; color: #1D9E75; }

  .badge-success {
    background: #EAF3DE; color: #3B6D11;
    padding: 3px 10px; border-radius: 20px;
    font-size: 12px; font-weight: 600;
  }
  .badge-warn {
    background: #FFF3CD; color: #856404;
    padding: 3px 10px; border-radius: 20px;
    font-size: 12px; font-weight: 600;
  }

  /* Cabeçalho das tabelas */
  div[data-testid="stDataFrame"] thead th {
    background: #2E4057 !important;
    color: white !important;
  }

  .panel-title { font-size: 14px; font-weight: 600; color: #aaa; margin-bottom: 8px; }
  .stTabs [data-baseweb="tab"] { font-size: 13px; }

  /* Barra de progresso customizada */
  .barra-wrapper { background: #333; border-radius: 6px; height: 14px; width: 100%; }
  .barra-fill { height: 14px; border-radius: 6px; background: #1D9E75; }
</style>
""", unsafe_allow_html=True)

# ─── Firebase ───────────────────────────────────────────────────────────────
@st.cache_resource
def init_firebase():
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
            st.error("⚠️ Credenciais do Firebase não encontradas.")
            st.stop()
        firebase_admin.initialize_app(cred)
        return firestore.client()
    except Exception as e:
        st.error(f"Erro ao inicializar Firebase: {e}")
        st.stop()

# ─── Funções Auxiliares ─────────────────────────────────────────────────────
def hms_para_minutos(horario: str) -> int:
    """Converte HH:MM ou HH:MM:SS para minutos totais."""
    if not horario or not isinstance(horario, str):
        return 0
    partes = horario.strip().split(":")
    try:
        h = int(partes[0])
        m = int(partes[1]) if len(partes) > 1 else 0
        return h * 60 + m
    except (ValueError, IndexError):
        return 0

def formatar_minutos(minutos: int) -> str:
    if not minutos or minutos <= 0:
        return "—"
    h = int(minutos) // 60
    m = round(int(minutos) % 60)
    return f"{h:02d}:{m:02d}"

def calcular_intervalo(inicio: str, fim: str) -> int:
    """Calcula diferença em minutos entre dois horários HH:MM[:SS], com virada de meia-noite."""
    if not inicio or not fim:
        return 0
    mi = hms_para_minutos(inicio)
    mf = hms_para_minutos(fim)
    if mf < mi:
        mf += 1440  # virada de meia-noite
    return mf - mi

def extrair_hora(val) -> str:
    """Extrai HH:MM:SS de uma string ou datetime."""
    if not val:
        return ""
    if isinstance(val, (datetime,)):
        return val.strftime("%H:%M:%S")
    m = re.search(r"(\d{1,2}:\d{2}(?::\d{2})?)", str(val).strip())
    return m.group(1) if m else ""

def extrair_data_str(val) -> str:
    """
    Extrai a data YYYY-MM-DD de um valor.
    PRIORIDADE: data embutida no próprio valor (timestamp/string da base),
    NÃO usa date.today() como padrão — retorna "" se não encontrar.
    """
    if not val:
        return ""
    # Se for objeto datetime/date
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, date):
        return val.isoformat()
    # Busca padrão YYYY-MM-DD na string
    m = re.search(r"(\d{4}-\d{2}-\d{2})", str(val).strip())
    return m.group(1) if m else ""

def is_supervisor(email: str) -> bool:
    return "supervisor" in email.lower() or "admin" in email.lower()

def validar_horario(h: str) -> bool:
    """Valida formato HH:MM ou HH:MM:SS."""
    if not h:
        return False
    return bool(re.match(r"^\d{1,2}:\d{2}(:\d{2})?$", h.strip()))

# ─── Autenticação ───────────────────────────────────────────────────────────
def tela_login(db):
    st.markdown("""
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:20px">
      <div style="background:#2E4057;border-radius:10px;width:44px;height:44px;
                  display:flex;align-items:center;justify-content:center;font-size:22px">🔧</div>
      <div>
        <h2 style="margin:0;color:#2E4057">Controle de Turnos</h2>
        <small style="color:#888">Painel do Usuário</small>
      </div>
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

# ─── ABA: Upload ────────────────────────────────────────────────────────────
def aba_upload(db):
    st.markdown('<div class="panel-title">📂 Importar arquivo do dia</div>', unsafe_allow_html=True)

    arquivo = st.file_uploader(
        "Selecione a planilha (.xlsx, .xls)",
        type=["xlsx", "xls"],
        label_visibility="collapsed",
    )

    if not arquivo:
        return

    try:
        df = pd.read_excel(arquivo)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        return

    # Detecção flexível de colunas
    col_prefixo = next((c for c in df.columns if re.search(r"prefixo|equipe|nom_equipe", c, re.I)), None)
    col_inicio  = next((c for c in df.columns if re.search(r"início|inicio|ini|start", c, re.I)), None)
    col_fim     = next((c for c in df.columns if re.search(r"fim|end|saída|saida|interv", c, re.I)), None)

    if not col_prefixo:
        st.error("Coluna 'Prefixo' ou 'Equipe' não encontrada. Colunas disponíveis: " + ", ".join(df.columns))
        return

    # Filtra prefixos alvo
    mask = df[col_prefixo].astype(str).apply(
        lambda x: any(x.strip().upper().startswith(p) for p in PREFIXOS_ALVO)
    )
    filtrado = df[mask].copy()

    if filtrado.empty:
        st.warning("Nenhum registro com os prefixos alvo encontrado.")
        return

    registros = []
    erros_data = []
    for _, row in filtrado.iterrows():
        prefixo = str(row[col_prefixo]).strip()

        # ── DATA: vem EXCLUSIVAMENTE da base, não do computador ──
        data_ref = ""
        if col_inicio:
            data_ref = extrair_data_str(row.get(col_inicio, ""))
        if not data_ref and col_fim:
            data_ref = extrair_data_str(row.get(col_fim, ""))
        if not data_ref:
            erros_data.append(prefixo)
            data_ref = "DATA_INVALIDA"

        inicio = extrair_hora(row.get(col_inicio, "")) if col_inicio else ""
        fim    = extrair_hora(row.get(col_fim, ""))    if col_fim    else ""
        tempo  = calcular_intervalo(inicio, fim)

        registros.append({
            "prefixo":     prefixo,
            "equipe":      get_equipe(prefixo),
            "grupo":       get_grupo(prefixo),
            "dataRef":     data_ref,
            "inicioTurno": inicio,
            "fimIntervalo": fim,
            "tempoMinutos": tempo,
            "motivo":      "",
        })

    if erros_data:
        st.warning(f"⚠️ {len(erros_data)} registro(s) sem data válida na planilha: {', '.join(erros_data[:5])}{'…' if len(erros_data)>5 else ''}")

    # Exibe tabela de prévia
    exibir = [{
        "Prefixo":        r["prefixo"],
        "Equipe":         r["equipe"],
        "Data (da base)": r["dataRef"],
        "Início Turno":   r["inicioTurno"] or "—",
        "Fim Intervalo":  r["fimIntervalo"] or "—",
        "Duração":        formatar_minutos(r["tempoMinutos"]),
    } for r in registros if r["dataRef"] != "DATA_INVALIDA"]

    st.success(f"{len(exibir)} registros válidos encontrados.")
    st.dataframe(pd.DataFrame(exibir), use_container_width=True, hide_index=True)
    st.session_state["dados_do_dia"] = registros

    if st.button("☁️ Salvar no Firebase", type="primary"):
        validos = [r for r in registros if r["dataRef"] != "DATA_INVALIDA"]
        try:
            batch = db.batch()
            for r in validos:
                doc_id = f"{r['dataRef']}_{r['prefixo']}"
                ref = db.collection("turnos").document(doc_id)
                batch.set(ref, {**r, "salvoEm": firestore.SERVER_TIMESTAMP})
            batch.commit()
            st.success(f"✅ {len(validos)} registros salvos com sucesso!")
        except Exception as e:
            st.error(f"Erro ao salvar: {e}")

# ─── ABA: Edição ────────────────────────────────────────────────────────────
def aba_edicao(db):
    st.markdown('<div class="panel-title">✏️ Gerenciamento de Intervalos e Edição</div>', unsafe_allow_html=True)

    tabs_edit = st.tabs(["⏱ Finalizar Intervalos Pendentes", "➕ Atribuir Início de Intervalo"])

    # ── Sub-aba 1: Finalizar intervalo ──────────────────────────────────────
    with tabs_edit[0]:
        snap = db.collection("turnos").stream()
        todos = [doc.to_dict() | {"_id": doc.id} for doc in snap]

        # Pendentes = têm inicioTurno mas NÃO têm fimIntervalo
        pendentes = [r for r in todos if r.get("inicioTurno") and not r.get("fimIntervalo")]

        if pendentes:
            st.warning(f"⚠️ Há **{len(pendentes)}** equipe(s) com o intervalo em andamento (Sem horário de fim).")
        else:
            st.success("✅ Nenhum intervalo pendente no momento.")

        if pendentes:
            st.markdown("**Selecione a equipe que terminou o intervalo para atualizar:**")
            opcoes = {
                r["_id"]: f"Data: {r['dataRef']} | Equipe: {r['prefixo']} (Início do Intervalo às: {r.get('inicioTurno','')})"
                for r in pendentes
            }
            selecionado_id = st.selectbox("", list(opcoes.keys()), format_func=lambda x: opcoes[x])
            reg = next(r for r in pendentes if r["_id"] == selecionado_id)

            st.markdown(f"### Atualizando: `{reg['prefixo']}`")
            col1, col2 = st.columns(2)
            with col1:
                # ── CORREÇÃO: o usuário digita o horário — NÃO usa datetime.now() ──
                fim_digitado = st.text_input(
                    "⏱ Definir Fim de Intervalo (HH:MM ou HH:MM:SS)",
                    value="",          # campo vazio — usuário preenche
                    placeholder="Ex: 14:35:00",
                    key="fim_intervalo_input",
                )
            with col2:
                motivo = st.text_input("📌 Motivo", value="", placeholder="Opcional", key="motivo_fim")

            ocorrencia = st.text_area("📋 Ocorrência", value="", placeholder="Descreva a ocorrência se houver…", key="ocorrencia_fim")

            if st.button("✅ Confirmar Fim de Intervalo", type="primary"):
                if not validar_horario(fim_digitado):
                    st.error("Horário inválido. Use o formato HH:MM ou HH:MM:SS.")
                else:
                    inicio = reg.get("inicioTurno", "")
                    tempo  = calcular_intervalo(inicio, fim_digitado)
                    try:
                        db.collection("turnos").document(selecionado_id).update({
                            "fimIntervalo":  fim_digitado,
                            "tempoMinutos":  tempo,
                            "motivo":        motivo,
                            "ocorrencia":    ocorrencia,
                            "editadoEm":     firestore.SERVER_TIMESTAMP,
                        })
                        st.success(f"✅ Fim de intervalo registrado! Duração: **{formatar_minutos(tempo)}**")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao atualizar: {e}")

    # ── Sub-aba 2: Atribuir início de intervalo ─────────────────────────────
    with tabs_edit[1]:
        snap2 = db.collection("turnos").stream()
        todos2 = [doc.to_dict() | {"_id": doc.id} for doc in snap2]

        # Sem início de intervalo
        sem_inicio = [r for r in todos2 if not r.get("inicioTurno")]

        if not sem_inicio:
            st.info("Todos os registros já possuem horário de início.")
        else:
            opcoes2 = {
                r["_id"]: f"Data: {r['dataRef']} | Equipe: {r['prefixo']}"
                for r in sem_inicio
            }
            selecionado2 = st.selectbox("Selecione o registro", list(opcoes2.keys()), format_func=lambda x: opcoes2[x])
            reg2 = next(r for r in sem_inicio if r["_id"] == selecionado2)

            st.markdown(f"### Atribuindo início para: `{reg2['prefixo']}`")

            inicio_digitado = st.text_input(
                "⏱ Horário de Início de Intervalo (HH:MM ou HH:MM:SS)",
                value="",
                placeholder="Ex: 09:10:00",
                key="inicio_intervalo_input",
            )

            if st.button("✅ Confirmar Início de Intervalo", type="primary"):
                if not validar_horario(inicio_digitado):
                    st.error("Horário inválido. Use o formato HH:MM ou HH:MM:SS.")
                else:
                    try:
                        db.collection("turnos").document(selecionado2).update({
                            "inicioTurno": inicio_digitado,
                            "editadoEm":   firestore.SERVER_TIMESTAMP,
                        })
                        st.success(f"✅ Início de intervalo registrado: {inicio_digitado}")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao atualizar: {e}")

        # ── Edição livre de qualquer registro ──
        st.markdown("---")
        st.markdown("**🔧 Edição avançada de qualquer registro**")

        snap3 = db.collection("turnos").stream()
        todos3 = [doc.to_dict() | {"_id": doc.id} for doc in snap3]

        if todos3:
            datas    = sorted({r["dataRef"] for r in todos3}, reverse=True)
            prefixos = sorted({r["prefixo"] for r in todos3})

            c1, c2 = st.columns(2)
            with c1: filtro_data = st.selectbox("Data", ["Todas"] + datas, key="edit_data")
            with c2: filtro_pref = st.selectbox("Prefixo", ["Todos"] + prefixos, key="edit_pref")

            filtrado3 = [
                r for r in todos3
                if (filtro_data == "Todas" or r.get("dataRef") == filtro_data)
                and (filtro_pref == "Todos" or r.get("prefixo") == filtro_pref)
            ]

            if filtrado3:
                opcoes3 = {r["_id"]: f"{r.get('dataRef','')} — {r.get('prefixo','')}" for r in filtrado3}
                sel3 = st.selectbox("Registro", list(opcoes3.keys()), format_func=lambda x: opcoes3[x], key="edit_reg")
                reg3 = next(r for r in filtrado3 if r["_id"] == sel3)

                with st.form("form_edicao_avancada"):
                    e1, e2 = st.columns(2)
                    with e1:
                        novo_inicio = st.text_input("Início do Turno", value=reg3.get("inicioTurno", ""), placeholder="HH:MM:SS")
                    with e2:
                        novo_fim    = st.text_input("Fim do Intervalo", value=reg3.get("fimIntervalo", ""), placeholder="HH:MM:SS")
                    novo_motivo     = st.text_input("Motivo", value=reg3.get("motivo", ""))
                    nova_ocorrencia = st.text_area("Ocorrência", value=reg3.get("ocorrencia", ""))
                    salvar3 = st.form_submit_button("💾 Salvar Alteração", type="primary")

                if salvar3:
                    novo_tempo = calcular_intervalo(novo_inicio, novo_fim)
                    try:
                        db.collection("turnos").document(sel3).update({
                            "inicioTurno":  novo_inicio,
                            "fimIntervalo": novo_fim,
                            "tempoMinutos": novo_tempo,
                            "motivo":       novo_motivo,
                            "ocorrencia":   nova_ocorrencia,
                            "editadoEm":    firestore.SERVER_TIMESTAMP,
                        })
                        st.success(f"✅ Registro atualizado! Duração: **{formatar_minutos(novo_tempo)}**")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao atualizar: {e}")

# ─── ABA: Histórico e Análises ──────────────────────────────────────────────
def aba_historico(db):
    st.markdown('<div class="panel-title">📈 Relatórios, Médias e Análise de Frequência</div>', unsafe_allow_html=True)

    snap = db.collection("turnos").stream()
    todos = [doc.to_dict() for doc in snap]

    if not todos:
        st.info("Nenhum registro encontrado na base.")
        return

    df = pd.DataFrame(todos)
    df["tempoMinutos"] = pd.to_numeric(df.get("tempoMinutos", 0), errors="coerce").fillna(0)

    # ── Filtro de período ──
    datas_disp = sorted(df["dataRef"].dropna().unique(), reverse=True)
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        data_ini = st.selectbox("De (data)", ["Todas"] + datas_disp, key="hist_di")
    with col_f2:
        data_fim_sel = st.selectbox("Até (data)", ["Todas"] + datas_disp, key="hist_df")
    with col_f3:
        grupo_sel = st.selectbox("Prefixo", ["Todos"] + sorted(df["grupo"].dropna().unique().tolist()), key="hist_grupo")

    df_f = df.copy()
    if data_ini != "Todas":
        df_f = df_f[df_f["dataRef"] >= data_ini]
    if data_fim_sel != "Todas":
        df_f = df_f[df_f["dataRef"] <= data_fim_sel]
    if grupo_sel != "Todos":
        df_f = df_f[df_f["grupo"] == grupo_sel]

    if df_f.empty:
        st.warning("Nenhum dado para os filtros selecionados.")
        return

    # ── Métricas gerais ──
    total_registros = len(df_f)
    com_intervalo   = df_f[df_f["tempoMinutos"] > 0]
    total_tempo_min = int(com_intervalo["tempoMinutos"].sum())
    media_geral_min = int(com_intervalo["tempoMinutos"].mean()) if not com_intervalo.empty else 0
    pendentes_count = len(df_f[df_f["fimIntervalo"].isna() | (df_f["fimIntervalo"] == "")])

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📋 Registros", total_registros)
    c2.metric("⏱ Tempo Total", formatar_minutos(total_tempo_min))
    c3.metric("📊 Média Geral", formatar_minutos(media_geral_min))
    c4.metric("⚠️ Pendentes", pendentes_count)

    st.markdown("---")

    # ── Tabela por grupo ──
    st.markdown("#### 🏆 Ranking por Grupo (Média de Intervalo)")
    grupos = df_f.groupby("grupo")
    tabela_grupos = []
    for grp, sub in grupos:
        com = sub[sub["tempoMinutos"] > 0]
        media = int(com["tempoMinutos"].mean()) if not com.empty else 0
        tabela_grupos.append({
            "Grupo":           grp,
            "Total Registros": len(sub),
            "Com Intervalo":   len(com),
            "Sem Intervalo":   len(sub) - len(com),
            "Média Intervalo": formatar_minutos(media),
            "Tempo Total":     formatar_minutos(int(com["tempoMinutos"].sum())),
        })
    tabela_grupos.sort(key=lambda x: hms_para_minutos(x["Média Intervalo"].replace("—","0:00")), reverse=True)
    st.dataframe(pd.DataFrame(tabela_grupos), use_container_width=True, hide_index=True)

    st.markdown("---")

    # ── Análise percentual de frequência por prefixo ──
    st.markdown("#### 📊 Frequência e Participação Percentual por Prefixo")
    total_com_intervalo = len(com_intervalo)

    tabela_freq = []
    for pref, sub in df_f.groupby("prefixo"):
        com_pref = sub[sub["tempoMinutos"] > 0]
        n        = len(com_pref)
        media    = int(com_pref["tempoMinutos"].mean()) if not com_pref.empty else 0
        pct_freq = (n / total_com_intervalo * 100) if total_com_intervalo > 0 else 0
        tabela_freq.append({
            "Prefixo":        pref,
            "Equipe":         get_equipe(pref),
            "Qtd Intervalos": n,
            "% Frequência":   round(pct_freq, 1),
            "Média Intervalo":formatar_minutos(media),
        })
    tabela_freq.sort(key=lambda x: x["% Frequência"], reverse=True)

    df_freq = pd.DataFrame(tabela_freq)
    st.dataframe(df_freq, use_container_width=True, hide_index=True)

    # Barras de frequência visual (top 10)
    st.markdown("##### 🔢 Top 10 Prefixos por Frequência de Intervalo")
    top10 = tabela_freq[:10]
    max_pct = max((r["% Frequência"] for r in top10), default=1)
    for row in top10:
        pct = row["% Frequência"]
        largura = int((pct / max_pct) * 100) if max_pct > 0 else 0
        st.markdown(
            f"<div style='display:flex;align-items:center;gap:10px;margin:4px 0'>"
            f"<span style='width:100px;font-size:12px;color:#ccc'>{row['Prefixo']}</span>"
            f"<div class='barra-wrapper' style='flex:1'>"
            f"  <div class='barra-fill' style='width:{largura}%'></div>"
            f"</div>"
            f"<span style='width:50px;font-size:12px;color:#1D9E75;text-align:right'>{pct}%</span>"
            f"</div>",
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # ── Distribuição por dia ──
    st.markdown("#### 📅 Registros por Data")
    por_data = df_f.groupby("dataRef").agg(
        Total=("prefixo", "count"),
        Com_Intervalo=("tempoMinutos", lambda x: (x > 0).sum()),
    ).reset_index().rename(columns={"dataRef": "Data", "Com_Intervalo": "Com Intervalo"})
    st.dataframe(por_data.sort_values("Data", ascending=False), use_container_width=True, hide_index=True)

    # ── Tabela detalhada ──
    with st.expander("🔍 Ver todos os registros detalhados"):
        detalhado = df_f[[c for c in ["dataRef","prefixo","equipe","inicioTurno","fimIntervalo","tempoMinutos","motivo","ocorrencia"] if c in df_f.columns]].copy()
        detalhado["tempoMinutos"] = detalhado["tempoMinutos"].apply(lambda x: formatar_minutos(int(x)) if x > 0 else "—")
        st.dataframe(detalhado.sort_values(["dataRef","prefixo"], ascending=[False,True]), use_container_width=True, hide_index=True)

# ─── ABA: Configurações ─────────────────────────────────────────────────────
def aba_config():
    st.markdown('<div class="panel-title">⚙️ Configurações</div>', unsafe_allow_html=True)
    st.info("Sistema operando com credenciais seguras via `st.secrets` do Streamlit.")
    st.markdown("""
**Como configurar o arquivo `.streamlit/secrets.toml`:**

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

**Observações importantes:**
- A **data dos registros** é extraída exclusivamente da planilha importada, nunca do relógio do servidor.
- **Horários de início/fim de intervalo** devem ser digitados manualmente — o sistema não preenche automaticamente com a hora atual.
- O ID do documento no Firestore é composto por `DATA_PREFIXO` para evitar duplicatas.
    """)

# ─── App Principal ──────────────────────────────────────────────────────────
def main():
    db = init_firebase()

    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
    if "user_email" not in st.session_state:
        st.session_state["user_email"] = ""

    if not st.session_state["logged_in"]:
        col_c = st.columns([1, 1.2, 1])[1]
        with col_c:
            tela_login(db)
        return

    email      = st.session_state["user_email"]
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

    if supervisor:
        tabs = st.tabs(["📋 Turnos do Dia", "✏️ Editar Registros", "📈 Histórico e Análises", "⚙️ Configurações"])
        with tabs[0]: aba_upload(db)
        with tabs[1]: aba_edicao(db)
        with tabs[2]: aba_historico(db)
        with tabs[3]: aba_config()
    else:
        tabs = st.tabs(["📈 Histórico e Análises"])
        with tabs[0]: aba_historico(db)


if __name__ == "__main__":
    main()
