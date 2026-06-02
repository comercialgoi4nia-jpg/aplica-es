import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from io import BytesIO
from datetime import datetime, timedelta

st.set_page_config(page_title="Comparador de Bases", page_icon="🔍", layout="wide")

# ─── mapeamento fixo de colunas por sistema ──────────────────────────────────
COLUNAS_CBILL = {
    "servico": "Serviço",
    "data":    "Prazo de execução",
    "tipo":    "Tipo Serviço",
}

COLUNAS_OPER = {
    "servico":  "Numero",
    "data":     "Data/Hora Limite",
    "tipo":     "Subtipo",
    "situacao": "Situação",
}

ALIASES_COLUNAS = {
    "Data/Hora Limite":  ["Data/Hora Limite", "Data Limite", "Data limite", "DataHora Limite"],
    "Numero":            ["Numero", "Número", "numero", "número"],
    "Subtipo":           ["Subtipo", "subtipo"],
    "Situação":          ["Situação", "Situacao", "situação", "situacao", "SITUAÇÃO"],
    "Tipo Serviço":      ["Tipo Serviço", "Tipo Servico", "Tipo de Serviço", "Tipo de Servico"],
    "Prazo de execução": ["Prazo de execução", "Prazo de Execução", "Prazo execução"],
    "Serviço":           ["Serviço", "Servico", "serviço"],
}

# Lista de tipos excluídos do Oper
TIPOS_EXCLUIDOS_OPER = [
    "RESTABELECIMENTO FORNEC. NORMAL",
    "RESTABELECIMENTO FORNEC. NORMAL - MUDANÇA TITULARIDADE",
]

# ─── helpers ─────────────────────────────────────────────────────────────────

def ler_excel(arquivo) -> pd.DataFrame:
    nome = arquivo.name.lower()
    conteudo = arquivo.read()
    arquivo.seek(0)

    amostra = conteudo[:10]
    if amostra.startswith(b"<") or amostra.startswith(b"\xef\xbb\xbf<"):
        for header_row in range(0, 8):
            for flavor in ("lxml", "html5lib"):
                try:
                    todas = pd.read_html(BytesIO(conteudo), header=header_row, flavor=flavor)
                except Exception:
                    continue
                for df in todas:
                    df.columns = [str(c).strip() for c in df.columns]
                    colunas_validas = [
                        c for c in df.columns
                        if c and not c.startswith("Unnamed") and c.lower() != "nan"
                    ]
                    if len(colunas_validas) >= 3:
                        df = df.dropna(how="all").reset_index(drop=True)
                        return df
        return pd.DataFrame()

    if nome.endswith(".xls"):
        return pd.read_excel(BytesIO(conteudo), engine="xlrd")
    else:
        return pd.read_excel(BytesIO(conteudo), engine="openpyxl")


def resolver_coluna(df: pd.DataFrame, col_canonica: str) -> str | None:
    aliases = ALIASES_COLUNAS.get(col_canonica, [col_canonica])
    for alias in aliases:
        if alias in df.columns:
            return alias
    col_lower = col_canonica.lower()
    for c in df.columns:
        if c.lower() == col_lower:
            return c
    return None


def resolver_colunas(df: pd.DataFrame, mapa: dict) -> dict:
    return {
        chave: resolver_coluna(df, col_canonica) or col_canonica
        for chave, col_canonica in mapa.items()
    }


def validar_colunas_obrigatorias(df: pd.DataFrame, cols: dict, nome_arquivo: str) -> bool:
    obrigatorias = {k: v for k, v in cols.items() if k != "situacao"}
    ausentes = [v for v in obrigatorias.values() if v not in df.columns]
    if ausentes:
        st.error(
            f"❌ **{nome_arquivo}** — colunas não encontradas: `{'`, `'.join(ausentes)}`\n\n"
            f"Colunas disponíveis: `{'`, `'.join(df.columns.tolist())}`"
        )
        return False
    return True


def normalizar_datas(df: pd.DataFrame, col: str) -> pd.DataFrame:
    df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")
    return df


def deduplicar(df: pd.DataFrame, col_servico: str, nome_base: str) -> pd.DataFrame:
    total_antes = len(df)
    df = df.drop_duplicates(subset=[col_servico], keep="first").reset_index(drop=True)
    removidos = total_antes - len(df)
    if removidos > 0:
        st.info(f"🔁 **{nome_base}:** {removidos} duplicata(s) removida(s) ({total_antes} → {len(df)} registros)")
    return df


def exportar_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Divergentes")
        ws = writer.sheets["Divergentes"]
        for col_cells in ws.columns:
            max_len = max(len(str(c.value)) if c.value else 0 for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = max(max_len + 4, 14)
    return buf.getvalue()


def gauge(titulo: str, valor: float, cor: str) -> go.Figure:
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=valor,
        number={"suffix": "%", "font": {"size": 28}},
        title={"text": titulo, "font": {"size": 14}},
        gauge={
            "axis": {"range": [0, 100], "tickwidth": 1},
            "bar":  {"color": cor},
            "steps": [
                {"range": [0,  30], "color": "#d4edda"},
                {"range": [30, 60], "color": "#fff3cd"},
                {"range": [60, 100], "color": "#f8d7da"},
            ],
            "threshold": {
                "line": {"color": "black", "width": 3},
                "thickness": 0.75,
                "value": valor,
            },
        },
    ))
    fig.update_layout(margin=dict(t=60, b=20, l=20, r=20), height=230)
    return fig


def grafico_barras(n_cbill: int, n_oper: int, s_cbill: int, s_oper: int) -> go.Figure:
    categorias = ["Total Cbill", "Total Oper", "Só no Cbill", "Só no Oper"]
    valores    = [n_cbill,       n_oper,        s_cbill,       s_oper]
    cores      = ["#4e8cff", "#ff8c42", "#c0392b", "#e67e22"]

    fig = go.Figure(go.Bar(
        x=categorias,
        y=valores,
        marker_color=cores,
        text=valores,
        textposition="outside",
    ))
    fig.update_layout(
        title="Visão geral — Cbill vs Oper",
        yaxis_title="Qtd. de serviços",
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(t=50, b=40, l=40, r=20),
        height=350,
    )
    return fig

# ─── sidebar com abas ────────────────────────────────────────────────────────

aba_config, aba_dash = st.sidebar.tabs(["⚙️ Configurações", "📊 Dashboard"])

with aba_config:
    # ── Modo de filtro de data ──
    modo_data = st.radio(
        "Modo de filtro de data",
        ["📅 Dia único", "📆 Intervalo de datas"],
        horizontal=True,
    )

    if modo_data == "📅 Dia único":
        data_inicio = st.date_input("Data a comparar", value=datetime.today())
        data_fim = data_inicio
        label_periodo = data_inicio.strftime("%d/%m/%Y")
    else:
        col_di, col_df = st.columns(2)
        with col_di:
            data_inicio = st.date_input("Data início", value=datetime.today())
        with col_df:
            data_fim = st.date_input("Data fim", value=datetime.today() + timedelta(days=6))
        if data_fim < data_inicio:
            st.error("⚠️ A data fim deve ser igual ou posterior à data início.")
            st.stop()
        label_periodo = f"{data_inicio.strftime('%d/%m/%Y')} → {data_fim.strftime('%d/%m/%Y')}"

    st.markdown("---")
    st.markdown("**Colunas utilizadas:**")
    st.markdown("🔵 **Cbill:** `Serviço` · `Prazo de execução` · `Tipo Serviço`")
    st.markdown("🟠 **Oper:** `Numero` · `Data/Hora Limite` · `Subtipo` · `Situação`")
    st.markdown("---")
    st.markdown("**Filtros automáticos Oper:**")
    for t in TIPOS_EXCLUIDOS_OPER:
        st.markdown(f"🚫 `{t}`")
    st.markdown("---")
    st.info(
        "💡 **Padrão de nomes esperado:**\n"
        "`base_DD.MM_Cbill.xlsx`\n"
        "`base_DD.MM_oper.xls` _(Comercial)_\n"
        "`base_gd_DD.MM_oper.xls` _(GD)_"
    )

# placeholder para o dashboard (preenchido após processamento)
dash_placeholder = aba_dash.empty()
dash_placeholder.info("📂 Faça upload das bases para visualizar o dashboard.")

# ─── título principal ────────────────────────────────────────────────────────

st.title("🔍 Comparador de Bases — Serviços Divergentes")
st.markdown("Carregue as bases **Cbill** e **Oper** para identificar serviços divergentes por data limite.")

# ─── upload ───────────────────────────────────────────────────────────────────

col1, col2 = st.columns(2)

with col1:
    st.subheader("📂 Base Cbill")
    arquivo_cbill = st.file_uploader("Selecione a base Cbill", type=["xlsx", "xls"], key="cbill")

with col2:
    st.subheader("📂 Base Oper")
    arquivo_oper_com = st.file_uploader("Base Oper — Comercial", type=["xlsx", "xls"], key="oper_com")
    arquivo_oper_gd  = st.file_uploader("Base Oper — GD (opcional)", type=["xlsx", "xls"], key="oper_gd")

# ─── processamento ────────────────────────────────────────────────────────────

if arquivo_cbill and arquivo_oper_com:
    try:
        with st.spinner("Carregando bases..."):
            df_cbill    = ler_excel(arquivo_cbill)
            df_oper_com = ler_excel(arquivo_oper_com)
            df_oper_gd  = ler_excel(arquivo_oper_gd) if arquivo_oper_gd else pd.DataFrame()

        # ── resolve colunas ──
        cols_cbill    = resolver_colunas(df_cbill,    COLUNAS_CBILL)
        cols_oper_com = resolver_colunas(df_oper_com, COLUNAS_OPER)

        if not validar_colunas_obrigatorias(df_cbill, cols_cbill, arquivo_cbill.name):
            st.stop()
        if not validar_colunas_obrigatorias(df_oper_com, cols_oper_com, arquivo_oper_com.name):
            st.stop()

        # ── deduplica ──
        df_cbill    = deduplicar(df_cbill,    cols_cbill["servico"],    "Cbill")
        df_oper_com = deduplicar(df_oper_com, cols_oper_com["servico"], "Oper Comercial")

        # ── normaliza datas ──
        df_cbill    = normalizar_datas(df_cbill,    cols_cbill["data"])
        df_oper_com = normalizar_datas(df_oper_com, cols_oper_com["data"])

        # ── processa e une GD se enviada ──
        if not df_oper_gd.empty:
            cols_oper_gd = resolver_colunas(df_oper_gd, COLUNAS_OPER)
            if not validar_colunas_obrigatorias(df_oper_gd, cols_oper_gd, arquivo_oper_gd.name):
                st.stop()
            df_oper_gd = deduplicar(df_oper_gd, cols_oper_gd["servico"], "Oper GD")
            df_oper_gd = normalizar_datas(df_oper_gd, cols_oper_gd["data"])

            rename_gd = {
                cols_oper_gd["servico"]:  cols_oper_com["servico"],
                cols_oper_gd["data"]:     cols_oper_com["data"],
                cols_oper_gd["tipo"]:     cols_oper_com["tipo"],
            }
            col_sit_gd  = cols_oper_gd.get("situacao")
            col_sit_com = cols_oper_com.get("situacao")
            if col_sit_gd and col_sit_gd in df_oper_gd.columns and col_sit_com:
                rename_gd[col_sit_gd] = col_sit_com

            df_oper_gd   = df_oper_gd.rename(columns=rename_gd)
            df_oper_full = pd.concat([df_oper_com, df_oper_gd], ignore_index=True)
        else:
            df_oper_full = df_oper_com.copy()

        # ── filtra pelo intervalo de datas ──
        ts_inicio = pd.Timestamp(data_inicio)
        ts_fim    = pd.Timestamp(data_fim) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

        base_cbill = df_cbill[
            (df_cbill[cols_cbill["data"]] >= ts_inicio) &
            (df_cbill[cols_cbill["data"]] <= ts_fim)
        ].copy()

        base_oper = df_oper_full[
            (df_oper_full[cols_oper_com["data"]] >= ts_inicio) &
            (df_oper_full[cols_oper_com["data"]] <= ts_fim)
        ].copy()

        # ── remove tipos excluídos apenas do Oper ──
        col_tipo_oper = cols_oper_com["tipo"]
        tipos_upper   = [t.upper() for t in TIPOS_EXCLUIDOS_OPER]
        base_oper = base_oper[
            ~base_oper[col_tipo_oper].astype(str).str.strip().str.upper().isin(tipos_upper)
        ].copy()

        total_cbill = len(base_cbill)
        total_oper  = len(base_oper)

        # ── métricas ──
        st.markdown("---")
        st.markdown(f"### 📅 Período analisado: **{label_periodo}**")
        m1, m2, m3 = st.columns(3)
        m1.metric("Serviços Cbill", total_cbill)
        m2.metric("Serviços Oper (sem Restab.)", total_oper)
        m3.metric("Diferença", abs(total_cbill - total_oper))

        # ── identifica divergências ──
        col_srv_cbill  = cols_cbill["servico"]
        col_srv_oper   = cols_oper_com["servico"]
        col_tipo_cbill = cols_cbill["tipo"]
        col_sit_oper   = cols_oper_com.get("situacao")
        tem_situacao   = col_sit_oper and col_sit_oper in base_oper.columns

        srvs_cbill = set(base_cbill[col_srv_cbill].astype(str).str.strip())
        srvs_oper  = set(base_oper[col_srv_oper].astype(str).str.strip())

        apenas_cbill = srvs_cbill - srvs_oper
        apenas_oper  = srvs_oper  - srvs_cbill

        registros = []

        for srv in sorted(apenas_cbill):
            linhas = base_cbill[base_cbill[col_srv_cbill].astype(str).str.strip() == srv]
            for _, row in linhas.iterrows():
                registros.append({
                    "servico":        row[col_srv_cbill],
                    "tipo_servico":   str(row.get(col_tipo_cbill, "")).strip(),
                    "situacao":       "—",
                    "data_limite":    row[cols_cbill["data"]].date() if pd.notna(row[cols_cbill["data"]]) else data_inicio,
                    "ausente_em":     "Oper",
                    "sistema_origem": "Cbill",
                })

        for srv in sorted(apenas_oper):
            linhas = base_oper[base_oper[col_srv_oper].astype(str).str.strip() == srv]
            for _, row in linhas.iterrows():
                sit = str(row[col_sit_oper]).strip() if tem_situacao else "—"
                registros.append({
                    "servico":        row[col_srv_oper],
                    "tipo_servico":   str(row.get(col_tipo_oper, "")).strip(),
                    "situacao":       sit,
                    "data_limite":    row[cols_oper_com["data"]].date() if pd.notna(row[cols_oper_com["data"]]) else data_inicio,
                    "ausente_em":     "Cbill",
                    "sistema_origem": "Oper",
                })

        df_resultado = pd.DataFrame(registros)

        # ── cálculo dos gauges ──
        total_universo = total_cbill + total_oper
        n_div          = len(apenas_cbill) + len(apenas_oper)
        pct_geral      = round(n_div / total_universo * 100, 1) if total_universo else 0
        pct_cbill_fora = round(len(apenas_cbill) / total_cbill * 100, 1) if total_cbill else 0
        pct_oper_fora  = round(len(apenas_oper)  / total_oper  * 100, 1) if total_oper  else 0

        # ── três gauges no topo ──
        st.markdown("---")
        g1, g2, g3 = st.columns(3)
        g1.plotly_chart(gauge("% Divergência Geral",       pct_geral,      "#6c63ff"), use_container_width=True)
        g2.plotly_chart(gauge("% Cbill ausente no Oper",   pct_cbill_fora, "#e74c3c"), use_container_width=True)
        g3.plotly_chart(gauge("% Oper ausente no Cbill",   pct_oper_fora,  "#e67e22"), use_container_width=True)

        # ── dashboard na sidebar ──
        with dash_placeholder.container():
            st.markdown(f"**📅 {label_periodo}**")
            st.markdown("---")
            st.markdown(f"🔵 **Cbill:** {total_cbill} serviços")
            st.markdown(f"🟠 **Oper:** {total_oper} serviços")
            st.markdown(f"⚠️ **Divergentes:** {n_div}")
            st.markdown("---")
            st.plotly_chart(
                grafico_barras(total_cbill, total_oper, len(apenas_cbill), len(apenas_oper)),
                use_container_width=True,
            )

        st.markdown("---")

        if df_resultado.empty:
            st.success("✅ Nenhuma divergência encontrada! As bases estão alinhadas para este período.")
        else:
            st.warning(f"⚠️ **{len(df_resultado)} serviço(s) divergente(s)** no período {label_periodo}")

            tab1, tab2, tab3 = st.tabs([
                f"📋 Todos ({len(df_resultado)})",
                f"🔵 Só no Cbill ({len(apenas_cbill)})",
                f"🟠 Só no Oper ({len(apenas_oper)})",
            ])

            with tab1:
                st.dataframe(df_resultado, use_container_width=True)

            with tab2:
                sub = df_resultado[df_resultado["sistema_origem"] == "Cbill"]
                if not sub.empty:
                    st.dataframe(sub, use_container_width=True)
                else:
                    st.info("Sem exclusivos.")

            with tab3:
                sub = df_resultado[df_resultado["sistema_origem"] == "Oper"]
                if not sub.empty:
                    st.dataframe(sub, use_container_width=True)
                else:
                    st.info("Sem exclusivos.")

            st.markdown("---")
            sufixo_arquivo = (
                data_inicio.strftime("%d.%m")
                if data_inicio == data_fim
                else f"{data_inicio.strftime('%d.%m')}_a_{data_fim.strftime('%d.%m')}"
            )
            nome_saida = f"divergentes_{sufixo_arquivo}_Cbill_vs_Oper.xlsx"
            st.download_button(
                label="⬇️ Baixar planilha de divergentes (.xlsx)",
                data=exportar_excel(df_resultado),
                file_name=nome_saida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        # ── expanders de visualização ──
        base_oper_exib = base_oper.copy()
        if not tem_situacao:
            base_oper_exib["Situação"] = "—"

        with st.expander(f"👁️ Ver base Cbill filtrada ({total_cbill} registros)"):
            st.dataframe(base_cbill, use_container_width=True)
        with st.expander(f"👁️ Ver base Oper filtrada ({total_oper} registros, sem Restabelecimento)"):
            st.dataframe(base_oper_exib, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Erro ao processar: {e}")
        st.exception(e)

else:
    st.info("👆 Faça upload da base Cbill e ao menos da base Oper Comercial para iniciar a comparação.")
