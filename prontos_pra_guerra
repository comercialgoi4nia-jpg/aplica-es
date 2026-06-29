import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Conferência MISCELÂNEA", page_icon="📋", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f5f7fa; }
    .stButton>button { background-color: #1a73e8; color: white; border-radius: 8px; font-weight: bold; }
    .metric-card { background: white; padding: 16px; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
    </style>
""", unsafe_allow_html=True)

st.title("📋 Conferência de MISCELÂNEA")
st.markdown("Compare a coluna **MISCELÂNEA** entre a base de abertas e a base de executadas.")

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("📂 Base Abertas")
    st.caption("Planilha com os chamados abertos (referência)")
    file_abertas = st.file_uploader("Selecione o arquivo", type=["xlsx", "xls", "csv"], key="abertas")

with col2:
    st.subheader("✅ Base Executadas")
    st.caption("Planilha com os chamados já executados")
    file_executadas = st.file_uploader("Selecione o arquivo", type=["xlsx", "xls", "csv"], key="executadas")

COLUNAS_ABERTAS = [
    "Bloco", "N° Ticket", "Município", "Qtd Câmeras",
    "Carga Inst. (W)", "Consumo (kWh/mês)", "Latitude", "Ponto",
    "Longitude", "PONTO DE SERVIÇO", "MISCELANEA", "Endereço"
]

def load_file(f):
    if f.name.endswith(".csv"):
        return pd.read_csv(f, dtype=str)
    return pd.read_excel(f, dtype=str)

def normalize(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    return str(val).strip().upper()

def gerar_excel(df_resultado):
    wb = Workbook()
    ws = wb.active
    ws.title = "Conferência"

    # Colors
    verde_fill   = PatternFill("solid", fgColor="C6EFCE")  # executada
    vermelho_fill = PatternFill("solid", fgColor="FFC7CE")  # não executada
    amarelo_fill  = PatternFill("solid", fgColor="FFEB9C")  # abrir miscelanea
    header_fill  = PatternFill("solid", fgColor="1A73E8")

    bold_white = Font(bold=True, color="FFFFFF")
    bold_black = Font(bold=True)
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    headers = list(df_resultado.columns)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = bold_white
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    status_col = headers.index("STATUS") + 1

    for row_idx, row in enumerate(df_resultado.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = border

        status = ws.cell(row=row_idx, column=status_col).value
        if status == "EXECUTADA":
            fill = verde_fill
            ws.cell(row=row_idx, column=status_col).font = Font(bold=True, color="375623")
        elif status == "NÃO EXECUTADA":
            fill = vermelho_fill
            ws.cell(row=row_idx, column=status_col).font = Font(bold=True, color="9C0006")
        else:
            fill = PatternFill()

        # Mark MISCELANEA col if "ABRIR MISCELANEA"
        misc_col_idx = headers.index("MISCELANEA_EXECUTADAS") + 1 if "MISCELANEA_EXECUTADAS" in headers else None

        for col_idx2 in range(1, len(headers) + 1):
            ws.cell(row=row_idx, column=col_idx2).fill = fill

        if misc_col_idx:
            misc_val = ws.cell(row=row_idx, column=misc_col_idx).value
            if misc_val == "ABRIR MISCELANEA":
                ws.cell(row=row_idx, column=misc_col_idx).fill = amarelo_fill
                ws.cell(row=row_idx, column=misc_col_idx).font = Font(bold=True, color="7B5E00")

    # Auto-width
    for col_idx, col_cells in enumerate(ws.columns, 1):
        max_len = max((len(str(c.value or "")) for c in col_cells), default=10)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 40)

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    # Legend sheet
    ws2 = wb.create_sheet("Legenda")
    ws2["A1"] = "Legenda de Cores"
    ws2["A1"].font = Font(bold=True, size=13)
    legenda = [
        ("Verde", "C6EFCE", "Executada — MISCELÂNEA encontrada na base de executadas"),
        ("Vermelho", "FFC7CE", "Não Executada — MISCELÂNEA não encontrada na base de executadas"),
        ("Amarelo", "FFEB9C", "ABRIR MISCELANEA — linha da executada sem MISCELÂNEA preenchida"),
    ]
    for i, (label, color, desc) in enumerate(legenda, 3):
        ws2.cell(row=i, column=1, value=label).fill = PatternFill("solid", fgColor=color)
        ws2.cell(row=i, column=1).font = Font(bold=True)
        ws2.cell(row=i, column=2, value=desc)
        ws2.column_dimensions["B"].width = 60

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

if file_abertas and file_executadas:
    try:
        df_ab = load_file(file_abertas)
        df_ex = load_file(file_executadas)

        # Normalize column names
        df_ab.columns = df_ab.columns.str.strip()
        df_ex.columns = df_ex.columns.str.strip()

        if "MISCELANEA" not in df_ab.columns:
            st.error("❌ Coluna **MISCELANEA** não encontrada na base de Abertas. Verifique o arquivo.")
            st.stop()
        if "MISCELANEA" not in df_ex.columns:
            st.error("❌ Coluna **MISCELANEA** não encontrada na base de Executadas. Verifique o arquivo.")
            st.stop()

        # Fill empty MISCELANEA in executadas
        df_ex["MISCELANEA"] = df_ex["MISCELANEA"].apply(
            lambda v: "ABRIR MISCELANEA" if normalize(v) == "" else v
        )

        # Build set of miscelaneas from executadas (normalized)
        exec_misc_set = set(df_ex["MISCELANEA"].apply(normalize)) - {"", "ABRIR MISCELANEA"}

        # Build result from abertas
        cols_abertas = [c for c in COLUNAS_ABERTAS if c in df_ab.columns]
        df_resultado = df_ab[cols_abertas].copy()

        # Add status
        df_resultado["STATUS"] = df_resultado["MISCELANEA"].apply(
            lambda v: "EXECUTADA" if normalize(v) in exec_misc_set else "NÃO EXECUTADA"
        )

        # Add MISCELANEA from executadas side (for reference)
        # Map: misc_ab -> misc_ex entry
        exec_misc_map = {}
        for _, row in df_ex.iterrows():
            k = normalize(row["MISCELANEA"])
            if k not in ("", "ABRIR MISCELANEA"):
                exec_misc_map[k] = row["MISCELANEA"]

        df_resultado["MISCELANEA_EXECUTADAS"] = df_resultado["MISCELANEA"].apply(
            lambda v: exec_misc_map.get(normalize(v), "ABRIR MISCELANEA" if normalize(v) == "" else "NÃO ENCONTRADA")
        )

        # Reorder columns: put STATUS first
        cols_order = ["STATUS"] + [c for c in df_resultado.columns if c != "STATUS"]
        df_resultado = df_resultado[cols_order]

        st.divider()
        st.subheader("📊 Resumo")

        total = len(df_resultado)
        executadas = (df_resultado["STATUS"] == "EXECUTADA").sum()
        nao_executadas = (df_resultado["STATUS"] == "NÃO EXECUTADA").sum()
        abrir_misc = (df_resultado["MISCELANEA_EXECUTADAS"] == "ABRIR MISCELANEA").sum()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total (Abertas)", total)
        c2.metric("✅ Executadas", executadas, delta=f"{executadas/total*100:.1f}%")
        c3.metric("❌ Não Executadas", nao_executadas, delta=f"-{nao_executadas/total*100:.1f}%", delta_color="inverse")
        c4.metric("⚠️ Abrir Miscelânea", abrir_misc)

        st.divider()
        st.subheader("📋 Resultado")

        filtro = st.radio("Filtrar por:", ["Todos", "Executadas", "Não Executadas"], horizontal=True)
        if filtro == "Executadas":
            df_show = df_resultado[df_resultado["STATUS"] == "EXECUTADA"]
        elif filtro == "Não Executadas":
            df_show = df_resultado[df_resultado["STATUS"] == "NÃO EXECUTADA"]
        else:
            df_show = df_resultado

        def highlight_status(row):
            if row["STATUS"] == "EXECUTADA":
                return ["background-color: #C6EFCE"] * len(row)
            elif row["STATUS"] == "NÃO EXECUTADA":
                return ["background-color: #FFC7CE"] * len(row)
            return [""] * len(row)

        st.dataframe(
            df_show.style.apply(highlight_status, axis=1),
            use_container_width=True,
            height=420
        )

        st.divider()
        excel_bytes = gerar_excel(df_resultado)
        st.download_button(
            label="⬇️ Baixar Planilha com Resultado",
            data=excel_bytes,
            file_name="conferencia_miscelanea.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    except Exception as e:
        st.error(f"Erro ao processar os arquivos: {e}")
        st.exception(e)
else:
    st.info("📎 Carregue os dois arquivos acima para iniciar a conferência.")
    with st.expander("ℹ️ Como funciona o aplicativo"):
        st.markdown("""
        1. **Carregue a Base de Abertas** — planilha com os chamados abertos, deve conter a coluna `MISCELANEA` e as demais colunas obrigatórias.
        2. **Carregue a Base de Executadas** — planilha com os chamados já executados, também com a coluna `MISCELANEA`.
        3. O app compara os valores de `MISCELANEA` entre as duas bases.
        4. Linhas da base de executadas com `MISCELANEA` vazia receberão o valor **ABRIR MISCELANEA**.
        5. O resultado é exibido na tela e pode ser baixado como planilha Excel colorida:
           - 🟢 **Verde** = Executada
           - 🔴 **Vermelho** = Não Executada
           - 🟡 **Amarelo** = Abrir Miscelânea (sem valor na executada)
        """)
