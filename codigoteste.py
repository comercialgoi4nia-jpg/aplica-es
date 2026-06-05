import streamlit as st
import pandas as pd
import io
import os
import time
import tempfile
import subprocess
import requests
import base64
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── Credenciais WhatsApp Cloud API ───────────────────────────────────────────
# Configure esses valores nos Secrets do Streamlit Cloud:
# Settings → Secrets → adicione as linhas abaixo
#
# WHATSAPP_TOKEN = "seu_token_aqui"
# PHONE_NUMBER_ID = "seu_phone_number_id_aqui"
# RECIPIENT_NUMBER = "5562999999999"  # seu número com DDI+DDD

WA_TOKEN      = st.secrets.get("WHATSAPP_TOKEN", "")
WA_PHONE_ID   = st.secrets.get("PHONE_NUMBER_ID", "")
WA_RECIPIENT  = st.secrets.get("RECIPIENT_NUMBER", "")


# ── Gerar Excel ───────────────────────────────────────────────────────────────

def gerar_excel(tabela: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Acomp. P1"

    AZUL_HEADER = "2E5F9E"
    AZUL_TITULO = "1F4E79"
    BRANCO      = "FFFFFF"
    CINZA_LINHA = "DEEAF1"

    thin = Side(style="thin", color="AAAAAA")
    borda = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:F1")
    c = ws["A1"]
    c.value = "Acompanhamento de Suspensão/Vistoria P1"
    c.font = Font(name="Arial", bold=True, color=BRANCO, size=13)
    c.fill = PatternFill("solid", fgColor=AZUL_TITULO)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    headers = ["DATA", "P1 SUSPENSÃO - GRUPO A", "P1 SUSPENSÃO - POSTE",
               "P1 VISTORIA - RETIRADA DE RAMAL", "Total Geral", "Total Dívida"]
    for col_idx, h in enumerate(headers, start=1):
        c = ws.cell(row=2, column=col_idx, value=h)
        c.font = Font(name="Arial", bold=True, color=BRANCO, size=10)
        c.fill = PatternFill("solid", fgColor=AZUL_HEADER)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = borda
    ws.row_dimensions[2].height = 30

    meses = {"Jan":"jan","Feb":"fev","Mar":"mar","Apr":"abr","May":"mai","Jun":"jun",
             "Jul":"jul","Aug":"ago","Sep":"set","Oct":"out","Nov":"nov","Dec":"dez"}

    for i, row in enumerate(tabela.itertuples(index=False), start=3):
        fill_cor = PatternFill("solid", fgColor=CINZA_LINHA if i % 2 == 0 else BRANCO)
        data_str = row[0].strftime("%d/%b")
        for en, pt in meses.items():
            data_str = data_str.replace(en, pt)

        c = ws.cell(row=i, column=1, value=data_str)
        c.font = Font(name="Arial", size=10)
        c.alignment = Alignment(horizontal="center")
        c.fill = fill_cor
        c.border = borda

        for col_idx, val in enumerate([row[1], row[2], row[3], row[4]], start=2):
            c = ws.cell(row=i, column=col_idx, value=int(val))
            c.font = Font(name="Arial", size=10, bold=(col_idx == 5))
            c.alignment = Alignment(horizontal="center")
            c.fill = fill_cor
            c.border = borda

        c = ws.cell(row=i, column=6, value=row[5])
        c.number_format = 'R$ #,##0.00'
        c.font = Font(name="Arial", size=10)
        c.alignment = Alignment(horizontal="right")
        c.fill = fill_cor
        c.border = borda

    total_row = ws.max_row + 1
    data_start, data_end = 3, total_row - 1
    ws.cell(row=total_row, column=1, value="Total Geral").font = Font(name="Arial", bold=True, color=BRANCO, size=10)
    ws.cell(row=total_row, column=1).fill = PatternFill("solid", fgColor=AZUL_HEADER)
    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=total_row, column=1).border = borda

    for col_idx, col_letter in enumerate(["B","C","D","E","F"], start=2):
        c = ws.cell(row=total_row, column=col_idx)
        c.value = f"=SUM({col_letter}{data_start}:{col_letter}{data_end})"
        c.font = Font(name="Arial", bold=True, color=BRANCO, size=10)
        c.fill = PatternFill("solid", fgColor=AZUL_HEADER)
        c.alignment = Alignment(horizontal="center" if col_idx < 6 else "right")
        c.border = borda
        if col_idx == 6:
            c.number_format = 'R$ #,##0.00'

    for i, w in enumerate([12, 26, 22, 34, 14, 18], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ── WhatsApp Cloud API ────────────────────────────────────────────────────────

def upload_midia_whatsapp(excel_bytes: bytes) -> str | None:
    """Faz upload do Excel para a API da Meta e retorna o media_id."""
    url = f"https://graph.facebook.com/v19.0/{WA_PHONE_ID}/media"
    headers = {"Authorization": f"Bearer {WA_TOKEN}"}
    files = {
        "file": ("relatorio_p1.xlsx", excel_bytes,
                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        "type": (None, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        "messaging_product": (None, "whatsapp"),
    }
    resp = requests.post(url, headers=headers, files=files)
    if resp.status_code == 200:
        return resp.json().get("id")
    else:
        st.error(f"Erro no upload da mídia: {resp.status_code} — {resp.text}")
        return None


def enviar_documento_whatsapp(media_id: str, caption: str = "") -> bool:
    """Envia o documento já carregado para o número destinatário."""
    url = f"https://graph.facebook.com/v19.0/{WA_PHONE_ID}/messages"
    headers = {
        "Authorization": f"Bearer {WA_TOKEN}",
        "Content-Type": "application/json",
    }
    payload = {
        "messaging_product": "whatsapp",
        "to": WA_RECIPIENT,
        "type": "document",
        "document": {
            "id": media_id,
            "filename": "relatorio_suspensao_p1.xlsx",
            "caption": caption,
        },
    }
    resp = requests.post(url, headers=headers, json=payload)
    if resp.status_code == 200:
        return True
    else:
        st.error(f"Erro ao enviar mensagem: {resp.status_code} — {resp.text}")
        return False


def enviar_whatsapp(excel_bytes: bytes) -> bool:
    if not WA_TOKEN or not WA_PHONE_ID or not WA_RECIPIENT:
        st.error("⚠️ Credenciais do WhatsApp não configuradas nos Secrets do Streamlit.")
        return False

    with st.spinner("📤 Fazendo upload do arquivo..."):
        media_id = upload_midia_whatsapp(excel_bytes)

    if not media_id:
        return False

    caption = "📋 *Relatório Acompanhamento Suspensão/Vistoria P1*\nGerado automaticamente."

    with st.spinner("💬 Enviando pelo WhatsApp..."):
        return enviar_documento_whatsapp(media_id, caption)


# ── UI ────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Acompanhamento Suspensão/Vistoria P1", layout="wide")
st.title("📋 Acompanhamento de Suspensão/Vistoria P1")

uploaded_file = st.file_uploader("Suba a base (.xlsx)", type=["xlsx"])

SUBTIPOS_P1 = [
    "P1 SUSPENSÃO - GRUPO A",
    "P1 SUSPENSÃO - POSTE",
    "P1 VISTORIA - RETIRADA DE RAMAL",
]

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Sheet1")

        required_cols = {"Subtipo", "Data Inclusão", "Numero", "Valor Faturas"}
        if not required_cols.issubset(df.columns):
            st.error(f"Colunas esperadas não encontradas. Necessário: {required_cols}")
            st.stop()

        df_p1 = df[df["Subtipo"].isin(SUBTIPOS_P1)].copy()
        df_p1["Data Inclusão"] = pd.to_datetime(df_p1["Data Inclusão"]).dt.date

        pivot_qtd = (
            df_p1.groupby(["Data Inclusão", "Subtipo"])["Numero"]
            .count()
            .unstack(fill_value=0)
            .reindex(columns=SUBTIPOS_P1, fill_value=0)
        )
        pivot_qtd["Total Geral"] = pivot_qtd.sum(axis=1)

        divida_dia = (
            df_p1.groupby("Data Inclusão")["Valor Faturas"]
            .sum()
            .rename("Total Dívida")
        )

        tabela = pivot_qtd.join(divida_dia).reset_index()
        tabela = tabela.rename(columns={"Data Inclusão": "DATA"})
        tabela = tabela.sort_values("DATA")

        totais = tabela.drop(columns="DATA").sum()
        totais_row = pd.DataFrame([["Total Geral"] + totais.tolist()], columns=tabela.columns)
        tabela_exibir = pd.concat([tabela, totais_row], ignore_index=True)

        tabela_exibir["DATA"] = tabela_exibir["DATA"].apply(
            lambda x: x.strftime("%d/%b").replace("/0", "/").replace(
                "Jan","jan").replace("Feb","fev").replace("Mar","mar").replace(
                "Apr","abr").replace("May","mai").replace("Jun","jun").replace(
                "Jul","jul").replace("Aug","ago").replace("Sep","set").replace(
                "Oct","out").replace("Nov","nov").replace("Dec","dez")
            if hasattr(x, "strftime") else str(x)
        )

        tabela_exibir["Total Dívida"] = tabela_exibir["Total Dívida"].apply(
            lambda x: f"R$ {float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            if x != "Total Dívida" else x
        )

        # Métricas
        st.subheader("Resumo do Período")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("P1 Suspensão - Grupo A", int(pivot_qtd["P1 SUSPENSÃO - GRUPO A"].sum()))
        col2.metric("P1 Suspensão - Poste", int(pivot_qtd["P1 SUSPENSÃO - POSTE"].sum()))
        col3.metric("P1 Vistoria - Ret. Ramal", int(pivot_qtd["P1 VISTORIA - RETIRADA DE RAMAL"].sum()))
        total_divida = divida_dia.sum()
        total_fmt = f"R$ {total_divida:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        col4.metric("Total Dívida", total_fmt)

        st.divider()

        # Cabeçalho + botões
        col_titulo, col_excel, col_whats = st.columns([3, 1, 1])
        col_titulo.subheader("Detalhamento por Dia")

        excel_bytes = gerar_excel(tabela)

        col_excel.download_button(
            label="⬇️ Baixar Excel",
            data=excel_bytes,
            file_name="acompanhamento_suspensao_p1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        with col_whats:
            if st.button("💬 Enviar pelo WhatsApp", use_container_width=True):
                st.session_state["mostrar_envio_zap"] = True

        # Painel de envio
        if st.session_state.get("mostrar_envio_zap"):
            with st.expander("📤 Enviar relatório pelo WhatsApp", expanded=True):
                st.info(f"📱 O relatório será enviado para o número configurado nos Secrets.")

                col_a, col_b = st.columns([1, 1])

                with col_a:
                    if st.button("✅ Confirmar envio", use_container_width=True):
                        sucesso = enviar_whatsapp(excel_bytes)
                        if sucesso:
                            st.success("✅ Relatório enviado com sucesso!")
                            st.session_state["mostrar_envio_zap"] = False

                with col_b:
                    if st.button("❌ Cancelar", use_container_width=True):
                        st.session_state["mostrar_envio_zap"] = False
                        st.rerun()

        st.dataframe(
            tabela_exibir,
            use_container_width=True,
            hide_index=True,
            column_config={
                "DATA": st.column_config.TextColumn("DATA"),
                "P1 SUSPENSÃO - GRUPO A": st.column_config.NumberColumn("P1 SUSPENSÃO - GRUPO A"),
                "P1 SUSPENSÃO - POSTE": st.column_config.NumberColumn("P1 SUSPENSÃO - POSTE"),
                "P1 VISTORIA - RETIRADA DE RAMAL": st.column_config.NumberColumn("P1 VISTORIA - RETIRADA DE RAMAL"),
                "Total Geral": st.column_config.NumberColumn("Total Geral"),
                "Total Dívida": st.column_config.TextColumn("Total Dívida"),
            },
        )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
else:
    st.info("👆 Suba um arquivo .xlsx para começar.")
