# stremlit_pesquisa
import streamlit as st

def front():
    st.set_page_config(page_title="CONVERSOR DE CARACTERES", layout="wide")
    st.write("FORMATAR SS PARA PESQUISA NO OPER")

    entrada = st.text_input(
        "Insira as solicitações de serviço (separe por espaço ou vírgula):"
    )
    st.text("Resultado: ")

    if entrada:
        valores = entrada.replace(",", " ").split()
        resultado = " OR ".join(valores)
        st.success(resultado)

front()
