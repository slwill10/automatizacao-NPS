import streamlit as st
import pandas as pd
from nps_contratante import gerar_aba_contratante
from nps_gestor import gerar_aba_gestor
from nps_aluno import gerar_aba_aluno
from gerar_tabela import gerar_tabela_resumo_nps
import tempfile
import os

st.set_page_config(page_title="Calculadora de NPS", layout="centered")

st.title("📊 Gerador de Relatório NPS")

st.markdown("Envie os arquivos necessários para gerar o relatório.")

file_gestor = st.file_uploader("📎 Envie o arquivo de **Gestor** ", type=["xlsx", "csv"], key="gestor")
file_contratante = st.file_uploader("📎 Envie o arquivo de **Contratante** ", type=["xlsx", "csv"], key="contratante")
file_aluno = st.file_uploader("📎 Envie o arquivo de **Aluno**", type=["xlsx", "csv"], key="aluno")

if file_gestor and file_contratante:
    st.success("Arquivos carregados com sucesso!")

    if st.button("📥 Gerar relatório NPS"):
        with tempfile.TemporaryDirectory() as tempdir:
            ext_gestor = os.path.splitext(file_gestor.name)[-1].lower()
            path_gestor = os.path.join(tempdir, f"planilha_de_gestor{ext_gestor}")

            ext_contratante = os.path.splitext(file_contratante.name)[-1].lower()
            path_contratante = os.path.join(tempdir, f"planilha_de_contratante{ext_contratante}")

            ext_aluno = os.path.splitext(file_aluno.name)[-1].lower()
            path_aluno = os.path.join(tempdir, f"planilha_de_aluno{ext_aluno}")

            with open(path_gestor, "wb") as f:
                f.write(file_gestor.read())

            with open(path_contratante, "wb") as f:
                f.write(file_contratante.read())

            with open(path_aluno, "wb") as f:
                f.write(file_aluno.read())

            output_path = os.path.join(tempdir, "relatorio_NPS.xlsx")
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                gerar_aba_gestor(writer, path_gestor)
                gerar_aba_aluno(writer, path_aluno)
                gerar_aba_contratante(writer, path_contratante)
                gerar_tabela_resumo_nps(writer)

            with open(output_path, "rb") as f:
                st.download_button(
                    label="📤 Baixar Relatório",
                    data=f,
                    file_name="relatorio_NPS.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )