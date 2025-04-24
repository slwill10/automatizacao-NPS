from nps_contratante import gerar_aba_contratante
from nps_gestor import gerar_aba_gestor
from nps_aluno import gerar_aba_aluno
from gerar_tabela import gerar_tabela_resumo_nps
import pandas as pd

with pd.ExcelWriter("relatorio_NPS.xlsx", engine='xlsxwriter') as writer:
    gerar_aba_gestor(writer)
    gerar_aba_contratante(writer)
    gerar_aba_aluno(writer)
    gerar_tabela_resumo_nps(writer)


# import streamlit as st
# import pandas as pd
# from nps_contratante import gerar_aba_contratante
# from nps_gestor import gerar_aba_gestor
# from nps_aluno import gerar_aba_aluno
# from gerar_tabela import gerar_tabela_resumo_nps
# import tempfile
# import os

# st.set_page_config(page_title="Calculadora de NPS", layout="centered")

# st.title("📊 Gerador de Relatório NPS")

# st.markdown("Envie os arquivos necessários para gerar o relatório.")

# file_gestor = st.file_uploader("📎 Envie o arquivo de **Gestor** (.xlsx)", type=["xlsx"], key="gestor")
# file_contratante = st.file_uploader("📎 Envie o arquivo de **Contratante** (.xlsx)", type=["xlsx"], key="contratante")
# file_aluno = st.file_uploader("📎 Envie o arquivo de **Aluno** (.xlsx)", type=["xlsx"], key="aluno")

# if file_gestor and file_contratante:
#     st.success("Arquivos carregados com sucesso!")

#     if st.button("📥 Gerar relatório NPS"):
#         with tempfile.TemporaryDirectory() as tempdir:
#             path_gestor = os.path.join(tempdir, "planilha_de_gestor.xlsx")
#             path_contratante = os.path.join(tempdir, "Onde_ela_obtem_o_contratente_e_aluno.xlsx")
#             path_aluno = os.path.join(tempdir, "Onde_ela_obtem_o_contratente_e_aluno.xlsx")

#             with open(path_gestor, "wb") as f:
#                 f.write(file_gestor.read())

#             with open(path_contratante, "wb") as f:
#                 f.write(file_contratante.read())

#             output_path = os.path.join(tempdir, "relatorio_NPS.xlsx")
#             with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
#                 gerar_aba_gestor(writer)
#                 gerar_aba_aluno(writer)
#                 gerar_aba_contratante(writer)
#                 gerar_tabela_resumo_nps(writer)

#             with open(output_path, "rb") as f:
#                 st.download_button(
#                     label="📤 Baixar Relatório",
#                     data=f,
#                     file_name="relatorio_NPS.xlsx",
#                     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#                 )
