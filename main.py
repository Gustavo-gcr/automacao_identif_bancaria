# import pandas as pd
# import streamlit as st
# from io import BytesIO

# # Título da aplicação
# st.title("Processador de Planilha Excel")

# # Upload do arquivo Excel
# uploaded_file = st.file_uploader("Envie a planilha Excel", type=["xlsx", "xls"])

# if uploaded_file is not None:
#     # Ler o arquivo enviado
#     df = pd.read_excel(uploaded_file)

#     # Verificar se as colunas esperadas estão presentes
#     colunas_esperadas = ["Empresa", "Conta", "Descricao", "emp", "cont", "descr"]
#     if all(coluna in df.columns for coluna in colunas_esperadas):
#         # Inicializar as colunas de Result
#         df["Result_Emp"] = ""
#         df["Result_Conta"] = ""
#         df["Result_Descricao"] = ""

#         # Verificar as contas que estão na base, mas não nas novas colunas
#         contas_novas = df["cont"].tolist()

#         for index, row in df.iterrows():
#             if row["Conta"] not in contas_novas:
#                 # Preencher Result_Emp, Result_Conta e Result_Descricao
#                 df.at[index, "Result_Emp"] = row["Empresa"]
#                 df.at[index, "Result_Conta"] = row["Conta"]
#                 df.at[index, "Result_Descricao"] = row["Descricao"]

#         # Reorganizar as colunas para que os resultados fiquem com três colunas de distância
#         colunas_reordenadas = (
#             df.columns[:3].tolist()  # Primeiras três colunas
#             + df.columns[3:-3].tolist()  # Colunas intermediárias
#             + ["Result_Emp", "Result_Conta", "Result_Descricao"]  # Colunas de resultado ao final
#         )
#         df = df[colunas_reordenadas]

#         # Exibir o DataFrame processado no Streamlit
#         st.write("Planilha processada com sucesso!")
#         #st.dataframe(df)

#         # Permitir download do arquivo processado
#         output = BytesIO()
#         df.to_excel(output, index=False, engine="openpyxl")
#         output.seek(0)

#         st.download_button(
#             label="Baixar planilha processada",
#             data=output,
#             file_name="planilha_resultado.xlsx",
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         )
#     else:
#         st.error(f"A planilha enviada não contém as colunas necessárias: {', '.join(colunas_esperadas)}")
import pandas as pd
import streamlit as st
from io import BytesIO

# Título da aplicação
st.title("Processador de Planilha Excel")

# Upload do arquivo Excel
uploaded_file = st.file_uploader("Envie a planilha Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Ler o arquivo enviado
    df = pd.read_excel(uploaded_file)

    # Verificar se as colunas esperadas estão presentes
    colunas_esperadas = ["Empresa", "Conta", "Descricao", "emp", "cont", "descr"]
    if all(coluna in df.columns for coluna in colunas_esperadas):

        # Inicializar DataFrame para resultados
        resultados = pd.DataFrame(columns=["Result_Emp", "Result_Conta", "Result_Descricao"])

        # Criar DataFrame apenas com as colunas novas
        novas_contas = df[["emp", "cont", "descr"]].drop_duplicates()

        # Verificar as contas que estão na base, mas não nas novas colunas
        for index, row in df.iterrows():
            if not ((novas_contas["cont"] == row["Conta"]) & (novas_contas["descr"] == row["Descricao"])).any():
                resultados = pd.concat([resultados, pd.DataFrame({
                    "Result_Emp": [row["Empresa"]],
                    "Result_Conta": [row["Conta"]],
                    "Result_Descricao": [row["Descricao"]]
                })], ignore_index=True)

        # Adicionar resultados ao DataFrame principal
        df = pd.concat([df, resultados], axis=1)

        # Exibir o DataFrame processado no Streamlit
        st.write("Planilha processada com sucesso!")
        st.dataframe(df)

        # Permitir download do arquivo processado
        output = BytesIO()
        df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        st.download_button(
            label="Baixar planilha processada",
            data=output,
            file_name="planilha_resultado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error(f"A planilha enviada não contém as colunas necessárias: {', '.join(colunas_esperadas)}")
