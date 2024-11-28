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
    colunas_esperadas = ["Empresa", "Conta", "Descricao", "emp", "cont", "desc"]
    if all(coluna in df.columns for coluna in colunas_esperadas):
        # Inicializar a coluna Result
        df["Result"] = ""

        # Verificar as contas que estão na base, mas não nas novas colunas
        contas_base = df["Conta"].tolist()
        contas_novas = df["cont"].tolist()

        for index, row in df.iterrows():
            # Se a conta da linha não estiver nas novas colunas, preenchê-la no Result
            if row["Conta"] not in contas_novas:
                df.at[index, "Result"] = f"{row['Empresa']}- {row['Conta']} - {row['Descricao']}"

        # Reorganizar as colunas para que "Result" fique com três colunas de distância
        colunas_reordenadas = (
            df.columns[:3].tolist()       # Primeiras três colunas
            + df.columns[3:-1].tolist()  # Colunas intermediárias
            + ["Result"]  # Coluna "Result" ao final
        )
        df = df[colunas_reordenadas]

        # Exibir o DataFrame processado no Streamlit
        st.write("Planilha processada com sucesso!")
       # st.dataframe(df)

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
