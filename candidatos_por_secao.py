import streamlit as st
import pandas as pd
from io import BytesIO

# =========================
# CONFIG
# =========================
st.set_page_config(
    page_title='Comparativo de candidatos',
    layout='wide'
)

# =========================
# LOAD DATA
# =========================
@st.cache_data
def read_parquet():
    return pd.read_parquet('eleicoes_2016_2024_agrupadas.parquet')

df = read_parquet()

# =========================
# CANDIDATOS
# =========================
@st.cache_data
def read_candidates():
    return sorted(df['NM_URNA_CANDIDATO'].dropna().unique())

# =========================
# GERAR EXCEL
# =========================
def gerar_excel(df_filtrado):

    output = BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:

        for ano in sorted(df_filtrado['ANO_ELEICAO'].unique()):

            candidatos_ano = df_filtrado[
                df_filtrado['ANO_ELEICAO'] == ano
            ]['NM_URNA_CANDIDATO'].unique()

            for candidato in candidatos_ano:

                temp = df_filtrado[
                    (df_filtrado['ANO_ELEICAO'] == ano) &
                    (df_filtrado['NM_URNA_CANDIDATO'] == candidato)
                ]

                # Nome da aba
                nome_aba = f"{candidato[:20]}_{ano}"

                # Excel não aceita alguns caracteres
                nome_aba = (
                    nome_aba
                    .replace('/', '_')
                    .replace('\\', '_')
                    .replace(':', '_')
                    .replace('*', '_')
                    .replace('?', '_')
                    .replace('[', '_')
                    .replace(']', '_')
                )

                temp.to_excel(
                    writer,
                    sheet_name=nome_aba[:31],  # limite excel
                    index=False
                )

                # Ajustar largura das colunas
                workbook = writer.book
                worksheet = writer.sheets[nome_aba[:31]]

                for i, col in enumerate(temp.columns):

                    largura = max(
                        temp[col].astype(str).map(len).max(),
                        len(col)
                    ) + 2

                    worksheet.set_column(i, i, min(largura, 40))

    output.seek(0)

    return output

# =========================
# UI
# =========================
st.title('Candidatos por seção - Comparativo')

anos = st.multiselect(
    'Selecione o(s) ano(s)',
    options=sorted(df['ANO_ELEICAO'].unique())
)

lista_candidatos = read_candidates()

candidatos = st.multiselect(
    'Selecione o(s) candidato(s)',
    options=lista_candidatos
)

# =========================
# GERAR TABELA
# =========================
if st.button('Gerar tabela'):

    if not anos:
        st.warning('Selecione ao menos um ano.')
        st.stop()

    if not candidatos:
        st.warning('Selecione ao menos um candidato.')
        st.stop()

    df_filtrado = df[
        (df['ANO_ELEICAO'].isin(anos)) &
        (df['NM_URNA_CANDIDATO'].isin(candidatos))
    ]

    if df_filtrado.empty:
        st.error('Nenhum dado encontrado.')
        st.stop()

    df_filtrado = df_filtrado.sort_values(
        by=['ANO_ELEICAO', 'NM_URNA_CANDIDATO']
    )

    st.success(f'{len(df_filtrado)} registros encontrados.')

    st.dataframe(
        df_filtrado,
        use_container_width=True,
        height=600
    )

    # =========================
    # DOWNLOAD EXCEL
    # =========================
    excel_file = gerar_excel(df_filtrado)

    st.download_button(
        label='📥 Baixar Excel',
        data=excel_file,
        file_name='comparativo_candidatos.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )