import pandas as pd
import os
from io import BytesIO
from openpyxl.styles import PatternFill
import streamlit as st

# Função para buscar preço baseado no SKU
def buscar_preco(sku2, df_base, coluna_sku_base, coluna_preco_base):
    sku2 = str(sku2).strip().lower()
    for _, row in df_base.iterrows():
        sku1 = str(row[coluna_sku_base]).strip().lower()
        if sku1 in sku2 or sku2 in sku1:
            return row[coluna_preco_base]
    return None

# Função para atualizar planilhas
def atualizar_planilha(planilha, planilha_base, coluna_sku_base, coluna_preco_base, novo_nome_coluna):
    if "Número de referência SKU" in planilha.columns:
        coluna_sku_planilha = "Número de referência SKU"  # Shopee
    elif "SKU" in planilha.columns:
        coluna_sku_planilha = "SKU"  # Mercado Livre
    else:
        raise ValueError("Nenhuma coluna de SKU encontrada (esperado 'Número de referência SKU' ou 'SKU').")

    planilha[novo_nome_coluna] = planilha[coluna_sku_planilha].apply(
        lambda sku: buscar_preco(sku, planilha_base, coluna_sku_base, coluna_preco_base)
    )
    planilha[novo_nome_coluna].fillna("SKU não encontrado", inplace=True)
    return planilha

# Interface do Streamlit
st.title("Atualizador de CMV")

arquivo_base = st.file_uploader("Faça upload da planilha CMV (Base de dados)", type=["xlsx"])
arquivos_para_atualizar = st.file_uploader("Faça upload das planilhas para atualizar", type=["xlsx"], accept_multiple_files=True)

if st.button("Gerar CMV Atualizado"):
    if not arquivo_base or not arquivos_para_atualizar:
        st.error("Por favor, envie a planilha CMV e as planilhas a serem atualizadas.")
    else:
        coluna_sku_base = "SKU"
        coluna_preco_base = "Preço"
        novo_nome_coluna = "CMV"

        planilha_base = pd.read_excel(arquivo_base)

        progresso = st.progress(0)
        total_arquivos = len(arquivos_para_atualizar)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for idx, arquivo in enumerate(arquivos_para_atualizar):
                try:
                    planilha = pd.read_excel(arquivo, header=0)

                    if "Número de referência SKU" not in planilha.columns and "SKU" not in planilha.columns:
                        planilha = pd.read_excel(arquivo, header=5)

                    planilha_atualizada = atualizar_planilha(
                        planilha, planilha_base, coluna_sku_base, coluna_preco_base, novo_nome_coluna
                    )

                    # Grava cada arquivo atualizado em uma aba do Excel
                    nome_aba = os.path.splitext(arquivo.name)[0][:30]  # limita nome da aba a 30 chars
                    planilha_atualizada.to_excel(writer, index=False, sheet_name=nome_aba)

                    # Pinta a coluna CMV de amarelo
                    workbook = writer.book
                    sheet = workbook[nome_aba]
                    col_idx = list(planilha_atualizada.columns).index(novo_nome_coluna) + 1
                    amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                    for row in range(2, sheet.max_row + 1):
                        cell = sheet.cell(row=row, column=col_idx)
                        cell.fill = amarelo

                except Exception as e:
                    st.error(f"Erro ao processar {arquivo.name}: {e}")

                progresso.progress((idx + 1) / total_arquivos)

        output.seek(0)

        st.download_button(
            label="📥 Baixar CMV Atualizado",
            data=output,
            file_name="CMV_Atualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.success("Planilhas atualizadas com sucesso! ✅")
