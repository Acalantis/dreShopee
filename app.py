
import pandas as pd
import streamlit as st
import os

# Função para processar Shopee
def processar_shopee(file_shopee):
    if file_shopee is None:
        return "Erro: Nenhuma planilha da Shopee foi carregada."

    try:
        df_shopee = pd.read_excel(file_shopee, header=0)
        df_shopee = df_shopee.loc[:, ~df_shopee.columns.str.contains('^Unnamed', na=False)]
        df_shopee = df_shopee.dropna(axis=1, how='all')
        df_shopee.columns = df_shopee.columns.str.strip()
    except Exception as e:
        return f"Erro ao ler a planilha de Shopee: {e}"

    colunas_necessarias = ['Preço acordado', 'Cupom do vendedor', 'Taxa de comissão', 'Taxa de serviço']
    for coluna in colunas_necessarias:
        if coluna not in df_shopee.columns:
            return f"Erro: A coluna '{coluna}' não foi encontrada na planilha de Shopee."

    status_coluna = next((col for col in df_shopee.columns if 'status' in col.lower()), None)
    if status_coluna is None:
        return "Erro: A coluna de status do pedido não foi encontrada na planilha de Shopee."

    df_shopee = df_shopee[~df_shopee[status_coluna].str.contains('cancelado|não pago', case=False, na=False)]
    df_shopee['Preço acordado'] = pd.to_numeric(df_shopee['Preço acordado'], errors='coerce')
    faturamento_total = df_shopee['Preço acordado'].sum()

    df_shopee['Taxa de comissão'] = pd.to_numeric(df_shopee['Taxa de comissão'], errors='coerce')
    df_shopee['Taxa de serviço'] = pd.to_numeric(df_shopee['Taxa de serviço'], errors='coerce')
    df_shopee['Cupom do vendedor'] = pd.to_numeric(df_shopee['Cupom do vendedor'], errors='coerce')
    df_shopee['Cupom do vendedor'].fillna(0, inplace=True)
    comissao_total = df_shopee['Taxa de comissão'].sum() + df_shopee['Taxa de serviço'].sum() + df_shopee['Cupom do vendedor'].sum()

    status_devolucao_coluna = next((col for col in df_shopee.columns if 'status da devolução' in col.lower() or 'reembolso' in col.lower()), None)
    if status_devolucao_coluna:
        df_devolucoes = df_shopee[df_shopee[status_devolucao_coluna].notna()]
        valor_devolucao = df_devolucoes['Preço acordado'].sum()
    else:
        valor_devolucao = 0

    tabela_resumo = {
        'Faturamento Shopee': faturamento_total,
        'Comissão Shopee': comissao_total,
        'Valor Devolvido': valor_devolucao
    }

    df_dre = pd.DataFrame(tabela_resumo.items(), columns=['Descrição', 'Valor'])
    output_dir = 'uploads'
    os.makedirs(output_dir, exist_ok=True)
    output_filepath = os.path.join(output_dir, "DRE_shopee.xlsx")
    try:
        df_dre.to_excel(output_filepath, index=False)
    except Exception as e:
        return f"Erro ao gerar o arquivo Excel: {e}"

    return output_filepath

def main():
    st.title("📊 **Gerador de DRE - Faturamento e Comissão**")
    st.write("Escolha o Marketplace e faça o upload das planilhas necessárias.")
    marketplace = st.radio("🛒 **Selecione o Marketplace:**", ["Shopee", "Mercado Livre", "Amazon"], horizontal=True)

    if marketplace == "Shopee":
        file_shopee = st.file_uploader("🔽 **Envie sua planilha da Shopee**:", type=["xls", "xlsx"])
        if file_shopee is not None:
            if st.button("📊 **Gerar Relatório de Shopee**"):
                st.info("🔄 **Processando... Aguarde!**")
                output_filepath_shopee = processar_shopee(file_shopee)
                if "Erro" in output_filepath_shopee:
                    st.error(output_filepath_shopee)
                else:
                    st.success("✅ Relatório gerado com sucesso!")
                    with open(output_filepath_shopee, "rb") as f:
                        st.download_button(
                            label="📥 **Baixar Relatório de Shopee (DRE)**",
                            data=f,
                            file_name="DRE_shopee.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

if __name__ == '__main__':
    main()
