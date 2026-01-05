import pandas as pd
import streamlit as st
import os

# Função para processar Shopee
def processar_shopee(file_shopee):
    if file_shopee is None:
        return "Erro: Nenhuma planilha da Shopee foi carregada."

    try:
        # Lê a planilha Shopee
        df_shopee = pd.read_excel(file_shopee, header=0)

        # Remover colunas "Unnamed" e colunas em branco
        df_shopee = df_shopee.loc[:, ~df_shopee.columns.str.contains('^Unnamed', na=False)]
        df_shopee = df_shopee.dropna(axis=1, how='all')
        df_shopee.columns = df_shopee.columns.str.strip()
    except Exception as e:
        return f"Erro ao ler a planilha de Shopee: {e}"

    # -------------------- FILTRAGEM --------------------
    status_coluna = next((col for col in df_shopee.columns if 'status do pedido' in col.lower()), None)
    if status_coluna is None:
        return "Erro: A coluna 'Status do pedido' não foi encontrada."

    padroes_exclusao = ['cancelado', 'cancelados', 'cancelada', 'canceladas', 'não pago', 'não pagos']
    df_shopee = df_shopee[
        ~df_shopee[status_coluna].str.contains('|'.join(padroes_exclusao), case=False, na=False)
    ]

    # -------------------- FATURAMENTO --------------------
    if 'Subtotal do produto' not in df_shopee.columns:
        return "Erro: A coluna 'Subtotal do produto' não foi encontrada."

    df_shopee['Subtotal do produto'] = pd.to_numeric(
        df_shopee['Subtotal do produto'], errors='coerce'
    )
    faturamento_total = df_shopee['Subtotal do produto'].sum()

    # -------------------- COMISSÕES --------------------
    df_comissao = df_shopee.copy()

    if 'ID do pedido' in df_comissao.columns:
        df_comissao = df_comissao.drop_duplicates(subset=['ID do pedido'], keep='first')

    colunas_comissao = [
        'Taxa de comissão bruta',
        'Taxa de serviço bruta',
        'Cupom do vendedor',
        'Cupom Shopee'
    ]

    for coluna in colunas_comissao:
        if coluna not in df_comissao.columns:
            return f"Erro: A coluna '{coluna}' não foi encontrada."

        df_comissao[coluna] = pd.to_numeric(df_comissao[coluna], errors='coerce').fillna(0)

    comissoes_detalhadas = df_comissao[colunas_comissao].sum()
    comissao_total = comissoes_detalhadas.sum()

    # -------------------- DEVOLUÇÕES --------------------
    if 'Status da Devolução / Reembolso' not in df_shopee.columns:
        valor_devolucao = 0
    else:
        df_devolucoes = df_shopee[df_shopee['Status da Devolução / Reembolso'].notna()].copy()
        df_devolucoes['Subtotal do produto'] = pd.to_numeric(
            df_devolucoes['Subtotal do produto'], errors='coerce'
        )
        valor_devolucao = df_devolucoes['Subtotal do produto'].sum()

    # -------------------- ENTREGA DIRETA --------------------
    valor_entrega_direta = 0
    if (
        'ID do pedido' in df_shopee.columns
        and 'Opção de envio' in df_shopee.columns
        and 'Valor estimado do frete' in df_shopee.columns
    ):
        df_entrega = df_shopee.copy()

        df_entrega = df_entrega.drop_duplicates(subset=['ID do pedido'], keep='first')
        df_entrega = df_entrega[
            ~df_entrega[status_coluna].str.contains('|'.join(padroes_exclusao), case=False, na=False)
        ]

        df_entrega = df_entrega[
            df_entrega['Opção de envio'].str.contains('Shopee Entrega Direta', case=False, na=False)
        ]

        df_entrega['Valor estimado do frete'] = pd.to_numeric(
            df_entrega['Valor estimado do frete'], errors='coerce'
        ).fillna(0)

        valor_entrega_direta = df_entrega['Valor estimado do frete'].sum()

    # -------------------- QUANTIDADE DE PEDIDOS --------------------
    qtd_pedidos = 0
    if 'ID do pedido' in df_shopee.columns:
        df_pedidos = df_shopee.copy()
        df_pedidos = df_pedidos.drop_duplicates(subset=['ID do pedido'], keep='first')
        df_pedidos = df_pedidos[
            ~df_pedidos[status_coluna].str.contains('|'.join(padroes_exclusao), case=False, na=False)
        ]
        qtd_pedidos = len(df_pedidos)

    # -------------------- PLANILHA DE SAÍDA --------------------
    tabela_resumo = {
        'Faturamento Shopee': faturamento_total,
        'Taxa de comissão bruta': comissoes_detalhadas['Taxa de comissão bruta'],
        'Taxa de serviço bruta': comissoes_detalhadas['Taxa de serviço bruta'],
        'Cupom do vendedor': comissoes_detalhadas['Cupom do vendedor'],
        'Cupom Shopee': comissoes_detalhadas['Cupom Shopee'],
        'Comissão Total': comissao_total,
        'Valor Devolvido': valor_devolucao,
        'Entrega Direta (Frete)': valor_entrega_direta,
        'Quantidade de Pedidos': qtd_pedidos
    }

    df_dre = pd.DataFrame(tabela_resumo.items(), columns=['Descrição', 'Valor'])

    destaques = [
        'Faturamento Shopee',
        'Comissão Total',
        'Valor Devolvido',
        'Entrega Direta (Frete)',
        'Quantidade de Pedidos'
    ]

    def highlight_rows(s):
        return ['background-color: yellow' if v in destaques else '' for v in s]

    df_styled = df_dre.style.apply(highlight_rows, subset=['Descrição'])

    # -------------------- SALVAR EXCEL --------------------
    output_dir = 'uploads'
    os.makedirs(output_dir, exist_ok=True)

    output_filepath = os.path.join(output_dir, "DRE_shopee.xlsx")
    try:
        df_styled.to_excel(output_filepath, index=False, engine="openpyxl")
    except Exception as e:
        return f"Erro ao gerar o arquivo Excel: {e}"

    return output_filepath


# -------------------- STREAMLIT --------------------
def main():
    st.title("📊 Gerador de DRE - Shopee")
    st.write("Envie sua planilha da Shopee para gerar o relatório.")

    marketplace = st.radio(
        "🛒 Selecione o Marketplace:",
        ["Shopee", "Mercado Livre", "Amazon"],
        horizontal=True
    )

    if marketplace == "Shopee":
        file_shopee = st.file_uploader(
            "🔽 Envie a planilha Shopee:",
            type=["xls", "xlsx"]
        )

        if file_shopee is not None and st.button("📊 Gerar Relatório"):
            st.info("🔄 Processando... Aguarde.")
            output = processar_shopee(file_shopee)

            if "Erro" in output:
                st.error(output)
            else:
                st.success("✅ Relatório gerado com sucesso!")
                with open(output, "rb") as f:
                    st.download_button(
                        label="📥 Baixar Relatório Shopee (DRE)",
                        data=f,
                        file_name="DRE_shopee.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

if __name__ == "__main__":
    main()
