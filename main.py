import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors # Importar para cores personalizadas

def carregar_dados_vendas(caminho_arquivo):
    """Carrega a planilha de vendas em um DataFrame do Pandas."""
    try:
        df = pd.read_excel(caminho_arquivo)
        print("Dados da planilha carregados com sucesso!")
        return df
    except FileNotFoundError:
        print(f"Erro: O arquivo '{caminho_arquivo}' não foi encontrado.")
        return None
    except Exception as e:
        print(f"Ocorreu um erro ao carregar o arquivo: {e}")
        return None

def encontrar_pagamentos_pendentes(dataframe):
    """Filtra e retorna os clientes com pagamentos pendentes."""
    pagamentos_pendentes = dataframe[dataframe['Status Pagamento'] == 'Pendente']
    if pagamentos_pendentes.empty:
        print("Nenhum pagamento pendente encontrado.")
        return None
    else:
        print(f"Pagamentos pendentes encontrados: {len(pagamentos_pendentes)} clientes.")
        return pagamentos_pendentes

def gerar_relatorio(dados_pendentes, nome_arquivo_saida):
    """Gera um novo arquivo Excel com os pagamentos pendentes."""
    if dados_pendentes is not None:
        dados_pendentes.to_excel(nome_arquivo_saida, index=False)
        print(f"Relatório de pagamentos pendentes gerado em '{nome_arquivo_saida}'.")
    else:
        print("Não foi possível gerar o relatório.")

def gerar_grafico_pagamentos(dataframe):
    """
    Gera um gráfico de pizza profissional mostrando a proporção de pagamentos.
    """
    if dataframe is None:
        print("Não foi possível gerar o gráfico, o DataFrame está vazio.")
        return

    # Usar um estilo mais profissional para o Matplotlib
    plt.style.use('seaborn-v0_8-darkgrid') # 'seaborn-v0_8-darkgrid' ou 'ggplot' ou 'fivethirtyeight'

    contagem_status = dataframe['Status Pagamento'].value_counts()
    
    if contagem_status.empty:
        print("Não há dados de status de pagamento para gerar o gráfico.")
        return

    labels = contagem_status.index
    sizes = contagem_status.values

    # Cores personalizadas e mais profissionais
    # Você pode ajustar essas cores. Ex: tons de azul, cinza e vermelho.
    cores_personalizadas = ['#66BB6A', '#FF7043'] # Verde-claro para Pago, Laranja-vermelho para Pendente

    # Criar um "explode" para destacar a fatia de "Pendente", se existir
    explode = [0] * len(labels)
    if 'Pendente' in labels:
        idx_pendente = list(labels).index('Pendente')
        explode[idx_pendente] = 0.1 # Destaca a fatia em 10%

    fig1, ax1 = plt.subplots(figsize=(10, 8)) # Aumenta o tamanho do gráfico

    # Desenha o gráfico de pizza
    wedges, texts, autotexts = ax1.pie(
        sizes, 
        explode=explode, 
        labels=labels, 
        colors=cores_personalizadas, 
        autopct='%1.1f%%', 
        shadow=True, # Adiciona uma sombra
        startangle=90,
        pctdistance=0.85 # Distância da porcentagem do centro
    )

    # Melhorar a aparência do texto da porcentagem
    plt.setp(autotexts, size=12, weight="bold", color="white")
    plt.setp(texts, size=11, color="black") # Cor dos rótulos

    ax1.axis('equal')  # Garante que o círculo seja desenhado como um círculo.

    plt.title('Distribuição Percentual dos Status de Pagamento', fontsize=16, weight='bold', pad=20)
    
    # Adiciona uma caixa de texto com o total de clientes
    total_clientes = dataframe.shape[0]
    plt.text(0, -1.2, f'Total de Clientes Analisados: {total_clientes}', 
             horizontalalignment='center', fontsize=12, bbox=dict(facecolor='white', alpha=0.5))

    # Salva o gráfico em um arquivo de imagem
    nome_arquivo_grafico = 'distribuicao_pagamentos_profissional.png'
    plt.savefig(nome_arquivo_grafico, bbox_inches='tight', dpi=300) # dpi=300 para alta resolução
    print(f"Gráfico de distribuição de pagamentos profissional gerado em '{nome_arquivo_grafico}'.")

    plt.show() # Exibe o gráfico em uma janela

if __name__ == "__main__":
    caminho_planilha = 'dados/vendas_consultoria.xlsx'
    nome_relatorio = 'relatorio_pendencias.xlsx'

    df_vendas = carregar_dados_vendas(caminho_planilha)

    if df_vendas is not None:
        df_pendentes = encontrar_pagamentos_pendentes(df_vendas)
        
        gerar_relatorio(df_pendentes, nome_relatorio)
        
        # Chama a nova função de gráfico
        gerar_grafico_pagamentos(df_vendas)