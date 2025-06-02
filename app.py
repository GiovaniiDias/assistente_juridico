import pandas as pd
import os

def padronizar_e_atualizar_planilha(caminho_entrada, caminho_saida, mapeamento_colunas, nova_coluna_nome, nova_coluna_valor):
    """
    Lê uma planilha Excel, padroniza nomes de colunas, adiciona uma nova coluna
    com um valor padrão e salva a planilha modificada.

    Args:
        caminho_entrada (str): Caminho completo para o arquivo Excel de entrada.
        caminho_saida (str): Caminho completo para o arquivo Excel de saída.
        mapeamento_colunas (dict): Dicionário com 'nome_antigo': 'novo_nome' para as colunas.
        nova_coluna_nome (str): Nome da nova coluna a ser adicionada.
        nova_coluna_valor (str): Valor padrão para a nova coluna.
    """
    try:
        # Ler a planilha Excel
        df = pd.read_excel(caminho_entrada)
        print(f"Planilha '{os.path.basename(caminho_entrada)}' lida com sucesso.")

        # 1. Padronizar nomes de colunas
        df.rename(columns=mapeamento_colunas, inplace=True)
        print("Nomes de colunas padronizados.")

        # 2. Adicionar nova coluna com valor padrão
        df[nova_coluna_nome] = nova_coluna_valor
        print(f"Coluna '{nova_coluna_nome}' adicionada com valor '{nova_coluna_valor}'.")

        # 3. Salvar a planilha modificada
        df.to_excel(caminho_saida, index=False)
        print(f"Planilha modificada salva em '{os.path.basename(caminho_saida)}'.")

    except FileNotFoundError:
        print(f"Erro: O arquivo '{caminho_entrada}' não foi encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

def separar_por_data_e_salvar_planilhas(caminho_entrada, coluna_data, pasta_saida):
    """
    Lê uma planilha Excel, separa por datas únicas na coluna especificada
    e salva uma nova planilha para cada data encontrada.

    Args:
        caminho_entrada (str): Caminho do arquivo Excel de entrada.
        coluna_data (str): Nome da coluna de data para separar.
        pasta_saida (str): Pasta onde as novas planilhas serão salvas.
    """
    try:
        df = pd.read_excel(caminho_entrada)
        if not os.path.exists(pasta_saida):
            os.makedirs(pasta_saida)
        datas_unicas = df[coluna_data].dropna().unique()
        for data in datas_unicas:
            df_data = df[df[coluna_data] == data]
            nome_arquivo = f"{coluna_data}_{str(data).replace('/', '-')}.xlsx"
            caminho_arquivo = os.path.join(pasta_saida, nome_arquivo)
            df_data.to_excel(caminho_arquivo, index=False)
            print(f"Planilha criada para data {data}: {nome_arquivo}")
    except Exception as e:
        print(f"Ocorreu um erro ao separar por data: {e}")

# --- Exemplo de Uso ---
if __name__ == "__main__":
    # Crie um arquivo Excel de exemplo para teste (ex: 'processos_forum.xlsx')
    # com as seguintes colunas (ou nomes similares):
    # Num. Proc. | Cliente | Descr. Processo | Data Entr.
    # Ex:
    # Num. Proc. | Cliente    | Descr. Processo     | Data Entr.
    # 12345      | Empresa X  | Ação Trabalhista    | 2024-01-15
    # 67890      | João Silva | Revisão de Contrato | 2024-02-20

    # Defina o mapeamento de colunas antigas para novas
    mapeamento = {
        'Num. Proc.': 'Numero_Processo',
        'Cliente': 'Nome_Cliente',
        'Descr. Processo': 'Descricao_Processo',
        'Data Entr.': 'Data_Entrada'
    }

    # Defina a nova coluna a ser adicionada
    nova_coluna = 'Status'
    valor_nova_coluna = 'A Analisar'

    # Caminhos dos arquivos
    arquivo_entrada = 'processos_forum.xlsx'
    arquivo_saida = 'processos_padronizados.xlsx'

    # # Crie um arquivo Excel de exemplo para simular a entrada
    # df_exemplo = pd.DataFrame({
    #     'Num. Proc.': ['12345', '67890', '11223'],
    #     'Cliente': ['Empresa X', 'João Silva', 'Maria Santos'],
    #     'Descr. Processo': ['Ação Trabalhista', 'Revisão de Contrato', 'Divórcio'],
    #     'Data Entr.': ['2024-01-15', '2024-02-20', '2024-03-01']
    # })
    # df_exemplo.to_excel(arquivo_entrada, index=False)
    # print(f"Arquivo de entrada '{arquivo_entrada}' criado para demonstração.\n")


    # Executar a função principal
    padronizar_e_atualizar_planilha(arquivo_entrada, arquivo_saida, mapeamento, nova_coluna, valor_nova_coluna)

    # Separar por data e criar planilhas individuais
    separar_por_data_e_salvar_planilhas(
        arquivo_saida,  # usa o arquivo já padronizado
        'Data_Entrada', # nome da coluna já padronizada
        'planilhas_por_data' # pasta de saída
    )

    # Opcional: Remover arquivos de teste após execução
    # os.remove(arquivo_entrada)
    # os.remove(arquivo_saida)