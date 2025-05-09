import os

def excluir_arquivo(caminho_arquivo):
    if os.path.exists(caminho_arquivo):
        try:
            os.remove(caminho_arquivo)
            print(f"Arquivo {caminho_arquivo} excluído com sucesso!")
        except Exception as e:
            print(f"Erro ao excluir o arquivo {caminho_arquivo}: {e}")
    else:
        print(f"O arquivo {caminho_arquivo} não foi encontrado na pasta especificada.")

# Exemplo de uso da função
caminho_arquivo1 = 'C:\\Users\\rasopinheiro\\Desktop\\Bot portal da Governança Lista CGTIC progresso\\Planejada.xlsx'
excluir_arquivo(caminho_arquivo1)
