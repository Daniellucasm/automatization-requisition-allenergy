import shutil
import os

def copiar_arquivo(caminho_origem, caminho_destino, caminho_destino_2):
    try:
        # Verifica se o arquivo de origem existe
        if os.path.exists(caminho_origem):
            # Move o arquivo
            shutil.copy(caminho_origem, caminho_destino)
            shutil.copy(caminho_origem, caminho_destino_2)
            print(f"Arquivo movido para {caminho_destino}")
            print(f"Arquivo movido para {caminho_destino_2}")
        else:
            print("Arquivo de origem não encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

# Caminhos de origem e destino
caminho_origem = "/Users/daniellucas/Library/Mobile Documents/com~apple~CloudDocs/All Energy/Programa RCO/ModelosXLSX/XXXX-RCO-000.xlsx"
caminho_destino = "/Users/daniellucas/Library/Mobile Documents/com~apple~CloudDocs/All Energy/Programa RCO/Move"
caminho_destino_2 = "/Users/daniellucas/Library/Mobile Documents/com~apple~CloudDocs/All Energy/Programa RCO/Move2"

# Chama a função para mover o arquivo
copiar_arquivo(caminho_origem, caminho_destino, caminho_destino_2)
