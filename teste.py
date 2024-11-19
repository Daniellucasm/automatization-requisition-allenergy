import subprocess
import platform

# Caminho do arquivo Excel
file_path = '1804-RCO-008 - Copiar.xlsx'

# Detectar o sistema operacional
system = platform.system()

try:
    if system == 'Windows':
        # No Windows, use 'start'
        subprocess.run(['start', file_path], shell=True)
    elif system == 'Darwin':  # macOS
        # No macOS, use 'open'
        subprocess.run(['open', file_path])
    else:  # Linux
        # No Linux, use 'xdg-open'
        subprocess.run(['xdg-open', file_path])
    print("Arquivo Excel aberto com sucesso!")
except Exception as e:
    print(f"Erro ao tentar abrir o arquivo: {e}")
