import platform
import shutil
import os
import re
import subprocess
from tkinter import messagebox, simpledialog
from Utilities.ExcelHandler import ExcelHandler
from difflib import get_close_matches

class FileManager():
    def encontrar_arquivo_mais_proximo(self, diretorio, nome_arquivo_base):
        # Lista todos os arquivos no diretório
        arquivos = os.listdir(diretorio)

        # Obtém os nomes mais próximos usando `difflib.get_close_matches`
        mais_proximos = get_close_matches(nome_arquivo_base, arquivos, n=1)

        # Retorna o mais próximo ou None se nenhum for encontrado
        return mais_proximos[0] if mais_proximos else None

    def requisicoes_existentes(self, diretorio):
        if self.tipo_var.get() == "RCO":
            prefixo = self.projeto.get()[:4] + "-" + self.tipo_var.get()  # Número do código de projeto + RCO EX: 1804-RCO
            nova_pasta = FileManager.ler_requisicoes_existentes(self, diretorio, prefixo)
        else:
            prefixo = self.projeto.get()[:4] + "-" + self.tipo_var.get()  # Número do código de projeto + RSE EX: 1804-RSE
            nova_pasta = FileManager.ler_requisicoes_existentes(self, diretorio, prefixo)
        return nova_pasta

    def abrir_arquivo(self, file_path):
        # Detectar o sistema operacional
        system = platform.system()

        try:
            # Inicializa o comando correto para abrir o arquivo
            if system == 'Windows':
                if self.acao_var.get() == "nova":
                    partes = self.caminho_arquivo.get().split(".")
                    novo_nome = partes[0] + "-Cópia." + partes[1]
                    shutil.copy(self.caminho_arquivo.get(), novo_nome)
                    
                    processo = subprocess.Popen(['cmd', '/c', 'start', '', novo_nome], shell=True)

                    self.caminho_arquivo.set(novo_nome)
                else:
                    cmd = f'start "" "{file_path}"'
                    print(f"Comando gerado: {cmd}")
                    processo = subprocess.Popen(cmd, shell=True)

                    #print(f"Comando gerado: cmd /c start \"\" \"{file_path}\"")
                    #processo = subprocess.Popen(['cmd', '/c', 'start', "", f'"{file_path}"'], shell=True)
            elif system == 'Darwin':  # macOS
                # No macOS, use 'open' e espere pelo término
                if self.acao_var.get() == "nova":
                    partes = self.caminho_arquivo.get().split(".")
                    novo_nome = partes[0] + " - Cópia." + partes[1]
                    shutil.copy(self.caminho_arquivo.get(), novo_nome)
                    processo = subprocess.Popen(['open', novo_nome])
                    self.caminho_arquivo.set(novo_nome)
                else:
                    processo = subprocess.Popen(['open', file_path])      
            else:  # Linux
                # No Linux, use 'xdg-open' e espere pelo término
                processo = subprocess.Popen(['xdg-open', file_path])

            # Aguarda o fechamento do arquivo
            processo.wait()
        except Exception as e:
            print(f"Erro ao tentar abrir o arquivo: {e}")

    def ler_requisicoes_existentes(self, diretorio, prefixo):
        try:
            print(prefixo)
            # Lista todos os diretórios no caminho especificado
            pastas = [nome for nome in os.listdir(diretorio) if os.path.isdir(os.path.join(diretorio, nome))]

            # Regex para extrair o número sequencial da pasta no formato prefixo-###
            padrao = re.compile(rf"{prefixo}-(\d{{3}})")

            # Lista para armazenar os números já utilizados
            numeros_utilizados = []

            for pasta in pastas:
                correspondencia = padrao.match(pasta)
                if correspondencia:
                    numero = int(correspondencia.group(1))
                    numeros_utilizados.append(numero)

            # Determina o próximo número da sequência
            if numeros_utilizados:
                proximo_numero = max(numeros_utilizados) + 1
            else:
                proximo_numero = 0

            # Abre uma janela para o usuário inserir um nome
            nome_personalizado = self.obter_nome_personalizado()
            if not nome_personalizado:
                return None  # Se o usuário cancelar, retorna None

            # Formata o próximo número com três dígitos e o nome personalizado
            proxima_pasta = f"{prefixo}-{proximo_numero:03d} - {nome_personalizado}"
            print(f"Próxima pasta sugerida: {proxima_pasta}")
            return proxima_pasta

        except Exception as e:
            print(f"Ocorreu um erro: {e}")

    def obter_proxima_revisao(self, caminho_pasta, prefixo):
        # Expressão regular para capturar o prefixo e a revisão
        padrao = re.compile(rf"^({prefixo})(?:.*-R(\d+))?")
        
        # Variável para armazenar a maior revisão encontrada
        maior_revisao = -1
        
        # Percorre todos os arquivos na pasta
        for arquivo in os.listdir(caminho_pasta):
            match = padrao.match(arquivo)
            if match:
                # Captura o número da revisão (se existir)
                revisao_atual = match.group(2)
                
                if revisao_atual is not None:
                    maior_revisao = max(maior_revisao, int(revisao_atual))
        
        # Incrementa a maior revisão encontrada
        nova_revisao = maior_revisao + 1
        
        # Retorna o novo nome do arquivo com a próxima revisão
        novo_nome_arquivo = f"{prefixo}-R{nova_revisao}"
        return novo_nome_arquivo

    def encontrar_pasta_por_prefixo(self, diretorio_base, prefixo):
        try:
            pastas = [pasta for pasta in os.listdir(diretorio_base) if os.path.isdir(os.path.join(diretorio_base, pasta))]

            for pasta in pastas:
                if pasta.startswith(prefixo):
                    return os.path.join(diretorio_base, pasta)
                
            messagebox.showinfo("Pasta não encontrada", f"Nenhuma pasta encontrada com o prefixo: {prefixo}")
        except Exception as e:
            print(f"Ocorreu um erro: {e}")
            return None

    def criar_pasta_superado(self, caminho_engenharia, caminho_suprimentos):
        print("Pasta Superado")
        if self.acao_var.get() == "nova":
            os.mkdir(caminho_engenharia + "/Superado")
            os.mkdir(caminho_suprimentos + "/Superado")
        else:
            print("Jogar arquivo antigo para a pasta superado Suprimentos/Engenharia")

    def copiar_arquivo(self):
        try:
            # Caminho base da pasta "Documents" do usuário
            path = "/Users/daniellucas/Library/Mobile Documents/com~apple~CloudDocs/All Energy/"
            #/Users/daniellucas/Library/Mobile Documents/com~apple~CloudDocs/All Energy/
            #C:\\Users\\daniel.murta\\All Energy\\Apropriação de Horas - Documentos\\Testes
            caminho_base = os.path.expanduser(path)

            # Caminho de destino específico para o projeto
            caminho_destino = os.path.join(caminho_base, self.projeto.get())
            caminho_destino_engenharia = os.path.join(caminho_destino, "Requisicao/Engenharia")
            caminho_destino_suprimentos = os.path.join(caminho_destino, "Requisicao/Suprimentos")

            if self.acao_var.get() == "nova":
                # Gera o nome da nova pasta
                nome_nova_pasta = FileManager.requisicoes_existentes(self, caminho_destino_engenharia) 
                if not nome_nova_pasta:  # Verifica se o usuário cancelou a operação
                    return
                
                caminho_nova_pasta_engenharia = os.path.join(caminho_destino_engenharia, nome_nova_pasta)
                caminho_nova_pasta_suprimentos = os.path.join(caminho_destino_suprimentos, nome_nova_pasta)
                os.makedirs(caminho_nova_pasta_engenharia, exist_ok=True)
                os.makedirs(caminho_nova_pasta_suprimentos, exist_ok=True)
                # Define o novo nome do arquivo com base no nome da nova pasta
                nome_arquivo_novo = nome_nova_pasta[:12] + "-R0" + os.path.splitext(self.caminho_arquivo.get())[1]
                self.caminho_arquivo_destino_engenharia = os.path.join(caminho_nova_pasta_engenharia, nome_arquivo_novo)
                self.caminho_arquivo_destino_suprimentos = os.path.join(caminho_nova_pasta_suprimentos, nome_arquivo_novo)

                ExcelHandler.abrir_alterar_arquivo(self, nome_nova_pasta[:12])
                
            else:
                # Obtém o nome do arquivo original
                nome_arquivo_original = os.path.basename(self.caminho_arquivo.get())
                
                caminho_pasta_existente_engenharia = FileManager.encontrar_pasta_por_prefixo(self, caminho_destino_engenharia, self.requisicao.get()[:12])
                caminho_pasta_existente_suprimentos = FileManager.encontrar_pasta_por_prefixo(self, caminho_destino_suprimentos, self.requisicao.get()[:12])
                print("Engenharia: ", caminho_pasta_existente_engenharia)
                print("Suprimentos: ", caminho_pasta_existente_suprimentos)
                #Define o caminho da pasta existente onde o arquivo será copiado
                #caminho_pasta_existente_engenharia = os.path.join(caminho_destino_engenharia, nome_arquivo_original.split(" ")[0])
                #caminho_pasta_existente_suprimentos = os.path.join(caminho_destino_suprimentos, nome_arquivo_original.split(" ")[0])

                nome_arquivo_novo = FileManager.obter_proxima_revisao(self, caminho_pasta_existente_engenharia, nome_arquivo_original[:12]) + "." + nome_arquivo_original.split(".")[1]
                self.caminho_arquivo_destino_engenharia = os.path.join(caminho_pasta_existente_engenharia, nome_arquivo_novo)
                self.caminho_arquivo_destino_suprimentos = os.path.join(caminho_pasta_existente_suprimentos, nome_arquivo_novo)

                ExcelHandler.abrir_alterar_arquivo(self, nome_arquivo_novo[:12])


            # Verifica se o arquivo de origem existe
            if os.path.exists(self.caminho_arquivo.get()):
                if self.acao_var.get() == "nova":
                    # Copia o arquivo e renomeia
                    shutil.copy(self.caminho_arquivo.get(), self.caminho_arquivo_destino_engenharia)
                    shutil.move(self.caminho_arquivo.get(), self.caminho_arquivo_destino_suprimentos)
                    FileManager.criar_pasta_superado(self, caminho_nova_pasta_engenharia, caminho_nova_pasta_suprimentos)
                    messagebox.showinfo("Atenção", "Arquivo criado com sucesso!")
                elif self.acao_var.get() == "existente":
                    shutil.copy(self.caminho_arquivo.get(), self.caminho_arquivo_destino_engenharia)
                    shutil.copy(self.caminho_arquivo.get(), self.caminho_arquivo_destino_suprimentos)
                    self.param = 1
                    print(self.caminho_arquivo_destino_engenharia)
                    FileManager.abrir_arquivo(self, self.caminho_arquivo_destino_engenharia)
            else:
                messagebox.showinfo("Atenção", "Arquivo não foi copiado!")
        except Exception as e:
            print(f"Ocorreu um erro: {e}")

