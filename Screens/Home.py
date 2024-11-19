import tkinter as tk
import shutil
import os
import re
import subprocess
import platform
import time
from tkinter import ttk
from tkinter import messagebox, filedialog, simpledialog
from Screens.Base import BaseScreen
from openpyxl import load_workbook

class HomeScreen(BaseScreen):

    def __init__(self, manager, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # -----------------------------
        #       Configurações
        # -----------------------------
        self.title("Programa Requisição")
        self.geometry("400x300")
        #self.configure(bg='grey')
        self.acao_var = tk.StringVar(value="")

        # ---------------------------
        #        Componentes
        # ---------------------------
        label = tk.Label(self, text="Menu Requisição", font=("Helvetica", 20), pady=20) #background="grey"
        label.pack()

        # Variáveis para armazenar a ação selecionada (Criar Nova ou Usar Existente), tipo de requisição e projeto
        self.tipo_var = tk.StringVar(value="")
        self.projeto = tk.StringVar(value="")
        self.revisao = tk.StringVar(value="")
        self.caminho_arquivo = tk.StringVar(value="")

        # Frame principal com botões de seleção de requisição
        opcoes_frame = tk.Frame(self)
        opcoes_frame.pack(fill="both", padx=10, pady=10)

        button_new_request = tk.Button(opcoes_frame, text="Criar Uma Nova Requisição",
                                       command=lambda: self.abrir_nova_janela("nova"))
        button_new_request.pack(pady=10)

        button_update_request = tk.Button(opcoes_frame, text="Editar Requisição Existente",
                                          command=lambda: self.abrir_nova_janela("existente"))
        button_update_request.pack(pady=10)


    def abrir_nova_janela(self, acao):
        """Abre uma nova janela para criar uma nova requisição."""
        self.acao_var.set(acao)
        self.projeto.set("")
        nova_janela = tk.Toplevel(self)
        nova_janela.title("Nova Requisição" if acao == "nova" else "Editar Requisicao")
        nova_janela.geometry("400x300")
        nova_janela.geometry("+{}+{}".format(self.positionRight, self.positionDown))

        # Combox para listar projetos
        label_projeto = tk.Label(nova_janela, text="Selecione o Projeto:")
        label_projeto.pack(pady=5)

        self.projeto_combobox = ttk.Combobox(nova_janela, values=self.projetos, width=30)
        self.projeto_combobox.pack(pady=5)

        # Frame para os botões de tipo de requisição
        tipo_requisicao_frame = tk.Frame(nova_janela)
        tipo_requisicao_frame.pack(pady=20)
        self.projeto_combobox.bind("<<ComboboxSelected>>", self.atualizar_projeto)

        if self.acao_var.get() == "nova":
            # Botão para selecionar RCO
            rco_btn = ttk.Button(tipo_requisicao_frame, text="RCO", 
                                command=lambda: self.exibir_botao_salvar("RCO", nova_janela))  
            rco_btn.grid(row=0, column=0, padx=10)

            # Botão para selecionar RSE
            rse_btn = ttk.Button(tipo_requisicao_frame, text="RSE", 
                                command=lambda: self.exibir_botao_salvar("RCO", nova_janela))
            rse_btn.grid(row=0, column=1, padx=10)
        elif self.acao_var.get() == "existente":
            # Botão para selecionar RCO
            rco_btn = ttk.Button(tipo_requisicao_frame, text="RCO", 
                                command=lambda: self.selecionar_arquivo("RCO"))
            rco_btn.grid(row=0, column=0, padx=10)

            # Botão para selecionar RSE
            rse_btn = ttk.Button(tipo_requisicao_frame, text="RSE", 
                                command=lambda: self.selecionar_arquivo("RSE"))
            rse_btn.grid(row=0, column=1, padx=10)

    def exibir_botao_salvar(self, tipo, janela):
        """Exibe o botão Salvar após selecionar RCO ou RSE."""
        if not self.projeto.get():
            messagebox.showwarning("Aviso", "Por favor, selecione um projeto antes de continuar.")
            return

        self.abrir_arquivo(tipo)

        # Adiciona o botão Salvar
        salvar_frame = tk.Frame(janela)
        salvar_frame.pack(pady=20)

        salvar_btn = ttk.Button(
            salvar_frame,
            text="Salvar",
            command=lambda: print("SALVAR")
        )
        salvar_btn.pack(pady=5)
    
    def atualizar_projeto(self, event):
        self.projeto.set(self.projeto_combobox.get()) 

    def abrir_arquivo(self, tipo):
        # Caminho do arquivo Excel
        self.tipo_var.set(tipo)
        file_path = 'XXXX-RCO-000.xlsx' if self.tipo_var.get() == "RCO" else "XXXX-RSE-000.xlsx"

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

    def selecionar_arquivo(self, tipo):
        """ Abre uma janela de diálogo para selecionar o arquivo .xlsx """
        self.tipo_var.set(tipo)
        if not self.projeto.get():
            messagebox.showwarning("Aviso", "Por favor, selecione um projeto antes de continuar.")
            return

        file_path = filedialog.askopenfilename(
            title="Selecione um Arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx"), ["Arquivos Excel", "*.xls"]])
        
        if file_path:
            messagebox.showinfo("Arquivo Selecionado", "Arquivo anexado com sucesso!")
            self.caminho_arquivo.set(file_path)
            self.copiar_arquivo()
        else:
            messagebox.showinfo("Arquivo não selecionado", "Arquivo não foi anexado")
        

    def requisicoes_existentes(self, diretorio):
        print("Fazer a leitura das pastas de requisições")
        if self.tipo_var.get() == "RCO":
            prefixo = self.projeto.get()[:4] + "-" + self.tipo_var.get()  # Número do código de projeto + RCO EX: 1804-RCO
            nova_pasta = self.ler_requisicoes_existentes(diretorio, prefixo)
        else:
            prefixo = self.projeto.get()[:4] + "-" + self.tipo_var.get()  # Número do código de projeto + RSE EX: 1804-RSE
            nova_pasta = self.ler_requisicoes_existentes(diretorio, prefixo)
        return nova_pasta

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

    def obter_nome_personalizado(self):
        # Abre uma janela de diálogo para o usuário inserir o nome
        nome_personalizado = simpledialog.askstring("Nome da Nova Pasta", "Digite um nome para a nova pasta:")      
        if not nome_personalizado:
            messagebox.showinfo("Atenção", "Nenhum nome foi inserido. Operação cancelada.")
        return nome_personalizado
    
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
        
    def abrir_alterar_arquivo(self, num_requisicao):
        workbook = load_workbook(filename=self.caminho_arquivo.get())
        sheet = workbook["Formulário"]

        print(num_requisicao)

        # Alterar o valor da célula B3
        sheet["G4"] = num_requisicao
        # Quebrando a string pelo hífen e armazenando em uma lista
        partes_projeto = self.projeto.get().split("-") #número requisição
        # Removendo espaços em branco de cada parte e armazenando em uma nova lista
        partes_limpas = [parte.strip() for parte in partes_projeto]

        sheet["A6"] = partes_limpas[0] #número projeto
        sheet["D6"] = partes_limpas[1] #nome projeto

        # Salvar as mudanças no arquivo Excel
        workbook.save(self.caminho_arquivo.get())


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
                nome_nova_pasta = self.requisicoes_existentes(caminho_destino_engenharia) 
                if not nome_nova_pasta:  # Verifica se o usuário cancelou a operação
                    return
                
                caminho_nova_pasta_engenharia = os.path.join(caminho_destino_engenharia, nome_nova_pasta)
                caminho_nova_pasta_suprimentos = os.path.join(caminho_destino_suprimentos, nome_nova_pasta)
                os.makedirs(caminho_nova_pasta_engenharia, exist_ok=True)
                os.makedirs(caminho_nova_pasta_suprimentos, exist_ok=True)
                # Define o novo nome do arquivo com base no nome da nova pasta
                nome_arquivo_novo = nome_nova_pasta[:12] + "-R0" + os.path.splitext(self.caminho_arquivo.get())[1]
                caminho_arquivo_destino_engenharia = os.path.join(caminho_nova_pasta_engenharia, nome_arquivo_novo)
                caminho_arquivo_destino_suprimentos = os.path.join(caminho_nova_pasta_suprimentos, nome_arquivo_novo)

                self.abrir_alterar_arquivo(nome_nova_pasta[:12])
                
            else:
                # Obtém o nome do arquivo original
                nome_arquivo_original = os.path.basename(self.caminho_arquivo.get())
                
                caminho_pasta_existente_engenharia = self.encontrar_pasta_por_prefixo(caminho_destino_engenharia, nome_arquivo_original[:12])
                caminho_pasta_existente_suprimentos = self.encontrar_pasta_por_prefixo(caminho_destino_suprimentos, nome_arquivo_original[:12])
                print("Engenharia: ", caminho_pasta_existente_engenharia)
                print("Suprimentos: ", caminho_pasta_existente_suprimentos)
                # Define o caminho da pasta existente onde o arquivo será copiado
                #caminho_pasta_existente_engenharia = os.path.join(caminho_destino_engenharia, nome_arquivo_original.split(" ")[0])
                #caminho_pasta_existente_suprimentos = os.path.join(caminho_destino_suprimentos, nome_arquivo_original.split(" ")[0])

                nome_arquivo_novo = self.obter_proxima_revisao(caminho_pasta_existente_engenharia, nome_arquivo_original[:12]) + "." + nome_arquivo_original.split(".")[1]
                caminho_arquivo_destino_engenharia = os.path.join(caminho_pasta_existente_engenharia, nome_arquivo_novo)
                caminho_arquivo_destino_suprimentos = os.path.join(caminho_pasta_existente_suprimentos, nome_arquivo_novo)

                self.abrir_alterar_arquivo(nome_arquivo_novo[:12])


            # Verifica se o arquivo de origem existe
            if os.path.exists(self.caminho_arquivo.get()):
                # Copia o arquivo e renomeia
                shutil.copy(self.caminho_arquivo.get(), caminho_arquivo_destino_engenharia)
                shutil.copy(self.caminho_arquivo.get(), caminho_arquivo_destino_suprimentos)
                messagebox.showinfo("Atenção", "Arquivo copiado com sucesso")
            else:
                messagebox.showinfo("Atenção", "Arquivo não foi copiado!")
        except Exception as e:
            print(f"Ocorreu um erro: {e}")

        