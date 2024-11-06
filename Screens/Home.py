import tkinter as tk
import shutil
import os
import re
from tkinter import ttk
from tkinter import messagebox, filedialog, simpledialog
from Screens.Base import BaseScreen


class HomeScreen(BaseScreen):

    def __init__(self, manager, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # -----------------------------
        #       Configurações
        # -----------------------------
        self.title("Programa Requisição")
        self.geometry("400x300")
        #self.configure(bg='grey')

        # ---------------------------
        #        Componentes
        # ---------------------------
        label = tk.Label(self, text="Menu Requisição", font=("Helvetica", 20), pady=20) #background="grey"
        label.pack()

        # Variáveis para armazenar a ação selecionada (Criar Nova ou Usar Existente), tipo de requisição e projeto
        self.acao_var = tk.StringVar(value="")
        self.tipo_var = tk.StringVar(value="")
        self.projeto = tk.StringVar(value="")
        self.caminho_arquivo = tk.StringVar(value="")

        # Frame principal com botões de seleção de requisição
        opcoes_frame = tk.Frame(self)#bg='gray'
        opcoes_frame.pack(fill="both", padx=10, pady=10)

        button_new_request = tk.Button(opcoes_frame, text="Criar Uma Nova Requisição",
                                       command=lambda: self.atualizar_opcoes("nova"))
        button_new_request.pack(pady=10)

        button_update_request = tk.Button(opcoes_frame, text="Editar Requisição Existente",
                                          command=lambda: self.atualizar_opcoes("existente"))
        button_update_request.pack(pady=10)

        self.projeto_combox_frame = tk.Frame(self)
        self.projeto_combox_frame.columnconfigure(0, weight=1)

        #Combox para listar projetos
        self.projeto_combobox = ttk.Combobox(self.projeto_combox_frame, values=self.projetos, width=30)
        self.projeto_combobox.pack(pady=5)
        self.projeto_combobox.bind("<<ComboboxSelected>>", self.atualizar_projeto)

        # Frame para os botões de tipo de requisição (RCO ou RSE)
        self.tipo_requisicao_frame = tk.Frame(self) #bg='gray'
        # Configurar o frame para centralizar os botões
        self.tipo_requisicao_frame.columnconfigure(1, weight=1)
        self.tipo_requisicao_frame.columnconfigure(2, weight=1)

        # Botão para selecionar RCO
        self.rco_btn = ttk.Button(self.tipo_requisicao_frame, text="RCO", command=lambda: self.selecionar_arquivo("RCO")
                                                                        if self.projeto.get() != "" 
                                                                        else messagebox.showinfo("Projeto não Selecionado", "Escolher um projeto antes de gerar requisição!"))
        self.rco_btn.grid(row=0, column=1, padx=20, sticky="ew")  # "ew" para expandir horizontalmente

        # Botão para selecionar RSE
        self.rse_btn = ttk.Button(self.tipo_requisicao_frame, text="RSE", command=lambda: self.selecionar_arquivo("RSE") 
                                                                        if self.projeto.get() != "" 
                                                                        else messagebox.showinfo("Projeto não Selecionado", "Escolher um projeto antes de gerar requisição!"))
        self.rse_btn.grid(row=0, column=2, padx=20, sticky="ew")  # "ew" para expandir horizontalmente

    def atualizar_projeto(self, event):
        self.projeto.set(self.projeto_combobox.get()) 
          
    def atualizar_opcoes(self, acao):
        """ Atualiza a interface para mostrar os botões de seleção de RCO ou RSE com base na ação selecionada """
        self.acao_var.set(acao)

        # Exibe o frame para seleção de RCO ou RSE dependendo da ação
        if acao == "nova":
            self.projeto_combox_frame.pack()
            self.tipo_requisicao_frame.pack(fill="x", padx=20, pady=10)
        elif acao == "existente":
            self.projeto_combox_frame.pack()
            self.tipo_requisicao_frame.pack(fill="x", padx=20, pady=10)
        else:
            self.projeto_combox_frame.pack_forget()
            self.tipo_requisicao_frame.pack_forget()

    def selecionar_arquivo(self, tipo):
        """ Abre uma janela de diálogo para selecionar o arquivo .xlsx """
        self.tipo_var.set(tipo)

        file_path = filedialog.askopenfilename(
            title="Selecione um Arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx"), ["Arquivos Excel", "*.xls"]])
        
        if file_path:
            messagebox.showinfo("Arquivo Selecionado", "Arquivo anexado com sucesso!")
            self.caminho_arquivo.set(file_path)
            if self.acao_var.get() == "nova":
                print("nova")
            else:
                print("existente")

            self.copiar_arquivo()
        else:
            messagebox.showinfo("Arquivo não selecionado", "Arquivo não foi anexado")
        
        print(self.caminho_arquivo.get())
        print("Tipo: " + self.tipo_var.get())
        print("Acao: " + self.acao_var.get())
        print("Projeto: " + self.projeto.get())

    def requisicoes_existentes(self, diretorio):
        print("Fazer a leitura das pastas de requisições")
        if self.tipo_var.get() == "RCO":
            print("RCO")
            prefixo = self.projeto.get()[:4] + "-" + self.tipo_var.get()  # Número do código de projeto + RCO EX: 1804-RCO
            nova_pasta = self.ler_requisicoes_existentes(diretorio, prefixo)
        else:
            print("RSE")
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

    def copiar_arquivo(self):
        try:
            # Caminho base da pasta "Documents" do usuário
            path = "/Users/daniellucas/Library/Mobile Documents/com~apple~CloudDocs/All Energy/"
            #/Users/daniellucas/Library/Mobile Documents/com~apple~CloudDocs/All Energy/
            #C:\\Users\\daniel.murta\\All Energy\\Apropriação de Horas - Documentos\\Testes
            caminho_base = os.path.expanduser(path)

            # Caminho de destino específico para o projeto
            caminho_destino = os.path.join(caminho_base, self.projeto.get())
            caminho_destino_engenharia = os.path.join(caminho_destino, "Engenharia")
            caminho_destino_suprimentos = os.path.join(caminho_destino, "Suprimentos")

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
                nome_arquivo_novo = nome_nova_pasta + os.path.splitext(self.caminho_arquivo.get())[1]
                caminho_arquivo_destino_engenharia = os.path.join(caminho_nova_pasta_engenharia, nome_arquivo_novo)
                caminho_arquivo_destino_suprimentos = os.path.join(caminho_nova_pasta_suprimentos, nome_arquivo_novo)

            else:
                # Obtém o nome do arquivo original
                nome_arquivo_original = os.path.basename(self.caminho_arquivo.get())
                # Define o caminho da pasta existente onde o arquivo será copiado
                caminho_pasta_existente_engenharia = os.path.join(caminho_destino_engenharia, nome_arquivo_original.split(".")[0])
                caminho_pasta_existente_suprimentos = os.path.join(caminho_destino_suprimentos, nome_arquivo_original.split(".")[0])
                caminho_arquivo_destino_engenharia = os.path.join(caminho_pasta_existente_engenharia, nome_arquivo_original)
                caminho_arquivo_destino_suprimentos = os.path.join(caminho_pasta_existente_suprimentos, nome_arquivo_original)


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

        