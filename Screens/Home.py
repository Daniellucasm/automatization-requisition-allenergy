import tkinter as tk
import shutil
import os
from tkinter import ttk
from tkinter import messagebox, filedialog
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

        button_update_request = tk.Button(opcoes_frame, text="Utilizar Uma Requisição Existente",
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
            if self.acao_var.get() == "novo":
                print("novo")
            else:
                print("existente")

            self.copiar_arquivo()
        else:
            messagebox.showinfo("Arquivo não selecionado", "Arquivo não foi anexado")
        
        print(self.caminho_arquivo.get())
        print("Tipo: " + self.tipo_var.get())
        print("Acao: " + self.acao_var.get())
        print("Projeto: " + self.projeto.get())

    def ler_requisicoes_existentes(self):
        print("Fazer a leitura das pastas de requisições")
        if self.tipo_var.get() == "RCO":
            print("RCO")
        else:
            print("RSE")
    
    def copiar_arquivo(self):
        try:
            # Caminho base da pasta "Documents" do usuário
            caminho_base = os.path.expanduser("/Users/daniellucas/Library/Mobile Documents/com~apple~CloudDocs/All Energy/Programa RCO/")
        
            # Caminho de destino específico para o projeto
            caminho_destino = os.path.join(caminho_base, self.projeto.get())
        
            # Cria o diretório de destino se ele não existir
            os.makedirs(caminho_destino, exist_ok=True)
        
            # Verifica se o arquivo de origem existe
            if os.path.exists(self.caminho_arquivo.get()):
                # Copia o arquivo
                shutil.copy(self.caminho_arquivo.get(), caminho_destino)
                print(f"Arquivo copiado com sucesso para o caminho: {caminho_destino}")
            else:
                print("Arquivo de origem não encontrado!")
        except Exception as e:
            print(f"Ocorreu um erro: {e}")
        