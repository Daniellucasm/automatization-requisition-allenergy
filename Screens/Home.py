import tkinter as tk
import shutil
import os
from Utilities.ExcelHandler import ExcelHandler
from Utilities.FileManager import FileManager
from tkinter import ttk
from tkinter import messagebox, simpledialog
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
        self.requisicao = tk.StringVar(value="")
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
                                command=lambda: self.exibir_botao_salvar("RSE", nova_janela))
            rse_btn.grid(row=0, column=1, padx=10)
        elif self.acao_var.get() == "existente":
            # Botão para selecionar RCO
            rco_btn = ttk.Button(tipo_requisicao_frame, text="RCO", 
                                command=lambda: self.janela_requisicoes_existentes("RCO"))
            rco_btn.grid(row=0, column=0, padx=10)

            # Botão para selecionar RSE
            rse_btn = ttk.Button(tipo_requisicao_frame, text="RSE", 
                                command=lambda: self.janela_requisicoes_existentes("RSE"))
            rse_btn.grid(row=0, column=1, padx=10)
    
    def janela_requisicoes_existentes(self, tipo):
        self.requisicao.set("")
        if not self.projeto.get():
            messagebox.showwarning("Aviso", "Por favor, selecione um projeto antes de continuar.")
            return

        nova_janela = tk.Toplevel(self)
        nova_janela.title("Editar Requisicao")
        nova_janela.geometry("400x150")
        nova_janela.geometry("+{}+{}".format(self.positionRight, self.positionDown))

        # Combox para listar projetos
        label_projeto = tk.Label(nova_janela, text=f"Selecione a {tipo}:")
        label_projeto.pack(pady=5)

        diretorio = "/Users/daniellucas/Library/Mobile Documents/com~apple~CloudDocs/All Energy/"
        #"/Users/daniellucas/Library/Mobile Documents/com~apple~CloudDocs/All Energy/" - MAC
        #C:\Users\daniel.murta\All Energy\Apropriação de Horas - Documentos\Testes - Windows
        diretorio = os.path.join(diretorio, self.projeto.get())
        diretorio = os.path.join(diretorio, "Requisicao/Engenharia/")
        # Lista todos os diretórios no caminho especificado
        pastas = [nome for nome in os.listdir(diretorio) if os.path.isdir(os.path.join(diretorio, nome))]
        pastas = sorted(pastas)

        self.requisicao_combobox = ttk.Combobox(nova_janela, values=pastas, width=30)
        self.requisicao_combobox.pack(pady=5)

        salvar_frame = tk.Frame(nova_janela)
        salvar_frame.pack(pady=20)
        self.requisicao_combobox.bind("<<ComboboxSelected>>", lambda event: self.atualizar_requisicao(event, diretorio))

        self.param = 0
        salvar_btn = ttk.Button(
            salvar_frame,
            text="Salvar Alterações",
            command=self.atualizar_requisicao_alterada
        )
        salvar_btn.pack(pady=5)

    def atualizar_requisicao(self, event, diretorio):
        self.requisicao.set(self.requisicao_combobox.get())
        if self.requisicao.get() != "":
            diretorio = os.path.join(diretorio, self.requisicao.get())
            diretorio = os.path.join(diretorio, FileManager.encontrar_arquivo_mais_proximo(self, diretorio, self.requisicao.get()[:12]))    
            self.caminho_arquivo.set(diretorio)
            FileManager.copiar_arquivo(self)

    def atualizar_requisicao_alterada(self):
        if self.param == 0:
            messagebox.showwarning("Aviso", "Por favor, selecione uma requisicao antes de continuar.")
            return
        shutil.copy(self.caminho_arquivo_destino_engenharia, self.caminho_arquivo_destino_suprimentos)
        messagebox.showinfo("Aviso", "Alterações salvas com sucesso")

    def exibir_botao_salvar(self, tipo, janela):
        self.tipo_var.set(tipo)
        if not self.projeto.get():
            messagebox.showwarning("Aviso", "Por favor, selecione um projeto antes de continuar.")
            return
        
        self.caminho_arquivo.set('XXXX-RCO-000.xlsx' if self.tipo_var.get() == "RCO" else "XXXX-RSE-000.xlsx")
        FileManager.abrir_arquivo(self, self.caminho_arquivo.get())

        # Remove o frame anterior, se existir
        if hasattr(self, "salvar_frame") and self.salvar_frame.winfo_exists():
            self.salvar_frame.destroy()

        # Adiciona o botão Salvar
        self.salvar_frame = tk.Frame(janela)
        self.salvar_frame.pack(pady=20)

        salvar_btn = ttk.Button(
            self.salvar_frame,
            text="Salvar",
            command=lambda: FileManager.copiar_arquivo(self)
        )
        salvar_btn.pack(pady=5)
    
    def atualizar_projeto(self, event):
        self.projeto.set(self.projeto_combobox.get()) 

    def obter_nome_personalizado(self):
        # Abre uma janela de diálogo para o usuário inserir o nome
        nome_personalizado = simpledialog.askstring("Nome da Nova Pasta", "Digite um nome para a nova pasta:")      
        if not nome_personalizado:
            messagebox.showinfo("Atenção", "Nenhum nome foi inserido. Operação cancelada.")
        return nome_personalizado
        