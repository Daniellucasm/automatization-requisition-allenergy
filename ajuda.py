import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import re
from datetime import date

# Lista para armazenar os materiais
materiais = []  

# Lista de projetos
projetos = ["1804 - PCH Foz do Estrela", "2001 - Miringuava", "2002 - PCH Gafanhoto", "2101 - Ampliação SE UHE Itutinga",
                "2102 - Reforços SE Itutinga", "2103 - SE Barreiro", "2104 - Painel UHE NLP", "2201 - UHE Simplício", 
                "2202 - Ampliação SE Sete Lagoas 4", "2203 - UHE São Simão", "2204 - SE Teresina III", "2205 - UHE Henry Borden", 
                "2206 - SE Iriri", "2207 - Eclusa", "2301 - SE Délio Bernardino", "2302 - ELs SE GV6 e SE Verona", 
                "2303 - Implantação de SECIs", "2304 - LT PCH Santa Luzia", "2305 - Monovia da PCH Gafanhoto", "2306 - SE Ibicoara",
                "2307 - Itutinga-Ipatinga", "2308 - SEs Híbridas", "2309 - ELs 500kV SEs Janaúba e Pres. Juscelino", 
                "2401 - Barragem de Ceraíma", "2402 - Híbridas 3 - CEMIG", "2403 - Modernização UHE Salto Grande"]

# Função para adicionar um material à lista
def adicionar_material():
    quantidade = entrada_quantidade.get()
    unidade = entrada_unidade.get()
    descricao_material = entrada_descricao_material.get()
    
    if not (quantidade and unidade and descricao_material):
        messagebox.showwarning("Atenção", "Por favor, preencha todos os campos de material!")
        return
    
    materiais.append({
        'quantidade': quantidade,
        'unidade': unidade,
        'descricao': descricao_material
    })
    
    # Limpa os campos de entrada após adicionar
    entrada_quantidade.delete(0, tk.END)
    entrada_unidade.delete(0, tk.END)
    entrada_descricao_material.delete(0, tk.END)
    
    # Atualiza a lista visual dos materiais
    atualizar_lista_materiais()

# Função para exibir os materiais na interface
def atualizar_lista_materiais():
    lista_materiais_text.delete(1.0, tk.END)  # Limpa a exibição anterior
    for i, material in enumerate(materiais, 1):
        lista_materiais_text.insert(tk.END, f"{i}. Quantidade: {material['quantidade']}, Unidade: {material['unidade']}, Descrição: {material['descricao']}\n")

# Função para validar a data no formato DD/MM/AAAA
def validar_data(data):
    padrao_data = r"\d{2}/\d{2}/\d{4}"
    if re.fullmatch(padrao_data, data):
        if date.today().strftime('%d/%m/%Y') < data:
            return True
        else: 
            return False
    else: 
        return False

# Função para capturar os dados preenchidos e gerar a requisição
def gerar_requisicao():
    projeto_selecionado = projeto_combobox.get()
    finalidade = entrada_finalidade.get()
    data_entrega = entrada_data_entrega.get()
    local_entrega = entrada_local_entrega.get()
    
    # Validação do formato da data
    if not validar_data(data_entrega):
        texto = f"Por favor, insira a data no formato e posterior ao dia {date.today().strftime('%d/%m/%Y')}."
        messagebox.showwarning("Data Inválida", texto)
        return

    if not (projeto_selecionado and finalidade and data_entrega and local_entrega and materiais):
        messagebox.showwarning("Atenção", "Por favor, preencha todos os campos e adicione pelo menos um material!")
        return
    
    # Exibir as informações coletadas
    requisicao_info = f"""
    Projeto: {projeto_selecionado}
    Finalidade: {finalidade}
    Data de Entrega: {data_entrega}
    Local de Entrega: {local_entrega}
    Materiais:\n"""
    
    for i, material in enumerate(materiais, 1):
        requisicao_info += f"{i}. Quantidade: {material['quantidade']}, Unidade: {material['unidade']}, Descrição: {material['descricao']}\n"
    
    messagebox.showinfo("Requisição Gerada", requisicao_info)

# Função para abrir o formulário em uma nova janela, dependendo do tipo selecionado
def abrir_formulario():
    tipo = tipo_var.get()
    if tipo == "":
        messagebox.showwarning("Seleção Necessária", "Selecione RCO ou RSE para continuar.")
        return
    elif tipo == "RCO":
        abrir_tela_rco_janela()
    else:
        abrir_tela_rse()

def abrir_tela_rco_janela():
    global projeto_combobox, entrada_finalidade, entrada_data_entrega, entrada_local_entrega
    global entrada_quantidade, entrada_unidade, entrada_descricao_material, lista_materiais_text
    
    janela_rco = tk.Toplevel()  # Cria uma nova janela
    janela_rco.title("Requisição de Compra")
    janela_rco.geometry("400x600")
    janela_rco.anchor("center")

    # Fecha o programa quando a janela é fechada
    #janela_rco.protocol("WM_DELETE_WINDOW", fechar_programa)

    # Label e Combobox para seleção do projeto
    label_projeto = tk.Label(janela_rco, text="Projeto:")
    label_projeto.pack(pady=5)

    projeto_combobox = ttk.Combobox(janela_rco, values=projetos, width=30)
    projeto_combobox.pack(pady=5)

    # Campos para preenchimento das informações gerais
    label_finalidade = tk.Label(janela_rco, text="Finalidade:")
    label_finalidade.pack(pady=5)
    entrada_finalidade = tk.Entry(janela_rco, width=30)
    entrada_finalidade.pack(pady=5)

    label_data_entrega = tk.Label(janela_rco, text="Data de Entrega (DD/MM/AAAA):")
    label_data_entrega.pack(pady=5)
    entrada_data_entrega = tk.Entry(janela_rco, width=30)
    entrada_data_entrega.pack(pady=5)

    label_local_entrega = tk.Label(janela_rco, text="Local de Entrega:")
    label_local_entrega.pack(pady=5)
    entrada_local_entrega = tk.Entry(janela_rco, width=30)
    entrada_local_entrega.pack(pady=5)

    # Campos para adicionar materiais
    label_quantidade = tk.Label(janela_rco, text="Quantidade:")
    label_quantidade.pack(pady=5)
    entrada_quantidade = tk.Entry(janela_rco, width=30)
    entrada_quantidade.pack(pady=5)

    label_unidade = tk.Label(janela_rco, text="Unidade:")
    label_unidade.pack(pady=5)
    entrada_unidade = tk.Entry(janela_rco, width=30)
    entrada_unidade.pack(pady=5)

    label_descricao_material = tk.Label(janela_rco, text="Descrição do Material:")
    label_descricao_material.pack(pady=5)
    entrada_descricao_material = tk.Entry(janela_rco, width=30)
    entrada_descricao_material.pack(pady=5)

    # Botão para adicionar material à lista
    botao_adicionar_material = tk.Button(janela_rco, text="Adicionar Material", command=adicionar_material)
    botao_adicionar_material.pack(pady=10)

    # Texto para exibir a lista de materiais adicionados
    lista_materiais_text = tk.Text(janela_rco, height=10, width=50)
    lista_materiais_text.pack(pady=5)

    # Botão para gerar a requisição
    botao_gerar = tk.Button(janela_rco, text="Gerar Requisição", command=gerar_requisicao)
    botao_gerar.pack(pady=20)

def abrir_tela_rse():
    global projeto_combobox, entrada_finalidade, entrada_data_entrega, entrada_local_entrega
    
    janela_rse = tk.Toplevel()  # Cria uma nova janela
    janela_rse.title("Requisição de Serviço")
    janela_rse.geometry("400x600")
    janela_rse.anchor("center")

    # Fecha o programa quando a janela é fechada
    #janela_rco.protocol("WM_DELETE_WINDOW", fechar_programa)

    # Label e Combobox para seleção do projeto
    label_projeto = tk.Label(janela_rse, text="Projeto:")
    label_projeto.pack(pady=5)

    projeto_combobox = ttk.Combobox(janela_rse, values=projetos, width=30)
    projeto_combobox.pack(pady=5)

    # Campos para preenchimento das informações gerais
    label_finalidade = tk.Label(janela_rse, text="Finalidade:")
    label_finalidade.pack(pady=5)
    entrada_finalidade = tk.Entry(janela_rse, width=30)
    entrada_finalidade.pack(pady=5)

    label_data_entrega = tk.Label(janela_rse, text="Data de Entrega (DD/MM/AAAA):")
    label_data_entrega.pack(pady=5)
    entrada_data_entrega = tk.Entry(janela_rse, width=30)
    entrada_data_entrega.pack(pady=5)

    label_local_entrega = tk.Label(janela_rse, text="Local de Entrega:")
    label_local_entrega.pack(pady=5)
    entrada_local_entrega = tk.Entry(janela_rse, width=30)
    entrada_local_entrega.pack(pady=5)

    # Botão para gerar a requisição
    botao_gerar = tk.Button(janela_rse, text="Gerar Requisição", command=gerar_requisicao)
    botao_gerar.pack(pady=20)

    # Botão para Anexar arquivo
    botao_gerar = tk.Button(janela_rse, text="Anexar arquivo já existente", command=None)
    botao_gerar.pack(pady=20)

# Função para enviar os dados do formulário
def enviar_formulario(tipo, campo1, campo2):
    print(f"Enviando {tipo} - Campo 1: {campo1}, Campo 2: {campo2}")
    # Implementar o salvamento na requisição

# Função para mostrar opções de requisições existentes
def selecionar_requisicao_existente():
    tipo_requisicao_frame.pack(fill="x", padx=10, pady=(5, 10))
    lista_requisicoes.pack_forget()

# Função para listar RCOs ou RSEs disponíveis após seleção
def listar_requisicoes(tipo):
    lista_requisicoes.pack(fill="x", padx=10, pady=(5, 10))
    lista_requisicoes_label.config(text=f"Lista de {tipo}:")
    # Atualizar com as requisições existentes, exemplo fictício:
    opcoes_requisicoes = ["RCO001", "RCO002", "RCO003"] if tipo == "RCO" else ["RSE001", "RSE002", "RSE003"]
    lista_requisicoes_combobox['values'] = opcoes_requisicoes
    lista_requisicoes_combobox.set("Selecione uma Requisição")

# Função para mostrar ou ocultar opções de RCO e RSE
def atualizar_opcoes_tipo():
    if acao_var.get() == "nova":
        tipo_requisicao_frame.pack(fill="x", padx=10, pady=(5, 10))
        lista_requisicoes.pack_forget()  # Esconde a lista de requisições
    elif acao_var.get() == "existente":
        tipo_requisicao_frame.pack(fill="x", padx=10, pady=(5, 10))
        lista_requisicoes.pack_forget()  # Esconde inicialmente
        tipo_var.set("")  # Limpa a seleção anterior de RCO ou RSE

# Inicialização da janela principal
root = tk.Tk()
root.title("Menu de Requisição")
root.geometry("400x300")
root.anchor("center")

# Variável para armazenar a ação selecionada (Criar Nova ou Usar Existente)
acao_var = tk.StringVar(value="")
tipo_var = tk.StringVar(value="")  # Variável para armazenar o tipo selecionado (RCO ou RSE)

# Frame principal com botões de seleção de requisição
opcoes_frame = tk.Frame(root)
opcoes_frame.pack(fill="both", padx=10, pady=10)

# Botão para criar uma nova requisição ou usar uma existente
nova_requisicao_radio = ttk.Radiobutton(opcoes_frame, text="Criar uma Requisição Nova", variable=acao_var, value="nova", command=atualizar_opcoes_tipo)
nova_requisicao_radio.grid(row=0, column=0, sticky="w")

existente_requisicao_radio = ttk.Radiobutton(opcoes_frame, text="Utilizar uma Já Existente", variable=acao_var, value="existente", command=atualizar_opcoes_tipo)
existente_requisicao_radio.grid(row=1, column=0, sticky="w")

# Frame para selecionar RCO ou RSE após escolher "Utilizar uma Já Existente"
tipo_requisicao_frame = tk.Frame(root)

# Botão para selecionar RCO
rco_btn = ttk.Button(tipo_requisicao_frame, text="RCO", command=lambda: abrir_tela_rco_janela() if acao_var.get() == "nova" else listar_requisicoes("RCO"))
rco_btn.grid(row=0, column=0, padx=5)

# Botão para selecionar RSE
rse_btn = ttk.Button(tipo_requisicao_frame, text="RSE", command=lambda: abrir_tela_rse() if acao_var.get() == "nova" else listar_requisicoes("RSE"))
rse_btn.grid(row=0, column=1, padx=5)

# Frame para listar as requisições disponíveis
lista_requisicoes = tk.Frame(root)
lista_requisicoes_label = tk.Label(lista_requisicoes, text="Lista de Requisições:")
lista_requisicoes_label.pack(side="top")

# Combobox para exibir a lista de RCOs ou RSEs existentes
lista_requisicoes_combobox = ttk.Combobox(lista_requisicoes)
lista_requisicoes_combobox.pack(fill="x")

root.mainloop()
