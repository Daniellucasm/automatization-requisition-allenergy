from openpyxl import load_workbook

class ExcelHandler():
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