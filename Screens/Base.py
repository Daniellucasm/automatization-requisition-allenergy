import tkinter as tk

class BaseScreen(tk.Tk):

    # Lista de projetos
    projetos = ["1804 - PCH FOZ DO ESTRELA", "2001 - MIRINGUAVA", "2002 - PCH GAFANHOTO", "2101 - AMPLIAÇÃO SE UHE ITUTINGA",
                "2102 - REFORÇOS SE ITUTINGA", "2103 - SE BARREIRO 1R", "2104 - PAINEL UHE NLP", "2201 - UHE SIMPLÍCIO - LOG BOOM", 
                "2202 - Ampliação SE Sete Lagoas 4", "2203 - UHE São Simão - BOP Mecânico", "2204 - SE Teresina III - 69 kV", 
                "2205 - UHE Henry Borden - Tubulação de Água de Refrigeração", "2206 - SE Iriri", "2207 - Eclusa", 
                "2301 - SE Délio Bernardino", "2302 - ELs SE GV6 e SE Verona", "2303 - Implantação de SECIs", 
                "2304 - LT PCH Santa Luzia", "2305 - Monovia da PCH Gafanhoto", "2306 - SE Ibicoara",
                "2307 - Itutinga-Ipatinga", "2308 - SEs Híbridas", "2309 - ELs 500kV SEs Janaúba e Pres. Juscelino", 
                "2401 - Barragem Ceraíma", "2402 - Híbridas 3 - CEMIG", "2403 - Modernização UHE Salto Grande"]

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        #-----------------------------
        #       Configurações
        #-----------------------------

        self.title("Programa Requisição")
        self.geometry("400x300")
        #self.configure(bg='black')

        windowWidth = self.winfo_reqwidth()
        windowHeight = self.winfo_reqheight()

        self.positionRight = int(self.winfo_screenwidth()/2 - windowWidth/2)
        self.positionDown = int(self.winfo_screenheight()/2 - windowHeight/2)

        # Positions the window in the center of the page.
        self.geometry("+{}+{}".format(self.positionRight, self.positionDown))

        #---------------------------
        #        Componentes
        #---------------------------