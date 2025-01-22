import os
import tkinter as tk

class BaseScreen(tk.Tk):

    # Lista de projetos
    projetos = []

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

        self.load_projetos()

    def load_projetos(self):
        directory = "/Users/daniellucas/Library/Mobile Documents/com~apple~CloudDocs/All Energy/Projetos"
        BaseScreen.projetos = [
            nome for nome in os.listdir(directory) if os.path.isdir(os.path.join(directory, nome))
        ]
    

