from Screens.Home import HomeScreen


class ScreenManager:
    def __init__(self):
        self.tela = HomeScreen(manager=self)
        self.tela.mainloop()