import customtkinter as ctk
from gui.dashboard import Dashboard


class App(ctk.CTk):

    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title("Test Case Management System")
        self.geometry("1400x800")

        dashboard = Dashboard(self)
        dashboard.pack(fill="both", expand=True)