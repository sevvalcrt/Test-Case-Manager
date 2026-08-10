import customtkinter as ctk

from gui.dashboard import Dashboard
from gui.project_manager import ProjectManager
from database.database import Database


class App(ctk.CTk):

    def __init__(self):

        super().__init__()

        # =================================================
        # THEME
        # =================================================

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # =================================================
        # WINDOW
        # =================================================

        self.title(
            "Test Senaryosu Yönetim Sistemi"
        )

        self.geometry(
            "1400x800"
        )

        # =================================================
        # VARIABLES
        # =================================================

        self.db = None

        self.current_project = None

        self.current_screen = None

        # Geri / ileri sistemi
        self.next_project = None

        # =================================================
        # BAŞLANGIÇ
        # =================================================

        self.show_project_manager()

        # =================================================
        # CLOSE
        # =================================================

        self.protocol(
            "WM_DELETE_WINDOW",
            self.on_close
        )

    # =====================================================
    # PROJECT MANAGER
    # =====================================================

    def show_project_manager(self):

        self.clear_screen()

        project_manager = ProjectManager(
            self,
            self.open_project
        )

        project_manager.pack(
            fill="both",
            expand=True
        )

        self.current_screen = project_manager

    # =====================================================
    # PROJECT AÇ
    # =====================================================

    def open_project(self, project_name):

        # -------------------------------------------------
        # Aynı proje zaten açıksa tekrar açma
        # -------------------------------------------------

        if (
            self.current_project == project_name
            and self.current_screen is not None
        ):
            return

        # -------------------------------------------------
        # Eski database bağlantısını kapat
        # -------------------------------------------------

        if self.db is not None:

            try:
                self.db.close()

            except Exception:
                pass

            self.db = None

        # -------------------------------------------------
        # İleri geçmişini temizle
        # -------------------------------------------------

        self.next_project = None

        # -------------------------------------------------
        # Yeni proje
        # -------------------------------------------------

        self.current_project = project_name

        # -------------------------------------------------
        # Database
        # -------------------------------------------------

        self.db = Database(
            project_name
        )

        # -------------------------------------------------
        # Dashboard
        # -------------------------------------------------

        self.clear_screen()

        dashboard = Dashboard(
            self,
            self.db,
            project_name,
            self.go_back,
            self.go_forward
        )

        dashboard.pack(
            fill="both",
            expand=True
        )

        self.current_screen = dashboard

    # =====================================================
    # GERİ
    # =====================================================

    def go_back(self):

        # Açık proje yoksa
        if self.current_project is None:
            return

        # -------------------------------------------------
        # Şu anki projeyi ileri için sakla
        # -------------------------------------------------

        self.next_project = self.current_project

        # -------------------------------------------------
        # Database'i kapat
        # -------------------------------------------------

        if self.db is not None:

            try:
                self.db.close()

            except Exception:
                pass

            self.db = None

        # -------------------------------------------------
        # ARTIK PROJE AÇIK DEĞİL
        # -------------------------------------------------

        self.current_project = None

        # -------------------------------------------------
        # Proje listesine dön
        # -------------------------------------------------

        self.show_project_manager()

    # =====================================================
    # İLERİ
    # =====================================================

    def go_forward(self):

        # İleri gidilecek proje yoksa
        if self.next_project is None:
            return

        project_name = self.next_project

        # Önce temizle
        self.next_project = None

        # Projeyi tekrar aç
        self.open_project(
            project_name
        )

    # =====================================================
    # SCREEN CLEAR
    # =====================================================

    def clear_screen(self):

        for widget in self.winfo_children():
            widget.destroy()

    # =====================================================
    # CLOSE
    # =====================================================

    def on_close(self):

        if self.db is not None:

            try:
                self.db.close()

            except Exception:
                pass

        self.destroy()