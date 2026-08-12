import os
import customtkinter as ctk
from excel.exporter import ExcelExporter
from tkinter import messagebox


class ProjectManager(ctk.CTkFrame):

    def __init__(self, master, open_project_callback):
        super().__init__(master)

        self.open_project_callback = open_project_callback

        self.configure(
            fg_color="#1E1E1E"
        )

        self.create_layout()
        self.load_projects()

    # =====================================================
    # LAYOUT
    # =====================================================

    def create_layout(self):

        # -------------------------------------------------
        # Ana başlık
        # -------------------------------------------------

        title = ctk.CTkLabel(
            self,
            text="🧪 Test Case Yönetim Sistemi",
            font=("Segoe UI", 30, "bold")
        )

        title.pack(
            pady=(40, 10)
        )

        # -------------------------------------------------
        # Alt başlık
        # -------------------------------------------------

        subtitle = ctk.CTkLabel(
            self,
            text="Test yapmak istediğiniz uygulamayı seçin",
            font=("Segoe UI", 16),
            text_color="#AAAAAA"
        )

        subtitle.pack(
            pady=(0, 25)
        )

        # -------------------------------------------------
        # Proje alanı
        # -------------------------------------------------

        self.project_frame = ctk.CTkFrame(
            self,
            width=700,
            corner_radius=15
        )

        self.project_frame.pack(
            fill="both",
            expand=True,
            padx=250,
            pady=(0, 20)
        )

        # -------------------------------------------------
        # Başlık
        # -------------------------------------------------

        list_title = ctk.CTkLabel(
            self.project_frame,
            text="📁 Projeler / Uygulamalar",
            font=("Segoe UI", 20, "bold")
        )

        list_title.pack(
            anchor="w",
            padx=25,
            pady=(20, 10)
        )

        # -------------------------------------------------
        # Scrollable proje listesi
        # -------------------------------------------------

        self.project_list = ctk.CTkScrollableFrame(
            self.project_frame,
            fg_color="transparent"
        )

        self.project_list.pack(
            fill="both",
            expand=True,
            padx=15,
            pady=10
        )

        # -------------------------------------------------
        # Yeni proje butonu
        # -------------------------------------------------

        new_project_btn = ctk.CTkButton(
            self.project_frame,
            text="➕ Yeni Proje",
            height=45,
            font=("Segoe UI", 15, "bold"),
            command=self.create_project
        )

        new_project_btn.pack(
            fill="x",
            padx=25,
            pady=(10, 25)
        )

    # =====================================================
    # PROJELERİ YÜKLE
    # =====================================================

    def load_projects(self):

        # Mevcut listeyi temizle

        for widget in self.project_list.winfo_children():
            widget.destroy()

        # Proje klasörünü belirle

        project_root = os.path.dirname(
            os.path.dirname(
                os.path.abspath(__file__)
            )
        )

        projects_path = os.path.join(
            project_root,
            "projects"
        )

        # Klasör yoksa oluştur

        os.makedirs(
            projects_path,
            exist_ok=True
        )

        # Projeleri bul

        projects = []

        for item in os.listdir(projects_path):

            item_path = os.path.join(
                projects_path,
                item
            )

            if os.path.isdir(item_path):
                projects.append(item)

        projects.sort(
            key=lambda x: x.lower()
        )

        # -------------------------------------------------
        # Hiç proje yoksa
        # -------------------------------------------------

        if not projects:

            empty_label = ctk.CTkLabel(
                self.project_list,
                text="Henüz bir proje oluşturulmadı.",
                font=("Segoe UI", 15),
                text_color="#999999"
            )

            empty_label.pack(
                pady=40
            )

            return

        # -------------------------------------------------
        # Projeleri oluştur
        # -------------------------------------------------

        for project in projects:

            self.create_project_button(
                project
            )

    # =====================================================
    # PROJE BUTONU
    # =====================================================

    def create_project_button(
        self,
        project_name
    ):

        frame = ctk.CTkFrame(
            self.project_list,
            corner_radius=10
        )

        frame.pack(
            fill="x",
            pady=5
        )

        # -------------------------------------------------
        # Proje ikonu
        # -------------------------------------------------

        icon = ctk.CTkLabel(
            frame,
            text="📁",
            font=("Segoe UI", 24)
        )

        icon.pack(
            side="left",
            padx=(15, 10),
            pady=10
        )

        # -------------------------------------------------
        # Proje adı
        # -------------------------------------------------

        name_label = ctk.CTkLabel(
            frame,
            text=project_name,
            font=("Segoe UI", 16, "bold")
        )

        name_label.pack(
            side="left",
            padx=5
        )

        # =================================================
        # AÇ BUTONU
        # =================================================

        open_button = ctk.CTkButton(
            frame,
            text="Aç",
            width=80,
            height=35,
            command=lambda:
                self.open_project(project_name)
        )

        open_button.pack(
            side="right",
            padx=(5, 10)
        )

        # =================================================
        # SİL BUTONU
        # =================================================

        delete_button = ctk.CTkButton(
            frame,
            text="🗑",
            width=50,
            height=35,
            fg_color="#C0392B",
            hover_color="#922B21",
            command=lambda:
                self.delete_project(project_name)
        )

        delete_button.pack(
            side="right",
            padx=(0, 5)
        )

        # -------------------------------------------------
        # Tüm satıra tıklama
        # -------------------------------------------------

        frame.bind(
            "<Button-1>",
            lambda event:
                self.open_project(project_name)
        )

        icon.bind(
            "<Button-1>",
            lambda event:
                self.open_project(project_name)
        )

        name_label.bind(
            "<Button-1>",
            lambda event:
                self.open_project(project_name)
        )

    # =====================================================
    # PROJE AÇ
    # =====================================================

    def open_project(
        self,
        project_name
    ):

        self.open_project_callback(
            project_name
        )

    # =====================================================
    # YENİ PROJE
    # =====================================================

    def create_project(self):

        dialog = ctk.CTkToplevel(
            self
        )

        dialog.title(
            "Yeni Proje"
        )

        dialog.geometry(
            "500x280"
        )

        dialog.resizable(
            False,
            False
        )

        # NOT: grab_set() kasıtlı olarak kullanılmıyor - bu,
        # NewTestWindow'da tespit edilen aynı hataya (küçültme
        # tuşunun tüm pencereleri dondurması) yol açıyordu.
        # Bunun yerine pencereyi öne getiriyoruz.

        dialog.lift()
        dialog.focus_force()
        dialog.after(150, lambda: (dialog.lift(), dialog.focus_force()))

        # -------------------------------------------------
        # Başlık
        # -------------------------------------------------

        title = ctk.CTkLabel(
            dialog,
            text="➕ Yeni Proje Oluştur",
            font=("Segoe UI", 22, "bold")
        )

        title.pack(
            pady=(30, 15)
        )

        # -------------------------------------------------
        # Açıklama
        # -------------------------------------------------

        label = ctk.CTkLabel(
            dialog,
            text="Proje / Uygulama Adı"
        )

        label.pack(
            pady=(5, 5)
        )

        # -------------------------------------------------
        # Entry
        # -------------------------------------------------

        entry = ctk.CTkEntry(
            dialog,
            width=350,
            height=40,
            placeholder_text="Örn: Mango"
        )

        entry.pack(
            pady=5
        )

        entry.focus()

        # -------------------------------------------------
        # Oluştur
        # -------------------------------------------------

        def save_project():

            project_name = entry.get().strip()

            if not project_name:

                messagebox.showwarning(
                    "Uyarı",
                    "Lütfen proje adı girin.",
                    parent=dialog
                )

                return

            # Yasak karakterler

            invalid_chars = '<>:"/\\|?*'

            if any(
                char in project_name
                for char in invalid_chars
            ):

                messagebox.showerror(
                    "Hata",
                    "Proje adında şu karakterler kullanılamaz:\n"
                    '< > : " / \\ | ? *',
                    parent=dialog
                )

                return

            project_root = os.path.dirname(
                os.path.dirname(
                    os.path.abspath(__file__)
                )
            )

            projects_path = os.path.join(
                project_root,
                "projects"
            )

            project_path = os.path.join(
                projects_path,
                project_name
            )

            # Aynı proje var mı?

            if os.path.exists(project_path):

                messagebox.showwarning(
                    "Uyarı",
                    "Bu proje zaten mevcut.",
                    parent=dialog
                )

                return

            # Klasörü oluştur

            os.makedirs(
                project_path,
                exist_ok=True
            )

            # ============================================
            # PROJE İÇİN EXCEL OLUŞTUR
            # ============================================

            try:

                from database.database import Database

                db = Database(
                    project_name
                )

                exporter = ExcelExporter(
                    db
                )

                exporter.export_project(
                    project_name
                )

                db.close()

            except Exception as e:

                messagebox.showerror(
                    "Excel Hatası",
                    f"Proje oluşturuldu fakat Excel oluşturulamadı:\n\n{e}",
                    parent=dialog
                )

                return

        dialog.destroy()

        self.load_projects()

        messagebox.showinfo(
            "Başarılı",
            f"'{project_name}' projesi oluşturuldu.\n\n"
            f"Excel dosyası da oluşturuldu."
        )

        self.open_project(
            project_name
        )

        # -------------------------------------------------
        # Buton
        # -------------------------------------------------

        create_button = ctk.CTkButton(
            dialog,
            text="Oluştur",
            width=350,
            height=40,
            command=save_project
        )

        create_button.pack(
            pady=(20, 10)
        )

    # =====================================================
    # PROJE SİL
    # =====================================================

    def delete_project(
        self,
        project_name
    ):

        answer = messagebox.askyesno(
            "Projeyi Sil",
            f"'{project_name}' projesini silmek istediğinize emin misiniz?\n\n"
            "Bu işlem proje içindeki tüm test caselerini ve verileri silecektir.",
            parent=self
        )

        if not answer:
            return

        # -------------------------------------------------
        # Proje klasörü
        # -------------------------------------------------

        project_root = os.path.dirname(
            os.path.dirname(
                os.path.abspath(__file__)
            )
        )

        project_path = os.path.join(
            project_root,
            "projects",
            project_name
        )

        # -------------------------------------------------
        # Sil
        # -------------------------------------------------

        try:

            import shutil

            shutil.rmtree(
                project_path
            )

            self.load_projects()

            messagebox.showinfo(
                "Başarılı",
                f"'{project_name}' projesi silindi.",
                parent=self
            )

        except Exception as e:

            messagebox.showerror(
                "Hata",
                f"Proje silinemedi.\n\n{str(e)}",
                parent=self
            )