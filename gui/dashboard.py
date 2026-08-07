import customtkinter as ctk
from gui.new_test import NewTestWindow
from database.database import Database
from tkinter import messagebox
from gui.new_test import NewTestWindow
from excel.exporter import ExcelExporter


class Dashboard(ctk.CTkFrame):

    def __init__(self, master):
        super().__init__(master)

        self.db = Database()

        self.configure(fg_color="#1E1E1E")

        self.create_layout()

    def create_layout(self):

        # ==========================
        # Başlık
        # ==========================

        title = ctk.CTkLabel(
            self,
            text="🧪 Test Case Management System",
            font=("Segoe UI", 28, "bold")
        )

        title.pack(pady=20)

        # ==========================
        # Ana Alan
        # ==========================

        content = ctk.CTkFrame(self)

        content.pack(
            fill="both",
            expand=True,
            padx=20,
            pady=10
        )

        # Sol panel

        self.left_panel = ctk.CTkFrame(
            content,
            width=320,
            corner_radius=15
        )

        self.left_panel.pack(
            side="left",
            fill="y",
            padx=(0, 10)
        )

        # Sağ panel (Scrollable)

        self.right_panel = ctk.CTkScrollableFrame(
            content,
            corner_radius=15
        )

        self.right_panel.pack(
            side="right",
            fill="both",
            expand=True
        )

        search = ctk.CTkEntry(
            self.left_panel,
            placeholder_text="🔍 Test Ara..."
        )

        search.pack(
            fill="x",
            padx=15,
            pady=(20,10)
        )

        title = ctk.CTkLabel(
            self.left_panel,
            text="Test Cases",
            font=("Segoe UI",18,"bold")
        )

        title.pack(pady=5)

        # Test listesinin bulunduğu alan
        self.test_list = ctk.CTkScrollableFrame(
            self.left_panel,
            fg_color="transparent"
        )

        self.test_list.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10
        )

        self.load_tests()

        # Yeni Test Butonu
        new_btn = ctk.CTkButton(
            self.left_panel,
            text="➕ Yeni Test",
            height=40,
            command=self.open_new_test
        )

        new_btn.pack(
            fill="x",
            padx=10,
            pady=(5,5)
        )

        # Excel Butonu
        export_btn = ctk.CTkButton(
            self.left_panel,
            text="📤 Excel'e Aktar",
            height=40,
            fg_color="#2E8B57",
            hover_color="#246B45",
            command=self.export_excel
        )

        export_btn.pack(
            fill="x",
            padx=10,
            pady=(0,10)
        )

        # Sağ Panel
        title = ctk.CTkLabel(
            self.right_panel,
            text="Test Detayı",
            font=("Segoe UI",24,"bold")
        )

        title.pack(
            anchor="nw",
            padx=20,
            pady=(20,10)
        )

        info = ctk.CTkLabel(
            self.right_panel,
            text="Soldan bir test seçerek detaylarını görüntüleyebilirsiniz.",
            font=("Segoe UI",16)
        )

        info.pack(
            anchor="nw",
            padx=20
        )

    def clear_right_panel(self):

        for widget in self.right_panel.winfo_children():
            widget.destroy()

    def open_new_test(self):

        NewTestWindow(self, self.load_tests)

    def show_test(self, tc_id):

        self.clear_right_panel()

        test = self.db.get_test_case(tc_id)

        pre_conditions = self.db.get_pre_conditions(tc_id)

        test_data = self.db.get_test_data(tc_id)

        steps = self.db.get_test_steps(tc_id)

        post_conditions = self.db.get_post_conditions(tc_id)

        if test is None:
            return

        # ==========================
        # Başlık
        # ==========================

        header = ctk.CTkFrame(
            self.right_panel,
            fg_color="transparent"
        )

        header.pack(fill="x", padx=20, pady=20)

        title = ctk.CTkLabel(
            header,
            text="Test Detayı",
            font=("Segoe UI",24,"bold")
        )

        title.pack(side="left")

        edit_btn = ctk.CTkButton(
            header,
            text="✏ Düzenle",
            width=100,
            command=lambda: self.edit_test(tc_id)
        )

        edit_btn.pack(side="right", padx=5)

        delete_btn = ctk.CTkButton(
            header,
            text="🗑 Sil",
            width=100,
            fg_color="#C0392B",
            hover_color="#922B21",
            command=lambda: self.delete_test(tc_id)
        )
        

        delete_btn.pack(side="right", padx=5)

        labels = [
            ("TC ID", test[0]),
            ("Test Adı", test[1]),
            ("Priority", test[2]),
            ("Application", test[3]),
            ("Version", test[4]),
            ("Creator", test[5]),
            ("Create Date", test[6]),
            ("Automation", "Evet" if test[7] else "Hayır")
        ]

        for key, value in labels:

            ctk.CTkLabel(
                self.right_panel,
                text=key,
                font=("Segoe UI",16,"bold")
            ).pack(anchor="w", padx=20, pady=(10,0))

            ctk.CTkLabel(
                self.right_panel,
                text=str(value),
                font=("Segoe UI",15)
            ).pack(anchor="w", padx=40)

        # ==========================
        # Pre Conditions
        # ==========================

        ctk.CTkLabel(
            self.right_panel,
            text="Pre Conditions",
            font=("Segoe UI",18,"bold")
        ).pack(anchor="w", padx=20, pady=(20,5))

        for condition in pre_conditions:

            ctk.CTkLabel(
                self.right_panel,
                text="• " + condition[0],
                font=("Segoe UI",15)
            ).pack(anchor="w", padx=40)

        # ==========================
        # Test Data
        # ==========================

        ctk.CTkLabel(
            self.right_panel,
            text="Test Data",
            font=("Segoe UI",18,"bold")
        ).pack(anchor="w", padx=20, pady=(20,5))

        for data in test_data:

            ctk.CTkLabel(
                self.right_panel,
                text=f"{data[0]} : {data[1]}",
                font=("Segoe UI",15)
            ).pack(anchor="w", padx=40)

        # ==========================
        # Test Steps
        # ==========================

        ctk.CTkLabel(
            self.right_panel,
            text="Test Steps",
            font=("Segoe UI",18,"bold")
        ).pack(anchor="w", padx=20, pady=(20,5))

        for step in steps:

            frame = ctk.CTkFrame(self.right_panel)
            frame.pack(fill="x", padx=20, pady=5)

            ctk.CTkLabel(
                frame,
                text=f"Step {step[0]}",
                font=("Segoe UI",16,"bold")
            ).pack(anchor="w", padx=10, pady=5)

            ctk.CTkLabel(
                frame,
                text=f"Action : {step[1]}"
            ).pack(anchor="w", padx=20)

            ctk.CTkLabel(
                frame,
                text=f"Test Data : {step[2]}"
            ).pack(anchor="w", padx=20)

            ctk.CTkLabel(
                frame,
                text=f"Expected : {step[3]}"
            ).pack(anchor="w", padx=20, pady=(0,10))

        # ==========================
        # Post Conditions
        # ==========================

        ctk.CTkLabel(
            self.right_panel,
            text="Post Conditions",
            font=("Segoe UI",18,"bold")
        ).pack(anchor="w", padx=20, pady=(20,5))

        for condition in post_conditions:

            ctk.CTkLabel(
                self.right_panel,
                text="• " + condition[0],
                font=("Segoe UI",15)
            ).pack(anchor="w", padx=40)


    def load_tests(self):

        # Listeyi temizle
        for widget in self.test_list.winfo_children():
            widget.destroy()

        tests = self.db.get_all_test_cases()

        for test in tests:

            btn = ctk.CTkButton(
                self.test_list,
                text=test[0],
                height=35,
                anchor="w",
                command=lambda tc=test[0]: self.show_test(tc)
            )

            btn.pack(
                fill="x",
                pady=3
            )

    def delete_test(self, tc_id):

        answer = messagebox.askyesno(
            "Sil",
            f"{tc_id} silinsin mi?"
        )

        if not answer:
            return

        self.db.delete_test_case(tc_id)

        self.load_tests()

        self.clear_right_panel()

        ctk.CTkLabel(
            self.right_panel,
            text="Test silindi.",
            font=("Segoe UI",18)
        ).pack(pady=30)

    

    def edit_test(self, tc_id):

        NewTestWindow(
            self,
            self.load_tests,
            tc_id
        )

    def export_excel(self):

        try:
            exporter = ExcelExporter()
            exporter.export()

            messagebox.showinfo(
                "Başarılı",
                "Excel başarıyla oluşturuldu."
            )

        except Exception as e:

            messagebox.showerror(
                "Hata",
                str(e)
            )
