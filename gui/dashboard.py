import customtkinter as ctk
from gui.new_test import NewTestWindow
from tkinter import messagebox
from excel.exporter import ExcelExporter


class Dashboard(ctk.CTkFrame):

    def __init__(
        self,
        master,
        db,
        project_name,
        go_back,
        go_forward
    ):

        super().__init__(master)

        self.db = db
        self.project_name = project_name

        self.go_back = go_back
        self.go_forward = go_forward

        self.active_test_window = None

        self.configure(
            fg_color="#1E1E1E"
        )

        self.create_layout()

    # =====================================================
    # STATUS COLOR
    # =====================================================

    def get_status_color(self, status):

        colors = {
            "Başarılı": "#2ECC71",
            "Başarısız": "#E74C3C",
            "Beklemede": "#F1C40F",
            "Devam Ediyor": "#E67E22",
            "Yeni": "#3498DB",
            "Çalıştırılmadı": "#3498DB",
            "Engellendi": "#9B59B6"
        }

        return colors.get(
            status,
            "#95A5A6"
        )

    # =====================================================
    # CREATE LAYOUT
    # =====================================================

    def create_layout(self):

        # =================================================
        # ÜST NAVİGASYON
        # =================================================

        top_bar = ctk.CTkFrame(
            self,
            fg_color="transparent"
        )

        top_bar.pack(
            fill="x",
            padx=20,
            pady=(15, 0)
        )

        # -------------------------------------------------
        # Geri
        # -------------------------------------------------

        back_btn = ctk.CTkButton(
            top_bar,
            text="← Geri",
            width=100,
            height=35,
            command=self.go_back
        )

        back_btn.pack(
            side="left",
            padx=(0, 5)
        )

        # -------------------------------------------------
        # İleri
        # -------------------------------------------------

        forward_btn = ctk.CTkButton(
            top_bar,
            text="İleri →",
            width=100,
            height=35,
            command=self.go_forward
        )

        forward_btn.pack(
            side="left",
            padx=5
        )

        # -------------------------------------------------
        # Proje adı
        # -------------------------------------------------

        project_label = ctk.CTkLabel(
            top_bar,
            text=f"📁 {self.project_name}",
            font=("Segoe UI", 16, "bold"),
            text_color="#5DADE2"
        )

        project_label.pack(
            side="right",
            padx=10
        )

        # =================================================
        # BAŞLIK
        # =================================================

        title = ctk.CTkLabel(
            self,
            text="🧪 Test Case Yönetim Sistemi",
            font=("Segoe UI", 28, "bold")
        )

        title.pack(
            pady=20
        )

        # =================================================
        # ANA ALAN
        # =================================================

        content = ctk.CTkFrame(
            self
        )

        content.pack(
            fill="both",
            expand=True,
            padx=20,
            pady=10
        )

        # =================================================
        # SOL PANEL
        # =================================================

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

        self.left_panel.pack_propagate(False)

        # =================================================
        # SAĞ PANEL
        # =================================================

        self.right_panel = ctk.CTkScrollableFrame(
            content,
            corner_radius=15
        )

        self.right_panel.pack(
            side="right",
            fill="both",
            expand=True
        )

        # ==========================
        # Arama
        # ==========================

        self.search_entry = ctk.CTkEntry(
            self.left_panel,
            placeholder_text="🔍 Test Ara..."
        )

        self.search_entry.pack(
            fill="x",
            padx=15,
            pady=(20, 10)
        )

        self.search_entry.bind(
            "<KeyRelease>",
            lambda event: self.load_tests()
        )

        # ==========================
        # TEST FİLTRESİ
        # ==========================

        filter_title = ctk.CTkLabel(
            self.left_panel,
            text="🔽 Test Filtresi",
            font=("Segoe UI", 15, "bold")
        )

        filter_title.pack(
            anchor="w",
            padx=15,
            pady=(5, 5)
        )

        # -------------------------------------------------
        # FİLTRELER
        # -------------------------------------------------

        self.filter_frame = ctk.CTkFrame(
            self.left_panel,
            fg_color="transparent"
        )

        self.filter_frame.pack(
            fill="x",
            padx=15,
            pady=(0, 10)
        )

        # -------------------------------------------------
        # TEST DURUMU
        # -------------------------------------------------

        self.status_filter = ctk.CTkComboBox(
            self.filter_frame,
            values=[
                "Tümü",
                "Yeni",
                "Beklemede",
                "Devam Ediyor",
                "Başarılı",
                "Başarısız",
                "Engellendi",
                "Çalıştırılmadı"
            ],
            command=lambda value: self.load_tests()
        )

        self.status_filter.pack(
            fill="x",
            pady=3
        )

        self.status_filter.set("Tümü")

        # -------------------------------------------------
        # OTOMASYON TALEBİ
        # -------------------------------------------------

        self.automation_requested_filter = ctk.CTkComboBox(
            self.filter_frame,
            values=[
                "Otomasyon Talebi: Tümü",
                "Otomasyonlaştırılsın",
                "Otomasyonlaştırılmasın"
            ],
            command=lambda value: self.load_tests()
        )

        self.automation_requested_filter.pack(
            fill="x",
            pady=3
        )

        self.automation_requested_filter.set(
            "Otomasyon Talebi: Tümü"
        )

        # -------------------------------------------------
        # OTOMASYON DURUMU
        # -------------------------------------------------

        self.automation_completed_filter = ctk.CTkComboBox(
            self.filter_frame,
            values=[
                "Otomasyon Durumu: Tümü",
                "Otomasyonlaştırıldı",
                "Otomasyonlaştırılmadı"
            ],
            command=lambda value: self.load_tests()
        )

        self.automation_completed_filter.pack(
            fill="x",
            pady=3
        )

        self.automation_completed_filter.set(
            "Otomasyon Durumu: Tümü"
        )

        # -------------------------------------------------
        # SONUÇ SAYISI
        # -------------------------------------------------

        self.result_label = ctk.CTkLabel(
            self.left_panel,
            text="",
            text_color="#AAAAAA",
            font=("Segoe UI", 12)
        )

        self.result_label.pack(
            pady=(0, 5)
        )

        # ==========================
        # TEST LİSTESİ BAŞLIĞI
        # ==========================

        title = ctk.CTkLabel(
            self.left_panel,
            text="Test Listesi",
            font=("Segoe UI", 18, "bold")
        )

        title.pack(
            pady=5
        )

        # =================================================
        # TEST LİSTESİ
        # =================================================

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

        # =================================================
        # YENİ TEST
        # =================================================

        new_btn = ctk.CTkButton(
            self.left_panel,
            text="➕ Yeni Test",
            height=40,
            command=self.open_new_test
        )

        new_btn.pack(
            fill="x",
            padx=10,
            pady=(5, 5)
        )

        # =================================================
        # EXCEL
        # =================================================

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
            pady=(0, 10)
        )

        # =================================================
        # SAĞ PANEL BAŞLANGIÇ
        # =================================================

        title = ctk.CTkLabel(
            self.right_panel,
            text="Test Detayı",
            font=("Segoe UI", 24, "bold")
        )

        title.pack(
            anchor="nw",
            padx=20,
            pady=(20, 10)
        )

        info = ctk.CTkLabel(
            self.right_panel,
            text="Soldan bir test seçerek detaylarını görüntüleyebilirsiniz.",
            font=("Segoe UI", 16)
        )

        info.pack(
            anchor="nw",
            padx=20
        )

        # =================================================
        # TESTLERİ YÜKLE
        # =================================================

        self.load_tests()

    # =====================================================
    # CLEAR RIGHT PANEL
    # =====================================================

    def clear_right_panel(self):

        for widget in self.right_panel.winfo_children():
            widget.destroy()

    # =====================================================
    # OPEN NEW TEST
    # =====================================================

    def open_new_test(self):

        if (
            self.active_test_window is not None
            and self.active_test_window.winfo_exists()
        ):

            self.active_test_window.focus()

            return

        self.active_test_window = NewTestWindow(
            self,
            self.load_tests,
            db=self.db
        )

    # =====================================================
    # SHOW TEST
    # =====================================================

    def show_test(self, tc_id):

        self.clear_right_panel()

        test = self.db.get_test_case(tc_id)

        if test is None:
            return

        pre_conditions = self.db.get_pre_conditions(
            tc_id
        )

        test_data = self.db.get_test_data(
            tc_id
        )

        steps = self.db.get_test_steps(
            tc_id
        )

        post_conditions = self.db.get_post_conditions(
            tc_id
        )

        # =================================================
        # HEADER
        # =================================================

        header = ctk.CTkFrame(
            self.right_panel,
            fg_color="transparent"
        )

        header.pack(
            fill="x",
            padx=20,
            pady=20
        )

        title = ctk.CTkLabel(
            header,
            text="Test Detayı",
            font=("Segoe UI", 24, "bold")
        )

        title.pack(
            side="left"
        )

        # =================================================
        # EDIT BUTTON
        # =================================================

        edit_btn = ctk.CTkButton(
            header,
            text="✏ Düzenle",
            width=100,
            command=lambda: self.edit_test(tc_id)
        )

        edit_btn.pack(
            side="right",
            padx=5
        )

        # =================================================
        # DELETE BUTTON
        # =================================================

        delete_btn = ctk.CTkButton(
            header,
            text="🗑 Sil",
            width=100,
            fg_color="#C0392B",
            hover_color="#922B21",
            command=lambda: self.delete_test(tc_id)
        )

        delete_btn.pack(
            side="right",
            padx=5
        )

        # =================================================
        # STATUS
        # =================================================

        status = (
            test[8]
            if len(test) > 8 and test[8]
            else "Yeni"
        )

        status_color = self.get_status_color(
            status
        )

        status_frame = ctk.CTkFrame(
            self.right_panel,
            fg_color="#252525",
            corner_radius=10
        )

        status_frame.pack(
            fill="x",
            padx=20,
            pady=(0, 15)
        )

        ctk.CTkLabel(
            status_frame,
            text="Durum",
            font=("Segoe UI", 16, "bold")
        ).pack(
            side="left",
            padx=(15, 10),
            pady=12
        )

        ctk.CTkLabel(
            status_frame,
            text="●",
            text_color=status_color,
            font=("Segoe UI", 22, "bold")
        ).pack(
            side="left"
        )

        ctk.CTkLabel(
            status_frame,
            text=status,
            text_color=status_color,
            font=("Segoe UI", 16, "bold")
        ).pack(
            side="left",
            padx=(5, 15)
        )

        # =================================================
        # TEMEL TEST BİLGİLERİ
        # =================================================

        basic_labels = [

            ("TC ID", test[0]),
            ("Test Adı", test[1]),
            ("Öncelik", test[2]),
            ("Uygulama", test[3]),
            ("Versiyon", test[4]),
            ("Oluşturan", test[5]),
            ("Oluşturma Tarihi", test[6]),

        ]

        for key, value in basic_labels:

            self.add_detail_row(
                key,
                value
            )

        # =================================================
        # TEST TÜRÜ
        # =================================================

        self.add_detail_row(
            "Test Türü",
            test[12]
            if len(test) > 12 and test[12]
            else "-"
        )

        # =================================================
        # TEST ORTAMI
        # =================================================

        self.add_detail_row(
            "Test Ortamı",
            test[13]
            if len(test) > 13 and test[13]
            else "-"
        )

        # =================================================
        # OTOMASYON BİLGİLERİ
        # =================================================

        self.add_section_title(
            "Otomasyon Bilgileri"
        )

        automation_requested = (
            "Evet"
            if len(test) > 9 and test[9]
            else "Hayır"
        )

        automation_completed = (
            "Evet"
            if len(test) > 10 and test[10]
            else "Hayır"
        )

        automation_scenario = (
            test[11]
            if len(test) > 11 and test[11]
            else "-"
        )

        self.add_detail_row(
            "Otomasyonlaştırılsın mı?",
            automation_requested
        )

        self.add_detail_row(
            "Otomasyonlaştırıldı mı?",
            automation_completed
        )

        self.add_detail_row(
            "Otomasyon Senaryo Karşılığı",
            automation_scenario
        )

        # =================================================
        # HATA BİLGİLERİ
        # =================================================

        self.add_section_title(
            "Hata Bilgileri"
        )

        error_code = (
            test[14]
            if len(test) > 14 and test[14]
            else "-"
        )

        error_priority = (
            test[15]
            if len(test) > 15 and test[15]
            else "Yok"
        )

        self.add_detail_row(
            "Hata Kodu",
            error_code
        )

        self.add_detail_row(
            "Hata Önceliği",
            error_priority
        )

        # =================================================
        # PRE CONDITIONS
        # =================================================

        self.add_section_title(
            "Ön Koşullar"
        )

        if pre_conditions:

            for condition in pre_conditions:

                ctk.CTkLabel(
                    self.right_panel,
                    text="• " + str(condition[0]),
                    font=("Segoe UI", 15)
                ).pack(
                    anchor="w",
                    padx=40,
                    pady=2
                )

        else:

            ctk.CTkLabel(
                self.right_panel,
                text="- Ön koşul bulunmuyor.",
                text_color="#999999",
                font=("Segoe UI", 14)
            ).pack(
                anchor="w",
                padx=40
            )

        # =================================================
        # TEST DATA
        # =================================================

        self.add_section_title(
            "Test Verileri"
        )

        if test_data:

            for data in test_data:

                ctk.CTkLabel(
                    self.right_panel,
                    text=f"{data[0]} : {data[1]}",
                    font=("Segoe UI", 15)
                ).pack(
                    anchor="w",
                    padx=40,
                    pady=2
                )

        else:

            ctk.CTkLabel(
                self.right_panel,
                text="- Test verisi bulunmuyor.",
                text_color="#999999",
                font=("Segoe UI", 14)
            ).pack(
                anchor="w",
                padx=40
            )

        # =================================================
        # TEST STEPS
        # =================================================

        self.add_section_title(
            "Test Adımları"
        )

        if steps:

            for step in steps:

                frame = ctk.CTkFrame(
                    self.right_panel
                )

                frame.pack(
                    fill="x",
                    padx=20,
                    pady=5
                )

                ctk.CTkLabel(
                    frame,
                    text=f"Adım {step[0]}",
                    font=("Segoe UI", 16, "bold")
                ).pack(
                    anchor="w",
                    padx=10,
                    pady=5
                )

                ctk.CTkLabel(
                    frame,
                    text=f"İşlem : {step[1]}",
                    font=("Segoe UI", 14)
                ).pack(
                    anchor="w",
                    padx=20
                )

                ctk.CTkLabel(
                    frame,
                    text=f"Test Verisi : {step[2]}",
                    font=("Segoe UI", 14)
                ).pack(
                    anchor="w",
                    padx=20
                )

                ctk.CTkLabel(
                    frame,
                    text=f"Beklenen Sonuç : {step[3]}",
                    font=("Segoe UI", 14)
                ).pack(
                    anchor="w",
                    padx=20,
                    pady=(0, 10)
                )

        else:

            ctk.CTkLabel(
                self.right_panel,
                text="- Test adımı bulunmuyor.",
                text_color="#999999",
                font=("Segoe UI", 14)
            ).pack(
                anchor="w",
                padx=40
            )

        # =================================================
        # POST CONDITIONS
        # =================================================

        self.add_section_title(
            "Son Koşullar"
        )

        if post_conditions:

            for condition in post_conditions:

                ctk.CTkLabel(
                    self.right_panel,
                    text="• " + str(condition[0]),
                    font=("Segoe UI", 15)
                ).pack(
                    anchor="w",
                    padx=40,
                    pady=2
                )

        else:

            ctk.CTkLabel(
                self.right_panel,
                text="- Son koşul bulunmuyor.",
                text_color="#999999",
                font=("Segoe UI", 14)
            ).pack(
                anchor="w",
                padx=40
            )

    # =====================================================
    # DETAIL ROW
    # =====================================================

    def add_detail_row(
        self,
        key,
        value
    ):

        ctk.CTkLabel(
            self.right_panel,
            text=key,
            font=("Segoe UI", 16, "bold")
        ).pack(
            anchor="w",
            padx=20,
            pady=(10, 0)
        )

        ctk.CTkLabel(
            self.right_panel,
            text=str(value),
            font=("Segoe UI", 15)
        ).pack(
            anchor="w",
            padx=40
        )

    # =====================================================
    # SECTION TITLE
    # =====================================================

    def add_section_title(
        self,
        title
    ):

        ctk.CTkLabel(
            self.right_panel,
            text=title,
            font=("Segoe UI", 18, "bold")
        ).pack(
            anchor="w",
            padx=20,
            pady=(20, 5)
        )

    # =====================================================
    # TEST FILTER
    # =====================================================

    def filter_tests(self, tests):

        selected_filter = self.filter_var.get()

        # -------------------------------------------------
        # TÜM TESTLER
        # -------------------------------------------------

        if selected_filter == "Tüm Testler":
            return tests

        # -------------------------------------------------
        # OTOMASYONLAŞTIRILABİLİR
        # -------------------------------------------------

        if selected_filter == "Otomasyonlaştırılabilir":

            return [
                test
                for test in tests
                if len(test) > 9
                and test[9] == 1
            ]

        # -------------------------------------------------
        # OTOMASYONLAŞTIRILDI
        # -------------------------------------------------

        if selected_filter == "Otomasyonlaştırıldı":

            return [
                test
                for test in tests
                if len(test) > 10
                and test[10] == 1
            ]

        # -------------------------------------------------
        # OTOMASYON BEKLİYOR
        # -------------------------------------------------

        if selected_filter == "Otomasyon Bekliyor":

            return [
                test
                for test in tests
                if (
                    len(test) > 10
                    and test[9] == 1
                    and test[10] == 0
                )
            ]

        # -------------------------------------------------
        # MANUEL TESTLER
        # -------------------------------------------------

        if selected_filter == "Manuel Testler":

            return [
                test
                for test in tests
                if (
                    len(test) > 9
                    and test[9] == 0
                )
            ]

        # -------------------------------------------------
        # DURUM FİLTRELERİ
        # -------------------------------------------------

        status_filters = [
            "Başarılı",
            "Başarısız",
            "Beklemede",
            "Devam Ediyor",
            "Engellendi"
        ]

        if selected_filter in status_filters:

            return [
                test
                for test in tests
                if (
                    len(test) > 8
                    and test[8] == selected_filter
                )
            ]

        return tests

    # =====================================================
    # LOAD TESTS
    # =====================================================

    def load_tests(self):

        # =================================================
        # LİSTEYİ TEMİZLE
        # =================================================

        for widget in self.test_list.winfo_children():
            widget.destroy()

        # =================================================
        # TÜM TESTLER
        # =================================================

        tests = self.db.get_all_test_cases()

        # =================================================
        # ARAMA
        # =================================================

        query = (
            self.search_entry
            .get()
            .strip()
            .lower()
        )

        if query:

            tests = [
                test
                for test in tests
                if (
                    query in str(test[0]).lower()
                    or
                    query in str(test[1]).lower()
                )
            ]

        # =================================================
        # DURUM FİLTRESİ
        # =================================================

        status_filter = self.status_filter.get()

        if status_filter != "Tümü":

            tests = [
                test
                for test in tests
                if (
                    test[8]
                    if len(test) > 8 and test[8]
                    else "Yeni"
                ) == status_filter
            ]

        # =================================================
        # OTOMASYON TALEBİ
        # =================================================

        automation_requested_filter = (
            self.automation_requested_filter.get()
        )

        if automation_requested_filter == "Otomasyonlaştırılsın":

            tests = [
                test
                for test in tests
                if len(test) > 9 and test[9] == 1
            ]

        elif automation_requested_filter == "Otomasyonlaştırılmasın":

            tests = [
                test
                for test in tests
                if not (len(test) > 9 and test[9] == 1)
            ]

        # =================================================
        # OTOMASYON DURUMU
        # =================================================

        automation_completed_filter = (
            self.automation_completed_filter.get()
        )

        if automation_completed_filter == "Otomasyonlaştırıldı":

            tests = [
                test
                for test in tests
                if len(test) > 10 and test[10] == 1
            ]

        elif automation_completed_filter == "Otomasyonlaştırılmadı":

            tests = [
                test
                for test in tests
                if not (len(test) > 10 and test[10] == 1)
            ]

        # =================================================
        # SONUÇ SAYISI
        # =================================================

        self.result_label.configure(
            text=f"{len(tests)} test bulundu"
        )

        # =================================================
        # TESTLER
        # =================================================

        for test in tests:

            tc_id = test[0]

            # -------------------------------------------------
            # STATUS
            # -------------------------------------------------

            status = (
                test[8]
                if len(test) > 8 and test[8]
                else "Yeni"
            )

            color = self.get_status_color(
                status
            )

            # -------------------------------------------------
            # TEST SATIRI
            # -------------------------------------------------

            row = ctk.CTkFrame(
                self.test_list,
                fg_color="transparent"
            )

            row.pack(
                fill="x",
                pady=3
            )

            # -------------------------------------------------
            # DURUM NOKTASI
            # -------------------------------------------------

            status_dot = ctk.CTkLabel(
                row,
                text="●",
                text_color=color,
                font=("Segoe UI", 18, "bold"),
                width=25
            )

            status_dot.pack(
                side="left",
                padx=(3, 2)
            )

            # -------------------------------------------------
            # TEST BUTONU
            # -------------------------------------------------

            btn = ctk.CTkButton(
                row,
                text=tc_id,
                height=35,
                anchor="w",
                command=lambda tc=tc_id:
                    self.show_test(tc)
            )

            btn.pack(
                side="left",
                fill="x",
                expand=True
            )

    def delete_test(
        self,
        tc_id
    ):

        answer = messagebox.askyesno(
            "Sil",
            f"{tc_id} silinsin mi?"
        )

        if not answer:
            return

        self.db.delete_test_case(
            tc_id
        )

        self.load_tests()

        self.clear_right_panel()

        ctk.CTkLabel(
            self.right_panel,
            text="Test silindi.",
            font=("Segoe UI", 18)
        ).pack(
            pady=30
        )

    # =====================================================
    # EDIT TEST
    # =====================================================

    def edit_test(
        self,
        tc_id
    ):

        if (
            self.active_test_window is not None
            and self.active_test_window.winfo_exists()
        ):

            self.active_test_window.focus()

            return

        def refresh():

            self.load_tests()

            self.show_test(
                tc_id
            )

        self.active_test_window = NewTestWindow(
            self,
            refresh,
            tc_id,
            db=self.db
        )

    # =====================================================
    # EXPORT EXCEL
    # =====================================================

    def export_excel(self):

        try:

            exporter = ExcelExporter(
                self.db
            )

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