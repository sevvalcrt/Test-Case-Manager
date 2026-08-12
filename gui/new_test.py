from tkinter import messagebox
import sqlite3
import customtkinter as ctk
from datetime import datetime
import tkinter as tk

from database.database import Database
from excel.exporter import ExcelExporter


class NewTestWindow(ctk.CTkToplevel):

    def __init__(self, master, refresh_callback, tc_id=None, db=None):

        self.refresh_callback = refresh_callback
        self.edit_mode = tc_id is not None
        self.edit_tc_id = tc_id
        self.loading_data = False

        super().__init__(master)

        if db is not None:
            self.db = db
            self.owns_db = False
        else:
            self.db = Database()
            self.owns_db = True

        self.title(
            "Test Düzenle" if self.edit_mode else "Yeni Test Senaryosu"
        )

        self.geometry("850x800")
        self.resizable(True, True)

        # NOT: Bu pencere kasıtlı olarak grab_set() ile modal
        # yapılmıyor. grab_set(), pencerenin küçültme (minimize)
        # tuşunu tamamen tepkisiz bırakıyordu (hem bu pencerede
        # hem de ana pencerede). Modal davranıştan (ana pencereye
        # tıklanamaması) vazgeçerek küçültme sorunu çözülüyor.

        # =====================================================
        # BAŞLIK
        # =====================================================

        title = ctk.CTkLabel(
            self,
            text="Test Düzenle" if self.edit_mode else "Yeni Test Senaryosu",
            font=("Segoe UI", 24, "bold")
        )

        title.pack(pady=20)

        # =====================================================
        # SCROLLABLE FORM
        # =====================================================

        form = ctk.CTkScrollableFrame(
            self,
            width=700,
        )

        form.pack(
            fill="both",
            expand=True,
            padx=20,
            pady=(5, 5)
        )

        # =====================================================
        # SCROLL KONTROLLERİ
        # =====================================================

        self.form = form

        self.bind(
            "<MouseWheel>",
            self._on_mousewheel
        )

        self.bind(
            "<Button-4>",
            self._on_mousewheel_linux
        )

        self.bind(
            "<Button-5>",
            self._on_mousewheel_linux
        )

        self.bind(
            "<Prior>",
            self._scroll_page_up
        )

        self.bind(
            "<Next>",
            self._scroll_page_down
        )

        self.bind(
            "<Home>",
            self._scroll_home
        )

        self.bind(
            "<End>",
            self._scroll_end
        )

        # =====================================================
        # TC ID
        # =====================================================

        ctk.CTkLabel(
            form,
            text="TC ID"
        ).pack(anchor="w", pady=(10, 0))

        self.tc_id = ctk.CTkEntry(form)
        self.tc_id.pack(fill="x")

        # =====================================================
        # TEST ADI
        # =====================================================

        ctk.CTkLabel(
            form,
            text="Test Adı"
        ).pack(anchor="w", pady=(10, 0))

        self.test_name = ctk.CTkEntry(form)
        self.test_name.pack(fill="x")

        # =====================================================
        # APPLICATION
        # =====================================================

        ctk.CTkLabel(
            form,
            text="Uygulama"
        ).pack(anchor="w", pady=(10, 0))

        self.application = ctk.CTkEntry(form)
        self.application.pack(fill="x")

        # =====================================================
        # VERSION
        # =====================================================

        ctk.CTkLabel(
            form,
            text="Sürüm"
        ).pack(anchor="w", pady=(10, 0))

        self.version = ctk.CTkEntry(form)

        self.version.insert(
            0,
            "1"
        )

        self.version.pack(fill="x")

        # =====================================================
        # CREATOR
        # =====================================================

        ctk.CTkLabel(
            form,
            text="Oluşturan"
        ).pack(anchor="w", pady=(10, 0))

        self.creator = ctk.CTkEntry(form)
        self.creator.pack(fill="x")

        # =====================================================
        # CREATE DATE
        # =====================================================

        ctk.CTkLabel(
            form,
            text="Oluşturma Tarihi"
        ).pack(anchor="w", pady=(10, 0))

        self.create_date = ctk.CTkEntry(form)

        self.create_date.insert(
            0,
            datetime.now().strftime("%d.%m.%Y")
        )

        self.create_date.pack(fill="x")

        # =====================================================
        # PRIORITY
        # =====================================================

        ctk.CTkLabel(
            form,
            text="Öncelik"
        ).pack(anchor="w", pady=(10, 0))

        self.priority = ctk.CTkOptionMenu(
            form,
            values=[
                "Düşük",
                "Orta",
                "Yüksek"
            ]
        )

        self.priority.set("Orta")

        self.priority.pack(fill="x")

        # =====================================================
        # STATUS
        # =====================================================

        ctk.CTkLabel(
            form,
            text="Durum"
        ).pack(anchor="w", pady=(10, 0))

        self.status = ctk.CTkOptionMenu(
            form,
            values=[
                "Beklemede",
                "Başarılı",
                "Başarısız",
                "Engellendi"
            ]
        )

        self.status.set("Beklemede")

        self.status.pack(fill="x")

        # =====================================================
        # TEST TÜRÜ
        # =====================================================

        ctk.CTkLabel(
            form,
            text="Test Türü"
        ).pack(anchor="w", pady=(10, 0))

        self.test_type = ctk.CTkOptionMenu(
            form,
            values=[
                "Fonksiyonel",
                "Smoke",
                "Regression",
                "Integration",
                "UI",
                "Performance",
                "Security",
                "Usability"
            ]
        )

        self.test_type.set("Fonksiyonel")

        self.test_type.pack(fill="x")


        # =====================================================
        # TEST ORTAMI
        # =====================================================

        ctk.CTkLabel(
            form,
            text="Test Ortamı"
        ).pack(anchor="w", pady=(10, 0))

        self.test_environment = ctk.CTkEntry(
            form,
            placeholder_text="Örn: https://www.mango.com"
        )

        self.test_environment.pack(fill="x")

        # =====================================================
        # OTOMASYON BİLGİLERİ
        # =====================================================

        self.create_section_title(
            form,
            "Otomasyon Bilgileri"
        )


        # -----------------------------------------------------
        # OTOMASYONA UYGUN MU?
        # -----------------------------------------------------

        self.automation_requested = ctk.CTkCheckBox(
            form,
            text="Otomasyonlaştırılsın mı?"
        )

        self.automation_requested.pack(
            anchor="w",
            pady=(10, 5)
        )


        # -----------------------------------------------------
        # OTOMASYON TAMAMLANDI MI?
        # -----------------------------------------------------

        self.automation_completed = ctk.CTkCheckBox(
            form,
            text="Otomasyonlaştırıldı mı?"
        )

        self.automation_completed.pack(
            anchor="w",
            pady=5
        )


        # -----------------------------------------------------
        # OTOMASYON SENARYO KARŞILIĞI
        # -----------------------------------------------------

        ctk.CTkLabel(
            form,
            text="Otomasyon Senaryo Karşılığı"
        ).pack(
            anchor="w",
            pady=(10, 0)
        )

        self.automation_scenario = ctk.CTkEntry(
            form,
            placeholder_text="Örn: F01_TC_0007"
        )

        self.automation_scenario.pack(
            fill="x"
        )

        # =====================================================
        # HATA BİLGİLERİ
        # =====================================================

        self.create_section_title(
            form,
            "Hata Bilgileri"
        )


        # -----------------------------------------------------
        # HATA KODU
        # -----------------------------------------------------

        ctk.CTkLabel(
            form,
            text="Hata Kodu"
        ).pack(
            anchor="w",
            pady=(10, 0)
        )

        self.error_code = ctk.CTkEntry(
            form,
            placeholder_text="Örn: BUG-001"
        )

        self.error_code.pack(
            fill="x"
        )


        # -----------------------------------------------------
        # HATA ÖNCELİĞİ
        # -----------------------------------------------------

        ctk.CTkLabel(
            form,
            text="Hata Önceliği"
        ).pack(
            anchor="w",
            pady=(10, 0)
        )

        self.error_priority = ctk.CTkOptionMenu(
            form,
            values=[
                "Yok",
                "Düşük",
                "Orta",
                "Yüksek",
                "Kritik"
            ]
        )

        self.error_priority.set("Yok")

        self.error_priority.pack(
            fill="x"
        )

        # =====================================================
        # PRECONDITIONS
        # =====================================================

        self.create_section_title(
            form,
            "Ön Koşullar"
        )

        self.pre_frame = ctk.CTkFrame(
            form,
            fg_color="transparent"
        )

        self.pre_frame.pack(
            fill="x",
            pady=5
        )

        self.pre_rows = []

        self.add_precondition_row()

        # =====================================================
        # TEST DATA
        # =====================================================

        self.create_section_title(
            form,
            "Test Verileri"
        )

        data_header = ctk.CTkFrame(
            form,
            fg_color="transparent"
        )

        data_header.pack(
            fill="x",
            pady=(5, 0)
        )

        ctk.CTkLabel(
            data_header,
            text="Veri Anahtarı"
        ).pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 5)
        )

        ctk.CTkLabel(
            data_header,
            text="Veri Değeri"
        ).pack(
            side="left",
            fill="x",
            expand=True,
            padx=(5, 45)
        )

        self.data_frame = ctk.CTkFrame(
            form,
            fg_color="transparent"
        )

        self.data_frame.pack(
            fill="x",
            pady=5
        )

        self.data_rows = []

        self.add_test_data_row()

        # =====================================================
        # TEST STEPS
        # =====================================================

        self.create_section_title(
            form,
            "Test Adımları"
        )

        step_header = ctk.CTkFrame(
            form,
            fg_color="transparent"
        )

        step_header.pack(
            fill="x",
            pady=(5, 0)
        )

        headers = [
            "Adım No",
            "İşlem",
            "Test Verileri",
            "Beklenen Sonuç"
        ]

        for text in headers:

            ctk.CTkLabel(
                step_header,
                text=text
            ).pack(
                side="left",
                fill="x",
                expand=True,
                padx=2
            )

        self.steps_frame = ctk.CTkFrame(
            form,
            fg_color="transparent"
        )

        self.steps_frame.pack(
            fill="x",
            pady=5
        )

        self.step_rows = []

        self.add_test_step_row()

        # =====================================================
        # POSTCONDITIONS
        # =====================================================

        self.create_section_title(
            form,
            "Son Koşullar"
        )

        self.post_frame = ctk.CTkFrame(
            form,
            fg_color="transparent"
        )

        self.post_frame.pack(
            fill="x",
            pady=5
        )

        self.post_rows = []

        self.add_postcondition_row()

        # =====================================================
        # SABİT ALT BUTON BAR
        # =====================================================

        button_frame = ctk.CTkFrame(
            self,
            height=70,
            fg_color="#202020",
            corner_radius=8
        )

        button_frame.pack(
            side="bottom",
            fill="x",
            padx=20,
            pady=(5, 15)
        )

        button_frame.pack_propagate(False)

        # -----------------------------------------------------
        # KAYDET / GÜNCELLE
        # -----------------------------------------------------

        save_btn = ctk.CTkButton(
            button_frame,
            text="💾 Güncelle" if self.edit_mode else "💾 Kaydet",
            command=self.save_test,
            height=60,
            font=("Segoe UI", 16, "bold")
        )

        save_btn.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(12, 6),
            pady=12
        )

        # -----------------------------------------------------
        # İPTAL
        # -----------------------------------------------------

        cancel_btn = ctk.CTkButton(
            button_frame,
            text="İptal",
            fg_color="#666666",
            hover_color="#555555",
            command=self.close_window,
            height=60,
            font=("Segoe UI", 16, "bold")
        )

        cancel_btn.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(6, 12),
            pady=12
        )

        # =====================================================
        # EDIT MODE
        # =====================================================

        if self.edit_mode:
            self.load_test_data()

        # =====================================================
        # ÖNE GETİR
        # =====================================================
        # CTkToplevel, koyu tema başlık çubuğunu uygulamak için
        # açılışta kısa süreliğine gizlenip tekrar gösteriliyor;
        # bu yüzden bazen ana pencerenin arkasında kalabiliyor.
        # lift()/focus_force() ile öne getiriyoruz; after() ile
        # tekrarlamak, Windows'taki gecikmeli render'a karşı
        # güvence sağlıyor.

        self.lift()
        self.focus_force()
        self.after(150, self._bring_to_front)
        self.after(
            200,
            lambda: self.bind_context_menus(self.form)
        )

        # =====================================================
        # SAĞ TIK MENÜLERİ
        # =====================================================

        self.after(
            200,
            lambda: self.bind_context_menus(self.form)
        )

    def _bring_to_front(self):

        try:
            self.lift()
            self.focus_force()
        except Exception:
            pass

    def bind_context_menus(self, parent):

        for widget in parent.winfo_children():

            if isinstance(
                widget,
                (ctk.CTkEntry, ctk.CTkTextbox)
            ):
                widget.bind(
                    "<Button-3>",
                    self.show_context_menu
                )

            if widget.winfo_children():
                self.bind_context_menus(widget)

    def show_context_menu(self, event):

        widget = event.widget

        menu = tk.Menu(
            self,
            tearoff=0
        )

        menu.add_command(
            label="Kes",
            command=lambda: widget.event_generate("<<Cut>>")
        )

        menu.add_command(
            label="Kopyala",
            command=lambda: widget.event_generate("<<Copy>>")
        )

        menu.add_command(
            label="Yapıştır",
            command=lambda: widget.event_generate("<<Paste>>")
        )

        menu.add_separator()

        menu.add_command(
            label="Tümünü Seç",
            command=lambda: widget.event_generate("<<SelectAll>>")
        )

        try:
            menu.tk_popup(
                event.x_root,
                event.y_root
            )
        finally:
            menu.grab_release()

    # =========================================================
    # SCROLL KONTROLLERİ
    # =========================================================

    def _on_mousewheel(self, event):

        try:
            direction = -1 if event.delta > 0 else 1
            self.form._parent_canvas.yview_scroll(
                direction * 21,
                "units"
            )
        except Exception:
            pass

        return "break"


    def _on_mousewheel_linux(self, event):

        try:

            if event.num == 4:

                self.form._parent_canvas.yview_scroll(
                    -3,
                    "units"
                )

            elif event.num == 5:

                self.form._parent_canvas.yview_scroll(
                    3,
                    "units"
                )

        except Exception:
            pass

        return "break"


    def _scroll_page_up(self, event=None):

        try:
            self.form._parent_canvas.yview_scroll(
                -17,
                "units"
            )
        except Exception:
            pass

        return "break"


    def _scroll_page_down(self, event=None):

        try:
            self.form._parent_canvas.yview_scroll(
                17,
                "units"
            )
        except Exception:
            pass

        return "break"


    def _scroll_home(self, event=None):

        try:
            self.form._parent_canvas.yview_moveto(0)
        except Exception:
            pass

        return "break"


    def _scroll_end(self, event=None):

        try:
            self.form._parent_canvas.yview_moveto(1)
        except Exception:
            pass

        return "break"

    # =========================================================
    # SECTION TITLE
    # =========================================================

    def create_section_title(self, parent, text):

        

        frame = ctk.CTkFrame(
            parent,
            height=40,
            fg_color="#1F4E78",
            corner_radius=5
        )

        frame.pack(
            fill="x",
            pady=(15, 5)
        )

        label = ctk.CTkLabel(
            frame,
            text=text,
            text_color="white",
            font=("Segoe UI", 14, "bold")
        )

        label.pack(
            anchor="w",
            padx=12,
            pady=7
        )

    # =========================================================
    # PRECONDITION
    # =========================================================

    def add_precondition_row(self, value=""):

        row = ctk.CTkFrame(
            self.pre_frame,
            fg_color="transparent"
        )

        row.pack(
            fill="x",
            pady=3
        )

        entry = ctk.CTkEntry(row)

        entry.insert(
            0,
            value
        )

        entry.pack(
            side="left",
            fill="x",
            expand=True
        )

        delete_btn = ctk.CTkButton(
            row,
            text="Sil",
            width=50,
            fg_color="#C0392B",
            command=lambda: self.remove_row(
                row,
                self.pre_rows
            )
        )

        delete_btn.pack(
            side="right",
            padx=(5, 0)
        )

        self.pre_rows.append(
            (row, entry)
        )

        # Otomatik satır kontrolü
        entry.bind(
            "<KeyRelease>",
            lambda event: self.check_precondition_row()
        )

    def check_precondition_row(self):

        if self.loading_data:
            return

        if not self.pre_rows:
            self.add_precondition_row()
            return

        last_row, last_entry = self.pre_rows[-1]

        if last_entry.get().strip():

            self.add_precondition_row()

    # =========================================================
    # TEST DATA
    # =========================================================

    def add_test_data_row(
        self,
        data_key="",
        data_value=""
    ):

        row = ctk.CTkFrame(
            self.data_frame,
            fg_color="transparent"
        )

        row.pack(
            fill="x",
            pady=3
        )

        key_entry = ctk.CTkEntry(row)

        key_entry.insert(
            0,
            data_key
        )

        key_entry.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 5)
        )

        value_entry = ctk.CTkEntry(row)

        value_entry.insert(
            0,
            data_value
        )

        value_entry.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(5, 5)
        )

        delete_btn = ctk.CTkButton(
            row,
            text="Sil",
            width=50,
            fg_color="#C0392B",
            command=lambda: self.remove_row(
                row,
                self.data_rows
            )
        )

        delete_btn.pack(
            side="right"
        )

        self.data_rows.append(
            (
                row,
                key_entry,
                value_entry
            )
        )

        key_entry.bind(
            "<KeyRelease>",
            lambda event: self.check_test_data_row()
        )

        value_entry.bind(
            "<KeyRelease>",
            lambda event: self.check_test_data_row()
        )

    def check_test_data_row(self):

        if self.loading_data:
            return

        if not self.data_rows:
            self.add_test_data_row()
            return

        row = self.data_rows[-1]

        key_entry = row[1]
        value_entry = row[2]

        if (
            key_entry.get().strip()
            or value_entry.get().strip()
        ):

            self.add_test_data_row()

    # =========================================================
    # TEST STEP
    # =========================================================

    def add_test_step_row(
        self,
        step_no=None,
        action="",
        test_data="",
        expected_result=""
    ):

        row = ctk.CTkFrame(
            self.steps_frame,
            fg_color="transparent"
        )

        row.pack(
            fill="x",
            pady=3
        )

        if step_no is None:

            step_no = len(self.step_rows) + 1

        step_entry = ctk.CTkEntry(
            row,
            width=70
        )

        step_entry.insert(
            0,
            str(step_no)
        )

        step_entry.pack(
            side="left",
            padx=2
        )

        action_entry = ctk.CTkEntry(row)

        action_entry.insert(
            0,
            action
        )

        action_entry.pack(
            side="left",
            fill="x",
            expand=True,
            padx=2
        )

        data_entry = ctk.CTkEntry(row)

        data_entry.insert(
            0,
            test_data
        )

        data_entry.pack(
            side="left",
            fill="x",
            expand=True,
            padx=2
        )

        expected_entry = ctk.CTkEntry(row)

        expected_entry.insert(
            0,
            expected_result
        )

        expected_entry.pack(
            side="left",
            fill="x",
            expand=True,
            padx=2
        )

        delete_btn = ctk.CTkButton(
            row,
            text="Sil",
            width=50,
            fg_color="#C0392B",
            command=lambda: self.remove_row(
                row,
                self.step_rows
            )
        )

        delete_btn.pack(
            side="right",
            padx=2
        )

        self.step_rows.append(
            (
                row,
                step_entry,
                action_entry,
                data_entry,
                expected_entry
            )
        )

        # Herhangi bir alana yazıldığında kontrol et
        for entry in (
            step_entry,
            action_entry,
            data_entry,
            expected_entry
        ):

            entry.bind(
                "<KeyRelease>",
                lambda event: self.check_test_step_row()
            )

    def check_test_step_row(self):

        if self.loading_data:
            return

        if not self.step_rows:
            self.add_test_step_row()
            return

        last_row = self.step_rows[-1]

        (
            row,
            step_entry,
            action_entry,
            data_entry,
            expected_entry
        ) = last_row

        if (
            action_entry.get().strip()
            or data_entry.get().strip()
            or expected_entry.get().strip()
        ):

            # Sadece Step No yazıldıysa hemen yeni satır açma
            # Action/Test Data/Expected Result girildiğinde aç.
            self.add_test_step_row()

    # =========================================================
    # POSTCONDITION
    # =========================================================

    def add_postcondition_row(self, value=""):

        row = ctk.CTkFrame(
            self.post_frame,
            fg_color="transparent"
        )

        row.pack(
            fill="x",
            pady=3
        )

        entry = ctk.CTkEntry(row)

        entry.insert(
            0,
            value
        )

        entry.pack(
            side="left",
            fill="x",
            expand=True
        )

        delete_btn = ctk.CTkButton(
            row,
            text="Sil",
            width=50,
            fg_color="#C0392B",
            command=lambda: self.remove_row(
                row,
                self.post_rows
            )
        )

        delete_btn.pack(
            side="right",
            padx=(5, 0)
        )

        self.post_rows.append(
            (row, entry)
        )

        entry.bind(
            "<KeyRelease>",
            lambda event: self.check_postcondition_row()
        )

    def check_postcondition_row(self):

        if self.loading_data:
            return

        if not self.post_rows:
            self.add_postcondition_row()
            return

        last_row, last_entry = self.post_rows[-1]

        if last_entry.get().strip():

            self.add_postcondition_row()

    # =========================================================
    # SATIR SİL
    # =========================================================

    def remove_row(self, row, row_list):

        row.destroy()

        row_list[:] = [
            item
            for item in row_list
            if item[0] != row
        ]

        # Hiç satır kalmadıysa bir boş satır oluştur
        if not row_list:

            if row_list is self.pre_rows:
                self.add_precondition_row()

            elif row_list is self.data_rows:
                self.add_test_data_row()

            elif row_list is self.step_rows:
                self.add_test_step_row()

            elif row_list is self.post_rows:
                self.add_postcondition_row()

            return

        # Son satır doluysa bir boş satır ekle
        if row_list is self.pre_rows:

            if self.pre_rows[-1][1].get().strip():
                self.add_precondition_row()

        elif row_list is self.data_rows:

            last = self.data_rows[-1]

            if (
                last[1].get().strip()
                or last[2].get().strip()
            ):
                self.add_test_data_row()

        elif row_list is self.step_rows:

            last = self.step_rows[-1]

            if (
                last[2].get().strip()
                or last[3].get().strip()
                or last[4].get().strip()
            ):
                self.add_test_step_row()

        elif row_list is self.post_rows:

            if self.post_rows[-1][1].get().strip():
                self.add_postcondition_row()

        # Step numaralarını yeniden düzenle
        if row_list is self.step_rows:

            for index, item in enumerate(
                self.step_rows,
                start=1
            ):

                item[1].delete(0, "end")
                item[1].insert(0, str(index))

    # =========================================================
    # SAVE
    # =========================================================

    def save_test(self):

        tc_id = self.tc_id.get().strip()
        test_name = self.test_name.get().strip()

        if not tc_id or not test_name:

            messagebox.showerror(
                "Hata",
                "TC ID ve Test Adı boş bırakılamaz!"
            )

            return

        try:

            version = int(
                self.version.get().strip()
            )

            # =================================================
            # TEST CASE
            # =================================================

            if self.edit_mode:

                self.db.update_test_case(
                tc_id=tc_id,
                name=test_name,
                priority=self.priority.get(),
                application=self.application.get(),
                version=version,
                creator=self.creator.get(),
                create_date=self.create_date.get(),

                automation=(
                    1
                    if self.automation_completed.get()
                    else 0
                ),

                status=self.status.get(),

                automation_requested=(
                    1
                    if self.automation_requested.get()
                    else 0
                ),

                automation_completed=(
                    1
                    if self.automation_completed.get()
                    else 0
                ),

                automation_scenario=self.automation_scenario.get().strip(),

                test_type=self.test_type.get(),

                test_environment=self.test_environment.get().strip(),

                error_code=self.error_code.get().strip(),

                error_priority=self.error_priority.get()
            )
                

                # Eski detayları temizle

                self.db.delete_pre_conditions(tc_id)
                self.db.delete_test_data(tc_id)
                self.db.delete_test_steps(tc_id)
                self.db.delete_post_conditions(tc_id)

            else:

                self.db.add_test_case(
                tc_id=tc_id,
                name=test_name,
                priority=self.priority.get(),
                application=self.application.get(),
                version=version,
                creator=self.creator.get(),
                create_date=self.create_date.get(),

                automation=(
                    1
                    if self.automation_completed.get()
                    else 0
                ),

                status=self.status.get(),

                automation_requested=(
                    1
                    if self.automation_requested.get()
                    else 0
                ),

                automation_completed=(
                    1
                    if self.automation_completed.get()
                    else 0
                ),

                automation_scenario=self.automation_scenario.get().strip(),

                test_type=self.test_type.get(),

                test_environment=self.test_environment.get().strip(),

                error_code=self.error_code.get().strip(),

                error_priority=self.error_priority.get()
            )

            # =================================================
            # PRECONDITIONS
            # =================================================

            for row, entry in self.pre_rows:

                condition = entry.get().strip()

                if condition:

                    self.db.add_pre_condition(
                        tc_id,
                        condition
                    )

            # =================================================
            # TEST DATA
            # =================================================

            for (
                row,
                key_entry,
                value_entry
            ) in self.data_rows:

                data_key = key_entry.get().strip()
                data_value = value_entry.get().strip()

                if not data_key and not data_value:
                    continue

                self.db.add_test_data(
                    tc_id,
                    data_key,
                    data_value
                )

            # =================================================
            # TEST STEPS
            # =================================================

            for (
                row,
                step_entry,
                action_entry,
                data_entry,
                expected_entry
            ) in self.step_rows:

                step_no_text = step_entry.get().strip()
                action = action_entry.get().strip()
                test_data = data_entry.get().strip()
                expected_result = expected_entry.get().strip()

                # Tamamen boş satırı geç

                if (
                    not action
                    and not test_data
                    and not expected_result
                ):
                    continue

                if not step_no_text:

                    messagebox.showerror(
                        "Hata",
                        "Adım No boş bırakılamaz."
                    )

                    return

                step_no = int(step_no_text)

                self.db.add_test_step(
                    tc_id,
                    step_no,
                    action,
                    test_data,
                    expected_result
                )

            # =================================================
            # POSTCONDITIONS
            # =================================================

            for row, entry in self.post_rows:

                condition = entry.get().strip()

                if condition:

                    self.db.add_post_condition(
                        tc_id,
                        condition
                    )

            # =================================================
            # EXCEL OLUŞTUR
            # =================================================

            try:

                exporter = ExcelExporter(
                    self.db
                )

                exporter.export_single_test(
                    tc_id
                )

            except Exception as e:

                messagebox.showerror(
                    "Excel Hatası",
                    f"Test kaydedildi fakat Excel oluşturulamadı:\n\n{e}"
                )

                self.refresh_callback()
                self.close_window()
                return


            # =================================================
            # BAŞARILI
            # =================================================

            messagebox.showinfo(
                "Başarılı",
                "Test başarıyla kaydedildi ve Excel dosyası oluşturuldu."
            )

            self.refresh_callback()

            self.close_window()

        except sqlite3.IntegrityError:

            messagebox.showerror(
                "Hata",
                "Bu TC ID zaten mevcut!"
            )

        except ValueError:

            messagebox.showerror(
                "Hata",
                "Sürüm ve Adım No alanları sayı olmalıdır."
            )

        except Exception as e:

            messagebox.showerror(
                "Hata",
                f"Test kaydedilirken hata oluştu:\n\n{e}"
            )

    # =========================================================
    # EDIT MODE
    # =========================================================

    def load_test_data(self):

        self.loading_data = True

        test = self.db.get_test_case(
            self.edit_tc_id
        )

        if not test:
            self.loading_data = False
            return
        
        # =====================================================
        # TEST BİLGİLERİ
        # =====================================================

        self.tc_id.insert(
            0,
            test[0]
        )

        self.tc_id.configure(
            state="disabled"
        )

        self.test_name.insert(
            0,
            test[1]
        )

        self.priority.set(
            test[2]
        )

        self.application.insert(
            0,
            test[3]
        )

        self.version.delete(
            0,
            "end"
        )

        self.version.insert(
            0,
            str(test[4])
        )

        self.creator.insert(
            0,
            test[5]
        )

        self.create_date.delete(
            0,
            "end"
        )

        self.create_date.insert(
            0,
            test[6]
        )

        # =====================================================
        # OTOMASYON
        # =====================================================

        if len(test) > 9 and test[9]:
            self.automation_requested.select()

        if len(test) > 10 and test[10]:
            self.automation_completed.select()


        # =====================================================
        # DURUM
        # =====================================================

        if len(test) > 8 and test[8]:
            self.status.set(test[8])


        # =====================================================
        # TEST TÜRÜ
        # =====================================================

        if len(test) > 12 and test[12]:
            self.test_type.set(test[12])


        # =====================================================
        # TEST ORTAMI
        # =====================================================

        if len(test) > 13 and test[13]:

            self.test_environment.insert(
                0,
                test[13]
            )


        # =====================================================
        # OTOMASYON SENARYO
        # =====================================================

        if len(test) > 11 and test[11]:

            self.automation_scenario.insert(
                0,
                test[11]
            )


        # =====================================================
        # HATA KODU
        # =====================================================

        if len(test) > 14 and test[14]:

            self.error_code.insert(
                0,
                test[14]
            )


        # =====================================================
        # HATA ÖNCELİĞİ
        # =====================================================

        if len(test) > 15 and test[15]:

            self.error_priority.set(
                test[15]
            )

        # =====================================================
        # DEFAULT SATIRLARI TEMİZLE
        # =====================================================

        self.clear_rows(self.pre_rows)
        self.clear_rows(self.data_rows)
        self.clear_rows(self.step_rows)
        self.clear_rows(self.post_rows)

        # =====================================================
        # PRECONDITIONS
        # =====================================================

        pre = self.db.get_pre_conditions(
            self.edit_tc_id
        )

        for item in pre:

            self.add_precondition_row(
                item[0]
            )

        # En sona boş satır
        self.add_precondition_row()

        # =====================================================
        # TEST DATA
        # =====================================================

        data = self.db.get_test_data(
            self.edit_tc_id
        )

        for item in data:

            self.add_test_data_row(
                item[0],
                item[1]
            )

        self.add_test_data_row()

        # =====================================================
        # TEST STEPS
        # =====================================================

        steps = self.db.get_test_steps(
            self.edit_tc_id
        )

        for step in steps:

            self.add_test_step_row(
                step_no=step[0],
                action=step[1],
                test_data=step[2],
                expected_result=step[3]
            )

        self.add_test_step_row()

        # =====================================================
        # POSTCONDITIONS
        # =====================================================

        post = self.db.get_post_conditions(
            self.edit_tc_id
        )

        for item in post:

            self.add_postcondition_row(
                item[0]
            )

        self.add_postcondition_row()

        self.loading_data = False

    # =========================================================
    # SATIRLARI TEMİZLE
    # =========================================================

    def clear_rows(self, row_list):

        for item in row_list:

            item[0].destroy()

        row_list.clear()

    # =========================================================
    # WINDOW KAPAT
    # =========================================================

    def close_window(self):

        if self.owns_db:
            try:
                self.db.close()
            except Exception:
                pass

        self.destroy()

    # =========================================================
    # WINDOW DESTROY
    # =========================================================

    def destroy(self):

        try:
            self.unbind("<MouseWheel>")
            self.unbind("<Button-4>")
            self.unbind("<Button-5>")
            self.unbind("<Prior>")
            self.unbind("<Next>")
            self.unbind("<Home>")
            self.unbind("<End>")
        except Exception:
            pass

        try:
            if hasattr(self, "db") and self.owns_db:
                self.db.close()
        except Exception:
            pass

        super().destroy()