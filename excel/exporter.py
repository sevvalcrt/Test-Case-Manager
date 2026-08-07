import os
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from database.database import Database

# Proje kök dizini: excel/exporter.py -> excel/ -> TestCaseManager/
PROJECT_ROOT = os.path.dirname(
    os.path.dirname(os.path.abspath(__file__))
)

TEMPLATE_PATH = os.path.join(PROJECT_ROOT, "excel", "template.xlsx")
EXPORT_PATH = os.path.join(PROJECT_ROOT, "export.xlsx")


class ExcelExporter:

    def __init__(self):
        self.db = Database()

    # -------------------------------------------------
    # Hücre stilini başka hücreye kopyala
    # -------------------------------------------------

    def copy_style(self, source, target):

        if source.has_style:
            target._style = copy(source._style)

        if source.number_format:
            target.number_format = source.number_format

        if source.alignment:
            target.alignment = copy(source.alignment)

        if source.font:
            target.font = copy(source.font)

        if source.fill:
            target.fill = copy(source.fill)

        if source.border:
            target.border = copy(source.border)

        if source.protection:
            target.protection = copy(source.protection)

    # -------------------------------------------------
    # Satır stilini kopyala
    # -------------------------------------------------

    def copy_row_style(self, sheet, source_row, target_row):

        for col in range(1, 5):

            source = sheet.cell(
                row=source_row,
                column=col
            )

            target = sheet.cell(
                row=target_row,
                column=col
            )

            self.copy_style(source, target)

    # -------------------------------------------------
    # Bölüm başlığı oluştur
    # -------------------------------------------------

    def create_section_header(self, sheet, row, text):

        fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )

        font = Font(
            color="FFFFFF",
            bold=True,
            size=11
        )

        alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        for col in range(1, 5):

            cell = sheet.cell(
                row=row,
                column=col
            )

            cell.fill = fill
            cell.font = font
            cell.alignment = alignment

        sheet.cell(
            row=row,
            column=1
        ).value = text

        sheet.row_dimensions[row].height = 22

    # -------------------------------------------------
    # Test Steps başlıkları
    # -------------------------------------------------

    def create_step_headers(self, sheet, row):

        headers = [
            "Step No",
            "Action",
            "Expected Result",
            "Actual Result"
        ]

        fill = PatternFill(
            fill_type="solid",
            fgColor="D9EAF7"
        )

        font = Font(
            color="000000",
            bold=True,
            size=11
        )

        alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        for col, value in enumerate(headers, start=1):

            cell = sheet.cell(
                row=row,
                column=col
            )

            cell.value = value
            cell.fill = fill
            cell.font = font
            cell.alignment = alignment

        sheet.row_dimensions[row].height = 22

    # -------------------------------------------------
    # EXPORT
    # -------------------------------------------------

    def export(self):

        workbook = load_workbook(
            TEMPLATE_PATH
        )

        summary = workbook["Test Senaryoları"]

        # =================================================
        # TEST SENARYOLARI BAŞLIKLARI
        # =================================================

        headers = [
            "Test ID",
            "Test Name",
            "Priority",
            "Module",
            "Status",
            "Created Date"
        ]

        header_fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )

        header_font = Font(
            color="FFFFFF",
            bold=True,
            size=11
        )

        header_alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        for col, header in enumerate(headers, start=1):

            cell = summary.cell(
                row=1,
                column=col
            )

            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment

        summary.row_dimensions[1].height = 25

        # Eski testleri temizle
        if summary.max_row > 1:
            summary.delete_rows(
                2,
                summary.max_row - 1
            )

        tests = self.db.get_all_test_cases()

        row = 2

        # =================================================
        # TESTLER
        # =================================================

        for test in tests:

            tc_id = test[0]

            # ---------------------------------------------
            # ÖZET SAYFASI
            # ---------------------------------------------

            cell = summary.cell(
                row=row,
                column=1
            )

            cell.value = tc_id

            cell.hyperlink = (
                f"#'{tc_id}'!A1"
            )

            cell.style = "Hyperlink"

            summary.cell(
                row=row,
                column=2
            ).value = test[1]

            summary.cell(
                row=row,
                column=3
            ).value = test[2]

            summary.cell(
                row=row,
                column=4
            ).value = test[3]

            summary.cell(
                row=row,
                column=5
            ).value = test[4]

            summary.cell(
                row=row,
                column=6
            ).value = test[6]

            row += 1

            # =================================================
            # DETAY SAYFASI
            # =================================================

            if tc_id in workbook.sheetnames:
                del workbook[tc_id]

            template = workbook["Template"]

            detail = workbook.copy_worksheet(
                template
            )

            detail.title = tc_id

            # ---------------------------------------------
            # KOLON GENİŞLİKLERİ
            # ---------------------------------------------

            detail.column_dimensions["A"].width = 22
            detail.column_dimensions["B"].width = 35
            detail.column_dimensions["C"].width = 35
            detail.column_dimensions["D"].width = 35

            # ---------------------------------------------
            # DATABASE
            # ---------------------------------------------

            pre = self.db.get_pre_conditions(
                tc_id
            )

            data = self.db.get_test_data(
                tc_id
            )

            steps = self.db.get_test_steps(
                tc_id
            )

            post = self.db.get_post_conditions(
                tc_id
            )

            # =================================================
            # TEST BİLGİLERİ
            # =================================================

            detail["A1"] = "← Test Listesine Dön"

            detail["A1"].hyperlink = (
                "#'Test Senaryoları'!A1"
            )

            detail["A1"].style = "Hyperlink"

            labels = [
                "Test ID",
                "Test Name",
                "Priority",
                "Module",
                "Status",
                "Created By",
                "Created Date",
                "Automated"
            ]

            values = [
                tc_id,
                test[1],
                test[2],
                test[3],
                test[4],
                test[5],
                test[6],
                "Yes" if test[7] else "No"
            ]

            for i, (label, value) in enumerate(
                zip(labels, values),
                start=2
            ):

                detail.cell(
                    row=i,
                    column=1
                ).value = label

                detail.cell(
                    row=i,
                    column=2
                ).value = value

                # Label stili
                cell = detail.cell(
                    row=i,
                    column=1
                )

                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor="D9EAF7"
                )

                cell.font = Font(
                    color="000000",
                    bold=True,
                    size=11
                )

                cell.alignment = Alignment(
                    horizontal="left",
                    vertical="center"
                )

            # =================================================
            # DİNAMİK BÖLÜMLER
            # =================================================

            # Bilgiler 2-9 arası olduğu için
            # 11. satırdan başlıyoruz.

            current_row = 11

            # =================================================
            # PRECONDITIONS
            # =================================================

            self.create_section_header(
                detail,
                current_row,
                "Preconditions"
            )

            current_row += 1

            # Veri yoksa bile 1 boş satır
            if not pre:

                detail.cell(
                    row=current_row,
                    column=2
                ).value = ""

                current_row += 1

            else:

                for item in pre:

                    detail.cell(
                        row=current_row,
                        column=2
                    ).value = item[0]

                    current_row += 1

            # Bölümler arasında boşluk
            current_row += 1

            # =================================================
            # TEST DATA
            # =================================================

            self.create_section_header(
                detail,
                current_row,
                "Test Data"
            )

            current_row += 1

            if not data:

                # En az 1 boş satır
                current_row += 1

            else:

                for item in data:

                    detail.cell(
                        row=current_row,
                        column=2
                    ).value = item[0]

                    detail.cell(
                        row=current_row,
                        column=3
                    ).value = item[1]

                    current_row += 1

            current_row += 1

            # =================================================
            # TEST STEPS
            # =================================================

            self.create_section_header(
                detail,
                current_row,
                "Test Steps"
            )

            current_row += 1

            # Kolon başlıkları
            self.create_step_headers(
                detail,
                current_row
            )

            current_row += 1

            # Veri yoksa 1 boş step
            if not steps:

                current_row += 1

            else:

                for step in steps:

                    detail.cell(
                        row=current_row,
                        column=1
                    ).value = step[0]

                    detail.cell(
                        row=current_row,
                        column=2
                    ).value = step[1]

                    detail.cell(
                        row=current_row,
                        column=3
                    ).value = step[2]

                    detail.cell(
                        row=current_row,
                        column=4
                    ).value = step[3]

                    # Metinleri hücre içinde göster
                    for col in range(1, 5):

                        detail.cell(
                            row=current_row,
                            column=col
                        ).alignment = Alignment(
                            vertical="top",
                            wrap_text=True
                        )

                    current_row += 1

            current_row += 1

            # =================================================
            # POSTCONDITIONS
            # =================================================

            self.create_section_header(
                detail,
                current_row,
                "Postconditions"
            )

            current_row += 1

            if not post:

                current_row += 1

            else:

                for item in post:

                    detail.cell(
                        row=current_row,
                        column=2
                    ).value = item[0]

                    current_row += 1

        # =================================================
        # ÖZET KOLON GENİŞLİKLERİ
        # =================================================

        summary.column_dimensions["A"].width = 18
        summary.column_dimensions["B"].width = 40
        summary.column_dimensions["C"].width = 15
        summary.column_dimensions["D"].width = 20
        summary.column_dimensions["E"].width = 10
        summary.column_dimensions["F"].width = 18

        # =================================================
        # TEMPLATE'İ EN SONA TAŞI
        # =================================================

        template = workbook["Template"]

        workbook._sheets.remove(template)
        workbook._sheets.append(template)

        template.sheet_state = "hidden"

        # =================================================
        # KAYDET
        # =================================================

        workbook.save(
            EXPORT_PATH
        )

        self.db.close()