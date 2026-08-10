import os
import sqlite3


# =========================================================
# PROJECT ROOT
# =========================================================

# database/database.py
# database/ -> TestCaseManager/

PROJECT_ROOT = os.path.dirname(
    os.path.dirname(
        os.path.abspath(__file__)
    )
)


# =========================================================
# PROJECTS DIRECTORY
# =========================================================

PROJECTS_PATH = os.path.join(
    PROJECT_ROOT,
    "projects"
)


# =========================================================
# DATABASE
# =========================================================

class Database:

    def __init__(self, project_name=None):

        """
        project_name verilirse:

            projects/
                ProjeAdi/
                    testcases.db

        şeklinde çalışır.

        project_name verilmezse eski database'i kullanır.
        """

        # -------------------------------------------------
        # ESKİ SİSTEM
        # -------------------------------------------------

        if not project_name:

            self.db_path = os.path.join(
                PROJECT_ROOT,
                "data",
                "testcases.db"
            )

        # -------------------------------------------------
        # YENİ PROJE SİSTEMİ
        # -------------------------------------------------

        else:

            # Proje adındaki gereksiz boşlukları temizle

            project_name = project_name.strip()

            # Proje klasörü

            project_path = os.path.join(
                PROJECTS_PATH,
                project_name
            )

            # Klasör yoksa oluştur

            os.makedirs(
                project_path,
                exist_ok=True
            )

            # Database yolu

            self.db_path = os.path.join(
                project_path,
                "testcases.db"
            )

        # =================================================
        # DATABASE KLASÖRÜ
        # =================================================

        os.makedirs(
            os.path.dirname(
                self.db_path
            ),
            exist_ok=True
        )

        # =================================================
        # SQLITE CONNECTION
        # =================================================

        self.connection = sqlite3.connect(
            self.db_path
        )

        self.cursor = self.connection.cursor()

        # =================================================
        # TABLOLARI OLUŞTUR
        # =================================================

        self.create_tables()


    # =========================================================
    # TABLES
    # =========================================================

    def create_tables(self):

        # -----------------------------------------------------
        # TEST CASES
        # -----------------------------------------------------

        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS test_cases(

                tc_id TEXT PRIMARY KEY,

                name TEXT NOT NULL,

                priority TEXT,

                application TEXT,

                version INTEGER,

                creator TEXT,

                create_date TEXT,

                automation INTEGER DEFAULT 0,

                status TEXT DEFAULT 'Beklemede'

            )
        """)


        # -----------------------------------------------------
        # PRE CONDITIONS
        # -----------------------------------------------------

        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS pre_conditions(

                id INTEGER PRIMARY KEY AUTOINCREMENT,

                tc_id TEXT,

                condition TEXT,

                FOREIGN KEY(tc_id)
                REFERENCES test_cases(tc_id)

            )
        """)


        # -----------------------------------------------------
        # TEST DATA
        # -----------------------------------------------------

        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS test_data(

                id INTEGER PRIMARY KEY AUTOINCREMENT,

                tc_id TEXT,

                data_key TEXT,

                data_value TEXT,

                FOREIGN KEY(tc_id)
                REFERENCES test_cases(tc_id)

            )
        """)


        # -----------------------------------------------------
        # TEST STEPS
        # -----------------------------------------------------

        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS test_steps(

                id INTEGER PRIMARY KEY AUTOINCREMENT,

                tc_id TEXT,

                step_no INTEGER,

                action TEXT,

                test_data TEXT,

                expected_result TEXT,

                FOREIGN KEY(tc_id)
                REFERENCES test_cases(tc_id)

            )
        """)


        # -----------------------------------------------------
        # POST CONDITIONS
        # -----------------------------------------------------

        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS post_conditions(

                id INTEGER PRIMARY KEY AUTOINCREMENT,

                tc_id TEXT,

                condition TEXT,

                FOREIGN KEY(tc_id)
                REFERENCES test_cases(tc_id)

            )
        """)


        # =====================================================
        # ESKİ DATABASE'LERE YENİ KOLONLARI EKLE
        # =====================================================

        self.add_missing_columns()

        self.connection.commit()


    # =========================================================
    # MISSING COLUMNS
    # =========================================================

    def add_missing_columns(self):

        self.cursor.execute("""
            PRAGMA table_info(test_cases)
        """)

        columns = [
            column[1]
            for column in self.cursor.fetchall()
        ]


        # -----------------------------------------------------
        # STATUS
        # -----------------------------------------------------

        if "status" not in columns:

            self.cursor.execute("""
                ALTER TABLE test_cases
                ADD COLUMN status TEXT DEFAULT 'Beklemede'
            """)


        # -----------------------------------------------------
        # OTOMASYONLAŞTIRILSIN MI?
        # -----------------------------------------------------

        if "automation_requested" not in columns:

            self.cursor.execute("""
                ALTER TABLE test_cases
                ADD COLUMN automation_requested
                INTEGER DEFAULT 0
            """)


        # -----------------------------------------------------
        # OTOMASYONLAŞTIRILDI MI?
        # -----------------------------------------------------

        if "automation_completed" not in columns:

            self.cursor.execute("""
                ALTER TABLE test_cases
                ADD COLUMN automation_completed
                INTEGER DEFAULT 0
            """)


        # -----------------------------------------------------
        # OTOMASYON SENARYO KARŞILIĞI
        # -----------------------------------------------------

        if "automation_scenario" not in columns:

            self.cursor.execute("""
                ALTER TABLE test_cases
                ADD COLUMN automation_scenario TEXT
            """)


        # -----------------------------------------------------
        # TEST TÜRÜ
        # -----------------------------------------------------

        if "test_type" not in columns:

            self.cursor.execute("""
                ALTER TABLE test_cases
                ADD COLUMN test_type TEXT
            """)


        # -----------------------------------------------------
        # TEST ORTAMI
        # -----------------------------------------------------

        if "test_environment" not in columns:

            self.cursor.execute("""
                ALTER TABLE test_cases
                ADD COLUMN test_environment TEXT
            """)


        # -----------------------------------------------------
        # HATA KODU
        # -----------------------------------------------------

        if "error_code" not in columns:

            self.cursor.execute("""
                ALTER TABLE test_cases
                ADD COLUMN error_code TEXT
            """)


        # -----------------------------------------------------
        # HATA ÖNCELİĞİ
        # -----------------------------------------------------

        if "error_priority" not in columns:

            self.cursor.execute("""
                ALTER TABLE test_cases
                ADD COLUMN error_priority TEXT
            """)


    # =========================================================
    # GET ALL TEST CASES
    # =========================================================

    def get_all_test_cases(self):

        self.cursor.execute("""
            SELECT *
            FROM test_cases
            ORDER BY tc_id
        """)

        return self.cursor.fetchall()


    # =========================================================
    # ADD TEST CASE
    # =========================================================

    def add_test_case(
        self,
        tc_id,
        name,
        priority,
        application,
        version,
        creator,
        create_date,
        automation,
        status="Beklemede",
        automation_requested=0,
        automation_completed=0,
        automation_scenario="",
        test_type="",
        test_environment="",
        error_code="",
        error_priority=""
    ):

        self.cursor.execute("""
            INSERT INTO test_cases(

                tc_id,
                name,
                priority,
                application,
                version,
                creator,
                create_date,
                automation,
                status,
                automation_requested,
                automation_completed,
                automation_scenario,
                test_type,
                test_environment,
                error_code,
                error_priority

            )

            VALUES(
                ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
            )

        """, (

            tc_id,
            name,
            priority,
            application,
            version,
            creator,
            create_date,
            automation,
            status,
            automation_requested,
            automation_completed,
            automation_scenario,
            test_type,
            test_environment,
            error_code,
            error_priority

        ))

        self.connection.commit()


    # =========================================================
    # GET TEST CASE
    # =========================================================

    def get_test_case(
        self,
        tc_id
    ):

        self.cursor.execute("""
            SELECT *
            FROM test_cases
            WHERE tc_id = ?
        """, (
            tc_id,
        ))

        return self.cursor.fetchone()


    # =========================================================
    # UPDATE TEST CASE
    # =========================================================

    def update_test_case(
        self,
        tc_id,
        name,
        priority,
        application,
        version,
        creator,
        create_date,
        automation,
        status="Beklemede",
        automation_requested=0,
        automation_completed=0,
        automation_scenario="",
        test_type="",
        test_environment="",
        error_code="",
        error_priority=""
    ):

        self.cursor.execute("""
            UPDATE test_cases

            SET

                name=?,

                priority=?,

                application=?,

                version=?,

                creator=?,

                create_date=?,

                automation=?,

                status=?,

                automation_requested=?,

                automation_completed=?,

                automation_scenario=?,

                test_type=?,

                test_environment=?,

                error_code=?,

                error_priority=?

            WHERE tc_id=?

        """, (

            name,
            priority,
            application,
            version,
            creator,
            create_date,
            automation,
            status,
            automation_requested,
            automation_completed,
            automation_scenario,
            test_type,
            test_environment,
            error_code,
            error_priority,
            tc_id

        ))

        self.connection.commit()


    # =========================================================
    # PRE CONDITIONS
    # =========================================================

    def add_pre_condition(
        self,
        tc_id,
        condition
    ):

        self.cursor.execute("""
            INSERT INTO pre_conditions(
                tc_id,
                condition
            )
            VALUES(?,?)
        """, (
            tc_id,
            condition
        ))

        self.connection.commit()


    def get_pre_conditions(
        self,
        tc_id
    ):

        self.cursor.execute("""
            SELECT condition
            FROM pre_conditions
            WHERE tc_id = ?
        """, (
            tc_id,
        ))

        return self.cursor.fetchall()


    def delete_pre_conditions(
        self,
        tc_id
    ):

        self.cursor.execute("""
            DELETE FROM pre_conditions
            WHERE tc_id = ?
        """, (
            tc_id,
        ))

        self.connection.commit()


    # =========================================================
    # TEST DATA
    # =========================================================

    def add_test_data(
        self,
        tc_id,
        data_key,
        data_value
    ):

        self.cursor.execute("""
            INSERT INTO test_data(
                tc_id,
                data_key,
                data_value
            )
            VALUES(?,?,?)
        """, (
            tc_id,
            data_key,
            data_value
        ))

        self.connection.commit()


    def get_test_data(
        self,
        tc_id
    ):

        self.cursor.execute("""
            SELECT
                data_key,
                data_value
            FROM test_data
            WHERE tc_id = ?
        """, (
            tc_id,
        ))

        return self.cursor.fetchall()


    def delete_test_data(
        self,
        tc_id
    ):

        self.cursor.execute("""
            DELETE FROM test_data
            WHERE tc_id = ?
        """, (
            tc_id,
        ))

        self.connection.commit()


    # =========================================================
    # TEST STEPS
    # =========================================================

    def add_test_step(
        self,
        tc_id,
        step_no,
        action,
        test_data,
        expected_result
    ):

        self.cursor.execute("""
            INSERT INTO test_steps(
                tc_id,
                step_no,
                action,
                test_data,
                expected_result
            )
            VALUES(?,?,?,?,?)
        """, (
            tc_id,
            step_no,
            action,
            test_data,
            expected_result
        ))

        self.connection.commit()


    def get_test_steps(
        self,
        tc_id
    ):

        self.cursor.execute("""
            SELECT
                step_no,
                action,
                test_data,
                expected_result

            FROM test_steps

            WHERE tc_id = ?

            ORDER BY step_no
        """, (
            tc_id,
        ))

        return self.cursor.fetchall()


    def delete_test_steps(
        self,
        tc_id
    ):

        self.cursor.execute("""
            DELETE FROM test_steps
            WHERE tc_id = ?
        """, (
            tc_id,
        ))

        self.connection.commit()


    # =========================================================
    # POST CONDITIONS
    # =========================================================

    def add_post_condition(
        self,
        tc_id,
        condition
    ):

        self.cursor.execute("""
            INSERT INTO post_conditions(
                tc_id,
                condition
            )
            VALUES(?,?)
        """, (
            tc_id,
            condition
        ))

        self.connection.commit()


    def get_post_conditions(
        self,
        tc_id
    ):

        self.cursor.execute("""
            SELECT condition
            FROM post_conditions
            WHERE tc_id = ?
        """, (
            tc_id,
        ))

        return self.cursor.fetchall()


    def delete_post_conditions(
        self,
        tc_id
    ):

        self.cursor.execute("""
            DELETE FROM post_conditions
            WHERE tc_id = ?
        """, (
            tc_id,
        ))

        self.connection.commit()


    # =========================================================
    # DELETE TEST CASE
    # =========================================================

    def delete_test_case(
        self,
        tc_id
    ):

        # -----------------------------------------------------
        # PRE CONDITIONS
        # -----------------------------------------------------

        self.cursor.execute("""
            DELETE FROM pre_conditions
            WHERE tc_id = ?
        """, (
            tc_id,
        ))


        # -----------------------------------------------------
        # POST CONDITIONS
        # -----------------------------------------------------

        self.cursor.execute("""
            DELETE FROM post_conditions
            WHERE tc_id = ?
        """, (
            tc_id,
        ))


        # -----------------------------------------------------
        # TEST DATA
        # -----------------------------------------------------

        self.cursor.execute("""
            DELETE FROM test_data
            WHERE tc_id = ?
        """, (
            tc_id,
        ))


        # -----------------------------------------------------
        # TEST STEPS
        # -----------------------------------------------------

        self.cursor.execute("""
            DELETE FROM test_steps
            WHERE tc_id = ?
        """, (
            tc_id,
        ))


        # -----------------------------------------------------
        # TEST CASE
        # -----------------------------------------------------

        self.cursor.execute("""
            DELETE FROM test_cases
            WHERE tc_id = ?
        """, (
            tc_id,
        ))

        self.connection.commit()


    # =========================================================
    # CLOSE
    # =========================================================

    def close(self):

        self.connection.close()