import os
import sqlite3

# Proje kök dizini: database/database.py -> database/ -> TestCaseManager/
PROJECT_ROOT = os.path.dirname(
    os.path.dirname(os.path.abspath(__file__))
)

DB_PATH = os.path.join(PROJECT_ROOT, "data", "testcases.db")


class Database:

    def __init__(self):

        os.makedirs(
            os.path.dirname(DB_PATH),
            exist_ok=True
        )

        self.connection = sqlite3.connect(DB_PATH)

        self.cursor = self.connection.cursor()

        self.create_tables()

    def create_tables(self):

        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS test_cases(

                tc_id TEXT PRIMARY KEY,

                name TEXT NOT NULL,

                priority TEXT,

                application TEXT,

                version INTEGER,

                creator TEXT,

                create_date TEXT,

                automation INTEGER

            )
        """)

        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS pre_conditions(

                id INTEGER PRIMARY KEY AUTOINCREMENT,

                tc_id TEXT,

                condition TEXT,

                FOREIGN KEY(tc_id) REFERENCES test_cases(tc_id)

            )
        """)

        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS test_data(

                id INTEGER PRIMARY KEY AUTOINCREMENT,

                tc_id TEXT,

                data_key TEXT,

                data_value TEXT,

                FOREIGN KEY(tc_id) REFERENCES test_cases(tc_id)

            )
        """)

        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS test_steps(

                id INTEGER PRIMARY KEY AUTOINCREMENT,

                tc_id TEXT,

                step_no INTEGER,

                action TEXT,

                test_data TEXT,

                expected_result TEXT,

                FOREIGN KEY(tc_id) REFERENCES test_cases(tc_id)

            )
        """)

        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS post_conditions(

                id INTEGER PRIMARY KEY AUTOINCREMENT,

                tc_id TEXT,

                condition TEXT,

                FOREIGN KEY(tc_id) REFERENCES test_cases(tc_id)

            )
        """)

        self.connection.commit()

    def get_all_test_cases(self):

        self.cursor.execute("""
            SELECT *
            FROM test_cases
            ORDER BY tc_id
        """)

        return self.cursor.fetchall()

    def add_test_case(
        self,
        tc_id,
        name,
        priority,
        application,
        version,
        creator,
        create_date,
        automation
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
                automation
            )
            VALUES(?,?,?,?,?,?,?,?)
        """, (
            tc_id,
            name,
            priority,
            application,
            version,
            creator,
            create_date,
            automation
        ))

        self.connection.commit()

    def add_pre_condition(self, tc_id, condition):

        self.cursor.execute("""
            INSERT INTO pre_conditions(tc_id, condition)
            VALUES(?,?)
        """, (tc_id, condition))

        self.connection.commit()


    def get_test_case(self, tc_id):

        self.cursor.execute("""
            SELECT *
            FROM test_cases
            WHERE tc_id = ?
        """, (tc_id,))

        return self.cursor.fetchone()


    def add_test_data(self, tc_id, data_key, data_value):

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

    def add_post_condition(self, tc_id, condition):

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

    def get_pre_conditions(self, tc_id):

        self.cursor.execute("""
            SELECT condition
            FROM pre_conditions
            WHERE tc_id = ?
        """, (tc_id,))

        return self.cursor.fetchall()

    def get_test_data(self, tc_id):

        self.cursor.execute("""
            SELECT data_key, data_value
            FROM test_data
            WHERE tc_id = ?
        """, (tc_id,))

        return self.cursor.fetchall()

    def get_test_steps(self, tc_id):

        self.cursor.execute("""
            SELECT
                step_no,
                action,
                test_data,
                expected_result
            FROM test_steps
            WHERE tc_id = ?
            ORDER BY step_no
        """, (tc_id,))

        return self.cursor.fetchall()

    def get_post_conditions(self, tc_id):

        self.cursor.execute("""
            SELECT condition
            FROM post_conditions
            WHERE tc_id = ?
        """, (tc_id,))

        return self.cursor.fetchall()

    def update_test_case(
        self,
        tc_id,
        name,
        priority,
        application,
        version,
        creator,
        create_date,
        automation
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
                automation=?
            WHERE tc_id=?
        """, (
            name,
            priority,
            application,
            version,
            creator,
            create_date,
            automation,
            tc_id
        ))

        self.connection.commit()

    def delete_pre_conditions(self, tc_id):

        self.cursor.execute("""
            DELETE FROM pre_conditions
            WHERE tc_id = ?
        """, (tc_id,))

        self.connection.commit()

    def delete_post_conditions(self, tc_id):

        self.cursor.execute("""
            DELETE FROM post_conditions
            WHERE tc_id = ?
        """, (tc_id,))

        self.connection.commit()

    def delete_test_case(self, tc_id):

        self.cursor.execute("""
            DELETE FROM pre_conditions
            WHERE tc_id = ?
        """, (tc_id,))

        self.cursor.execute("""
            DELETE FROM post_conditions
            WHERE tc_id = ?
        """, (tc_id,))

        self.cursor.execute("""
            DELETE FROM test_data
            WHERE tc_id = ?
        """, (tc_id,))

        self.cursor.execute("""
            DELETE FROM test_steps
            WHERE tc_id = ?
        """, (tc_id,))

        self.cursor.execute("""
            DELETE FROM test_cases
            WHERE tc_id = ?
        """, (tc_id,))

        self.connection.commit()

    def delete_test_data(self, tc_id):

        self.cursor.execute("""
            DELETE FROM test_data
            WHERE tc_id = ?
        """, (tc_id,))

        self.connection.commit()

    def delete_test_steps(self, tc_id):

        self.cursor.execute("""
            DELETE FROM test_steps
            WHERE tc_id = ?
        """, (tc_id,))

        self.connection.commit()

    def close(self):

        self.connection.close()