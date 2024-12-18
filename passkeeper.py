import random
import string
import sqlite3
from datetime import datetime
from tkinter import messagebox, filedialog
import pandas as pd

class PassKeeper:
    def __init__(self):
        self.db_connected = False
        self.db_connection = None
        self.db_cursor = None

        # 尝试连接SQLite数据库
        self.connect_db()

    def connect_db(self):
        try:
            self.db_connection = sqlite3.connect("PassKeeper.db")
            self.db_cursor = self.db_connection.cursor()
            self.db_connected = True
            self.db_cursor.execute('''CREATE TABLE IF NOT EXISTS data (
                                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                                    website TEXT,
                                    account TEXT,
                                    password TEXT,
                                    notes TEXT,
                                    created_at TEXT,
                                    updated_at TEXT)''')
        except Exception as e:
            print(f"数据库连接失败: {e}")
            self.db_connected = False

    def save_to_db(self):
        if not self.db_connected:
            return
        # 保存所有数据
        self.db_connection.commit()

    def load_from_db(self):
        if not self.db_connected:
            return []
        self.db_cursor.execute("SELECT * FROM data")
        rows = self.db_cursor.fetchall()
        return [{
            "id": row[0],
            "website": row[1],
            "account": row[2],
            "password": row[3],
            "notes": row[4],
            "created_at": row[5],
            "updated_at": row[6]
        } for row in rows]

    def add_password(self, website, account, password,  notes):
        created_at = updated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.db_cursor.execute('''INSERT INTO data (website, account, password,  notes, created_at, updated_at)
                                VALUES (?, ?, ?, ?, ?, ?)''', (
            website, account, password, notes, created_at, updated_at
        ))
        self.db_connection.commit()

    def update_password(self, password_id, website, account, password, notes):
        updated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if password == "******":
            self.db_cursor.execute(
                '''UPDATE data SET website=?, account=?, notes=?, updated_at=? WHERE id=?''',
                (website, account, notes, updated_at, password_id))
        else:
            self.db_cursor.execute(
                '''UPDATE data SET website=?, account=?, password=?, notes=?, updated_at=? WHERE id=?''',
                (website, account, password, notes, updated_at, password_id))
        self.db_connection.commit()

    def delete_password(self, password_id):
        self.db_cursor.execute('''DELETE FROM data WHERE id=?''', (password_id,))
        self.db_connection.commit()

    def generate_password(self, length=12):
        chars = string.ascii_letters + string.digits + string.punctuation
        return ''.join(random.choice(chars) for _ in range(length))

    def export_to_excel(self):
        # 自动生成默认文件名（应用名称 + 当前日期）
        app_name = "PassKeeper"  # 可根据需要更改为实际应用名称
        current_date = datetime.now().strftime("%Y-%m-%d")
        default_filename = f"{app_name}_{current_date}.xlsx"

        # 弹出文件保存对话框，让用户选择保存路径和文件名
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile=default_filename,
            title="保存密码数据"
        )

        # 如果用户选择了路径
        if file_path:
            try:
                data = self.load_from_db()

                # 将数据转换为 DataFrame
                df = pd.DataFrame(data)

                # 删除 id 列，不导出 id
                df = df.drop(columns=['id'])

                # 使用中文列名，并将“登记时间”修改为“更新时间”
                df.columns = ['网站', '账号', '密码', '备注', '创建时间', '更新时间']

                # 导出数据到 Excel
                df.to_excel(file_path, index=False)
                messagebox.showinfo("导出成功", f"文件已保存到：{file_path}")

            except Exception as e:
                messagebox.showerror("导出失败", f"导出过程中发生错误: {e}")

    def import_from_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx")])
        if not file_path:
            return
        df = pd.read_excel(file_path)
        for _, row in df.iterrows():
            self.add_password(row["账号"], row["密码"], row["网站"], row["备注"])
        # 刷新数据