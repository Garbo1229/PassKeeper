from passkeeper import PassKeeper
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pyperclip

class PassKeeperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("网站密码管理器 -- PassKeeper -- Github:Garbo1229")
        self.passkeeper = PassKeeper()

        self.setup_main_screen()

    def import_from_excel(self):
        self.passkeeper.import_from_excel()
        self.update_password_listbox()  # 导入后刷新数据

    def setup_main_screen(self):
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(padx=20, pady=20, fill="both", expand=True)

        # 搜索框
        self.search_label = tk.Label(self.main_frame, text="搜索：", font=("Arial", 12))
        self.search_label.grid(row=0, column=0, padx=10, pady=10)

        self.search_entry = tk.Entry(self.main_frame, width=50, font=("Arial", 10), bd=2)
        # self.search_entry.grid(row=0, column=1)
        self.search_entry.grid(row=0, column=1, padx=10, pady=20, sticky="ew")  # 增加上下内边距
        self.search_entry.bind("<KeyRelease>", self.search_data)

        # 创建表格和滚动条
        self.treeview_frame = tk.Frame(self.main_frame)
        self.treeview_frame.grid(row=1, column=0, columnspan=3, pady=10, sticky="nsew")

        # 为 grid 布局的行和列设置权重，确保表格可以自动扩展
        self.main_frame.grid_rowconfigure(1, weight=1)  # 第二行的表格行允许伸缩
        self.main_frame.grid_columnconfigure(1, weight=1)  # 第1列搜索框的空间伸缩
        self.treeview_frame.grid_rowconfigure(0, weight=1)  # 表格容器的第一行允许伸缩
        self.treeview_frame.grid_columnconfigure(0, weight=1)  # 表格容器的第一列允许伸缩

        # 创建 Treeview 表格
        self.treeview = ttk.Treeview(self.treeview_frame,
                                     columns=("website", "account", "password", "notes", "created_at", "updated_at"),
                                     show="headings")
        self.treeview.pack(side="left", fill="both", expand=True)

        # 创建滚动条
        self.treeview_scrollbar = tk.Scrollbar(self.treeview_frame, orient="vertical", command=self.treeview.yview)
        self.treeview_scrollbar.pack(side="right", fill="y")

        # 配置滚动条与Treeview的绑定
        self.treeview.configure(yscrollcommand=self.treeview_scrollbar.set)

        # 定义列标题
        self.treeview.heading("website", text="网站")
        self.treeview.heading("account", text="账号")
        self.treeview.heading("password", text="密码")
        self.treeview.heading("notes", text="备注")
        self.treeview.heading("created_at", text="创建时间")
        self.treeview.heading("updated_at", text="更新时间")

        # 调整列宽
        self.treeview.column("website", width=120)
        self.treeview.column("account", width=120)
        self.treeview.column("password", width=120)
        self.treeview.column("notes", width=150)
        self.treeview.column("created_at", width=120)
        self.treeview.column("updated_at", width=120)

        # 绑定点击事件到表格
        self.treeview.bind("<ButtonRelease-1>", self.on_item_click)

        # 按钮样式
        button_style = {
            "bd": 1,  # 较细的边框
            "relief": "flat",  # 去掉阴影效果
            "bg": "#e0e0e0",  # 柔和的背景颜色
            "fg": "#333333",  # 深色前景文字
            "width": 15
        }

        # 添加数据按钮
        self.add_button = tk.Button(self.main_frame, text="添加新数据", command=self.add_entry, **button_style)
        self.add_button.grid(row=2, column=0, pady=10)

        # 更新和删除数据按钮
        self.update_button = tk.Button(self.main_frame, text="更新选中数据", command=self.update_entry, **button_style)
        self.update_button.grid(row=2, column=1, pady=10)

        self.delete_button = tk.Button(self.main_frame, text="删除选中数据", command=self.delete_entry, **button_style)
        self.delete_button.grid(row=2, column=2, pady=10)

        # 导入导出按钮
        self.export_button = tk.Button(self.main_frame, text="导出到Excel", command=self.passkeeper.export_to_excel,
                                       **button_style)
        self.export_button.grid(row=3, column=0, pady=10)

        self.import_button = tk.Button(self.main_frame, text="从Excel导入", command=self.import_from_excel,
                                       **button_style)
        self.import_button.grid(row=3, column=1, pady=10)

        # 生成随机密码按钮
        self.generate_button = tk.Button(self.main_frame, text="生成随机密码", command=self.generate_random_password,
                                         **button_style)
        self.generate_button.grid(row=3, column=2, pady=10)

        # 刷新数据
        self.update_password_listbox()

    def generate_random_password(self):
        password = self.passkeeper.generate_password(length=12)

        # 将密码复制到剪贴板
        pyperclip.copy(password)

        # 显示信息提示
        messagebox.showinfo("密码生成成功", f"已生成随机密码: {password}\n\n密码已复制到剪贴板，您可以粘贴使用！")

    def update_password_listbox(self):
        # 清空表格数据
        for item in self.treeview.get_children():
            self.treeview.delete(item)

        # 获取数据并显示
        password_data = self.passkeeper.load_from_db()
        for entry in password_data:
            # 使用数据库的 id 作为 iid
            self.treeview.insert("", "end", iid=entry['id'], values=(
                entry['website'], entry['account'], "******", entry['notes'], entry['created_at'], entry['updated_at']))
        # 自动滑动到最后
        self.treeview.yview_moveto(1)

    def search_data(self, event=None):
        search_term = self.search_entry.get().lower()
        filtered_data = [entry for entry in self.passkeeper.load_from_db() if
                         search_term in entry['website'].lower() or search_term in entry['account'].lower() or search_term in entry['notes'].lower()]
        for item in self.treeview.get_children():
            self.treeview.delete(item)
        for entry in filtered_data:
            self.treeview.insert("", "end", iid=entry['id'], values=(
                entry['website'], entry['account'], "******", entry['notes'], entry['created_at'], entry['updated_at']))

    def add_entry(self):
        new_window = tk.Toplevel(self.root)
        new_window.title("添加新数据")
        new_window.geometry("300x250")

        # 样式
        label_style = {
            "font": ("Arial", 10, "normal"),  # 使用普通字体，不加粗
            "fg": "#333333",  # 深色文字
            "bg": "#f0f0f0",  # 浅色背景
        }
        entry_style = {"font": ("Arial", 10), "bd": 2, "highlightthickness": 0}

        tk.Label(new_window, text="网站：", **label_style).grid(row=0, column=0, padx=10, pady=5)
        website_entry = tk.Entry(new_window, **entry_style)
        website_entry.grid(row=0, column=1)

        tk.Label(new_window, text="账号：", **label_style).grid(row=1, column=0, padx=10, pady=5)
        account_entry = tk.Entry(new_window, **entry_style)
        account_entry.grid(row=1, column=1)

        tk.Label(new_window, text="密码：", **label_style).grid(row=2, column=0, padx=10, pady=5)
        password_entry = tk.Entry(new_window, **entry_style)
        password_entry.grid(row=2, column=1)

        tk.Label(new_window, text="备注：", **label_style).grid(row=3, column=0, padx=10, pady=5)
        notes_entry = tk.Entry(new_window, **entry_style)
        notes_entry.grid(row=3, column=1)

        def save_new_entry():
            website = website_entry.get()
            account = account_entry.get()
            password = password_entry.get()
            notes = notes_entry.get()

            self.passkeeper.add_password(website, account, password, notes)
            self.update_password_listbox()
            new_window.destroy()

        save_button = tk.Button(new_window, text="保存", command=save_new_entry, font=("Arial", 10), bd=2)
        save_button.grid(row=4, columnspan=2, pady=10)

    def update_entry(self):
        selected_item = self.treeview.selection()
        if not selected_item:
            messagebox.showwarning("选择错误", "请先选择一个数据！")
            return
        item_length = len(selected_item)
        if item_length != 1:
            messagebox.showwarning("选择错误", "不能修改多条数据！")
            return

        # 取得选中的项 ID
        selected_item_id = selected_item[0]

        # 获取该项的所有数据
        item_values = self.treeview.item(selected_item_id)["values"]

        # 如果 item_values 为空，直接返回
        if not item_values:
            messagebox.showwarning("选择错误", "无法获取该数据的数据！")
            return

        # 创建更新窗口
        new_window = tk.Toplevel(self.root)
        new_window.title("更新数据")
        new_window.geometry("300x250")

        tk.Label(new_window, text="网站：", font=("Arial", 10)).grid(row=0, column=0, padx=10, pady=5)
        website_entry = tk.Entry(new_window, font=("Arial", 10), bd=2)
        website_entry.insert(0, item_values[0])
        website_entry.grid(row=0, column=1)

        tk.Label(new_window, text="账号：", font=("Arial", 10)).grid(row=1, column=0, padx=10, pady=5)
        account_entry = tk.Entry(new_window, font=("Arial", 10), bd=2)
        account_entry.insert(0, item_values[1])
        account_entry.grid(row=1, column=1)

        tk.Label(new_window, text="密码：", font=("Arial", 10)).grid(row=2, column=0, padx=10, pady=5)
        password_entry = tk.Entry(new_window, font=("Arial", 10), bd=2)
        password_entry.insert(0, item_values[2])
        password_entry.grid(row=2, column=1)

        tk.Label(new_window, text="备注：", font=("Arial", 10)).grid(row=3, column=0, padx=10, pady=5)
        notes_entry = tk.Entry(new_window, font=("Arial", 10), bd=2)
        notes_entry.insert(0, item_values[3])
        notes_entry.grid(row=3, column=1)

        # 弹出确认对话框，提示用户删除操作不可恢复

        def save_updated_entry():
            account = account_entry.get()
            password = password_entry.get()
            website = website_entry.get()
            notes = notes_entry.get()
            password_id = selected_item_id
            confirm = messagebox.askyesno("确认修改", "您确定要修改此条吗？\n修改后无法恢复！")
            if not confirm:  # 如果用户确认删除
                messagebox.showinfo("取消修改", "修改操作已取消。")
            else:
                self.passkeeper.update_password(password_id, website, account, password, notes)
                self.update_password_listbox()
                new_window.destroy()

        save_button = tk.Button(new_window, text="保存", command=save_updated_entry, font=("Arial", 10), bd=2)
        save_button.grid(row=4, columnspan=2, pady=10)

    def delete_entry(self):
        selected_item = self.treeview.selection()

        if not selected_item:
            messagebox.showwarning("选择错误", "请先选择一个数据")
            return
        item_length = len(selected_item)
        confirm = messagebox.askyesno("确认删除", "您确定要删除数据吗？\n删除后无法恢复！")
        if confirm:  # 如果用户确认删除
            if item_length == 1:
                # 弹出确认对话框，提示用户删除操作不可恢复
                password_id = selected_item[0]
                self.passkeeper.delete_password(password_id)
                self.update_password_listbox()
            else:
                for item in selected_item:
                    # 弹出确认对话框，提示用户删除操作不可恢复
                    if confirm:  # 如果用户确认删除
                        self.passkeeper.delete_password(item)
                        self.update_password_listbox()

            messagebox.showinfo("删除成功", str(item_length) + "条数据已成功删除！")
        else:
            messagebox.showinfo("取消删除", "删除操作已取消。")

    def on_item_click(self, event):
        item_id = self.treeview.identify_row(event.y)
        if not item_id:
            return

        # 获取点击的列ID
        column_id = self.treeview.identify_column(event.x)

        # 只处理点击密码列
        if column_id == "#3":
            # 获取数据库中对应的数据
            password_data = self.passkeeper.load_from_db()

            # 遍历数据，找到与 item_id 匍配的数据
            matching_item = next((entry for entry in password_data if str(entry['id']) == item_id), None)

            if matching_item:
                current_password = matching_item['password']
                item_values = self.treeview.item(item_id)['values']

                # 切换显示/隐藏密码
                if item_values[2] == "******":
                    # 显示密码
                    self.treeview.item(item_id, values=(
                        item_values[0], item_values[1], current_password, item_values[3], item_values[4],
                        item_values[5]))
                else:
                    # 隐藏密码
                    self.treeview.item(item_id, values=(
                        item_values[0], item_values[1], "******", item_values[3], item_values[4], item_values[5]))

