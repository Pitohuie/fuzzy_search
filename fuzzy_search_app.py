import pandas as pd
from fuzzywuzzy import process
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class FuzzySearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("模糊搜索")
        self.root.geometry("800x600")
        self.root.minsize(400, 300)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(7, weight=1)

        # 应用 ttk 主题
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # 文件选择按钮
        self.select_file_btn = ttk.Button(root, text="选择Excel文件", command=self.load_file)
        self.select_file_btn.grid(row=0, column=0, pady=10, padx=10, sticky='ew')

        # 多列名输入框
        self.column_label1 = ttk.Label(root, text="请输入第一个列名:", font=("Arial", 12))
        self.column_label1.grid(row=1, column=0, pady=5, padx=10, sticky='w')
        self.column_entry1 = ttk.Entry(root, width=40, font=("Arial", 12))
        self.column_entry1.grid(row=2, column=0, pady=5, padx=10, sticky='ew')

        self.column_label2 = ttk.Label(root, text="请输入第二个列名:", font=("Arial", 12))
        self.column_label2.grid(row=3, column=0, pady=5, padx=10, sticky='w')
        self.column_entry2 = ttk.Entry(root, width=40, font=("Arial", 12))
        self.column_entry2.grid(row=4, column=0, pady=5, padx=10, sticky='ew')

        self.column_label3 = ttk.Label(root, text="请输入第三个列名:", font=("Arial", 12))
        self.column_label3.grid(row=5, column=0, pady=5, padx=10, sticky='w')
        self.column_entry3 = ttk.Entry(root, width=40, font=("Arial", 12))
        self.column_entry3.grid(row=6, column=0, pady=5, padx=10, sticky='ew')

        # 搜索关键词输入框和搜索按钮
        self.entry_label = ttk.Label(root, text="请输入要搜索的关键词:", font=("Arial", 12))
        self.entry_label.grid(row=7, column=0, pady=5, padx=10, sticky='w')
        self.search_entry = ttk.Entry(root, width=40, font=("Arial", 12))
        self.search_entry.grid(row=8, column=0, pady=5, padx=10, sticky='ew')
        self.search_btn = ttk.Button(root, text="搜索", command=self.search_name)
        self.search_btn.grid(row=9, column=0, pady=10, padx=10, sticky='ew')

        # 结果显示的Treeview
        self.tree = None
        self.df = None  # 用于存储加载的DataFrame

    def load_file(self):
        file_path = filedialog.askopenfilename(title="选择Excel文件", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            try:
                self.df = pd.read_excel(file_path)
                if self.tree:
                    self.tree.destroy()  # 删除之前的Treeview
                self.create_treeview()
                messagebox.showinfo("成功", "文件加载成功！请在下方输入关键词进行搜索。")
            except Exception as e:
                messagebox.showerror("错误", f"无法加载文件: {e}")

    def create_treeview(self):
        # 在现有列的基础上添加“相似度”列
        columns = list(self.df.columns) + ["相似度"]
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", style="Custom.Treeview")

        for col in columns:
            self.tree.heading(col, text=col, anchor="center")
            self.tree.column(col, anchor="center", minwidth=100, stretch=True)

        self.tree.grid(row=10, column=0, sticky='nsew', padx=10, pady=10)
        self.root.rowconfigure(10, weight=1)

    def search_name(self):
        if self.df is None:
            messagebox.showwarning("警告", "请先选择一个Excel文件。")
            return

        # 获取输入的列名
        columns = []
        if self.column_entry1.get():
            columns.append(self.column_entry1.get())
        if self.column_entry2.get():
            columns.append(self.column_entry2.get())
        if self.column_entry3.get():
            columns.append(self.column_entry3.get())

        if not columns:
            messagebox.showerror("错误", "请至少输入一个列名。")
            return

        search_name = self.search_entry.get()
        missing_columns = [col for col in columns if col not in self.df.columns]

        if missing_columns:
            messagebox.showerror("错误", f"以下列名不存在于Excel文件中，请检查列名: {', '.join(missing_columns)}")
            return

        combined_results = {}
        for col in columns:
            matches = process.extract(search_name, self.df[col], limit=10)
            for match in matches:
                index = match[2]
                if index in combined_results:
                    combined_results[index]["相似度"] += match[1]
                else:
                    combined_results[index] = {
                        "数据": self.df.iloc[index].tolist(),
                        "相似度": match[1]
                    }

        # 按相似度排序并插入到Treeview中
        for item in self.tree.get_children():
            self.tree.delete(item)  # 清空之前的搜索结果

        sorted_results = sorted(combined_results.items(), key=lambda x: x[1]["相似度"], reverse=True)
        for i, (index, result) in enumerate(sorted_results):
            values = result["数据"] + [f"{result['相似度']}%"]
            self.tree.insert("", "end", values=values)

        self.auto_resize_columns()

    def auto_resize_columns(self):
        # 根据内容自动调整列宽
        for col in self.tree["columns"]:
            max_len = max([len(str(self.tree.set(item, col))) for item in self.tree.get_children()] + [len(col)])
            self.tree.column(col, width=max_len * 10)

if __name__ == "__main__":
    root = tk.Tk()
    app = FuzzySearchApp(root)
    root.mainloop()
