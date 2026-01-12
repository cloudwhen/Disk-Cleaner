import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
from cleaner import DiskCleaner
import os


class DiskCleanerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("C盘深度清理工具")
        self.root.geometry("900x700")
        self.cleaner = DiskCleaner()
        self.scan_results = {}
        self.category_nodes = {}
        self.item_nodes = {}
        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)

        title_label = ttk.Label(
            main_frame,
            text="C盘深度清理工具",
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 20))

        info_label = ttk.Label(
            main_frame,
            text="扫描并清理系统临时文件、浏览器缓存、Edge旧版本、剪映缓存、系统日志、回收站和软件残留",
            font=("Arial", 10)
        )
        info_label.grid(row=1, column=0, pady=(0, 10))

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, pady=(0, 10))

        self.scan_button = ttk.Button(
            button_frame,
            text="扫描C盘",
            command=self.start_scan,
            width=15
        )
        self.scan_button.pack(side=tk.LEFT, padx=5)

        self.clean_button = ttk.Button(
            button_frame,
            text="清理选中",
            command=self.start_clean,
            width=15,
            state=tk.DISABLED
        )
        self.clean_button.pack(side=tk.LEFT, padx=5)

        ttk.Separator(button_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, padx=10, fill=tk.Y)

        self.expand_all_button = ttk.Button(
            button_frame,
            text="展开全部",
            command=self.expand_all,
            width=10,
            state=tk.DISABLED
        )
        self.expand_all_button.pack(side=tk.LEFT, padx=5)

        self.collapse_all_button = ttk.Button(
            button_frame,
            text="折叠全部",
            command=self.collapse_all,
            width=10,
            state=tk.DISABLED
        )
        self.collapse_all_button.pack(side=tk.LEFT, padx=5)

        ttk.Separator(button_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, padx=10, fill=tk.Y)

        self.select_all_button = ttk.Button(
            button_frame,
            text="全选",
            command=self.select_all,
            width=10,
            state=tk.DISABLED
        )
        self.select_all_button.pack(side=tk.LEFT, padx=5)

        self.deselect_all_button = ttk.Button(
            button_frame,
            text="取消全选",
            command=self.deselect_all,
            width=10,
            state=tk.DISABLED
        )
        self.deselect_all_button.pack(side=tk.LEFT, padx=5)

        results_frame = ttk.LabelFrame(main_frame, text="扫描结果", padding="10")
        results_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            results_frame,
            columns=("path", "size", "selected"),
            show="tree headings"
        )
        self.tree.heading("#0", text="类别")
        self.tree.heading("path", text="路径")
        self.tree.heading("size", text="大小")
        self.tree.heading("selected", text="选择")

        self.tree.column("#0", width=200)
        self.tree.column("path", width=400)
        self.tree.column("size", width=100)
        self.tree.column("selected", width=60)

        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<Double-1>", self.on_tree_double_click)

        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=4, column=0, sticky=(tk.W, tk.E))

        self.status_label = ttk.Label(
            status_frame,
            text="准备就绪",
            font=("Arial", 10)
        )
        self.status_label.pack(side=tk.LEFT)

        self.progress = ttk.Progressbar(
            status_frame,
            mode="indeterminate",
            length=200
        )
        self.progress.pack(side=tk.RIGHT)

        log_frame = ttk.LabelFrame(main_frame, text="操作日志", padding="10")
        log_frame.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, state=tk.DISABLED)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    def log(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            if column == "#4":
                item = self.tree.identify_row(event.y)
                if item:
                    current_value = self.tree.set(item, "selected")
                    new_value = "是" if current_value == "否" else "否"
                    self.tree.set(item, "selected", new_value)

                    parent = self.tree.parent(item)
                    if parent:
                        self._update_parent_selection(parent)

    def on_tree_double_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            if self.tree.item(item, "open"):
                self.tree.item(item, open=False)
            else:
                self.tree.item(item, open=True)

    def _update_parent_selection(self, parent):
        children = self.tree.get_children(parent)
        if not children:
            return

        all_selected = all(self.tree.set(child, "selected") == "是" for child in children)
        none_selected = all(self.tree.set(child, "selected") == "否" for child in children)

        if all_selected:
            self.tree.set(parent, "selected", "是")
        elif none_selected:
            self.tree.set(parent, "selected", "否")
        else:
            self.tree.set(parent, "selected", "部分")

        grandparent = self.tree.parent(parent)
        if grandparent:
            self._update_parent_selection(grandparent)

    def select_all(self):
        for item in self.tree.get_children():
            self._select_recursive(item, True)

    def deselect_all(self):
        for item in self.tree.get_children():
            self._select_recursive(item, False)

    def _select_recursive(self, item, select):
        self.tree.set(item, "selected", "是" if select else "否")
        for child in self.tree.get_children(item):
            self._select_recursive(child, select)

    def expand_all(self):
        for item in self.tree.get_children():
            self._expand_recursive(item)

    def collapse_all(self):
        for item in self.tree.get_children():
            self._collapse_recursive(item)

    def _expand_recursive(self, item):
        self.tree.item(item, open=True)
        for child in self.tree.get_children(item):
            self._expand_recursive(child)

    def _collapse_recursive(self, item):
        self.tree.item(item, open=False)
        for child in self.tree.get_children(item):
            self._collapse_recursive(child)

    def start_scan(self):
        self.scan_button.config(state=tk.DISABLED)
        self.clean_button.config(state=tk.DISABLED)
        self.select_all_button.config(state=tk.DISABLED)
        self.deselect_all_button.config(state=tk.DISABLED)
        self.expand_all_button.config(state=tk.DISABLED)
        self.collapse_all_button.config(state=tk.DISABLED)
        self.progress.start()
        self.status_label.config(text="正在扫描...")
        self.log("开始扫描C盘...")

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.category_nodes = {}
        self.item_nodes = {}

        thread = threading.Thread(target=self.scan)
        thread.start()

    def scan(self):
        try:
            self.scan_results = self.cleaner.scan_c_drive()

            self.root.after(0, self.display_results)
            self.root.after(0, lambda: self.status_label.config(text=f"扫描完成！发现 {len(self.scan_results)} 个可清理项"))
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.scan_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.clean_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.select_all_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.deselect_all_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.expand_all_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.collapse_all_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.log(f"扫描完成！发现 {len(self.scan_results)} 个可清理项"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"扫描失败: {str(e)}"))
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.scan_button.config(state=tk.NORMAL))

    def display_results(self):
        total_size = 0
        categories = {}

        for item_id, data in self.scan_results.items():
            category = data["category"]
            size = data["size"]
            total_size += size

            if category not in categories:
                categories[category] = []
            categories[category].append({
                "id": item_id,
                "path": data["path"],
                "size": size
            })

        for category, items in categories.items():
            category_size = sum(item["size"] for item in items)
            category_node = self.tree.insert("", tk.END, text=category, values=(
                f"{len(items)} 项",
                self.format_size(category_size),
                "是"
            ), open=True)
            self.category_nodes[category] = category_node

            for item in items:
                item_node = self.tree.insert(category_node, tk.END, text="", values=(
                    item["path"],
                    self.format_size(item["size"]),
                    "是"
                ))
                self.item_nodes[item["id"]] = item_node

        self.log(f"总共可释放空间: {self.format_size(total_size)}")

    def start_clean(self):
        selected_items = []
        for item_id, item_node in self.item_nodes.items():
            if self.tree.set(item_node, "selected") == "是":
                path = self.tree.set(item_node, "path")
                selected_items.append(path)

        if not selected_items:
            messagebox.showwarning("警告", "请选择要清理的项目")
            return

        confirm = messagebox.askyesno(
            "确认清理",
            f"确定要清理 {len(selected_items)} 个项目吗？\n此操作不可恢复！"
        )

        if confirm:
            self.clean_button.config(state=tk.DISABLED)
            self.scan_button.config(state=tk.DISABLED)
            self.select_all_button.config(state=tk.DISABLED)
            self.deselect_all_button.config(state=tk.DISABLED)
            self.expand_all_button.config(state=tk.DISABLED)
            self.collapse_all_button.config(state=tk.DISABLED)
            self.progress.start()
            self.status_label.config(text="正在清理...")
            self.log(f"开始清理 {len(selected_items)} 个项目...")

            thread = threading.Thread(target=self.clean, args=(selected_items,))
            thread.start()

    def clean(self, selected_items):
        try:
            jianying_categories = [
                "剪映如下缓存内容", "剪映音频缓存", "剪映工作平台缓存",
                "剪映艺术特效缓存", "剪映临时文件", "剪映预览缓存", "剪映导出缓存"
            ]

            has_jianying_cache = False
            for item in selected_items:
                for item_id, item_node in self.item_nodes.items():
                    if self.tree.set(item_node, "path") == item:
                        category = self.tree.set(item_node, "category")
                        if category in jianying_categories:
                            has_jianying_cache = True
                            break
                if has_jianying_cache:
                    break

            success_count, total_size = self.cleaner.clean_items(selected_items)

            self.root.after(0, lambda: self.status_label.config(text=f"清理完成！释放 {self.format_size(total_size)}"))
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.scan_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.log(f"清理完成！成功清理 {success_count} 个项目，释放 {self.format_size(total_size)}"))

            message = f"清理完成！\n成功清理 {success_count} 个项目\n释放空间: {self.format_size(total_size)}"

            if has_jianying_cache:
                message += "\n\n注意：清理剪映缓存后，贴纸、特效等素材可能需要重新下载。"

            self.root.after(0, lambda: messagebox.showinfo("完成", message))

            for item in self.tree.get_children():
                self.tree.delete(item)

            self.category_nodes = {}
            self.item_nodes = {}

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"清理失败: {str(e)}"))
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.scan_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.clean_button.config(state=tk.NORMAL))

    @staticmethod
    def format_size(size):
        for unit in ["B", "KB", "MB", "GB"]:
            if size < 1024.0:
                return f"{size:.2f} {unit}"
            size /= 1024.0
        return f"{size:.2f} TB"


def main():
    root = tk.Tk()
    app = DiskCleanerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
