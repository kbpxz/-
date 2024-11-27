import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import pyperclip
import csv
import re
import logging
from datetime import datetime
from ui.ocr_capture import start_capture
import json
import os
import keyboard
from tkinter import simpledialog

class HelperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("客服小助手")
        self.root.geometry("677x397")
        
        # 配置日志
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            filename='app.log',
            filemode='a'
        )
        self.logger = logging.getLogger(__name__)

        # 界面布局配置
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # 数据管理
        self.current_page = 1
        self.total_pages = 1
        self.all_data = []
        self.filtered_data = []
        self.search_after_id = None
        
        # 快捷键配置
        self.hotkey = 'F1'  # 默认快捷键
        self.load_hotkey()
        self.register_global_hotkey()

        # 设置用户界面
        self.setup_ui()

    def load_hotkey(self):
        """加载快捷键配置"""
        try:
            config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "hotkey.json")
            if os.path.exists(config_path):
                with open(config_path, 'r') as f:
                    config = json.load(f)
                    self.hotkey = config.get('hotkey', 'F1')
        except Exception as e:
            self.logger.error(f"加载快捷键配置失败: {e}")

    def save_hotkey(self):
        """保存快捷键配置"""
        try:
            config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "hotkey.json")
            with open(config_path, 'w') as f:
                json.dump({'hotkey': self.hotkey}, f)
        except Exception as e:
            self.logger.error(f"保存快捷键配置失败: {e}")

    def register_global_hotkey(self):
        """注册全局快捷键"""
        try:
            keyboard.unhook_all()  # 清除所有已注册的快捷键
            keyboard.add_hotkey(self.hotkey, self.on_hotkey_pressed, suppress=True)
        except Exception as e:
            self.logger.error(f"注册快捷键失败: {e}")
            messagebox.showerror("错误", f"注册快捷键失败: {e}")

    def change_hotkey(self):
        """修改快捷键"""
        class HotkeyDialog(simpledialog.Dialog):
            def body(self, master):
                tk.Label(master, text="请按下新的快捷键组合\n(支持组合键如: ctrl+`, alt+1)").pack()
                self.result = None
                return None

            def keypress(self, event):
                if event.keysym == 'Escape':
                    self.cancel()
                    return
                
                key = []
                if event.state & 0x4:
                    key.append('ctrl')
                if event.state & 0x8:
                    key.append('alt')
                if event.state & 0x1:
                    key.append('shift')
                
                if event.keysym != 'Control_L' and event.keysym != 'Alt_L' and event.keysym != 'Shift_L':
                    key.append(event.keysym.lower())
                    self.result = '+'.join(key)
                    self.ok()

            def buttonbox(self):
                self.bind('<Key>', self.keypress)
                box = tk.Frame(self)
                tk.Button(box, text="取消", width=10, command=self.cancel).pack(side=tk.RIGHT, padx=5, pady=5)
                self.bind('<Escape>', lambda e: self.cancel())
                box.pack()

        dialog = HotkeyDialog(self.root, title="设置快捷键")
        if dialog.result:
            try:
                old_hotkey = self.hotkey
                self.hotkey = dialog.result
                self.register_global_hotkey()
                self.save_hotkey()
                self.status_bar.config(text=f"快捷键已更改为: {self.hotkey}")
            except Exception as e:
                self.hotkey = old_hotkey
                self.register_global_hotkey()
                messagebox.showerror("错误", f"快捷键设置失败: {e}")

    def setup_ui(self):
        """设置用户界面"""
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self._create_search_frame(main_frame)
        self.create_treeview(main_frame)
        self._create_pagination_frame(main_frame)

        # 创建状态栏
        status_frame = tk.Frame(self.root)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_bar = tk.Label(status_frame, text=f"当前快捷键: {self.hotkey}", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 添加修改快捷键按钮
        ttk.Button(status_frame, text="修改快捷键", command=self.change_hotkey).pack(side=tk.RIGHT, padx=5)

    def _create_search_frame(self, parent):
        """创建搜索框架"""
        search_frame = ttk.Frame(parent)
        search_frame.pack(fill=tk.X, pady=5)

        ttk.Label(search_frame, text="搜索:").pack(side=tk.LEFT, padx=(0, 5))

        self.search_var = tk.StringVar()
        self.search_var.trace('w', self._on_search_change)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def create_treeview(self, parent):
        """创建表格视图"""
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        # 创建滚动条
        y_scrollbar = ttk.Scrollbar(tree_frame)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        x_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        # 创建表格并设置样式
        style = ttk.Style()
        style.configure("Custom.Treeview", background="white")
        style.map("Custom.Treeview",
                 background=[("selected", "#0078D7")],
                 foreground=[("selected", "white")])

        self.tree = ttk.Treeview(
            tree_frame,
            columns=("客户昵称", "订单号", "商家", "创建时间"),
            show="headings",
            yscrollcommand=y_scrollbar.set,
            xscrollcommand=x_scrollbar.set,
            style="Custom.Treeview"
        )

        # 设置列
        for col in ("客户昵称", "订单号", "商家", "创建时间"):
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_column(c))
            self.tree.column(col, width=150, minwidth=100)

        self.tree.pack(fill=tk.BOTH, expand=True)
        y_scrollbar.config(command=self.tree.yview)
        x_scrollbar.config(command=self.tree.xview)

        # 绑定双击事件和右键菜单
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Button-3>", self.show_context_menu)
        
        # 创建右键菜单
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="复制客户昵称", command=lambda: self.copy_cell_value("客户昵称"))
        self.context_menu.add_command(label="复制订单号", command=lambda: self.copy_cell_value("订单号"))
        self.context_menu.add_command(label="复制商家", command=lambda: self.copy_cell_value("商家"))
        self.context_menu.add_separator()
        self.context_menu.add_command(label="复制整行", command=self.copy_row)

    def on_double_click(self, event):
        """双击处理函数"""
        item = self.tree.selection()[0]
        column = self.tree.identify_column(event.x)
        value = self.tree.item(item)['values'][int(column[1]) - 1]
        pyperclip.copy(str(value))
        self.show_status_message(f"已复制: {value}")

    def show_context_menu(self, event):
        """显示右键菜单"""
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def copy_cell_value(self, column):
        """复制单元格值"""
        selection = self.tree.selection()
        if selection:
            item = selection[0]
            col_idx = self.tree["columns"].index(column)
            value = self.tree.item(item)['values'][col_idx]
            pyperclip.copy(str(value))
            self.show_status_message(f"已复制{column}: {value}")

    def copy_row(self):
        """复制整行数据"""
        selection = self.tree.selection()
        if selection:
            values = self.tree.item(selection[0])['values']
            text = "\t".join(str(v) for v in values)
            pyperclip.copy(text)
            self.show_status_message("已复制整行数据")

    def show_status_message(self, message, duration=3000):
        """显示状态栏消息"""
        self.status_bar.config(text=message)
        self.root.after(duration, lambda: self.status_bar.config(text=f"当前快捷键: {self.hotkey}"))

    def _create_pagination_frame(self, parent):
        """创建分页框架"""
        pagination_frame = ttk.Frame(parent)
        pagination_frame.pack(fill=tk.X, pady=5)

        # 左侧按钮
        ttk.Button(pagination_frame, text="导出数据", command=self.export_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(pagination_frame, text="清空数据", command=self.clear_data).pack(side=tk.LEFT, padx=5)

        # 右侧分页
        jump_frame = ttk.Frame(pagination_frame)
        jump_frame.pack(side=tk.RIGHT)
        
        ttk.Label(jump_frame, text="跳转到:").pack(side=tk.LEFT, padx=2)
        self.page_entry = ttk.Entry(jump_frame, width=5)
        self.page_entry.pack(side=tk.LEFT, padx=2)
        self.page_entry.bind('<Return>', lambda e: self._jump_to_page())
        ttk.Button(jump_frame, text="跳转", command=self._jump_to_page).pack(side=tk.LEFT, padx=2)
        
        self.next_button = ttk.Button(pagination_frame, text="下一页", command=self.next_page)
        self.next_button.pack(side=tk.RIGHT, padx=5)

        self.page_label = ttk.Label(pagination_frame, text=self.get_page_text())
        self.page_label.pack(side=tk.RIGHT, padx=10)

        self.prev_button = ttk.Button(pagination_frame, text="上一页", command=self.prev_page)
        self.prev_button.pack(side=tk.RIGHT, padx=5)

    def _jump_to_page(self):
        """跳转到指定页码"""
        try:
            page = int(self.page_entry.get())
            if 1 <= page <= self.total_pages:
                self.current_page = page
                self.update_treeview()
            else:
                messagebox.showwarning("警告", f"页码必须在1和{self.total_pages}之间")
        except ValueError:
            messagebox.showwarning("警告", "请输入有效的页码")
        self.page_entry.delete(0, tk.END)

    def _on_search_change(self, *args):
        """处理搜索框内容变化（带防抖）"""
        if self.search_after_id:
            self.root.after_cancel(self.search_after_id)
        self.search_after_id = self.root.after(300, self._perform_search)

    def _perform_search(self):
        """执行搜索"""
        search_term = self.search_var.get().strip().lower()
        if search_term:
            self.filtered_data = []
            for item in self.all_data:
                for key, value in item.items():
                    if search_term in str(value).lower():
                        self.filtered_data.append(item)
                        break
        else:
            self.filtered_data = self.all_data.copy()
        
        self.total_pages = max(1, (len(self.filtered_data) + 9) // 10)
        self.current_page = 1
        self.update_treeview()

    def update_treeview(self):
        """更新表格显示"""
        self.tree.delete(*self.tree.get_children())
        start_idx = (self.current_page - 1) * 10
        end_idx = start_idx + 10
        
        search_term = self.search_var.get().strip().lower()
        
        for item in self.filtered_data[start_idx:end_idx]:
            values = (
                item["客户昵称"],
                item["订单号"],
                item["商家"],
                item["创建时间"]
            )
            
            # 创建带有高亮标记的行
            tags = []
            if search_term:
                for value in values:
                    if search_term in str(value).lower():
                        tags.append('search_match')
                        break
            
            self.tree.insert("", "end", values=values, tags=tags)
        
        # 设置高亮样式
        self.tree.tag_configure('search_match', background='yellow')
        
        self.update_page_controls()

    def get_page_text(self):
        """获取页码文本"""
        return f"第 {self.current_page} / {self.total_pages} 页"

    def update_page_controls(self):
        """更新分页控件状态"""
        self.page_label.config(text=self.get_page_text())
        self.prev_button.state(['!disabled'] if self.current_page > 1 else ['disabled'])
        self.next_button.state(['!disabled'] if self.current_page < self.total_pages else ['disabled'])

    def prev_page(self):
        """上一页"""
        if self.current_page > 1:
            self.current_page -= 1
            self.update_treeview()

    def next_page(self):
        """下一页"""
        if self.current_page < self.total_pages:
            self.current_page += 1
            self.update_treeview()

    def sort_column(self, col):
        """排序表格列"""
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
        l.sort(reverse=self.tree.heading(col).get("reverse", False))
        
        for index, (_, k) in enumerate(l):
            self.tree.move(k, "", index)
        
        self.tree.heading(col, reverse=not self.tree.heading(col).get("reverse", False))

    def on_hotkey_pressed(self):
        """全局快捷键处理函数"""
        self.root.deiconify()  # 确保窗口可见
        self.root.focus_force()  # 强制获取焦点
        
        self.show_status_message("正在执行自动识别...")
        self.root.update()
        
        try:
            if start_capture(self):
                self.show_status_message("自动识别完成")
            else:
                self.show_status_message("自动识别失败，请重试")
        except Exception as e:
            self.logger.error(f"自动识别失败: {e}")
            self.show_status_message(f"错误: {str(e)}")
            messagebox.showerror("错误", f"自动识别失败: {str(e)}")

    def add_new_data(self, 客户昵称, 订单号, 商家):
        """添加新数据到表格"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_data = {
            "客户昵称": 客户昵称,
            "订单号": 订单号,
            "商家": 商家,
            "创建时间": current_time
        }
        self.all_data.insert(0, new_data)
        self.filtered_data = self.all_data.copy()
        self.total_pages = max(1, (len(self.filtered_data) + 9) // 10)
        self.current_page = 1
        self.update_treeview()
        self.show_status_message(f"成功添加数据：{客户昵称}")

    def export_data(self):
        """导出数据为CSV文件"""
        if not self.all_data:
            messagebox.showwarning("警告", "没有数据可导出")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.DictWriter(f, fieldnames=["客户昵称", "订单号", "商家", "创建时间"])
                    writer.writeheader()
                    writer.writerows(self.all_data)
                self.show_status_message("数据导出成功")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")

    def clear_data(self):
        """清空所有数据"""
        if messagebox.askyesno("确认", "确定要清空所有数据吗？"):
            self.tree.delete(*self.tree.get_children())
            self.all_data.clear()
            self.filtered_data.clear()
            self.current_page = 1
            self.total_pages = 1
            self.update_page_controls()
            self.show_status_message("已清空所有数据")