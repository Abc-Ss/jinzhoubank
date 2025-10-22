# main.py
import os
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from utils import LocalOffer,LocalReply,OtherOffer,OtherReply


class App(tk.Tk):
    NAV_NAMES = ['本行报盘', '本行回盘', '他行报盘', '他行回盘']

    def __init__(self):
        super().__init__()
        self.title('文件处理工具')
        self.geometry('700x450')

        # 顶部导航
        nav_frame = tk.Frame(self, bg='#f0f0f0')
        nav_frame.pack(fill='x')
        self.nav_buttons = {}
        for name in self.NAV_NAMES:
            btn = tk.Button(nav_frame, text=name, width=12, height=2,
                            command=lambda n=name: self.show_page(n))
            btn.pack(side='left', padx=5, pady=5)
            self.nav_buttons[name] = btn

        # 页面容器
        self.container = tk.Frame(self)
        self.container.pack(fill='both', expand=True)
        self.frames = {}
        for F in (LocalOfferPage, LocalReplyPage, OtherOfferPage, OtherReplyPage):
            page_name = F.page_name
            frame = F(parent=self.container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky='nsew')
        self.show_page('本行报盘')

    # 切换页面：仅 raise + 高亮，不清空
    def show_page(self, page_name: str):
        frame = self.frames[page_name]
        frame.tkraise()
        for btn in self.nav_buttons.values():
            btn.config(bg='#f0f0f0', fg='black')
        self.nav_buttons[page_name].config(bg='#005fb8', fg='white')


# -------------------- 页面基类（含所有公共逻辑） --------------------
class _BasePage(tk.Frame):
    """统一提供：状态栏、建行、浏览、秒级时间戳、异常处理、手动清除"""
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        # 底部状态栏
        self.status_label = tk.Label(
            self, text='', font=12, fg='green', anchor='w', justify='left',
            wraplength=650, width=90, height=5
        )
        self.status_label.pack(side='bottom', fill='x', padx=15, pady=5)

        # 子类只需把 StringVar 填到这里即可
        self.path_vars = []

    # 子类返回需要清空的 StringVar 列表
    def get_path_vars(self):
        return self.path_vars

    # 一键清除：路径 + 提示
    def clear_record(self):
        for var in self.get_path_vars():
            var.set('')
        self.display_status('')

    # 统一状态显示
    def display_status(self, msg, success=True):
        self.status_label.config(text=msg, fg='green' if success else 'red')

    # 公共：构造「标签 + 输入框 + 浏览」一行
    def build_row(self, label, var, file_type):
        row = tk.Frame(self)
        row.pack(pady=8)
        tk.Label(row, text=label, width=10).pack(side='left')
        tk.Entry(row, textvariable=var, width=50).pack(side='left', padx=5)
        tk.Button(row, text='浏览',
                  command=lambda: self.browse_file(var, file_type)).pack(side='left')

    def browse_file(self, entry: tk.StringVar, file_type: str):
        ftypes = [('Text', '*.txt')] if file_type == 'txt' else \
                 [('Excel', '*.xls *.xlsx')] if file_type == 'excel' else \
                 [('All', '*.*')]
        path = filedialog.askopenfilename(filetypes=ftypes)
        if path:
            entry.set(path)

    # 统一处理成功/失败提示（带秒级时间）
    def run_job(self, job_func, success_msg):
        self.display_status('')          # 先清旧提示
        try:
            out_path = job_func()
            now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.display_status(f'{success_msg}（{now}），文件已保存至 {out_path}', True)
            return out_path
        except Exception as e:
            now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.display_status(f'操作失败（{now}）', False)
            messagebox.showerror('错误', str(e))


# -------------------- 1. 本行报盘 --------------------
class LocalOfferPage(_BasePage):
    page_name = '本行报盘'

    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        tk.Label(self, text='报盘TXT → 报盘Excel', font=('Microsoft YaHei', 14)).pack(pady=20)
        self.txt_path = tk.StringVar()
        self.path_vars = [self.txt_path]
        self.build_row('选择TXT:', self.txt_path, 'txt')

        # 按钮区
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=20)
        tk.Button(btn_frame, text='开始转换', width=15, command=self.convert).pack(side='left', padx=5)
        tk.Button(btn_frame, text='清除记录', width=15, command=self.clear_record).pack(side='left', padx=5)

    # 业务逻辑仅返回生成路径，其余交给基类
    def convert(self):
        if not self.txt_path.get():
            messagebox.showwarning('提示', '请先选择文件')
            return
        def job():
            out_path = os.path.splitext(self.txt_path.get())[0] + '_报盘结果.xls'
            LocalOffer(self.txt_path.get(), out_path)
            return out_path
        self.run_job(job, '操作成功')


# -------------------- 2. 本行回盘 --------------------
class LocalReplyPage(_BasePage):
    page_name = '本行回盘'

    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        tk.Label(self, text='报盘TXT + 回盘Excel → 回盘TXT', font=('Microsoft YaHei', 14)).pack(pady=20)
        self.txt_report = tk.StringVar()
        self.xls_reply = tk.StringVar()
        self.path_vars = [self.txt_report, self.xls_reply]
        self.build_row('报盘TXT:', self.txt_report, 'txt')
        self.build_row('回盘Excel:', self.xls_reply, 'excel')

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=20)
        tk.Button(btn_frame, text='开始处理', width=15, command=self.process).pack(side='left', padx=5)
        tk.Button(btn_frame, text='清除记录', width=15, command=self.clear_record).pack(side='left', padx=5)

    def process(self):
        if not self.txt_report.get() or not self.xls_reply.get():
            messagebox.showwarning('提示', '请先选择两个文件')
            return
        def job():
            out_path = os.path.splitext(self.txt_report.get())[0] + '_回盘结果.txt'
            LocalReply(self.txt_report.get(), self.xls_reply.get(), out_path)
            return out_path
        self.run_job(job, '操作成功')


# -------------------- 3. 他行报盘 --------------------
class OtherOfferPage(_BasePage):
    page_name = '他行报盘'

    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        tk.Label(self, text='报盘TXT → 报盘Excel', font=('Microsoft YaHei', 14)).pack(pady=20)
        self.txt_path = tk.StringVar()
        self.path_vars = [self.txt_path]
        self.build_row('选择TXT:', self.txt_path, 'txt')

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=20)
        tk.Button(btn_frame, text='开始转换', width=15, command=self.convert).pack(side='left', padx=5)
        tk.Button(btn_frame, text='清除记录', width=15, command=self.clear_record).pack(side='left', padx=5)

    def convert(self):
        if not self.txt_path.get():
            messagebox.showwarning('提示', '请先选择文件')
            return
        def job():
            out_path = os.path.splitext(self.txt_path.get())[0] + '_报盘结果.xls'
            OtherOffer(self.txt_path.get(), out_path)
            return out_path
        self.run_job(job, '操作成功')


# -------------------- 4. 他行回盘 --------------------
class OtherReplyPage(_BasePage):
    page_name = '他行回盘'

    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        tk.Label(self, text='报盘TXT + 回盘Excel → 回盘', font=('Microsoft YaHei', 14)).pack(pady=20)
        self.txt_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.path_vars = [self.txt_path, self.excel_path]
        self.build_row('报盘TXT:', self.txt_path, 'txt')
        self.build_row('回盘Excel:', self.excel_path, 'excel')

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=20)
        tk.Button(btn_frame, text='开始处理', width=15, command=self.process).pack(side='left', padx=5)
        tk.Button(btn_frame, text='清除记录', width=15, command=self.clear_record).pack(side='left', padx=5)

    def process(self):
        if not self.txt_path.get() or not self.excel_path.get():
            messagebox.showwarning('提示', '请先选择两个文件')
            return
        def job():
            out_path = os.path.splitext(self.txt_path.get())[0] + '_回盘结果.txt'
            OtherReply(self.txt_path.get(), self.excel_path.get(), out_path)
            return out_path
        self.run_job(job, '操作成功')


# ============================================================
if __name__ == '__main__':
    App().mainloop()