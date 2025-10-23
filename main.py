import os
import datetime
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox,scrolledtext,ttk
from utils import LocalOffer, LocalReply, OtherOffer, OtherReply


class App(tk.Tk):
    NAV_NAMES = ['本行报盘', '本行回盘', '他行报盘', '他行回盘']
    # 优化配色方案（更现代柔和）
    COLORS = {
        'bg': '#f0f2f5',  # 页面背景（浅灰）
        'nav_bg': '#ffffff',  # 导航栏背景（纯白）
        'active_nav': '#165dff',  # 选中导航按钮（品牌蓝）
        'btn_normal': '#e6e6e6',  # 普通按钮背景（白）
        'btn_hover': '#f0f5ff',  # 按钮悬停（浅蓝灰）
        'btn_active': '#e6f7ff',  # 按钮点击（更浅蓝）
        'text': '#1f2329',  # 文本颜色（深灰）
        'text_light': '#6e7072',  # 次要文本（浅灰）
        'border': '#e5e6eb',  # 边框颜色（淡灰）
        'status_bg': '#ffffff',  # 状态栏背景（白）
        'success': '#00b42a',  # 成功色（绿）
        'error': '#f53f3f',  # 错误色（红）
    }

    def __init__(self):
        super().__init__()
        self.title('文件处理工具')
        self.geometry('850x520')  # 适度放大窗口
        self.minsize(800, 550)  # 最小尺寸限制
        self.configure(bg=self.COLORS['bg'])

        # 顶部导航栏（带阴影效果）
        nav_frame = tk.Frame(
            self,
            bg=self.COLORS['nav_bg'],
            height=60,
            relief=tk.FLAT,
            bd=0,
            highlightbackground=self.COLORS['border'],
            highlightthickness=1
        )
        nav_frame.pack(fill='x', padx=0, pady=0)
        nav_frame.pack_propagate(False)  # 固定导航栏高度

        # 导航按钮布局（居中）
        nav_inner = tk.Frame(nav_frame, bg=self.COLORS['nav_bg'])
        nav_inner.place(relx=0.5, rely=0.5, anchor='center')

        self.nav_buttons = {}
        for i, name in enumerate(self.NAV_NAMES):
            btn = tk.Button(
                nav_inner, text=name,
                width=12, height=2,
                bg=self.COLORS['btn_normal'],
                fg=self.COLORS['text'],
                font=('Microsoft YaHei', 10, 'bold'),
                relief=tk.FLAT,
                bd=0,
                command=lambda n=name: self.show_page(n)
            )
            btn.grid(row=0, column=i, padx=2, pady=5)  # 减小间距更紧凑
            self.nav_buttons[name] = btn
            # 添加悬停/点击效果
            self.bind_button_events(btn)

        # 页面容器（卡片式设计）
        self.container = tk.Frame(
            self,
            bg=self.COLORS['bg'],
            relief=tk.FLAT,
            bd=0,
            padx=20,
            pady=15
        )
        self.container.pack(fill='both', expand=True)

        # 初始化页面（每个页面用卡片包裹）
        self.frames = {}
        for F in (LocalOfferPage, LocalReplyPage, OtherOfferPage, OtherReplyPage):
            page_name = F.page_name
            # 页面外层卡片
            card = tk.Frame(
                self.container,
                bg=self.COLORS['nav_bg'],
                relief=tk.FLAT,
                bd=0,
                highlightbackground=self.COLORS['border'],
                highlightthickness=1,
                padx=10,
                pady=10
            )
            card.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)
            # 页面内容
            frame = F(parent=card, controller=self)
            frame.pack(fill='both', expand=True)
            self.frames[page_name] = card

        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)
        self.show_page('本行报盘')

    # 按钮事件绑定（悬停+点击）
    def bind_button_events(self, widget):
        def on_enter(e):
            if widget['bg'] != self.COLORS['active_nav']:
                widget['bg'] = self.COLORS['btn_hover']

        def on_leave(e):
            if widget['bg'] != self.COLORS['active_nav']:
                widget['bg'] = self.COLORS['btn_normal']

        def on_press(e):
            if widget['bg'] != self.COLORS['active_nav']:
                widget['bg'] = self.COLORS['btn_active']

        def on_release(e):
            if widget['bg'] != self.COLORS['active_nav']:
                widget['bg'] = self.COLORS['btn_hover']

        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)
        widget.bind("<ButtonPress-1>", on_press)
        widget.bind("<ButtonRelease-1>", on_release)

    # 切换页面：优化过渡效果
    def show_page(self, page_name: str):
        frame = self.frames[page_name]
        frame.tkraise()
        # 导航按钮样式切换
        for name, btn in self.nav_buttons.items():
            if name == page_name:
                btn.config(
                    bg=self.COLORS['active_nav'],
                    fg='white',
                    relief=tk.FLAT
                )
            else:
                btn.config(
                    bg=self.COLORS['btn_normal'],
                    fg=self.COLORS['text']
                )


# -------------------- 页面基类（优化样式） --------------------
class _BasePage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.configure(bg=controller.COLORS['nav_bg'])  # 继承卡片背景

        # 内容容器（增加内边距）
        self.content = tk.Frame(self, bg=controller.COLORS['nav_bg'])
        self.content.pack(expand=True, fill='both', padx=30, pady=25)

        # 标题区（带分割线）
        title_frame = tk.Frame(self.content, bg=controller.COLORS['nav_bg'])
        title_frame.pack(fill='x', pady=(0, 20))
        self.title_label = tk.Label(
            title_frame,
            text='',  # 子类设置具体标题
            font=('Microsoft YaHei', 16, 'bold'),
            bg=controller.COLORS['nav_bg'],
            fg=self.controller.COLORS['text']
        )
        self.title_label.pack(anchor='w')
        # 标题下方分割线
        tk.Frame(
            title_frame,
            height=2,
            bg=self.controller.COLORS['active_nav'],
            relief=tk.FLAT
        ).pack(fill='x', pady=(8, 0))

        # 输入区容器
        self.input_frame = tk.Frame(self.content, bg=controller.COLORS['nav_bg'])
        self.input_frame.pack(fill='x', pady=10)

        # 按钮区容器
        self.btn_frame = tk.Frame(self.content, bg=controller.COLORS['nav_bg'])
        self.btn_frame.pack(fill='x', pady=25)

        # 底部状态栏（卡片式）
        self.status_frame = tk.Frame(
            self.content,
            bg=self.controller.COLORS['status_bg'],
            relief=tk.FLAT,
            bd=1,
            highlightbackground=self.controller.COLORS['border'],
            highlightthickness=1,
            # height=80
        )
        self.status_frame.pack(fill='both', expand=True, pady=(10, 0))
        # 行/列权重，保证 Text 区域占满
        self.status_frame.rowconfigure(0, weight=1)
        self.status_frame.columnconfigure(0, weight=1)

        self.status_text = scrolledtext.ScrolledText(
            self.status_frame,
            wrap='word',
            font=('Microsoft YaHei', 10),
            fg=self.controller.COLORS['text_light'],
            bg=self.controller.COLORS['status_bg'],
            padx=15,
            pady=8,
            height=5,  # 字符行数
            borderwidth=0,
            highlightthickness=0,
            state='disabled'
        )
        self.status_text.grid(row=0, column=0, sticky='nsew')

        # 右下角 Sizegrip（小三角拖柄）
        ttk.Sizegrip(self.status_frame).grid(row=0, column=0, sticky='se')

        self.path_vars = []

    def get_path_vars(self):
        return self.path_vars

    def clear_record(self):
        for var in self.get_path_vars():
            var.set('')
        self.display_status('请选择文件并执行操作', success=None)

    def display_status(self, msg, success=True):
        # 1. 先解锁
        self.status_text.config(state='normal')
        # 2. 清空旧内容
        self.status_text.delete('1.0', 'end')
        # 3. 写入新消息
        self.status_text.insert('end', msg)
        # 4. 设置颜色
        if success is None:
            color = self.controller.COLORS['text_light']
        else:
            color = (self.controller.COLORS['success'] if success
                     else self.controller.COLORS['error'])
        # Text 控件没有 fg，所以用 tag 给整行上色
        self.status_text.tag_config('whole', foreground=color)
        self.status_text.tag_add('whole', '1.0', 'end')
        # 5. 重新设为只读并滚到底
        self.status_text.config(state='disabled')
        self.status_text.see('end')

    # 优化输入行样式（带图标感）
    def build_row(self, label, var, file_type):
        row = tk.Frame(self.input_frame, bg=self.controller.COLORS['nav_bg'], height=45)
        row.pack(fill='x', pady=8)
        row.pack_propagate(False)

        # 标签（带图标占位符）
        tk.Label(
            row,
            text=label,
            width=12,
            font=('Microsoft YaHei', 10),
            bg=self.controller.COLORS['nav_bg'],
            fg=self.controller.COLORS['text'],
            anchor='e'
        ).pack(side='left', padx=(0, 12))

        # 输入框（带焦点效果）
        entry = tk.Entry(
            row,
            textvariable=var,
            width=50,
            font=('Microsoft YaHei', 10),
            relief=tk.FLAT,
            bd=1,
            highlightthickness=2,
            highlightbackground=self.controller.COLORS['border'],
            highlightcolor=self.controller.COLORS['active_nav'],  # 焦点时边框色
            bg='#f7f8fa'  # 输入框背景色
        )
        entry.pack(side='left', fill='x', expand=True, padx=(0, 12))
        # 输入框获得焦点时的样式
        entry.bind("<FocusIn>", lambda e: e.widget.config(bg='white'))
        entry.bind("<FocusOut>", lambda e: e.widget.config(bg='#f7f8fa'))

        # 浏览按钮（美化样式）
        btn = tk.Button(
            row,
            text='浏览',
            width=8,
            font=('Microsoft YaHei', 10),
            bg=self.controller.COLORS['btn_normal'],
            fg=self.controller.COLORS['text'],
            relief=tk.FLAT,
            bd=1,
            highlightbackground=self.controller.COLORS['border'],
            highlightthickness=1,
            command=lambda: self.browse_file(var, file_type)
        )
        btn.pack(side='left')
        self.controller.bind_button_events(btn)

    def browse_file(self, entry: tk.StringVar, file_type: str):
        ftypes = [('Text文件', '*.txt')] if file_type == 'txt' else \
            [('Excel文件', '*.xls *.xlsx')] if file_type == 'excel' else \
                [('所有文件', '*.*')]
        path = filedialog.askopenfilename(
            title=f'选择{file_type}文件',
            filetypes=ftypes
        )
        if path:
            entry.set(path)

    def run_job(self, job_func, success_msg):
        self.display_status('处理中...', success=None)
        try:
            out_path = job_func()
            now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.display_status(
                f'✅ {success_msg}（{now}）\n文件路径：{out_path}',
                True
            )
            return out_path
        except Exception as e:
            now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.display_status(
                f'❌ 操作失败（{now}）\n错误：{str(e)}',
                False
            )
            messagebox.showerror('错误', str(e))


# -------------------- 1. 本行报盘 --------------------
class LocalOfferPage(_BasePage):
    page_name = '本行报盘'

    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.title_label.config(text='报盘TXT → 报盘Excel')

        self.txt_path = tk.StringVar()
        self.path_vars = [self.txt_path]
        self.build_row('选择TXT:', self.txt_path, 'txt')

        # 按钮区布局
        self.convert_btn = tk.Button(
            self.btn_frame,
            text='开始转换',
            width=15,
            height=1,
            font=('Microsoft YaHei', 10, 'bold'),
            bg=self.controller.COLORS['active_nav'],
            fg='white',
            relief=tk.FLAT,
            bd=0,
            padx=10,
            command=self.convert
        )
        self.convert_btn.pack(side='left', padx=(120, 20))
        self.controller.bind_button_events(self.convert_btn)

        self.clear_btn = tk.Button(
            self.btn_frame,
            text='清除记录',
            width=15,
            height=1,
            font=('Microsoft YaHei', 10),
            bg=self.controller.COLORS['btn_normal'],
            fg=self.controller.COLORS['text'],
            relief=tk.FLAT,
            bd=1,
            highlightbackground=self.controller.COLORS['border'],
            highlightthickness=1,
            command=self.clear_record
        )
        self.clear_btn.pack(side='left')
        self.controller.bind_button_events(self.clear_btn)

        # 初始状态提示
        self.display_status('请选择TXT文件并点击"开始转换"', success=None)

    def convert(self):
        if not self.txt_path.get():
            messagebox.showwarning('提示', '请先选择TXT文件')
            return
        elif "本行报盘.txt" not in Path(self.txt_path.get()).name:
            messagebox.showwarning('提示', '本行报盘.txt文件选择错误！')
            return

        def job():
            out_path = Path(self.txt_path.get()).with_name("工行本行报盘.xls")
            # out_path = os.path.splitext(self.txt_path.get())[0] + '_报盘结果.xls'
            LocalOffer(self.txt_path.get(), out_path)
            return out_path

        self.run_job(job, '本行报盘文件转换成功')


# -------------------- 2. 本行回盘 --------------------
class LocalReplyPage(_BasePage):
    page_name = '本行回盘'

    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.title_label.config(text='报盘TXT + 回盘Excel → 回盘TXT')

        self.txt_report = tk.StringVar()
        self.xls_reply = tk.StringVar()
        self.path_vars = [self.txt_report, self.xls_reply]
        self.build_row('报盘TXT:', self.txt_report, 'txt')
        self.build_row('回盘Excel:', self.xls_reply, 'excel')

        # 按钮区
        self.process_btn = tk.Button(
            self.btn_frame,
            text='开始处理',
            width=15,
            height=1,
            font=('Microsoft YaHei', 10, 'bold'),
            bg=self.controller.COLORS['active_nav'],
            fg='white',
            relief=tk.FLAT,
            bd=0,
            padx=10,
            command=self.process
        )
        self.process_btn.pack(side='left', padx=(120, 20))
        self.controller.bind_button_events(self.process_btn)

        self.clear_btn = tk.Button(
            self.btn_frame,
            text='清除记录',
            width=15,
            height=1,
            font=('Microsoft YaHei', 10),
            bg=self.controller.COLORS['btn_normal'],
            fg=self.controller.COLORS['text'],
            relief=tk.FLAT,
            bd=1,
            highlightbackground=self.controller.COLORS['border'],
            highlightthickness=1,
            command=self.clear_record
        )
        self.clear_btn.pack(side='left')
        self.controller.bind_button_events(self.clear_btn)

        self.display_status('请选择报盘TXT和回盘Excel并点击"开始处理"', success=None)

    def process(self):
        if not self.txt_report.get() or not self.xls_reply.get():
            messagebox.showwarning('提示', '请先选择两个文件')
            return
        elif "本行报盘.txt" not in Path(self.txt_report.get()).name:
            messagebox.showwarning('提示', '本行报盘.txt文件选择错误！')
            return
        elif "本行回盘.xls" not in Path(self.xls_reply.get()).name:
            messagebox.showwarning('提示', '本行回盘.xls文件选择错误！')
            return

        out_path = Path(self.txt_report.get()).with_name("自来水本行回盘.txt")
        # out_path = os.path.splitext(self.txt_report.get())[0] + '_回盘结果.txt'
        mess, txt_xls, xls_txt = LocalReply(self.txt_report.get(), self.xls_reply.get(), out_path)

        self.display_status('处理中...', success=None)
        if mess == "文件信息一致":
            now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.display_status(
                f'✅ 本行回盘文件转换成功（{now}）\n文件路径：{out_path}',
                True
            )
        else:
            now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.display_status(
                f'❌ 操作失败（{now}）：{mess}\n❌ txt有但xls没有：{txt_xls}\n❌ xls有但txt没有：{xls_txt}',
                False
            )


# -------------------- 3. 他行报盘 --------------------
class OtherOfferPage(_BasePage):
    page_name = '他行报盘'

    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.title_label.config(text='报盘TXT → 报盘Excel')

        self.txt_path = tk.StringVar()
        self.path_vars = [self.txt_path]
        self.build_row('选择TXT:', self.txt_path, 'txt')

        # 按钮区
        self.convert_btn = tk.Button(
            self.btn_frame,
            text='开始转换',
            width=15,
            height=1,
            font=('Microsoft YaHei', 10, 'bold'),
            bg=self.controller.COLORS['active_nav'],
            fg='white',
            relief=tk.FLAT,
            bd=0,
            padx=10,
            command=self.convert
        )
        self.convert_btn.pack(side='left', padx=(120, 20))
        self.controller.bind_button_events(self.convert_btn)

        self.clear_btn = tk.Button(
            self.btn_frame,
            text='清除记录',
            width=15,
            height=1,
            font=('Microsoft YaHei', 10),
            bg=self.controller.COLORS['btn_normal'],
            fg=self.controller.COLORS['text'],
            relief=tk.FLAT,
            bd=1,
            highlightbackground=self.controller.COLORS['border'],
            highlightthickness=1,
            command=self.clear_record
        )
        self.clear_btn.pack(side='left')
        self.controller.bind_button_events(self.clear_btn)

        self.display_status('请选择TXT文件并点击"开始转换"', success=None)

    def convert(self):
        if not self.txt_path.get():
            messagebox.showwarning('提示', '请先选择TXT文件')
            return
        elif "他行报盘.txt" not in Path(self.txt_path.get()).name:
            messagebox.showwarning('提示', '他行报盘.txt文件选择错误！')
            return

        def job():
            out_path = Path(self.txt_path.get()).with_name("工行他行报盘.xls")
            # out_path = os.path.splitext(self.txt_path.get())[0] + '_报盘结果.xls'
            OtherOffer(self.txt_path.get(), out_path)
            return out_path

        self.run_job(job, '他行报盘文件转换成功')


# -------------------- 4. 他行回盘 --------------------
class OtherReplyPage(_BasePage):
    page_name = '他行回盘'

    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.title_label.config(text='报盘TXT + 回盘Excel → 回盘TXT')

        self.txt_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.path_vars = [self.txt_path, self.excel_path]
        self.build_row('报盘TXT:', self.txt_path, 'txt')
        self.build_row('回盘Excel:', self.excel_path, 'excel')

        # 按钮区
        self.process_btn = tk.Button(
            self.btn_frame,
            text='开始处理',
            width=15,
            height=1,
            font=('Microsoft YaHei', 10, 'bold'),
            bg=self.controller.COLORS['active_nav'],
            fg='white',
            relief=tk.FLAT,
            bd=0,
            padx=10,
            command=self.process
        )
        self.process_btn.pack(side='left', padx=(120, 20))
        self.controller.bind_button_events(self.process_btn)

        self.clear_btn = tk.Button(
            self.btn_frame,
            text='清除记录',
            width=15,
            height=1,
            font=('Microsoft YaHei', 10),
            bg=self.controller.COLORS['btn_normal'],
            fg=self.controller.COLORS['text'],
            relief=tk.FLAT,
            bd=1,
            highlightbackground=self.controller.COLORS['border'],
            highlightthickness=1,
            command=self.clear_record
        )
        self.clear_btn.pack(side='left')
        self.controller.bind_button_events(self.clear_btn)

        self.display_status('请选择报盘TXT和回盘Excel并点击"开始处理"', success=None)

    def process(self):
        if not self.txt_path.get() or not self.excel_path.get():
            messagebox.showwarning('提示', '请先选择两个文件')
            return
        elif "他行报盘.txt" not in Path(self.txt_path.get()).name:
            messagebox.showwarning('提示', '他行报盘.txt文件选择错误！')
            return
        elif "他行回盘.xls" not in Path(self.excel_path.get()).name:
            messagebox.showwarning('提示', '他行回盘.xls文件选择错误！')
            return

        out_path = Path(self.txt_path.get()).with_name("自来水他行回盘")
        # out_path = os.path.splitext(self.txt_report.get())[0] + '_回盘结果.txt'
        mess, txt_xls, xls_txt = OtherReply(self.txt_path.get(), self.excel_path.get(), out_path)

        self.display_status('处理中...', success=None)
        if mess == "文件信息一致":
            now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.display_status(
                f'✅ 他行回盘文件转换成功（{now}）\n文件路径：{out_path}',
                True
            )
        else:
            now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.display_status(
                f'❌ 操作失败（{now}）：{mess}\n❌ txt有但xls没有：{txt_xls}\n❌ xls有但txt没有：{xls_txt}',
                False
            )


if __name__ == '__main__':
    App().mainloop()