import re
import sys
import tkinter as tk
import unicodedata
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from bank_excel import default_library, load_excel, read_banks, save_bank, write_template


def enable_high_dpi():
    """让 Windows 按实际屏幕像素绘制界面，避免系统拉伸造成模糊。"""
    if sys.platform != 'win32':
        return
    import ctypes
    try:
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except (AttributeError, OSError):
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass


enable_high_dpi()

ROOT = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))
SAMPLE_IMAGE = ROOT / 'docs' / 'excel-import.png'


def normalized(value):
    return ''.join(c for c in unicodedata.normalize('NFKC', str(value)).lower()
                   if not c.isspace() and not unicodedata.category(c).startswith(('P', 'S')) and c != '_')


def subsequence_span(text, term):
    start = text.find(term[0])
    best = None
    while start >= 0:
        end = start
        for char in term[1:]:
            end = text.find(char, end + 1)
            if end < 0:
                return best
        span = end - start + 1
        best = span if best is None else min(best, span)
        start = text.find(term[0], start + 1)
    return best


def search(questions, query):
    terms = [normalized(part) for part in re.split(r'[\s,，、;；]+', query.strip())]
    terms = [term for term in terms if term]
    if not terms:
        return []
    ranked = []
    phrase = ''.join(terms)
    for q in questions:
        title = q['_search']
        if phrase in title:
            ranked.append((0 if title.startswith(phrase) else 1, title.find(phrase), q))
        elif len(terms) > 1 and all(term in title for term in terms):
            ranked.append((2, sum(title.find(term) for term in terms), q))
        elif len(terms) == 1 and len(phrase) >= 2:
            span = subsequence_span(title, phrase)
            if span is not None:
                ranked.append((3, span, q))
    return [q for _, _, q in sorted(ranked, key=lambda match: (match[0], match[1]))]


class Application(tk.Tk):
    def __init__(self, library=None):
        super().__init__()
        self.questions = []
        self.library = library if library is not None else default_library()
        self.title('题库速查')
        self.geometry('1040x760')
        self.minsize(650, 450)
        self.configure(bg='#f3f6fa')
        self.query = tk.StringVar()
        self.after_id = None
        self.render_id = 0
        top = ttk.Frame(self, padding=(18, 16, 18, 8))
        top.pack(fill='x')
        ttk.Label(top, text='题库速查', font=('Microsoft YaHei UI', 19, 'bold')).pack(anchor='w')
        bank_row = ttk.Frame(top)
        bank_row.pack(fill='x', pady=(5, 10))
        ttk.Label(bank_row, text='当前题库').pack(side='left')
        self.bank_picker = ttk.Combobox(bank_row, state='readonly', width=23)
        self.bank_picker.pack(side='left', padx=8)
        self.bank_picker.bind('<<ComboboxSelected>>', self.switch_bank)
        ttk.Button(bank_row, text='导入 Excel', command=self.import_excel).pack(side='left')
        ttk.Button(bank_row, text='导入格式', command=self.show_import_format).pack(side='left', padx=(6, 0))
        self.delete_button = ttk.Button(bank_row, text='删除题库', command=self.delete_bank)
        self.delete_button.pack(side='left', padx=(6, 0))
        self.bank_count = ttk.Label(bank_row)
        self.bank_count.pack(side='left', padx=8)
        row = ttk.Frame(top)
        row.pack(fill='x')
        self.entry = ttk.Entry(row, textvariable=self.query, font=('Microsoft YaHei UI', 15))
        self.entry.pack(side='left', fill='x', expand=True, ipady=6)
        ttk.Button(row, text='清空', command=lambda: self.query.set('')).pack(side='left', padx=(8, 8), ipady=4)
        self.status = ttk.Label(top, text='输入题目的几个字即可搜索；多个关键词可用空格隔开。', font=('Microsoft YaHei UI', 10))
        self.status.pack(anchor='w', pady=(10, 0))
        ttk.Label(top, text='选项可能随机排序，请优先核对答案文字。', foreground='#a23a21').pack(anchor='w', pady=(3, 0))
        body = ttk.Frame(self, padding=(18, 0, 18, 16))
        body.pack(fill='both', expand=True)
        self.view = tk.Text(body, wrap='word', font=('Microsoft YaHei UI', 12), bg='white', fg='#17243b',
                            relief='flat', padx=18, pady=15, spacing1=3, spacing3=2, cursor='arrow')
        scroll = ttk.Scrollbar(body, orient='vertical', command=self.view.yview)
        self.view.configure(yscrollcommand=scroll.set)
        scroll.pack(side='right', fill='y')
        self.view.pack(side='left', fill='both', expand=True)
        self.view.tag_configure('title', font=('Microsoft YaHei UI', 13, 'bold'), foreground='#17345d', spacing1=8)
        self.view.tag_configure('meta', font=('Microsoft YaHei UI', 10), foreground='#62748a')
        self.view.tag_configure('answer', font=('Microsoft YaHei UI', 12, 'bold'), foreground='#137344', spacing3=10)
        self.view.tag_configure('correct', foreground='#137344')
        self.view.tag_configure('separator', foreground='#b6c4d6', spacing1=4, spacing3=5)
        self.view.configure(state='disabled')
        self.query.trace_add('write', self.schedule_search)
        self.bind('<Control-f>', self.focus_search)
        self.bind('<Escape>', lambda event: self.query.set(''))
        self.refresh_banks()
        self.entry.focus_set()

    def refresh_banks(self, selected=None):
        self.banks = read_banks(self.library)
        labels = []
        for _, name, _ in self.banks:
            label = name
            suffix = 2
            while label in labels:
                label = f'{name} ({suffix})'
                suffix += 1
            labels.append(label)
        self.bank_picker.configure(values=labels, state='readonly' if labels else 'disabled')
        if labels:
            selected_index = next((i for i, (key, _, _) in enumerate(self.banks) if key == selected), 0)
            self.bank_picker.current(selected_index)
        else:
            self.bank_picker.set('暂无题库')
        self.switch_bank()

    def switch_bank(self, event=None):
        index = self.bank_picker.current()
        self.questions = self.banks[index][2] if index >= 0 else []
        self.delete_button.configure(state='normal' if index >= 0 else 'disabled')
        for question in self.questions:
            if '_search' not in question:
                question['_search'] = normalized(question['title'])
        self.bank_count.configure(text=f'{len(self.questions)} 题')
        self.schedule_search()

    def delete_bank(self):
        index = self.bank_picker.current()
        if index < 0:
            return
        path, _, _ = self.banks[index]
        label = self.bank_picker.get()
        if not messagebox.askyesno('删除题库', f'删除「{label}」的导入副本？\n如有需要，可从原 Excel 再次导入。', parent=self):
            return
        try:
            path.unlink(missing_ok=True)
        except OSError as exc:
            messagebox.showerror('删除失败', str(exc), parent=self)
            return
        self.refresh_banks()
        self.entry.focus_set()

    def import_excel(self):
        filename = filedialog.askopenfilename(parent=self, title='导入题库', filetypes=[('Excel 工作簿', '*.xlsx')])
        if not filename:
            return
        try:
            count = self.import_bank(Path(filename))
        except Exception as exc:
            messagebox.showerror('导入失败', f'{exc}\n\n请使用 .xlsx 文件，首行为题目、选项和答案等列名。', parent=self)
            return
        messagebox.showinfo('导入完成', f'已导入 {count} 道题，并切换到新题库。', parent=self)
        self.entry.focus_set()

    def show_import_format(self):
        from PIL import Image, ImageTk

        window = tk.Toplevel(self)
        window.title('Excel 导入格式')
        window.transient(self)
        body = ttk.Frame(window, padding=16)
        body.pack(fill='both', expand=True)
        width = min(1100, max(400, self.winfo_screenwidth() - 130))
        ttk.Label(body, text='Excel 导入示例', font=('Microsoft YaHei UI', 17, 'bold')).pack(anchor='w')
        ttk.Label(body, text='使用 .xlsx 文件。第一行写列名，每行填写一道题。').pack(anchor='w', pady=(3, 9))
        with Image.open(SAMPLE_IMAGE) as picture:
            picture.thumbnail((width, 280), Image.Resampling.LANCZOS)
            window.example_image = ImageTk.PhotoImage(picture, master=window)
        ttk.Label(body, image=window.example_image).pack(anchor='w')
        ttk.Label(body, text=(
            '题目与答案必填。单选填一个字母，多选把所有正确选项写在同一格，例如 A、C、F 或 ACF。\n'
            '判断题可在选项中填“正确”“错误”，答案写对应字母。填空题留空选项，直接填写答案。\n'
            '题型、知识点为可选列。选项最多支持 A 到 F，也可在“选项”一列用 | 分隔。'
        ), wraplength=width, justify='left').pack(anchor='w', pady=(9, 13))
        actions = ttk.Frame(body)
        actions.pack(fill='x')
        ttk.Button(actions, text='另存空白模板', command=lambda: self.save_template(window)).pack(side='left')
        ttk.Button(actions, text='关闭', command=window.destroy).pack(side='right')
        window.bind('<Escape>', lambda event: window.destroy())
        return window

    def save_template(self, window):
        filename = filedialog.asksaveasfilename(parent=window, title='保存空白模板',
            initialfile='题库导入模板.xlsx', defaultextension='.xlsx', filetypes=[('Excel 工作簿', '*.xlsx')])
        if filename:
            try:
                write_template(filename)
            except OSError as exc:
                messagebox.showerror('保存模板失败', str(exc), parent=window)

    def import_bank(self, path):
        questions = load_excel(path)
        key = save_bank(self.library, path, questions)
        self.refresh_banks(key)
        return len(questions)

    def focus_search(self, event=None):
        self.entry.focus_set()
        self.entry.selection_range(0, 'end')
        return 'break'

    def schedule_search(self, *args):
        if self.after_id:
            self.after_cancel(self.after_id)
        self.after_id = self.after(80, self.update_results)

    def update_results(self):
        self.after_id = None
        self.render_id += 1
        version = self.render_id
        matches = search(self.questions, self.query.get())
        self.view.configure(state='normal')
        self.view.delete('1.0', 'end')
        if not self.banks:
            self.status.configure(text='暂无题库，请先导入 Excel。')
            self.view.insert('end', '点「导入 Excel」选择题库。点「导入格式」查看示例或另存空白模板。', 'meta')
            self.view.configure(state='disabled')
            return
        if not normalized(self.query.get()):
            self.status.configure(text='输入题干片段。跳字可直接输入，多个关键词可用空格隔开。')
            self.view.insert('end', '快捷键：Ctrl+F 聚焦搜索框，Esc 清空。', 'meta')
            self.view.configure(state='disabled')
            return
        self.status.configure(text=f'找到 {len(matches)} 道题' + ('，连续匹配优先。' if matches else '，可尝试减少字数或换用题目中的其他词。'))
        if not matches:
            self.view.insert('end', '没有找到匹配题目。请检查字词，或输入更短的题干片段。', 'meta')
        self.view.configure(state='disabled')
        self.render_batch(matches, 0, version)

    def render_batch(self, matches, start, version):
        if version != self.render_id:
            return
        self.view.configure(state='normal')
        for i, q in enumerate(matches[start:start + 40], start + 1):
            self.view.insert('end', f'{i}. {q["title"]}\n', 'title')
            self.view.insert('end', f'{q["kind"]}  ·  {q["category"]}\n', 'meta')
            for opt in q['options']:
                self.view.insert('end', f'{opt["letter"]}. {opt["text"]}\n', 'correct' if opt.get('correct') else ())
            answer_text = '；'.join(q['answer']) if q['answer'] else '题库未提供可识别的答案文字'
            self.view.insert('end', f'答案：{q["letters"]}  {answer_text}\n', 'answer')
            self.view.insert('end', '─' * 45 + '\n', 'separator')
        self.view.configure(state='disabled')
        if start + 40 < len(matches):
            self.after(1, lambda: self.render_batch(matches, start + 40, version))


if __name__ == '__main__':
    try:
        Application().mainloop()
    except Exception as exc:
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror('题库速查无法启动', f'{exc}\n\n请确保程序文件夹完整，且已安装 Python（含 tkinter）。')
            root.destroy()
        except Exception:
            print(exc, file=sys.stderr)
