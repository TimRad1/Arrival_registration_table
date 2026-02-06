import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from datetime import datetime, timedelta
import json, os
import pandas as pd

DATA_FILE = "data.json"

POSITIONS = ["Неизвестно", "Инженер", "Менеджер", "Техник", "Администратор", "Директор"]
STATUSES = ["Присутствует", "Болен", "Командировка"]
ARRIVAL_TIMES = ["1 час", "1.5 часа", "2 часа", "2.5 часа", "3 часа"]

FONT_SIZE = 20
COMBO_FONT_SIZE = 20
SEARCH_FONT_SIZE = 25

def parse_hours(text):
    return float(text.replace("час", "").replace("а", "").strip())

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("БЕТА-прибытие")
        self.state("zoomed")

        self.start_time = None
        self.rows = {}
        self.clipboard_data = []  # Для копирования/вырезания
        self.current_item = None   # Текущая выделенная строка
        self.edit_widget = None
        self.edit_popup = None
        self.suppress_select = False
        self.last_selection = ()
        self.prev_grab = None
        self.search_matches = []
        self.search_index = -1

        self.setup_style()
        self.create_widgets()
        self.load_data()
        self.sort_table()
        self.update_numbers()
        self.update_counters()

        self.bind("<Control-f>", self.search_dialog)
        self.bind("<Escape>", self.clear_selection)
        #self.bind_all("<Button-1>", self.global_click, add="+")  # Удалено
        self.tree.bind("<MouseWheel>", self.fast_scroll)
        
        # Правый клик: редактирование ячеек со списком (блокируем стандартное выделение)
        self.tree.bind_class("Treeview", "<ButtonPress-3>", lambda e: "break")
        self.tree.bind_class("Treeview", "<ButtonRelease-3>", lambda e: "break")
        self.tree.bind_class("Treeview", "<Button-3>", lambda e: "break")
        self.tree.bind("<ButtonPress-3>", self.on_tree_right_click, add="+")
        self.tree.bind("<ButtonRelease-3>", lambda e: "break", add="+")
        self.tree.bind("<Button-1>", self.on_tree_left_click, add="+")

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_style(self):
        style = ttk.Style(self)
        style.theme_use("default")

        bg = "#1e1e1e"
        fg = "#ffffff"
        field = "#eeeeee"

        self.configure(bg=bg)

        style.configure(".", background=bg, foreground=fg, font=("Segoe UI", FONT_SIZE))
        style.configure("Treeview", font=("Segoe UI", FONT_SIZE), rowheight=40,
                        background="#2a2a2a", fieldbackground="#2a2a2a", foreground="white")
        style.configure("Treeview.Heading", font=("Segoe UI", FONT_SIZE, "bold"),
                        background="#2255aa", foreground="white")
        style.map("Treeview.Heading", background=[("active", "#3366cc")])

        style.configure("TButton", padding=10)
        style.configure("Green.TButton", background="#1e8f3a", foreground="white")
        style.configure("Red.TButton", background="#aa2222", foreground="white")
        style.configure("Blue.TButton", background="#2255aa", foreground="white")
        style.map("TButton", background=[("active", "#555555")])

        style.configure("TEntry", fieldbackground=field, foreground="black")
        style.configure("TCombobox", fieldbackground=field, foreground="black",
                        font=("Segoe UI", COMBO_FONT_SIZE))
        style.configure("Big.TCombobox", fieldbackground=field, foreground="black",
                        font=("Segoe UI", COMBO_FONT_SIZE))

        # Кнопочный стиль для выпадающих списков
        combo_bg = "#1f1f1f"
        combo_border = "#4a4a4a"
        style.configure("Fancy.TCombobox",
                        fieldbackground=combo_bg,
                        foreground="white",
                        background=combo_bg,
                        arrowcolor="white",
                        bordercolor=combo_border,
                        lightcolor=combo_border,
                        darkcolor="#101010",
                        font=("Segoe UI", COMBO_FONT_SIZE),
                        relief="raised",
                        borderwidth=2,
                        padding=6)
        style.map("Fancy.TCombobox",
                  fieldbackground=[("readonly", combo_bg), ("active", combo_bg)],
                  background=[("readonly", combo_bg), ("active", combo_bg)],
                  foreground=[("readonly", "white"), ("active", "white")],
                  bordercolor=[("readonly", combo_border), ("active", combo_border)],
                  lightcolor=[("readonly", combo_border), ("active", combo_border)],
                  darkcolor=[("readonly", "#101010"), ("active", "#101010")])

        self.option_add("*TCombobox*Listbox.font", ("Segoe UI", COMBO_FONT_SIZE))
        self.option_add("*TCombobox*Listbox.background", "#1f1f1f")
        self.option_add("*TCombobox*Listbox.foreground", "white")
        self.option_add("*TCombobox*Listbox.selectBackground", "#2f2f2f")
        self.option_add("*TCombobox*Listbox.selectForeground", "white")

        style.configure("Dialog.TEntry", fieldbackground="#2a2a2a", foreground="white")
        style.configure("Dialog.TCombobox", fieldbackground="#2a2a2a", foreground="white",
                        font=("Segoe UI", COMBO_FONT_SIZE))
        style.configure("Dialog.Fancy.TCombobox",
                        fieldbackground=combo_bg,
                        foreground="white",
                        background=combo_bg,
                        arrowcolor="white",
                        bordercolor=combo_border,
                        lightcolor=combo_border,
                        darkcolor="#101010",
                        font=("Segoe UI", COMBO_FONT_SIZE),
                        relief="raised",
                        borderwidth=2,
                        padding=6)

    def create_widgets(self):
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=5)

        self.start_btn = ttk.Button(top, text="ЗАПУСК", command=self.start_timer,
                                    style="Green.TButton", width=14)
        self.start_btn.pack(side="left", padx=5)

        self.clear_btn = ttk.Button(top, text="ОЧИСТИТЬ ДАННЫЕ", command=self.clear_data,
                                    style="Red.TButton", width=18)
        self.clear_btn.pack(side="left", padx=10)

        self.start_lbl = ttk.Label(top, text="Время запуска: ---",
                                   font=("Segoe UI", FONT_SIZE * 2), foreground="red")
        self.start_lbl.pack(side="left", padx=20)

        self.search_clear_btn = ttk.Button(top, text="X", command=self.clear_search,
                                           style="Red.TButton", width=3)
        self.search_clear_btn.pack(side="right", padx=6)
        self.search_next_btn = ttk.Button(top, text="Далее", command=self.search_next,
                                          style="Blue.TButton", width=8)
        self.search_next_btn.pack(side="right", padx=6)
        self.search_ent = ttk.Entry(top, font=("Segoe UI", SEARCH_FONT_SIZE), width=18)
        self.search_ent.pack(side="right", padx=10)
        self.search_ent.bind("<Return>", self.search_from_entry)
        ttk.Label(top, text="Поиск:").pack(side="right", padx=6)

        main = ttk.Frame(self)
        main.pack(fill="both", expand=True, padx=10, pady=5)

        # Включаем множественное выделение
        self.tree = ttk.Treeview(main, columns=("num","pos","fio","arr","fact","late","status"), 
                                 show="headings", selectmode="extended")

        headers = ["№","Должность","ФИО","Время прибытия","Факт","Опоздание","Статус"]
        for c,t in zip(self.tree["columns"], headers):
            self.tree.heading(c, text=t)

        self.tree.column("num", anchor="center", width=60)
        self.tree.column("pos", anchor="w", width=220)
        self.tree.column("fio", anchor="w", width=320)
        self.tree.column("arr", anchor="center", width=200)
        self.tree.column("fact", anchor="center", width=160)
        self.tree.column("late", anchor="center", width=160)
        self.tree.column("status", anchor="center", width=200)

        self.tree.tag_configure("late", background="#661111")
        self.tree.tag_configure("ontime", background="#114411")
        self.tree.tag_configure("other", background="#112244")

        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Button-1>", self.on_tree_left_click, add="+")
        self.tree.bind("<ButtonRelease-1>", self.on_tree_left_release, add="+")
        self.tree.bind("<ButtonPress-1>", self.on_tree_left_press, add="+")

        scroll = ttk.Scrollbar(main, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="left", fill="y")

        right = ttk.Frame(main)
        right.pack(side="right", fill="y", padx=10)
        right.configure(width=360)

        def label(txt):
            ttk.Label(right, text=txt).pack(anchor="w", pady=(6,0))

        btn_add = ttk.Button(right, text="Добавить", command=self.open_add_dialog)
        btn_add.pack(fill="x", pady=5)

        btn_del = ttk.Button(right, text="Удалить строку", command=self.delete_row)
        btn_del.pack(fill="x")

        ttk.Button(right, text="Загрузка ФИО из Excel", command=self.import_excel).pack(fill="x", pady=5)

        ttk.Separator(right).pack(fill="x", pady=8)

        label("Выгрузка")
        self.export_btn = ttk.Button(right, text="Выбрать", command=self.open_export_menu, style="Blue.TButton")
        self.export_btn.pack(fill="x")

        ttk.Separator(right).pack(fill="x", pady=8)

        self.total_lbl = ttk.Label(right)
        self.present_lbl = ttk.Label(right)
        self.sick_lbl = ttk.Label(right)
        self.trip_lbl = ttk.Label(right)
        self.arrived_lbl = ttk.Label(right)
        self.percent_lbl = ttk.Label(right)

        self.total_lbl.pack(anchor="w")
        self.present_lbl.pack(anchor="w")
        self.sick_lbl.pack(anchor="w")
        self.trip_lbl.pack(anchor="w")
        self.arrived_lbl.pack(anchor="w")
        self.percent_lbl.pack(anchor="w")

        self.chart = tk.Canvas(right, width=320, height=180, bg="#2a2a2a", highlightthickness=0)
        self.chart.pack(fill="x", pady=10)

        ttk.Separator(right).pack(fill="x", pady=15)
        
        self.delete_all_btn = ttk.Button(right, text="УДАЛИТЬ ВСЮ ТАБЛИЦУ", 
                                         command=self.delete_all_table,
                                         style="Blue.TButton")
        self.delete_all_btn.pack(fill="x", side="bottom", pady=(10, 0))

        # Двусторонняя привязка удалена: редактирование теперь в таблице

    # ====== Методы ======
    def fast_scroll(self, event):
        self.tree.yview_scroll(int(-3*(event.delta/120)), "units")

    def format_timedelta(self, td):
        total_minutes = int(td.total_seconds() // 60)
        hours = total_minutes // 60
        minutes = total_minutes % 60
        return f"{hours:02d}:{minutes:02d}"

    def clear_selection(self, event=None):
        self.tree.selection_remove(self.tree.selection())
        self.current_item = None
        if hasattr(self, "search_ent"):
            self.search_ent.selection_clear()

    def clear_search(self):
        if hasattr(self, "search_ent"):
            self.search_ent.delete(0, tk.END)
        self.search_matches = []
        self.search_index = -1
        self.clear_selection()

    def on_tree_right_click(self, event):
        saved_selection = self.tree.selection()
        self.suppress_select = True
        item = self.tree.identify_row(event.y)
        if not item:
            self.restore_selection(saved_selection)
            return "break"
        col = self.tree.identify_column(event.x)
        if col == "#0":
            self.restore_selection(saved_selection)
            return "break"
        col_index = int(col[1:]) - 1
        col_id = self.tree["columns"][col_index]
        if col_id == "pos":
            self.start_combo_edit(item, col, col_id, POSITIONS)
        elif col_id == "arr":
            self.start_combo_edit(item, col, col_id, ARRIVAL_TIMES)
        elif col_id == "status":
            self.start_combo_edit(item, col, col_id, STATUSES)
        elif col_id == "fio":
            self.start_text_edit(item, col, col_id)
        self.restore_selection(saved_selection)
        self.force_restore_selection(saved_selection)
        return "break"

    def restore_selection(self, selection):
        def _restore():
            if selection:
                self.tree.selection_set(selection)
            self.suppress_select = False
        self.after_idle(_restore)

    def force_restore_selection(self, selection):
        def _force():
            if selection:
                self.tree.selection_set(selection)
        self.after(10, _force)

    def on_tree_left_press(self, event):
        self.last_selection = self.tree.selection()

    def on_tree_left_release(self, event):
        self.last_selection = self.tree.selection()

    def on_tree_left_click(self, event):
        if getattr(self, "edit_widget", None):
            self.destroy_editor()
        col = self.tree.identify_column(event.x)
        if col == "#0":
            return
        item = self.tree.identify_row(event.y)
        if not item:
            return
        col_index = int(col[1:]) - 1
        col_id = self.tree["columns"][col_index]
        if col_id == "pos":
            self.start_combo_edit(item, col, col_id, POSITIONS)
            return "break"
        if col_id == "arr":
            self.start_combo_edit(item, col, col_id, ARRIVAL_TIMES)
            return "break"
        if col_id == "status":
            self.start_combo_edit(item, col, col_id, STATUSES)
            return "break"

    def copy_rows(self):
        selected = self.tree.selection()
        if not selected:
            return
        self.clipboard_data = []
        for item in selected:
            row_data = {
                "values": self.tree.item(item)["values"],
                "row_info": self.rows[item].copy()
            }
            self.clipboard_data.append(row_data)
        messagebox.showinfo("Копирование", f"Скопировано строк: {len(selected)}")

    def cut_rows(self):
        selected = self.tree.selection()
        if not selected:
            return
        self.copy_rows()
        self.delete_selected_rows()

    def delete_selected_rows(self):
        selected = self.tree.selection()
        if not selected:
            return
        if not messagebox.askyesno("Удаление", f"Удалить выбранных строк: {len(selected)}?"):
            return
        for item in selected:
            self.tree.delete(item)
            del self.rows[item]
        self.update_numbers()
        self.update_counters()
        self.save_data()

    def delete_all_table(self):
        if not self.rows:
            messagebox.showinfo("Информация", "Таблица уже пуста")
            return
        if not messagebox.askyesno("Удаление всей таблицы", 
                                   "Вы уверены, что хотите удалить ВСЕ строки из таблицы?\n\n"
                                   "Это действие нельзя отменить!"):
            return
        for item in list(self.rows.keys()):
            self.tree.delete(item)
        self.rows.clear()
        self.update_numbers()
        self.update_counters()
        self.save_data()
        messagebox.showinfo("Готово", "Все строки удалены")

    def start_timer(self):
        self.start_time = datetime.now()
        self.start_lbl.config(text=f"Время запуска: {self.start_time.strftime('%H:%M')}")
        self.save_data()

    def add_row_values(self, pos, fio, arr, status="Присутствует"):
        if not fio:
            return
        for r in self.rows.values():
            if fio.lower() == r["fio"].lower():
                if not messagebox.askyesno("Совпадение", "Есть совпадение. Продолжить?"):
                    return
        item = self.tree.insert("", "end",
            values=("", pos, fio,
                    arr,"","",status))
        self.rows[item] = {
            "fio": fio,
            "arrival": parse_hours(arr),
            "fact": None,
            "status": status
        }
        self.sort_table()
        self.update_numbers()
        self.update_counters()
        self.save_data()

    def open_add_dialog(self):
        dlg = tk.Toplevel(self)
        dlg.title("Добавить строку")
        dlg.configure(bg="#1e1e1e")
        dlg.transient(self)
        dlg.grab_set()

        tk.Label(dlg, text="Должность", bg="#1e1e1e", fg="white",
                 font=("Segoe UI", FONT_SIZE)).pack(anchor="w", padx=10, pady=(10, 0))
        pos_btn = ttk.Button(dlg, text="Должность", style="Blue.TButton")
        pos_btn.pack(fill="x", padx=10)

        tk.Label(dlg, text="ФИО", bg="#1e1e1e", fg="white",
                 font=("Segoe UI", FONT_SIZE)).pack(anchor="w", padx=10, pady=(10, 0))
        fio_ent = tk.Entry(dlg, font=("Segoe UI", FONT_SIZE), bg="#2a2a2a", fg="white",
                           insertbackground="white", highlightthickness=1,
                           highlightbackground="#4a4a4a", relief="flat")
        fio_ent.pack(fill="x", padx=10)

        tk.Label(dlg, text="Время прибытия", bg="#1e1e1e", fg="white",
                 font=("Segoe UI", FONT_SIZE)).pack(anchor="w", padx=10, pady=(10, 0))
        arr_btn = ttk.Button(dlg, text="ч+", style="Blue.TButton")
        arr_btn.pack(fill="x", padx=10, pady=(0, 10))

        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))

        def submit():
            fio = fio_ent.get().strip()
            if not fio:
                return
            pos_val = pos_btn.cget("text")
            arr_val = arr_btn.cget("text")
            if pos_val == "Должность":
                pos_val = POSITIONS[0]
            if arr_val == "ч+":
                arr_val = ARRIVAL_TIMES[0]
            self.add_row_values(pos_val, fio, arr_val, "Присутствует")
            dlg.destroy()

        ttk.Button(btn_frame, text="Добавить", command=submit, style="Green.TButton").pack(side="left", fill="x", expand=True)
        ttk.Button(btn_frame, text="Отмена", command=dlg.destroy, style="Red.TButton").pack(side="left", fill="x", expand=True, padx=(10, 0))

        fio_ent.focus_set()
        dlg.update_idletasks()
        w = dlg.winfo_width()
        h = dlg.winfo_height()
        x = (dlg.winfo_screenwidth() // 2) - (w // 2)
        y = (dlg.winfo_screenheight() // 2) - (h // 2)
        dlg.geometry(f"{w}x{h}+{x}+{y}")

        def open_pos_menu():
            bx = pos_btn.winfo_rootx()
            by = pos_btn.winfo_rooty() + pos_btn.winfo_height()
            self.open_selector(POSITIONS, bx, by, pos_btn.winfo_width(),
                               lambda v: (pos_btn.config(text=v), fio_ent.focus_set()),
                               parent=dlg, modal=True)

        def open_arr_menu():
            bx = arr_btn.winfo_rootx()
            by = arr_btn.winfo_rooty() + arr_btn.winfo_height()
            self.open_selector(ARRIVAL_TIMES, bx, by, arr_btn.winfo_width(),
                               lambda v: (arr_btn.config(text=v), fio_ent.focus_set()),
                               parent=dlg, modal=True)

        pos_btn.config(command=open_pos_menu)
        arr_btn.config(command=open_arr_menu)

    def delete_row(self):
        sel = self.tree.selection()
        if not sel: return
        item = sel[0]
        self.tree.delete(item)
        del self.rows[item]
        self.update_numbers()
        self.update_counters()
        self.save_data()

    def import_excel(self):
        file = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")])
        if not file: return
        try:
            df = pd.read_excel(file)
            if df.empty or len(df.columns) == 0:
                messagebox.showerror("Ошибка", "Файл Excel пустой или не содержит данных")
                return
            imported_count = 0
            for name in df.iloc[:, 0].dropna():
                fio = str(name).strip()
                if not fio:
                    continue
                item = self.tree.insert("", "end",
                    values=("", "Неизвестно", fio,
                            ARRIVAL_TIMES[0], "", "", "Присутствует"))
                self.rows[item] = {
                    "fio": fio,
                    "arrival": 1,
                    "fact": None,
                    "status": "Присутствует"
                }
                imported_count += 1
            self.sort_table()
            self.update_numbers()
            self.update_counters()
            self.save_data()
            messagebox.showinfo("Импорт завершен", f"Импортировано ФИО: {imported_count}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{str(e)}")

    def clear_data(self):
        if not messagebox.askyesno("Очистка","Обнулить факт, опоздание и статус?"):
            return
        for i in self.rows:
            self.tree.set(i,"fact","")
            self.tree.set(i,"late","")
            self.tree.set(i,"status","Присутствует")
            self.rows[i]["fact"]=None
            self.rows[i]["status"]="Присутствует"
            self.tree.item(i,tags=())
        self.start_time=None
        self.start_lbl.config(text="Время запуска: ---")
        self.update_counters()
        self.save_data()

    def on_tree_double_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        self.mark_arrival(item)

    def mark_arrival(self, item):
        if not self.start_time:
            return
        if self.rows[item]["fact"] is not None:
            return
        if self.rows[item]["status"] != "Присутствует":
            return
        now = datetime.now()
        self.rows[item]["fact"] = now
        self.tree.set(item, "fact", now.strftime("%H:%M"))
        expected = self.start_time + timedelta(hours=self.rows[item]["arrival"])
        late = now - expected
        if late.total_seconds() < 0:
            late = timedelta(0)
        self.tree.set(item, "late", self.format_timedelta(late))
        self.apply_color(item)
        self.update_counters()
        self.save_data()

    def start_combo_edit(self, item, col, col_id, values):
        self.destroy_editor()
        x, y, w, h = self.tree.bbox(item, col)
        if w <= 0:
            return
        abs_x = self.tree.winfo_rootx() + x
        abs_y = self.tree.winfo_rooty() + y + h
        self.open_selector(values, abs_x, abs_y, w,
                           lambda v: self.commit_combo(item, col_id, v))

    def open_selector(self, values, x, y, width, on_select, parent=None, modal=False):
        self.destroy_editor()
        popup_parent = parent if parent else self
        popup = tk.Toplevel(popup_parent)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.configure(bg="#1e1e1e")
        item_h = 32
        height = item_h * len(values) + 2
        width_px = max(200, width)
        # Подгоняем по экрану
        screen_h = popup.winfo_screenheight()
        screen_w = popup.winfo_screenwidth()
        if y + height > screen_h:
            y = max(0, y - height)
        if x + width_px > screen_w:
            x = max(0, screen_w - width_px)
        popup.geometry(f"{width_px}x{height}+{x}+{y}")
        popup.update_idletasks()
        popup.attributes("-topmost", True)
        if parent:
            popup.transient(parent)
        if modal:
            try:
                self.prev_grab = popup.grab_current()
                if self.prev_grab:
                    self.prev_grab.grab_release()
                popup.grab_set()
            except Exception:
                self.prev_grab = None
        popup.lift()
        popup.focus_force()

        container = tk.Frame(popup, bg="#1f1f1f", highlightthickness=1, highlightbackground="#4a4a4a")
        container.pack(fill="both", expand=True)

        def make_row(text):
            row = tk.Frame(container, bg="#1f1f1f", height=item_h)
            row.pack_propagate(False)
            row.pack(fill="x")
            lbl = tk.Label(row, text=text, bg="#1f1f1f", fg="white",
                           font=("Segoe UI", COMBO_FONT_SIZE), anchor="w", padx=10)
            lbl.pack(fill="both", expand=True)
            sep = tk.Frame(container, bg="#3a3a3a", height=1)
            sep.pack(fill="x")

            def set_hover(on):
                color = "#5a5a5a" if on else "#1f1f1f"
                row.configure(bg=color)
                lbl.configure(bg=color)

            def choose():
                on_select(text)
                self.destroy_editor()

            row.bind("<Enter>", lambda e: set_hover(True))
            row.bind("<Leave>", lambda e: set_hover(False))
            lbl.bind("<Enter>", lambda e: set_hover(True))
            lbl.bind("<Leave>", lambda e: set_hover(False))
            row.bind("<ButtonRelease-1>", lambda e: choose())
            lbl.bind("<ButtonRelease-1>", lambda e: choose())
            return row

        for v in values:
            make_row(v)

        popup.bind("<Escape>", lambda e: self.destroy_editor())
        popup.focus_set()
        self.edit_widget = container
        self.edit_popup = popup

    def start_text_edit(self, item, col, col_id):
        self.destroy_editor()
        x, y, w, h = self.tree.bbox(item, col)
        if w <= 0:
            return
        ent = tk.Entry(self.tree, font=("Segoe UI", FONT_SIZE), bg="#3a3a3a", fg="white",
                       insertbackground="white", highlightthickness=1,
                       highlightbackground="#555555", relief="flat")
        ent.place(x=x, y=y, width=w, height=h)
        ent.insert(0, self.tree.set(item, col_id))
        ent.select_range(0, tk.END)
        ent.bind("<Return>", lambda e: self.commit_text(item, col_id, ent.get()))
        ent.bind("<FocusOut>", lambda e: self.commit_text(item, col_id, ent.get()))
        ent.bind("<Escape>", lambda e: self.destroy_editor())
        ent.focus_set()
        self.edit_widget = ent

    def destroy_editor(self):
        if getattr(self, "edit_widget", None):
            try:
                self.edit_widget.destroy()
            finally:
                self.edit_widget = None
        if getattr(self, "edit_popup", None):
            try:
                try:
                    self.edit_popup.grab_release()
                except Exception:
                    pass
                self.edit_popup.destroy()
            finally:
                self.edit_popup = None
        if getattr(self, "prev_grab", None):
            try:
                self.prev_grab.grab_set()
            except Exception:
                pass
            self.prev_grab = None

    def commit_combo(self, item, col_id, value):
        self.destroy_editor()
        if not self.tree.exists(item):
            return
        self.tree.set(item, col_id, value)
        if col_id == "pos":
            pass
        elif col_id == "arr":
            self.rows[item]["arrival"] = parse_hours(value)
        elif col_id == "status":
            self.rows[item]["status"] = value
        self.apply_color(item)
        self.update_counters()
        self.save_data()

    def commit_text(self, item, col_id, value):
        self.destroy_editor()
        if not self.tree.exists(item):
            return
        value = value.strip()
        if not value:
            return
        self.tree.set(item, col_id, value)
        if col_id == "fio":
            self.rows[item]["fio"] = value
        self.sort_table()
        self.update_numbers()
        self.update_counters()
        self.save_data()

    def apply_color(self, item):
        status = self.rows[item]["status"]
        late_text = self.tree.set(item,"late")
        if status != "Присутствует":
            self.tree.item(item, tags=("other",))
        elif late_text and late_text != "00:00":
            self.tree.item(item, tags=("late",))
        elif self.rows[item]["fact"]:
            self.tree.item(item, tags=("ontime",))
        else:
            self.tree.item(item, tags=())

    def sort_table(self):
        items = list(self.tree.get_children())
        items.sort(key=lambda i: self.tree.set(i,"fio").lower())
        for i in items:
            self.tree.move(i, "", "end")

    def update_numbers(self):
        for idx, item in enumerate(self.tree.get_children(), start=1):
            self.tree.set(item, "num", idx)

    def update_chart(self, total, present, sick, trip):
        self.chart.delete("all")
        max_h = 150
        w = int(self.chart.winfo_width() or 320)
        pad = 10
        bar_w = 40
        gap = 12
        x = pad

        show_stats = self.start_time is not None

        def bar(count, color, label=None):
            nonlocal x
            height = int((count / total) * max_h) if total else 0
            y0 = max_h - height + 10
            y1 = max_h + 10
            self.chart.create_rectangle(x, y0, x + bar_w, y1, fill=color, outline="")
            if label:
                self.chart.create_text(x + bar_w / 2, y1 + 12, text=label, fill="white", font=("Segoe UI", 10))
            x += bar_w + gap

        # Всего всегда активно
        bar(total, "#d6b400", "Всего")

        if show_stats:
            late = sum(1 for i, r in self.rows.items()
                       if r["status"] == "Присутствует"
                       and r["fact"]
                       and self.tree.set(i, "late") not in ("", "00:00"))
            ontime = sum(1 for i, r in self.rows.items()
                         if r["status"] == "Присутствует"
                         and r["fact"]
                         and self.tree.set(i, "late") in ("", "00:00"))
            blue_total = sick + trip
            bar(late, "#aa2222")
            bar(blue_total, "#2255aa")
            bar(ontime, "#1e8f3a")
        else:
            bar(0, "#aa2222")
            bar(0, "#2255aa")
            bar(0, "#1e8f3a")

    def update_counters(self):
        total = len(self.rows)
        present = sum(1 for r in self.rows.values() if r["status"]=="Присутствует")
        sick = sum(1 for r in self.rows.values() if r["status"]=="Болен")
        trip = sum(1 for r in self.rows.values() if r["status"]=="Командировка")
        arrived = sum(1 for r in self.rows.values() if r["status"]=="Присутствует" and r["fact"])
        percent = (arrived/present*100) if present else 0

        def pct(n):
            return (n/total*100) if total else 0

        self.total_lbl.config(text=f"Всего по списку: {total}")
        self.present_lbl.config(text=f"Присутствует: {present} ({pct(present):.1f}%)")
        self.sick_lbl.config(text=f"Болен: {sick} ({pct(sick):.1f}%)")
        self.trip_lbl.config(text=f"Командировка: {trip} ({pct(trip):.1f}%)")
        self.arrived_lbl.config(text=f"Пришли: {arrived}")
        self.percent_lbl.config(text=f"Процент прибытия: {percent:.1f}%")
        self.update_chart(total, present, sick, trip)

    def on_select(self, event):
        if self.suppress_select:
            return
        sel = self.tree.selection()
        if not sel: return
        item = sel[0]
        self.current_item = item

    def search_dialog(self, event=None):
        if hasattr(self, "search_ent"):
            self.search_ent.focus_set()
            self.search_ent.select_range(0, tk.END)

    def search_from_entry(self, event=None):
        if not hasattr(self, "search_ent"):
            return
        q = self.search_ent.get().strip()
        if not q:
            return
        self.search_matches = [i for i in self.tree.get_children()
                               if q.lower() in self.tree.set(i,"fio").lower()]
        if not self.search_matches:
            messagebox.showinfo("Поиск","Не найдено")
            return
        self.search_index = 0
        item = self.search_matches[self.search_index]
        self.tree.selection_set(item)
        self.tree.see(item)

    def search_next(self):
        if not self.search_matches:
            self.search_from_entry()
            return
        self.search_index = (self.search_index + 1) % len(self.search_matches)
        item = self.search_matches[self.search_index]
        self.tree.selection_set(item)
        self.tree.see(item)

    def export_excel(self, hours):
        if not self.start_time:
            messagebox.showerror("Ошибка","Нажмите Запуск")
            return
        data=[]
        present_total = 0
        present_arrived = 0
        end = self.start_time + timedelta(hours=hours)
        for i,r in self.rows.items():
            row = list(self.tree.item(i)["values"])
            if not r["fact"]:
                row[4] = "Еще не прибыл"
                row[5] = ""
            data.append(row[1:])
            if r["status"]=="Присутствует":
                present_total += 1
                if r["fact"] and self.start_time <= r["fact"] <= end:
                    present_arrived += 1
        percent = (present_arrived/present_total*100) if present_total else 0
        date_str = (self.start_time or datetime.now()).strftime("%Y-%m-%d")
        header_text = f"{date_str}.прибытие ч+{hours}"
        data = [[header_text, "", "", "", "", ""]] + data
        df=pd.DataFrame(data,columns=["Должность","ФИО","Время прибытия","Факт","Опоздание","Статус"])
        df.loc[len(df)] = ["","","","","",""]
        df.loc[len(df)] = ["Процент прибытия","","","","",f"{percent:.1f}%"]

        fname=f"export_{hours}h.xlsx"
        df.to_excel(fname,index=False)
        full_path = os.path.abspath(fname)
        messagebox.showinfo("Готово",f"Сохранено: {full_path}")

    def open_export_menu(self):
        values = [f"{h} ч" for h in [1, 1.5, 2, 2.5, 3]]
        bx = self.export_btn.winfo_rootx()
        by = self.export_btn.winfo_rooty() + self.export_btn.winfo_height()
        self.open_selector(values, bx, by, self.export_btn.winfo_width(),
                           lambda v: self.export_excel(float(v.replace("ч", "").replace(" ", ""))))

    def save_data(self):
        data={"start": self.start_time.isoformat() if self.start_time else None,
              "rows":[]}
        for i,r in self.rows.items():
            vals = list(self.tree.item(i)["values"])
            data["rows"].append({
                "values": vals,
                "fact": r["fact"].isoformat() if r["fact"] else None,
                "arrival": r["arrival"],
                "status": r["status"]
            })
        with open(DATA_FILE,"w",encoding="utf8") as f:
            json.dump(data,f,ensure_ascii=False,indent=2)

    def load_data(self):
        if not os.path.exists(DATA_FILE): return
        with open(DATA_FILE,"r",encoding="utf8") as f:
            data=json.load(f)
        if data["start"]:
            self.start_time=datetime.fromisoformat(data["start"])
            self.start_lbl.config(text=f"Время запуска: {self.start_time.strftime('%H:%M')}")
        for row in data["rows"]:
            vals = list(row["values"])
            i=self.tree.insert("", "end", values=vals)
            self.rows[i]={ 
                "fio": row["values"][2],
                "arrival":row["arrival"],
                "fact": datetime.fromisoformat(row["fact"]) if row["fact"] else None,
                "status":row["status"]
            }
            self.apply_color(i)

    def on_close(self):
        self.save_data()
        self.destroy()


if __name__=="__main__":
    App().mainloop()
