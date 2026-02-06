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

        self.setup_style()
        self.create_widgets()
        self.load_data()
        self.sort_table()
        self.update_numbers()
        self.update_counters()

        self.bind("<Control-f>", self.search_dialog)
        #self.bind_all("<Button-1>", self.global_click, add="+")  # Удалено
        self.tree.bind("<MouseWheel>", self.fast_scroll)
        
        # Контекстное меню по правой кнопке мыши
        self.tree.bind("<Button-3>", self.show_context_menu)

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
        style.configure("Treeview.Heading", font=("Segoe UI", FONT_SIZE, "bold"))

        style.configure("TButton", padding=10)
        style.configure("Green.TButton", background="#1e8f3a", foreground="white")
        style.configure("Red.TButton", background="#aa2222", foreground="white")
        style.configure("Blue.TButton", background="#2255aa", foreground="white")
        style.map("TButton", background=[("active", "#555555")])

        style.configure("TEntry", fieldbackground=field, foreground="black")
        style.configure("TCombobox", fieldbackground=field, foreground="black")

    def create_widgets(self):
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=5)

        self.start_btn = ttk.Button(top, text="ЗАПУСК", command=self.start_timer,
                                    style="Green.TButton", width=14)
        self.start_btn.pack(side="left", padx=5)

        self.start_lbl = ttk.Label(top, text="Время запуска: ---")
        self.start_lbl.pack(side="left", padx=20)

        self.clear_btn = ttk.Button(top, text="ОЧИСТИТЬ ДАННЫЕ", command=self.clear_data,
                                    style="Red.TButton", width=18)
        self.clear_btn.pack(side="right", padx=10)

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
        self.tree.bind("<Double-1>", self.on_double)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        scroll = ttk.Scrollbar(main, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="left", fill="y")

        right = ttk.Frame(main)
        right.pack(side="right", fill="y", padx=10)
        right.configure(width=360)

        def label(txt):
            ttk.Label(right, text=txt).pack(anchor="w", pady=(6,0))

        label("Должность")
        self.pos_cb = ttk.Combobox(right, values=POSITIONS, state="readonly")
        self.pos_cb.current(0)
        self.pos_cb.pack(fill="x")

        label("ФИО")
        self.fio_ent = ttk.Entry(right)
        self.fio_ent.pack(fill="x")

        label("Время прибытия")
        self.arr_cb = ttk.Combobox(right, values=ARRIVAL_TIMES, state="readonly")
        self.arr_cb.pack(fill="x")

        label("Статус")
        self.status_cb = ttk.Combobox(right, values=STATUSES, state="readonly")
        self.status_cb.pack(fill="x")
        self.status_cb.bind("<<ComboboxSelected>>", self.update_row_from_right)

        btn_add = ttk.Button(right, text="Добавить", command=self.add_row)
        btn_add.pack(fill="x", pady=5)

        btn_del = ttk.Button(right, text="Удалить строку", command=self.delete_row)
        btn_del.pack(fill="x")

        ttk.Button(right, text="Загрузка ФИО из Excel", command=self.import_excel).pack(fill="x", pady=5)

        ttk.Separator(right).pack(fill="x", pady=8)

        for h in [1,1.5,2,2.5,3]:
            ttk.Button(right, text=f"Выгрузка за {h} ч",
                       command=lambda x=h:self.export_excel(x)).pack(fill="x")

        ttk.Separator(right).pack(fill="x", pady=8)

        self.present_lbl = ttk.Label(right)
        self.sick_lbl = ttk.Label(right)
        self.trip_lbl = ttk.Label(right)
        self.arrived_lbl = ttk.Label(right)
        self.percent_lbl = ttk.Label(right)

        self.present_lbl.pack(anchor="w")
        self.sick_lbl.pack(anchor="w")
        self.trip_lbl.pack(anchor="w")
        self.arrived_lbl.pack(anchor="w")
        self.percent_lbl.pack(anchor="w")

        ttk.Separator(right).pack(fill="x", pady=15)
        
        self.delete_all_btn = ttk.Button(right, text="УДАЛИТЬ ВСЮ ТАБЛИЦУ", 
                                         command=self.delete_all_table,
                                         style="Blue.TButton")
        self.delete_all_btn.pack(fill="x", side="bottom", pady=(10, 0))

        # Двусторонняя привязка
        self.pos_cb.bind("<<ComboboxSelected>>", self.update_row_from_right)
        self.fio_ent.bind("<KeyRelease>", self.update_row_from_right)
        self.arr_cb.bind("<<ComboboxSelected>>", self.update_row_from_right)

    # ====== Методы ======
    def fast_scroll(self, event):
        self.tree.yview_scroll(int(-3*(event.delta/120)), "units")

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item and item not in self.tree.selection():
            self.tree.selection_set(item)
        if self.tree.selection():
            menu = tk.Menu(self, tearoff=0, font=("Segoe UI", 14))
            menu.add_command(label="Копировать", command=self.copy_rows)
            menu.add_command(label="Вырезать", command=self.cut_rows)
            menu.add_separator()
            menu.add_command(label="Удалить", command=self.delete_selected_rows)
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

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
        self.start_lbl.config(text=f"Время запуска: {self.start_time.strftime('%H:%M:%S')}")
        self.save_data()

    def add_row(self):
        fio = self.fio_ent.get().strip()
        if not fio:
            return
        for r in self.rows.values():
            if fio.lower() == r["fio"].lower():
                if not messagebox.askyesno("Совпадение", "Есть совпадение. Продолжить?"):
                    return
        item = self.tree.insert("", "end",
            values=("", self.pos_cb.get(), fio,
                    self.arr_cb.get(),"","",self.status_cb.get()))
        self.rows[item] = {
            "fio": fio,
            "arrival": parse_hours(self.arr_cb.get()),
            "fact": None,
            "status": self.status_cb.get()
        }
        self.fio_ent.delete(0, tk.END)
        self.sort_table()
        self.update_numbers()
        self.update_counters()
        self.save_data()

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
                    values=("", "Неизвестно", fio, ARRIVAL_TIMES[0], "", "", "Присутствует"))
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

    def on_double(self, event):
        if not self.start_time:
            return
        item = self.tree.identify_row(event.y)
        if not item: return
        if self.rows[item]["fact"] is not None: return
        if self.rows[item]["status"] != "Присутствует": return
        now = datetime.now()
        self.rows[item]["fact"] = now
        self.tree.set(item,"fact",now.strftime("%H:%M:%S"))
        expected = self.start_time + timedelta(hours=self.rows[item]["arrival"])
        late = now - expected
        if late.total_seconds()<0:
            late = timedelta(0)
        self.tree.set(item,"late",str(late))
        self.apply_color(item)
        self.update_counters()
        self.save_data()

    def update_row_from_right(self, event=None):
        if not self.current_item:
            return
        item = self.current_item
        if not self.tree.exists(item):
            return
        self.tree.set(item, "pos", self.pos_cb.get())
        self.tree.set(item, "fio", self.fio_ent.get())
        self.tree.set(item, "arr", self.arr_cb.get())
        self.tree.set(item, "status", self.status_cb.get())
        self.rows[item]["fio"] = self.fio_ent.get()
        self.rows[item]["arrival"] = parse_hours(self.arr_cb.get())
        self.rows[item]["status"] = self.status_cb.get()
        self.apply_color(item)
        self.update_counters()
        self.save_data()

    def apply_color(self, item):
        status = self.rows[item]["status"]
        late_text = self.tree.set(item,"late")
        if status != "Присутствует":
            self.tree.item(item, tags=("other",))
        elif late_text and late_text != "0:00:00":
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

    def update_counters(self):
        present = sum(1 for r in self.rows.values() if r["status"]=="Присутствует")
        sick = sum(1 for r in self.rows.values() if r["status"]=="Болен")
        trip = sum(1 for r in self.rows.values() if r["status"]=="Командировка")
        arrived = sum(1 for r in self.rows.values() if r["status"]=="Присутствует" and r["fact"])
        percent = (arrived/present*100) if present else 0
        self.present_lbl.config(text=f"Присутствует: {present}")
        self.sick_lbl.config(text=f"Болен: {sick}")
        self.trip_lbl.config(text=f"Командировка: {trip}")
        self.arrived_lbl.config(text=f"Пришли: {arrived}")
        self.percent_lbl.config(text=f"Процент: {percent:.1f}%")

    def on_select(self, event):
        sel = self.tree.selection()
        if not sel: return
        item = sel[0]
        vals = self.tree.item(item)["values"]
        self.pos_cb.set(vals[1])
        self.fio_ent.delete(0,tk.END)
        self.fio_ent.insert(0,vals[2])
        self.arr_cb.set(vals[3])
        self.status_cb.set(vals[6])
        self.current_item = item

    def search_dialog(self, event=None):
        q = simpledialog.askstring("Поиск","Введите фамилию:")
        if not q: return
        for i in self.tree.get_children():
            if q.lower() in self.tree.set(i,"fio").lower():
                self.tree.selection_set(i)
                self.tree.see(i)
                return
        messagebox.showinfo("Поиск","Не найдено")

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
        df=pd.DataFrame(data,columns=["Должность","ФИО","Время прибытия","Факт","Опоздание","Статус"])
        df.loc[len(df)] = ["","","","","",""]
        df.loc[len(df)] = ["Процент прибытия","","","","",f"{percent:.1f}%"]

        fname=f"export_{hours}h.xlsx"
        df.to_excel(fname,index=False)
        messagebox.showinfo("Готово",f"Сохранено: {fname}")

    def save_data(self):
        data={"start": self.start_time.isoformat() if self.start_time else None,
              "rows":[]}
        for i,r in self.rows.items():
            data["rows"].append({
                "values": self.tree.item(i)["values"],
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
            self.start_lbl.config(text=f"Время запуска: {self.start_time.strftime('%H:%M:%S')}")
        for row in data["rows"]:
            i=self.tree.insert("", "end", values=row["values"])
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
