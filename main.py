# ===================== LISTAT =====================

DEFAULT_PARTS = [
    # Pezullim & Drejtim
    "Amortizator para",
    "Amortizator mbrapa",
    "Suporte amortizatori",
    "Susta para",
    "Susta mbrapa",
    "Kit Amortizimi (Amorizator + Susta + Suporte)",
    "Krah Hekuri",
    "Krah Alumini",
    "Gomina krahu",
    "Stabilizator para",
    "Stabilizator mbrapa",
    "Kok Paraleli",
    "Tirant",
    "Grusht i Jashtëm",
    "Kushinetë",

    # Frenim
    "Tamburo",
    "Ferrota",
    "Sensor ABS",

    # Motor & Servis
    "Filtri vajit",
    "Filtri ajrit",
    "Filtri kabinës",
    "Filtri karburantit",
    "Vaj motori",
    "Vaj kambioje",
    "Antifriz",
    "Alkool Frenash",
    "Kandele",
    "Bombina",
    "Kandeleta (Naftë)",
    "Kit Zinxhiri",
    "Kit Faze",
    
    
    # Rripa & Ftohje
    "Rrip dinamo",
    "Rrip faze",
    "Pompa e ujit",
    "Termostat",
    "Radiator",

    # Elektrike
    "Bateri",
    "Dinamo",
    "Motorrino",

    # Transmision
    "Kit Freksioni",

    # Shërbim
    "Servis i përgjithshëm"
]

FUEL_TYPES = ["Naftë", "Benzinë", "Benzinë + Gaz", "Gaz (LPG)", "Hibrid", "Elektrik", "Tjetër"]


# ===================== IMPORTS =====================

import os
import sqlite3
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter.font as tkfont

import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

import win32api
import win32print

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet


APP_FOLDER = str(Path.home() / "Documents" / "Preventiva")
DB = str(Path.home() / "Documents" / "Preventiva" / "preventiv.db")
PDF_FOLDER = str(Path.home() / "Documents" / "Preventiva")


# ===================== DB =====================

def init_folders():
    os.makedirs(APP_FOLDER, exist_ok=True)
    os.makedirs(PDF_FOLDER, exist_ok=True)

def db_connect():
    return sqlite3.connect(DB)

def init_db():
    init_folders()
    conn = db_connect()
    c = conn.cursor()

    c.execute("""
    CREATE TABLE IF NOT EXISTS vehicles (
        plate TEXT PRIMARY KEY,
        client TEXT,
        model TEXT,
        fuel TEXT,
        vtype TEXT,
        updated_at TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS preventiv (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT,
        plate TEXT,
        client TEXT,
        model TEXT,
        fuel TEXT,
        vtype TEXT,
        labor INTEGER,
        parts_total INTEGER,
        grand_total INTEGER,
        notes TEXT,
        pdf_path TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS preventiv_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        preventiv_id INTEGER,
        description TEXT,
        qty INTEGER,
        price INTEGER,
        total INTEGER
    )
    """)

    conn.commit()
    conn.close()


# ===================== LOGIN =====================

class LoginWindow:
    def __init__(self, root):
        self.root = root
        root.title("Login - Preventiv")
        root.geometry("420x270")
        root.resizable(False, False)

        tk.Label(root, text="AUTO SERVIS ALDO", font=("Segoe UI", 18, "bold")).pack(pady=14)

        tk.Label(root, text="Username", font=("Segoe UI", 12)).pack()
        self.username = tk.Entry(root, font=("Segoe UI", 12))
        self.username.pack(pady=(2, 10))

        tk.Label(root, text="Password", font=("Segoe UI", 12)).pack()
        self.password = tk.Entry(root, show="*", font=("Segoe UI", 12))
        self.password.pack(pady=(2, 14))

        tk.Button(root, text="HYR", width=18, command=self.check,
                  font=("Segoe UI", 12, "bold")).pack(pady=10)

        self.username.focus_set()

    def check(self):
        if self.username.get() == "admin" and self.password.get() == "1234":
            self.root.destroy()
            open_main_app()
        else:
            messagebox.showerror("Gabim", "Username ose Password gabim!")


# ===================== MAIN APP =====================

class PreventivApp:
    def __init__(self, root):
        self.root = root
        root.title("PREVENTIV - AUTO SERVIS ALDO")
        root.state("zoomed")
        root.configure(bg="#f2f4f7")

        self.last_pdf = None
        self._plate_lookup_after_id = None

        # ✅ Rrit font global
        self._boost_global_fonts()

        self.style = ttk.Style()
        self.style.theme_use("clam")
        self._apply_styles()

        self._build_ui()
        self._update_totals()
        self._update_dashboard()

    def _boost_global_fonts(self):
        try:
            tkfont.nametofont("TkDefaultFont").configure(size=12)
            tkfont.nametofont("TkTextFont").configure(size=12)
            tkfont.nametofont("TkFixedFont").configure(size=12)
        except Exception:
            pass

    def _apply_styles(self):
        self.style.configure("TButton", padding=8)
        self.style.configure("Treeview", font=("Segoe UI", 12), rowheight=36)
        self.style.configure("Treeview.Heading", font=("Segoe UI", 12, "bold"))

    # ---------- UI ----------

    def _build_ui(self):
        header = tk.Frame(self.root, bg="#1f3c88", height=90)
        header.pack(fill="x")

        tk.Label(header, text="AUTO SERVIS ALDO",
                 bg="#1f3c88", fg="white",
                 font=("Segoe UI", 24, "bold")).pack(pady=(12, 0))

        self.date_str = datetime.now().strftime("%d/%m/%Y")
        tk.Label(header, text=f"PREVENTIV   |   Data: {self.date_str}",
                 bg="#1f3c88", fg="white",
                 font=("Segoe UI", 15)).pack(pady=(2, 12))

        body = tk.Frame(self.root, bg="#f2f4f7")
        body.pack(fill="both", expand=True, padx=18, pady=14)

        left = tk.Frame(body, bg="white", bd=1, relief="solid")
        left.pack(side="left", fill="y", padx=(0, 12))

        center = tk.Frame(body, bg="white", bd=1, relief="solid")
        center.pack(side="left", fill="both", expand=True)

        right = tk.Frame(body, bg="white", bd=1, relief="solid")
        right.pack(side="left", fill="y", padx=(12, 0))

        # LEFT
        tk.Label(left, text="KLIENTI / MAKINA", bg="white",
                 font=("Segoe UI", 14, "bold")).pack(pady=12)

        self.client = self._field(left, "Emri")
        self.plate = self._field(left, "Targa")
        self.model = self._field(left, "Modeli")

        tk.Label(left, text="Karburanti", bg="white", font=("Segoe UI", 12)).pack(anchor="w", padx=10)
        self.fuel = ttk.Combobox(left, values=FUEL_TYPES, state="readonly", width=26, font=("Segoe UI", 12))
        self.fuel.pack(padx=10, pady=(2, 8))
        self.fuel.set(FUEL_TYPES[0])

        self.vtype = self._field(left, "Lloji i makinës")

        self.plate.bind("<KeyRelease>", self._plate_keyrelease)

        tk.Label(left, text="Shënime", bg="white", font=("Segoe UI", 12)).pack(pady=(12, 2))
        self.notes = tk.Text(left, height=6, width=28, font=("Segoe UI", 12))
        self.notes.pack(padx=10, pady=(0, 10))

        # CENTER
        tk.Label(center, text="PJESËT", bg="white",
                 font=("Segoe UI", 14, "bold")).pack(pady=12)

        table_frame = tk.Frame(center, bg="white")
        table_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        cols = ("desc", "qty", "price", "total")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings")
        self.tree.heading("desc", text="Përshkrimi")
        self.tree.heading("qty", text="Sasia")
        self.tree.heading("price", text="Çmimi (Lek)")
        self.tree.heading("total", text="Totali (Lek)")

        self.tree.column("desc", width=460, anchor="w")
        self.tree.column("qty", width=90, anchor="center")
        self.tree.column("price", width=160, anchor="e")
        self.tree.column("total", width=170, anchor="e")

        self.tree.pack(fill="both", expand=True)

        center_btns = tk.Frame(center, bg="white")
        center_btns.pack(pady=(0, 14))

        ttk.Button(center_btns, text="Shto pjesë", command=self._add_part_dialog).pack(side="left", padx=7)
        ttk.Button(center_btns, text="Fshi të zgjedhurën", command=self._delete_selected_part).pack(side="left", padx=7)
        ttk.Button(center_btns, text="Pastro tabelën", command=self._clear_parts).pack(side="left", padx=7)

        # RIGHT (me TOTAL FINAL poshtë djathtas)
        right.pack_propagate(False)
        right.configure(width=380)

        for r in range(0, 30):
            right.grid_rowconfigure(r, weight=0)
        right.grid_rowconfigure(29, weight=1)
        right.grid_columnconfigure(0, weight=1)

        row = 0
        tk.Label(right, text="TOTALI", bg="white",
                 font=("Segoe UI", 14, "bold")).grid(row=row, column=0, sticky="w", padx=12, pady=(12, 6))
        row += 1

        self.parts_total_label = tk.Label(right, text="Totali pjesëve: 0 Lek",
                                          bg="white", font=("Segoe UI", 13))
        self.parts_total_label.grid(row=row, column=0, sticky="w", padx=12, pady=(6, 6))
        row += 1

        tk.Label(right, text="Haku i punës (Lek)", bg="white", font=("Segoe UI", 12)).grid(
            row=row, column=0, sticky="w", padx=12, pady=(6, 2)
        )
        row += 1

        self.labor = tk.Entry(right, width=18, font=("Segoe UI", 12))
        self.labor.grid(row=row, column=0, sticky="w", padx=12, pady=(2, 10))
        self.labor.bind("<KeyRelease>", lambda e: self._update_totals())
        row += 1

        actions = tk.Frame(right, bg="white")
        actions.grid(row=row, column=0, sticky="ew", padx=12, pady=(6, 10))
        actions.grid_columnconfigure(0, weight=1)

        ttk.Button(actions, text="Ruaj", command=self.save).grid(row=0, column=0, sticky="ew", pady=3)
        ttk.Button(actions, text="Ruaj & Printo", command=self.save_and_print).grid(row=1, column=0, sticky="ew", pady=3)
        ttk.Button(actions, text="Historiku (targa)", command=self.show_history).grid(row=2, column=0, sticky="ew", pady=3)
        ttk.Button(actions, text="Raport (periudhë)", command=self.report_dialog).grid(row=3, column=0, sticky="ew", pady=3)
        ttk.Button(actions, text="Eksport Excel", command=self.export_excel).grid(row=4, column=0, sticky="ew", pady=3)
        row += 1

        dash = tk.Frame(right, bg="white")
        dash.grid(row=row, column=0, sticky="nsew", padx=12, pady=(6, 10))
        dash.grid_rowconfigure(1, weight=1)
        dash.grid_columnconfigure(0, weight=1)

        tk.Label(dash, text="DASHBOARD - XHIRO DITORE (30 ditë)",
                 bg="white", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w")

        self.fig = Figure(figsize=(3.8, 2.4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, dash)
        self.canvas.get_tk_widget().grid(row=1, column=0, sticky="nsew", pady=(6, 0))

        # ✅ Total poshtë djathtas
        self.grand_total_label = tk.Label(
            right, text="TOTAL FINAL: 0 Lek",
            bg="white", fg="#1f3c88",
            font=("Segoe UI", 26, "bold")
        )
        self.grand_total_label.grid(row=29, column=0, sticky="se", padx=12, pady=12)

    def _field(self, parent, label):
        tk.Label(parent, text=label, bg="white", font=("Segoe UI", 12)).pack(anchor="w", padx=10)
        e = tk.Entry(parent, width=28, font=("Segoe UI", 12))
        e.pack(padx=10, pady=(2, 8))
        return e

    # ---------- Helpers ----------

    def _money(self, n: int) -> str:
        return f"{int(n):,} Lek"

    def _get_parts_total(self) -> int:
        total = 0
        for item in self.tree.get_children():
            total += int(self.tree.item(item)["values"][3])
        return total

    def _get_labor(self) -> int:
        v = self.labor.get().strip()
        return int(v) if v.isdigit() else 0

    def _update_totals(self):
        parts_total = self._get_parts_total()
        labor = self._get_labor()
        grand = parts_total + labor
        self.parts_total_label.config(text=f"Totali pjesëve: {self._money(parts_total)}")
        self.grand_total_label.config(text=f"TOTAL FINAL: {self._money(grand)}")

    # ---------- Parts dialog (dropdown + search) ----------

    def _add_part_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Shto pjesë")
        win.geometry("660x520")
        win.resizable(False, False)
        win.grab_set()

        # ── Kërkimi live ──────────────────────────────────────────
        search_frame = tk.Frame(win, bg="#f2f4f7")
        search_frame.pack(fill="x", padx=14, pady=(14, 4))

        tk.Label(search_frame, text="🔍 Kërko pjesën:",
                 font=("Segoe UI", 12, "bold"), bg="#f2f4f7").pack(anchor="w")

        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var,
                                font=("Segoe UI", 13), width=50)
        search_entry.pack(fill="x", pady=(4, 0))

        # ── Lista me scroll ────────────────────────────────────────
        list_frame = tk.Frame(win)
        list_frame.pack(fill="both", expand=True, padx=14, pady=(6, 4))

        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        listbox = tk.Listbox(
            list_frame,
            font=("Segoe UI", 12),
            selectbackground="#1f3c88",
            selectforeground="white",
            activestyle="none",
            height=10,
            yscrollcommand=scrollbar.set
        )
        scrollbar.config(command=listbox.yview)
        scrollbar.pack(side="right", fill="y")
        listbox.pack(side="left", fill="both", expand=True)

        # Mbush listën fillestare
        for part in DEFAULT_PARTS:
            listbox.insert("end", part)

        # ── Filter live ────────────────────────────────────────────
        def filter_list(*args):
            text = search_var.get().lower().strip()
            listbox.delete(0, "end")
            filtered = [p for p in DEFAULT_PARTS if text in p.lower()] if text else DEFAULT_PARTS
            for part in filtered:
                listbox.insert("end", part)
            if filtered:
                listbox.selection_set(0)

        search_var.trace_add("write", filter_list)

        # Klik në listë → plotëso search me atë vlerë
        def on_select(event=None):
            sel = listbox.curselection()
            if sel:
                search_var.set(listbox.get(sel[0]))

        listbox.bind("<<ListboxSelect>>", on_select)

        # Enter në search → zgjidh të parën në listë
        def search_enter(event=None):
            if listbox.size() > 0:
                listbox.selection_clear(0, "end")
                listbox.selection_set(0)
                listbox.see(0)
                search_var.set(listbox.get(0))
            qty_entry.focus_set()

        search_entry.bind("<Return>", search_enter)
        search_entry.bind("<Down>", lambda e: listbox.focus_set() or listbox.selection_set(0))

        # Navigim me tastierë brenda listës
        def listbox_enter(event=None):
            on_select()
            qty_entry.focus_set()

        listbox.bind("<Return>", listbox_enter)

        # ── Sasia & Çmimi ──────────────────────────────────────────
        bottom = tk.Frame(win, bg="white", bd=1, relief="solid")
        bottom.pack(fill="x", padx=14, pady=(4, 0))

        row_f = tk.Frame(bottom, bg="white")
        row_f.pack(padx=10, pady=10)

        tk.Label(row_f, text="Sasia:", font=("Segoe UI", 12), bg="white").grid(row=0, column=0, padx=8, sticky="e")
        qty_entry = tk.Entry(row_f, width=10, font=("Segoe UI", 12))
        qty_entry.grid(row=0, column=1, padx=8)
        qty_entry.insert(0, "1")

        tk.Label(row_f, text="Çmimi (Lek):", font=("Segoe UI", 12), bg="white").grid(row=0, column=2, padx=8, sticky="e")
        price_entry = tk.Entry(row_f, width=12, font=("Segoe UI", 12))
        price_entry.grid(row=0, column=3, padx=8)

        # ── Butoni Shto ────────────────────────────────────────────
        def add():
            d = search_var.get().strip()
            q = qty_entry.get().strip()
            p = price_entry.get().strip()

            if not d:
                messagebox.showwarning("Kujdes", "Zgjidh ose shkruaj përshkrimin.", parent=win)
                return
            if not q.isdigit() or not p.isdigit():
                messagebox.showwarning("Kujdes", "Sasia dhe Çmimi duhet të jenë numra.", parent=win)
                return

            q_i, p_i = int(q), int(p)
            self.tree.insert("", "end", values=(d, q_i, p_i, q_i * p_i))
            self._update_totals()

            # Reset për të shtuar pjesën tjetër
            search_var.set("")
            qty_entry.delete(0, "end")
            qty_entry.insert(0, "1")
            price_entry.delete(0, "end")
            search_entry.focus_set()

        def add_and_close():
            add_silent()
            win.destroy()

        def add_silent():
            d = search_var.get().strip()
            q = qty_entry.get().strip()
            p = price_entry.get().strip()
            if not d or not q.isdigit() or not p.isdigit():
                return
            q_i, p_i = int(q), int(p)
            self.tree.insert("", "end", values=(d, q_i, p_i, q_i * p_i))
            self._update_totals()

        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="➕ Shto tjetër", command=add).pack(side="left", padx=8)
        ttk.Button(btn_frame, text="✅ Shto & Mbyll", command=add_and_close).pack(side="left", padx=8)
        ttk.Button(btn_frame, text="Anulo", command=win.destroy).pack(side="left", padx=8)

        price_entry.bind("<Return>", lambda e: add())
        search_entry.focus_set()

    def _delete_selected_part(self):
        sel = self.tree.selection()
        if not sel:
            return
        for item in sel:
            self.tree.delete(item)
        self._update_totals()

    def _clear_parts(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._update_totals()

    # ---------- Auto lookup by plate ----------

    def _plate_keyrelease(self, event=None):
        if self._plate_lookup_after_id:
            self.root.after_cancel(self._plate_lookup_after_id)
        self._plate_lookup_after_id = self.root.after(250, self.auto_lookup)

    def auto_lookup(self):
        plate = self.plate.get().strip().upper()
        if not plate:
            return

        conn = db_connect()
        c = conn.cursor()
        c.execute("SELECT client, model, fuel, vtype FROM vehicles WHERE plate=?", (plate,))
        row = c.fetchone()

        if not row:
            c.execute("""
                SELECT client, model, fuel, vtype
                FROM preventiv
                WHERE plate=?
                ORDER BY id DESC
                LIMIT 1
            """, (plate,))
            row = c.fetchone()
        conn.close()

        if row:
            client, model, fuel, vtype = row
            if client:
                self.client.delete(0, "end")
                self.client.insert(0, client)
            if model:
                self.model.delete(0, "end")
                self.model.insert(0, model)
            if vtype:
                self.vtype.delete(0, "end")
                self.vtype.insert(0, vtype)
            if fuel:
                try:
                    self.fuel.set(fuel)
                except Exception:
                    self.fuel.set("Tjetër")

    # ---------- SAVE ----------

    def save(self, silent=False):
        plate = self.plate.get().strip().upper()
        if not plate:
            messagebox.showwarning("Kujdes", "Vendos targën.")
            return

        parts_total = self._get_parts_total()
        labor = self._get_labor()
        grand = parts_total + labor

        client = self.client.get().strip()
        model = self.model.get().strip()
        fuel = self.fuel.get().strip()
        vtype = self.vtype.get().strip()
        notes = self.notes.get("1.0", "end").strip()

        if parts_total == 0 and labor == 0:
            messagebox.showwarning("Kujdes", "Shto pjesë ose vendos hak pune.")
            return

        conn = db_connect()
        c = conn.cursor()

        c.execute("""
            INSERT OR REPLACE INTO vehicles (plate, client, model, fuel, vtype, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (plate, client, model, fuel, vtype, datetime.now().strftime("%d/%m/%Y %H:%M:%S")))

        c.execute("""
            INSERT INTO preventiv
            (date, plate, client, model, fuel, vtype, labor, parts_total, grand_total, notes, pdf_path)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            datetime.now().strftime("%d/%m/%Y"),
            plate, client, model, fuel, vtype,
            labor, parts_total, grand, notes, ""
        ))
        preventiv_id = c.lastrowid

        for item in self.tree.get_children():
            d, q, p, t = self.tree.item(item)["values"]
            c.execute("""
                INSERT INTO preventiv_items (preventiv_id, description, qty, price, total)
                VALUES (?, ?, ?, ?, ?)
            """, (preventiv_id, str(d), int(q), int(p), int(t)))

        conn.commit()
        conn.close()

        # ✅ Gjenero PDF automatikisht pas ruajtjes
        pdf_path = self._generate_pdf_silent()

        # ✅ Përditëso pdf_path në DB
        if pdf_path:
            conn2 = db_connect()
            c2 = conn2.cursor()
            c2.execute("UPDATE preventiv SET pdf_path=? WHERE id=?", (pdf_path, preventiv_id))
            conn2.commit()
            conn2.close()

        self._update_dashboard()

        if pdf_path and not silent:
            messagebox.showinfo("✅ U ruajt", f"Preventivi u ruajt dhe PDF u gjenerua:\n{os.path.basename(pdf_path)}")
        elif not pdf_path and not silent:
            messagebox.showinfo("✅ U ruajt", "Preventivi u ruajt në bazën e të dhënave.")

        return pdf_path

    def save_and_print(self):
        """Ruan + gjeneron PDF automatikisht, pastaj hap dialog printimi."""
        pdf_path = self.save(silent=True)
        if pdf_path:
            self.last_pdf = pdf_path
            self.print_dialog()

    # ---------- PDF (Punë + Pjesë + Total poshtë djathtas, pa TVSH) ----------

    def _generate_pdf_silent(self):
        """Gjeneron PDF automatikisht pa asnjë dialog - kthehet path-i."""
        from reportlab.lib.units import mm
        from reportlab.platypus import TableStyle

        parts_total = self._get_parts_total()
        labor_total = self._get_labor()
        grand = parts_total + labor_total

        plate = self.plate.get().strip().upper()
        if not plate:
            messagebox.showwarning("Kujdes", "Vendos targën.")
            return
        if parts_total == 0 and labor_total == 0:
            messagebox.showwarning("Kujdes", "Shto pjesë ose vendos hak pune.")
            return

        filename = os.path.join(PDF_FOLDER, f"{plate}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")

        doc = SimpleDocTemplate(
            filename,
            pagesize=A4,
            leftMargin=14*mm,
            rightMargin=14*mm,
            topMargin=12*mm,
            bottomMargin=12*mm
        )
        styles = getSampleStyleSheet()
        elements = []

        # rrit fontet PDF
        styles["Title"].fontSize = 22
        styles["Title"].leading = 26
        styles["Normal"].fontSize = 12
        styles["Normal"].leading = 15
        styles["Heading2"].fontSize = 16
        styles["Heading2"].leading = 20

        elements.append(Paragraph("<b>PREVENTIV</b>", styles["Title"]))
        elements.append(Paragraph("AUTO SERVIS ALDO - TIRANË", styles["Normal"]))
        elements.append(Paragraph("KONTAKT : +355 69 34 94 214", styles["Normal"]))
        elements.append(Paragraph(f"Data: {datetime.now().strftime('%d/%m/%Y')}", styles["Normal"]))
        elements.append(Spacer(1, 8))

        elements.append(Paragraph(f"Emri: {self.client.get().strip()}", styles["Normal"]))
        elements.append(Paragraph(f"Targa: {plate}", styles["Normal"]))
        elements.append(Paragraph(f"Modeli: {self.model.get().strip()}", styles["Normal"]))
        elements.append(Paragraph(f"Karburanti: {self.fuel.get().strip()}", styles["Normal"]))
        elements.append(Paragraph(f"Lloji i makinës: {self.vtype.get().strip()}", styles["Normal"]))
        elements.append(Spacer(1, 10))

        # PUNË (vetëm Hak Pune)
        work_data = [["PUNË", "SHUMA (Lek)"], ["Hak Pune", f"{labor_total:,}"]]
        work_tbl = Table(work_data, colWidths=[110*mm, 70*mm])
        work_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
            ("GRID", (0,0), (-1,-1), 1, colors.black),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (1,1), (1,-1), "RIGHT"),
            ("FONTSIZE", (0,0), (-1,-1), 11),
        ]))
        elements.append(work_tbl)
        elements.append(Spacer(1, 10))

        # PJESË
        parts_data = [["PJESË", "SASIA", "ÇMIMI", "SHUMA"]]
        for item in self.tree.get_children():
            d, q, p, t = self.tree.item(item)["values"]
            parts_data.append([str(d), str(q), f"{int(p):,}", f"{int(t):,}"])

        parts_tbl = Table(parts_data, colWidths=[85*mm, 20*mm, 35*mm, 40*mm])
        parts_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
            ("GRID", (0,0), (-1,-1), 1, colors.black),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (1,1), (-1,-1), "RIGHT"),
            ("ALIGN", (0,0), (0,-1), "LEFT"),
            ("FONTSIZE", (0,0), (-1,-1), 10),
        ]))
        elements.append(parts_tbl)
        elements.append(Spacer(1, 10))

        # Blloku totals poshtë djathtas (PA TVSH)
        totals_block = Table(
            [
                ["Totali Punë", f"{labor_total:,} Lek"],
                ["Totali Pjesë", f"{parts_total:,} Lek"],
                ["TOTALI FINAL", f"{grand:,} Lek"],
            ],
            colWidths=[55*mm, 55*mm],
            hAlign="RIGHT"
        )
        totals_block.setStyle(TableStyle([
            ("GRID", (0,0), (-1,-1), 1, colors.black),
            ("BACKGROUND", (0,2), (-1,2), colors.lightgrey),
            ("FONTNAME", (0,2), (-1,2), "Helvetica-Bold"),
            ("ALIGN", (1,0), (1,-1), "RIGHT"),
            ("FONTSIZE", (0,0), (-1,-1), 11),
        ]))
        elements.append(totals_block)

        notes = self.notes.get("1.0", "end").strip()
        if notes:
            elements.append(Spacer(1, 10))
            elements.append(Paragraph(f"Shënime: {notes}", styles["Normal"]))

        doc.build(elements)
        self.last_pdf = filename
        return filename

    # ---------- PRINT (Dialog + fallback) ----------

    def print_dialog(self):
        if not self.last_pdf or not os.path.exists(self.last_pdf):
            messagebox.showerror("Gabim", "Nuk ka PDF të gjeneruar.")
            return

        printers = [p[2] for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        )]

        if not printers:
            messagebox.showerror("Gabim", "Nuk u gjet asnjë printer i instaluar!")
            return

        try:
            default_printer = win32print.GetDefaultPrinter()
        except Exception:
            default_printer = printers[0]

        win = tk.Toplevel(self.root)
        win.title("Printo")
        win.geometry("640x260")
        win.resizable(False, False)

        tk.Label(win, text="Zgjidh printerin:", font=("Segoe UI", 12, "bold")).pack(pady=(14, 8))

        printer_var = tk.StringVar(value=default_printer if default_printer in printers else printers[0])
        cb = ttk.Combobox(win, textvariable=printer_var, values=printers, state="readonly", width=70)
        cb.pack(pady=6)

        tk.Label(win, text=f"PDF: {os.path.basename(self.last_pdf)}", font=("Segoe UI", 11)).pack(pady=6)

        def do_print():
            prn = printer_var.get().strip()
            if not prn:
                messagebox.showerror("Gabim", "Zgjidh një printer.")
                return

            try:
                win32api.ShellExecute(0, "printto", self.last_pdf, f'"{prn}"', ".", 0)
                messagebox.showinfo("OK", f"U dërgua për print te: {prn}")
                win.destroy()
                return
            except Exception:
                pass

            try:
                os.startfile(self.last_pdf)
                messagebox.showinfo("Info", "Printimi direkt dështoi. PDF u hap — shtyp CTRL+P.")
                win.destroy()
            except Exception as e:
                messagebox.showerror("Gabim", f"Nuk u hap dot PDF:\n{e}")

        ttk.Button(win, text="Printo", command=do_print).pack(pady=14)

    # ---------- HISTORY ----------

    def show_history(self):
        plate = self.plate.get().strip().upper()
        if not plate:
            messagebox.showwarning("Kujdes", "Shkruaj targën për të parë historikun.")
            return

        conn = db_connect()
        df = pd.read_sql_query("""
            SELECT id, date, grand_total, pdf_path
            FROM preventiv
            WHERE plate=?
            ORDER BY id DESC
        """, conn, params=(plate,))
        conn.close()

        if df.empty:
            messagebox.showinfo("Historiku", "Nuk ka preventiva të ruajtura për këtë targë.")
            return

        win = tk.Toplevel(self.root)
        win.title(f"Historiku - {plate}")
        win.geometry("860x460")

        tree = ttk.Treeview(win, columns=("id", "date", "total", "pdf"), show="headings")
        tree.heading("id", text="ID")
        tree.heading("date", text="Data")
        tree.heading("total", text="Totali (Lek)")
        tree.heading("pdf", text="PDF")

        tree.column("id", width=70, anchor="center")
        tree.column("date", width=140, anchor="center")
        tree.column("total", width=160, anchor="e")
        tree.column("pdf", width=450, anchor="w")
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        for _, r in df.iterrows():
            tree.insert("", "end", values=(int(r["id"]), r["date"], f"{int(r['grand_total']):,}", r["pdf_path"] or ""))

        btns = tk.Frame(win)
        btns.pack(pady=8)

        def get_selected_pdf():
            sel = tree.selection()
            if not sel:
                return None
            return tree.item(sel[0])["values"][3]

        def open_pdf():
            pdf = get_selected_pdf()
            if pdf and os.path.exists(pdf):
                os.startfile(pdf)
            else:
                messagebox.showwarning("Kujdes", "PDF nuk ekziston.")

        ttk.Button(btns, text="Hap PDF", command=open_pdf).pack(side="left", padx=6)

    # ---------- DASHBOARD ----------

    def _update_dashboard(self):
        conn = db_connect()
        df = pd.read_sql_query("SELECT date, grand_total FROM preventiv", conn)
        conn.close()

        self.ax.clear()
        self.ax.set_title("Xhiro ditore", fontsize=10)

        if not df.empty:
            df["date"] = pd.to_datetime(df["date"], format="%d/%m/%Y", errors="coerce")
            df = df.dropna(subset=["date"]).sort_values("date")
            cutoff = pd.Timestamp.now() - pd.Timedelta(days=30)
            df = df[df["date"] >= cutoff]
            daily = df.groupby(df["date"].dt.date)["grand_total"].sum()

            x = list(daily.index)
            y = list(daily.values)
            self.ax.plot(x, y)
            self.ax.tick_params(axis="x", labelrotation=45, labelsize=8)
            self.ax.tick_params(axis="y", labelsize=8)

        self.fig.tight_layout()
        self.canvas.draw()

    # ---------- REPORT ----------

    def report_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Raport - Zgjidh periudhën")
        win.geometry("760x560")

        top = tk.Frame(win)
        top.pack(fill="x", padx=10, pady=10)

        tk.Label(top, text="Nga (DD/MM/YYYY):", font=("Segoe UI", 12)).grid(row=0, column=0, padx=6, pady=6, sticky="e")
        start_e = tk.Entry(top, width=14, font=("Segoe UI", 12))
        start_e.grid(row=0, column=1, padx=6, pady=6)

        tk.Label(top, text="Deri (DD/MM/YYYY):", font=("Segoe UI", 12)).grid(row=0, column=2, padx=6, pady=6, sticky="e")
        end_e = tk.Entry(top, width=14, font=("Segoe UI", 12))
        end_e.grid(row=0, column=3, padx=6, pady=6)

        today = datetime.now()
        start_e.insert(0, today.replace(day=1).strftime("%d/%m/%Y"))
        end_e.insert(0, today.strftime("%d/%m/%Y"))

        summary = tk.Label(win, text="", font=("Segoe UI", 12, "bold"))
        summary.pack(pady=(0, 6))

        tree = ttk.Treeview(win, columns=("date", "plate", "total"), show="headings")
        tree.heading("date", text="Data")
        tree.heading("plate", text="Targa")
        tree.heading("total", text="Totali (Lek)")
        tree.column("date", width=160, anchor="center")
        tree.column("plate", width=160, anchor="center")
        tree.column("total", width=200, anchor="e")
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        def load():
            for i in tree.get_children():
                tree.delete(i)

            start = start_e.get().strip()
            end = end_e.get().strip()

            conn = db_connect()
            df = pd.read_sql_query("""
                SELECT date, plate, grand_total
                FROM preventiv
                WHERE date BETWEEN ? AND ?
                ORDER BY date ASC
            """, conn, params=(start, end))
            conn.close()

            if df.empty:
                summary.config(text="Nuk ka të dhëna në këtë periudhë.")
                return

            total_sum = int(df["grand_total"].sum())
            count = int(len(df))
            avg = int(total_sum / count) if count else 0
            summary.config(text=f"Preventiva: {count}   |   Xhiro: {total_sum:,} Lek   |   Mesatare: {avg:,} Lek")

            for _, r in df.iterrows():
                tree.insert("", "end", values=(r["date"], r["plate"], f"{int(r['grand_total']):,}"))

        ttk.Button(top, text="Gjenero", command=load).grid(row=0, column=4, padx=10)
        load()

    # ---------- Excel export ----------

    def export_excel(self):
        conn = db_connect()
        df = pd.read_sql_query("""
            SELECT date, plate, client, model, fuel, vtype, labor, parts_total, grand_total, notes, pdf_path
            FROM preventiv
            ORDER BY id DESC
        """, conn)
        conn.close()

        if df.empty:
            messagebox.showinfo("Info", "Nuk ka të dhëna për eksport.")
            return

        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return

        df.to_excel(path, index=False)
        messagebox.showinfo("OK", "Excel u krijua me sukses!")


def open_main_app():
    root = tk.Tk()
    PreventivApp(root)
    root.mainloop()


if __name__ == "__main__":
    init_db()
    login_root = tk.Tk()
    LoginWindow(login_root)
    login_root.mainloop()