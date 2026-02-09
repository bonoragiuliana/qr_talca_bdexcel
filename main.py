import os
import json
import textwrap
import sqlite3
from datetime import datetime
from dateutil.relativedelta import relativedelta

import pandas as pd
import qrcode
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import messagebox, filedialog, ttk


# ---------- CONFIGURACIÓN ----------
EXCEL_PATH = "productos.xlsx"
HOJA_PRODUCTOS = "productos"
CACHE_FILE = "config.json"
DB_PATH = "talca_qr.db"  # Base SQLite local


# ---------- SQLITE ----------
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("PRAGMA journal_mode=WAL;")
    cur.execute("PRAGMA synchronous=NORMAL;")

    # Crear tabla (solo 6 campos)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS pallet_scans (
        descripcion   TEXT NOT NULL,
        nro_serie     INTEGER NOT NULL,
        id_producto   TEXT NOT NULL,
        lote          TEXT NOT NULL,
        creacion      TEXT NOT NULL,
        vencimiento   TEXT NOT NULL,
        UNIQUE(id_producto, nro_serie, lote)
    )
    """)

    # Índices útiles
    cur.execute("CREATE INDEX IF NOT EXISTS idx_lote ON pallet_scans(lote)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_id_producto ON pallet_scans(id_producto)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_creacion ON pallet_scans(creacion)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_vencimiento ON pallet_scans(vencimiento)")

    conn.commit()
    return conn


def parse_qr_payload(raw: str) -> dict:
    """
    Formato esperado (nuevo, robusto):
    NS=000001|PRD=12|DSC=Descripcion...|LOT=090226|FEC=2026-02-09|VTO=2026-08-09

    Devuelve:
    descripcion, nro_serie, id_producto, lote, creacion, vencimiento
    """
    raw = raw.strip()

    # 1) Formato nuevo K=V con "|"
    if "|" in raw and "=" in raw:
        parts = raw.split("|")
        data = {}
        for p in parts:
            if "=" in p:
                k, v = p.split("=", 1)
                data[k.strip()] = v.strip()

        required = ["NS", "PRD", "DSC", "LOT", "FEC", "VTO"]
        missing = [k for k in required if k not in data or not data[k]]
        if missing:
            raise ValueError(f"QR inválido, faltan campos: {', '.join(missing)}")

        return {
            "descripcion": data["DSC"],
            "nro_serie": int(data["NS"]),
            "id_producto": data["PRD"],
            "lote": data["LOT"],
            "creacion": data["FEC"],
            "vencimiento": data["VTO"],
        }

    # 2) (Fallback) Formato viejo con saltos de línea
    if "\n" in raw:
        lines = [l.strip() for l in raw.splitlines() if l.strip()]

        def pick(prefix):
            for l in lines:
                if l.lower().startswith(prefix.lower()):
                    return l.split(":", 1)[1].strip()
            return None

        ns = pick("Num de serie") or pick("N de serie") or pick("N° de serie")
        prd = pick("ID producto")
        lote = pick("Lote")
        cre = pick("Creacion") or pick("Creación")
        vto = pick("Vencimiento")

        known_prefixes = (
            "num de serie", "n de serie", "n° de serie",
            "id producto", "lote", "creacion", "creación", "vencimiento"
        )
        desc_candidates = [l for l in lines if not l.lower().startswith(known_prefixes)]
        desc = desc_candidates[0] if desc_candidates else ""

        if not all([ns, prd, lote, cre, vto]) or not desc:
            raise ValueError("QR inválido (formato viejo): no pude extraer todos los campos.")

        def normalize_date(s):
            s = s.strip()
            if "/" in s:
                try:
                    d = datetime.strptime(s, "%d/%m/%y").date()
                    return d.isoformat()
                except:
                    return s
            return s

        return {
            "descripcion": desc,
            "nro_serie": int(ns),
            "id_producto": prd,
            "lote": lote,
            "creacion": normalize_date(cre),
            "vencimiento": normalize_date(vto),
        }

    raise ValueError("QR inválido: formato no reconocido.")


def save_scan(conn: sqlite3.Connection, raw_payload: str):
    data = parse_qr_payload(raw_payload)

    cur = conn.cursor()
    try:
        cur.execute("""
        INSERT INTO pallet_scans (descripcion, nro_serie, id_producto, lote, creacion, vencimiento)
        VALUES (?, ?, ?, ?, ?, ?)
        """, (
            data["descripcion"],
            data["nro_serie"],
            data["id_producto"],
            data["lote"],
            data["creacion"],
            data["vencimiento"]
        ))
        conn.commit()
        return "OK", f"✅ Guardado: {data['id_producto']} - Serie {data['nro_serie']} - Lote {data['lote']}"
    except sqlite3.IntegrityError:
        return "DUP", f"⚠️ YA REGISTRADO: {data['id_producto']} - Serie {data['nro_serie']} - Lote {data['lote']}"


def fetch_latest_scans(conn: sqlite3.Connection, limit=300):
    cur = conn.cursor()
    cur.execute("""
        SELECT descripcion, nro_serie, id_producto, lote, creacion, vencimiento
        FROM pallet_scans
        ORDER BY rowid DESC
        LIMIT ?
    """, (limit,))
    return cur.fetchall()


# ---------- FUNCIONES EXCEL ----------
def cargar_productos():
    df = pd.read_excel(EXCEL_PATH, sheet_name=HOJA_PRODUCTOS)
    df.columns = df.columns.str.strip()
    return df


def guardar_productos(df):
    df.to_excel(EXCEL_PATH, sheet_name=HOJA_PRODUCTOS, index=False)


def obtener_productos():
    df = cargar_productos()
    return list(zip(df["id_producto"], df["descripcion"]))


def dividir_texto(texto, max_caracteres):
    return textwrap.wrap(texto, width=max_caracteres)


# ---------- CACHÉ ----------
def guardar_config(seleccion, cantidad):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump({"producto": seleccion, "cantidad": cantidad}, f, ensure_ascii=False)


def cargar_config():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


# ---------- GENERAR PDF ----------
def generar_y_imprimir_qrs(id_producto, descripcion, cantidad):
    df = cargar_productos()
    fila = df[df["id_producto"] == id_producto].index
    if fila.empty:
        messagebox.showerror("Error", "Producto no encontrado.")
        return

    nro_serie = int(df.loc[fila[0], "ultimo_nro_serie"])

    fecha_actual = datetime.now()
    fec_iso = fecha_actual.strftime("%Y-%m-%d")
    vto_iso = (fecha_actual + relativedelta(months=6)).strftime("%Y-%m-%d")

    fecha_str = fecha_actual.strftime("%d/%m/%y")
    fecha_venc_str = (fecha_actual + relativedelta(months=6)).strftime("%d/%m/%y")

    numero_lote = fecha_actual.strftime("%d%m%y")

    pdf_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF", "*.pdf")],
        initialfile=f"qr_lote_{numero_lote}.pdf"
    )
    if not pdf_path:
        return

    c = canvas.Canvas(pdf_path, pagesize=A4)
    _, alto = A4

    # 4 slots por hoja (como ya venías usando)
    y_positions = [alto - 230, alto - 430, alto - 630, alto - 830]

    x_qr = 40
    qr_size = 215
    text_x = x_qr + qr_size + 40
    posicion_actual = 0

    desc_clean = str(descripcion).replace("\n", " ").replace("|", "/").replace("=", "-").strip()
    if len(desc_clean) > 90:
        desc_clean = desc_clean[:90]

    for _ in range(cantidad):
        nro_serie += 1

        payload_qr = (
            f"NS={nro_serie:06d}"
            f"|PRD={id_producto}"
            f"|DSC={desc_clean}"
            f"|LOT={numero_lote}"
            f"|FEC={fec_iso}"
            f"|VTO={vto_iso}"
        )

        qr = qrcode.make(payload_qr)
        qr_path = f"temp_qr_{id_producto}_{nro_serie}.png"
        qr.save(qr_path)

        # Dos copias por número de serie
        for _ in range(2):
            y = y_positions[posicion_actual]
            c.drawImage(qr_path, x_qr, y, width=qr_size, height=qr_size)

            titulo_lineas = dividir_texto(descripcion, 40)
            resto_lineas = [
                f"N° de serie: {nro_serie}",
                f"ID producto: {id_producto}",
                f"Lote: {numero_lote}",
                f"Creación: {fecha_str}",
                f"Vencimiento: {fecha_venc_str}",
            ]

            titulo_height = len(titulo_lineas) * 18
            resto_height = len(resto_lineas) * 15
            total_height = titulo_height + resto_height

            centro_qr_y = y + qr_size / 2
            text_y = centro_qr_y + total_height / 2

            c.setFont("Helvetica-Bold", 15)
            for i, linea_txt in enumerate(titulo_lineas):
                c.drawString(text_x, text_y - i * 20, linea_txt)

            offset = titulo_height

            c.setFont("Helvetica-Bold", 18)
            c.drawString(text_x, text_y - offset, resto_lineas[0])
            offset += 20

            c.setFont("Helvetica", 15)
            for linea_txt in resto_lineas[1:]:
                c.drawString(text_x, text_y - offset, linea_txt)
                offset += 15

            posicion_actual += 1
            if posicion_actual == 4:
                c.showPage()
                posicion_actual = 0

        os.remove(qr_path)

    c.save()

    df.loc[fila[0], "ultimo_nro_serie"] = nro_serie
    guardar_productos(df)

    messagebox.showinfo("PDF generado", f"El archivo se guardó correctamente:\n{pdf_path}")


# ---------- INTERFAZ ----------
conn_db = init_db()

root = tb.Window(themename="minty")
root.title("Sistema QRs – Talca")
root.geometry("980x560")

notebook = tb.Notebook(root)
notebook.pack(fill="both", expand=True, padx=10, pady=10)

tab_gen = tb.Frame(notebook, padding=20)
tab_scan = tb.Frame(notebook, padding=20)
tab_view = tb.Frame(notebook, padding=20)

notebook.add(tab_gen, text="Generar QRs")
notebook.add(tab_scan, text="Escanear pallet")
notebook.add(tab_view, text="Registros")


# ----- TAB 1: Generar -----
tb.Label(tab_gen, text="Generador de QRs", font=("Segoe UI", 18, "bold")).pack(pady=10)
tb.Label(tab_gen, text="Seleccioná un producto:", font=("Segoe UI", 12)).pack(pady=5)

productos = obtener_productos()
producto_dict = {f"{d} (ID: {i})": (i, d) for i, d in productos}

combo = tb.Combobox(tab_gen, values=list(producto_dict.keys()), width=80)
combo.pack(pady=4)

tb.Label(tab_gen, text="Cantidad de números de serie:", font=("Segoe UI", 12)).pack(pady=10)
cantidad_entry = tb.Entry(tab_gen, width=12)
cantidad_entry.pack()


def al_hacer_click_generar():
    if not combo.get():
        messagebox.showwarning("Aviso", "Seleccioná un producto.")
        return

    try:
        cantidad = int(cantidad_entry.get())
        if cantidad <= 0:
            raise ValueError
    except:
        messagebox.showwarning("Aviso", "Cantidad inválida.")
        return

    pid, desc = producto_dict[combo.get()]
    generar_y_imprimir_qrs(pid, desc, cantidad)
    guardar_config(combo.get(), cantidad)


tb.Button(tab_gen, text="GENERAR", bootstyle=SUCCESS, command=al_hacer_click_generar).pack(pady=18)


# ----- TAB 3: Registros (tabla tipo Excel) -----
tb.Label(tab_view, text="Registros escaneados", font=("Segoe UI", 18, "bold")).pack(pady=10)

count_var = tb.StringVar(value="Cargando…")
tb.Label(tab_view, textvariable=count_var, font=("Segoe UI", 11)).pack(pady=(0, 8))

table_frame = tb.Frame(tab_view)
table_frame.pack(fill="both", expand=True)

columns = ("descripcion", "nro_serie", "id_producto", "lote", "creacion", "vencimiento")

tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=16)
tree.pack(side="left", fill="both", expand=True)

tree.heading("descripcion", text="Descripción")
tree.heading("nro_serie", text="N° Serie")
tree.heading("id_producto", text="ID Producto")
tree.heading("lote", text="Lote")
tree.heading("creacion", text="Creación")
tree.heading("vencimiento", text="Vencimiento")

tree.column("descripcion", width=420, anchor="w")
tree.column("nro_serie", width=90, anchor="center")
tree.column("id_producto", width=110, anchor="center")
tree.column("lote", width=90, anchor="center")
tree.column("creacion", width=120, anchor="center")
tree.column("vencimiento", width=120, anchor="center")

scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
scroll_y.pack(side="right", fill="y")
tree.configure(yscrollcommand=scroll_y.set)


def refresh_table(limit=300):
    for item in tree.get_children():
        tree.delete(item)

    rows = fetch_latest_scans(conn_db, limit=limit)

    # Mostrar en pantalla de más viejo -> más nuevo
    for r in reversed(rows):
        tree.insert("", "end", values=r)

    count_var.set(f"Mostrando {len(rows)} registros (últimos)")


btns_frame = tb.Frame(tab_view)
btns_frame.pack(pady=10)

tb.Button(btns_frame, text="Refrescar", bootstyle=INFO, command=refresh_table).pack(side="left", padx=6)
tb.Button(btns_frame, text="Últimos 100", bootstyle=SECONDARY, command=lambda: refresh_table(100)).pack(side="left", padx=6)
tb.Button(btns_frame, text="Últimos 500", bootstyle=SECONDARY, command=lambda: refresh_table(500)).pack(side="left", padx=6)

refresh_table()


# ----- TAB 2: Escanear (auto con Enter) -----
tb.Label(tab_scan, text="Escaneo de pallets", font=("Segoe UI", 18, "bold")).pack(pady=10)
tb.Label(
    tab_scan,
    text="Posicionate en el campo y escaneá.\nEl guardado es automático cuando el escáner envía Enter.",
    font=("Segoe UI", 11)
).pack(pady=6)

scan_var = tb.StringVar()
entry_scan = tb.Entry(tab_scan, textvariable=scan_var, width=95, font=("Segoe UI", 14))
entry_scan.pack(pady=12)
entry_scan.focus_set()

status_var = tb.StringVar(value="Listo para escanear…")
status_lbl = tb.Label(tab_scan, textvariable=status_var, font=("Segoe UI", 12))
status_lbl.pack(pady=8)


def on_scan_return(event=None):
    raw = scan_var.get().strip()
    scan_var.set("")
    entry_scan.focus_set()

    if not raw:
        return

    try:
        state, msg = save_scan(conn_db, raw)
        status_var.set(msg)
        root.bell()
        refresh_table()  # ✅ actualiza la pestaña "Registros"
    except Exception as e:
        status_var.set(f"❌ ERROR: {e}")
        root.bell()


entry_scan.bind("<Return>", on_scan_return)


# Mantener foco en el campo al volver a la pestaña de escaneo
def on_tab_change(event=None):
    try:
        current = notebook.tab(notebook.select(), "text")
        if current == "Escanear pallet":
            entry_scan.focus_set()
    except:
        pass


notebook.bind("<<NotebookTabChanged>>", on_tab_change)


def on_close():
    try:
        conn_db.close()
    except:
        pass
    root.destroy()


root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()
