import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import math
import time

from docx import Document
from docx.shared import RGBColor, Pt, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT

try:
import win32com.client
import pythoncom
PDF_AVAILABLE = True
except:
PDF_AVAILABLE = False

# ============================================================

# COLORES

# ============================================================

COLOR_HEADER  = “1F4E79”   # azul oscuro
COLOR_HEADER2 = “2E75B6”   # azul medio
COLOR_ROW_ALT = “D6E4F0”   # azul claro filas alternas

# ============================================================

# MAPAS

# ============================================================

REGION_MAP = {
‘RM’:   (‘Las Condes’, ‘Región Metropolitana’),
‘Viña’: (‘Viña del Mar’, ‘Región de Valparaíso’),
‘VALP’: (‘Valparaíso’, ‘Región de Valparaíso’),
‘BIO’:  (‘Concepción’, ‘Región del Biobío’),
‘ARAU’: (‘Temuco’, ‘Región de La Araucanía’),
‘OHIG’: (“Rancagua”, “Región del Libertador General Bernardo O’Higgins”),
‘MAULE’:(‘Talca’, ‘Región del Maule’),
‘LOS R’:(‘Valdivia’, ‘Región de Los Ríos’),
‘LOS L’:(‘Puerto Montt’, ‘Región de Los Lagos’),
‘AYSN’: (‘Coyhaique’, ‘Región de Aysén’),
‘MAG’:  (‘Punta Arenas’, ‘Región de Magallanes’),
‘ATF’:  (‘Iquique’, ‘Región de Tarapacá’),
‘ANT’:  (‘Antofagasta’, ‘Región de Antofagasta’),
‘ATA’:  (‘Copiapó’, ‘Región de Atacama’),
‘COQ’:  (‘La Serena’, ‘Región de Coquimbo’),
}

# Cargos fijos netos (sin IVA) por producto

CARGO_FIJO_NETO = {
‘BAF 200’:  12597,
‘BAF 300’:  15118,
‘BAF 400’:  17639,
‘BAF 500’:  18479,
‘BAF 600’:  19319,
‘BAF 700’:  21840,
‘BAF 800’:  23521,
‘BAF 940’:  32765,
‘BAF 2000’: 19319,
‘IPTV’:     19319,
‘LINEAS’:   11756,
}
CARGO_FIJO_BRUTO = {
‘BAF 200’:  14990,
‘BAF 300’:  17990,
‘BAF 400’:  22990,
‘BAF 500’:  21990,
‘BAF 600’:  22990,
‘BAF 700’:  25990,
‘BAF 800’:  27990,
‘BAF 940’:  38990,
‘BAF 2000’: 22990,
‘IPTV’:     22990,
‘LINEAS’:   13990,
}
PRODUCTOS_ORDER = [‘BAF 200’,‘BAF 300’,‘BAF 400’,‘BAF 500’,‘BAF 600’,‘BAF 700’,‘BAF 800’,‘BAF 940’,‘BAF 2000’,‘IPTV’,‘LINEAS’]

def tramo_pct(pct_cumpl):
if pct_cumpl < 0.60: return 0.0
if pct_cumpl < 0.70: return 0.15
if pct_cumpl < 0.80: return 0.20
if pct_cumpl < 0.90: return 0.30
if pct_cumpl < 1.00: return 0.40
if pct_cumpl < 1.30: return 0.50
return 0.60

def bono_baf(total_fo_tv):
t = int(total_fo_tv)
if t == 20: return 75000
if 21 <= t <= 22: return 100000
if 23 <= t <= 25: return 150000
if 26 <= t <= 29: return 200000
if 30 <= t <= 33: return 250000
if 34 <= t <= 37: return 300000
if 38 <= t <= 43: return 400000
if 44 <= t <= 48: return 450000
if 49 <= t <= 52: return 500000
if 53 <= t <= 56: return 550000
if 57 <= t <= 60: return 650000
if 61 <= t <= 66: return 750000
if t >= 67: return 850000
return 0

def tramo_emitidas(cant):
if cant <= 24: return 1000
if cant <= 40: return 2500
return 3500

def nz(v):
if pd.isna(v) or str(v).strip().lower() in (‘nan’,’’):
return 0
try:
s = str(v).replace(’$’,’’).strip()
if ‘,’ in s: s = s.replace(’.’,’’).replace(’,’,’.’)
return float(s)
except:
return 0

def fmt_m(v):
if not v: return “$ -”
r = math.floor(v + 0.5)
return f”$ {r:,}”.replace(”,”,”.”)

def fmt_pct(v):
return f”{int(round(v * 100))}%”

# ============================================================

# UTILIDADES DOCX

# ============================================================

def set_cell_margins(cell, top=80, start=150, bottom=80, end=150):
tc = cell._tc
tcPr = tc.get_or_add_tcPr()
tcMar = OxmlElement(‘w:tcMar’)
for m in [(‘top’,top),(‘start’,start),(‘bottom’,bottom),(‘end’,end)]:
node = OxmlElement(f’w:{m[0]}’)
node.set(qn(‘w:w’), str(m[1]))
node.set(qn(‘w:type’), ‘dxa’)
tcMar.append(node)
tcPr.append(tcMar)

def set_cell_bg(cell, hex_color):
tcPr = cell._tc.get_or_add_tcPr()
shd = OxmlElement(‘w:shd’)
shd.set(qn(‘w:fill’), hex_color)
shd.set(qn(‘w:val’), ‘clear’)
shd.set(qn(‘w:color’), ‘auto’)
tcPr.append(shd)

def set_table_borders(table):
tblPr = table._element.xpath(‘w:tblPr’)
if tblPr:
tblBorders = OxmlElement(‘w:tblBorders’)
for b_name in [‘top’,‘left’,‘bottom’,‘right’,‘insideH’,‘insideV’]:
b = OxmlElement(f’w:{b_name}’)
b.set(qn(‘w:val’), ‘single’)
b.set(qn(‘w:sz’), ‘4’)
b.set(qn(‘w:color’), ‘AAAAAA’)
tblBorders.append(b)
tblPr[0].append(tblBorders)

# ============================================================

# APP

# ============================================================

class MasivoApp:
def **init**(self, root):
self.root = root
self.root.title(“Masivo — Generador Comisiones”)
self.root.geometry(“700x880”)
self.root.configure(bg=”#0f0f0f”)
self.archivos = {“sim”: None, “cfijo”: None, “cmov”: None, “pla”: None}
self.ejecutivos = {}   # nombre -> dict datos
self.ventas_fijo = []
self.ventas_movil = []
self.setup_ui()

```
# ----------------------------------------------------------
#  UI
# ----------------------------------------------------------
def setup_ui(self):
    bg, fg, blue = "#0f0f0f", "#f0f0f0", "#1F6EC5"
    tk.Label(self.root, text="Masivo — Comisiones", font=("Arial",22,"bold"),
             bg=bg, fg=fg).pack(pady=20)

    f = tk.LabelFrame(self.root, text=" 1. Archivos ", bg=bg, fg=blue,
                      font=("Arial",10,"bold"))
    f.pack(fill="x", padx=40, pady=5)

    self.btn_sim  = self._btn(f, "Simulador Masivo (.xlsm)", "sim")
    self.btn_cfijo= self._btn(f, "Cierre Masivo Fijo (.xlsx)", "cfijo")
    self.btn_cmov = self._btn(f, "Cierre Masivo Móvil (.xlsx)", "cmov")
    self.btn_pla  = self._btn(f, "Plantilla DC Base (.docx)", "pla")

    tk.Label(self.root, text=" 2. Selección de Ejecutivos ", bg=bg, fg=blue,
             font=("Arial",10,"bold")).pack(padx=40, anchor="w", pady=(10,0))
    self.listbox = tk.Listbox(self.root, bg="#1a1a1a", fg="white", height=10,
                              selectmode="multiple", borderwidth=0,
                              highlightthickness=1, highlightbackground="#333")
    self.listbox.pack(fill="x", padx=40, pady=5)

    f_fmt = tk.Frame(self.root, bg=bg)
    f_fmt.pack(pady=8)
    self.var_fmt = tk.StringVar(value="docx")
    tk.Radiobutton(f_fmt, text="Word (.docx)", variable=self.var_fmt, value="docx",
                   bg=bg, fg=fg, selectcolor=bg).pack(side="left", padx=20)
    tk.Radiobutton(f_fmt, text="PDF (.pdf)", variable=self.var_fmt, value="pdf",
                   bg=bg, fg=fg, selectcolor=bg).pack(side="left", padx=20)

    self.btn_gen = tk.Button(self.root, text="⚡ GENERAR SELECCIONADOS",
                             bg=blue, fg="white", font=("Arial",12,"bold"),
                             command=self.generar, state=tk.DISABLED, width=30)
    self.btn_gen.pack(pady=10, ipady=5)

    self.log_box = tk.Text(self.root, height=10, bg="#000", fg="#0f0",
                           font=("Consolas",9), state=tk.DISABLED)
    self.log_box.pack(fill="both", padx=40, pady=10)

def _btn(self, parent, label, key):
    f = tk.Frame(parent, bg="#0f0f0f"); f.pack(fill="x", pady=2)
    tk.Button(f, text=label, bg="#222", fg="white", width=26,
              command=lambda: self.cargar(key)).pack(side="left", padx=5)
    lbl = tk.Label(f, text="Esperando...", bg="#0f0f0f", fg="#555", font=("Arial",8))
    lbl.pack(side="left")
    return lbl

def log(self, m):
    self.log_box.config(state=tk.NORMAL)
    self.log_box.insert(tk.END, f"> {m}\n")
    self.log_box.see(tk.END)
    self.log_box.config(state=tk.DISABLED)
    self.root.update()

def cargar(self, key):
    r = filedialog.askopenfilename()
    if r:
        self.archivos[key] = r
        getattr(self, f"btn_{key}").config(text=os.path.basename(r), fg="#22c55e")
        if key == "sim":
            self.cargar_ejecutivos()
        if all(self.archivos.values()):
            self.btn_gen.config(state=tk.NORMAL)

def cargar_ejecutivos(self):
    try:
        df = pd.read_excel(self.archivos["sim"], sheet_name='Datos', header=None)
        self.listbox.delete(0, tk.END)
        self.listbox.insert(tk.END, "--- TODOS ---")
        skip = {'nan','NOMBRE_COMPLETO',''}
        cargos_ok = {'Junior','Senior'}
        for _, row in df.iterrows():
            nombre = str(row[1]).strip()   # col B = APODO
            cargo  = str(row[3]).strip()   # col D = Detalle Cargo
            if nombre not in skip and cargo in cargos_ok:
                self.listbox.insert(tk.END, nombre)
        self.log("Lista de ejecutivos cargada.")
    except Exception as e:
        self.log(f"Error leyendo Simulador: {e}")

# ----------------------------------------------------------
#  GENERACIÓN PRINCIPAL
# ----------------------------------------------------------
def generar(self):
    idxs = self.listbox.curselection()
    if not idxs:
        return messagebox.showwarning("!", "Selecciona ejecutivos.")

    self.log("Leyendo archivos...")
    try:
        self._leer_datos()
    except Exception as e:
        self.log(f"Error leyendo datos: {e}")
        return

    dest = filedialog.askdirectory(title="Carpeta de destino")
    if not dest: return
    dest = os.path.abspath(dest)

    seleccionados = [self.listbox.get(i) for i in idxs]
    if "--- TODOS ---" in seleccionados:
        seleccionados = list(self.ejecutivos.keys())

    fmt = self.var_fmt.get()
    if fmt == "pdf":
        try:
            pythoncom.CoInitialize()
        except:
            pass

    for nombre in seleccionados:
        if nombre not in self.ejecutivos:
            self.log(f"Sin datos: {nombre}")
            continue
        d = self.ejecutivos[nombre]
        path_docx = os.path.join(dest, f"Comisiones_{nombre.replace(' ','_')}.docx")
        try:
            vf = [v for v in self.ventas_fijo  if v['ej'].upper() == nombre.upper()]
            vm = [v for v in self.ventas_movil if v['ej'].upper() == nombre.upper()]
            self.crear_doc(d, vf, vm, path_docx)
            if fmt == "pdf":
                try:
                    abs_docx = os.path.abspath(path_docx)
                    abs_pdf  = abs_docx.replace('.docx','.pdf')
                    word = win32com.client.DispatchEx("Word.Application")
                    word.Visible = False
                    wd = word.Documents.Open(abs_docx)
                    wd.SaveAs(abs_pdf, FileFormat=17)
                    wd.Close(); word.Quit()
                    time.sleep(0.5)
                    os.remove(path_docx)
                    self.log(f"✓ PDF: {nombre}")
                except Exception as e:
                    self.log(f"x Error PDF {nombre}: {e}")
                    try: word.Quit()
                    except: pass
            else:
                self.log(f"✓ Word: {nombre}")
        except Exception as e:
            self.log(f"x Error {nombre}: {e}")

    try: os.startfile(dest)
    except: pass
    messagebox.showinfo("Masivo", "Generación finalizada.")

# ----------------------------------------------------------
#  LECTURA DE DATOS
# ----------------------------------------------------------
def _leer_datos(self):
    self.ejecutivos = {}
    self.ventas_fijo = []
    self.ventas_movil = []

    # ---- Simulador hoja Datos ----
    df_d = pd.read_excel(self.archivos["sim"], sheet_name='Datos', header=None)
    # ---- Simulador hoja Pago ----
    df_p = pd.read_excel(self.archivos["sim"], sheet_name='Pago', header=None)
    # ---- Simulador hoja Fijas ----
    df_f = pd.read_excel(self.archivos["sim"], sheet_name='Fijas', header=None)
    # ---- Simulador hoja Moviles ----
    df_mv = pd.read_excel(self.archivos["sim"], sheet_name='Moviles', header=None)

    # Construir índice Pago por APODO (col 0 = nombre en Pago)
    pago_idx = {}
    for i, row in df_p.iterrows():
        n = str(row[0]).strip()
        if n and n not in ('nan','Nombre','SENIOR','JUNIOR','Junior','Senior',
                           'EJECUTIVOS','SUPERVISORES','PROMOTORES DE VENTAS','LIDERES',
                           'Nombre'):
            pago_idx[n.upper()] = row

    # Construir índice Fijas por ASESOR (col 1)
    fijas_idx = {}
    for i, row in df_f.iterrows():
        if i < 2: continue
        n = str(row[1]).strip()
        if n and n != 'nan':
            fijas_idx[n.upper()] = row

    # Construir índice Moviles por EJECUTIVO (col 1)
    movil_idx = {}
    for i, row in df_mv.iterrows():
        if i < 2: continue
        n = str(row[1]).strip()
        if n and n != 'nan':
            movil_idx[n.upper()] = row

    # Construir ejecutivos desde Datos
    for i, row in df_d.iterrows():
        if i < 2: continue
        nombre   = str(row[1]).strip()  # APODO
        cargo_d  = str(row[3]).strip()  # Detalle Cargo
        estado   = str(row[19]).strip()
        if nombre in ('nan','','APODO') or cargo_d not in ('Junior','Senior'):
            continue
        # Incluir todos los cargos (ACTIVO, LICENCIA, NUEVO, etc.)

        rut     = str(row[4]).strip()
        region  = str(row[6]).strip()
        ventas_cerradas = int(nz(row[8]))
        ventas_movil_q  = int(nz(row[9]))
        meta    = int(nz(row[17])) if nz(row[17]) else 21
        ciudad, region_nombre = REGION_MAP.get(region, ('Las Condes', 'Región Metropolitana'))

        # Porcentaje cumplimiento
        pct = ventas_cerradas / meta if meta > 0 else 0
        pct_tramo = tramo_pct(pct)

        # Cantidades BAF de hoja Fijas
        frow = fijas_idx.get(nombre.upper())
        baf = {
            'BAF 200': 0, 'BAF 300': 0, 'BAF 400': 0, 'BAF 500': 0,
            'BAF 600': 0, 'BAF 700': 0, 'BAF 800': 0, 'BAF 940': 0,
            'BAF 2000': 0, 'IPTV': 0, 'LINEAS': 0,
        }
        if frow is not None:
            baf['BAF 300'] = int(nz(frow[2]))
            baf['BAF 400'] = int(nz(frow[3]))
            baf['BAF 500'] = int(nz(frow[4]))
            baf['BAF 600'] = int(nz(frow[5]))
            baf['BAF 700'] = int(nz(frow[6]))
            baf['BAF 800'] = int(nz(frow[7]))
            baf['BAF 940'] = int(nz(frow[8]))
            baf['IPTV']    = int(nz(frow[9]))
            baf['LINEAS']  = int(nz(frow[10]))

        # Comisión ventas cerradas = sum(cant * CFN * pct_tramo)
        com_cerradas = sum(
            baf[p] * CARGO_FIJO_NETO[p] * pct_tramo
            for p in PRODUCTOS_ORDER
        )
        # Valores monetarios por producto para tabla CERRADAS
        baf_vals = {
            p: baf[p] * CARGO_FIJO_NETO[p] * pct_tramo
            for p in PRODUCTOS_ORDER
        }

        # Ventas emitidas
        ventas_emitidas = math.floor(ventas_cerradas * 1.15 + 0.5)
        com_emitidas = ventas_emitidas * tramo_emitidas(ventas_emitidas)

        # Bono BAF no carterizado
        total_fo_tv = sum(baf[p] for p in PRODUCTOS_ORDER if p != 'LINEAS')
        bono = bono_baf(total_fo_tv)

        # Comisión movil desde Pago o hoja Moviles
        bono_movil = 0
        mrow = movil_idx.get(nombre.upper())
        if mrow is not None:
            bono_movil = int(nz(mrow[7]))  # col 7 = BONO MOVIL

        self.ejecutivos[nombre] = {
            'nombre':      nombre,
            'rut':         rut,
            'cargo':       cargo_d,
            'region':      region,
            'ciudad':      ciudad,
            'region_nom':  region_nombre,
            'meta':        meta,
            'cerradas':    ventas_cerradas,
            'emitidas':    ventas_emitidas,
            'pct':         pct,
            'pct_tramo':   pct_tramo,
            'baf':         baf,
            'baf_vals':    baf_vals,
            'com_cerradas':com_cerradas,
            'com_emitidas':com_emitidas,
            'total_fo_tv': total_fo_tv,
            'bono':        bono,
            'bono_movil':  bono_movil,
            'movil_q':     ventas_movil_q,
        }

    # ---- Cierre Fijo Maestro ----
    df_mf = pd.read_excel(self.archivos["cfijo"], sheet_name='Maestro', header=None)
    for i, row in df_mf.iterrows():
        if i == 0: continue
        ej   = str(row[4]).strip()    # col E = EJECUTIVO
        per  = str(row[0]).strip()    # col A = PERIODO
        ord_ = str(row[1]).strip()    # col B = ORDEN
        prod = str(row[2]).strip()    # col C = PRODUCTO
        rut  = str(row[25]).strip()   # col Z = RUT_CLIENTE
        com  = str(row[9]).strip()    # col J = COMISIONABLE
        # Solo registros COMISIONABLE (excluye PYME, NO COMISIONABLE, FREELANCE)
        if (ej and ej != 'nan'
                and not ej.startswith('FREELANCE')
                and com == 'COMISIONABLE'):
            # Normalizar TV → IPTV para que coincida con tabla de productos
            prod_norm = 'IPTV' if prod.upper() == 'TV' else prod
            self.ventas_fijo.append({
                'ej': ej, 'per': per, 'ord': ord_,
                'prod': prod_norm, 'rut': rut
            })

    # ---- Cierre Movil Maestro ----
    df_mm = pd.read_excel(self.archivos["cmov"], sheet_name='MAESTRO', header=None)
    for i, row in df_mm.iterrows():
        if i == 0: continue
        ej   = str(row[159]).strip()  # col EJECUTIVO
        per  = str(row[0]).strip()    # col A = PERIODO
        ord_ = str(row[1]).strip()    # col B = ORDEN
        cel  = str(row[11]).strip()   # col L = CELULAR
        rut  = str(row[15]).strip()   # col P = RUT_CLIENTE
        if ej and ej != 'nan':
            self.ventas_movil.append({'ej':ej,'per':per,'ord':ord_,'cel':cel,'rut':rut})

    self.log(f"Datos cargados: {len(self.ejecutivos)} ejecutivos, "
             f"{len(self.ventas_fijo)} ventas fijo, {len(self.ventas_movil)} ventas movil.")

# ----------------------------------------------------------
#  CONSTRUCCIÓN DEL DOCUMENTO
# ----------------------------------------------------------
def crear_doc(self, d, ventas_fijo, ventas_movil, path):
    doc = Document(self.archivos["pla"])

    # Encabezado dinámico: reemplazar placeholders
    for section in doc.sections:
        for tbl in section.header.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if '[CIUDAD]' in run.text:
                                run.text = run.text.replace('[CIUDAD]', d['ciudad'])
                            if '[REGION]' in run.text:
                                run.text = run.text.replace('[REGION]', d['region_nom'])

    # También intentar en párrafos del header
    for section in doc.sections:
        for para in section.header.paragraphs:
            for run in para.runs:
                run.text = run.text.replace('[CIUDAD]', d['ciudad'])
                run.text = run.text.replace('[REGION]', d['region_nom'])

    # Buscar marcador
    ptr = None
    for p in doc.paragraphs:
        if "[INSERTAR_TABLAS]" in p.text:
            ptr = p; p.text = ""; break
    if not ptr:
        # Si no hay marcador, insertar al final
        ptr = doc.add_paragraph()
    ref = ptr._p

    # ---- helpers internos ----
    def add_p(txt, bold=False, size=10, align=WD_ALIGN_PARAGRAPH.LEFT,
              color="000000", space=2):
        nonlocal ref
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_after = Pt(space)
        p.paragraph_format.space_before = Pt(0)
        run = p.add_run(txt)
        run.bold = bold
        run.font.size = Pt(size)
        run.font.name = 'Arial'
        run.font.color.rgb = RGBColor.from_string(color)
        ref.addnext(p._p); ref = p._p

    def add_t(headers, rows, widths, header_color=COLOR_HEADER, det_mode=False):
        nonlocal ref
        nc = len(headers)
        t = doc.add_table(rows=0, cols=nc)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        set_table_borders(t)

        tblPr = t._element.xpath('w:tblPr')[0]
        tblW = OxmlElement('w:tblW')
        tblW.set(qn('w:type'), 'pct')
        tblW.set(qn('w:w'), '5000')
        tblPr.append(tblW)

        # Header row que se repite
        h_row = t.add_row()
        trPr = h_row._tr.get_or_add_trPr()
        hdrEl = OxmlElement('w:tblHeader')
        hdrEl.set(qn('w:val'), 'true')
        trPr.append(hdrEl)
        for i, h in enumerate(headers):
            c = h_row.cells[i]
            c.width = Cm(widths[i])
            set_cell_margins(c)
            set_cell_bg(c, header_color)
            pa = c.paragraphs[0]
            pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
            rn = pa.add_run(h)
            rn.bold = True
            rn.font.color.rgb = RGBColor(255,255,255)
            rn.font.size = Pt(8.5)
            rn.font.name = 'Arial'

        for idx, r_data in enumerate(rows):
            dr = t.add_row()
            trPr2 = dr._tr.get_or_add_trPr()
            cs = OxmlElement('w:cantSplit')
            cs.set(qn('w:val'), 'true')
            trPr2.append(cs)
            for i, v in enumerate(r_data):
                c = dr.cells[i]
                c.width = Cm(widths[i])
                set_cell_margins(c)
                pa = c.paragraphs[0]
                pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
                val = "" if str(v).lower() == "nan" else str(v)
                rn = pa.add_run(val)
                rn.font.size = Pt(8.5)
                rn.font.name = 'Arial'
                if det_mode and idx % 2 == 1:
                    set_cell_bg(c, COLOR_ROW_ALT)
                elif not det_mode and idx == 0:
                    set_cell_bg(c, "E8EEF6")
                    rn.bold = True

        ref.addnext(t._tbl); ref = t._tbl

    def add_spacer(pts=4):
        nonlocal ref
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        pPr = p._p.get_or_add_pPr()
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:before'), '0')
        sp.set(qn('w:after'), str(int(pts * 20)))
        pPr.append(sp)
        ref.addnext(p._p); ref = p._p

    def add_page_break():
        nonlocal ref
        p = doc.add_paragraph()
        p.add_run().add_break(WD_BREAK.PAGE)
        ref.addnext(p._p); ref = p._p

    baf   = d['baf']
    bafv  = d['baf_vals']

    # ============================================================
    #  PÁGINA 1 — RESUMEN COMISIÓN
    # ============================================================
    add_p(f"NOMBRE: {d['nombre']}", True, 12, space=0)
    add_p(f"RUT: {d['rut']}", False, 10, space=0)
    add_p(f"META: {d['meta']} RGU", False, 10, space=8)

    add_p("RESUMEN COMISIÓN", True, 11, WD_ALIGN_PARAGRAPH.CENTER,
          COLOR_HEADER, space=4)

    # ---- Tabla RESUMEN COMISIÓN — replica exacta del simulador ----
    # 4 columnas: etiq_izq | val_izq | etiq_der | val_der
    # Col 2 (etiquetas der) siempre azul oscuro + texto blanco
    # Fila 1 col 0-1 también azul oscuro
    def add_resumen_table():
        nonlocal ref
        t = doc.add_table(rows=0, cols=4)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        set_table_borders(t)
        tblPr = t._element.xpath('w:tblPr')[0]
        tblW = OxmlElement('w:tblW')
        tblW.set(qn('w:w'), '5000'); tblW.set(qn('w:type'), 'pct')
        tblPr.append(tblW)

        widths = [3.5, 2.0, 5.5, 2.5]
        rows_data = [
            ["EJECUTIVO",             d['nombre'],
             "COMISION POR VENTAS CERRADAS",        fmt_m(d['com_cerradas'])],
            ["",                      "",
             "COMISION POR VENTAS EMITIDAS",         fmt_m(d['com_emitidas'])],
            ["META FIJO",             str(d['meta']),
             "NUMERO TOTAL DE FIBRA OPTICA + IPTV",  str(int(d['total_fo_tv']))],
            ["VENTAS FIJAS CERRADAS", str(d['cerradas']),
             "BONO BAF NO CARTERIZADO",              fmt_m(d['bono'])],
            ["VENTAS FIJAS EMITIDAS", str(d['emitidas']),
             "NUMERO TOTAL DE MOVILES",              str(d['movil_q'])],
            ["% CUMPLIMIENTO",        fmt_pct(d['pct']),
             "BONO MOVIL",                           fmt_m(d['bono_movil'])],
        ]
        for r_idx, r_data in enumerate(rows_data):
            dr = t.add_row()
            trPr2 = dr._tr.get_or_add_trPr()
            cs = OxmlElement('w:cantSplit'); cs.set(qn('w:val'), 'true')
            trPr2.append(cs)
            for c_idx, val in enumerate(r_data):
                cell = dr.cells[c_idx]
                cell.width = Cm(widths[c_idx])
                set_cell_margins(cell)
                pa = cell.paragraphs[0]
                pa.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in (1, 3) else WD_ALIGN_PARAGRAPH.LEFT
                rn = pa.add_run(str(val))
                rn.font.size = Pt(8.5); rn.font.name = 'Arial'
                # Fila 1 cols izq → azul oscuro
                if r_idx == 0 and c_idx in (0, 1):
                    set_cell_bg(cell, COLOR_HEADER)
                    rn.bold = True; rn.font.color.rgb = RGBColor(255, 255, 255)
                # Etiquetas der (col 2) siempre azul oscuro
                elif c_idx == 2:
                    set_cell_bg(cell, COLOR_HEADER)
                    rn.bold = True; rn.font.color.rgb = RGBColor(255, 255, 255)
                # Etiquetas izq filas 2-6 → negrita si tienen texto
                elif c_idx == 0:
                    rn.bold = (val != "")
        ref.addnext(t._tbl); ref = t._tbl

    add_resumen_table()

    add_spacer(14)   # <-- espacio visible entre resumen y CANTIDADES

    # ---- Tabla CANTIDADES con etiquetas de fila ----
    # Estructura: col 0 = etiqueta (CANTIDADES / CARGO FIJO / CERRADAS)
    #             cols 1-11 = productos BAF 200 … LINEAS
    add_p("CANTIDADES", True, 9, WD_ALIGN_PARAGRAPH.LEFT, space=2)

    def add_tabla_con_etiqueta(etiq_filas, data_filas, header_color):
        """Tabla con columna de etiqueta izquierda + 11 columnas de productos."""
        nonlocal ref
        n_cols = 12  # 1 etiq + 11 productos
        t = doc.add_table(rows=0, cols=n_cols)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        set_table_borders(t)
        tblPr = t._element.xpath('w:tblPr')[0]
        tblW = OxmlElement('w:tblW')
        tblW.set(qn('w:w'), '5000'); tblW.set(qn('w:type'), 'pct')
        tblPr.append(tblW)

        # Ancho: col etiqueta 1.8cm, resto distribuido
        etiq_w = 1.8
        prod_w = 1.27
        widths = [etiq_w] + [prod_w] * 11

        # Fila encabezado (nombres productos)
        h_row = t.add_row()
        trPr = h_row._tr.get_or_add_trPr()
        hdrEl = OxmlElement('w:tblHeader'); hdrEl.set(qn('w:val'), 'true')
        trPr.append(hdrEl)
        # col 0 encabezado vacío
        c0 = h_row.cells[0]; c0.width = Cm(etiq_w)
        set_cell_margins(c0); set_cell_bg(c0, header_color)
        c0.paragraphs[0].add_run("")
        for i, p in enumerate(PRODUCTOS_ORDER):
            c = h_row.cells[i + 1]; c.width = Cm(prod_w)
            set_cell_margins(c); set_cell_bg(c, header_color)
            pa = c.paragraphs[0]; pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
            rn = pa.add_run(p)
            rn.bold = True; rn.font.color.rgb = RGBColor(255, 255, 255)
            rn.font.size = Pt(8); rn.font.name = 'Arial'

        # Filas de datos
        for row_idx, (etiq, datos) in enumerate(zip(etiq_filas, data_filas)):
            dr = t.add_row()
            trPr2 = dr._tr.get_or_add_trPr()
            cs = OxmlElement('w:cantSplit'); cs.set(qn('w:val'), 'true')
            trPr2.append(cs)
            # Celda etiqueta
            ce = dr.cells[0]; ce.width = Cm(etiq_w)
            set_cell_margins(ce); set_cell_bg(ce, "E8EEF6")
            pa_e = ce.paragraphs[0]; pa_e.alignment = WD_ALIGN_PARAGRAPH.LEFT
            rn_e = pa_e.add_run(etiq)
            rn_e.bold = True; rn_e.font.size = Pt(8); rn_e.font.name = 'Arial'
            # Celdas datos
            for i, val in enumerate(datos):
                c = dr.cells[i + 1]; c.width = Cm(prod_w)
                set_cell_margins(c)
                if row_idx % 2 == 1:
                    set_cell_bg(c, COLOR_ROW_ALT)
                pa = c.paragraphs[0]; pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
                rn = pa.add_run(str(val) if str(val).lower() != 'nan' else "")
                rn.font.size = Pt(8); rn.font.name = 'Arial'
        ref.addnext(t._tbl); ref = t._tbl

    cant_row = [baf[p] if baf[p] != 0 else 0 for p in PRODUCTOS_ORDER]
    add_tabla_con_etiqueta(
        ["CARGO FIJO", "CERRADAS"],
        [
            [fmt_m(CARGO_FIJO_BRUTO[p]) for p in PRODUCTOS_ORDER],
            [str(v) for v in cant_row],
        ],
        COLOR_HEADER2,
    )

    add_spacer(8)

    # ---- Subtítulo 60% ----
    add_p("60% DEL CARGO FIJO NETO", True, 10, WD_ALIGN_PARAGRAPH.CENTER,
          COLOR_HEADER, space=4)

    val_row = [fmt_m(bafv[p]) if bafv[p] else "$ -" for p in PRODUCTOS_ORDER]
    add_tabla_con_etiqueta(
        ["CARGO FIJO", "CERRADAS"],
        [
            [fmt_m(CARGO_FIJO_BRUTO[p]) for p in PRODUCTOS_ORDER],
            val_row,
        ],
        COLOR_HEADER2,
    )

    add_page_break()

    # ============================================================
    #  PÁGINA 2 — DETALLE VENTAS FIJAS
    # ============================================================
    add_p("DETALLE DE VENTAS:", True, 11, WD_ALIGN_PARAGRAPH.CENTER,
          COLOR_HEADER, space=6)
    add_p("VENTAS FIJAS TERMINADAS:", True, 10, WD_ALIGN_PARAGRAPH.CENTER,
          COLOR_HEADER2, space=4)

    # Solo productos comisionables (excluye TV, etc.)
    PRODS_COMISIONABLES_UP = {p.upper() for p in CARGO_FIJO_NETO.keys()}
    ventas_fijo_com = [
        v for v in ventas_fijo
        if v['prod'].strip().upper() in PRODS_COMISIONABLES_UP
    ]

    if not ventas_fijo_com:
        add_p("No posee ventas", False, 9, WD_ALIGN_PARAGRAPH.CENTER, "777777")
    else:
        rows_fijo = [[v['per'], v['ord'], v['prod'], v['rut']] for v in ventas_fijo_com]
        add_t(['PERIODO','ORDEN','PRODUCTO','RUT_CLIENTE'],
              rows_fijo, [2.0, 3.5, 3.0, 3.5],
              header_color=COLOR_HEADER2, det_mode=True)

    add_spacer(10)
    add_p("VENTAS MOVILES:", True, 10, WD_ALIGN_PARAGRAPH.CENTER,
          COLOR_HEADER2, space=4)

    if not ventas_movil:
        add_p("No posee ventas", False, 9, WD_ALIGN_PARAGRAPH.CENTER, "777777")
    else:
        rows_movil = [[v['per'], v['ord'], v['cel'], v['rut']] for v in ventas_movil]
        add_t(['PERIODO','ORDEN','CELULAR','RUT_CLIENTE'],
              rows_movil, [2.0, 3.5, 3.0, 3.5],
              header_color=COLOR_HEADER2, det_mode=True)

    # ============================================================
    #  PÁGINA 3 — EXPLICACIÓN (siempre en página nueva)
    # ============================================================
    add_page_break()

    doc.save(path)
```

if **name** == “**main**”:
root = tk.Tk()
app = MasivoApp(root)
root.mainloop()