# ============================================================
#  RECORDATORIO DE INCONVENIENTES — ENVÍO POR WHATSAPP
#  Requiere: pip install openpyxl
#  Uso: python enviar_whatsapp.py
# ============================================================

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl
import webbrowser
import urllib.parse
import re
import os
from auto_envio import envio
import time
import sys
import os

# Obtiene la ruta absoluta para recursos (imágenes, etc) 

def ruta_recurso(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
    
# ── Colores y fuentes ────────────────────────────────────────
BG        = "#F7F8FA"
SIDEBAR   = "#1E2432"
VERDE_WA  = "#25D366"
VERDE_OSC = "#128C7E"
ROJO      = "#E74C3C"
NARANJA   = "#E67E22"
AZUL      = "#2980B9"
BLANCO    = "#FFFFFF"
TEXTO     = "#1A1A2E"
MUTED     = "#6B7280"
BORDE     = "#E2E8F0"

TIPOS_INCONVENIENTE = [
    "— Seleccione tipo —",
    "PERSONALIZADO",
    "CANCELA CITA",
    "NO PUEDE ASISTIR — CITA MÉDICA",
    "RECORDATORIO DE CITA MEDICA",
    "INASISTENCIA SIN AVISO",
    "TERMINA SESIONES",
    "RETRASO EN LLEGADA",
    "CALAMIDAD DOMÉSTICA",
    "PROBLEMA DE ORDEN PÚBLICO",
    "ZONA RURAL — SIN TRANSPORTE"
]

MENSAJES_PLANTILLA = {
    "PERSONALIZADO":
        "Hola {nombre}, {mensaje}",
    "CANCELA CITA":
        "Hola {nombre}, le informamos que su cita ha sido cancelada. "
        "Por favor comuníquese con nosotros para reprogramar. Gracias.",
    "NO PUEDE ASISTIR — CITA MÉDICA":
        "Hola {nombre}, entendemos que tiene una cita médica y no podrá asistir hoy. "
        "No se preocupe, comuníquese para reagendar su sesión. Que se mejore.",
    "RECORDATORIO DE CITA MEDICA":
        "Hola {nombre}, esperamos que se encuentre bien. "
        "Para recordale que el dia cita. Cuídese mucho.",
    "INASISTENCIA SIN AVISO":
        "Hola {nombre}, notamos que no asistió a su sesión programada de hoy. "
        "Por favor confirme si desea continuar su tratamiento. Gracias.",
    "TERMINA SESIONES":
        "Hola {nombre}, ha completado exitosamente sus sesiones de terapia física. "
        "Ha sido un gusto acompañarle en su recuperación. Que siga muy bien.",
    "RETRASO EN LLEGADA":
        "Hola {nombre}, le recordamos que tiene sesión de terapia hoy. "
        "Si se va a demorar, por favor avísenos para coordinar su atención. Gracias.",
    "CALAMIDAD DOMÉSTICA":
        "Hola {nombre}, lamentamos su situación. "
        "Cuando todo esté mejor, con gusto reagendamos sus sesiones. Estamos aquí para apoyarle.",
    "PROBLEMA DE ORDEN PÚBLICO":
        "Hola {nombre}, entendemos la situación. Su seguridad es lo más importante. "
        "Cuando pueda movilizarse, contáctenos para retomar sus sesiones.",
    "ZONA RURAL — SIN TRANSPORTE":
        "Hola {nombre}, entendemos las dificultades de transporte. "
        "Por favor avísenos con anticipación cuando pueda venir. Gracias.",
    "OTRO INCONVENIENTE":
        "Hola {nombre}, le contactamos del centro de fisioterapia. "
        "Hemos notado un inconveniente con su cita. Por favor comuníquese con nosotros. Gracias.",
}


def limpiar_telefono(tel):
    """Deja solo dígitos y agrega 57 si es colombiano de 10 dígitos."""
    if tel is None:
        return ""
    # El teléfono puede venir como int directamente desde Excel
    tel = str(int(tel)) if isinstance(tel, (int, float)) else str(tel)
    tel = re.sub(r"[^0-9]", "", tel.split("-")[0].split("/")[0])
    if len(tel) == 10:
        tel = "57" + tel
    return tel


def detectar_tipo(obs):
    obs = obs.upper()
    if "CANCELA" in obs:                              return "CANCELA CITA"
    if "CITA MEDICA" in obs or "CITA MÉDICA" in obs: return "NO PUEDE ASISTIR — CITA MÉDICA"
    if any(p in obs for p in ["SALUD","ENFERM","MALUC"]): return "PROBLEMA DE SALUD"
    if "INASIST" in obs:                              return "INASISTENCIA SIN AVISO"
    if "TERMINA" in obs or "TERMINADO" in obs:        return "TERMINA SESIONES"
    if "CALAMIDAD" in obs:                            return "CALAMIDAD DOMÉSTICA"
    if "ORDEN PUBLICO" in obs:                        return "PROBLEMA DE ORDEN PÚBLICO"
    if "RURAL" in obs:                                return "ZONA RURAL — SIN TRANSPORTE"
    if obs.strip():                                   return "OTRO INCONVENIENTE"
    return ""


def cargar_pacientes(ruta):
    wb = openpyxl.load_workbook(ruta)
    ws = wb.active
    pacientes = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        # Columnas reales del archivo:
        # 0=PROFESIONAL, 1=SESIONES_TERAPIA, 2=NOMBRE, 3=CEDULA,
        # 4=ENTIDAD, 5=LUNES, 6=MARTES, 7=MIERCOLES, 8=JUEVES,
        # 9=VIERNES, 10=OBSERVASION, 11=LISTA CITA, 12=TELEFONO
        if len(row) < 13:
            continue
        profesional = str(row[0]).strip() if row[0] else ""
        nombre      = str(row[2]).strip() if row[2] else ""
        entidad     = str(row[4]).strip() if row[4] else ""
        obs         = str(row[10]).strip() if row[10] else ""
        lista_cita  = str(row[11]).strip() if row[11] else ""
        telefono    = str(int(row[12])) if isinstance(row[12], (int, float)) else (str(row[12]).strip() if row[12] else "")
        tel_limpio  = limpiar_telefono(telefono)

        if not nombre or str(nombre).strip() in ("None", "NOMBRE", ""):
            continue

        tipo = detectar_tipo(obs)
        nombre_corto = nombre.split()[0].capitalize()
        msg = MENSAJES_PLANTILLA.get(tipo, "").replace("{nombre}", nombre_corto) if tipo else ""

        pacientes.append({
            "profesional": profesional,
            "nombre":      nombre,
            "telefono":    telefono,
            "tel_limpio":  tel_limpio,
            "obs":         obs,
            "lista_cita":  lista_cita,
            "tipo":        tipo,
            "mensaje":     msg,
        })
    return pacientes


def construir_url(tel, mensaje):
    msg_enc = urllib.parse.quote(mensaje)
    return f"https://web.whatsapp.com/send?phone={tel}&text={msg_enc}"


# ── Aplicación principal ─────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Mensajes Automáticos — WhatsApp")
        self.geometry("1100x700")
        self.minsize(900, 580)
        self.configure(bg=BG)
        self.pacientes = []
        self._build_ui()

    # ── Interfaz ─────────────────────────────────────────────
    def _build_ui(self):
        # Sidebar izquierdo
        sidebar = tk.Frame(self, bg=SIDEBAR, width=240)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        tk.Label(sidebar, text="💬", font=("Segoe UI Emoji", 32),
                 bg=SIDEBAR, fg=VERDE_WA).pack(pady=(32, 4))
        tk.Label(sidebar, text="WhatsApp\nAuto Mensajes",
                 font=("Segoe UI", 14, "bold"), bg=SIDEBAR,
                 fg=BLANCO, justify="center").pack()

        ttk.Separator(sidebar, orient="horizontal").pack(fill="x", padx=20)

        # Botón cargar Excel
        self._sidebar_btn(sidebar, "📂  Cargar Excel", self.cargar_excel, VERDE_OSC)

        # Filtro profesional
        tk.Label(sidebar, text="FILTRAR POR PROFESIONAL",
                 font=("Segoe UI", 8, "bold"), bg=SIDEBAR,
                 fg="#8892A4").pack(pady=(24, 4), padx=16, anchor="w")

        self.var_filtro = tk.StringVar(value="Todos")
        self.combo_filtro = ttk.Combobox(sidebar, textvariable=self.var_filtro,
                                          state="readonly", width=22)
        self.combo_filtro.pack(padx=16, pady=(0, 8))
        self.combo_filtro.bind("<<ComboboxSelected>>", lambda e: self.actualizar_tabla())

        # Filtro tipo inconveniente
        tk.Label(sidebar, text="TIPO DE INCONVENIENTE",
                 font=("Segoe UI", 8, "bold"), bg=SIDEBAR,
                 fg="#8892A4").pack(pady=(8, 4), padx=16, anchor="w")

        self.var_tipo_filtro = tk.StringVar(value="PERSONALIZADO")
        tipos_filtro = ["PERSONALIZADO"] + TIPOS_INCONVENIENTE[1:]
        self.combo_tipo = ttk.Combobox(sidebar, textvariable=self.var_tipo_filtro,
                                        values=tipos_filtro, state="readonly", width=22)
        self.combo_tipo.pack(padx=16)
        self.combo_tipo.bind("<<ComboboxSelected>>", lambda e: self.actualizar_tabla())

        # Personalizado
        self.personalizado = tk.Frame(sidebar, bg=SIDEBAR)
        self.personalizado.pack(fill="x")

        self.label_personalizado = tk.Label(self.personalizado, text="MENSAJE PERSONALIZADO",
                 font=("Segoe UI", 8, "bold"), bg=SIDEBAR,
                 fg="#8892A4")
        self.label_personalizado.pack(pady=(8, 4), padx=16, anchor="w")
        self.text_personalizado = tk.Text(self.personalizado, height=5)
        self.text_personalizado.pack(pady=(8, 4), padx=16)
       


        # Contador
        self.lbl_count = tk.Label(sidebar, text="0 pacientes cargados",
                                   font=("Segoe UI", 9), bg=SIDEBAR, fg="#8892A4")
        self.lbl_count.pack(pady=(20, 0))

        # Enviar todos
        self._sidebar_btn(sidebar, "📤  Enviar a todos los\npacientes filtrados",
                          self.enviar_todos, NARANJA)

        # Panel derecho
        main = tk.Frame(self, bg=BG)
        main.pack(side="right", fill="both", expand=True)

        # Barra superior
        topbar = tk.Frame(main, bg=BLANCO, height=64,
                          highlightbackground=BORDE, highlightthickness=1)
        topbar.pack(fill="x")
        topbar.pack_propagate(False)

        tk.Label(topbar, text="Lista De Pacientes",
                 font=("Segoe UI", 14, "bold"), bg=BLANCO, fg=TEXTO).pack(
                 side="left", padx=20, pady=16)

        # Buscador
        self.var_buscar = tk.StringVar()
        self.var_buscar.trace_add("write", lambda *_: self.actualizar_tabla())
        entry_frame = tk.Frame(topbar, bg=BLANCO)
        entry_frame.pack(side="right", padx=20, pady=12)
        tk.Label(entry_frame, text="Buscar pacientes 🔍", font=("Segoe UI Emoji", 12),
                 bg=BLANCO).pack(side="left")
        tk.Entry(entry_frame, textvariable=self.var_buscar,
                 font=("Segoe UI", 11), width=24,
                 bd=0, bg="#F1F3F5", fg=TEXTO,
                 insertbackground=TEXTO).pack(side="left", padx=4)

        # Tabla
        tabla_frame = tk.Frame(main, bg=BG)
        tabla_frame.pack(fill="both", expand=True, padx=16, pady=12)

        cols = ("profesional", "nombre", "telefono")
        self.tabla = ttk.Treeview(tabla_frame, columns=cols,
                                   show="headings", selectmode="browse")

        hdrs = {"profesional": ("Profesional", 110),
                "nombre":      ("Nombre paciente", 200),
                "telefono":    ("Teléfono", 130)}

        for col, (label, w) in hdrs.items():
            self.tabla.heading(col, text=label,
                               command=lambda c=col: self._ordenar(c))
            self.tabla.column(col, width=w, minwidth=80)

        vsb = ttk.Scrollbar(tabla_frame, orient="vertical",
                             command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(tabla_frame, orient="horizontal",
                             command=self.tabla.xview)
        self.tabla.configure(xscrollcommand=hsb.set)
        hsb.pack(side="bottom", fill="x")
        self.tabla.pack(fill="both", expand=True)
        self.tabla.tag_configure("impar", background="#F9FAFB")
        self.tabla.tag_configure("par",   background=BLANCO)
        self.tabla.bind("<Double-1>", lambda e: self.ver_detalle())

        # Panel inferior — detalle
        detalle = tk.Frame(main, bg=BLANCO,
                           highlightbackground=BORDE, highlightthickness=1)
        detalle.pack(fill="x", padx=16, pady=(0, 12))

        tk.Label(detalle, text="Paciente seleccionado",
                 font=("Segoe UI", 10, "bold"), bg=BLANCO, fg=TEXTO).grid(
                 row=0, column=0, columnspan=4, sticky="w", padx=12, pady=(10, 4))

        # Campos de detalle editables
        labels_d = ["Nombre:", "Teléfono (con código):", "Tipo:", "Mensaje:"]
        self.campos = {}
        for i, lbl in enumerate(labels_d):
            tk.Label(detalle, text=lbl, font=("Segoe UI", 9), bg=BLANCO,
                     fg=MUTED, width=20, anchor="e").grid(row=i+1, column=0, padx=(12, 4), pady=3)
            if lbl == "Mensaje:":
                w = tk.Text(detalle, font=("Segoe UI", 10), height=3,
                             width=60, bd=1, relief="solid",
                             bg="#F9FAFB", fg=TEXTO, wrap="word")
                w.grid(row=i+1, column=1, columnspan=2, sticky="ew", padx=(0, 8), pady=3)
                self.campos["mensaje"] = w
            elif lbl == "Tipo:":
                self.var_tipo_det = tk.StringVar()
                cb = ttk.Combobox(detalle, textvariable=self.var_tipo_det,
                                   values=TIPOS_INCONVENIENTE, state="readonly",
                                   font=("Segoe UI", 10), width=40)
                cb.grid(row=i+1, column=1, sticky="w", padx=(0, 8), pady=3)
                cb.bind("<<ComboboxSelected>>", self._actualizar_mensaje_plantilla)
                self.campos["tipo"] = cb
            else:
                key = "nombre" if "Nombre" in lbl else "telefono"
                var = tk.StringVar()
                tk.Entry(detalle, textvariable=var, font=("Segoe UI", 10),
                         width=38, bd=1, relief="solid",
                         bg="#F9FAFB", fg=TEXTO).grid(
                         row=i+1, column=1, sticky="w", padx=(0, 8), pady=3)
                self.campos[key] = var

        detalle.columnconfigure(1, weight=1)

        # Botones acción
        btn_frame = tk.Frame(detalle, bg=BLANCO)
        btn_frame.grid(row=5, column=0, columnspan=4, sticky="e",
                       padx=12, pady=(4, 12))

        self._btn(btn_frame, "📲  Enviar WhatsApp", self.enviar_seleccionado,
                  VERDE_WA).pack(side="left", padx=4)

        self.tabla.bind("<<TreeviewSelect>>", self._on_select)

        # Estilos ttk
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview", font=("Segoe UI", 10),
                         rowheight=28, background=BLANCO,
                         fieldbackground=BLANCO, foreground=TEXTO)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"),
                         background="#F1F5F9", foreground=TEXTO,
                         relief="flat", padding=6)
        style.map("Treeview", background=[("selected", "#EBF5FB")],
                  foreground=[("selected", TEXTO)])

    # ── Widgets helpers ──────────────────────────────────────
    def _sidebar_btn(self, parent, text, cmd, color):
        f = tk.Frame(parent, bg=color, cursor="hand2")
        f.pack(fill="x", padx=16, pady=6)
        lbl = tk.Label(f, text=text, font=("Segoe UI", 10, "bold"),
                        bg=color, fg=BLANCO, pady=10, justify="center")
        lbl.pack(fill="x")
        for w in (f, lbl):
            w.bind("<Button-1>", lambda e, c=cmd: c())

    def _btn(self, parent, text, cmd, color, small=False):
        sz = 9 if small else 10
        b = tk.Button(parent, text=text, font=("Segoe UI", sz, "bold"),
                       bg=color, fg=BLANCO, bd=0, padx=14, pady=6,
                       cursor="hand2", relief="flat",
                       activebackground=color, activeforeground=BLANCO,
                       command=cmd)
        return b

    # ── Lógica ───────────────────────────────────────────────
    def cargar_excel(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Todos", "*.*")])
        if not ruta:
            return
        try:
            self.pacientes = cargar_pacientes(ruta)
            # Actualizar filtro profesional
            profs = sorted(set(p["profesional"] for p in self.pacientes if p["profesional"]))
            self.combo_filtro["values"] = ["Todos"] + profs
            self.var_filtro.set("Todos")
            self.actualizar_tabla()
            messagebox.showinfo("Cargado",
                f"✅ {len(self.pacientes)} pacientes cargados correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo:\n{e}")

    def _pacientes_filtrados(self):
        filtro_prof = self.var_filtro.get()
        #filtro_tipo = self.var_tipo_filtro.get()
        buscar      = self.var_buscar.get().lower()
        result = []
        for p in self.pacientes:
            if filtro_prof != "Todos" and p["profesional"] != filtro_prof:
                continue
            #if filtro_tipo != "Todos" and p["tipo"] != filtro_tipo:
               # continue
            if buscar and buscar not in p["nombre"].lower() \
                      and buscar not in p["telefono"].lower():
                continue
            result.append(p)
        return result

    def actualizar_tabla(self):
        for item in self.tabla.get_children():
            self.tabla.delete(item)
        filtrados = self._pacientes_filtrados()
        for i, p in enumerate(filtrados):
            tag = "impar" if i % 2 else "par"
            self.tabla.insert("", "end", iid=str(i), tags=(tag,),
                values=(p["profesional"], p["nombre"],
                        p["telefono"], p["tipo"],
                        p["mensaje"][:80] + ("…" if len(p["mensaje"]) > 80 else "")))
        total = len(self.pacientes)
        shown = len(filtrados)
        self.lbl_count.config(
            text=f"{shown} de {total} pacientes")
        if self.var_tipo_filtro.get() == "PERSONALIZADO":
            self.label_personalizado.pack(pady=(8, 4), padx=16, anchor="w")
            self.text_personalizado.pack(pady=(8, 4))
        else:
            self.label_personalizado.pack_forget()
            self.text_personalizado.pack_forget()

    def _on_select(self, event=None):
        sel = self.tabla.selection()
        if not sel:
            return
        idx   = int(sel[0])
        filtrados = self._pacientes_filtrados()
        if idx >= len(filtrados):
            return
        p = filtrados[idx]
        self.campos["nombre"].set(p["nombre"])
        self.campos["telefono"].set(p["tel_limpio"])
        self.var_tipo_det.set(p["tipo"] or TIPOS_INCONVENIENTE[0])
        self.campos["mensaje"].delete("1.0", "end")
        self.campos["mensaje"].insert("1.0", p["mensaje"])

    def _actualizar_mensaje_plantilla(self, event=None):
        tipo  = self.var_tipo_det.get()
        nombre = self.campos["nombre"].get().split()[0].capitalize() \
                 if self.campos["nombre"].get() else "Paciente"
        msg = MENSAJES_PLANTILLA.get(tipo, "").replace("{nombre}", nombre)
        self.campos["mensaje"].delete("1.0", "end")
        self.campos["mensaje"].insert("1.0", msg)

    def ver_detalle(self):
        self._on_select()

    def enviar_seleccionado(self):
        tel = self.campos["telefono"].get().strip()
        msg = self.campos["mensaje"].get("1.0", "end").strip()
        nombre = self.campos["nombre"].get()

        if not tel:
            messagebox.showwarning("Sin teléfono", "No hay número de teléfono.")
            return
        if not msg:
            messagebox.showwarning("Sin mensaje", "El mensaje está vacío.")
            return

        tel_clean = re.sub(r"[^0-9]", "", tel)
        if len(tel_clean) < 10:
            messagebox.showwarning("Teléfono inválido",
                "El número debe tener mínimo 10 dígitos con código de país.\n"
                "Ejemplo Colombia: 573001234567")
            return

        url = construir_url(tel_clean, msg)
        webbrowser.open(url)
        envio(tel_clean, msg)
        
    def enviar_todos(self):
        filtrados = self._pacientes_filtrados()
        
        # Obtenemos el mensaje personalizado si está activo
        es_personalizado = self.var_tipo_filtro.get() == "PERSONALIZADO"
        mensaje_fijo = self.text_personalizado.get("1.0", "end-1c").strip() if es_personalizado else None

        # Validamos pacientes
        validos = [p for p in filtrados if len(p["tel_limpio"]) >= 10]

        if not validos:
            messagebox.showwarning("Sin datos", "No hay pacientes válidos con teléfono.")
            return

        if es_personalizado and not mensaje_fijo:
            messagebox.showwarning("Mensaje vacío", "Escribe el mensaje personalizado primero.")
            return

        if not messagebox.askyesno("Confirmar", f"¿Enviar {len(validos)} mensajes?"):
            return

        for p in validos:
            # Si es personalizado, usamos el texto del cuadro; si no, el mensaje de la tabla
            msg_final = mensaje_fijo if es_personalizado else p["mensaje"]
            
            # Reemplazar {nombre} si existe en el mensaje personalizado
            nombre_corto = p["nombre"].split()[0].capitalize()
            msg_final = msg_final.replace("{nombre}", nombre_corto)

            url = construir_url(p["tel_limpio"], msg_final)
            webbrowser.open(url)
            
            # Forzamos a que la interfaz se actualice para que no diga "No responde"
            self.update()
            
            # Tiempo de espera para carga de WhatsApp Web
            time.sleep(10) 
            
            # Llamamos a tu función de auto_envio
            envio(p["tel_limpio"], msg_final)
            
            # Pausa breve antes del siguiente
            time.sleep(2)
            self.update()

        messagebox.showinfo("Finalizado", "Se ha completado la lista de envíos.")
    def _ordenar(self, col):
        data = [(self.tabla.set(k, col), k) for k in self.tabla.get_children()]
        data.sort()
        for i, (_, k) in enumerate(data):
            self.tabla.move(k, "", i)


if __name__ == "__main__":
    app = App()
    app.mainloop()
