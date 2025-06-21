# ───────── GUI Traductor + Plantilla WCA (Prime 10) ─────────
from __future__ import annotations
from pathlib import Path
import os, subprocess, sys, tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

import traductor                         # ← contiene WARNINGS
from auto_mathcad import rellenar_plantilla_wca


# ──────────── GUI callbacks ─────────────────────────────────
def seleccionar_archivo() -> None:
    ruta = filedialog.askopenfilename(
        title="Selecciona un archivo de datos",
        filetypes=[
            ("LTspice / SIMetrix", "*.net *.bom"),
            ("Archivos de texto",   "*.txt *.csv"),
            ("Todos los archivos",  "*.*"),
        ]
    )
    if ruta:
        ruta_archivo.set(ruta)
        _mostrar_contenido(ruta)


def _mostrar_contenido(ruta: str) -> None:
    text_area.delete("1.0", tk.END)
    for enc in ("utf-8", "latin-1"):
        try:
            text_area.insert(tk.END, Path(ruta).read_text(encoding=enc))
            break
        except UnicodeDecodeError:
            continue
    else:
        messagebox.showerror(
            "Error de lectura",
            "No se pudo leer el archivo con codificación UTF-8 ni Latin-1."
        )


def _abrir_excel(path: Path) -> None:
    """Lanza el .xlsx con la app asociada al sistema (Windows / macOS / Linux)."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)                       # type: ignore[attr-defined]
        elif sys.platform.startswith("darwin"):
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showwarning("Excel",
                               f"No se pudo abrir el archivo automáticamente:\n{e}\n\n"
                               f"Ábrelo manualmente:\n{path}")


def procesar_archivo() -> None:
    archivo = ruta_archivo.get()
    if not archivo:
        messagebox.showwarning("Advertencia", "Selecciona un archivo primero.")
        return

    # 1) Preguntamos H(s)…  (se puede omitir)
    h_s = simpledialog.askstring(
        "Función de transferencia (opcional)",
        "Escribe H(s) o pulsa Cancelar si tu plantilla ya la define:",
        parent=ventana
    )
    if h_s is None:
        h_s = ""

    # 2) Traductor  ─ genera Entrada_Datos_01.xlsx y llena traductor.WARNINGS
    try:
        ext = Path(archivo).suffix.lower()
        if ext == ".net":
            info_trad = traductor.procesar_net(archivo, h_s)
        elif ext == ".bom":
            info_trad = traductor.procesar_bom(archivo, h_s)
        else:
            info_trad = traductor.procesar_generico(archivo, h_s)
    except Exception as err:
        messagebox.showerror("Traductor", f"Error durante la conversión:\n{err}")
        return

    xlsx_path = Path(traductor.DEST_XLSX).resolve()

    # 2-bis) Avisos inmediatos ────────────────────────────────────────────
    if traductor.WARNINGS:
        messagebox.showwarning(
            "Advertencias de conversión",
            "Se detectaron algunos datos que no pudieron interpretarse "
            "automáticamente:\n\n"
            + "\n".join(f"• {w}" for w in traductor.WARNINGS) +
            "\n\nRevisa y corrige el Excel si es necesario antes de continuar."
        )

    # 3) ¿Revisar / editar manualmente el Excel?
    if messagebox.askyesno(
        "Editar Excel",
        "El archivo ‘Entrada_Datos_01.xlsx’ se ha generado.\n\n"
        "¿Quieres revisarlo o modificar algún dato a mano\n"
        "antes de enviarlo a la plantilla de Mathcad?"
    ):
        _abrir_excel(xlsx_path)
        messagebox.showinfo(
            "Edición manual",
            "Realiza los cambios, guarda y cierra el Excel.\n"
            "Pulsa Aceptar para continuar cuando hayas terminado."
        )

    # 4) ¿Plantilla ya abierta?
    usar_abierta = messagebox.askyesno(
        "Plantilla WCA",
        "¿Tienes abierta la plantilla de Mathcad?\n\n"
        "•  Sí  → se usará la hoja activa.\n"
        "•  No  → podrás seleccionar el archivo .mcdx."
    )
    plantilla = None
    if not usar_abierta:
        plantilla = filedialog.askopenfilename(
            title="Selecciona la plantilla WCA (.mcdx)",
            filetypes=[("Mathcad Prime", "*.mcdx")],
        )
        if not plantilla:
            messagebox.showinfo(
                "Cancelado",
                "No se seleccionó plantilla.\n"
                "Operación cancelada."
            )
            return

    # 5) Llamada a auto_mathcad
    try:
        rellenar_plantilla_wca(xlsx_path, plantilla if plantilla else None)
        msg_prime = "Plantilla WCA actualizada correctamente."
    except Exception as err:
        msg_prime = f"Mathcad no pudo completarse:\n{err}"

    
    # 6) Resumen  ─ solo la primera línea (sin advertencias)
    resumen = info_trad.splitlines()[0]          
    messagebox.showinfo("Procesamiento completo", f"{resumen}\n\n{msg_prime}")


# ──────────── Construcción GUI ─────────────────────────────
ventana = tk.Tk()
ventana.title("Traductor → Plantilla WCA (Mathcad Prime 10)")
ventana.geometry("850x520")

ruta_archivo = tk.StringVar(value="")

frame_top = tk.Frame(ventana)
frame_top.pack(pady=10, fill="x")

tk.Button(
    frame_top, text="Seleccionar archivo (.net / .bom / .csv)",
    command=seleccionar_archivo
).pack(side="left", padx=6)

tk.Entry(
    frame_top, textvariable=ruta_archivo, width=95, state="readonly"
).pack(side="left", fill="x", expand=True)

text_area = tk.Text(ventana, height=20, width=110, wrap="none")
text_area.pack(padx=10, pady=10, fill="both", expand=True)

tk.Button(
    ventana, text="Procesar y rellenar plantilla WCA",
    command=procesar_archivo
).pack(pady=12)

ventana.mainloop()
