import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
from converter import convertir_pdf
import os
import threading


def limpiar_ruta(path):
    return path.strip().replace("{", "").replace("}", "")


def seleccionar_pdf():
    archivo = filedialog.askopenfilename(
        filetypes=[("Archivos PDF", "*.pdf")]
    )
    if archivo:
        entry_pdf.delete(0, tk.END)
        entry_pdf.insert(0, archivo)


def drop_pdf(event):
    archivo = limpiar_ruta(event.data)

    if archivo.lower().endswith(".pdf") and os.path.exists(archivo):
        entry_pdf.delete(0, tk.END)
        entry_pdf.insert(0, archivo)
        estado.set("PDF cargado correctamente")
    else:
        messagebox.showerror("Error", "Solo se permiten archivos PDF válidos")


def convertir_thread():
    convertir_btn.config(state="disabled")
    progress.start(10)
    estado.set("Convirtiendo...")

    try:
        pdf = entry_pdf.get()
        salida = os.path.splitext(pdf)[0] + ".docx"
        resultado = convertir_pdf(pdf, salida)

        messagebox.showinfo(
            "Conversión finalizada",
            resultado + f"\n\nArchivo generado:\n{salida}"
        )
    except Exception as e:
        messagebox.showerror("Error", str(e))
    finally:
        progress.stop()
        progress["value"] = 0
        estado.set("Arrastra un PDF o selecciónalo")
        convertir_btn.config(state="normal")


def convertir():
    pdf = entry_pdf.get()
    if not pdf or not os.path.exists(pdf):
        messagebox.showerror("Error", "Selecciona un archivo PDF válido")
        return

    threading.Thread(target=convertir_thread, daemon=True).start()


# ================== UI ==================

root = TkinterDnD.Tk()
root.title("Conversor PDF a Word (OCR)")
root.geometry("540x300")
root.resizable(False, False)

tk.Label(
    root,
    text="Conversor PDF a Word con OCR",
    font=("Segoe UI", 12, "bold")
).pack(pady=10)

# === ZONA DE ARRASTRE ===
drop_label = tk.Label(
    root,
    text="📄 Arrastra aquí tu archivo PDF",
    relief="ridge",
    width=55,
    height=4,
    bg="#f2f2f2"
)
drop_label.pack(pady=8)

drop_label.drop_target_register(DND_FILES)
drop_label.dnd_bind("<<Drop>>", drop_pdf)

entry_pdf = tk.Entry(root, width=70)
entry_pdf.pack(pady=5)

tk.Button(
    root,
    text="Seleccionar PDF",
    command=seleccionar_pdf,
    width=20
).pack(pady=6)

convertir_btn = tk.Button(
    root,
    text="Convertir a Word",
    command=convertir,
    width=20
)
convertir_btn.pack(pady=8)

progress = ttk.Progressbar(
    root,
    orient="horizontal",
    length=460,
    mode="indeterminate"
)
progress.pack(pady=6)

estado = tk.StringVar(value="Arrastra un PDF o selecciónalo")
tk.Label(root, textvariable=estado).pack(pady=6)

root.mainloop()
