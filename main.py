import os
import re
import subprocess
import threading
from dataclasses import dataclass, asdict
from typing import Optional, Dict, Any, List

import flet as ft
import pdfplumber
import pandas as pd

# ✅ Tkinter (selector nativo Windows)
from tkinter import Tk, filedialog


@dataclass
class ConvaHeader:
    apellidos_nombres: Optional[str] = None
    codigo: Optional[str] = None
    carrera_upn: Optional[str] = None
    campus: Optional[str] = None
    plan_estudios: Optional[str] = None
    fecha: Optional[str] = None
    version_excelconva: Optional[str] = None
    total_creditos: Optional[str] = None
    observaciones: Optional[str] = None
    nombre_pdf: Optional[str] = None
    ruta_pdf: Optional[str] = None


def _clean_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def _limpiar_nombre(nombre_raw: Optional[str]) -> Optional[str]:
    if not nombre_raw:
        return None
    cortes = [" ID Estudiante:", " Código:", " CODIGO:", " ID ESTUDIANTE:"]
    nombre = nombre_raw
    low = nombre.lower()
    for c in cortes:
        idx = low.find(c.lower())
        if idx != -1:
            nombre = nombre[:idx]
            break
    return _clean_spaces(nombre)


def _extraer_observacion_paquete(text: str) -> Optional[str]:
    m = re.search(r"(Convalidación\s+por\s+paquete\s*\([^\)]+\))", text, re.IGNORECASE)
    if m:
        return _clean_spaces(m.group(1))
    m2 = re.search(r"(Convalidación\s+por\s+paquete)", text, re.IGNORECASE)
    if m2:
        return _clean_spaces(m2.group(1))
    return None


def _compile_patterns(patterns: List[str], flags=re.IGNORECASE) -> List[re.Pattern]:
    return [re.compile(p, flags) for p in patterns]


def _extract_first_compiled(patterns: List[re.Pattern], text: str) -> Optional[str]:
    for pat in patterns:
        m = pat.search(text)
        if m:
            val = m.group(1)
            if val is not None:
                return _clean_spaces(val)
    return None


PATRONES_NOMBRE = _compile_patterns([r"Apellidos\s+y\s+Nombres:\s*([^\n]+)"])
PATRONES_CODIGO = _compile_patterns([r"\bID\s*Estudiante:\s*(N\d+)", r"\bCódigo:\s*(N\d+)"])
PATRONES_CARRERA = _compile_patterns([r"Carrera\s+en\s+UPN:\s*([^\n]+)", r"Carrera\s+UPN:\s*([^\n]+)"])
PATRONES_CAMPUS = _compile_patterns([r"Campus:\s*([^\n]+)"])
PATRONES_PLAN = _compile_patterns([r"Plan\s+de\s+Estudios:\s*([0-9]+)"])
PATRONES_FECHA = _compile_patterns([r"\bFecha:\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})"])
PATRONES_VERSION = _compile_patterns([
    r"Versión\s+ExcelConva:\s*([0-9.]+)",
    r"Versión\s+Conva2025G\s*:\s*([0-9.]+)",
    r"Versión\s+Conva\s*:\s*([^\n]+)",
])
PATRONES_TOTAL = _compile_patterns([
    r"TOTAL\s+DE\s+CRÉDITOS\s*(?:o\s*Total\s*[:])?\s*([0-9]+)",
    r"\bTotal\s+([0-9]+)\b",
])


def extract_conva_header(pdf_path: str, max_pages: int = 4) -> ConvaHeader:
    nombre_pdf = os.path.basename(pdf_path)

    nombre_raw = None
    codigo = None
    carrera_raw = None
    campus = None
    plan = None
    fecha = None
    version = None
    total = None
    observ = None

    found = 0
    need_min = 7

    with pdfplumber.open(pdf_path) as pdf:
        if getattr(pdf, "is_encrypted", False):
            raise ValueError("PDF encriptado/no legible")

        n = min(len(pdf.pages), max_pages)

        for i in range(n):
            txt = pdf.pages[i].extract_text() or ""
            if not txt.strip():
                continue

            if nombre_raw is None:
                nombre_raw = _extract_first_compiled(PATRONES_NOMBRE, txt)
                if nombre_raw:
                    found += 1

            if codigo is None:
                codigo = _extract_first_compiled(PATRONES_CODIGO, txt)
                if codigo:
                    found += 1

            if carrera_raw is None:
                carrera_raw = _extract_first_compiled(PATRONES_CARRERA, txt)
                if carrera_raw:
                    found += 1

            if campus is None:
                campus = _extract_first_compiled(PATRONES_CAMPUS, txt)
                if campus:
                    found += 1

            if plan is None:
                plan = _extract_first_compiled(PATRONES_PLAN, txt)
                if plan:
                    found += 1

            if fecha is None:
                fecha = _extract_first_compiled(PATRONES_FECHA, txt)
                if fecha:
                    found += 1

            if version is None:
                version = _extract_first_compiled(PATRONES_VERSION, txt)
                if version:
                    found += 1

            if total is None:
                total = _extract_first_compiled(PATRONES_TOTAL, txt)
                if total:
                    found += 1

            if observ is None:
                observ = _extraer_observacion_paquete(txt)
                if observ:
                    found += 1

            if found >= need_min:
                break

    carrera_limpia = None
    if carrera_raw:
        carrera_limpia = _clean_spaces(re.sub(r"\s+Modalidad:.*$", "", carrera_raw, flags=re.IGNORECASE))

    return ConvaHeader(
        apellidos_nombres=_limpiar_nombre(nombre_raw),
        codigo=codigo,
        carrera_upn=carrera_limpia,
        campus=campus,
        plan_estudios=plan,
        fecha=fecha,
        version_excelconva=version,
        total_creditos=total,
        observaciones=observ,
        nombre_pdf=nombre_pdf,
        ruta_pdf=pdf_path,
    )


def header_to_row(h: ConvaHeader) -> Dict[str, Any]:
    d = asdict(h)
    return {
        "Apellidos y Nombres": d["apellidos_nombres"],
        "Código": d["codigo"],
        "Carrera en UPN": d["carrera_upn"],
        "Campus": d["campus"],
        "Plan de Estudios": d["plan_estudios"],
        "Fecha": d["fecha"],
        "Versión ExcelConva": d["version_excelconva"],
        "Total de Créditos": d["total_creditos"],
        "Observaciones": d["observaciones"],
        "Nombre_PDF": d["nombre_pdf"],
    }


def abrir_archivo(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError(f"No existe: {path}")
    if os.name == "nt":
        os.startfile(path)  # type: ignore
    elif hasattr(os, "uname") and os.uname().sysname == "Darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def main(page: ft.Page):
    page.title = "Extractor Convalidaciones (PDF → Excel)"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.window_width = 1100
    page.window_height = 520

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(BASE_DIR, "resultado_conva.xlsx")

    columns = [
        "Apellidos y Nombres",
        "Código",
        "Carrera en UPN",
        "Campus",
        "Plan de Estudios",
        "Fecha",
        "Versión ExcelConva",
        "Total de Créditos",
        "Observaciones",
        "Nombre_PDF",
    ]

    registros: List[Dict[str, Any]] = []
    lock = threading.Lock()

    cancel_flag = {"stop": False}
    running_flag = {"running": False}

    status = ft.Text("", selectable=True)
    lbl_actual = ft.Text("", selectable=True, size=12, color=ft.Colors.GREY_700)
    progress = ft.ProgressBar(value=0, expand=True, visible=False)

    # ✅ Flet 0.81+: usa Button
    btn_pick = ft.Button("Seleccionar PDFs", icon=ft.Icons.UPLOAD_FILE)
    btn_export = ft.Button("Exportar a Excel", icon=ft.Icons.SAVE_ALT)
    btn_open = ft.Button("Abrir Excel", icon=ft.Icons.FOLDER_OPEN)
    btn_clear = ft.OutlinedButton("Limpiar", icon=ft.Icons.DELETE_OUTLINE)
    btn_cancel = ft.OutlinedButton("Cancelar", icon=ft.Icons.STOP_CIRCLE, disabled=True)

    def set_enabled(enabled: bool):
        btn_pick.disabled = not enabled
        btn_export.disabled = not enabled
        btn_open.disabled = not enabled
        btn_clear.disabled = not enabled

    def on_pubsub_message(msg: Any):
        t = msg.get("type")

        if t == "progress":
            progress.visible = True
            progress.value = msg.get("p", 0)
            status.value = msg.get("status", "")
            lbl_actual.value = msg.get("current", "")
            page.update()

        elif t == "done":
            running_flag["running"] = False
            progress.visible = False
            btn_cancel.disabled = True
            set_enabled(True)
            status.value = msg.get("status", "Listo ✅")
            lbl_actual.value = ""
            page.update()

        elif t == "error":
            running_flag["running"] = False
            progress.visible = False
            btn_cancel.disabled = True
            set_enabled(True)
            status.value = msg.get("status", "Ocurrió un error.")
            page.update()

    page.pubsub.subscribe(on_pubsub_message)

    def export_excel(_):
        with lock:
            if not registros:
                status.value = "No hay datos para exportar."
                page.update()
                return
            df = pd.DataFrame(registros, columns=columns)
        df.to_excel(excel_path, index=False)
        status.value = f"Excel exportado ✅: {excel_path}"
        page.update()

    def open_excel(_):
        try:
            abrir_archivo(excel_path)
            status.value = f"Abierto ✅: {excel_path}"
        except Exception as ex:
            status.value = f"No se pudo abrir el Excel: {ex}"
        page.update()

    def limpiar(_):
        if running_flag["running"]:
            status.value = "No puedes limpiar mientras se procesa. Cancela primero."
            page.update()
            return
        with lock:
            registros.clear()
        progress.visible = False
        progress.value = 0
        lbl_actual.value = ""
        status.value = "Limpio ✅"
        page.update()

    def cancel(_):
        cancel_flag["stop"] = True
        btn_cancel.disabled = True
        status.value = "Cancelando... (terminará el PDF actual)"
        page.update()

    btn_export.on_click = export_excel
    btn_open.on_click = open_excel
    btn_clear.on_click = limpiar
    btn_cancel.on_click = cancel

    # ✅ Selección de PDFs con Tkinter (sin FilePicker de Flet)
    def pick_pdfs_with_tk(_):
        if running_flag["running"]:
            status.value = "Ya hay un proceso en ejecución."
            page.update()
            return

        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)  # asegura que el diálogo salga delante
        paths = filedialog.askopenfilenames(
            title="Seleccionar PDFs",
            filetypes=[("PDF", "*.pdf")]
        )
        root.destroy()

        if not paths:
            status.value = "No se seleccionaron archivos."
            page.update()
            return

        start_processing(list(paths))

    def start_processing(paths: List[str]):
        cancel_flag["stop"] = False
        running_flag["running"] = True

        total_files = len(paths)
        status.value = f"Procesando {total_files} PDF(s)..."
        progress.visible = True
        progress.value = 0
        btn_cancel.disabled = False
        set_enabled(False)
        page.update()

        def worker():
            ok = 0
            errores = 0
            try:
                for i, pdf_path in enumerate(paths, start=1):
                    if cancel_flag["stop"]:
                        break

                    current_name = os.path.basename(pdf_path)
                    try:
                        h = extract_conva_header(pdf_path, max_pages=4)
                        row = header_to_row(h)
                        ok += 1
                    except Exception as ex:
                        errores += 1
                        row = {
                            "Apellidos y Nombres": None,
                            "Código": None,
                            "Carrera en UPN": None,
                            "Campus": None,
                            "Plan de Estudios": None,
                            "Fecha": None,
                            "Versión ExcelConva": None,
                            "Total de Créditos": None,
                            "Observaciones": f"ERROR: {str(ex)}",
                            "Nombre_PDF": current_name,
                        }

                    with lock:
                        registros.append(row)

                    if i % 10 == 0 or i == total_files:
                        p = i / total_files
                        st = f"Procesando... {i}/{total_files} | OK: {ok} | Errores: {errores}"
                        page.pubsub.send_all({
                            "type": "progress",
                            "p": p,
                            "status": st,
                            "current": f"PDF actual: {current_name}",
                        })

                if cancel_flag["stop"]:
                    final = f"Cancelado 🛑 | Procesados: {len(registros)} | OK: {ok} | Errores: {errores}"
                else:
                    final = f"Listo ✅ | PDFs: {total_files} | OK: {ok} | Errores: {errores}"

                page.pubsub.send_all({"type": "done", "status": final})

            except Exception as ex:
                page.pubsub.send_all({"type": "error", "status": f"Error crítico: {ex}"})

        threading.Thread(target=worker, daemon=True).start()

    btn_pick.on_click = pick_pdfs_with_tk

    firma = ft.Row(
        controls=[ft.Text("Elaborado por: Ing Jesus Apolaya", italic=True, size=12, color=ft.Colors.GREY_700)],
        alignment=ft.MainAxisAlignment.END,
    )

    page.add(
        ft.Column(
            controls=[
                ft.Text("Extractor de Convalidaciones (PDF → Excel)", size=20, weight=ft.FontWeight.BOLD),
                ft.Text(
                    "Procesa 1 por 1 y genera un Excel final. Sin tabla en pantalla (más rápido).",
                    size=12, color=ft.Colors.GREY_700
                ),
                ft.Row([btn_pick, btn_export, btn_open, btn_clear, btn_cancel], wrap=True),
                progress,
                lbl_actual,
                ft.Divider(),
                status,
                ft.Container(height=10),
                firma
            ],
            expand=True
        )
    )


if __name__ == "__main__":
    ft.run(main)