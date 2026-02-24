import os
import re
import threading
from dataclasses import dataclass, asdict
from typing import Optional, Dict, Any, List

import flet as ft
import fitz  # PyMuPDF
import pandas as pd


# =========================
# 1) MODELO
# =========================
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


# =========================
# 2) UTILIDADES
# =========================
def _clean_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def _extract_first(patterns: List[str], text: str) -> Optional[str]:
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return _clean_spaces(m.group(1))
    return None


def read_pdf_text(pdf_path: str) -> str:
    text = ""
    with fitz.open(pdf_path) as doc:
        for page in doc:
            text += page.get_text()
    return text


def extract_conva_header(pdf_path: str) -> ConvaHeader:
    text = read_pdf_text(pdf_path)
    nombre_pdf = os.path.basename(pdf_path)

    header = ConvaHeader(
        apellidos_nombres=_extract_first([r"Apellidos\s+y\s+Nombres:\s*([^\n]+)"], text),
        codigo=_extract_first([r"\b(ID\s*Estudiante|Código):\s*(N\d+)"], text),
        carrera_upn=_extract_first([r"Carrera\s+(en\s+UPN|UPN):\s*([^\n]+)"], text),
        campus=_extract_first([r"Campus:\s*([^\n]+)"], text),
        plan_estudios=_extract_first([r"Plan\s+de\s+Estudios:\s*([0-9]+)"], text),
        fecha=_extract_first([r"\bFecha:\s*([0-9/]+)"], text),
        version_excelconva=_extract_first([r"Versión\s+.*?:\s*([^\n]+)"], text),
        total_creditos=_extract_first([r"TOTAL\s+DE\s+CRÉDITOS.*?([0-9]+)"], text),
        observaciones=_extract_first(
            [r"(Convalidación\s+por\s+paquete\s*\([^\)]+\))"], text
        ),
        nombre_pdf=nombre_pdf,
    )

    return header


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


# =========================
# 3) APP FLET
# =========================
def main(page: ft.Page):

    page.title = "Extractor Convalidaciones PRO"
    page.window_width = 1350
    page.window_height = 750

    registros: List[Dict[str, Any]] = []
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

    status = ft.Text("")
    progress = ft.ProgressBar(width=400, visible=False)

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(BASE_DIR, "resultado_conva.xlsx")

    tabla = ft.DataTable(
        columns=[ft.DataColumn(ft.Text(c)) for c in columns],
        rows=[],
        expand=True,
    )

    def render_table():
        tabla.rows = []
        for r in registros[:200]:  # 🔥 SOLO 200 FILAS
            tabla.rows.append(
                ft.DataRow(
                    cells=[ft.DataCell(ft.Text(str(r.get(c) or ""))) for c in columns]
                )
            )
        page.update()

    def procesar_pdfs(files):

        nonlocal registros
        registros.clear()

        total = len(files)
        errores = 0

        progress.visible = True
        progress.value = 0
        page.update()

        for i, f in enumerate(files, 1):
            try:
                h = extract_conva_header(f.path)
                registros.append(header_to_row(h))
            except Exception as ex:
                errores += 1

            if i % 25 == 0 or i == total:
                progress.value = i / total
                status.value = f"Procesando {i}/{total}"
                render_table()

        progress.visible = False
        status.value = f"Listo ✅ | PDFs: {total} | Errores: {errores}"
        render_table()
        page.update()

    def on_files_picked(e: ft.FilePickerResultEvent):
        if not e.files:
            return

        threading.Thread(
            target=procesar_pdfs,
            args=(e.files,),
            daemon=True
        ).start()

    file_picker = ft.FilePicker(on_result=on_files_picked)
    page.overlay.append(file_picker)

    def export_excel(_):
        if not registros:
            status.value = "No hay datos."
            page.update()
            return

        df = pd.DataFrame(registros, columns=columns)
        df.to_excel(excel_path, index=False)

        status.value = f"Excel exportado en: {excel_path}"
        page.update()

    page.add(
        ft.Column(
            [
                ft.Text("Extractor Masivo de Convalidaciones", size=20),
                ft.Row(
                    [
                        ft.ElevatedButton(
                            "Seleccionar PDFs",
                            on_click=lambda _: file_picker.pick_files(
                                allow_multiple=True,
                                allowed_extensions=["pdf"],
                            ),
                        ),
                        ft.ElevatedButton("Exportar Excel", on_click=export_excel),
                    ]
                ),
                progress,
                status,
                ft.Container(content=tabla, expand=True),
            ],
            expand=True,
        )
    )


if __name__ == "__main__":
    ft.app(target=main)