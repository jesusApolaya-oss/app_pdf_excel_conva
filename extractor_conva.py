import os
import re
import subprocess
from dataclasses import dataclass, asdict
from typing import Optional, Dict, Any, List

import flet as ft
import pdfplumber
import pandas as pd


# =========================
# 1) Modelo de datos
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
    ruta_pdf: Optional[str] = None  # opcional (trazabilidad)


# =========================
# 2) Utilidades
# =========================
def _clean_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def _extract_first(patterns: List[str], text: str, flags=re.IGNORECASE) -> Optional[str]:
    for pat in patterns:
        m = re.search(pat, text, flags)
        if m:
            val = m.group(1)
            if val is not None:
                return _clean_spaces(val)
    return None


def read_pdf_text(pdf_path: str) -> str:
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
    return "\n".join(parts).replace("\r", "\n")


def _limpiar_nombre(nombre_raw: Optional[str]) -> Optional[str]:
    """
    Quita 'ID Estudiante: ...' o 'Código: ...' si vienen pegados en la misma línea.
    """
    if not nombre_raw:
        return None
    # Cortes típicos
    cortes = [" ID Estudiante:", " Código:", " CODIGO:", " ID ESTUDIANTE:"]
    nombre = nombre_raw
    for c in cortes:
        if c.lower() in nombre.lower():
            # cortar por índice case-insensitive
            idx = nombre.lower().find(c.lower())
            nombre = nombre[:idx]
            break
    return _clean_spaces(nombre)


def _extraer_observacion_paquete(text: str) -> Optional[str]:
    """
    Busca el término 'Convalidación por paquete' y devuelve el texto completo esperado
    ejemplo: Convalidación por paquete (1-2-3-4)

    Si no encuentra, devuelve None.
    """
    # Caso típico con paréntesis
    m = re.search(r"(Convalidación\s+por\s+paquete\s*\([^\)]+\))", text, re.IGNORECASE)
    if m:
        return _clean_spaces(m.group(1))

    # Si viniera sin paréntesis (lo dejamos como frase base)
    m2 = re.search(r"(Convalidación\s+por\s+paquete)", text, re.IGNORECASE)
    if m2:
        return _clean_spaces(m2.group(1))

    return None


def extract_conva_header(pdf_path: str) -> ConvaHeader:
    text = read_pdf_text(pdf_path)
    nombre_pdf = os.path.basename(pdf_path)

    patrones_nombre = [
        r"Apellidos\s+y\s+Nombres:\s*([^\n]+)",
    ]

    patrones_codigo = [
        r"\bID\s*Estudiante:\s*(N\d+)",
        r"\bCódigo:\s*(N\d+)",
    ]

    patrones_carrera = [
        r"Carrera\s+en\s+UPN:\s*([^\n]+)",
        r"Carrera\s+UPN:\s*([^\n]+)",
    ]

    patrones_campus = [
        r"Campus:\s*([^\n]+)",
    ]

    patrones_plan = [
        r"Plan\s+de\s+Estudios:\s*([0-9]+)",
    ]

    patrones_fecha = [
        r"\bFecha:\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})",
    ]

    patrones_version = [
        r"Versión\s+ExcelConva:\s*([0-9.]+)",
        r"Versión\s+Conva2025G\s*:\s*([0-9.]+)",
        r"Versión\s+Conva\s*:\s*([^\n]+)",  # puede ser "Manual"
    ]

    patrones_total = [
        r"TOTAL\s+DE\s+CRÉDITOS\s*(?:o\s*Total\s*[:])?\s*([0-9]+)",
        r"\bTotal\s+([0-9]+)\b",
    ]

    nombre_raw = _extract_first(patrones_nombre, text)
    carrera_raw = _extract_first(patrones_carrera, text)

    header = ConvaHeader(
        apellidos_nombres=_limpiar_nombre(nombre_raw),
        codigo=_extract_first(patrones_codigo, text),
        carrera_upn=_clean_spaces(re.sub(r"\s+Modalidad:.*$", "", carrera_raw, flags=re.IGNORECASE)) if carrera_raw else None,
        campus=_extract_first(patrones_campus, text),
        plan_estudios=_extract_first(patrones_plan, text),
        fecha=_extract_first(patrones_fecha, text),
        version_excelconva=_extract_first(patrones_version, text),
        total_creditos=_extract_first(patrones_total, text),
        observaciones=_extraer_observacion_paquete(text),  # 👈 mejora clave
        nombre_pdf=nombre_pdf,
        ruta_pdf=pdf_path,
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
        # si luego quieres mostrar también ruta:
        # "Ruta_PDF": d["ruta_pdf"],
    }


def abrir_archivo(path: str):
    """Abre el archivo con el programa por defecto (Win/Mac/Linux)."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"No existe: {path}")

    if os.name == "nt":
        os.startfile(path)  # type: ignore
    elif os.uname().sysname == "Darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


# =========================
# 3) App Flet
# =========================
def main(page: ft.Page):
    page.title = "Extractor Convalidaciones (PDF → Tabla → Excel)"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.window_width = 1350
    page.window_height = 740

    registros: List[Dict[str, Any]] = []
    status = ft.Text("", selectable=True)

    # Guardar Excel en la MISMA carpeta del .py
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

    # Tabla con mejor separación visual (líneas + espaciado + borde)
    tabla = ft.DataTable(
        columns=[ft.DataColumn(ft.Text(c, weight=ft.FontWeight.BOLD)) for c in columns],
        rows=[],
        expand=True,
        column_spacing=22,
        horizontal_margin=12,
        divider_thickness=1.2,
        border=ft.border.all(1, ft.Colors.GREY_300),
        border_radius=10,
    )

    def cell_text(value: Any) -> ft.Text:
        """
        Evita que todo se "pegue" visualmente:
        - max_lines controla altura
        - overflow maneja textos largos
        """
        return ft.Text(
            str(value or ""),
            max_lines=3,
            overflow=ft.TextOverflow.ELLIPSIS,
        )

    def render_table():
        tabla.rows = []
        for r in registros:
            tabla.rows.append(
                ft.DataRow(
                    cells=[ft.DataCell(cell_text(r.get(c))) for c in columns]
                )
            )
        page.update()

    def on_files_picked(e: ft.FilePickerResultEvent):
        if not e.files:
            status.value = "No se seleccionaron archivos."
            page.update()
            return

        status.value = f"Procesando {len(e.files)} PDF(s)..."
        page.update()

        errores = 0
        for f in e.files:
            try:
                h = extract_conva_header(f.path)
                registros.append(header_to_row(h))
            except Exception as ex:
                errores += 1
                registros.append({
                    "Apellidos y Nombres": None,
                    "Código": None,
                    "Carrera en UPN": None,
                    "Campus": None,
                    "Plan de Estudios": None,
                    "Fecha": None,
                    "Versión ExcelConva": None,
                    "Total de Créditos": None,
                    "Observaciones": f"ERROR: {str(ex)}",
                    "Nombre_PDF": os.path.basename(f.path),
                })

        render_table()
        status.value = f"Listo ✅ | PDFs: {len(e.files)} | Errores: {errores}"
        page.update()

    file_picker = ft.FilePicker(on_result=on_files_picked)
    page.overlay.append(file_picker)

    def export_excel(_):
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
        registros.clear()
        tabla.rows = []
        status.value = "Tabla limpia."
        page.update()

    # Firma abajo a la derecha
    firma = ft.Row(
        controls=[
            ft.Text("Elaborado por: Ing Jesus Apolaya", italic=True, size=12, color=ft.Colors.GREY_700)
        ],
        alignment=ft.MainAxisAlignment.END,
    )

    page.add(
        ft.Column(
            [
                ft.Text("Extractor de Resultados de Convalidación", size=20, weight=ft.FontWeight.BOLD),
                ft.Text("PDF → extracción de cabecera → tabla → exportación Excel", size=12, color=ft.Colors.GREY_700),

                ft.Row(
                    [
                        ft.ElevatedButton(
                            "Seleccionar PDFs",
                            icon=ft.Icons.UPLOAD_FILE,
                            on_click=lambda _: file_picker.pick_files(
                                allow_multiple=True,
                                allowed_extensions=["pdf"],
                            ),
                        ),
                        ft.ElevatedButton(
                            "Exportar a Excel",
                            icon=ft.Icons.SAVE_ALT,
                            on_click=export_excel,
                        ),
                        ft.ElevatedButton(
                            "Abrir Excel",
                            icon=ft.Icons.FOLDER_OPEN,
                            on_click=open_excel,
                        ),
                        ft.OutlinedButton(
                            "Limpiar",
                            icon=ft.Icons.DELETE_OUTLINE,
                            on_click=limpiar,
                        ),
                    ],
                    wrap=True,
                ),

                ft.Divider(),
                ft.Container(content=tabla, expand=True),
                ft.Divider(),
                ft.Row([status], alignment=ft.MainAxisAlignment.START),
                firma,  # 👈 firma inferior derecha
            ],
            expand=True,
        )
    )


if __name__ == "__main__":
    # pip install flet pdfplumber pandas openpyxl
    ft.app(target=main)
