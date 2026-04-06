import flet as ft

def main(page: ft.Page):
    page.title = "Test FilePicker"

    def on_result(e: ft.FilePickerResultEvent):
        if e.files:
            txt.value = "Seleccionados: " + ", ".join([f.name for f in e.files])
        else:
            txt.value = "No se seleccionó nada"
        page.update()

    picker = ft.FilePicker(on_result=on_result)
    page.overlay.append(picker)

    txt = ft.Text("Sin seleccionar")

    page.add(
        ft.ElevatedButton(
            "Seleccionar PDF",
            on_click=lambda _: picker.pick_files(
                allow_multiple=True,
                allowed_extensions=["pdf"]
            )
        ),
        txt
    )

ft.run(main)