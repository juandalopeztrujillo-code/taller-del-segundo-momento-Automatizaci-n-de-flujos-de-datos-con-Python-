# Taller: Automatización de Flujos de Datos con Python
# De Excel (openpyxl) a Consola (tabulate) e Informe (ReportLab)


import os
import openpyxl
from tabulate import tabulate
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors


CARPETA = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_EXCEL = os.path.join(CARPETA, "ventas_tecnologia_v2.xlsx")
ARCHIVO_PDF = os.path.join(CARPETA, "reporte_ejecutivo_ventas.pdf")


ENCABEZADOS = ["ID", "Producto", "Categoría", "Precio ($)", "Ventas (Unid)"]
ALINEACION = ("center", "left", "center", "right", "right")



def cargar_datos():
    
    libro = openpyxl.load_workbook(ARCHIVO_EXCEL, data_only=True)
    hoja = libro.active
    print(f"Hoja cargada exitosamente: {hoja.title}")

    datos = []

    for fila in hoja.iter_rows(min_row=2, values_only=True):
        datos.append(list(fila))

    print(f"Total de registros leídos en memoria: {len(datos)}")
    return datos


def mostrar_tabla(datos, titulo):
    print(f"\n=== {titulo} ({len(datos)} PRODUCTOS) ===")
    print(tabulate(datos, headers=ENCABEZADOS, tablefmt="fancy_grid",
                   colalign=ALINEACION))


def filtrar_por_precio(datos, minimo=100):
    filtrados = []
    for fila in datos:
        precio = fila[3] 
        if precio > minimo:
            filtrados.append(fila)
    return filtrados

def calcular_metricas(datos):

    if len(datos) == 0:
        return 0, 0
    total_unidades = 0
    suma_precios = 0
    for fila in datos:
        suma_precios += fila[3]
        total_unidades += fila[4]
    precio_promedio = suma_precios / len(datos)
    return total_unidades, precio_promedio


def mostrar_metricas(datos):

    total_unidades, precio_promedio = calcular_metricas(datos)
    print("\n--- MÉTRICAS ---")
    print(f"Total de unidades vendidas: {total_unidades:,}")
    print(f"Precio promedio: ${precio_promedio:,.2f}")


def buscar_por_id(datos):
    """Pide un ID al usuario y muestra solo esa fila."""
    texto = input("Escribe el ID del producto: ").strip()

    if not texto.isdigit():
        print("Debes escribir un número entero.")
        return

    id_buscado = int(texto)
    for fila in datos:
        if fila[0] == id_buscado:
            mostrar_tabla([fila], f"RESULTADO DE BÚSQUEDA ID {id_buscado}")
            return


    print(f"⚠ Advertencia: el ID {id_buscado} no existe en el catálogo.")



def generar_pdf(datos, titulo="Informe Ejecutivo de Ventas Tecno"):
    """Crea el PDF con la tabla de productos y una fila de resumen."""
    doc = SimpleDocTemplate(ARCHIVO_PDF, pagesize=letter)
    estilos = getSampleStyleSheet()


    estilo_titulo = ParagraphStyle(
        "EstiloTitulo", parent=estilos["Title"],
        fontName="Helvetica-Bold", fontSize=18,
        textColor=colors.HexColor("#1A365D"), spaceAfter=10)
    estilo_cuerpo = ParagraphStyle(
        "EstiloCuerpo", parent=estilos["Normal"], fontSize=9)
    estilo_encabezado = ParagraphStyle(
        "EstiloEncabezado", parent=estilos["Normal"],
        fontName="Helvetica-Bold", fontSize=10,
        textColor=colors.whitesmoke, alignment=1)


    story = [
        Paragraph(titulo, estilo_titulo),
        Paragraph(f"<b>Total de Productos Procesados:</b> {len(datos)}", estilo_cuerpo),
        Spacer(1, 12),
    ]

    encabezados_pdf = [Paragraph(f"<b>{h}</b>", estilo_encabezado) for h in ENCABEZADOS]

    estilo_c = ParagraphStyle("C", parent=estilo_cuerpo, alignment=1)   
    estilo_l = ParagraphStyle("L", parent=estilo_cuerpo, alignment=0)  
    estilo_r = ParagraphStyle("R", parent=estilo_cuerpo, alignment=2)  

    filas_pdf = []
    for r in datos:
        filas_pdf.append([
            Paragraph(str(r[0]), estilo_c),
            Paragraph(str(r[1]), estilo_l),
            Paragraph(str(r[2]), estilo_c),
            Paragraph(f"${r[3]:,.2f}", estilo_r),
            Paragraph(str(r[4]), estilo_r),
        ])

    total_unidades, precio_promedio = calcular_metricas(datos)
    fila_resumen = [
        Paragraph("", estilo_c),
        Paragraph("<b>RESUMEN</b>", estilo_l),
        Paragraph("<b>Prom. precio / Total unid.</b>", estilo_c),
        Paragraph(f"<b>${precio_promedio:,.2f}</b>", estilo_r),
        Paragraph(f"<b>{total_unidades:,}</b>", estilo_r),
    ]


    anchos_cols = [40, 150, 110, 80, 80]
    tabla_pdf = Table([encabezados_pdf] + filas_pdf + [fila_resumen],
                      colWidths=anchos_cols, repeatRows=1)


    tabla_pdf.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F7FAFC")]),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E2E8F0")),
    ]))

    story.append(tabla_pdf)

    doc.build(story)
    print(f"\n[ÉXITO] Archivo '{os.path.basename(ARCHIVO_PDF)}' generado correctamente.")


def mostrar_menu():
    print("\n" + "=" * 40)
    print("   MENÚ - TECNOVENTAS")
    print("=" * 40)
    print("1) Ver catálogo completo")
    print("2) Filtrar por precio (> $100)")
    print("3) Buscar por ID")
    print("4) Exportar informe PDF")
    print("5) Salir")


def exportar_pdf_menu(datos):

    print("\n¿Qué deseas exportar?")
    print("  a) Catálogo completo")
    print("  b) Solo productos con precio > $100")
    opcion = input("Elige a o b: ").strip().lower()

    if opcion == "a":
        generar_pdf(datos)
    elif opcion == "b":
        generar_pdf(filtrar_por_precio(datos),
                    "Informe Ejecutivo de Ventas Tecno (Precio > $100)")
    else:
        print("⚠ Opción no válida. No se generó el PDF.")


def main():
    datos = cargar_datos()

    while True:
        mostrar_menu()
        opcion = input("Elige una opción (1-5): ").strip()

        if opcion == "1":
            mostrar_tabla(datos, "REPORTE DE VENTAS EN CONSOLA")
            mostrar_metricas(datos)
        elif opcion == "2":
            filtrados = filtrar_por_precio(datos)
            mostrar_tabla(filtrados, "PRODUCTOS CON PRECIO MAYOR A $100")
            mostrar_metricas(filtrados)
        elif opcion == "3":
            buscar_por_id(datos)
        elif opcion == "4":
            exportar_pdf_menu(datos)
        elif opcion == "5":
            print("¡Hasta luego!")
            break
        else:
            print("⚠ Opción no válida. Escribe un número del 1 al 5.")
5
if __name__ == "__main__":
    main()
