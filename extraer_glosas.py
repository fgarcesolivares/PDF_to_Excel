#!/usr/bin/env python3
"""
extraer_glosas.py


USO:
    python3 extraer_glosas.py <carpeta_con_pdfs> [archivo_salida.xlsx]

Ejemplo:
    python3 extraer_glosas.py ./pdfs_glosas ./consolidado_glosas.xlsx

"""

import sys
import glob
import os
import re
from datetime import datetime

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}

# Columnas de salida, en el orden exacto solicitado
COLUMNAS_SALIDA = [
    "N#", "PERIODO", "MES", "Fecha Rad.", "Num Rad.", "Pref Fact.", "Num Fact.",
    "Num Doc Afil.", "Nombre Completo Afiliado", "Num Autoriz.", "Valor Servicio",
    "Cant. Fact.", "Valor Unitario", "Valor Total Glosa", "Valor Glosa Detalle",
    "ESTADO PAGO", "Descripcion Motivo", "OBSERVACIONES", "Cod Mot Glosa Genr",
    "Motivo Glosa General", "Cod Mot Glosa Espec", "Motivo Glosa Especifico",
    "Nombre reporte de glosa",
]

# Encabezado esperado dentro de la tabla del PDF (para detectar filas de encabezado repetidas)
ENCABEZADO_PDF = "fecha rad."


def limpiar_texto(valor):
    """Quita saltos de línea y espacios repetidos dejados por pdfplumber."""
    if valor is None:
        return ""
    texto = str(valor).replace("\n", " ").strip()
    texto = re.sub(r"\s{2,}", " ", texto)
    return texto


def parsear_fecha(fecha_str):
    """Convierte 'dd/mm/aaaa' a un objeto date. Devuelve None si no se puede."""
    fecha_str = limpiar_texto(fecha_str)
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(fecha_str, fmt).date()
        except ValueError:
            continue
    return None


def periodo_y_mes(fecha_rad):
    """Aplica la regla: mes 1-6 -> PRIMER, mes 7-12 -> SEGUNDO."""
    fecha = parsear_fecha(fecha_rad)
    if fecha is None:
        return "", ""
    periodo = "PRIMER" if fecha.month <= 6 else "SEGUNDO"
    mes = MESES_ES.get(fecha.month, "")
    return periodo, mes


def extraer_filas_de_pdf(ruta_pdf, advertencias):
    """Devuelve una lista de diccionarios, una fila por servicio glosado.

    Cualquier PDF o fila que se ignore queda registrado en la lista `advertencias`
    (se le hace append de un texto descriptivo) para reportarlo al final.
    """
    filas = []
    nombre_archivo = os.path.basename(ruta_pdf)
    se_encontro_alguna_tabla = False

    with pdfplumber.open(ruta_pdf) as pdf:
        for num_pagina, pagina in enumerate(pdf.pages, start=1):
            tablas = pagina.extract_tables()
            for tabla in tablas:
                se_encontro_alguna_tabla = True
                for num_fila, fila in enumerate(tabla, start=1):
                    fila_limpia = [limpiar_texto(c) for c in fila]

                    # Saltar filas de encabezado repetidas
                    if fila_limpia and fila_limpia[0].lower() == ENCABEZADO_PDF:
                        continue

                    # Saltar filas vacías o de totales (sin Fecha Rad. ni Num Rad.)
                    if not fila_limpia[0] and not fila_limpia[1]:
                        continue

                    # La tabla del PDF tiene 19 columnas; si difiere, se ignora la fila
                    if len(fila_limpia) < 19:
                        adelanto = " | ".join(fila_limpia[:5])
                        advertencias.append(
                            f"[{nombre_archivo}] Página {num_pagina}, fila {num_fila} de la "
                            f"tabla: IGNORADA (tiene {len(fila_limpia)} columnas, se esperaban "
                            f"19). Inicio de la fila: {adelanto}"
                        )
                        continue

                    (fecha_rad, num_rad, pref_fact, num_fact, num_doc_afil,
                     nombre_completo, num_autoriz, servicio, valor_servicio,
                     cant_fact, valor_unitario, valor_total_glosa,
                     valor_glosa_detalle, descripcion_motivo, observaciones,
                     cod_mot_genr, motivo_genr, cod_mot_espec,
                     motivo_espec) = fila_limpia[:19]

                    # Corrige números que quedaron partidos por un salto de línea
                    # justo después de un guión, ej. "34311- 2459719138"
                    num_autoriz = re.sub(r"-\s+", "-", num_autoriz)
                    num_rad = re.sub(r"\s+", "", num_rad)

                    periodo, mes = periodo_y_mes(fecha_rad)

                    filas.append({
                        "PERIODO": periodo,
                        "MES": mes,
                        "Fecha Rad.": fecha_rad,
                        "Num Rad.": num_rad,
                        "Pref Fact.": pref_fact,
                        "Num Fact.": num_fact,
                        "Num Doc Afil.": num_doc_afil,
                        "Nombre Completo Afiliado": nombre_completo,
                        "Num Autoriz.": num_autoriz,
                        "Valor Servicio": valor_servicio,
                        "Cant. Fact.": cant_fact,
                        "Valor Unitario": valor_unitario,
                        "Valor Total Glosa": valor_total_glosa,
                        "Valor Glosa Detalle": valor_glosa_detalle,
                        "ESTADO PAGO": "",
                        "Descripcion Motivo": descripcion_motivo,
                        "OBSERVACIONES": observaciones,
                        "Cod Mot Glosa Genr": cod_mot_genr,
                        "Motivo Glosa General": motivo_genr,
                        "Cod Mot Glosa Espec": cod_mot_espec,
                        "Motivo Glosa Especifico": motivo_espec,
                        "Nombre reporte de glosa": nombre_archivo,
                    })

    if not se_encontro_alguna_tabla:
        advertencias.append(
            f"[{nombre_archivo}] PDF IGNORADO POR COMPLETO: no se detectó ninguna tabla "
            f"en el documento (puede ser un PDF escaneado/imagen, o con un formato distinto)."
        )
    elif not filas:
        advertencias.append(
            f"[{nombre_archivo}] Se encontró tabla pero 0 filas de datos válidas "
            f"(revisar si el formato de esa tabla es distinto al esperado)."
        )

    return filas


def a_numero(valor):
    """Intenta convertir un texto tipo '144000.00' o '$ 144000' a float."""
    if valor is None:
        return None
    texto = str(valor).replace("$", "").replace(",", "").strip()
    if texto == "":
        return None
    try:
        return float(texto)
    except ValueError:
        return None


def construir_excel(filas, archivo_salida):
    wb = Workbook()
    ws = wb.active
    ws.title = "Glosas"

    fuente_normal = Font(name="Arial", size=10)
    fuente_encabezado = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    relleno_encabezado = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

    # Encabezados
    for col_idx, nombre_col in enumerate(COLUMNAS_SALIDA, start=1):
        celda = ws.cell(row=1, column=col_idx, value=nombre_col)
        celda.font = fuente_encabezado
        celda.fill = relleno_encabezado
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"

    columnas_numericas = {
        "Valor Servicio", "Cant. Fact.", "Valor Unitario",
        "Valor Total Glosa", "Valor Glosa Detalle",
    }

    for i, fila in enumerate(filas, start=1):
        fila_out = {"N#": i, **fila}
        for col_idx, nombre_col in enumerate(COLUMNAS_SALIDA, start=1):
            valor = fila_out.get(nombre_col, "")
            if nombre_col in columnas_numericas:
                num = a_numero(valor)
                valor = num if num is not None else valor
            celda = ws.cell(row=i + 1, column=col_idx, value=valor)
            celda.font = fuente_normal
            celda.alignment = Alignment(vertical="top", wrap_text=(nombre_col in
                                         ("Descripcion Motivo", "Nombre Completo Afiliado",
                                          "Motivo Glosa Especifico", "OBSERVACIONES")))
            if nombre_col in columnas_numericas and isinstance(valor, float):
                celda.number_format = "#,##0.00"

    # Ancho de columnas aproximado
    anchos = {
        "N#": 5, "PERIODO": 10, "MES": 11, "Fecha Rad.": 12, "Num Rad.": 17,
        "Pref Fact.": 9, "Num Fact.": 10, "Num Doc Afil.": 13,
        "Nombre Completo Afiliado": 26, "Num Autoriz.": 14, "Valor Servicio": 13,
        "Cant. Fact.": 9, "Valor Unitario": 13, "Valor Total Glosa": 14,
        "Valor Glosa Detalle": 14, "ESTADO PAGO": 12, "Descripcion Motivo": 45,
        "OBSERVACIONES": 18, "Cod Mot Glosa Genr": 10, "Motivo Glosa General": 15,
        "Cod Mot Glosa Espec": 10, "Motivo Glosa Especifico": 22,
        "Nombre reporte de glosa": 40,
    }
    for col_idx, nombre_col in enumerate(COLUMNAS_SALIDA, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = anchos.get(nombre_col, 15)

    wb.save(archivo_salida)


def main():
    if len(sys.argv) < 2:
        print("Uso: python3 extraer_glosas.py <carpeta_con_pdfs> [archivo_salida.xlsx]")
        sys.exit(1)

    carpeta = sys.argv[1]
    archivo_salida = sys.argv[2] if len(sys.argv) > 2 else "consolidado_glosas.xlsx"

    pdfs = sorted(glob.glob(os.path.join(carpeta, "*.pdf")))
    if not pdfs:
        print(f"No se encontraron PDFs en: {carpeta}")
        sys.exit(1)

    todas_las_filas = []
    advertencias = []
    for ruta_pdf in pdfs:
        print(f"Procesando: {os.path.basename(ruta_pdf)}")
        filas = extraer_filas_de_pdf(ruta_pdf, advertencias)
        print(f"  -> {len(filas)} fila(s) extraída(s)")
        todas_las_filas.extend(filas)

    construir_excel(todas_las_filas, archivo_salida)
    print(f"\nListo. {len(todas_las_filas)} fila(s) en total guardadas en: {archivo_salida}")

    if advertencias:
        print(f"\n⚠ ADVERTENCIAS ({len(advertencias)}) — revisa lo siguiente manualmente:")
        for adv in advertencias:
            print(f"  - {adv}")

        # También se guarda un log en un .txt junto al Excel, por si son muchas
        archivo_log = os.path.splitext(archivo_salida)[0] + "_advertencias.txt"
        with open(archivo_log, "w", encoding="utf-8") as f:
            f.write(f"Advertencias generadas al procesar la carpeta: {carpeta}\n")
            f.write(f"Fecha de ejecución: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            for adv in advertencias:
                f.write(f"- {adv}\n")
        print(f"\n(También quedaron guardadas en: {archivo_log})")
    else:
        print("\nSin advertencias: todos los PDFs y filas se procesaron correctamente.")


if __name__ == "__main__":
    main()