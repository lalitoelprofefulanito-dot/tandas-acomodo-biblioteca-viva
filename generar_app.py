#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Regenera «Tandas de acomodo · Biblioteca Viva» desde el Registro de Acervo.

Uso (dentro de la carpeta de la app, donde están index.html y sw.js):

    python generar_app.py Registro_Acervo_BibliotecaViva_2026-11-09.xlsx

Qué hace:
  1. Lee la hoja «Registro» (título, autor, subserie, cantidad, ubicación, situación).
  2. Lee la hoja «Acomodo por anaquel» (orden dentro del entrepaño y «Dónde está hoy»).
  3. Extrae las portadas incrustadas en la columna I (imagen dentro de la celda)
     y las guarda en portadas/BV-####.jpg (máx. 420 px, JPEG calidad 72).
  4. Sustituye el bloque de datos de index.html y la versión de caché de sw.js.

Requiere: Python 3.9+, openpyxl y Pillow  (pip install openpyxl pillow)
"""
import io
import json
import posixpath
import re
import sys
import warnings
import zipfile
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
from PIL import Image

LADO_MAXIMO = 420
CALIDAD_JPEG = 72
MARCA_INICIO = "/*DATOS-INICIO*/"
MARCA_FIN = "/*DATOS-FIN*/"
VERSION_APP = "v5"  # cambia cuando cambia el código de index.html, para que los celulares renueven su copia

NS_HOJA = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL_DOC = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_REL_PAQ = "{http://schemas.openxmlformats.org/package/2006/relationships}"
NS_RICO = "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}"
NS_RICO_REL = "{http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel}"

# Columnas (índice desde 0) según el diccionario del Registro.
REG_UBICACION, REG_ID, REG_TITULO, REG_CANTIDAD, REG_AUTOR, REG_SITUACION = 0, 2, 3, 5, 11, 28
ACO_ORDEN, ACO_ID, ACO_DONDE = 1, 3, 11
PATRON_ID = re.compile(r"BV-\d{4}")
PATRON_UBICACION = re.compile(r"Anaquel\s*(\d+)\s*/\s*Entrepaño\s*(\d+)")


def texto(valor):
    return "" if valor is None else str(valor).strip()


def leer_registro(libro):
    """Devuelve {fila_excel: datos} de los libros vigentes con título."""
    hoja = libro["Registro"]
    encabezados = [texto(c) for c in next(hoja.iter_rows(min_row=1, max_row=1, values_only=True))]
    if "Id" not in encabezados[REG_ID] or "Título" not in encabezados[REG_TITULO]:
        sys.exit("La hoja «Registro» cambió de estructura: no encuentro «Id» en la columna C "
                 "ni «Título» en la columna D. No generé nada.")
    # «Subserie» se busca por su encabezado: sirve para decir «… de Pasos de luna» después del título.
    reg_subserie = next((i for i, e in enumerate(encabezados) if e.lower().startswith("subserie")), None)
    filas, omitidos = {}, []
    for numero, fila in enumerate(hoja.iter_rows(min_row=2, values_only=True), start=2):
        folio = texto(fila[REG_ID])
        if not PATRON_ID.fullmatch(folio):
            continue
        titulo = re.sub(r"\s+", " ", texto(fila[REG_TITULO]))
        if not titulo:
            omitidos.append(f"{folio} (sin título)")
            continue
        if texto(fila[REG_SITUACION]).lower() == "baja":
            omitidos.append(f"{folio} (baja)")
            continue
        lugar = PATRON_UBICACION.search(texto(fila[REG_UBICACION]))
        if not lugar:
            omitidos.append(f"{folio} (ubicación ilegible: «{texto(fila[REG_UBICACION])}»)")
            continue
        try:
            cantidad = int(fila[REG_CANTIDAD] or 1)
        except (TypeError, ValueError):
            cantidad = 1
        subserie = re.sub(r"\s+", " ", texto(fila[reg_subserie])) if reg_subserie is not None else ""
        filas[numero] = {
            "id": folio, "titulo": titulo, "autor": re.sub(r"\s+", " ", texto(fila[REG_AUTOR])),
            "subserie": "" if subserie.lower() == "sin registro" else subserie,
            "cantidad": cantidad, "anaquel": int(lugar.group(1)), "entrepano": int(lugar.group(2)),
        }
    return filas, omitidos, reg_subserie is not None


def leer_acomodo(libro):
    """Devuelve {folio: (orden, donde_esta_hoy)}; ignora filas de título intercaladas."""
    acomodo = {}
    for fila in libro["Acomodo por anaquel"].iter_rows(min_row=4, values_only=True):
        folio = texto(fila[ACO_ID]) if len(fila) > ACO_DONDE else ""
        if not PATRON_ID.fullmatch(folio):
            continue
        try:
            orden = int(float(texto(fila[ACO_ORDEN])))
        except ValueError:
            orden = 9999
        acomodo[folio] = (orden, texto(fila[ACO_DONDE]))
    return acomodo


def ruta_de_hoja(paquete, nombre):
    """Ruta interna (xl/worksheets/sheetN.xml) de una hoja a partir de su nombre."""
    relaciones = {r.get("Id"): r.get("Target")
                  for r in ET.fromstring(paquete.read("xl/_rels/workbook.xml.rels"))}
    for hoja in ET.fromstring(paquete.read("xl/workbook.xml")).iter(NS_HOJA + "sheet"):
        if hoja.get("name") == nombre:
            destino = relaciones[hoja.get(NS_REL_DOC + "id")]
            return destino.lstrip("/") if destino.startswith("/") else posixpath.join("xl", destino)
    sys.exit(f"No encuentro la hoja «{nombre}» en el libro.")


def portadas_por_fila(paquete, ruta_hoja):
    """Sigue la cadena de «imagen en celda» y devuelve {fila_excel: ruta de la imagen en el zip}."""
    nombres = set(paquete.namelist())
    necesarios = ["xl/metadata.xml", "xl/richData/rdrichvalue.xml", "xl/richData/richValueRel.xml",
                  "xl/richData/_rels/richValueRel.xml.rels"]
    if not all(n in nombres for n in necesarios):
        return {}
    metadatos = ET.fromstring(paquete.read("xl/metadata.xml"))
    # valueMetadata: bk n-1 → rc v → posición en futureMetadata «XLRICHVALUE» → rvb i
    valores = [bk.find(NS_HOJA + "rc").get("v")
               for bk in metadatos.find(NS_HOJA + "valueMetadata").findall(NS_HOJA + "bk")]
    futuros = []
    for bloque in metadatos.findall(NS_HOJA + "futureMetadata"):
        if bloque.get("name") == "XLRICHVALUE":
            futuros = [bk.find(".//" + NS_RICO + "rvb").get("i") for bk in bloque.findall(NS_HOJA + "bk")]
    # rdrichvalue: rv número i → el primer <v> es el índice de relación
    ricos = [rv.find(NS_RICO + "v").text
             for rv in ET.fromstring(paquete.read("xl/richData/rdrichvalue.xml")).findall(NS_RICO + "rv")]
    relaciones = [rel.get(NS_REL_DOC + "id")
                  for rel in ET.fromstring(paquete.read("xl/richData/richValueRel.xml")).findall(NS_RICO_REL + "rel")]
    destinos = {r.get("Id"): posixpath.normpath(posixpath.join("xl/richData", r.get("Target")))
                for r in ET.fromstring(paquete.read("xl/richData/_rels/richValueRel.xml.rels")).findall(NS_REL_PAQ + "Relationship")}
    resultado = {}
    hoja_xml = paquete.read(ruta_hoja).decode("utf-8")
    for celda in re.finditer(r'<c r="I(\d+)"([^>]*)>', hoja_xml):
        vm = re.search(r'\bvm="(\d+)"', celda.group(2))
        if not vm:
            continue
        try:
            rico = int(futuros[int(valores[int(vm.group(1)) - 1])])
            ruta = destinos[relaciones[int(ricos[rico])]]
        except (IndexError, KeyError, ValueError, TypeError):
            continue
        if ruta in nombres:
            resultado[int(celda.group(1))] = ruta
    return resultado


def guardar_portada(datos, destino):
    with Image.open(io.BytesIO(datos)) as imagen:
        imagen.seek(0)
        imagen = imagen.convert("RGBA")
        fondo = Image.new("RGB", imagen.size, "white")
        fondo.paste(imagen, mask=imagen.split()[3])
        fondo.thumbnail((LADO_MAXIMO, LADO_MAXIMO), Image.LANCZOS)
        fondo.save(destino, "JPEG", quality=CALIDAD_JPEG, optimize=True)


def fecha_de_datos(ruta_excel):
    en_nombre = re.search(r"(\d{4}-\d{2}-\d{2})", ruta_excel.name)
    return en_nombre.group(1) if en_nombre else date.fromtimestamp(ruta_excel.stat().st_mtime).isoformat()


def main():
    if len(sys.argv) < 2:
        sys.exit("Uso: python generar_app.py Registro_Acervo_BibliotecaViva_AAAA-MM-DD.xlsx [carpeta_de_la_app]")
    ruta_excel = Path(sys.argv[1])
    carpeta = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).resolve().parent
    ruta_index, ruta_sw, carpeta_portadas = carpeta / "index.html", carpeta / "sw.js", carpeta / "portadas"
    if not ruta_excel.is_file():
        sys.exit(f"No encuentro el Excel: {ruta_excel}")
    if not ruta_index.is_file() or not ruta_sw.is_file():
        sys.exit(f"En «{carpeta}» deben estar index.html y sw.js de la app.")
    pagina = ruta_index.read_text(encoding="utf-8")
    if pagina.count(MARCA_INICIO) != 1 or pagina.count(MARCA_FIN) != 1:
        sys.exit("index.html no tiene las marcas del bloque de datos; no lo modifiqué.")

    warnings.simplefilter("ignore")
    libro = openpyxl.load_workbook(ruta_excel, read_only=True, data_only=True)
    registro, omitidos, hay_columna_subserie = leer_registro(libro)
    acomodo = leer_acomodo(libro)
    libro.close()

    carpeta_portadas.mkdir(exist_ok=True)
    con_portada, fallidas = set(), []
    with zipfile.ZipFile(ruta_excel) as paquete:
        imagenes = portadas_por_fila(paquete, ruta_de_hoja(paquete, "Registro"))
        for fila, datos in registro.items():
            if fila not in imagenes:
                continue
            try:
                guardar_portada(paquete.read(imagenes[fila]), carpeta_portadas / f"{datos['id']}.jpg")
                con_portada.add(datos["id"])
            except Exception as error:  # imagen dañada: el libro queda «sin portada»
                fallidas.append(f"{datos['id']} ({error})")
    for sobrante in carpeta_portadas.glob("*.jpg"):
        if sobrante.stem not in con_portada:
            sobrante.unlink()

    acervo, sin_acomodo = [], []
    for datos in registro.values():
        orden, donde = acomodo.get(datos["id"], (9999, ""))
        if datos["id"] not in acomodo:
            sin_acomodo.append(datos["id"])
        acervo.append([datos["id"], datos["titulo"], datos["autor"], datos["cantidad"], datos["anaquel"],
                       datos["entrepano"], orden, "" if donde == "En anaquel" else donde,
                       1 if datos["id"] in con_portada else 0, datos["subserie"]])
    acervo.sort(key=lambda libro_: libro_[0])

    fecha = fecha_de_datos(ruta_excel)
    bloque = (f"{MARCA_INICIO}\nconst FECHA_DATOS = {json.dumps(fecha)};\nconst ACERVO = "
              + json.dumps(acervo, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
              + f";\n{MARCA_FIN}")
    inicio, fin = pagina.index(MARCA_INICIO), pagina.index(MARCA_FIN) + len(MARCA_FIN)
    ruta_index.write_text(pagina[:inicio] + bloque + pagina[fin:], encoding="utf-8")

    sw = ruta_sw.read_text(encoding="utf-8")
    sw_nuevo, cambios = re.subn(r"const CACHE = '[^']*';", f"const CACHE = 'tandas-bv-{fecha}-{len(acervo)}-{VERSION_APP}';", sw)
    if cambios != 1:
        sys.exit("sw.js no tiene la línea «const CACHE = '…';»; index.html sí quedó actualizado.")
    ruta_sw.write_text(sw_nuevo, encoding="utf-8")

    print(f"Libros en la app: {len(acervo)}  ·  con portada: {len(con_portada)}  ·  sin portada: {len(acervo) - len(con_portada)}")
    print(f"Fecha de los datos: {fecha}")
    if hay_columna_subserie:
        print(f"Con subserie: {sum(1 for l in acervo if l[9])}")
    else:
        print("No encontré la columna «Subserie» en la hoja «Registro»: la app no podrá filtrar por subserie.")
    if omitidos:
        print("Omitidos del Registro:", ", ".join(omitidos))
    if sin_acomodo:
        print("Sin renglón en «Acomodo por anaquel» (van al final de su entrepaño):", ", ".join(sin_acomodo))
    if fallidas:
        print("Portadas que no pude convertir:", ", ".join(fallidas))
    sin = sorted(l[0] for l in acervo if not l[8])
    if sin:
        print("Sin portada:", ", ".join(sin))


if __name__ == "__main__":
    main()
