"""Convierte el notebook ya ejecutado en un informe de Word (sin el código).

Toma el texto, las tablas y las gráficas de analisis_QQQ.ipynb tal como
quedaron la última vez que se corrió, así que primero hay que ejecutar el
notebook completo y guardarlo. Luego, desde la carpeta del proyecto:

    python informe/generar_informe_word.py
"""

import base64
import io
import json
import re
from datetime import date
from html.parser import HTMLParser
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

PROYECTO = Path(__file__).resolve().parent.parent
NOTEBOOK = PROYECTO / "analisis_QQQ.ipynb"
SALIDA = PROYECTO / "informe" / "Informe_Analisis_QQQ.docx"

TITULO = "Análisis del ETF QQQ (Invesco QQQ Trust)"
SUBTITULO = "Composición, pronóstico con modelos predictivos y cadena de suministro"
CURSO = "Taller de Ingeniería Económica — 2026"
AUTORES = ["Felipe López Chaparro"]
REPOSITORIO = "https://github.com/felipesimplest/analisis-qqq-ingenieria-economica"

AZUL = RGBColor(0x2A, 0x78, 0xD6)
TINTA = RGBColor(0x0B, 0x0B, 0x0B)
GRIS = RGBColor(0x52, 0x51, 0x4E)
FONDO_ENCABEZADO = "E3EEFC"
MAX_FILAS_TABLA = 40
ANCHO_UTIL = Inches(6.5)  # carta con márgenes de 1"

# Nombres de columnas del código -> nombres para leer en el informe
NOMBRES_COLUMNAS = {
    "valor": "Valor",
    "empresa": "Empresa",
    "tickers": "Tickers",
    "ticker": "Ticker",
    "peso_pct": "Peso (%)",
    "peso_acumulado_pct": "Peso acumulado (%)",
    "valor_miles_millones_usd": "Valor (miles de millones USD)",
    "num_empresas": "N.º de empresas",
    "sector_es": "Sector",
    "industria": "Industria",
    "pais": "País",
    "peso_invesco_pct": "Peso Invesco (%)",
    "peso_yahoo_pct": "Peso Yahoo (%)",
    "diferencia_pct": "Diferencia (pts)",
    "fecha_dato": "Fecha del pronóstico",
    "precio_base": "Precio base (USD)",
    "eslabon": "Eslabón",
    "peso_inicio_pct": "Peso al inicio (%)",
    "retorno_1a_pct": "Retorno 1 año (%)",
    "contribucion_pts": "Contribución (pts)",
    "True": "Sí",
    "False": "No",
}

# Frases que solo tienen sentido dentro del notebook
REEMPLAZOS_TEXTO = {
    "Este notebook analiza": "Este informe analiza",
    "Cada vez que se corre el notebook": "Cada vez que se ejecuta el análisis",
}


# ---------------------------------------------------------------- texto en línea

TOKENS_EN_LINEA = re.compile(r"(\*\*.+?\*\*|\$[^$]+\$|`[^`]+`|\*[^*\s][^*]*?\*|_\([^)]*\)_)")
SIMBOLOS_LATEX = {r"\times": "×", r"\propto": "∝", r"\sum": "Σ", r"\cdot": "·"}


def agregar_formula_en_linea(parrafo, formula, negrita=False):
    """Escribe LaTeX simple ($FC_t$, $(1+i)^t$) con subíndices y superíndices reales."""
    for comando, simbolo in SIMBOLOS_LATEX.items():
        formula = formula.replace(comando, simbolo)
    for parte in re.split(r"([_^](?:\{[^}]*\}|.))", formula):
        if not parte:
            continue
        run = parrafo.add_run(parte[1:].strip("{}") if parte[0] in "_^" else parte)
        run.italic = True
        run.bold = negrita
        run.font.subscript = parte[0] == "_"
        run.font.superscript = parte[0] == "^"


def agregar_texto(parrafo, texto, negrita=False):
    for token in TOKENS_EN_LINEA.split(texto):
        if not token:
            continue
        if token.startswith("**"):
            agregar_texto(parrafo, token[2:-2], negrita=True)
        elif token.startswith("$"):
            agregar_formula_en_linea(parrafo, token[1:-1], negrita)
        elif token.startswith("`"):
            run = parrafo.add_run(token[1:-1])
            run.font.name = "Consolas"
            run.font.size = Pt(9.5)
            run.bold = negrita
        elif token.startswith("_(") or (token.startswith("*") and len(token) > 2):
            run = parrafo.add_run(token[1:-1])
            run.italic = True
            run.bold = negrita
        else:
            parrafo.add_run(token).bold = negrita


# ---------------------------------------------------------------- tablas

class LectorTablaHTML(HTMLParser):
    """Lee la tabla HTML que pandas produce con display(df), respetando rowspan/colspan."""

    def __init__(self):
        super().__init__()
        self.filas, self.encabezado = [], 0
        self._fila = self._celda = None
        self._en_thead = False
        self._pendientes = {}  # columna -> (texto, filas restantes) por rowspan

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == "thead":
            self._en_thead = True
        elif tag == "tr":
            self._fila = []
        elif tag in ("th", "td") and self._fila is not None:
            self._celda = {"texto": "", "rowspan": int(attrs.get("rowspan", 1)), "colspan": int(attrs.get("colspan", 1))}

    def handle_endtag(self, tag):
        if tag == "thead":
            self._en_thead = False
        elif tag in ("th", "td") and self._celda is not None:
            self._fila.append(self._celda)
            self._celda = None
        elif tag == "tr" and self._fila is not None:
            self._cerrar_fila()

    def handle_data(self, data):
        if self._celda is not None:
            self._celda["texto"] += data

    def _cerrar_fila(self):
        fila, celdas = [], iter(self._fila)
        while True:
            col = len(fila)
            if col in self._pendientes:
                texto, restantes = self._pendientes.pop(col)
                fila.append(texto)
                if restantes > 1:
                    self._pendientes[col] = (texto, restantes - 1)
                continue
            celda = next(celdas, None)
            if celda is None:
                break
            for _ in range(celda["colspan"]):
                if celda["rowspan"] > 1:
                    self._pendientes[len(fila)] = (celda["texto"].strip(), celda["rowspan"] - 1)
                fila.append(celda["texto"].strip())
        self.filas.append(fila)
        self.encabezado += self._en_thead
        self._fila = None


def unir_encabezados(filas, n_encabezado):
    """pandas pone el nombre del índice en una segunda fila de encabezado; se combinan en una."""
    if n_encabezado < 2:
        return filas, n_encabezado
    ancho = max(len(f) for f in filas)
    combinada = []
    for col in range(ancho):
        partes = [f[col] for f in filas[:n_encabezado] if col < len(f) and f[col]]
        combinada.append(" ".join(dict.fromkeys(partes)))
    return [combinada] + filas[n_encabezado:], 1


def sombrear(celda, color):
    tcPr = celda._tc.get_or_add_tcPr()
    sombra = OxmlElement("w:shd")
    sombra.set(qn("w:val"), "clear")
    sombra.set(qn("w:color"), "auto")
    sombra.set(qn("w:fill"), color)
    tcPr.append(sombra)


def agregar_tabla(doc, filas, n_encabezado=1, texto_en_linea=False):
    ancho = max(len(f) for f in filas)
    filas = [f + [""] * (ancho - len(f)) for f in filas]
    recortadas = len(filas) - n_encabezado > MAX_FILAS_TABLA
    if recortadas:
        filas = filas[: n_encabezado + MAX_FILAS_TABLA]

    filas = [[NOMBRES_COLUMNAS.get(texto, texto) for texto in fila] for fila in filas]

    # Ancho de cada columna proporcional a su texto más largo (con mínimo y máximo)
    largos = [min(max(max(len(f[j]) for f in filas), 5), 32) for j in range(ancho)]
    anchos = [int(ANCHO_UTIL * l / sum(largos)) for l in largos]

    tabla = doc.add_table(rows=len(filas), cols=ancho)
    tabla.style = "Table Grid"
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabla.autofit = False
    tamaño = Pt(9) if ancho <= 5 else Pt(8)
    for j, columna in enumerate(tabla.columns):
        columna.width = anchos[j]
    for i, fila in enumerate(filas):
        for j, texto in enumerate(fila):
            celda = tabla.cell(i, j)
            celda.width = anchos[j]
            p = celda.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if texto_en_linea:
                agregar_texto(p, texto, negrita=i < n_encabezado)
            else:
                p.add_run(texto).bold = i < n_encabezado
            for run in p.runs:
                run.font.size = tamaño
            if i < n_encabezado:
                sombrear(celda, FONDO_ENCABEZADO)
            elif j > 0 and re.fullmatch(r"[-−+]?[\d.,]+%?", texto):
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if recortadas:
        nota(doc, f"(Se muestran las primeras {MAX_FILAS_TABLA} filas.)")
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def tabla_markdown(lineas):
    filas = [[c.strip() for c in l.strip().strip("|").split("|")] for l in lineas]
    return [f for f in filas if not all(re.fullmatch(r":?-+:?", c) for c in f)]


# ---------------------------------------------------------------- imágenes y fórmulas

def agregar_imagen(doc, origen, ancho=ANCHO_UTIL):
    doc.add_picture(origen, width=ancho)
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER


def agregar_formula(doc, latex):
    """Fórmula centrada: se dibuja con matplotlib; si trae \\text{...} se escribe como texto."""
    latex = latex.strip()
    if r"\text" in latex:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        agregar_formula_en_linea(p, re.sub(r"\\text\{([^}]*)\}", r"\1", latex))
        return
    fig = plt.figure(figsize=(0.01, 0.01))
    fig.text(0, 0, f"${latex}$", fontsize=16)
    buffer = io.BytesIO()
    fig.savefig(buffer, dpi=200, bbox_inches="tight", pad_inches=0.05, transparent=True)
    plt.close(fig)
    buffer.seek(0)
    agregar_imagen(doc, buffer, ancho=None)


# ---------------------------------------------------------------- markdown

def agregar_markdown(doc, md):
    for viejo, nuevo in REEMPLAZOS_TEXTO.items():
        md = md.replace(viejo, nuevo)
    lineas = md.splitlines()
    i = 0
    while i < len(lineas):
        linea = lineas[i].rstrip()
        if not linea.strip():
            i += 1
        elif linea.startswith("$$"):
            formula = [linea[2:]]
            while not formula[-1].rstrip().endswith("$$"):
                i += 1
                formula.append(lineas[i])
            agregar_formula(doc, " ".join(formula).replace("$$", ""))
            i += 1
        elif linea.startswith("|"):
            bloque = []
            while i < len(lineas) and lineas[i].startswith("|"):
                bloque.append(lineas[i])
                i += 1
            agregar_tabla(doc, tabla_markdown(bloque), texto_en_linea=True)
        elif m := re.match(r"!\[[^\]]*\]\(([^)]+)\)", linea):
            agregar_imagen(doc, str(PROYECTO / "imagenes" / Path(m.group(1)).name))  # la URL apunta a la copia local
            i += 1
        elif m := re.match(r"(#{1,4})\s+(.*)", linea):
            nivel = len(m.group(1))
            if nivel > 1:  # el título "#" ya está en la portada
                doc.add_heading(m.group(2).replace("*", ""), level=nivel - 1)
            i += 1
        elif m := re.match(r"\s*[-*]\s+(.*)", linea):
            agregar_texto(doc.add_paragraph(style="List Bullet"), m.group(1))
            i += 1
        elif m := re.match(r"\s*(\d+)\.\s+(.*)", linea):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.75)
            p.paragraph_format.first_line_indent = Cm(-0.5)
            p.add_run(f"{m.group(1)}. ")
            agregar_texto(p, m.group(2))
            i += 1
        else:
            parrafo = [linea]
            i += 1
            while i < len(lineas) and lineas[i].strip() and not re.match(r"\s*([-*#|!]|\d+\.\s|\$\$)", lineas[i]):
                parrafo.append(lineas[i].rstrip())
                i += 1
            if linea.startswith("**Taller"):  # la línea de curso/fecha ya está en la portada
                continue
            agregar_texto(doc.add_paragraph(), " ".join(parrafo))


# ---------------------------------------------------------------- salidas de código

def nota(doc, texto):
    p = doc.add_paragraph()
    run = p.add_run(texto)
    run.italic = True
    run.font.size = Pt(9)
    run.font.color.rgb = GRIS


def agregar_salidas(doc, salidas):
    for salida in salidas:
        datos = salida.get("data", {})
        if salida["output_type"] == "stream" and salida.get("name") == "stdout":
            for linea in "".join(salida["text"]).strip().splitlines():
                if linea.strip():
                    p = doc.add_paragraph()
                    p.paragraph_format.space_after = Pt(2)
                    run = p.add_run(linea.strip(" =").replace("===", "").strip())
                    run.font.size = Pt(10)
                    run.font.color.rgb = GRIS
        elif "image/png" in datos:
            agregar_imagen(doc, io.BytesIO(base64.b64decode(datos["image/png"])))
        elif "text/html" in datos and "<table" in "".join(datos["text/html"]):
            lector = LectorTablaHTML()
            lector.feed("".join(datos["text/html"]))
            filas, n_encabezado = unir_encabezados(lector.filas, lector.encabezado)
            agregar_tabla(doc, filas, n_encabezado)


# ---------------------------------------------------------------- documento

def campo(parrafo, instruccion):
    """Inserta un campo de Word (por ejemplo el número de página)."""
    run = parrafo.add_run()
    for tipo, texto in (("begin", None), (None, instruccion), ("separate", None), ("end", None)):
        if tipo:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tipo)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = texto
        run._r.append(el)


def preparar_estilos(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.font.color.rgb = TINTA
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.15
    for nivel, tamaño in ((1, 16), (2, 13), (3, 11.5)):
        estilo = doc.styles[f"Heading {nivel}"]
        estilo.font.name = "Calibri"
        estilo.font.size = Pt(tamaño)
        estilo.font.color.rgb = AZUL if nivel == 1 else TINTA
        estilo.font.bold = True
        estilo.paragraph_format.space_before = Pt(14 if nivel == 1 else 10)
        estilo.paragraph_format.space_after = Pt(4)
        estilo.paragraph_format.keep_with_next = True
        rFonts = estilo.element.rPr.rFonts
        for atributo in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            rFonts.attrib.pop(qn(atributo), None)


def portada(doc, fecha_datos):
    for _ in range(6):
        doc.add_paragraph()
    for texto, tamaño, color, negrita in (
        (TITULO, 26, AZUL, True),
        (SUBTITULO, 14, GRIS, False),
    ):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(texto)
        run.font.size = Pt(tamaño)
        run.font.color.rgb = color
        run.bold = negrita
    for _ in range(4):
        doc.add_paragraph()
    for texto in [CURSO, *AUTORES, f"Fecha de entrega: {date.today():%d/%m/%Y}", f"Datos de mercado al: {fecha_datos}"]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(texto).font.size = Pt(12)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"Código y datos: {REPOSITORIO}")
    run.font.size = Pt(10)
    run.font.color.rgb = AZUL
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def anexo_reproducir(doc):
    doc.add_heading("Anexo: cómo reproducir este análisis", level=1)
    for texto in (
        f"Todo el código, los datos y este informe están en el repositorio **{REPOSITORIO}**.",
        "Para ejecutarlo sin instalar nada, se usa el botón **Abrir en Colab** del repositorio. En un computador "
        "con Python basta con `pip install -r requirements.txt` y abrir el notebook `analisis_QQQ.ipynb`.",
        "Los precios se descargan en vivo de Yahoo Finance, así que al volver a ejecutar el notebook las cifras "
        "se actualizan a la fecha del día. Este informe se regenera con `python informe/generar_informe_word.py`.",
    ):
        agregar_texto(doc.add_paragraph(), texto)


def fecha_de_los_datos(notebook):
    for celda in notebook["cells"]:
        for salida in celda.get("outputs", []):
            texto = "".join(salida.get("text", ""))
            if m := re.search(r"Precio de \w+ al (\d{4}-\d{2}-\d{2})", texto):
                return m.group(1)
    return "sin ejecutar"


def main():
    notebook = json.loads(NOTEBOOK.read_text(encoding="utf-8"))
    if not any(c.get("outputs") for c in notebook["cells"] if c["cell_type"] == "code"):
        raise SystemExit("El notebook no tiene resultados: ejecútalo completo y guárdalo antes de generar el informe.")

    doc = Document()
    seccion = doc.sections[0]
    seccion.page_width, seccion.page_height = Inches(8.5), Inches(11)
    for lado in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(seccion, lado, Inches(1))
    preparar_estilos(doc)
    portada(doc, fecha_de_los_datos(notebook))

    cuerpo = doc.add_section(WD_SECTION.CONTINUOUS)
    pie = cuerpo.footer
    pie.is_linked_to_previous = False
    p = pie.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{TITULO} — página ")
    run.font.size = Pt(9)
    campo(p, "PAGE")

    for celda in notebook["cells"]:
        if celda["cell_type"] == "markdown":
            agregar_markdown(doc, "".join(celda["source"]))
        else:
            agregar_salidas(doc, celda.get("outputs", []))
    anexo_reproducir(doc)

    SALIDA.parent.mkdir(exist_ok=True)
    doc.save(SALIDA)
    print(f"Informe guardado en {SALIDA}")


if __name__ == "__main__":
    main()
