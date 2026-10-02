"""
Rutina de generación de informes de inspección estructural.

1. Carga una plantilla Word (.docx) con etiquetas docxtpl/Jinja.
2. Pide la ruta de la carpeta con las fotografías.
3. Clasifica cada fotografía por tipo de defecto, la inserta en su espacio
   de la plantilla y genera un resumen con el número total de cada tipo.

Cómo se determina el tipo de defecto de cada imagen (por orden de prioridad):
  a) Archivo "inspeccion.csv" en la carpeta de imágenes, con columnas
     archivo;defecto;elemento;ubicacion;severidad;observaciones
  b) Subcarpeta en la que está la imagen:   fotos/Fisura/IMG_001.jpg
  c) Nombre del archivo:  <defecto>_<elemento>_<n>.jpg   ->  fisura_viga-P2_01.jpg

Imágenes con nombre reservado (no cuentan como defecto) se colocan en las
etiquetas {{ img_<nombre> }} de la plantilla, p. ej. "portada.jpg" ->
{{ img_portada }}, "ubicacion.png" -> {{ img_ubicacion }}.

Uso:
    python generar_informe.py                      (modo interactivo)
    python generar_informe.py -i ./fotos -p plantilla.docx -o informe.docx
"""
import argparse
import csv
import json
import sys
import tempfile
import unicodedata
from collections import Counter
from datetime import date
from pathlib import Path

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from PIL import Image, ImageOps

BASE = Path(__file__).parent
PLANTILLA_POR_DEFECTO = BASE / "plantillas" / "plantilla_inspeccion.docx"
EXTENSIONES = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".gif", ".webp"}
ANCHO_FOTO_MM = 150
LADO_MAX_PX = 1600  # se reducen las fotos para que el .docx no pese demasiado

# Normalización de nombres de defecto: alias (sin tildes, minúsculas) -> nombre oficial.
# Amplíe esta tabla con la nomenclatura de su empresa.
CATALOGO_DEFECTOS = {
    "fisura": "Fisura",
    "fisuras": "Fisura",
    "grieta": "Grieta",
    "grietas": "Grieta",
    "corrosion": "Corrosión de armaduras",
    "oxido": "Corrosión de armaduras",
    "armadura expuesta": "Armadura expuesta",
    "acero expuesto": "Armadura expuesta",
    "desprendimiento": "Desprendimiento de recubrimiento",
    "descascaramiento": "Desprendimiento de recubrimiento",
    "spalling": "Desprendimiento de recubrimiento",
    "eflorescencia": "Eflorescencia",
    "eflorescencias": "Eflorescencia",
    "humedad": "Humedad / Filtración",
    "filtracion": "Humedad / Filtración",
    "deformacion": "Deformación / Flecha",
    "flecha": "Deformación / Flecha",
    "asentamiento": "Asentamiento",
    "coquera": "Coquera / Nido de grava",
    "nido de grava": "Coquera / Nido de grava",
    "segregacion": "Coquera / Nido de grava",
    "carbonatacion": "Carbonatación",
    "pudricion": "Pudrición (madera)",
    "xilofagos": "Ataque de xilófagos",
}

SEVERIDADES = ["Alta", "Media", "Baja", "Sin especificar"]
ALIAS_SEVERIDAD = {"a": "Alta", "alta": "Alta", "grave": "Alta",
                   "m": "Media", "media": "Media", "moderada": "Media",
                   "b": "Baja", "baja": "Baja", "leve": "Baja"}


# --------------------------------------------------------------------------- #
# Utilidades
# --------------------------------------------------------------------------- #
def _sin_tildes(texto: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", texto)
                   if unicodedata.category(c) != "Mn")


def _clave(texto: str) -> str:
    return " ".join(_sin_tildes(texto).lower().replace("_", " ").replace("-", " ").split())


def normalizar_defecto(texto: str) -> str:
    clave = _clave(texto)
    if not clave:
        return "Sin clasificar"
    return CATALOGO_DEFECTOS.get(clave, clave.capitalize())


def normalizar_severidad(texto: str) -> str:
    return ALIAS_SEVERIDAD.get(_clave(texto or ""), (texto or "").strip().capitalize() or "Sin especificar")


def preguntar(mensaje: str, defecto: str = "") -> str:
    sufijo = f" [{defecto}]" if defecto else ""
    try:
        valor = input(f"{mensaje}{sufijo}: ").strip().strip('"').strip("'")
    except EOFError:
        valor = ""
    return valor or defecto


def preparar_imagen(ruta: Path, tmpdir: Path) -> Path:
    """Corrige la orientación EXIF y reduce la resolución; devuelve un JPEG temporal."""
    with Image.open(ruta) as img:
        img = ImageOps.exif_transpose(img)
        img.thumbnail((LADO_MAX_PX, LADO_MAX_PX))
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        destino = tmpdir / f"{len(list(tmpdir.iterdir())):04d}_{ruta.stem}.jpg"
        img.save(destino, "JPEG", quality=85)
    return destino


# --------------------------------------------------------------------------- #
# Lectura y clasificación de imágenes
# --------------------------------------------------------------------------- #
def leer_csv(carpeta: Path) -> dict:
    """Lee inspeccion.csv (separador ; o ,) si existe. Devuelve {nombre_archivo: fila}."""
    ruta = carpeta / "inspeccion.csv"
    if not ruta.exists():
        return {}
    with open(ruta, encoding="utf-8-sig", newline="") as f:
        muestra = f.read(2048)
        f.seek(0)
        dialecto = csv.Sniffer().sniff(muestra, delimiters=";,\t")
        filas = csv.DictReader(f, dialect=dialecto)
        datos = {}
        for fila in filas:
            fila = {_clave(k): (v or "").strip() for k, v in fila.items() if k}
            if fila.get("archivo"):
                datos[fila["archivo"].lower()] = fila
    print(f"  · Leído {ruta.name}: {len(datos)} registros")
    return datos


def clasificar_imagenes(carpeta: Path, variables_img: set):
    """Devuelve (lista_defectos, imagenes_fijas)."""
    datos_csv = leer_csv(carpeta)
    defectos, fijas = [], {}

    imagenes = sorted(p for p in carpeta.rglob("*")
                      if p.is_file() and p.suffix.lower() in EXTENSIONES)

    for ruta in imagenes:
        nombre_var = f"img_{_clave(ruta.stem).replace(' ', '_')}"
        if nombre_var in variables_img and ruta.parent == carpeta:
            fijas[nombre_var] = ruta
            continue

        rel = ruta.relative_to(carpeta)
        fila = datos_csv.get(ruta.name.lower()) or datos_csv.get(str(rel).lower().replace("\\", "/"))
        partes_nombre = ruta.stem.split("_")

        if fila:
            tipo = fila.get("defecto", "")
            elemento = fila.get("elemento", "")
            ubicacion = fila.get("ubicacion", "")
            severidad = fila.get("severidad", "")
            obs = fila.get("observaciones", "")
        elif len(rel.parts) > 1:  # está dentro de una subcarpeta -> la subcarpeta es el defecto
            tipo = rel.parts[0]
            elemento = " ".join(partes_nombre).replace("-", " ")
            ubicacion = " / ".join(rel.parts[1:-1])
            severidad = obs = ""
        else:  # defecto_elemento_n.jpg
            tipo = partes_nombre[0]
            elemento = " ".join(partes_nombre[1:-1] if len(partes_nombre) > 2 else partes_nombre[1:])
            elemento = elemento.replace("-", " ")
            ubicacion = severidad = obs = ""

        defectos.append({
            "ruta": ruta,
            "archivo": str(rel),
            "tipo": normalizar_defecto(tipo),
            "elemento": elemento or "-",
            "ubicacion": ubicacion or "-",
            "severidad": normalizar_severidad(severidad),
            "observaciones": obs or "-",
        })

    defectos.sort(key=lambda d: (d["tipo"], d["archivo"]))
    for i, d in enumerate(defectos, 1):
        d["num"] = i
    return defectos, fijas


# --------------------------------------------------------------------------- #
# Resumen
# --------------------------------------------------------------------------- #
def construir_resumen(defectos):
    total = len(defectos)
    cuenta = Counter(d["tipo"] for d in defectos)
    resumen = [{"tipo": t, "cantidad": n,
                "porcentaje": f"{100 * n / total:.1f} %" if total else "0 %"}
               for t, n in sorted(cuenta.items(), key=lambda x: (-x[1], x[0]))]

    cuenta_sev = Counter(d["severidad"] for d in defectos)
    orden = SEVERIDADES + sorted(s for s in cuenta_sev if s not in SEVERIDADES)
    resumen_sev = [{"severidad": s, "cantidad": cuenta_sev[s]} for s in orden if cuenta_sev[s]]
    return resumen, resumen_sev


def redactar_conclusiones(resumen, resumen_sev, total):
    if not total:
        return "No se han detectado defectos en la documentación fotográfica analizada."
    partes = [f"{r['cantidad']} de tipo «{r['tipo']}»" for r in resumen]
    texto = (f"Durante la inspección se han registrado un total de {total} "
             f"{'defecto' if total == 1 else 'defectos'}: "
             + ", ".join(partes[:-1]) + (" y " if len(partes) > 1 else "") + partes[-1] + ". ")
    texto += f"El defecto más frecuente es «{resumen[0]['tipo']}» ({resumen[0]['porcentaje']} del total). "
    altas = next((s["cantidad"] for s in resumen_sev if s["severidad"] == "Alta"), 0)
    if altas:
        if altas == 1:
            texto += "Se ha identificado 1 defecto de severidad alta, para el que "
        else:
            texto += f"Se han identificado {altas} defectos de severidad alta, para los que "
        texto += "se recomienda una evaluación detallada y actuación prioritaria."
    return texto


# --------------------------------------------------------------------------- #
# Programa principal
# --------------------------------------------------------------------------- #
def generar_informe(plantilla: Path, carpeta: Path, salida: Path, datos_obra: dict) -> Path:
    tpl = DocxTemplate(plantilla)
    variables = tpl.get_undeclared_template_variables()
    variables_img = {v for v in variables if v.startswith("img_")}

    defectos, fijas = clasificar_imagenes(carpeta, variables_img)
    resumen, resumen_sev = construir_resumen(defectos)
    total = len(defectos)

    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        for d in defectos:
            d["imagen"] = InlineImage(tpl, str(preparar_imagen(d["ruta"], tmpdir)),
                                      width=Mm(ANCHO_FOTO_MM))
            d.pop("ruta")

        contexto = dict(datos_obra)
        contexto.update({
            "defectos": defectos,
            "resumen": resumen,
            "resumen_severidad": resumen_sev,
            "total_defectos": total,
            "conclusiones": datos_obra.get("conclusiones") or redactar_conclusiones(resumen, resumen_sev, total),
        })
        for var in variables_img:
            ruta = fijas.get(var)
            contexto[var] = (InlineImage(tpl, str(preparar_imagen(ruta, tmpdir)), width=Mm(ANCHO_FOTO_MM))
                             if ruta else "")

        tpl.render(contexto)
        salida.parent.mkdir(parents=True, exist_ok=True)
        tpl.save(salida)

    print("\nResumen de defectos")
    print("-" * 45)
    for r in resumen:
        print(f"  {r['tipo']:<32}{r['cantidad']:>5}")
    print("-" * 45)
    print(f"  {'TOTAL':<32}{total:>5}")
    if fijas:
        print(f"\nImágenes fijas colocadas: {', '.join(sorted(fijas))}")
    return salida


def main():
    ap = argparse.ArgumentParser(description="Genera un informe de inspección estructural.")
    ap.add_argument("-p", "--plantilla", help="Plantilla .docx (por defecto la incluida)")
    ap.add_argument("-i", "--imagenes", help="Carpeta con las fotografías")
    ap.add_argument("-o", "--salida", help="Ruta del informe .docx a generar")
    ap.add_argument("-d", "--datos", help="JSON con datos de la obra (obra, ubicacion, cliente, inspector, fecha...)")
    ap.add_argument("--no-preguntar", action="store_true", help="No pedir datos por teclado")
    args = ap.parse_args()

    print("=== Generador de informes de inspección estructural ===\n")

    plantilla = Path(args.plantilla or (PLANTILLA_POR_DEFECTO if args.no_preguntar else
                     preguntar("Ruta de la plantilla del informe", str(PLANTILLA_POR_DEFECTO))))
    if not plantilla.exists():
        if plantilla == PLANTILLA_POR_DEFECTO:
            from crear_plantilla import crear_plantilla
            crear_plantilla(plantilla)
            print(f"  · Plantilla por defecto creada en {plantilla}")
        else:
            sys.exit(f"ERROR: no existe la plantilla {plantilla}")

    carpeta = Path(args.imagenes or preguntar("Ruta de la carpeta con las imágenes")).expanduser()
    while not carpeta.is_dir():
        if args.no_preguntar:
            sys.exit(f"ERROR: la carpeta {carpeta} no existe")
        print(f"  ! La carpeta '{carpeta}' no existe.")
        carpeta = Path(preguntar("Ruta de la carpeta con las imágenes")).expanduser()

    # Datos de la obra: JSON indicado, datos_obra.json en la carpeta o preguntas
    datos = {}
    ruta_json = Path(args.datos) if args.datos else carpeta / "datos_obra.json"
    if ruta_json.exists():
        datos = json.loads(ruta_json.read_text(encoding="utf-8"))
        print(f"  · Datos de la obra leídos de {ruta_json}")
    campos = [("obra", "Obra / estructura", carpeta.name), ("ubicacion", "Ubicación", ""),
              ("cliente", "Cliente", ""), ("inspector", "Inspector", ""),
              ("fecha", "Fecha de inspección", date.today().strftime("%d/%m/%Y"))]
    for clave, texto, defecto in campos:
        if clave not in datos:
            datos[clave] = defecto if args.no_preguntar else preguntar(texto, defecto)

    salida = Path(args.salida) if args.salida else carpeta / f"Informe_inspeccion_{date.today():%Y%m%d}.docx"
    salida = generar_informe(plantilla, carpeta, salida, datos)
    print(f"\nInforme generado: {salida.resolve()}")


if __name__ == "__main__":
    main()
