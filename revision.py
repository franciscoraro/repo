"""
Ciclo de revisión del informe de inspección.

Tras generar la primera versión (v1) se muestra un menú para introducir cambios
y generar nuevas versiones (v2, v3...) hasta marcar una como definitiva.

Los cambios de contenido (defectos, datos de la obra, conclusiones) se guardan en
inspeccion.csv / datos_obra.json de la carpeta de fotos. Los cambios de formato y
de textos fijos se aplican sobre una *plantilla de trabajo* (copia de la plantilla
original). Al cerrar la versión definitiva, esa plantilla de trabajo, que conserva
las etiquetas {{ ... }}, se guarda como nueva plantilla de informe.
"""
import os
import re
import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path

from docx import Document
from docxtpl import DocxTemplate

import generar_informe as gi

ETIQUETA = re.compile(r"\{\{.*?\}\}|\{%.*?%\}")
CAMPOS_DEFECTO = [("tipo", "Tipo de defecto"), ("elemento", "Elemento"), ("ubicacion", "Ubicación"),
                  ("severidad", "Severidad (Alta/Media/Baja)"), ("observaciones", "Observaciones")]
CAMPOS_OBRA = [("obra", "Obra / estructura"), ("ubicacion", "Ubicación"), ("cliente", "Cliente"),
               ("inspector", "Inspector"), ("fecha", "Fecha de inspección")]


# --------------------------------------------------------------------------- #
# Utilidades
# --------------------------------------------------------------------------- #
def leer(mensaje: str) -> str:
    """input() que devuelve 'S' (salir) si se cierra la entrada estándar."""
    try:
        return input(mensaje).strip()
    except EOFError:
        print()
        return "S"


def confirmar(mensaje: str, defecto: bool = True) -> bool:
    opciones = "[S/n]" if defecto else "[s/N]"
    try:
        r = gi._clave(input(f"{mensaje} {opciones}: "))
    except EOFError:
        r = ""
    return defecto if not r else r in ("s", "si", "y", "yes")


def leer_parrafos(mensaje: str) -> str:
    """Lee texto de varias líneas hasta una línea vacía. Cada línea es un párrafo."""
    print(f"{mensaje} (termine con una línea vacía):")
    lineas = []
    while True:
        try:
            linea = input("  > ")
        except EOFError:
            break
        if not linea.strip():
            break
        lineas.append(linea.strip())
    return "\n".join(lineas)


def abrir(ruta: Path):
    """Abre un archivo con la aplicación predeterminada del sistema (Word, LibreOffice...)."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(ruta)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(ruta)])
        else:
            subprocess.Popen(["xdg-open", str(ruta)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print(f"  · Abriendo {ruta}")
    except Exception as e:  # noqa: BLE001
        print(f"  ! No se pudo abrir automáticamente ({e}). Ábralo manualmente: {ruta}")


def _slug(texto: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", gi._clave(texto)).strip("_")[:40] or "informe"


def _parrafos(doc):
    """Todos los párrafos del documento: cuerpo, tablas (anidadas), encabezados y pies."""
    def de_contenedor(cont):
        yield from cont.paragraphs
        for tabla in getattr(cont, "tables", []):
            for fila in tabla.rows:
                for celda in fila.cells:
                    yield from de_contenedor(celda)

    yield from de_contenedor(doc)
    for seccion in doc.sections:
        for parte in (seccion.header, seccion.footer, seccion.first_page_header,
                      seccion.first_page_footer, seccion.even_page_header, seccion.even_page_footer):
            if parte is not None and not parte.is_linked_to_previous:
                yield from de_contenedor(parte)


def reemplazar_texto(ruta_docx: Path, buscar: str, reemplazo: str):
    """Reemplaza texto fijo en un .docx conservando el formato y sin tocar las etiquetas.

    Devuelve (reemplazos_hechos, avisos)."""
    doc = Document(ruta_docx)
    hechos, avisos, vistos = 0, [], set()
    for p in _parrafos(doc):
        if id(p._p) in vistos or buscar not in p.text:
            continue
        vistos.add(id(p._p))
        etiquetas = ETIQUETA.findall(p.text)
        if ETIQUETA.findall(p.text.replace(buscar, reemplazo)) != etiquetas:
            avisos.append(f"Omitido (afectaría a una etiqueta de la plantilla): «{p.text[:60]}»")
            continue

        runs = p.runs
        textos = [r.text for r in runs]
        inicio_run, pos = [], 0
        for t in textos:
            inicio_run.append(pos)
            pos += len(t)
        completo = "".join(textos)

        # Se procesan las apariciones de la última a la primera para no desplazar posiciones
        for m in reversed(list(re.finditer(re.escape(buscar), completo))):
            ini, fin = m.start(), m.end()
            r_ini = max(i for i, s in enumerate(inicio_run) if s <= ini)
            r_fin = max(i for i, s in enumerate(inicio_run) if s < fin)
            implicados = runs[r_ini:r_fin + 1]
            if r_fin > r_ini and any(r._r.xpath(".//w:drawing|.//w:pict") for r in implicados):
                avisos.append(f"Omitido (el texto rodea una imagen): «{p.text[:60]}»")
                continue
            prefijo = textos[r_ini][:ini - inicio_run[r_ini]]
            sufijo = textos[r_fin][fin - inicio_run[r_fin]:]
            if r_ini == r_fin:
                textos[r_ini] = prefijo + reemplazo + sufijo
            else:
                textos[r_ini] = prefijo + reemplazo
                for k in range(r_ini + 1, r_fin):
                    textos[k] = ""
                textos[r_fin] = sufijo
            hechos += 1
        for r, t in zip(runs, textos):
            if r.text != t:
                r.text = t
    if hechos:
        doc.save(ruta_docx)
    return hechos, avisos


def validar_plantilla(ruta: Path):
    """Comprueba que la plantilla es legible y devuelve sus variables (lanza excepción si no)."""
    return DocxTemplate(ruta).get_undeclared_template_variables()


# --------------------------------------------------------------------------- #
# Ciclo de revisión
# --------------------------------------------------------------------------- #
class CicloRevision:
    def __init__(self, plantilla: Path, carpeta: Path, salida: Path, datos: dict):
        self.plantilla_origen = Path(plantilla)
        self.carpeta = Path(carpeta)
        self.datos = datos
        self.dir_salida = Path(salida).parent
        self.base = Path(salida).stem
        self.dir_salida.mkdir(parents=True, exist_ok=True)
        self.plantilla = self.dir_salida / f"{self.base}_plantilla_trabajo.docx"
        if not self.plantilla.exists():
            # Revisión empezada otro día (el nombre por defecto lleva la fecha): se retoma
            previas = sorted(self.dir_salida.glob("*_plantilla_trabajo.docx"),
                             key=lambda p: p.stat().st_mtime)
            if previas:
                self.plantilla = previas[-1]
                self.base = self.plantilla.name[:-len("_plantilla_trabajo.docx")]
        self.version = self._ultima_version()
        self.ultimo_informe = None
        self.pendiente = True
        self.cambios_plantilla = False

        if self.plantilla.exists() and confirmar(
                f"\nExiste una revisión anterior (v{self.version}). ¿Continuar con su plantilla de trabajo?"):
            print(f"  · Se continúa con {self.plantilla.name}")
            self.cambios_plantilla = True
        else:
            shutil.copy2(self.plantilla_origen, self.plantilla)

        self.defectos, self.fijas = gi.clasificar_imagenes(self.carpeta, gi.variables_imagen(self.plantilla))

    def _ultima_version(self) -> int:
        patron = re.compile(re.escape(self.base) + r"_v(\d+)\.docx$")
        nums = [int(m.group(1)) for p in self.dir_salida.glob(f"{self.base}_v*.docx")
                if (m := patron.match(p.name))]
        return max(nums, default=0)

    # ------------------------------------------------------------------ #
    def ejecutar(self):
        self.generar_version()
        acciones = {
            "1": self.ver_defectos, "2": self.modificar_defecto, "3": self.excluir_defectos,
            "4": self.restaurar_defectos, "5": self.renombrar_tipo, "6": self.modificar_obra,
            "7": self.editar_conclusiones, "8": self.reemplazar_texto_fijo,
            "9": self.editar_plantilla_word, "A": lambda: abrir(self.ultimo_informe),
            "G": self.generar_version,
        }
        while True:
            self.menu()
            op = leer("Opción: ").upper()
            if op == "D":
                if self.finalizar():
                    return
            elif op == "S":
                if self.salir():
                    return
            elif op in acciones:
                try:
                    acciones[op]()
                except (ValueError, IndexError):
                    print("  ! Entrada no válida.")
            else:
                print("  ! Opción no reconocida.")

    def menu(self):
        estado = "  (hay cambios sin generar)" if self.pendiente else ""
        print(f"\n===== Revisión del informe – versión actual: v{self.version}{estado} =====")
        print("  1. Ver lista de defectos")
        print("  2. Modificar un defecto (tipo, elemento, ubicación, severidad, observaciones)")
        print("  3. Excluir defectos del informe")
        print("  4. Restaurar defectos excluidos")
        print("  5. Renombrar / fusionar un tipo de defecto")
        print("  6. Modificar datos de la obra")
        print("  7. Redactar conclusiones")
        print("  8. Reemplazar un texto fijo de la plantilla")
        print("  9. Editar formato de la plantilla en Word (logos, estilos, textos...)")
        print("  A. Abrir la última versión generada")
        print("  G. Generar nueva versión con los cambios")
        print("  D. Marcar como DEFINITIVA y guardar como nueva plantilla")
        print("  S. Salir (se puede continuar la revisión más tarde)")

    def _cambio(self, mensaje="Cambio registrado"):
        self.pendiente = True
        gi.numerar(self.defectos)
        print(f"  ✓ {mensaje}. Pulse G para generar la nueva versión.")

    # ------------------------------------------------------------------ #
    def generar_version(self):
        self.version += 1
        destino = self.dir_salida / f"{self.base}_v{self.version}.docx"
        try:
            gi.renderizar(self.plantilla, self.defectos, self.fijas, self.datos, destino)
        except PermissionError:
            self.version -= 1
            print(f"  ! No se puede escribir {destino.name}: ¿está abierto en Word? Ciérrelo y reintente.")
            return
        gi.guardar_estado(self.carpeta, self.defectos, self.datos)
        self.ultimo_informe = destino
        self.pendiente = False
        gi.mostrar_resumen(self.defectos)
        print(f"\n  ✓ Versión v{self.version} generada: {destino.resolve()}")

    # ------------------------------------------------------------------ #
    def ver_defectos(self):
        print(f"\n  {'N.º':>4}  {'Tipo':<28}{'Elemento':<20}{'Ubicación':<16}{'Sev.':<9}Archivo")
        for d in gi.incluidos(self.defectos):
            print(f"  {d['num']:>4}  {d['tipo'][:27]:<28}{d['elemento'][:19]:<20}"
                  f"{d['ubicacion'][:15]:<16}{d['severidad'][:8]:<9}{d['archivo']}")
        excl = self._excluidos()
        if excl:
            print("\n  Excluidos del informe:")
            for i, d in enumerate(excl, 1):
                print(f"   x{i}  {d['tipo']} – {d['archivo']}")

    def _excluidos(self):
        return [d for d in self.defectos if d["excluido"]]

    def _pedir_defecto(self):
        activos = gi.incluidos(self.defectos)
        n = int(leer(f"N.º de defecto (1-{len(activos)}): "))
        if not 1 <= n <= len(activos):
            raise ValueError
        return activos[n - 1]

    def modificar_defecto(self):
        d = self._pedir_defecto()
        print(f"  Archivo: {d['archivo']}   (Enter = mantener, '-' = dejar vacío)")
        for clave, texto in CAMPOS_DEFECTO:
            valor = gi.preguntar(f"  {texto}", d[clave])
            valor = "" if valor == "-" else valor
            if clave == "tipo":
                valor = gi.normalizar_defecto(valor)
            elif clave == "severidad":
                valor = gi.normalizar_severidad(valor)
            d[clave] = valor
        self._cambio(f"Defecto «{d['archivo']}» modificado")

    def excluir_defectos(self):
        activos = gi.incluidos(self.defectos)
        nums = [int(x) for x in re.split(r"[,\s]+", leer("N.º de defectos a excluir (p. ej. 3, 7): ")) if x]
        if not nums or any(not 1 <= n <= len(activos) for n in nums):
            raise ValueError
        for n in nums:
            activos[n - 1]["excluido"] = True
        self._cambio(f"{len(nums)} defecto(s) excluido(s); se renumerarán los restantes")

    def restaurar_defectos(self):
        excl = self._excluidos()
        if not excl:
            print("  No hay defectos excluidos.")
            return
        for i, d in enumerate(excl, 1):
            print(f"   x{i}  {d['tipo']} – {d['archivo']}")
        sel = leer("Restaurar (p. ej. x1, x3 o 'todos'): ").lower()
        elegidos = excl if sel == "todos" else [excl[int(x.lstrip("x")) - 1]
                                                for x in re.split(r"[,\s]+", sel) if x]
        for d in elegidos:
            d["excluido"] = False
        self._cambio(f"{len(elegidos)} defecto(s) restaurado(s)")

    def renombrar_tipo(self):
        resumen, _ = gi.construir_resumen(gi.incluidos(self.defectos))
        for i, r in enumerate(resumen, 1):
            print(f"   {i}. {r['tipo']} ({r['cantidad']})")
        actual = resumen[int(leer("Tipo a renombrar (n.º): ")) - 1]["tipo"]
        nuevo = gi.normalizar_defecto(gi.preguntar("  Nuevo nombre (si ya existe, se fusionan)", actual))
        cambiados = 0
        for d in self.defectos:
            if d["tipo"] == actual:
                d["tipo"] = nuevo
                cambiados += 1
        self._cambio(f"«{actual}» → «{nuevo}» en {cambiados} defecto(s)")

    def modificar_obra(self):
        print("  (Enter = mantener)")
        claves = [c for c, _ in CAMPOS_OBRA] + [k for k in self.datos
                                                 if k not in dict(CAMPOS_OBRA) and k != "conclusiones"]
        textos = dict(CAMPOS_OBRA)
        for clave in claves:
            self.datos[clave] = gi.preguntar(f"  {textos.get(clave, clave)}", str(self.datos.get(clave, "")))
        self._cambio("Datos de la obra actualizados")

    def editar_conclusiones(self):
        activos = gi.incluidos(self.defectos)
        resumen, resumen_sev = gi.construir_resumen(activos)
        manual = self.datos.get("conclusiones")
        print("\n  Conclusiones actuales" + (" (redactadas manualmente):" if manual else " (automáticas):"))
        print("  " + (manual or gi.redactar_conclusiones(resumen, resumen_sev, len(activos))).replace("\n", "\n  "))
        print("\n  1. Escribir nuevo texto   2. Añadir texto a las actuales   3. Volver a automáticas   0. Cancelar")
        op = leer("Opción: ")
        if op == "1":
            texto = leer_parrafos("Nuevo texto")
            if texto:
                self.datos["conclusiones"] = texto
                self._cambio("Conclusiones actualizadas")
        elif op == "2":
            texto = leer_parrafos("Texto a añadir")
            if texto:
                base = manual or gi.redactar_conclusiones(resumen, resumen_sev, len(activos))
                self.datos["conclusiones"] = base + "\n" + texto
                self._cambio("Conclusiones ampliadas")
        elif op == "3":
            self.datos.pop("conclusiones", None)
            self._cambio("Se usarán conclusiones automáticas")

    def reemplazar_texto_fijo(self):
        print("  Cambia textos fijos de la plantilla (títulos, párrafos de alcance, pies...).")
        print("  No afecta a los datos variables ni a las etiquetas {{ ... }}.")
        buscar = leer("Texto a buscar: ")
        if not buscar or buscar == "S":
            return
        reemplazo = leer("Reemplazar por: ")
        n, avisos = reemplazar_texto(self.plantilla, buscar, reemplazo)
        for a in avisos:
            print(f"  ! {a}")
        if n:
            self.cambios_plantilla = True
            self._cambio(f"{n} reemplazo(s) en la plantilla")
        elif not avisos:
            print("  No se encontró el texto en la plantilla (los datos variables se cambian con las opciones 2, 6 y 7).")

    def editar_plantilla_word(self):
        copia = self.plantilla.with_name(self.plantilla.stem + "_respaldo.docx")
        shutil.copy2(self.plantilla, copia)
        print("\n  Se abrirá la PLANTILLA DE TRABAJO. Modifique formato, logos, estilos o textos fijos,")
        print("  respetando las etiquetas {{ ... }}, {%p ... %} y {%tr ... %}.")
        abrir(self.plantilla)
        leer("  Guarde y CIERRE el documento en Word y pulse Enter para continuar... ")
        try:
            variables = validar_plantilla(self.plantilla)
        except Exception as e:  # noqa: BLE001
            print(f"  ! La plantilla modificada no es válida ({e}).")
            print("    Se restaura la versión anterior.")
            shutil.copy2(copia, self.plantilla)
            copia.unlink()
            return
        copia.unlink()
        if "defectos" not in variables:
            print("  ! Aviso: la plantilla ya no contiene el bloque {%p for d in defectos %}.")
        nuevas = gi.variables_imagen(self.plantilla) - set(self.fijas)
        if nuevas:
            self.defectos, self.fijas = self._reclasificar()
        self.cambios_plantilla = True
        self._cambio("Plantilla de trabajo actualizada")

    def _reclasificar(self):
        """Tras añadir etiquetas img_xxx, vuelve a separar imágenes fijas conservando las ediciones."""
        previos = {d["archivo"]: d for d in self.defectos}
        nuevos, fijas = gi.clasificar_imagenes(self.carpeta, gi.variables_imagen(self.plantilla))
        return gi.numerar([previos.get(d["archivo"], d) for d in nuevos]), fijas

    # ------------------------------------------------------------------ #
    def finalizar(self) -> bool:
        if self.pendiente:
            print("  Hay cambios sin generar; se genera primero una nueva versión.")
            self.generar_version()
            if self.pendiente:
                return False
            if not confirmar(f"¿Revisada la v{self.version} y conforme para marcarla como definitiva?"):
                return False

        definitivo = self.dir_salida / f"{self.base}_DEFINITIVO.docx"
        shutil.copy2(self.ultimo_informe, definitivo)
        print(f"\n  ✓ Informe definitivo: {definitivo.resolve()}  (= v{self.version})")

        sugerida = gi.BASE / "plantillas" / f"plantilla_{_slug(self.datos.get('obra', ''))}_{date.today():%Y%m%d}.docx"
        destino = Path(gi.preguntar("Guardar la nueva plantilla de informe en", str(sugerida))).expanduser()
        if destino.suffix.lower() != ".docx":
            destino = destino.with_suffix(".docx")
        destino.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(self.plantilla, destino)
        print(f"  ✓ Nueva plantilla guardada: {destino.resolve()}")
        if not self.cambios_plantilla:
            print("    (No hubo cambios de formato ni de textos fijos: es igual a la plantilla de partida.)")

        if confirmar("¿Usar esta plantilla por defecto en los próximos informes?"):
            config = gi.leer_config()
            try:
                config["plantilla_por_defecto"] = str(destino.resolve().relative_to(gi.BASE.resolve()))
            except ValueError:
                config["plantilla_por_defecto"] = str(destino.resolve())
            gi.guardar_config(config)
            print(f"  ✓ Configurada como plantilla por defecto ({gi.CONFIG.name}).")

        self.plantilla.unlink(missing_ok=True)
        return True

    def salir(self) -> bool:
        if self.pendiente and not confirmar("Hay cambios sin generar. ¿Salir de todos modos?", defecto=False):
            return False
        gi.guardar_estado(self.carpeta, self.defectos, self.datos)
        print(f"\n  Revisión guardada (última versión: v{self.version}).")
        print("  Vuelva a ejecutar la rutina con la misma carpeta para continuar.")
        return True
