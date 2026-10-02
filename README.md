# Generador de informes de inspección estructural

Rutina en Python que:

1. Carga una **plantilla Word** (`.docx`) con el formato del informe.
2. Pide la **ruta de la carpeta con las fotografías**.
3. Inserta cada fotografía en su espacio de la plantilla, con su ficha de defecto.
4. Añade un **resumen final** con el número total de cada tipo de defecto (y por severidad) y unas conclusiones automáticas.
5. Abre un **ciclo de revisión**: se introducen cambios, se generan nuevas versiones (v1, v2, v3…) hasta marcar
   una como **definitiva**, cuyo formato se guarda como **nueva plantilla** para próximos informes.

## Instalación

```bash
pip install -r requirements.txt
```

## Uso

Modo interactivo (pregunta plantilla, carpeta de imágenes y datos de la obra):

```bash
python generar_informe.py
```

Modo directo:

```bash
python generar_informe.py -i "C:\Inspecciones\Puente\fotos" -p plantillas/plantilla_inspeccion.docx -o informe.docx
```

| Opción | Descripción |
|---|---|
| `-p`, `--plantilla` | Plantilla `.docx` (por defecto `plantillas/plantilla_inspeccion.docx`) |
| `-i`, `--imagenes` | Carpeta con las fotografías |
| `-o`, `--salida` | Informe a generar (por defecto `Informe_inspeccion_AAAAMMDD.docx` en la carpeta de fotos) |
| `-d`, `--datos` | JSON con los datos de la obra |
| `--no-preguntar` | No pedir nada por teclado ni abrir el ciclo de revisión |
| `--sin-revision` | Generar una única versión, sin ciclo de revisión |

## Ciclo de revisión

Tras generar la v1 aparece el menú:

```
===== Revisión del informe – versión actual: v1 =====
  1. Ver lista de defectos
  2. Modificar un defecto (tipo, elemento, ubicación, severidad, observaciones)
  3. Excluir defectos del informe
  4. Restaurar defectos excluidos
  5. Renombrar / fusionar un tipo de defecto
  6. Modificar datos de la obra
  7. Redactar conclusiones
  8. Reemplazar un texto fijo de la plantilla
  9. Editar formato de la plantilla en Word (logos, estilos, textos...)
  A. Abrir la última versión generada
  G. Generar nueva versión con los cambios
  D. Marcar como DEFINITIVA y guardar como nueva plantilla
  S. Salir (se puede continuar la revisión más tarde)
```

- Los cambios se acumulan y se aplican al pulsar **G**, que crea `Informe_..._v2.docx`, `_v3.docx`…
  Las versiones anteriores se conservan.
- **Cambios de contenido** (opciones 2 a 7): se guardan en `inspeccion.csv` y `datos_obra.json` de la carpeta
  de fotos. Si existía un `inspeccion.csv` propio, antes se copia como `inspeccion_original.csv`.
  Los defectos excluidos quedan con `incluir = no`.
- **Cambios de formato** (opciones 8 y 9): se hacen sobre una *plantilla de trabajo*
  (`Informe_..._plantilla_trabajo.docx`), nunca sobre la plantilla original. La opción 9 la abre en Word.
  Al volver se comprueba que sigue siendo válida y, si no lo es, se restaura. La opción 8 se niega a tocar
  las etiquetas `{{ ... }}`.
- **S** guarda el estado. Si vuelve a ejecutar la rutina con la misma carpeta, ofrece continuar la revisión
  donde se dejó.
- **D** genera la última versión si hay cambios pendientes y la copia como `Informe_..._DEFINITIVO.docx`.
  Después guarda la plantilla de trabajo como nueva plantilla (por defecto
  `plantillas/plantilla_<obra>_<fecha>.docx`) y pregunta si debe usarse **por defecto** en los próximos
  informes (se anota en `config.json`).

> La nueva plantilla conserva todo el formato y los textos fijos del informe definitivo, pero mantiene las
> etiquetas `{{ ... }}`, así que en cada informe nuevo se rellena con las fotos y los datos de esa inspección.

## Cómo se clasifica cada imagen

Por orden de prioridad:

**a) Archivo `inspeccion.csv`** en la carpeta de fotos (separador `;` o `,`), la opción más completa:

```csv
archivo;defecto;elemento;ubicacion;severidad;observaciones
IMG_0012.jpg;Fisura;Viga V-3;Planta 1;Alta;Fisura de cortante, abertura 0,4 mm
IMG_0013.jpg;Corrosión;Pilar P-2;Sótano;Media;
```

**b) Subcarpetas por tipo de defecto:**

```
fotos/
├── Fisura/       viga_V3.jpg, pilar_P2.jpg
├── Corrosion/    pilar_P1.jpg
└── Humedad/      Sotano/muro_M1.jpg     (subcarpetas intermedias = ubicación)
```

**c) Nombre del archivo** `defecto_elemento_n.jpg`, p. ej. `fisura_viga-P2_01.jpg`.

Los nombres de defecto se normalizan (mayúsculas, tildes, plurales y sinónimos) mediante
`CATALOGO_DEFECTOS` en `generar_informe.py`, de modo que `fisuras`, `Fisura` y `FISURA`
cuentan como el mismo tipo. Amplíe ese diccionario con su propia nomenclatura.
Severidades reconocidas: Alta / Media / Baja (también `grave`, `moderada`, `leve`, `A/M/B`).

### Imágenes fijas

Las imágenes en la raíz de la carpeta cuyo nombre coincide con una etiqueta `{{ img_<nombre> }}`
de la plantilla se colocan en ese lugar y **no** cuentan como defecto:
`portada.jpg` → `{{ img_portada }}`, `ubicacion.png` → `{{ img_ubicacion }}`.
Puede añadir las etiquetas que quiera a su plantilla (`{{ img_plano_planta }}` → `plano_planta.jpg`).

### Datos de la obra

Si existe `datos_obra.json` en la carpeta de fotos se usa automáticamente; si no, se preguntan:

```json
{ "obra": "Puente sobre el río Seco", "ubicacion": "PK 12+300", "cliente": "Ayuntamiento",
  "inspector": "Ing. F. Raro", "fecha": "02/10/2026" }
```

Cualquier clave adicional del JSON queda disponible en la plantilla como `{{ clave }}`.
Si incluye `"conclusiones"`, sustituye al texto automático.

## Usar su propio formato de informe

`python crear_plantilla.py` genera la plantilla por defecto. Puede abrirla en Word y cambiar
logos, estilos, encabezados, textos, etc., o crear la suya desde cero usando estas etiquetas:

| Etiqueta | Contenido |
|---|---|
| `{{ obra }}`, `{{ ubicacion }}`, `{{ cliente }}`, `{{ inspector }}`, `{{ fecha }}` | Datos generales |
| `{%p for d in defectos %}` … `{%p endfor %}` | Bloque que se repite por cada foto (cada etiqueta en su propio párrafo) |
| `{{ d.num }}`, `{{ d.tipo }}`, `{{ d.elemento }}`, `{{ d.ubicacion }}`, `{{ d.severidad }}`, `{{ d.observaciones }}`, `{{ d.archivo }}` | Ficha del defecto |
| `{{ d.imagen }}` | Fotografía del defecto |
| `{%tr for r in resumen %}` … `{%tr endfor %}` con `{{ r.tipo }}`, `{{ r.cantidad }}`, `{{ r.porcentaje }}` | Filas de la tabla resumen (cada `{%tr %}` en su propia fila) |
| `{%tr for s in resumen_severidad %}` con `{{ s.severidad }}`, `{{ s.cantidad }}` | Resumen por severidad |
| `{{ total_defectos }}` | Número total de defectos |
| `{{ conclusiones }}` | Conclusiones automáticas |
| `{{ img_xxx }}` | Imagen fija `xxx.jpg` |

Las fotos se corrigen de orientación (EXIF) y se reducen a 1600 px para que el informe no pese
demasiado; el ancho en el documento se ajusta con `ANCHO_FOTO_MM` (150 mm por defecto).
