"""Llamada a Claude: genera o corrige la spec del informe."""
import base64
import json
import os

import anthropic

MODELO = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5-5")
MAX_FOTOS_VISTAS = int(os.environ.get("MAX_FOTOS_VISTAS", "60"))

SISTEMA = """Eres un técnico experto en redactar informes de inspección en español.
Recibes: (a) la estructura de un documento que sirve de PLANTILLA, (b) lo que el usuario quiere que haga el informe,
(c) las fotografías a incorporar (cada una precedida por su identificador) y, según el caso, (d) un informe anterior
aprobado como modelo, o (e) la versión previa del informe con las correcciones pedidas por el usuario.

Reglas:
- Respeta la estructura, secciones, orden, tono y terminología de la plantilla o del informe modelo.
- Describe únicamente lo que se ve en las fotos o lo que dice el usuario. No inventes medidas, fechas, nombres ni
  normativa; si falta un dato que la plantilla requiere, escribe «[a completar]».
- Cada fotografía debe aparecer en un bloque "photo" con su identificador exacto y un pie descriptivo.
- Si hay una versión previa y correcciones, aplica exactamente las correcciones y conserva el resto.
- Devuelve el informe llamando a la herramienta emitir_informe."""

HERRAMIENTA = {
    "name": "emitir_informe",
    "description": "Entrega el informe completo como una lista ordenada de bloques.",
    "input_schema": {
        "type": "object",
        "properties": {
            "title": {"type": "string"},
            "blocks": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "type": {"type": "string",
                                 "enum": ["heading", "paragraph", "bullets", "table", "photo", "page_break"]},
                        "level": {"type": "integer", "description": "heading: 1, 2 o 3"},
                        "text": {"type": "string", "description": "heading o paragraph"},
                        "items": {"type": "array", "items": {"type": "string"}, "description": "bullets"},
                        "header": {"type": "array", "items": {"type": "string"}, "description": "table"},
                        "rows": {"type": "array", "items": {"type": "array", "items": {"type": "string"}},
                                 "description": "table"},
                        "file": {"type": "string", "description": "photo: identificador exacto de la foto"},
                        "caption": {"type": "string", "description": "photo: pie de foto"},
                    },
                    "required": ["type"],
                },
            },
        },
        "required": ["title", "blocks"],
    },
}


def construir_mensaje(*, nombre, instrucciones, estructura, fotos, ejemplo=None, previo=None, feedback=None):
    """fotos: lista de (id, jpeg_bytes | None). Devuelve el contenido del mensaje de usuario."""
    partes = [{"type": "text", "text": f"NOMBRE DEL INFORME: {nombre}\n\nQUÉ DEBE HACER EL INFORME:\n{instrucciones}"}]
    if estructura:
        partes.append({"type": "text", "text": f"ESTRUCTURA DE LA PLANTILLA:\n{estructura}"})
    if ejemplo:
        partes.append({"type": "text", "text": "INFORME MODELO YA APROBADO (sigue su estilo y estructura; "
                       "no copies sus datos concretos):\n" + json.dumps(ejemplo, ensure_ascii=False)})
    if previo:
        partes.append({"type": "text", "text": "VERSIÓN PREVIA DEL INFORME:\n" + json.dumps(previo, ensure_ascii=False)})
    if feedback:
        partes.append({"type": "text", "text": f"CORRECCIONES DEL USUARIO A APLICAR:\n{feedback}"})
    partes.append({"type": "text", "text": f"FOTOGRAFÍAS ({len(fotos)}):"})
    for fid, jpg in fotos:
        partes.append({"type": "text", "text": f"Foto: {fid}" + ("" if jpg else " (no visible, solo su nombre)")})
        if jpg:
            partes.append({"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                                       "data": base64.b64encode(jpg).decode()}})
    return partes


def generar_spec(**kw) -> dict:
    contenido = construir_mensaje(**kw)
    cliente = anthropic.Anthropic()
    with cliente.messages.stream(
        model=MODELO, max_tokens=16000, system=SISTEMA, tools=[HERRAMIENTA],
        tool_choice={"type": "tool", "name": "emitir_informe"},
        messages=[{"role": "user", "content": contenido}],
    ) as s:
        msg = s.get_final_message()
    for b in msg.content:
        if b.type == "tool_use":
            return b.input
    raise RuntimeError("Claude no devolvió el informe")
