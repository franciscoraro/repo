"""Estructura del informe (spec) que intercambian Claude, la vista previa y el generador de Word."""
from typing import Any

TIPOS = {"heading", "paragraph", "bullets", "table", "photo", "page_break"}


def _txt(v: Any) -> str:
    return "" if v is None else str(v).strip()


def normalizar(spec: dict, fotos_validas: list[str]) -> dict:
    """Limpia lo que devuelve el modelo y garantiza que TODAS las fotos aparecen en el informe.

    - descarta bloques desconocidos o vacíos
    - elimina referencias a fotos inexistentes y fotos repetidas
    - añade al final un anexo con las fotos que el modelo no haya colocado
    """
    validas = set(fotos_validas)
    usadas: set[str] = set()
    bloques: list[dict] = []
    for b in (spec or {}).get("blocks") or []:
        if not isinstance(b, dict) or b.get("type") not in TIPOS:
            continue
        t = b["type"]
        if t == "heading":
            texto = _txt(b.get("text"))
            if texto:
                nivel = b.get("level") if b.get("level") in (1, 2, 3) else 1
                bloques.append({"type": t, "level": nivel, "text": texto})
        elif t == "paragraph":
            texto = _txt(b.get("text"))
            if texto:
                bloques.append({"type": t, "text": texto})
        elif t == "bullets":
            items = [_txt(i) for i in b.get("items") or [] if _txt(i)]
            if items:
                bloques.append({"type": t, "items": items})
        elif t == "table":
            header = [_txt(c) for c in b.get("header") or []]
            filas = [[_txt(c) for c in (r if isinstance(r, list) else [r])] for r in b.get("rows") or []]
            if header or filas:
                n = max([len(header)] + [len(r) for r in filas])
                header += [""] * (n - len(header)) if header else []
                filas = [r + [""] * (n - len(r)) for r in filas]
                bloques.append({"type": t, "header": header, "rows": filas})
        elif t == "photo":
            f = _txt(b.get("file"))
            if f in validas and f not in usadas:
                usadas.add(f)
                bloques.append({"type": t, "file": f, "caption": _txt(b.get("caption"))})
        else:
            bloques.append({"type": t})
    faltan = [f for f in fotos_validas if f not in usadas]
    if faltan:
        bloques.append({"type": "heading", "level": 1, "text": "Anexo fotográfico"})
        for f in faltan:
            bloques.append({"type": "photo", "file": f, "caption": f})
    return {"title": _txt((spec or {}).get("title")) or "Informe", "blocks": bloques}
