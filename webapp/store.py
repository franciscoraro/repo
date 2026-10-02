"""Persistencia en disco: plantillas aprobadas e informes en curso (trabajos)."""
import json
import os
import re
import shutil
import threading
import uuid
from datetime import datetime
from pathlib import Path

from . import docx_io

DATA = Path(os.environ.get("INFORMES_DATA_DIR", Path(__file__).resolve().parent.parent / "data"))
_lock = threading.Lock()


def _dir(*p) -> Path:
    d = DATA.joinpath(*p)
    d.mkdir(parents=True, exist_ok=True)
    return d


def _id() -> str:
    return uuid.uuid4().hex[:12]


def _ahora() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _leer(p: Path):
    return json.loads(p.read_text(encoding="utf-8"))


def _escribir(p: Path, obj):
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(obj, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(p)


def _check_id(i: str) -> str:
    if not re.fullmatch(r"[0-9a-f]{12}", i):
        raise KeyError(i)
    return i


# ---------- plantillas ----------
def listar_plantillas() -> list[dict]:
    out = []
    for d in _dir("plantillas").iterdir():
        if (d / "meta.json").exists():
            out.append(_leer(d / "meta.json"))
    return sorted(out, key=lambda m: m["creada"], reverse=True)


def plantilla(tid: str) -> dict:
    d = _dir("plantillas") / _check_id(tid)
    if not (d / "meta.json").exists():
        raise KeyError(tid)
    return {"meta": _leer(d / "meta.json"), "base": (d / "base.docx").read_bytes(),
            "spec": _leer(d / "report.json")}


def borrar_plantilla(tid: str):
    d = _dir("plantillas") / _check_id(tid)
    if not d.exists():
        raise KeyError(tid)
    shutil.rmtree(d)


# ---------- trabajos ----------
def _jd(jid: str) -> Path:
    d = _dir("trabajos") / _check_id(jid)
    if not (d / "job.json").exists():
        raise KeyError(jid)
    return d


def crear_trabajo(*, nombre, instrucciones, base: bytes, estructura: str, plantilla_id: str | None,
                  ejemplo: dict | None) -> dict:
    jid = _id()
    d = _dir("trabajos", jid)
    (d / "base.docx").write_bytes(base)
    (d / "photos").mkdir()
    job = {"id": jid, "nombre": nombre, "instrucciones": instrucciones, "estructura": estructura,
           "plantilla_id": plantilla_id, "ejemplo": ejemplo, "estado": "preparando", "error": None,
           "versiones": [], "aprobado": False, "plantilla_guardada": None, "creado": _ahora()}
    _escribir(d / "job.json", job)
    return job


def trabajo(jid: str) -> dict:
    with _lock:
        return _leer(_jd(jid) / "job.json")


def actualizar_trabajo(jid: str, **campos) -> dict:
    with _lock:
        p = _jd(jid) / "job.json"
        job = _leer(p)
        job.update(campos)
        _escribir(p, job)
        return job


def anadir_version(jid: str, spec: dict, feedback: str | None) -> dict:
    with _lock:
        p = _jd(jid) / "job.json"
        job = _leer(p)
        job["versiones"].append({"n": len(job["versiones"]) + 1, "spec": spec, "feedback": feedback,
                                 "creada": _ahora()})
        job["estado"], job["error"] = "listo", None
        _escribir(p, job)
        return job


def guardar_foto(jid: str, ruta_rel: str, datos: bytes) -> str | None:
    """Guarda una foto normalizada a JPEG y devuelve su identificador (ruta relativa segura)."""
    partes = [re.sub(r"[^\w .()\-]", "_", s).strip(" .") for s in re.split(r"[\\/]+", ruta_rel)]
    partes = [s for s in partes if s and s != ".."]
    if not partes or Path(partes[-1]).suffix.lower() not in docx_io.EXT_IMAGEN:
        return None
    try:
        jpg = docx_io.jpeg_bytes(datos, 2400, 88)
    except Exception:
        return None
    partes[-1] = Path(partes[-1]).stem + ".jpg"
    base = _dir("trabajos", _check_id(jid), "photos")
    rel = Path(*partes)
    n = 1
    while (base / rel).exists():
        n += 1
        rel = Path(*partes[:-1], f"{Path(partes[-1]).stem}_{n}.jpg")
    destino = base / rel
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_bytes(jpg)
    return rel.as_posix()


def fotos(jid: str) -> list[str]:
    base = _jd(jid) / "photos"
    return sorted(p.relative_to(base).as_posix() for p in base.rglob("*.jpg"))


def ruta_foto(jid: str, rel: str) -> Path:
    base = (_jd(jid) / "photos").resolve()
    p = (base / rel).resolve()
    if base not in p.parents or not p.is_file():
        raise KeyError(rel)
    return p


def base_trabajo(jid: str) -> bytes:
    return (_jd(jid) / "base.docx").read_bytes()


def aprobar(jid: str, docx: bytes) -> dict:
    """El informe aprobado pasa a ser plantilla reutilizable."""
    with _lock:
        d = _jd(jid)
        job = _leer(d / "job.json")
        if not job["versiones"]:
            raise ValueError("No hay ninguna versión que aprobar")
        tid = _id()
        t = _dir("plantillas", tid)
        (t / "base.docx").write_bytes((d / "base.docx").read_bytes())
        (t / "informe.docx").write_bytes(docx)
        _escribir(t / "report.json", job["versiones"][-1]["spec"])
        meta = {"id": tid, "nombre": job["nombre"], "instrucciones": job["instrucciones"],
                "estructura": job["estructura"], "creada": _ahora(), "version_aprobada": len(job["versiones"])}
        _escribir(t / "meta.json", meta)
        job["aprobado"], job["plantilla_guardada"] = True, tid
        _escribir(d / "job.json", job)
        return meta
