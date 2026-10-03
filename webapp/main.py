"""App web para crear informes de inspección con Claude.  Arranque: uvicorn webapp.main:app"""
import base64
import logging
import os
import secrets
import threading
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, Response, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from . import docx_io, llm, spec as specmod, store

log = logging.getLogger("informes")
STATIC = Path(__file__).parent / "static"
PASSWORD = os.environ.get("APP_PASSWORD")  # si se define, se pide usuario cualquiera + esta contraseña

app = FastAPI(title="Informes de inspección")


@app.middleware("http")
async def autenticar(request: Request, call_next):
    if PASSWORD:
        ok = False
        h = request.headers.get("authorization", "")
        if h.startswith("Basic "):
            try:
                _, _, pw = base64.b64decode(h[6:]).decode().partition(":")
                ok = secrets.compare_digest(pw, PASSWORD)
            except Exception:
                pass
        if not ok:
            return Response(status_code=401, headers={"WWW-Authenticate": 'Basic realm="Informes"'})
    return await call_next(request)


def _404(e: KeyError):
    raise HTTPException(404, "No encontrado") from e


def _publico(job: dict) -> dict:
    return {k: v for k, v in job.items() if k not in ("estructura", "ejemplo")} | {"fotos": store.fotos(job["id"])}


# ---------- generación en segundo plano ----------
def _generar(jid: str, feedback: str | None = None):
    try:
        job = store.trabajo(jid)
        ids = store.fotos(jid)
        fotos = []
        for i, fid in enumerate(ids):
            jpg = docx_io.jpeg_bytes(store.ruta_foto(jid, fid), docx_io.LADO_LLM_PX, 80) \
                if i < llm.MAX_FOTOS_VISTAS else None
            fotos.append((fid, jpg))
        previo = job["versiones"][-1]["spec"] if job["versiones"] else None
        crudo = llm.generar_spec(nombre=job["nombre"], instrucciones=job["instrucciones"],
                                 estructura=job["estructura"], fotos=fotos, ejemplo=job["ejemplo"],
                                 previo=previo, feedback=feedback)
        store.anadir_version(jid, specmod.normalizar(crudo, ids), feedback)
    except Exception as e:  # el error se muestra al usuario en la interfaz
        log.exception("Fallo generando informe %s", jid)
        store.actualizar_trabajo(jid, estado="error", error=str(e))


def _lanzar(jid: str, feedback: str | None = None):
    store.actualizar_trabajo(jid, estado="generando", error=None)
    threading.Thread(target=_generar, args=(jid, feedback), daemon=True).start()


# ---------- API ----------
@app.get("/api/plantillas")
def plantillas():
    return store.listar_plantillas()


@app.delete("/api/plantillas/{tid}")
def borrar_plantilla(tid: str):
    try:
        store.borrar_plantilla(tid)
    except KeyError as e:
        _404(e)
    return {"ok": True}


@app.post("/api/trabajos")
async def crear_trabajo(
    nombre: str = Form(...),
    instrucciones: str = Form(...),
    plantilla_id: str = Form(""),
    plantilla_archivo: UploadFile | None = File(None),
    rutas: list[str] = Form([]),
    fotos: list[UploadFile] = File([]),
):
    nombre, instrucciones = nombre.strip(), instrucciones.strip()
    if not nombre or not instrucciones:
        raise HTTPException(400, "Falta el nombre o lo que debe hacer el informe")
    ejemplo = None
    if plantilla_id:
        try:
            t = store.plantilla(plantilla_id)
        except KeyError as e:
            _404(e)
        base, estructura, ejemplo = t["base"], t["meta"]["estructura"], t["spec"]
    elif plantilla_archivo and plantilla_archivo.filename:
        try:
            base, estructura = docx_io.cargar_plantilla(plantilla_archivo.filename, await plantilla_archivo.read())
        except Exception as e:
            raise HTTPException(400, f"Plantilla no válida: {e}") from e
    else:
        raise HTTPException(400, "Indique un archivo de plantilla o elija un informe existente")

    job = store.crear_trabajo(nombre=nombre, instrucciones=instrucciones, base=base, estructura=estructura,
                              plantilla_id=plantilla_id or None, ejemplo=ejemplo)
    n = 0
    for i, f in enumerate(fotos):
        ruta = rutas[i] if i < len(rutas) else f.filename or f"foto{i}.jpg"
        if store.guardar_foto(job["id"], ruta, await f.read()):
            n += 1
    if n == 0:
        store.actualizar_trabajo(job["id"], estado="error", error="No se recibió ninguna fotografía válida")
        raise HTTPException(400, "No se recibió ninguna fotografía válida")
    _lanzar(job["id"])
    return _publico(store.trabajo(job["id"]))


@app.get("/api/trabajos/{jid}")
def ver_trabajo(jid: str):
    try:
        return _publico(store.trabajo(jid))
    except KeyError as e:
        _404(e)


@app.post("/api/trabajos/{jid}/modificar")
def modificar(jid: str, cambios: str = Form(...)):
    try:
        job = store.trabajo(jid)
    except KeyError as e:
        _404(e)
    if job["estado"] == "generando":
        raise HTTPException(409, "Ya se está generando una versión")
    if job["aprobado"]:
        raise HTTPException(409, "El informe ya está aprobado")
    if not job["versiones"]:
        raise HTTPException(409, "Todavía no hay versión que modificar")
    if not cambios.strip():
        raise HTTPException(400, "Indique qué desea modificar")
    _lanzar(jid, cambios.strip())
    return _publico(store.trabajo(jid))


@app.post("/api/trabajos/{jid}/reintentar")
def reintentar(jid: str):
    try:
        job = store.trabajo(jid)
    except KeyError as e:
        _404(e)
    if job["estado"] != "error":
        raise HTTPException(409, "No hay error que reintentar")
    _lanzar(jid)
    return _publico(store.trabajo(jid))


def _docx(jid: str, n: int | None = None) -> bytes:
    job = store.trabajo(jid)
    if not job["versiones"]:
        raise HTTPException(409, "Todavía no hay versión")
    v = job["versiones"][-1] if n is None else next((v for v in job["versiones"] if v["n"] == n), None)
    if v is None:
        raise HTTPException(404, "Versión inexistente")
    return docx_io.renderizar(store.base_trabajo(jid), v["spec"], lambda f: store.ruta_foto(jid, f))


@app.post("/api/trabajos/{jid}/aprobar")
def aprobar(jid: str):
    try:
        job = store.trabajo(jid)
        if job["aprobado"]:
            raise HTTPException(409, "El informe ya está aprobado")
        if job["estado"] != "listo":
            raise HTTPException(409, "Espere a que termine la generación")
        return store.aprobar(jid, _docx(jid))
    except KeyError as e:
        _404(e)


@app.get("/api/trabajos/{jid}/descargar")
def descargar(jid: str, v: int | None = None):
    try:
        datos = _docx(jid, v)
        nombre = "".join(c if c.isalnum() or c in " -_" else "_" for c in store.trabajo(jid)["nombre"]).strip() or "informe"
    except KeyError as e:
        _404(e)
    return Response(datos, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    headers={"Content-Disposition": f'attachment; filename="{nombre}.docx"'})


@app.get("/api/trabajos/{jid}/fotos/{ruta:path}")
def foto(jid: str, ruta: str):
    try:
        return FileResponse(store.ruta_foto(jid, ruta), media_type="image/jpeg")
    except KeyError as e:
        _404(e)


@app.get("/api/estado")
def estado():
    return {"api_key": bool(llm.clave_api()), "clave": llm.clave_enmascarada(), "modelo": llm.MODELO}


app.mount("/", StaticFiles(directory=STATIC, html=True), name="static")
