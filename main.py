"""
main.py
=======
Aplicación FastAPI para el Optimizador de Horarios Docentes.
Sin base de datos; todo opera con el Excel subido por el usuario.
"""

from __future__ import annotations

import json
import os
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from core.optimizer import (
    generate_excel,
    get_horarios_labels,
    get_professors_list,
    load_excel,
    optimize_assignments,
    parse_disponibilidad,
)

# ---------------------------------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------------------------------
BASE_DIR  = Path(__file__).parent
TEMPLATES = Jinja2Templates(directory=str(BASE_DIR / "templates"))
TEMP_DIR  = Path(tempfile.gettempdir()) / "horarios_app"
TEMP_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Optimizador de Horarios Docentes", version="2.0.0")
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


# ---------------------------------------------------------------------------
# RUTAS
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return TEMPLATES.TemplateResponse(
        request=request,
        name="index.html",
        context={}
    )


@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    """
    Recibe el Excel, lo valida y devuelve:
    - session_id para identificar la sesión
    - lista de profesores
    - lista de clases/horarios con etiquetas
    - vista previa de horarios (primeras 20 filas)
    """
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="El archivo debe ser .xlsx")

    contents = await file.read()
    session_id = str(uuid.uuid4())
    session_path = TEMP_DIR / f"{session_id}.xlsx"
    session_path.write_bytes(contents)

    df_disp_raw, df_hor, error = load_excel(str(session_path))
    if error:
        session_path.unlink(missing_ok=True)
        raise HTTPException(status_code=422, detail=error)

    profesores  = get_professors_list(df_disp_raw)
    clases      = get_horarios_labels(df_hor)
    total_profs = len(profesores)
    total_hor   = len(df_hor)

    # Vista previa: primeras 20 filas, solo columnas relevantes
    preview_cols = ["id_horario", "semestre", "lengua", "curso", "grupo",
                    "dia", "hora_inicio", "hora_fin", "modalidad"]
    preview_cols = [c for c in preview_cols if c in df_hor.columns]
    preview = df_hor[preview_cols].head(20).fillna("").to_dict("records")

    return JSONResponse({
        "session_id":   session_id,
        "total_profs":  total_profs,
        "total_hor":    total_hor,
        "profesores":   profesores,
        "clases":       clases,
        "preview":      preview,
        "preview_cols": preview_cols,
    })


@app.post("/optimize")
async def optimize(
    session_id:    str  = Form(...),
    restricciones: str  = Form(default="[]"),
):
    """
    Ejecuta la optimización y devuelve:
    - métricas
    - tabla de resultados
    - result_id para descargar el Excel
    """
    session_path = TEMP_DIR / f"{session_id}.xlsx"
    if not session_path.exists():
        raise HTTPException(status_code=404,
                            detail="Sesión no encontrada. Por favor suba el archivo de nuevo.")

    try:
        rest_list = json.loads(restricciones)
    except json.JSONDecodeError:
        rest_list = []

    df_disp_raw, df_hor, error = load_excel(str(session_path))
    if error:
        raise HTTPException(status_code=422, detail=error)

    df_disp = parse_disponibilidad(df_disp_raw)
    df_result, metodo = optimize_assignments(df_disp, df_hor, rest_list)

    asignados   = int((df_result["Estado"] == "Asignado").sum())
    no_asig     = int((df_result["Estado"] == "No asignado").sum())
    total       = len(df_result)
    cobertura   = round(asignados / total * 100, 1) if total else 0.0

    # Guardar Excel de resultado
    result_id   = str(uuid.uuid4())
    excel_bytes = generate_excel(df_result, metodo)
    result_path = TEMP_DIR / f"result_{result_id}.xlsx"
    result_path.write_bytes(excel_bytes)

    # Tabla de resultados (todas las filas)
    tabla = df_result.fillna("").to_dict("records")

    return JSONResponse({
        "metodo":     metodo,
        "total":      total,
        "asignados":  asignados,
        "no_asig":    no_asig,
        "cobertura":  cobertura,
        "result_id":  result_id,
        "tabla":      tabla,
    })


@app.get("/download/{result_id}")
async def download(result_id: str):
    """Descarga el Excel de resultados generado."""
    # Sanitize: solo UUID-like
    if not all(c in "0123456789abcdef-" for c in result_id):
        raise HTTPException(status_code=400, detail="ID inválido.")
    result_path = TEMP_DIR / f"result_{result_id}.xlsx"
    if not result_path.exists():
        raise HTTPException(status_code=404, detail="Resultado no encontrado.")
    return FileResponse(
        path=str(result_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="asignacion_optimizada.xlsx",
    )


@app.get("/health")
async def health():
    return {"status": "ok"}
