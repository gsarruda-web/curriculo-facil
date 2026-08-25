from pathlib import Path
import tempfile
import re
import unicodedata

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from typing import Any
from starlette.background import BackgroundTask

from gerador_curriculo import gerar_curriculo

BASE = Path(__file__).resolve().parent
STATIC = BASE / "static"

app = FastAPI(
    title="Currículo Fácil",
    version="1.1.0",
    description="MVP para geração guiada de currículo em Word.",
)
app.mount("/static", StaticFiles(directory=STATIC), name="static")


class Payload(BaseModel):
    data: dict[str, Any] = Field(default_factory=dict)


def _cleanup(path: Path, folder: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    finally:
        try:
            folder.rmdir()
        except OSError:
            pass


@app.get("/")
def home():
    return FileResponse(STATIC / "index.html")


@app.get("/privacidade")
def privacidade():
    return FileResponse(STATIC / "privacidade.html")


@app.get("/health")
def health():
    return {"status": "ok", "versao": "1.1.0"}


@app.post("/gerar")
def gerar(payload: Payload):
    data = payload.data
    nome = str(data.get("nome_completo", "")).strip()
    telefone = str(data.get("telefone", "")).strip()
    cidade_uf = str(data.get("cidade_uf", "")).strip()
    cargo = str(data.get("cargo_objetivo", "")).strip()

    if not all([nome, telefone, cidade_uf, cargo]):
        raise HTTPException(
            status_code=422,
            detail="Nome, cidade/UF, telefone e objetivo profissional são obrigatórios.",
        )

    folder = Path(tempfile.mkdtemp(prefix="curriculo_facil_"))
    ascii_name = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode("ascii")
    ascii_name = re.sub(r"[^A-Za-z0-9]+", "_", ascii_name).strip("_") or "Curriculo"
    output = folder / f"Curriculo_{ascii_name}.docx"

    try:
        out = gerar_curriculo(data, output_path=str(output))
    except Exception as exc:
        _cleanup(output, folder)
        raise HTTPException(status_code=500, detail="Não foi possível gerar o currículo.") from exc

    # O arquivo é apagado do servidor após o envio ao usuário.
    download_name = out.name
    return FileResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=download_name,
        background=BackgroundTask(_cleanup, out, folder),
    )
