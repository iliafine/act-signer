from __future__ import annotations

import io
import shutil
import subprocess
import tempfile
import zipfile
from urllib.parse import quote
from dataclasses import asdict
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

from app.docx_signer import sign_document, SignReport
from app.image_utils import extract_signature, intensify_signature
from app.signature_db import add_signer, delete_signer, get_signer, load_db, update_signer, SIGNATURES_DIR

app = FastAPI(title="Сервис подписания актов")

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


def _report_to_dict(report: SignReport) -> dict:
    return {
        "all_signed": report.all_signed,
        "results": [
            {
                "role_text": r.role_text,
                "status": r.status,
                "match_score": r.match_score,
                "matched_signer": r.matched_signer.full_name if r.matched_signer else None,
                "matched_by": r.matched_by,
            }
            for r in report.results
        ],
    }


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    html_path = STATIC_DIR / "index.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


@app.get("/api/signers")
def list_signers() -> JSONResponse:
    return JSONResponse([asdict(s) for s in load_db()])


def _convert_to_pdf(docx_path: Path, out_dir: Path) -> Path:
    """Конвертирует .docx в .pdf через LibreOffice headless. Делаем это уже
    из подписанного документа, поэтому "плавающая" подпись (наложенная на
    текст) сохраняется и в PDF — в отличие от вставки картинки прямо в PDF."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise HTTPException(
            status_code=500,
            detail="Для экспорта в PDF на сервере должен быть установлен LibreOffice (soffice)",
        )
    with tempfile.TemporaryDirectory() as profile_dir:
        # Отдельный профиль на каждый вызов: параллельные/быстрые подряд
        # запуски soffice конфликтуют за общий профиль и портят результат.
        result = subprocess.run(
            [
                soffice, "--headless", "--norestore",
                f"-env:UserInstallation=file://{profile_dir}",
                "--convert-to", "pdf", "--outdir", str(out_dir), str(docx_path),
            ],
            capture_output=True, text=True, timeout=120,
        )
    pdf_path = out_dir / (docx_path.stem + ".pdf")
    if result.returncode != 0 or not pdf_path.exists():
        raise HTTPException(status_code=500, detail=f"Не удалось конвертировать в PDF: {result.stderr}")
    return pdf_path


def _extract_with_fallback(raw_path: Path, out_path: Path) -> None:
    """Сначала обычная бинаризация; если ничего не нашла (или скан с
    линиями таблицы мешает), пробуем выделение только синих/цветных чернил.
    Затем усиливаем (утолщаем + насыщаем цвет), чтобы тонкие бледные росчерки
    с чистого листа не превращались в волоски при вставке в документ."""
    try:
        extract_signature(str(raw_path), str(out_path), color_mode="auto")
    except ValueError:
        extract_signature(str(raw_path), str(out_path), color_mode="blue_ink")
    intensify_signature(str(out_path))


@app.post("/api/signers")
async def create_signer(
    full_name: str = Form(...),
    role: str = Form(""),
    aliases: str = Form(""),
    scale: float = Form(1.0),
    scan: UploadFile = File(...),
):
    if not full_name.strip():
        raise HTTPException(status_code=422, detail="ФИО обязательно для заполнения")

    SIGNATURES_DIR.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        raw_path = Path(tmp) / scan.filename
        raw_path.write_bytes(await scan.read())

        out_name = f"sig_{full_name.strip().replace(' ', '_')}.png"
        out_path = SIGNATURES_DIR / out_name
        try:
            _extract_with_fallback(raw_path, out_path)
        except ValueError as e:
            raise HTTPException(status_code=422, detail=str(e))

    alias_list = [a.strip() for a in aliases.split(";") if a.strip()]
    signer = add_signer(full_name=full_name, role=role, file=out_name, aliases=alias_list, scale=scale)
    return JSONResponse(asdict(signer))


@app.put("/api/signers/{signer_id}")
async def edit_signer(
    signer_id: str,
    full_name: str = Form(...),
    role: str = Form(""),
    aliases: str = Form(""),
    scale: float = Form(1.0),
    scan: UploadFile | None = File(None),
):
    if not full_name.strip():
        raise HTTPException(status_code=422, detail="ФИО обязательно для заполнения")
    if get_signer(signer_id) is None:
        raise HTTPException(status_code=404, detail="Подписант не найден")

    alias_list = [a.strip() for a in aliases.split(";") if a.strip()]
    new_file = None

    if scan is not None and scan.filename:
        SIGNATURES_DIR.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory() as tmp:
            raw_path = Path(tmp) / scan.filename
            raw_path.write_bytes(await scan.read())
            out_name = f"sig_{full_name.strip().replace(' ', '_')}.png"
            out_path = SIGNATURES_DIR / out_name
            try:
                _extract_with_fallback(raw_path, out_path)
            except ValueError as e:
                raise HTTPException(status_code=422, detail=str(e))
            new_file = out_name

    signer = update_signer(signer_id, full_name=full_name, role=role, aliases=alias_list, file=new_file, scale=scale)
    return JSONResponse(asdict(signer))


@app.delete("/api/signers/{signer_id}")
def remove_signer(signer_id: str):
    if not delete_signer(signer_id):
        raise HTTPException(status_code=404, detail="Подписант не найден")
    return JSONResponse({"ok": True})


@app.get("/api/signers/{signer_id}/signature")
def get_signature_image(signer_id: str):
    signer = get_signer(signer_id)
    if signer is None or not signer.signature_path().exists():
        raise HTTPException(status_code=404, detail="Подпись не найдена")
    return FileResponse(str(signer.signature_path()), media_type="image/png")


@app.post("/api/sign")
async def sign_act(file: UploadFile = File(...), export_format: str = Form("docx")):
    filename = file.filename or "act"
    suffix = Path(filename).suffix.lower()
    content = await file.read()
    export_format = export_format.lower().strip()
    if export_format not in ("docx", "pdf"):
        export_format = "docx"

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)

        if suffix == ".zip":
            in_zip_path = tmp_path / "input.zip"
            in_zip_path.write_bytes(content)
            extract_dir = tmp_path / "extracted"
            extract_dir.mkdir()
            with zipfile.ZipFile(in_zip_path) as zf:
                zf.extractall(extract_dir)

            output_dir = tmp_path / "output"
            output_dir.mkdir()
            full_report: dict[str, dict] = {}

            for docx_file in extract_dir.rglob("*.docx"):
                rel = docx_file.relative_to(extract_dir)
                out_file = output_dir / rel
                report = sign_document(str(docx_file), str(out_file))
                full_report[str(rel)] = _report_to_dict(report)

                if export_format == "pdf":
                    pdf_path = _convert_to_pdf(out_file, out_file.parent)
                    out_file.unlink()
                    pdf_path.rename(out_file.with_suffix(".pdf"))

            (output_dir / "_report.json").write_text(
                __import__("json").dumps(full_report, ensure_ascii=False, indent=2), encoding="utf-8"
            )

            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for item in output_dir.rglob("*"):
                    if item.is_file():
                        zf.write(item, item.relative_to(output_dir))
            zip_buf.seek(0)

            # Отдаём содержимое напрямую (не через FileResponse с диска):
            # временная папка удаляется сразу при выходе из `with` ниже, а
            # FileResponse читает файл лениво — после удаления папки это
            # привело бы к "File does not exist".
            return Response(
                content=zip_buf.read(),
                media_type="application/zip",
                headers={"Content-Disposition": 'attachment; filename="signed_acts.zip"'},
            )

        elif suffix == ".docx":
            in_path = tmp_path / filename
            in_path.write_bytes(content)
            out_path = tmp_path / f"signed_{filename}"
            report = sign_document(str(in_path), str(out_path))

            persisted = STATIC_DIR.parent / "_last_output"
            persisted.mkdir(exist_ok=True)

            if export_format == "pdf":
                pdf_path = _convert_to_pdf(out_path, tmp_path)
                final_name = pdf_path.stem + ".pdf"
                final_path = persisted / final_name
                shutil.copy(pdf_path, final_path)
                media_type = "application/pdf"
            else:
                final_name = out_path.name
                final_path = persisted / final_name
                shutil.copy(out_path, final_path)
                media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

            response = FileResponse(str(final_path), filename=final_name, media_type=media_type)
            report_json = __import__("json").dumps(_report_to_dict(report), ensure_ascii=False)
            response.headers["X-Sign-Report"] = quote(report_json)
            return response
        else:
            raise HTTPException(status_code=400, detail="Поддерживаются только .docx и .zip")
