from __future__ import annotations

import io
import json
import shutil
import subprocess
import uuid
import tempfile
import zipfile
from urllib.parse import quote
from dataclasses import asdict
from pathlib import Path, PurePosixPath

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


def _is_usable_docx(path: Path) -> bool:
    """Отсекает служебные файлы macOS из zip: __MACOSX и ._*.docx не являются Word-документами."""
    return (
        path.suffix.lower() == ".docx"
        and not path.name.startswith("._")
        and "__MACOSX" not in path.parts
    )


def _decode_zip_name(name: str) -> str:
    """Восстанавливает кириллицу из zip без UTF-8-флага."""
    try:
        raw = name.encode("cp437")
    except UnicodeEncodeError:
        return name
    for encoding in ("utf-8", "cp866"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            pass
    return name


def _safe_extract_zip(zip_path: Path, extract_dir: Path) -> None:
    with zipfile.ZipFile(zip_path) as zf:
        for info in zf.infolist():
            decoded_name = _decode_zip_name(info.filename)
            rel = PurePosixPath(decoded_name)
            if rel.is_absolute() or ".." in rel.parts:
                continue

            target = extract_dir.joinpath(*rel.parts)
            if info.is_dir():
                target.mkdir(parents=True, exist_ok=True)
                continue

            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(zf.read(info))


def _prefixed_output_path(output_dir: Path, rel: Path, export_format: str) -> Path:
    suffix = ".pdf" if export_format == "pdf" else rel.suffix
    return output_dir / rel.with_name(f"Подписанный {rel.stem}{suffix}")


def _report_status_label(status: str) -> str:
    return {
        "not_found_in_db": "нет в базе",
        "no_target_location": "не найдено место для подписи",
        "signed": "подписано",
    }.get(status, status)


def _friendly_error(error: Exception) -> str:
    if type(error).__name__ == "PackageNotFoundError":
        return "файл не похож на настоящий .docx-документ Word"
    return f"{type(error).__name__}: {error}"


def _write_issue_report(report_path: Path, issues: list[dict]) -> None:
    if not issues:
        return

    lines = ["Отчет по подписанию", ""]
    for issue in issues:
        lines.append(f"Файл: {issue['file']}")
        if issue.get("error"):
            lines.append(f"Ошибка: {issue['error']}")
        for result in issue.get("results", []):
            role = " ".join((result.get("role_text") or "").split())
            status = _report_status_label(result.get("status", ""))
            signer = result.get("matched_signer") or "не найден"
            lines.append(f"- {status}: {signer}")
            if role:
                lines.append(f"  Роль: {role}")
        lines.append("")

    report_path.write_text("\n".join(lines).strip() + "\n", encoding="utf-8")


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    html_path = STATIC_DIR / "index.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


@app.get("/api/signers")
def list_signers() -> JSONResponse:
    return JSONResponse([asdict(s) for s in load_db()])


def _find_soffice() -> str | None:
    """Ищет LibreOffice кросс-платформенно. На Windows soffice обычно не в PATH,
    поэтому проверяем стандартные пути установки; на macOS — бандл приложения."""
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/opt/homebrew/bin/soffice",
    ]
    for path in candidates:
        if Path(path).exists():
            return path
    return None


def _convert_to_pdf(docx_path: Path, out_dir: Path) -> Path:
    """Конвертирует .docx в .pdf через LibreOffice headless. Делаем это уже
    из подписанного документа, поэтому "плавающая" подпись (наложенная на
    текст) сохраняется и в PDF — в отличие от вставки картинки прямо в PDF."""
    soffice = _find_soffice()
    if not soffice:
        raise HTTPException(
            status_code=500,
            detail="Для экспорта в PDF нужен установленный LibreOffice. "
                   "Скачайте с libreoffice.org и установите, затем повторите. "
                   "Без него экспорт в .docx работает как обычно.",
        )
    with tempfile.TemporaryDirectory() as profile_dir:
        # Отдельный профиль на каждый вызов: параллельные/быстрые подряд
        # запуски soffice конфликтуют за общий профиль и портят результат.
        # as_uri() даёт корректный file:// и на Windows (file:///C:/...), и на *nix.
        result = subprocess.run(
            [
                soffice, "--headless", "--norestore",
                f"-env:UserInstallation={Path(profile_dir).as_uri()}",
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

        out_name = f"sig_{uuid.uuid4().hex}.png"  # ASCII-имя: ФИО бывает кириллическим, ломает перенос на Windows
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
            out_name = f"sig_{uuid.uuid4().hex}.png"  # ASCII-имя: ФИО бывает кириллическим, ломает перенос на Windows
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
    filename = Path(file.filename or "act").name
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
            _safe_extract_zip(in_zip_path, extract_dir)

            output_dir = tmp_path / "output"
            output_dir.mkdir()
            issues: list[dict] = []

            docx_files = [path for path in extract_dir.rglob("*.docx") if _is_usable_docx(path.relative_to(extract_dir))]
            if not docx_files:
                raise HTTPException(status_code=400, detail="В zip не найдено настоящих .docx-файлов")

            for docx_file in docx_files:
                rel = docx_file.relative_to(extract_dir)
                out_file = _prefixed_output_path(output_dir, rel, export_format)
                try:
                    signed_docx = out_file.with_suffix(".docx") if export_format == "pdf" else out_file
                    report = sign_document(str(docx_file), str(signed_docx))
                    report_dict = _report_to_dict(report)
                    if not report.all_signed:
                        issues.append({"file": str(rel), "results": report_dict["results"]})

                    if export_format == "pdf":
                        pdf_path = _convert_to_pdf(signed_docx, signed_docx.parent)
                        signed_docx.unlink()
                        if pdf_path != out_file:
                            if out_file.exists():
                                out_file.unlink()
                            pdf_path.rename(out_file)
                except Exception as e:  # noqa: BLE001
                    issues.append({"file": str(rel), "error": _friendly_error(e)})

            _write_issue_report(output_dir / "Отчет.txt", issues)

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
                headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote('Подписанные акты.zip')}"},
            )

        elif suffix == ".docx":
            in_path = tmp_path / filename
            in_path.write_bytes(content)
            out_path = tmp_path / f"Подписанный {filename}"
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
            report_json = json.dumps(_report_to_dict(report), ensure_ascii=False)
            response.headers["X-Sign-Report"] = quote(report_json)
            return response
        else:
            raise HTTPException(status_code=400, detail="Поддерживаются только .docx и .zip")
