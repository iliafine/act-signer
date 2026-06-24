"""
Простая JSON-база подписантов: ФИО (обязательно), должность/роль
(опционально, как она пишется в акте) и путь к файлу подписи.

Сопоставление с актом приоритетно идёт по ФИО (надёжнее — в акте обычно
напечатана фамилия с инициалами рядом с местом подписи). Должность/роль
используется только как запасной вариант, если по имени найти не удалось
(например, в акте напечатана только должность без ФИО).
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, asdict, field
from pathlib import Path

from rapidfuzz import fuzz

SIGNATURES_DIR = Path(__file__).resolve().parent.parent.parent / "signatures"
DB_PATH = SIGNATURES_DIR / "db.json"

FUZZY_ROLE_THRESHOLD = 80   # 0-100, насколько похож текст роли в акте на роль в базе
FUZZY_NAME_THRESHOLD = 85   # 0-100, насколько похоже ФИО в акте на ФИО в базе


@dataclass
class Signer:
    id: str
    full_name: str
    role: str = ""
    aliases: list[str] = field(default_factory=list)
    file: str = ""  # имя файла внутри SIGNATURES_DIR
    scale: float = 1.0  # множитель размера подписи (для размашистых росчерков)
    vertical_offset_frac: float = 0.0  # сдвиг вниз относительно "полки" в долях высоты подписи

    def signature_path(self) -> Path:
        return SIGNATURES_DIR / self.file


def load_db() -> list[Signer]:
    if not DB_PATH.exists():
        return []
    raw = json.loads(DB_PATH.read_text(encoding="utf-8"))
    return [Signer(**item) for item in raw]


def save_db(signers: list[Signer]) -> None:
    SIGNATURES_DIR.mkdir(parents=True, exist_ok=True)
    DB_PATH.write_text(
        json.dumps([asdict(s) for s in signers], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def get_signer(signer_id: str) -> Signer | None:
    return next((s for s in load_db() if s.id == signer_id), None)


def add_signer(full_name: str, file: str, role: str = "", aliases: list[str] | None = None, scale: float = 1.0) -> Signer:
    signers = load_db()
    new_id = str(max((int(s.id) for s in signers if s.id.isdigit()), default=0) + 1)
    signer = Signer(id=new_id, full_name=full_name.strip(), role=(role or "").strip(), aliases=aliases or [], file=file, scale=scale)
    signers.append(signer)
    save_db(signers)
    return signer


def update_signer(
    signer_id: str,
    full_name: str | None = None,
    role: str | None = None,
    aliases: list[str] | None = None,
    file: str | None = None,
    scale: float | None = None,
) -> Signer | None:
    signers = load_db()
    for signer in signers:
        if signer.id == signer_id:
            if full_name is not None:
                signer.full_name = full_name.strip()
            if role is not None:
                signer.role = role.strip()
            if aliases is not None:
                signer.aliases = aliases
            if file is not None:
                signer.file = file
            if scale is not None:
                signer.scale = scale
            save_db(signers)
            return signer
    return None


def delete_signer(signer_id: str) -> bool:
    signers = load_db()
    remaining = [s for s in signers if s.id != signer_id]
    if len(remaining) == len(signers):
        return False
    save_db(remaining)
    return True


# --- Сопоставление по ФИО -----------------------------------------------

_INITIAL_RE = re.compile(r"^[А-ЯЁA-Z]\.?$")


def _name_key(name_text: str) -> str | None:
    """Приводит ФИО (полное или 'Фамилия И.О.') к ключу 'фамилия и о'
    для устойчивого сравнения независимо от порядка и формата написания."""
    tokens = re.findall(r"[А-ЯЁа-яёA-Za-z]+\.?", name_text)
    surname = None
    initials: list[str] = []
    for tok in tokens:
        clean = tok.rstrip(".")
        if not clean:
            continue
        if len(clean) == 1:
            initials.append(clean.lower())
        elif surname is None:
            surname = clean.lower()
        else:
            initials.append(clean[0].lower())
    if surname is None:
        return None
    return surname + " " + "".join(sorted(initials))


def find_signer_by_name(name_text: str, signers: list[Signer] | None = None) -> tuple[Signer | None, int]:
    signers = signers if signers is not None else load_db()
    candidate_key = _name_key(name_text)
    if candidate_key is None:
        return None, 0

    best: tuple[Signer | None, int] = (None, 0)
    for signer in signers:
        signer_key = _name_key(signer.full_name)
        if signer_key is None:
            continue
        score = fuzz.ratio(candidate_key, signer_key)
        if score > best[1]:
            best = (signer, score)

    if best[1] < FUZZY_NAME_THRESHOLD:
        return None, best[1]
    return best


def find_signer_for_role(role_text: str, signers: list[Signer] | None = None) -> tuple[Signer | None, int]:
    """Запасной вариант: ищет подписанта по тексту должности/роли, найденному
    в акте. Используется только если сопоставление по ФИО не дало результата."""
    signers = signers if signers is not None else load_db()
    best: tuple[Signer | None, int] = (None, 0)
    role_norm = role_text.strip().lower()

    for signer in signers:
        candidates = [signer.role, *signer.aliases]
        candidates = [c for c in candidates if c.strip()]
        if not candidates:
            continue
        score = max(fuzz.token_sort_ratio(role_norm, c.strip().lower()) for c in candidates)
        if score > best[1]:
            best = (signer, score)

    if best[1] < FUZZY_ROLE_THRESHOLD:
        return None, best[1]
    return best
