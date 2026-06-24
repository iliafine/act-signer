"""
Извлечение чистой подписи (PNG с альфа-каналом) из скана/фото документа,
где подпись может лежать на нежелезном (не белом) фоне и пересекаться
с другими элементами (печать, линия, текст).

Идея: подпись — это связные тёмные/насыщенные (чернильные) штрихи.
Остальное — почти однородный светлый фон листа. Через адаптивную
бинаризацию + морфологию выделяем штрихи, обрезаем по их bbox и
кладём в альфа-канал, заливая сам штрих одним цветом чернил.
"""
from __future__ import annotations

import cv2
import numpy as np
from PIL import Image


def _load_bgr(path: str) -> np.ndarray:
    img = cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR)
    if img is None:
        raise ValueError(f"Не удалось прочитать изображение: {path}")
    return img


def _auto_ink_mask(img: np.ndarray) -> np.ndarray:
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)

    # Адаптивная бинаризация устойчива к неоднородному/нежелезному фону.
    thresh = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV, blockSize=35, C=10,
    )

    # Доп. отсев "глобально светлого" фона по Оцу, чтобы убрать крупные
    # однородные заливки (тени, цветной фон) — берём пересечение масок.
    _, otsu = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    mask = cv2.bitwise_and(thresh, otsu)

    kernel_small = np.ones((2, 2), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel_small, iterations=1)
    kernel_close = np.ones((3, 3), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel_close, iterations=2)
    return mask


def _is_ruled_line(w: int, h: int, area: int) -> bool:
    """Печатная линия для подчёркивания / граница ячейки таблицы — почти
    прямой и тонкий штрих, в отличие от росчерка подписи (который изгибается
    и поэтому занимает заметно бОльшую высоту/ширину bbox, чем толщина линии).
    Отличаем по: 1) очень тонкая сторона bbox (<=4px) при длинной другой,
    2) почти полная заливка своего bbox (типично для прямой линии)."""
    short_side, long_side = min(w, h), max(w, h)
    if short_side > 4 or long_side < 20:
        return False
    fill_ratio = area / max(w * h, 1)
    return fill_ratio > 0.5


def _blue_ink_mask(img: np.ndarray) -> np.ndarray:
    """Выделяет только синие/насыщенные чернила, игнорируя чёрные линии
    (разлиновку, границы таблицы, печатный текст) рядом с подписью."""
    f = img.astype(np.float32) / 255.0
    b, g, r = f[:, :, 0], f[:, :, 1], f[:, :, 2]
    maxc = f.max(axis=2)
    minc = f.min(axis=2)
    sat = np.where(maxc == 0, 0, (maxc - minc) / np.maximum(maxc, 1e-6))

    mask = ((b > r + 0.05) & (sat > 0.13)).astype(np.uint8) * 255
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, np.ones((2, 2), np.uint8), iterations=1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, np.ones((3, 3), np.uint8), iterations=1)
    return mask


def extract_signature(
    input_path: str,
    output_path: str,
    crop_box: tuple[int, int, int, int] | None = None,
    pad: int = 12,
    ink_rgb: tuple[int, int, int] | None = None,
    color_mode: str = "auto",
) -> str:
    """
    crop_box:   необязательная (x0, y0, x1, y1) — если известно, где на скане
                примерно находится подпись, сильно повышает качество результата.
    ink_rgb:    если задано — штрихи перекрашиваются в этот цвет (например (0,0,0)).
                Если None — сохраняется исходный цвет чернил.
    color_mode: "auto" — обычная бинаризация (тёмные штрихи на светлом фоне);
                "blue_ink" — выделять только синие/цветные чернила, чтобы
                игнорировать чёрные линии разметки/таблицы рядом с подписью
                (например, подпись внутри ячейки таблицы со своими линиями).
    """
    img = _load_bgr(input_path)

    if crop_box:
        x0, y0, x1, y1 = crop_box
        img = img[max(y0, 0):y1, max(x0, 0):x1]

    if color_mode == "blue_ink":
        mask = _blue_ink_mask(img)
    else:
        mask = _auto_ink_mask(img)

    # Оставляем только компоненты разумного размера (убираем точечный шум
    # и одновременно — сплошные крупные заливки типа печати, если она
    # значительно крупнее штрихов подписи, путём фильтра по площади/форме).
    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(mask, connectivity=8)
    clean_mask = np.zeros_like(mask)
    areas = stats[1:, cv2.CC_STAT_AREA]
    if len(areas) > 0:
        area_thresh = max(15, np.percentile(areas, 40) * 0.2)
        main_idx = int(np.argmax(areas)) + 1
        main_y0, main_y1 = stats[main_idx, cv2.CC_STAT_TOP], stats[main_idx, cv2.CC_STAT_TOP] + stats[main_idx, cv2.CC_STAT_HEIGHT]
        main_area = stats[main_idx, cv2.CC_STAT_AREA]
        for i in range(1, num_labels):
            area = stats[i, cv2.CC_STAT_AREA]
            if area < area_thresh:
                continue
            w = stats[i, cv2.CC_STAT_WIDTH]
            h = stats[i, cv2.CC_STAT_HEIGHT]
            if _is_ruled_line(w, h, area):
                continue
            if i != main_idx and area < 0.1 * main_area:
                # Маленький обрывок, никак не пересекающийся по высоте с
                # основным росчерком — почти наверняка засечка на линии
                # подчёркивания/границе ячейки, а не часть подписи.
                y0, y1 = stats[i, cv2.CC_STAT_TOP], stats[i, cv2.CC_STAT_TOP] + stats[i, cv2.CC_STAT_HEIGHT]
                if y1 <= main_y0 or y0 >= main_y1:
                    continue
            clean_mask[labels == i] = 255

    if clean_mask.sum() == 0:
        clean_mask = mask  # fallback: ничего не отфильтровывать, если порог слишком жёсткий

    ys, xs = np.where(clean_mask > 0)
    if len(xs) == 0:
        raise ValueError("Не удалось найти штрихи подписи на изображении")

    x0b, x1b = max(xs.min() - pad, 0), min(xs.max() + pad, clean_mask.shape[1])
    y0b, y1b = max(ys.min() - pad, 0), min(ys.max() + pad, clean_mask.shape[0])

    crop_mask = clean_mask[y0b:y1b, x0b:x1b]
    crop_img = img[y0b:y1b, x0b:x1b]

    alpha = crop_mask
    if ink_rgb is not None:
        color = np.zeros_like(crop_img)
        color[:, :] = (ink_rgb[2], ink_rgb[1], ink_rgb[0])  # BGR
        rgb = color
    else:
        rgb = crop_img

    rgba = cv2.cvtColor(rgb, cv2.COLOR_BGR2BGRA)
    rgba[:, :, 3] = alpha

    out = Image.fromarray(cv2.cvtColor(rgba, cv2.COLOR_BGRA2RGBA))
    out.save(output_path, "PNG")
    return output_path


# Насыщенный "чернильный" синий — как у эталонных подписей (Кудлай/Рожин).
SIGNATURE_INK_RGB = (32, 74, 174)


def intensify_signature(
    path: str,
    ink_rgb: tuple[int, int, int] = SIGNATURE_INK_RGB,
    thickness_frac: float = 0.012,
    out_path: str | None = None,
) -> str:
    """Делает уже извлечённую подпись «жирнее и насыщеннее»: утолщает штрихи
    (дилатация ядром, пропорциональным размеру картинки — чтобы после сжатия
    до ~1 см в документе линия не превращалась в волосок) и перекрашивает в
    ровный насыщенный синий. Нужно для тонких бледных росчерков с чистого
    листа (шариковая ручка), которые иначе теряются на фоне текста."""
    out_path = out_path or path
    im = np.array(Image.open(path).convert("RGBA"))
    alpha = im[:, :, 3]

    h = alpha.shape[0]
    k = max(3, int(h * thickness_frac) | 1)  # нечётный размер ядра
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (k, k))
    thick = cv2.dilate((alpha > 40).astype(np.uint8) * 255, kernel, iterations=1)
    # лёгкое сглаживание краёв, чтобы не было «лесенки» после сжатия
    thick = cv2.GaussianBlur(thick, (3, 3), 0)

    rgba = np.zeros_like(im)
    rgba[:, :, 0] = ink_rgb[0]
    rgba[:, :, 1] = ink_rgb[1]
    rgba[:, :, 2] = ink_rgb[2]
    rgba[:, :, 3] = thick

    ys, xs = np.where(thick > 10)
    if len(xs):
        pad = 4
        x0, x1 = max(xs.min() - pad, 0), min(xs.max() + pad, rgba.shape[1])
        y0, y1 = max(ys.min() - pad, 0), min(ys.max() + pad, rgba.shape[0])
        rgba = rgba[y0:y1, x0:x1]

    Image.fromarray(rgba).save(out_path, "PNG")
    return out_path


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("Использование: python image_utils.py <вход> <выход.png> [x0,y0,x1,y1]")
        raise SystemExit(1)

    box = None
    if len(sys.argv) > 3:
        box = tuple(int(v) for v in sys.argv[3].split(","))

    extract_signature(sys.argv[1], sys.argv[2], crop_box=box)
    print("Готово:", sys.argv[2])
