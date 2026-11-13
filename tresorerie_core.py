"""
tresorerie_core.py — logique métier (sans Streamlit) pour le suivi de trésorerie
des jardins familiaux.

  * détection automatique de la structure du tableau (mois, blocs Entrées / Sorties)
  * ajout d'un montant dans la cellule (catégorie × mois) avec sauvegarde de sécurité
  * lecture du tableau avec recalcul des totaux (openpyxl ne calcule pas les formules)
  * numérisation de document (OpenCV), OCR (Tesseract / EasyOCR), extraction
    du montant TTC, de la date et suggestion de catégorie
"""
from __future__ import annotations

import csv
import datetime as dt
import io
import os
import re
import shutil
import unicodedata
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path

from PIL import Image, ImageOps

MOIS_FR = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet",
           "août", "septembre", "octobre", "novembre", "décembre"]


def norm(s) -> str:
    """Minuscule, sans accents, espaces compactés : sert à comparer des libellés."""
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().lower()


MOIS_NORM = [norm(m) for m in MOIS_FR]


class LayoutError(ValueError):
    """La structure attendue du tableau n'a pas été trouvée."""


class OcrUnavailable(RuntimeError):
    """Le moteur OCR demandé n'est pas utilisable (non installé, etc.)."""


# ════════════════════════════════════════════════════════════════════════════
#  1. STRUCTURE DU TABLEAU EXCEL
# ════════════════════════════════════════════════════════════════════════════
@dataclass
class Block:
    title_row: int
    first_row: int      # première ligne de catégorie
    last_row: int       # dernière ligne de catégorie (juste avant « Total … »)
    total_row: int

    def rows(self):
        return range(self.first_row, self.last_row + 1)


@dataclass
class Layout:
    label_col: int
    header_row: int
    month_cols: dict            # {1: n° colonne de janvier, ..., 12: décembre}
    start_col: int | None       # « Début d'activité »
    total_col: int | None       # « TOTAL »
    entrees: Block
    sorties: Block
    year: int | None


@dataclass
class Row:
    label: str
    row: int


def _find_label(ws, target: str, cols=(1, 2, 3), min_row=1):
    for c in cols:
        for r in range(min_row, ws.max_row + 1):
            if norm(ws.cell(r, c).value) == target:
                return r, c
    return None


def _block(ws, title: str, total_prefix: str, col: int, row: int) -> Block:
    total_row = None
    for r in range(row + 1, ws.max_row + 1):
        if norm(ws.cell(r, col).value).startswith(total_prefix):
            total_row = r
            break
    if total_row is None:
        raise LayoutError(f"Ligne « {total_prefix} » introuvable sous « {title} ».")
    return Block(title_row=row, first_row=row + 1, last_row=total_row - 1, total_row=total_row)


def detect_layout(ws) -> Layout:
    # ligne des mois
    header_row, month_cols = None, {}
    for r in range(1, min(ws.max_row, 25) + 1):
        found = {}
        for c in range(1, ws.max_column + 1):
            n = norm(ws.cell(r, c).value)
            if n in MOIS_NORM:
                found[MOIS_NORM.index(n) + 1] = c
        if len(found) >= 6:
            header_row, month_cols = r, found
            break
    if header_row is None:
        raise LayoutError("Ligne des mois (janvier … décembre) introuvable.")

    s = _find_label(ws, "sorties")
    if s is None:
        raise LayoutError("Section « Sorties » introuvable (colonne A ou B).")
    s_row, label_col = s
    e = _find_label(ws, "entrees", cols=(label_col,))
    if e is None:
        raise LayoutError("Section « Entrées » introuvable.")
    e_row = e[0]

    start_col = total_col = None
    for c in range(1, ws.max_column + 1):
        n = norm(ws.cell(header_row, c).value)
        if n.startswith("debut"):
            start_col = c
        elif n == "total":
            total_col = c

    year = None
    m = re.search(r"(20\d{2})", ws.title)
    if m:
        year = int(m.group(1))

    return Layout(
        label_col=label_col, header_row=header_row, month_cols=month_cols,
        start_col=start_col, total_col=total_col,
        entrees=_block(ws, "Entrées", "total entrees", label_col, e_row),
        sorties=_block(ws, "Sorties", "total sorties", label_col, s_row),
        year=year,
    )


def valid_sheets(path) -> list[str]:
    """Feuilles dont la structure est reconnue."""
    import openpyxl

    wb = openpyxl.load_workbook(path)
    out = []
    for ws in wb.worksheets:
        try:
            detect_layout(ws)
            out.append(ws.title)
        except LayoutError:
            pass
    return out


def num(v) -> float:
    """Valeur de cellule → float (les formules du type =5+10 sont additionnées)."""
    if v is None or v == "" or isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s.startswith("="):
        expr = s[1:].replace(",", ".")
        if re.fullmatch(r"[\d.+\- ]+", expr):
            return sum(float(t) for t in re.findall(r"[+-]?\d+(?:\.\d+)?", expr))
        return 0.0
    try:
        return float(s.replace("\u00a0", "").replace(" ", "").replace(",", "."))
    except ValueError:
        return 0.0


def _data_cols(ws, lay: Layout) -> list[int]:
    cols = [lay.month_cols[m] for m in range(1, 13)]
    return ([lay.start_col] if lay.start_col else []) + cols


def list_rows(ws, lay: Layout, block: Block):
    """(catégories nommées, lignes libres) d'un bloc."""
    cats, free = [], []
    cols = _data_cols(ws, lay)
    for r in block.rows():
        label = str(ws.cell(r, lay.label_col).value or "").strip()
        if label:
            cats.append(Row(label, r))
        elif all(ws.cell(r, c).value in (None, "") for c in cols):
            free.append(r)
    return cats, free


# ════════════════════════════════════════════════════════════════════════════
#  2. LECTURE DU TABLEAU (avec totaux recalculés)
# ════════════════════════════════════════════════════════════════════════════
@dataclass
class Model:
    ws: object
    layout: Layout
    sorties: list[Row]
    sorties_free: list[int]
    entrees: list[Row]
    entrees_free: list[int]
    month_headers: list[str]
    entrees_df: pd.DataFrame = field(default=None)
    sorties_df: pd.DataFrame = field(default=None)
    summary_df: pd.DataFrame = field(default=None)

    def value(self, row: int, month: int) -> float:
        return num(self.ws.cell(row, self.layout.month_cols[month]).value)


def _block_frame(ws, lay: Layout, block: Block, headers: list[str]) -> pd.DataFrame:
    import pandas as pd

    cols = _data_cols(ws, lay)
    labels, data = [], []
    for r in block.rows():
        label = str(ws.cell(r, lay.label_col).value or "").strip()
        vals = [num(ws.cell(r, c).value) for c in cols]
        if label or any(vals):
            labels.append(label or f"(ligne {r})")
            data.append(vals)
    df = pd.DataFrame(data, index=labels, columns=headers, dtype=float)
    df["Total"] = df.sum(axis=1)
    return df


def open_model(path, sheet: str) -> Model:
    import openpyxl
    import pandas as pd

    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]
    lay = detect_layout(ws)
    s, s_free = list_rows(ws, lay, lay.sorties)
    e, e_free = list_rows(ws, lay, lay.entrees)

    headers = ([str(ws.cell(lay.header_row, lay.start_col).value).strip()] if lay.start_col else []) \
        + [m.capitalize() for m in MOIS_FR]
    ent = _block_frame(ws, lay, lay.entrees, headers)
    sor = _block_frame(ws, lay, lay.sorties, headers)

    tot_e, tot_s = ent[headers].sum(), sor[headers].sum()
    diff = tot_e - tot_s
    treso = diff.cumsum()
    summary = pd.DataFrame([tot_e, tot_s, diff, treso],
                           index=["Total entrées", "Total sorties", "Différence", "Trésorerie (cumul)"])
    summary["Total"] = [tot_e.sum(), tot_s.sum(), diff.sum(), treso.iloc[-1]]

    return Model(ws=ws, layout=lay, sorties=s, sorties_free=s_free, entrees=e, entrees_free=e_free,
                 month_headers=headers, entrees_df=ent, sorties_df=sor, summary_df=summary)


# ════════════════════════════════════════════════════════════════════════════
#  3. ÉCRITURE DANS EXCEL
# ════════════════════════════════════════════════════════════════════════════
def add_amount(cell, amount: float):
    """Additionne `amount` au contenu de la cellule. Retourne (ancienne, nouvelle valeur)."""
    amt = round(float(amount), 2)
    old = cell.value
    if old in (None, ""):
        new = amt
    elif isinstance(old, bool):
        raise ValueError(f"Cellule {cell.coordinate} : valeur booléenne inattendue.")
    elif isinstance(old, (int, float)):
        new = round(old + amt, 2)
    elif isinstance(old, str) and old.startswith("="):
        new = f"{old}+{amt}"                      # on prolonge la formule existante
    else:
        try:
            new = round(float(str(old).replace(" ", "").replace(",", ".")) + amt, 2)
        except ValueError:
            raise ValueError(f"La cellule {cell.coordinate} contient du texte non numérique : {old!r}")
    cell.value = new
    return old, new


def save_with_backup(wb, path, backup_dir="sauvegardes") -> Path | None:
    path = Path(path)
    backup = None
    if path.exists():
        bdir = path.parent / backup_dir
        bdir.mkdir(exist_ok=True)
        backup = bdir / f"{path.stem}_{dt.datetime.now():%Y%m%d_%H%M%S}{path.suffix}"
        shutil.copy2(path, backup)
    wb.save(path)
    return backup


def _copy_row_style(ws, lay: Layout, src_row: int, dst_row: int):
    cols = [lay.label_col] + _data_cols(ws, lay) + ([lay.total_col] if lay.total_col else [])
    for c in cols:
        ws.cell(dst_row, c)._style = copy(ws.cell(src_row, c)._style)


def record_amount(path, sheet: str, which: str, month: int, amount: float, *,
                  label: str | None = None, new_label: str | None = None,
                  make_backup: bool = True, detail: dict | None = None) -> dict:
    """
    Ajoute `amount` dans (catégorie × mois) du bloc `which` ('sorties' ou 'entrees').
    `new_label` crée la catégorie dans la première ligne libre du bloc.
    """
    if not 1 <= int(month) <= 12:
        raise ValueError("Mois invalide.")
    if amount is None or amount <= 0:
        raise ValueError("Le montant doit être strictement positif.")

    import openpyxl

    wb = openpyxl.load_workbook(path)          # formules conservées
    ws = wb[sheet]
    lay = detect_layout(ws)
    block = lay.sorties if which == "sorties" else lay.entrees
    cats, free = list_rows(ws, lay, block)

    if new_label and new_label.strip():
        known = {norm(c.label): c.row for c in cats}
        if norm(new_label) in known:
            row = known[norm(new_label)]
        else:
            if not free:
                raise LayoutError("Aucune ligne libre dans ce bloc pour créer une catégorie.")
            row = free[0]
            ref = [c.row for c in cats if c.row < row]
            if ref:
                _copy_row_style(ws, lay, ref[-1], row)
            ws.cell(row, lay.label_col).value = new_label.strip()
    else:
        row = next((c.row for c in cats if norm(c.label) == norm(label)), None)
        if row is None:
            raise LayoutError(f"Catégorie « {label} » introuvable dans le bloc {which}.")

    col = lay.month_cols[int(month)]
    cell = ws.cell(row, col)
    old, new = add_amount(cell, amount)
    detail_res = ecrire_detail(ws, lay, detail) if detail else None
    if make_backup:
        backup = save_with_backup(wb, path)
    else:                     # lot de plusieurs lignes : une seule sauvegarde pour tout le lot
        wb.save(path)
        backup = None
    return {"row": row, "col": col, "cell": cell.coordinate, "old": old, "new": new,
            "label": str(ws.cell(row, lay.label_col).value).strip(), "backup": backup, "detail": detail_res}


def log_operation(path, **fields):
    """Journal CSV (séparateur « ; », lisible dans Excel) pour garder une trace des saisies."""
    log = Path(path).parent / "journal_saisies.csv"
    new = not log.exists()
    with open(log, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        if new:
            w.writerow(["horodatage", *fields.keys()])
        w.writerow([dt.datetime.now().isoformat(timespec="seconds"), *fields.values()])


# ════════════════════════════════════════════════════════════════════════════
#  4. IMAGES : chargement, numérisation, OCR
# ════════════════════════════════════════════════════════════════════════════
MAX_COTE = 2600       # plus grand côté conservé (en pixels) : une photo de 12 Mpx ne sature plus la mémoire


def load_document(data: bytes, filename: str = "", max_pages: int = 5):
    """
    Retourne (aperçu PIL, texte intégré au PDF ou None, [images à OCRiser]).
    Un PDF « numérique » est lu directement (pas d'OCR, donc fiable à 100 %).
    """
    is_pdf = filename.lower().endswith(".pdf") or data[:5] == b"%PDF-"
    if not is_pdf:
        img = Image.open(io.BytesIO(data))
        try:
            img.draft("RGB", (MAX_COTE, MAX_COTE))        # JPEG : décodage directement à taille réduite (peu de mémoire)
        except Exception:  # noqa: BLE001
            pass
        img = ImageOps.exif_transpose(img).convert("RGB")
        img.thumbnail((MAX_COTE, MAX_COTE), Image.LANCZOS)
        return img, None, [img]

    try:
        import fitz  # PyMuPDF
    except ImportError as e:
        raise OcrUnavailable("Lecture des PDF : installez PyMuPDF (pip install pymupdf).") from e

    doc = fitz.open(stream=data, filetype="pdf")
    pages = [p for i, p in enumerate(doc) if i < max_pages]

    def render(page, dpi):
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72))
        return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

    text = "\n".join(p.get_text("text") for p in pages).strip()
    if len(text) >= 40:                       # PDF avec couche texte
        return render(pages[0], 110), text, []
    images = [render(p, 200) for p in pages]  # PDF scanné → OCR
    return images[0], None, images


def _order_points(pts):
    import numpy as np
    rect = np.zeros((4, 2), dtype="float32")
    s, d = pts.sum(axis=1), np.diff(pts, axis=1)
    rect[0], rect[2] = pts[s.argmin()], pts[s.argmax()]
    rect[1], rect[3] = pts[d.argmin()], pts[d.argmax()]
    return rect


def _page_quad(bgr):
    """
    Les 4 coins de la feuille, ou None. Le papier est clair et peu coloré, alors que la table, le bois
    ou le tissu autour sont plus sombres ou plus colorés : on isole le papier par sa couleur,
    ce qui marche beaucoup mieux que la recherche de contours sur un fond texturé.
    """
    import cv2
    import numpy as np

    h, w = bgr.shape[:2]
    k = min(1.0, 900.0 / max(h, w))
    small = cv2.resize(bgr, None, fx=k, fy=k, interpolation=cv2.INTER_AREA) if k < 1 else bgr.copy()
    hsv = cv2.cvtColor(small, cv2.COLOR_BGR2HSV)
    v = cv2.GaussianBlur(hsv[:, :, 2], (0, 0), 4)
    sat = cv2.GaussianBlur(hsv[:, :, 1], (0, 0), 4)
    seuil, _ = cv2.threshold(v, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    mask = ((v > seuil * 0.9) & (sat < 95)).astype(np.uint8) * 255
    taille = max(5, int(0.03 * max(small.shape[:2])) | 1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, np.ones((taille, taille), np.uint8))
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, np.ones((taille // 2 | 1, taille // 2 | 1), np.uint8))
    cnts, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cnts:
        return None
    grand = max(cnts, key=cv2.contourArea)
    if cv2.contourArea(grand) < 0.20 * small.shape[0] * small.shape[1]:
        return None
    hull = cv2.convexHull(grand)
    peri = cv2.arcLength(hull, True)
    quad = None
    for eps in (0.01, 0.02, 0.03, 0.05, 0.08):
        approx = cv2.approxPolyDP(hull, eps * peri, True)
        if len(approx) == 4:
            quad = approx.reshape(4, 2).astype("float32")
            break
    if quad is None:
        quad = cv2.boxPoints(cv2.minAreaRect(hull)).astype("float32")
    return quad / k


def _text_mask(bgr):
    """
    Pixels de texte imprimé (sombres par rapport à leur voisinage), débarrassés des petites taches :
    sert à vérifier qu'une découpe ne coupe aucun texte, et à découper autour du contenu.
    Retourne (masque, échelle).
    """
    import cv2
    import numpy as np

    h, w = bgr.shape[:2]
    k = min(1.0, 1000.0 / max(h, w))
    small = cv2.resize(bgr, None, fx=k, fy=k, interpolation=cv2.INTER_AREA) if k < 1 else bgr
    g = cv2.cvtColor(small, cv2.COLOR_BGR2GRAY)
    fond = cv2.medianBlur(g, 31)
    T = (cv2.subtract(fond, g) > 25).astype(np.uint8) * 255
    s = max(3, int(0.015 * max(g.shape)) | 1)
    fusion = cv2.dilate(T, cv2.getStructuringElement(cv2.MORPH_RECT, (s * 3, s)))
    n, labels, stats, _ = cv2.connectedComponentsWithStats(fusion)
    garde = np.zeros_like(T)
    if n > 1:
        # on garde les blocs de texte importants ; les petits motifs (broderie, rayures du support) sont ignorés
        plus_grand = int(stats[1:, cv2.CC_STAT_AREA].max())
        seuil = max(0.002 * g.shape[0] * g.shape[1], 0.08 * plus_grand)
        for i in range(1, n):
            if stats[i, cv2.CC_STAT_AREA] >= seuil:
                garde[labels == i] = 255
    return cv2.bitwise_and(T, garde), k


def _couverture(T, quad_petit):
    """Part du texte qui se trouve à l'intérieur de la découpe (1.0 = rien n'est coupé)."""
    import cv2
    import numpy as np

    total = int((T > 0).sum())
    if total == 0:
        return 1.0
    masque = np.zeros_like(T)
    cv2.fillConvexPoly(masque, quad_petit.astype("int32"), 255)
    return float(((T > 0) & (masque > 0)).sum()) / total


def scan_document(img: Image.Image, decouper: bool = True) -> Image.Image:
    """
    « Numérisation » : détecte la feuille, la découpe, redresse la perspective, puis uniformise l'éclairage
    (fond blanc, texte net) SANS tout passer en noir et blanc brutal : le texte fin reste lisible.
    Si la feuille n'est pas trouvée, l'image entière est conservée (éclairage uniformisé seulement).
    """
    import cv2
    import numpy as np

    bgr = cv2.cvtColor(np.array(ImageOps.exif_transpose(img).convert("RGB")), cv2.COLOR_RGB2BGR)
    H0, W0 = bgr.shape[:2]
    T, kt = _text_mask(bgr)
    page = _page_quad(bgr) if decouper else None
    if page is not None and T.any() and _couverture(T, page * kt) < 0.97:
        page = None                       # la découpe « par la couleur » couperait du texte : on l'abandonne
    if decouper and page is None and T.any():
        # secours : on découpe autour du CONTENU imprimé (marge de 2 %), sans jamais couper de texte
        x, y, wd, ht = cv2.boundingRect(cv2.findNonZero(T))
        m = 0.02 * max(wd, ht)
        x0, y0 = max(0, (x - m) / kt), max(0, (y - m) / kt)
        x1, y1 = min(W0, (x + wd + m) / kt), min(H0, (y + ht + m) / kt)
        if (x1 - x0) * (y1 - y0) < 0.92 * W0 * H0:
            page = np.array([[x0, y0], [x1, y0], [x1, y1], [x0, y1]], dtype="float32")
    if page is not None:
        tl, tr, br, bl = _order_points(page)
        wd = int(max(np.linalg.norm(br - bl), np.linalg.norm(tr - tl)))
        ht = int(max(np.linalg.norm(tr - br), np.linalg.norm(tl - bl)))
        if wd > 200 and ht > 200:
            ech = min(1.0, 3000.0 / max(wd, ht))
            wd, ht = int(wd * ech), int(ht * ech)
            dst = np.array([[0, 0], [wd - 1, 0], [wd - 1, ht - 1], [0, ht - 1]], dtype="float32")
            M = cv2.getPerspectiveTransform(np.array([tl, tr, br, bl], dtype="float32"), dst)
            bgr = cv2.warpPerspective(bgr, M, (wd, ht), flags=cv2.INTER_CUBIC)
            marge = max(2, int(0.006 * max(wd, ht)))          # petite marge : retire le bord sombre résiduel
            bgr = bgr[marge:-marge, marge:-marge]

    gris = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    # éclairage uniformisé : on divise par le fond estimé (ombres, dégradés) puis on étire le contraste
    h, w = gris.shape
    petit = cv2.resize(gris, None, fx=0.25, fy=0.25, interpolation=cv2.INTER_AREA)
    k = max(3, (min(petit.shape) // 12) | 1)
    fond = cv2.resize(cv2.medianBlur(petit, k), (w, h), interpolation=cv2.INTER_LINEAR)
    norm_ = cv2.divide(gris, np.maximum(fond, 1), scale=255)
    bas, haut = np.percentile(norm_, (1, 97))
    norm_ = np.clip((norm_.astype("float32") - bas) * 255.0 / max(haut - bas, 1.0), 0, 255).astype("uint8")
    return Image.fromarray(norm_).convert("RGB")


def image_quality(img: Image.Image) -> dict:
    """Netteté et taille de la photo : permet de prévenir « photo floue » avant de lire la facture."""
    import cv2
    import numpy as np

    g = cv2.cvtColor(np.array(img.convert("RGB")), cv2.COLOR_RGB2GRAY)
    h, w = g.shape
    k = 1000.0 / max(h, w)
    g2 = cv2.resize(g, None, fx=k, fy=k, interpolation=cv2.INTER_AREA)
    nettete = float(cv2.Laplacian(g2, cv2.CV_64F).var())
    return {"largeur": w, "hauteur": h, "nettete": round(nettete, 1),
            "floue": nettete < SEUIL_NETTETE or max(w, h) < 1000}


SEUIL_NETTETE = 150.0


def preprocess_for_ocr(img: Image.Image) -> Image.Image:
    """Niveaux de gris, agrandi vers ~2400 px sur le grand côté (le texte fin devient lisible pour l'OCR)."""
    img = ImageOps.exif_transpose(img).convert("L")
    w, h = img.size
    longest = max(w, h)
    cible = 1800
    if longest < cible:
        k = min(cible / longest, 3.0)
        img = img.resize((int(w * k), int(h * k)), Image.LANCZOS)
    elif longest > 2600:
        k = 2600 / longest
        img = img.resize((int(w * k), int(h * k)), Image.LANCZOS)
    return ImageOps.autocontrast(img, cutoff=1)


def _tesseract():
    try:
        import pytesseract
    except ImportError as e:
        raise OcrUnavailable("Installez pytesseract : pip install pytesseract") from e
    os.environ.setdefault("OMP_THREAD_LIMIT", "1")      # sur un petit serveur, plusieurs threads ralentissent Tesseract
    if os.environ.get("TESSERACT_CMD"):
        pytesseract.pytesseract.tesseract_cmd = os.environ["TESSERACT_CMD"]
    return pytesseract


def ocr_tesseract(img: Image.Image, psm: int = 6) -> str:
    pytesseract = _tesseract()
    img = preprocess_for_ocr(img)
    try:
        try:
            return pytesseract.image_to_string(img, lang="fra+eng", config=f"--psm {psm}", timeout=25)
        except pytesseract.TesseractError:      # pack de langue français absent
            return pytesseract.image_to_string(img, lang="eng", config=f"--psm {psm}", timeout=25)
    except pytesseract.TesseractNotFoundError as e:
        raise OcrUnavailable("Le programme Tesseract n'est pas installé (ou pas dans le PATH). "
                             "Voir les instructions d'installation.") from e


def ocr_donnees(img: Image.Image, psm: int = 6):
    """Une seule lecture Tesseract : le texte (ligne par ligne) ET la position de chaque mot."""
    pytesseract = _tesseract()
    prepared = preprocess_for_ocr(img)
    try:
        try:
            d = pytesseract.image_to_data(prepared, lang="fra+eng", config=f"--psm {psm}",
                                          output_type=pytesseract.Output.DICT, timeout=25)
        except pytesseract.TesseractError:          # pack de langue français absent
            d = pytesseract.image_to_data(prepared, lang="eng", config=f"--psm {psm}",
                                          output_type=pytesseract.Output.DICT, timeout=25)
    except pytesseract.TesseractNotFoundError as e:
        raise OcrUnavailable("Le programme Tesseract n'est pas installé (ou pas dans le PATH). "
                             "Voir les instructions d'installation.") from e
    mots, lignes = [], {}
    for k, t in enumerate(d["text"]):
        t = (t or "").strip()
        if not t:
            continue
        try:
            conf = float(d["conf"][k])
        except (TypeError, ValueError):
            conf = -1.0
        cle_ligne = (d["block_num"][k], d["par_num"][k], d["line_num"][k])
        mots.append({"t": t, "n": norm(t), "x": d["left"][k], "y": d["top"][k], "w": d["width"][k], "h": d["height"][k],
                     "c": conf, "l": cle_ligne})
        lignes.setdefault(cle_ligne, []).append(t)
    texte = "\n".join(" ".join(v) for _, v in sorted(lignes.items()))
    return texte, mots


def find_total_in_image(img: Image.Image):
    """Montant « Net à payer / Total TTC » repéré par sa position (voir find_total_in_words)."""
    try:
        _, mots = ocr_donnees(img)
    except OcrUnavailable:
        raise
    except Exception:  # noqa: BLE001
        return None
    return find_total_in_words(mots)


def find_total_in_words(mots):
    """
    Montant « Net à payer / Total TTC » repéré par sa POSITION : à droite du libellé, ou juste en dessous
    (cas des tableaux où le libellé est en en-tête de colonne). Retourne un float ou None.
    """
    for m in mots:
        m["cx"], m["cy"], m["x2"] = m["x"] + m["w"] / 2, m["y"] + m["h"] / 2, m["x"] + m["w"]
    montants = []
    for m in mots:
        t = m["t"].replace("€", "").replace("EUR", "").strip()
        mt = _AMOUNT.fullmatch(t)
        if mt and 0 < _to_float(mt) < 100_000:
            montants.append((m, _to_float(mt)))
    meilleurs = []                                           # (priorité, distance, montant)
    for k in (m for m in mots if m["n"].startswith("ttc")):
        ligne = sorted((m for m in mots if abs(m["cy"] - k["cy"]) < 0.6 * k["h"]), key=lambda m: m["x"])
        i = ligne.index(k)
        gauche, droite = i, i
        sep = re.compile(r"^[\W_]+$")               # tirets, barres... : séparent les colonnes d'un en-tête
        # un libellé = des mots serrés les uns contre les autres ; un espace large ou un tiret = autre colonne
        while (gauche > 0 and ligne[gauche]["x"] - ligne[gauche - 1]["x2"] < 1.1 * k["h"]
               and not sep.match(ligne[gauche - 1]["t"])):
            gauche -= 1
        while (droite < len(ligne) - 1 and ligne[droite + 1]["x"] - ligne[droite]["x2"] < 1.1 * k["h"]
               and not sep.match(ligne[droite + 1]["t"])):
            droite += 1
        libelle = norm(" ".join(m["t"] for m in ligne[gauche:droite + 1]))
        payer = bool(re.search(r"payer", libelle))                       # « net à payer », « total à payer »...
        if not (payer or re.search(r"total|montant|net", libelle)):
            continue
        prio = 0 if payer else 1
        x0, x1 = ligne[gauche]["x"], ligne[droite]["x2"]
        marge = 0.25 * (x1 - x0)
        for m, valeur in montants:
            if abs(m["cy"] - k["cy"]) < 0.6 * k["h"] and m["x"] >= x1 - 0.5 * k["h"] and m["x"] - x1 < 20 * k["h"]:
                meilleurs.append((prio, 0, m["x"] - x1, valeur))                  # même ligne, à droite
            elif payer and 0.5 * k["h"] < m["cy"] - k["cy"] < 12 * k["h"] and x0 - marge <= m["cx"] <= x1 + marge:
                # juste sous « Net à payer TTC » (libellé en en-tête de colonne). On ne le fait que pour « payer » :
                # « Total TTC » en tête d'un tableau d'articles est suivi des prix de chaque ligne, pas du total.
                meilleurs.append((prio, 1, m["cy"] - k["cy"], valeur))
    if not meilleurs:
        return None
    return sorted(meilleurs)[0][3]


_EASY = None


def ocr_easyocr(img: Image.Image) -> str:
    global _EASY
    try:
        import easyocr
        import numpy as np
    except ImportError as e:
        raise OcrUnavailable("Installez EasyOCR : pip install easyocr") from e
    if _EASY is None:
        _EASY = easyocr.Reader(["fr", "en"], gpu=False)
    arr = np.array(preprocess_for_ocr(img).convert("RGB"))
    return "\n".join(_EASY.readtext(arr, detail=0, paragraph=False))


def run_ocr(images, engine: str = "Tesseract", psm: int = 6) -> str:
    if engine.lower().startswith("easy"):
        return "\n".join(ocr_easyocr(im) for im in images)
    return "\n".join(ocr_tesseract(im, psm) for im in images)


def read_invoice(images, engine: str = "Tesseract") -> dict:
    """
    Lit la facture : texte, montant TTC le plus fiable et montants candidats.
    Une seule lecture par image (texte + positions). Le montant repéré par sa position
    (à côté ou sous « Net à payer TTC ») prime sur l'analyse du texte.
    """
    import time

    debut = time.time()
    total = None
    mots_premiere, entete = [], []
    if engine.lower().startswith("easy"):
        texte = run_ocr(images, engine)
    else:
        morceaux = []
        for i_im, im in enumerate(images):
            t_im, mots = ocr_donnees(im)
            morceaux.append(t_im)
            if total is None:
                total = find_total_in_words(mots)
            if i_im == 0:
                mots_premiere = mots
                try:
                    entete = lignes_fiables(ocr_donnees(im.crop((0, 0, im.size[0], max(1, int(im.size[1] * 0.28)))), psm=11)[1], 70)
                except Exception:  # noqa: BLE001      l'en-tête est un plus : jamais bloquant
                    entete = []
        texte = "\n".join(morceaux)
    if images and not engine.lower().startswith("easy") and time.time() - debut < 18 \
            and not find_invoice_number(texte + "\n" + "\n".join(entete)):
        try:                                        # n° de facture introuvable : on relit le haut de la page, agrandi
            haut = images[0].crop((0, 0, images[0].size[0], max(1, int(images[0].size[1] * 0.45))))
            if haut.size[0] < 2400:
                haut = haut.resize((2400, int(haut.size[1] * 2400 / haut.size[0])), Image.LANCZOS)
            entete = list(entete) + ocr_donnees(haut, psm=6)[0].splitlines()
        except Exception:  # noqa: BLE001
            pass
    meilleur, candidats = find_amount(texte)
    if meilleur is None and total is None and not engine.lower().startswith("easy") and time.time() - debut < 15:
        texte += "\n" + run_ocr(images, engine, psm=4)          # 2e passe seulement si la 1re a été rapide
        meilleur, candidats = find_amount(texte)
    somme = _total_par_somme([float(m.group(0).replace(",", ".")) for m in re.finditer(r"\d+[.,]\d{2}", texte)])
    final = total or meilleur
    source = "libellé Net à payer / TTC" if total else "texte"
    if final is None and mots_premiere:
        final = find_total_bas_droite(mots_premiere)
        source = "bas de page, à droite"
    if final is not None and somme and round(final, 2) in somme[1] and round(final, 2) != somme[0]:
        final = somme[0]                    # ex. 3,88 (TVA) ou 19,42 (HT) repéré à tort : 19,42 + 3,88 = 23,30
        source = "HT + TVA"
    toutes = [_to_float(m) for m in _AMOUNT.finditer(texte)]
    corrige = _eviter_tva_et_ht(final, texte.splitlines(), toutes)   # jamais un HT ni une TVA
    if corrige is not None and final is not None and round(corrige, 2) != round(final, 2):
        source = "HT + TVA (corrigé)"
    final = corrige if corrige is not None else final
    if final is not None:
        meilleur = final
        candidats = [final] + [c for c in candidats if c != final]
    return {"text": texte, "amount": meilleur, "amount_source": source, "candidates": candidats[:6],
            "mots": [{"t": m["t"], "c": m["c"], "l": m["l"]} for m in mots_premiere], "entete": entete}


# ════════════════════════════════════════════════════════════════════════════
#  5. EXTRACTION : montant TTC, date, catégorie
# ════════════════════════════════════════════════════════════════════════════
_AMOUNT = re.compile(
    r"(?<![\d.,/\-])(\d{1,3}(?:[ \u00a0.]\d{3})+|\d+)\s?[.,]\s?(\d{2})"
    r"(?!\d)(?![.,/\-]\d)(?!\s?%)"
)


def _to_float(m: re.Match) -> float:
    ent = re.sub(r"[ \u00a0.]", "", m.group(1))
    return float(f"{ent}.{m.group(2)}")


def _line_weight(n: str) -> int:
    """3 = ligne « TTC / net à payer », 2 = « total / montant / CB », 0 = à ignorer."""
    if "sous-total" in n or "sous total" in n:
        return 0
    if re.search(r"net a payer|total ttc|montant ttc|total a payer|montant a payer|"
                 r"a payer ttc|total general|total facture", n) or re.search(r"\bttc\b", n):
        return 3
    if re.search(r"\b(ht|tva|hors taxes?)\b", n):
        return 0
    if re.search(r"\b(total|montant|somme|carte|cb|especes|paye|reglement)\b", n):
        return 2
    return 0


def _total_par_somme(valeurs):
    """
    Cherche un montant égal à la somme de deux autres (HT + TVA = TTC, ou somme des lignes).
    Retourne (total, (a, b)) pour le plus grand total trouvé, ou None.
    """
    vals = sorted(round(v, 2) for v in valeurs if 0 < v < 100_000)
    meilleur = None
    for i, a in enumerate(vals):
        for b in vals[i + 1:]:
            c = round(a + b, 2)
            if c > b and c in vals and (meilleur is None or c > meilleur[0]):
                meilleur = (c, (a, b))
    return meilleur


def _montants_ht_tva(lines):
    """Montants des lignes libellées HT et TVA (jamais celles qui contiennent « TTC »)."""
    tva, ht = set(), set()
    for line in lines:
        n_l = norm(line)
        if "ttc" in n_l or "payer" in n_l:
            continue
        valeurs = {round(_to_float(m), 2) for m in _AMOUNT.finditer(line)}
        if re.search(r"tva|t\.v\.a|taxe", n_l):
            tva |= valeurs
        elif re.search(r"\bht\b|hors taxe", n_l):
            ht |= valeurs
    return ht, tva


_TAUX_TVA = (0.20, 0.10, 0.055, 0.021)


def _eviter_tva_et_ht(best, lines, every):
    """
    Garde-fou : le champ doit être le NET À PAYER TTC, jamais un montant HT ni de TVA.
    Règles : TTC >= HT > TVA, et HT + TVA = TTC. Si le montant retenu les viole, on le remplace par
    HT + TVA (retrouvé dans le texte, ou reconstitué si le taux de TVA est cohérent).
    """
    ht, tva = _montants_ht_tva(lines)
    tva = {v for v in tva if v > 0}
    if not ht and not tva:
        return best
    every = [round(a, 2) for a in every]
    reconstruit = None
    for h in sorted(ht, reverse=True):
        for v in sorted(tva, reverse=True):
            if v >= h:
                continue
            dans_texte = next((c for c in sorted(every, reverse=True) if abs(c - (h + v)) < 0.011), None)
            if dans_texte:
                reconstruit = dans_texte
            elif any(abs(h * t - v) <= 0.03 for t in _TAUX_TVA):
                reconstruit = round(h + v, 2)
            if reconstruit:
                break
        if reconstruit:
            break
    if best is None:
        return reconstruit
    b = round(best, 2)
    mauvais = ((b in tva) or (tva and b in ht) or any(b < h - 0.005 for h in ht) or any(b <= v for v in tva))
    if not mauvais:
        return best
    if reconstruit:
        return reconstruit
    plancher = max(ht | tva, default=0)
    autres = [a for a in every if a not in ht | tva and a > plancher]
    return max(autres) if autres else best


def find_total_bas_droite(mots):
    """
    Dernier recours quand aucun libellé « Net à payer / Total TTC » n'est lu : le montant le plus BAS de la page
    (à droite en cas d'égalité), hors lignes HT / TVA / sous-total.
    """
    if not mots:
        return None
    haut = max(m["y"] + m["h"] for m in mots)
    lignes = {}
    for m in mots:
        lignes.setdefault(m["l"], []).append(m)
    candidats = []
    for ms in lignes.values():
        n_l = norm(" ".join(m["t"] for m in ms))
        if re.search(r"\b(ht|tva)\b|hors taxe|sous.?total", n_l) and "ttc" not in n_l and "payer" not in n_l:
            continue
        for m in ms:
            t = m["t"].replace("€", "").replace("EUR", "").strip()
            mt = _AMOUNT.fullmatch(t)
            if mt and 0 < _to_float(mt) < 100_000 and m["y"] > 0.45 * haut:
                candidats.append((round(m["y"] / max(1, m["h"])), m["x"], _to_float(mt)))
    if not candidats:
        return None
    return max(candidats)[2]


def find_amount(text: str):
    """Retourne (meilleur montant TTC ou None, liste des montants candidats)."""
    lines = text.splitlines()
    weighted, every = [], []
    for i, line in enumerate(lines):
        found = [_to_float(m) for m in _AMOUNT.finditer(line)]
        every += found
        n_ligne = norm(line)
        w = _line_weight(n_ligne)
        entete = bool(re.search(r"\b(ht|tva)\b", n_ligne))          # en-tête de colonnes : la ligne suivante n'est pas fiable
        if w and not found and i + 1 < len(lines) and not entete:     # montant sur la ligne suivante
            found = [_to_float(m) for m in _AMOUNT.finditer(lines[i + 1])]
        weighted += [(w, a) for a in found if w]
    every = [a for a in every if 0 < a < 100_000]
    weighted = [(w, a) for w, a in weighted if 0 < a < 100_000]

    if weighted:
        best = max(weighted)[1]
    else:
        somme = _total_par_somme(every)
        best = somme[0] if somme else (max(every) if every else None)
    best = _eviter_tva_et_ht(best, lines, every)
    cands = []
    for a in ([best] if best else []) + [a for _, a in sorted(weighted, reverse=True)] + sorted(every, reverse=True):
        if a not in cands:
            cands.append(a)
    return best, cands[:6]


_MOIS_PREFIX = {"janv": 1, "jan": 1, "fevr": 2, "fev": 2, "mars": 3, "avr": 4, "mai": 5, "juin": 6,
                "juil": 7, "aout": 8, "sept": 9, "sep": 9, "oct": 10, "nov": 11, "dec": 12}
_D_NUM = re.compile(r"(?<!\d)(\d{1,2})\s?[/\-.]\s?(\d{1,2})\s?[/\-.]\s?(\d{4}|\d{2})(?!\d)")
_D_ISO = re.compile(r"(?<!\d)(\d{4})-(\d{2})-(\d{2})(?!\d)")
_D_TXT = re.compile(r"(?<!\d)(\d{1,2})(?:er)?\s+(janv|jan|fevr|fev|mars|avr|mai|juin|juil|aout|sept|sep|oct|nov|dec)"
                    r"[a-z]*\.?\s+(\d{4}|\d{2})(?!\d)")


def find_date(text: str, today: dt.date | None = None) -> dt.date | None:
    today = today or dt.date.today()
    best, best_score = None, None
    for line in text.splitlines():
        n = norm(line)
        cands = []
        for m in _D_NUM.finditer(n):
            cands.append((int(m.group(3)), int(m.group(2)), int(m.group(1)), len(m.group(3)) == 4))
        for m in _D_ISO.finditer(n):
            cands.append((int(m.group(1)), int(m.group(2)), int(m.group(3)), True))
        for m in _D_TXT.finditer(n):
            cands.append((int(m.group(3)), _MOIS_PREFIX[m.group(2)], int(m.group(1)), len(m.group(3)) == 4))
        for y, mo, d, four in cands:
            y = y + 2000 if y < 100 else y
            try:
                date = dt.date(y, mo, d)
            except ValueError:
                continue
            if not (2000 <= y <= today.year + 1):
                continue
            score = (2 if re.search(r"\bdate\b", n) and not re.search(r"echeance|limite|naissance", n) else 0) \
                - (2 if re.search(r"echeance|limite|avant le", n) else 0) + (1 if four else 0)
            if best_score is None or score > best_score:       # à score égal : la première trouvée
                best, best_score = date, score
    return best


_GENERIC = {"achat", "achats", "divers", "diverses", "fourniture", "fournitures", "frais", "total"}
_SYNONYMES = {"carburant": ["essence", "gazole", "diesel", "sans plomb", "sp95", "sp98", "e10"],
              "assurance": ["assurance", "cotisation assurance", "attestation"],
              "bureau": ["bureautique", "papeterie", "cartouche", "ramette", "stylo", "classeur", "agrafe"],
              "fumier": ["fumier", "compost"]}
# Lignes qui décrivent seulement le MOYEN DE PAIEMENT : elles ne disent rien sur la nature de l'achat
# (« carte bancaire » figure sur presque tous les tickets de caisse).
_PAIEMENT = re.compile(r"carte (bancaire|bleue|banc)|\b(cb|visa|mastercard|espece|especes|paiement|reglement|"
                       r"sans contact|contactless|monnaie rendue)\b")


def suggest_category(text: str, labels: list[str]) -> int | None:
    """Index de la catégorie la plus plausible d'après la nature de l'achat (ou None si aucun indice net)."""
    lignes = [l for l in text.splitlines() if not _PAIEMENT.search(norm(l))]
    t = norm(" ".join(lignes))
    best_i, best = None, 0
    for i, label in enumerate(labels):
        score = 0
        for tok in re.findall(r"[a-z0-9]{4,}", norm(label)):
            if tok in _GENERIC:
                continue
            if tok in t:
                score += 2
            score += sum(1 for syn in _SYNONYMES.get(tok, []) if syn in t)
        if score > best:
            best_i, best = i, score
    return best_i


def prepare_justificatif(data: bytes, filename: str, preview: Image.Image, max_pdf: int = 4_000_000):
    """
    Pièce justificative à conserver : (octets, type MIME).
    PDF d'origine s'il est léger, sinon image JPEG réduite (1600 px, qualité 80 : environ 300 Ko).
    """
    is_pdf = filename.lower().endswith(".pdf") or data[:5] == b"%PDF-"
    if is_pdf and len(data) <= max_pdf:
        return data, "application/pdf"
    img = ImageOps.exif_transpose(preview).convert("RGB")
    img.thumbnail((1600, 1600))
    out = io.BytesIO()
    img.save(out, "JPEG", quality=80, optimize=True)
    return out.getvalue(), "image/jpeg"


def excel_verrouille(path) -> bool:
    """
    Vrai si le fichier semble ouvert dans Excel (Windows interdit alors d'y écrire) :
    fichier verrou « ~$nom.xlsx » présent, ou écriture refusée par le système.
    """
    p = Path(path)
    for nom in ("~$" + p.name, "~$" + p.name[2:]):
        if (p.parent / nom).exists():
            return True
    try:
        with open(p, "ab"):
            pass
    except PermissionError:
        return True
    except OSError:
        return False
    return False


def maintenant_fr() -> dt.datetime:
    """Date et heure actuelles à Paris (le serveur en ligne est à l'heure UTC)."""
    try:
        from zoneinfo import ZoneInfo
        return dt.datetime.now(ZoneInfo("Europe/Paris"))
    except Exception:  # noqa: BLE001  (pas de base de fuseaux : on prend l'heure de l'appareil)
        return dt.datetime.now()


def reduire_photo(data: bytes, nom: str = "", max_cote: int = 1800, qualite: int = 85) -> bytes:
    """
    Photo reçue → JPEG réduit (1800 px max, EXIF appliqué) : mémoire du serveur et envois allégés.
    Les PDF sont rendus tels quels.
    """
    if nom.lower().endswith(".pdf") or data[:5] == b"%PDF-":
        return data
    img = Image.open(io.BytesIO(data))
    try:
        img.draft("RGB", (max_cote, max_cote))
    except Exception:  # noqa: BLE001
        pass
    img = ImageOps.exif_transpose(img).convert("RGB")
    img.thumbnail((max_cote, max_cote), Image.LANCZOS)
    sortie = io.BytesIO()
    img.save(sortie, "JPEG", quality=qualite, optimize=True)
    return sortie.getvalue()


# ════════════════════════════════════════════════════════════════════════════
#  TABLEAU DU BAS : « Fournisseur & Détail de l'achat · Désignation · N° Facture · Date · Mode de paiement ·
#  Montant · Justificatif »
# ════════════════════════════════════════════════════════════════════════════
# Plan des colonnes, par rapport à la colonne « Fournisseur » (A) : B = Désignation (C fusionnée, ignorée), D = N° facture,
# E = Date, F = Mode de paiement, G = Montant, H = Justificatif. Pour changer le plan, modifier ces décalages.
DECALAGES_DETAIL = {"fournisseur": 0, "objet": 1, "facture": 3, "date": 4, "paiement": 5, "montant": 6, "justificatif": 7}
_ENTETES_ATTENDUS = {"objet": ("designation", "objet"), "facture": ("facture",), "date": ("date",), "paiement": ("paiement",),
                     "montant": ("montant",), "justificatif": ("justificatif",)}


def trouver_tableau_detail(ws, lay: Layout):
    """(ligne d'en-tête, {rôle: n° de colonne}, avertissements) du tableau de détail, ou None s'il n'existe pas."""
    fin = min(ws.max_row, lay.sorties.total_row + 80)
    largeur = min(ws.max_column, 25)
    for r in range(lay.sorties.total_row + 1, fin + 1):
        c0 = next((c for c in range(1, largeur + 1) if norm(ws.cell(r, c).value).startswith("fournisseur")), None)
        if c0 is None:
            continue
        cols = {role: c0 + dec for role, dec in DECALAGES_DETAIL.items()}
        alertes = []
        for role, mots in _ENTETES_ATTENDUS.items():
            n = norm(ws.cell(r, cols[role]).value)
            if not any(m in n for m in mots):
                alertes.append(f"colonne {openpyxl_lettre(cols[role])} : « {role} » attendu, en-tête trouvé « {ws.cell(r, cols[role]).value or 'vide'} »")
        return r, cols, alertes
    return None


def openpyxl_lettre(n: int) -> str:
    from openpyxl.utils import get_column_letter
    return get_column_letter(n)


def ecrire_detail(ws, lay: Layout, d: dict):
    """
    Ajoute une ligne dans le tableau du bas (première ligne libre, en dessous des lignes déjà remplies).
    d : fournisseur, objet, facture, date (datetime.date), paiement, montant, justificatif_lien (chemin relatif ou adresse).
    Retourne {"ligne": n, "cellule_justificatif": "G52"} ou None si le tableau est introuvable.
    """
    from copy import copy as _copy
    from openpyxl.styles import Font

    trouve = trouver_tableau_detail(ws, lay)
    if not trouve:
        return None
    entete, cols, alertes = trouve
    cles = [cols[k] for k in ("fournisseur", "objet", "facture", "date", "paiement", "montant")]
    r = entete + 1
    while r < entete + 800 and any(ws.cell(r, c).value not in (None, "") for c in cles):
        r += 1
    # au-delà des lignes déjà dessinées : on reprend la mise en forme de la ligne du dessus
    if r > entete + 1 and ws.cell(r, min(cols.values())).border.left.style is None:
        for c in cols.values():
            ws.cell(r, c)._style = _copy(ws.cell(r - 1, c)._style)

    def poser(role, valeur, format_=None):
        if role not in cols or valeur in (None, ""):
            return
        cell = ws.cell(r, cols[role])
        try:
            format_avant = cell.number_format
            cell.value = valeur                  # openpyxl pose lui-même « aaaa-mm-jj » sur une date
            if format_ and format_avant == "General":
                cell.number_format = format_     # …on met le format français, sauf si la case a déjà le sien
        except AttributeError:          # cellule fusionnée : on la laisse
            pass

    poser("fournisseur", d.get("fournisseur"))
    poser("objet", d.get("objet"))
    poser("facture", d.get("facture"))
    poser("date", d.get("date"), "DD/MM/YYYY")
    poser("paiement", d.get("paiement"))
    poser("montant", round(float(d["montant"]), 2) if d.get("montant") is not None else None, '#,##0.00 "€"')
    cell_just = None
    if "justificatif" in cols and d.get("justificatif_lien"):
        cell = ws.cell(r, cols["justificatif"])
        try:
            cell.value = "📄 PDF"
            cell.hyperlink = d["justificatif_lien"]
            cell.font = Font(color="0563C1", underline="single")
            cell_just = cell.coordinate
        except AttributeError:
            pass
    return {"ligne": r, "cellule_justificatif": cell_just, "avertissement": " ; ".join(alertes)}


# ════════════════════════════════════════════════════════════════════════════
#  DÉTECTION DES CHAMPS D'UNE FACTURE ET PDF
# ════════════════════════════════════════════════════════════════════════════
_NON_FOURNISSEUR = re.compile(r"facture|ticket|re[cç]u|date|total|tva|siret|siren|\btel\b|fax|www|http|@|rue |avenue|chemin|route|"
                              r"\bcb\b|merci|client|duplicata|page |bon de|devis|commande|livraison|r[eè]glement|iban|bic\b|"
                              r"capital|r\.?c\.?s|n° tva|adresse|vendeur|caisse|op[eé]rateur")
_FORMES_JURIDIQUES = r"(?:SAS|SARL|SA|EURL|SASU|SNC|SCOP|SCI)"
_PAS_UN_COMMERCE = re.compile(r"banque|bank|credit|cr[eé]dit|caisse|assurance|epargne|[eé]pargne|palatine|mutuelle|poste")
_FOURNISSEURS_GENERIQUES = {"orange", "free", "gmail", "hotmail", "outlook", "yahoo", "wanadoo", "sfr", "laposte", "live", "icloud"}
_MOTS_DE_FACTURE = set("""facture total montant designation client date tva paiement carte bancaire net payer prix unitaire quantite
reference adresse france avenue route siege social capital euros tel fax email banque iban bic siret siren code article remise
base taux echeance reglement mode ticket caisse merci votre vous pour les des une avec sans par sur dans commande livraison
conditions generales vente page duplicata original exemplaire numero compte contact horaires ouvert""".split())


def lignes_fiables(mots, seuil: float = 55.0):
    """Lignes de texte reconstruites avec les seuls mots lus de façon fiable (confiance Tesseract ≥ seuil)."""
    lignes = {}
    for m in mots:
        if m.get("c", 100) >= seuil:
            lignes.setdefault(m["l"], []).append(m["t"])
    return [" ".join(v) for _, v in sorted(lignes.items())]


def _ligne_propre(ligne: str) -> bool:
    """Vrai si la ligne ressemble à du texte lu correctement (et non à du bruit venant d'un logo)."""
    mots = re.findall(r"[A-Za-zÀ-ÿ0-9&'\-]+", ligne)
    lettres = len(re.findall(r"[A-Za-zÀ-ÿ]", ligne))
    if not mots or lettres < 3 or lettres < 0.65 * len(re.sub(r"\s", "", ligne)):
        return False
    courts = sum(len(re.sub(r"[^A-Za-zÀ-ÿ]", "", m)) <= 2 for m in mots)
    return courts <= max(1, len(mots) // 4) and any(len(m) >= 4 for m in mots)


def _nom_propre(nom: str) -> str:
    nom = re.sub(r"(?i)^(si[eè]ge social|soci[eé]t[eé]|raison sociale|[eé]tablissements?|ets|ste)\s*[:\-]?\s*", "", nom)
    nom = re.sub(r"\s+", " ", nom).strip(" -:;,.|_*#")
    nom = re.sub(r"\s+" + _FORMES_JURIDIQUES + r"$", "", nom)
    return (nom.title() if nom.isupper() else nom)[:60]


def find_supplier(text: str, fiables=None, entete=None) -> str:
    """
    Nom du fournisseur : raison sociale (« VAUDAUX SA »), sinon nom du site web / de l'e-mail, sinon première ligne
    de l'en-tête lue de façon fiable. Les lignes illisibles (logo mal lu) sont écartées.
    """
    lignes = [re.sub(r"\s+", " ", l).strip() for l in (fiables if fiables else text.splitlines())]
    lignes += [re.sub(r"\s+", " ", l).strip() for l in (entete or [])]
    tout = "\n".join(lignes)
    # 1) raison sociale suivie de sa forme juridique
    for m in re.finditer(r"([A-ZÀ-Ý][A-Za-zÀ-ÿ&'\-]+(?:\s+[A-Za-zÀ-ÿ&'\-]+){0,3}?)\s+" + _FORMES_JURIDIQUES + r"\b", tout):
        nom = _nom_propre(m.group(1))
        n = norm(nom)
        if len(re.findall(r"[A-Za-zÀ-ÿ]", nom)) >= 4 and not _PAS_UN_COMMERCE.search(n) and not _NON_FOURNISSEUR.search(n) \
                and _ligne_propre(nom):
            return nom
    # 2) site web ou adresse e-mail du fournisseur
    for m in re.finditer(r"(?:www\.|@)([a-z0-9][a-z0-9\-]{2,24})\.(?:fr|com|net|org|eu|io)\b", tout.lower()):
        if m.group(1) not in _FOURNISSEURS_GENERIQUES and not _PAS_UN_COMMERCE.search(m.group(1)):
            return m.group(1).replace("-", " ").title()
    # 3) première ligne propre de l'en-tête
    for l in lignes[:14]:
        brut = l.strip(" -—_|.:;,*#")
        if _ligne_propre(brut) and not re.search(r"\d{4,}", brut) and not _NON_FOURNISSEUR.search(norm(brut)):
            return _nom_propre(brut)
    return ""


_NON_DESIGNATION = re.compile(r"\bqt[eé]?\b|quantit|prix|p\.?u\b|unit|montant|total|tva|taux|remise|r[eé]f[eé]rence|\bht\b|\bttc\b|"
                              r"d[eé]signation|libell|description|\bcode\b|\bn°|tel\b|www|siret|iban|merci|remercie|"
                              r"banque|capital|siege|conditions|garantie|\brcs\b|exemplaire|original|duplicata")
_ADMIN = re.compile(r"facture|client|repr[eé]sentant|adresse|\btel\b|t[eé]l[eé]phone|siret|\bdate\b|commande|livraison|\bpage\b|"
                    r"mail|www|code postal|\b\d{5}\b|n[°º]|r[eé]f[eé]rence|vendeur|[eé]ch[eé]ance|mode de r[eè]glement")
_FIN_TABLEAU = re.compile(r"total|montant ht|net a payer|sous-total|remercie|r[eè]glement|echeance|merci|tva|banque|capital|"
                          r"siege|conditions|garantie|iban|\brcs\b")
_OBJET_FORT = re.compile(r"r[eé]paration|remplacement|concernant|entretien|fourniture|achat|location|abonnement|cotisation|"
                         r"intervention|installation|dépannage|depannage|maintenance", re.I)
# mots courants des factures de jardin / bricolage : corrige les fautes de lecture (« concemant » → « concernant »)
_VOCABULAIRE = """réparation remplacement concernant entretien fourniture fournitures achat pièces pièce pneu pneus roue roues
autoportée autoportee tondeuse taille-haie débroussailleuse debroussailleuse huile essence carburant lame lames courroie bougie
filtre graines plants terreau engrais outils outil arrosage arrosoir tuyau grillage clôture cloture peinture ciment béton beton
bois vis clous main-d'œuvre mise place petites standard espace vert mèche meche constat intervention installation maintenance
location abonnement cotisation cadenas serrure clé clés boîte boite lettres porte portail poteau planche planches sable gravier
gants sécateur secateur râteau rateau pelle bêche beche brouette semences semis arbres arbuste arbustes haie""".split()


def _mots_de(s: str) -> set:
    return set(re.findall(r"[a-z]{3,}", norm(s)))


def _recoupe(mots: set, reference: set) -> float:
    """Part des mots de la ligne qui se retrouvent (à une faute de lecture près) dans la référence."""
    import difflib
    if not mots:
        return 0.0
    trouves = sum(any(m == r or difflib.SequenceMatcher(None, m, r).ratio() >= 0.8 for r in reference) for m in mots)
    return trouves / len(mots)


def _corriger_mots(texte: str) -> str:
    import difflib
    base = {norm(v): v for v in _VOCABULAIRE}
    sortie = []
    for mot in texte.split():
        brut = re.sub(r"[^A-Za-zÀ-ÿ'\-]", "", mot)
        if len(brut) >= 6 and norm(brut) not in base:
            proche = difflib.get_close_matches(norm(brut), list(base), n=1, cutoff=0.84)
            if proche:
                v = base[proche[0]]
                mot = mot.replace(brut, v.upper() if brut.isupper() else (v.capitalize() if brut[0].isupper() else v))
        sortie.append(mot)
    return " ".join(sortie)


def _couper(texte: str, longueur: int) -> str:
    if len(texte) <= longueur:
        return texte
    return texte[:longueur].rsplit(" ", 1)[0]


def find_designation(text: str, fiables=None, entete=None, fournisseur: str = "", max_lignes: int = 2) -> str:
    """
    Objet réel de la dépense. Écarte l'en-tête (enseigne, slogan), les adresses, les totaux, les mentions légales
    et les lignes qui répètent une ligne déjà prise ; corrige les fautes de lecture sur les mots courants.
    """
    lignes = [re.sub(r"\s+", " ", l).strip() for l in (fiables if fiables else text.splitlines())]
    debut = next((i for i, l in enumerate(lignes)
                  if re.search(r"d[eé]signation|libell[eé]|description", norm(l)) and not re.search(r"\d{5,}", l)), None)
    if debut is not None:
        depart = debut + 1
    else:                                           # sans en-tête de tableau : après la dernière ligne administrative du haut
        fin_idx = next((i for i, l in enumerate(lignes) if i > 2 and _FIN_TABLEAU.search(norm(l))), len(lignes))
        admin = [i for i in range(min(fin_idx, 35)) if _ADMIN.search(norm(lignes[i]))]     # l'adresse du pied de page ne compte pas
        depart = admin[-1] + 1 if admin else min(6, len(lignes))
    # mots de l'en-tête : enseigne, slogan, nom du fournisseur → jamais dans la désignation
    reference = set()
    for l in lignes[:min(depart, 8)] + list(entete or [])[:3] + [fournisseur]:      # enseigne et slogan : tout en haut seulement
        reference |= _mots_de(l)
    trouvees, vus = [], []
    for l in lignes[depart:]:
        n = norm(l)
        if _FIN_TABLEAU.search(n) and trouvees:
            break
        propre = re.sub(r"\b\d+(?:[.,]\d+)?\s*(?:€|eur|euros?|%)?", " ", l)
        propre = re.sub(r"[^A-Za-zÀ-ÿ'\- ]", " ", propre)
        propre = re.sub(r"(?<![A-Za-zÀ-ÿ'\-])[A-Za-zÀ-ÿ](?![A-Za-zÀ-ÿ'\-])", " ", propre)     # « 10 L » → reste « L » : écarté
        propre = re.sub(r"\s+", " ", propre).strip(" -'")
        if len(propre) < 5 or not _ligne_propre(propre) or _NON_DESIGNATION.search(norm(propre)) or _NON_FOURNISSEUR.search(norm(propre)):
            continue
        mots_l = _mots_de(propre)
        if mots_l and _recoupe(mots_l, reference) >= 0.6:
            continue                                # slogan / enseigne repris de l'en-tête
        if mots_l and any(mots_l <= v for v in vus):
            continue                                # même ligne répétée en majuscules
        vus.append(mots_l)
        trouvees.append(_corriger_mots(_couper(propre, 60)))
        if _OBJET_FORT.search(propre) or len(trouvees) >= max_lignes:
            break                                   # une ligne qui décrit l'objet suffit
    return " ; ".join(trouvees)[:130]


_MOT_FACTURE = r"f[a4][cg]t[uv]r[eéf]?"
_ETIQUETTE_FACTURE = r"\b(?:" + _MOT_FACTURE + r"|fact\.?|invoice|ticket|re[cç]u|avoir|pi[eè]ce|document|doc\.)"
_EXCLUS_NUMERO = re.compile(r"tva|siret|siren|iban|\bt[eé]l|fax|rcs|capital|\bbic\b|www|@|portable|mobile")


def _chiffres_ocr(jeton: str) -> str:
    """Corrige les confusions de lecture d'un numéro presque entièrement numérique (O→0, I/l→1, S→5, B→8)."""
    if jeton and sum(c.isdigit() for c in jeton) >= 0.6 * len(jeton):
        return jeton.translate(str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1", "|": "1", "S": "5", "B": "8", "Z": "2"}))
    return jeton


def _compacter_chiffres(text: str) -> str:
    """« 2026 0906 0335 » → « 202609060335 » (groupes de chiffres séparés par des espaces ou des points, au moins 9 chiffres)."""
    def joindre(m):
        return m.group(1) + re.sub(r"[ .]", "", m.group(2)) if len(re.sub(r"\D", "", m.group(0))) >= 9 else m.group(0)
    return re.sub(r"(?<![\d,.])(\d{2,5})((?:[ .]\d{2,5}){1,4})(?![\d,])", joindre, text)


def find_invoice_number(text: str) -> str:
    """
    Numéro de facture. Essaie dans l'ordre : « FACTURE N° … » (fautes de lecture tolérées : FACTURF, N“, O pour 0, chiffres
    séparés par des espaces), numéro sous l'en-tête d'un tableau, « N° … » près d'un mot clé, puis, à défaut d'étiquette lisible,
    un long numéro commençant par l'année (202609060335) en haut de la page.
    """
    def valide(jeton: str) -> bool:
        jeton = jeton.strip("-/")
        return (sum(ch.isdigit() for ch in jeton) >= 3 and norm(jeton) not in ("date", "numero", "client")
                and not re.fullmatch(r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", jeton))

    signe = r"(?:n\s*[°º\"“”'*®]?|num[eé]ro|#)"
    jeton_re = r"([A-Z0-9][A-Z0-9\-/]{3,24})"
    modeles = [r"n\s*[°º\"“”'*®]?\s*(?:de\s*)?" + _MOT_FACTURE + r"\s*[:#]?\s*" + jeton_re,
               _ETIQUETTE_FACTURE + r"\s*" + signe + r"?\s*[:#]?\s*" + jeton_re,
               _ETIQUETTE_FACTURE + r"\s*" + signe + r"?\s*[:#]?\s*\n\s*" + jeton_re]     # numéro à la ligne suivante
    compact = _compacter_chiffres(text)
    for variante in (compact, text):                # d'abord les chiffres regroupés : « 2026 0906 0335 » est un seul numéro
        for modele in modeles:
            for m in re.finditer(modele, variante, re.I):
                j = _chiffres_ocr(m.group(1)).strip("-/")
                if valide(j):
                    return j
        lignes = variante.splitlines()
        for i, l in enumerate(lignes):             # en-tête de tableau « N° FACTURE | DATE | CLIENT », valeurs dessous
            if re.search(r"facture|invoice", l, re.I) and re.search(r"\bdate\b", l, re.I) and not re.search(r"\d{5,}", l):
                for suivante in lignes[i + 1:i + 3]:
                    for jeton in re.findall(r"[A-Z0-9][A-Z0-9\-/]{4,24}", suivante, re.I):
                        j = _chiffres_ocr(jeton)
                        if valide(j) and sum(ch.isdigit() for ch in j) >= 5:
                            return j.strip("-/")
        for l in lignes:                           # « N° 00451 » sur une ligne qui parle de ticket / facture / caisse / commande
            if re.search(r"ticket|caisse|facture|re[cç]u|commande", l, re.I):
                m = re.search(r"n\s*[°º\"“”'*®]\s*[:#]?\s*([A-Z0-9][A-Z0-9\-/]{2,24})", l, re.I)
                if m and valide(_chiffres_ocr(m.group(1))):
                    return _chiffres_ocr(m.group(1)).strip("-/")
    # sans étiquette lisible : long numéro en haut de page (année + chiffres), hors téléphones, SIRET, TVA, IBAN
    for l in compact.splitlines()[:20]:
        if _EXCLUS_NUMERO.search(norm(l)):
            continue
        for m in re.finditer(r"(?<![\d,.])(20\d{2}\d{6,9})(?![\d,.])", l):
            return m.group(1)
    for l in compact.splitlines()[:20]:
        if _EXCLUS_NUMERO.search(norm(l)):
            continue
        for m in re.finditer(r"(?<![\d,.])([1-9]\d{8,12})(?![\d,.])", l):
            return m.group(1)
    return ""


def extraire_champs(lecture: dict | None = None, texte: str = "") -> dict:
    """Fournisseur, désignation, n° de facture et mode de paiement d'une facture lue (résultat de read_invoice, ou texte seul)."""
    if lecture:
        texte = lecture.get("text", "")
        fiables = lignes_fiables(lecture.get("mots", []), 60) or None
        entete = lecture.get("entete") or []
    else:
        fiables, entete = None, []
    fournisseur = find_supplier(texte, fiables, entete)
    return {"fournisseur": fournisseur, "objet": find_designation(texte, fiables, entete, fournisseur),
            "numero": find_invoice_number(texte + "\n" + "\n".join(entete)), "paiement": find_payment(texte)}


def find_payment(text: str) -> str:
    """Mode de paiement : « CB », « Chèque » ou « Espèces » (par défaut)."""
    n = norm(text)
    if re.search(r"\b(cheque|chq)\b", n):
        return "Chèque"
    if re.search(r"carte bancaire|carte bleue|\bcb\b|visa|mastercard|contactless|sans contact", n):
        return "CB"
    return "Espèces"


def en_pdf(data: bytes, nom: str = "") -> bytes:
    """Photo → PDF d'une page A4 (un PDF reçu est conservé tel quel)."""
    if nom.lower().endswith(".pdf") or data[:5] == b"%PDF-":
        return data
    img = Image.open(io.BytesIO(data))
    try:
        img.draft("RGB", (2600, 2600))
    except Exception:  # noqa: BLE001
        pass
    img = ImageOps.exif_transpose(img).convert("RGB")
    img.thumbnail((2400, 2400), Image.LANCZOS)
    w, h = img.size
    page_l, page_h = (8.27, 11.69) if h >= w else (11.69, 8.27)       # A4, portrait ou paysage
    resolution = max(w / page_l, h / page_h)
    sortie = io.BytesIO()
    img.save(sortie, "PDF", resolution=resolution, quality=85)
    return sortie.getvalue()


def nom_fichier_sur(*morceaux, max_len: int = 70) -> str:
    """Nom de fichier sans accent ni caractère spécial (utilisable dans un lien Excel)."""
    brut = "_".join(norm(m) for m in morceaux if m)
    brut = re.sub(r"[^a-z0-9]+", "-", brut).strip("-")
    return brut[:max_len] or "justificatif"
