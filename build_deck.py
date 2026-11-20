import os
import csv
import glob
import openpyxl
from PIL import Image, ImageDraw, ImageFont, ImageOps

# ==========================================
# 1. KONFIGURACJA ŚCIEŻEK I WYMIARÓW (300 DPI)
# ==========================================
BASE_DIR = r"C:\Users\zarow\Documents\nandeck\Nietoperze"
OUTPUT_DIR = os.path.join(BASE_DIR, "output_cards")
PDF_OUTPUT = os.path.join(BASE_DIR, "talia_nietoperzy_druk.pdf")

# 80 mm x 120 mm przy 300 DPI -> 945 x 1417 px
CARD_W, CARD_H = 945, 1417
os.makedirs(OUTPUT_DIR, exist_ok=True)

xlsx_files = glob.glob(os.path.join(BASE_DIR, "*.xlsx"))
csv_files = glob.glob(os.path.join(BASE_DIR, "*.csv"))

if xlsx_files:
    DATA_PATH = next(iter(xlsx_files))
elif csv_files:
    DATA_PATH = next(iter(csv_files))
else:
    raise FileNotFoundError(f"Brak pliku .xlsx ani .csv w folderze: {BASE_DIR}")

print(f"[INFO] Wczytuję dane z: {os.path.basename(DATA_PATH)}")

# Czcionki systemowe Windows z pełnym wsparciem polskich znaków
FONT_REGULAR = r"C:\Windows\Fonts\segoeui.ttf"
FONT_BOLD = r"C:\Windows\Fonts\segoeuib.ttf"
FONT_ITALIC = r"C:\Windows\Fonts\segoeuii.ttf"

font_title = ImageFont.truetype(FONT_BOLD, 37)
font_latin = ImageFont.truetype(FONT_ITALIC, 29)
font_meta = ImageFont.truetype(FONT_BOLD, 23)
font_desc = ImageFont.truetype(FONT_REGULAR, 27)
font_skill = ImageFont.truetype(FONT_BOLD, 24)
font_plus = ImageFont.truetype(FONT_BOLD, 36)
font_stat = ImageFont.truetype(FONT_BOLD, 42)


# ==========================================
# 2. FUNKCJE POMOCNICZE (TYPOGRAFIA DTP)
# ==========================================
def draw_text_with_shadow(draw, pos, text, font, fill, shadow="#080A0F", stroke_w=2):
    """Rysuje tekst z cieniem i ciemnym obrysem dla maksymalnego kontrastu w druku."""
    x, y = pos
    draw.text(
        (x + 2, y + 2),
        text,
        font=font,
        fill=shadow,
        stroke_width=stroke_w,
        stroke_fill=shadow,
    )
    draw.text(
        (x, y),
        text,
        font=font,
        fill=fill,
        stroke_width=stroke_w,
        stroke_fill=shadow,
    )


def wrap_text_lines(text, font, max_w):
    """Dzieli tekst na listę linii dopasowanych do szerokości max_w."""
    if not text:
        return []
    lines = []
    for paragraph in str(text).split("\n"):
        words = paragraph.split(" ")
        current_line = ""
        for word in words:
            test_line = f"{current_line} {word}".strip()
            if font.getlength(test_line) <= max_w:
                current_line = test_line
            else:
                if current_line:
                    lines.append(current_line)
                current_line = word
        if current_line:
            lines.append(current_line)
    return lines


def get_line_height(font, spacing=5):
    """Zwraca wysokość pojedynczej linii wraz z interlinią."""
    _, top_y, _, bottom_y = font.getbbox("Ag")
    return (bottom_y - top_y) + spacing


def draw_wrapped_text(draw, text, box, font, fill, align="center", line_spacing=5, stroke_w=2):
    """Zawija tekst po słowach wewnątrz (x1, y1, x2, y2) i centruje w pionie."""
    if not text:
        _, y1, _, _ = box
        return y1

    x1, y1, x2, y2 = box
    max_w = x2 - x1
    lines = wrap_text_lines(text, font, max_w)
    line_h = get_line_height(font, line_spacing)
    total_h = len(lines) * line_h

    curr_y = y1 + max(0, (y2 - y1 - total_h) // 2) if align == "center" else y1

    for line in lines:
        if curr_y + line_h > y2 + 6:
            break
        line_w = font.getlength(line)
        curr_x = x1 + (max_w - line_w) // 2 if align == "center" else x1
        draw_text_with_shadow(draw, (curr_x, curr_y), line, font, fill, stroke_w=stroke_w)
        curr_y += line_h

    return curr_y


def load_clean_path(raw_val):
    """Czyści wielokrotne cudzysłowy wokół ścieżek z arkusza."""
    if not raw_val:
        return ""
    return str(raw_val).strip().strip('"').strip("'")


def load_rows(path):
    """Wczytuje wiersze jako listę słowników z .xlsx lub .csv."""
    if path.lower().endswith(".xlsx"):
        wb = openpyxl.load_workbook(path, data_only=True)
        sheet = wb.active
        all_rows = list(sheet.iter_rows(values_only=True))

        header_idx = 0
        for idx, r in enumerate(all_rows):
            row_strs = [str(c).strip() for c in r if c is not None]
            if "Region" in row_strs or "Cechy" in row_strs:
                header_idx = idx
                break

        headers = [str(h).strip() if h is not None else "" for h in all_rows[header_idx]]
        records = []
        for r in all_rows[header_idx + 1:]:
            if not any(r):
                continue
            records.append(dict(zip(headers, r)))
        return records
    else:
        for enc in ("utf-8-sig", "utf-8", "cp1250"):
            try:
                with open(path, mode="r", encoding=enc, newline="") as f:
                    sample = f.read(2048)
                    f.seek(0)
                    delim = ";" if sample.count(";") > sample.count(",") else ","
                    reader = csv.DictReader(f, delimiter=delim)
                    return [{k.strip(): v for k, v in row.items() if k} for row in reader]
            except UnicodeDecodeError:
                continue
        raise ValueError("Nie udało się rozpoznać kodowania pliku CSV.")


def get_col(row, *candidates):
    """Pobiera wartość kolumny niezależnie od wariantu nazwy nagłówka."""
    for name in candidates:
        if name in row and row[name] not in (None, ""):
            return row[name]
    return ""


# ==========================================
# 3. GŁÓWNA PĘTLA GENERUJĄCA TALIĘ
# ==========================================
records = load_rows(DATA_PATH)
rendered_cards = []

for row_num, row in enumerate(records, start=1):
    nazwa_pl = get_col(row, "NAZWA_PL", "Nazwa nietoperza (PL)")
    if not nazwa_pl:
        continue

    nazwa_lac = get_col(row, "NAZWA_LAC", "Nazwa nietoperza (łac.)")
    region = get_col(row, "Region")
    rozmiar = get_col(row, "Rozmiar")
    cechy = get_col(row, "Cechy")
    raw_skill = str(get_col(row, "UMIEJETNOSC", "Umiejętność")).strip()
    umiejetnosc = raw_skill[:1].upper() + raw_skill[1:] if raw_skill else ""

    hp = int(float(get_col(row, "HP", "Punkty życia") or 0))
    atk_1 = int(float(get_col(row, "atak_1") or 0))
    atk_2 = int(float(get_col(row, "atak_2") or 0))

    art_path = load_clean_path(get_col(row, "Zdjęcie", "ZDJECIE"))
    tpl_path = load_clean_path(get_col(row, "template"))
    hp_icon_path = load_clean_path(get_col(row, "zycie"))
    atk_icon_path = load_clean_path(get_col(row, "atak"))

    card = Image.new("RGBA", (CARD_W, CARD_H), (15, 18, 25, 255))

    # Warstwa 1: Ilustracja nietoperza
    ax1, ay1, ax2, ay2 = 55, 75, 890, 710
    if os.path.exists(art_path):
        art_img = Image.open(art_path).convert("RGBA")
        art_fitted = ImageOps.fit(
            art_img,
            (ax2 - ax1, ay2 - ay1),
            method=Image.Resampling.LANCZOS,
        )
        card.paste(art_fitted, (ax1, ay1))

    # Warstwa 2: Ramka karty (Template PNG z przezroczystością)
    tpl_img = Image.open(tpl_path).convert("RGBA").resize(
        (CARD_W, CARD_H), Image.Resampling.LANCZOS
    )
    card = Image.alpha_composite(card, tpl_img)

    draw = ImageDraw.Draw(card)

    # Warstwa 3: Belka tytułu — zwarty układ z mniejszym odstępem między nazwą PL a ŁAC
    tx1, ty1, tx2, ty2 = 85, 710, 860, 810
    t_max_w = tx2 - tx1
    pl_lines = wrap_text_lines(nazwa_pl, font_title, t_max_w)
    lac_lines = wrap_text_lines(nazwa_lac, font_latin, t_max_w)

    pl_lh = get_line_height(font_title, spacing=2)
    lac_lh = get_line_height(font_latin, spacing=2)
    title_gap = 2  # Zmniejszony odstęp między nazwą PL a łacińską

    total_title_h = len(pl_lines) * pl_lh + (title_gap if lac_lines else 0) + len(lac_lines) * lac_lh
    curr_ty = ty1 + max(0, (ty2 - ty1 - total_title_h) // 2)

    for line in pl_lines:
        lw = font_title.getlength(line)
        draw_text_with_shadow(draw, (tx1 + (t_max_w - lw) // 2, curr_ty), line, font_title, "#FFFFFF", stroke_w=2)
        curr_ty += pl_lh

    curr_ty += title_gap

    for line in lac_lines:
        lw = font_latin.getlength(line)
        draw_text_with_shadow(draw, (tx1 + (t_max_w - lw) // 2, curr_ty), line, font_latin, "#FFF4D6", stroke_w=2)
        curr_ty += lac_lh

    # Warstwa 4: Pole opisu — pełne wyśrodkowanie w poziomie i w pionie (center-center)
    dx1, dy1, dx2, dy2 = 92, 825, 853, 1068
    d_max_w = dx2 - dx1
    meta_text = f"{region}  •  {rozmiar}"

    meta_lines = wrap_text_lines(meta_text, font_meta, d_max_w)
    desc_lines = wrap_text_lines(cechy, font_desc, d_max_w)

    meta_lh = get_line_height(font_meta, spacing=4)
    desc_lh = get_line_height(font_desc, spacing=5)
    meta_desc_gap = 12

    total_desc_h = len(meta_lines) * meta_lh + (meta_desc_gap if desc_lines else 0) + len(desc_lines) * desc_lh
    curr_dy = dy1 + max(0, (dy2 - dy1 - total_desc_h) // 2)

    for line in meta_lines:
        lw = font_meta.getlength(line)
        draw_text_with_shadow(draw, (dx1 + (d_max_w - lw) // 2, curr_dy), line, font_meta, "#8CE0FF", stroke_w=2)
        curr_dy += meta_lh

    curr_dy += meta_desc_gap

    for line in desc_lines:
        if curr_dy + desc_lh > dy2 + 6:
            break
        lw = font_desc.getlength(line)
        draw_text_with_shadow(draw, (dx1 + (d_max_w - lw) // 2, curr_dy), line, font_desc, "#FFFFFF", stroke_w=2)
        curr_dy += desc_lh
            
    # Warstwa 5: Pasek statystyk — HP wyśrodkowane w lewej połówce, Atak wyśrodkowany w prawej połówce
    ICON_SIZE = 50
    ICON_GAP = 5
    STAT_Y = 1094

    LEFT_HALF_CENTER = (92 + CARD_W // 2) // 2       # Środek lewej połówki (~282 px)
    RIGHT_HALF_CENTER = (CARD_W // 2 + 853) // 2     # Środek prawej połówki (~662 px)

    # 1. Blok HP wyśrodkowany na osi LEFT_HALF_CENTER
    if os.path.exists(hp_icon_path) and hp > 0:
        hp_icon = Image.open(hp_icon_path).convert("RGBA").resize(
            (ICON_SIZE, ICON_SIZE), Image.Resampling.LANCZOS
        )
        hp_str = str(hp)
        bx0, by0, bx1, by1 = font_stat.getbbox(hp_str)
        hp_text_w = bx1 - bx0
        hp_text_h = by1 - by0
        hp_gap = 10

        hp_block_w = hp_text_w + hp_gap + ICON_SIZE
        hp_start_x = LEFT_HALF_CENTER - hp_block_w // 2
        text_y = STAT_Y + (ICON_SIZE - hp_text_h) // 2 - by0

        draw_text_with_shadow(draw, (hp_start_x, text_y), hp_str, font_stat, "#FFFFFF", stroke_w=2)
        card.alpha_composite(hp_icon, (hp_start_x + hp_text_w + hp_gap, STAT_Y))

    # 2. Blok Ataku wyśrodkowany na osi RIGHT_HALF_CENTER
    if os.path.exists(atk_icon_path) and (atk_1 > 0 or atk_2 > 0):
        atk_icon = Image.open(atk_icon_path).convert("RGBA").resize(
            (ICON_SIZE, ICON_SIZE), Image.Resampling.LANCZOS
        )
        atk_block_w = 0
        if atk_1 > 0:
            atk_block_w += atk_1 * ICON_SIZE + max(0, atk_1 - 1) * ICON_GAP
        if atk_2 > 0:
            atk_block_w += 34 + atk_2 * ICON_SIZE + max(0, atk_2 - 1) * ICON_GAP

        curr_x = RIGHT_HALF_CENTER - atk_block_w // 2

        for i in range(atk_1):
            card.alpha_composite(atk_icon, (curr_x, STAT_Y))
            curr_x += ICON_SIZE + (ICON_GAP if i < atk_1 - 1 else 0)

        if atk_2 > 0:
            draw_text_with_shadow(draw, (curr_x + 5, STAT_Y + 2), "+", font_plus, "#FFFFFF", stroke_w=2)
            curr_x += 34
            for i in range(atk_2):
                card.alpha_composite(atk_icon, (curr_x, STAT_Y))
                curr_x += ICON_SIZE + (ICON_GAP if i < atk_2 - 1 else 0)

    # Warstwa 6: Dolna belka (Umiejętność - podniesiona o 8 px i wyśrodkowana w belce)
    draw_wrapped_text(
        draw, umiejetnosc, (92, 1177, 853, 1304), font_skill, "#FFD766", align="center", line_spacing=5, stroke_w=2
    )

    final_rgb = card.convert("RGB")
    out_filename = f"{row_num:02d}_{nazwa_pl}.png"
    out_path = os.path.join(OUTPUT_DIR, out_filename)
    final_rgb.save(out_path, dpi=(300, 300))
    rendered_cards.append(final_rgb)
    print(f"[OK] Wygenerowano: {out_filename}")

if rendered_cards:
    first_card, *other_cards = rendered_cards
    first_card.save(
        PDF_OUTPUT,
        save_all=True,
        append_images=other_cards,
        resolution=300.0,
    )
    print(f"\n[SUKCES] Zapisano komplet {len(rendered_cards)} kart oraz PDF: {PDF_OUTPUT}")

    # Arkusz testowy A4 (1:1 -> dokładnie 80 x 120 mm ze znacznikami cięcia)
    A4_W, A4_H = 2480, 3508
    a4_sheet = Image.new("RGB", (A4_W, A4_H), (255, 255, 255))
    draw_a4 = ImageDraw.Draw(a4_sheet)

    px = (A4_W - CARD_W) // 2
    py = (A4_H - CARD_H) // 2
    a4_sheet.paste(first_card, (px, py))

    m_len, m_off = 35, 10
    corners = [
        (px, py, -1, -1),
        (px + CARD_W, py, 1, -1),
        (px, py + CARD_H, -1, 1),
        (px + CARD_W, py + CARD_H, 1, 1),
    ]
    for cx, cy, dx, dy in corners:
        draw_a4.line([(cx + dx * m_off, cy), (cx + dx * (m_off + m_len), cy)], fill=(80, 80, 80), width=2)
        draw_a4.line([(cx, cy + dy * m_off), (cx, cy + dy * (m_off + m_len))], fill=(80, 80, 80), width=2)

    test_pdf_path = os.path.join(BASE_DIR, "test_karta_A4_8x12cm.pdf")
    a4_sheet.save(test_pdf_path, resolution=300.0)
    print(f"[TEST A4] Zapisano arkusz testowy 1:1 (80x120 mm): {test_pdf_path}")
