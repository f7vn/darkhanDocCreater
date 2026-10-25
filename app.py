from __future__ import annotations

import json
import hmac
import os
import re
import secrets
import shutil
import subprocess
import tempfile
import threading
import uuid
from copy import deepcopy
from decimal import Decimal, InvalidOperation
from http import HTTPStatus
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

import pymupdf as fitz
from docx import Document
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
TEMPLATE_DIR = ROOT / "templates"
GENERATED_DIR = ROOT / "generated"
GENERATED_DIR.mkdir(exist_ok=True)

AUTH_USERNAME = os.environ.get("DOCGEN_USERNAME", "")
AUTH_PASSWORD = os.environ.get("DOCGEN_PASSWORD", "")
AUTH_COOKIE_NAME = "docgen_session"
AUTH_SESSION_TOKEN = secrets.token_urlsafe(32)
AUTH_SESSION_SECONDS = 8 * 60 * 60

RATES = {
    "abdykhalykov": Decimal("25000"),
    "abdicul": Decimal("27000"),
    "archi": Decimal("29000"),
}

COMBINED_AREA_RATES = {
    "best": Decimal("25000"),
    "nalibai": Decimal("27000"),
    "archi": Decimal("29000"),
}

KTRU = {
    "aesthetic": {
        "label": "Эстетическое оформление",
        "short": "Эстетическое оформление",
        "ru_subject": "услуги по эстетическому оформлению",
        "ru_genitive": "эстетическому оформлению",
        "kz_subject": "эстетикалық безендіру қызметтері",
        "kz_genitive": "эстетикалық безендіру",
        "abdicul": "Настенная печать с помощью УФ-принтера",
        "abdykhalykov": "Эстетическое оформление строго в рамках Брендбука Адал Азамат",
        "archi": "Эстетическое оформление",
    },
    "finishing": {
        "label": "Декоративно-отделочные работы",
        "short": "Декоративно-отделочные работы",
        "ru_subject": "декоративно-отделочные работы",
        "ru_genitive": "декоративно-отделочным работам",
        "kz_subject": "сәндік-әрлеу жұмыстары",
        "kz_genitive": "сәндік-әрлеу жұмыстары",
        "abdicul": "Декоративно-отделочные работы",
        "abdykhalykov": "Декоративно-отделочные работы",
        "archi": "Декоративно-отделочные работы",
    },
    "graphics": {
        "label": "Нанесение надписей и изображений",
        "short": "Нанесение надписей и изображений на объекты",
        "ru_subject": "услуги по нанесению надписей и изображений на объекты",
        "ru_genitive": "нанесению надписей и изображений на объекты",
        "kz_subject": "объектілерге жазулар мен бейнелерді салу қызметтері",
        "kz_genitive": "объектілерге жазулар мен бейнелерді салу",
        "abdicul": "Нанесение надписей и изображений на объекты",
        "abdykhalykov": "Нанесение надписей и изображений на объекты",
        "archi": "Нанесение надписей и изображений",
    },
    "combined": {
        "label": "Декоративно-отделочные работы + стенды",
        "short": "Декор + стенды",
    },
}

CAP_TOTAL = Decimal("2162500")


def decimal_label(value: Decimal) -> str:
    if value == value.to_integral():
        return f"{int(value):,}".replace(",", " ")
    normalized = format(value.normalize(), "f")
    integer, fraction = normalized.split(".")
    return f"{int(integer):,}".replace(",", " ") + "," + fraction


def slug_area(value: Decimal) -> str:
    return decimal_label(value).replace(" ", "").replace(",", "-")


def parse_area(raw: object) -> Decimal:
    text = str(raw).strip().replace(" ", "").replace(",", ".")
    try:
        value = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError("Введите площадь числом") from exc
    if value <= 0:
        raise ValueError("Площадь должна быть больше нуля")
    if value > Decimal("100000"):
        raise ValueError("Площадь не может превышать 100 000 м²")
    if value.as_tuple().exponent < -2:
        raise ValueError("Используйте не более двух знаков после запятой")
    return value


def parse_count(raw: object) -> int:
    text = str(raw).strip().replace(" ", "")
    if not re.fullmatch(r"\d+", text):
        raise ValueError("Введите количество стендов целым числом")
    value = int(text)
    if value <= 0:
        raise ValueError("Количество стендов должно быть больше нуля")
    if value > 10000:
        raise ValueError("Количество стендов не может превышать 10 000")
    return value


def parse_money(raw: object, provider: str) -> Decimal:
    text = str(raw).strip().replace(" ", "").replace(",", ".")
    try:
        value = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"Укажите цену стенда для {provider}") from exc
    if value <= 0:
        raise ValueError(f"Цена стенда для {provider} должна быть больше нуля")
    if value > Decimal("100000000"):
        raise ValueError(f"Цена стенда для {provider} слишком велика")
    if value.as_tuple().exponent < -2:
        raise ValueError("В цене можно использовать не более двух знаков после запятой")
    return value


def iter_all_paragraphs(document: Document):
    yield from document.paragraphs
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs


def set_paragraph_text(paragraph, text: str, bold_prefix: str | None = None) -> None:
    run_properties = None
    if paragraph.runs and paragraph.runs[0]._r.rPr is not None:
        run_properties = deepcopy(paragraph.runs[0]._r.rPr)
    paragraph.clear()
    run = paragraph.add_run(text)
    if run_properties is not None:
        run._r.insert(0, run_properties)
    if run.font.name is None:
        run.font.name = "Times New Roman"
    if run.font.size is None:
        run.font.size = Pt(11)
    if bold_prefix and text.startswith(bold_prefix):
        # Keep the source layout stable; the provider remains easy to scan through numbering.
        run.bold = False


def update_justification(source: Path, destination: Path, area: Decimal, kind: dict) -> None:
    doc = Document(source)
    area_text = decimal_label(area)
    totals = {key: rate * area for key, rate in RATES.items()}

    provider_lines_ru = {
        "Абдихалыков": f"ИП «Абдихалыков»- 25 000 тг за 1 м2 * {area_text} м2= {decimal_label(totals['abdykhalykov'])} тг",
        "Әбдіқұл": f"ИП «Әбдіқұл»- 27 000 тг за 1 м2 * {area_text} м2= {decimal_label(totals['abdicul'])} тг",
        "Архиспецснабмонтаж": f"ТОО «Архиспецснабмонтаж»- 29 000 тг за 1 м2 * {area_text} м2= {decimal_label(totals['archi'])} тг",
    }
    provider_lines_kz = {
        "Абдихалыков": f"ЖК «Абдихалыков»- 25 000 тг за 1 м2 * {area_text} м2= {decimal_label(totals['abdykhalykov'])} тг",
        "Әбдіқұл": f"ЖК «Әбдіқұл»- 27 000 тг за 1 м2 * {area_text} м2= {decimal_label(totals['abdicul'])} тг",
        "Архиспецснабмонтаж": f"ЖШС «Архиспецснабмонтаж»- 29 000 тг за 1 м2 * {area_text} м2= {decimal_label(totals['archi'])} тг",
    }
    minimum = totals["abdykhalykov"]
    ru_comparison = "не превышает" if minimum <= CAP_TOTAL else "превышает"
    kz_comparison = "аспайды" if minimum <= CAP_TOTAL else "асады"

    for paragraph in doc.paragraphs:
        text = paragraph.text
        if "По итогам проведённого анализа" in text:
            set_paragraph_text(
                paragraph,
                "По итогам проведённого анализа установлено, что стоимость услуг по "
                f"{kind['ru_genitive']} составляет {decimal_label(minimum)} тенге без учета НДС, "
                f"что {ru_comparison} предельный объем закупа, установленный подпунктом 28 пункта 1 "
                "статьи 14 Закона — 500 МРП (500 × 4 325 = 2 162 500 тенге).",
            )
            continue
        if "Жүргізілген талдау қорытындысы" in text:
            set_paragraph_text(
                paragraph,
                "Жүргізілген талдау қорытындысы бойынша "
                f"{kind['kz_genitive']} бойынша қызметтердің құны ҚҚС есептемегенде "
                f"{decimal_label(minimum)} теңгені құрайтыны анықталды, бұл Заңның 14-бабы "
                "1-тармағының 28-тармақшасында белгіленген сатып алудың шекті көлемінен "
                f"{kz_comparison} — 500 МРП (500 × 4 325 = 2 162 500 теңге).",
            )
            continue
        for provider, replacement in provider_lines_ru.items():
            if provider in text and (text.startswith("ИП") or text.startswith("ТОО")) and "м2" in text:
                set_paragraph_text(paragraph, replacement)
                break
        else:
            for provider, replacement in provider_lines_kz.items():
                if provider in text and (text.startswith("ЖК") or text.startswith("ЖШС")) and "м2" in text:
                    set_paragraph_text(paragraph, replacement)
                    break

    doc.save(destination)


def update_specification(source: Path, destination: Path, area: Decimal, kind: dict) -> None:
    doc = Document(source)
    for paragraph in doc.paragraphs:
        if paragraph.text.startswith("Предметом закупки является"):
            set_paragraph_text(paragraph, f"Предметом закупки являются {kind['ru_subject']}. ")
        elif paragraph.text.startswith("Сатып алу нысаны"):
            set_paragraph_text(paragraph, f"Сатып алу нысаны {kind['kz_subject']} болып табылады. ")

    area_text = f"{decimal_label(area)} м2"
    for table in doc.tables:
        if len(table.rows) > 1 and len(table.columns) > 2:
            set_paragraph_text(table.cell(1, 2).paragraphs[0], area_text)

    # The source leaves eight empty lines before the Russian appendix. In Word,
    # this can strand the header row at the foot of one page and the data row on
    # the next. Keep four lines of visual separation and make this one table a
    # little denser, without changing the main body or the Kazakh section.
    russian_appendix_match = next(
        (
            (index, paragraph)
            for index, paragraph in enumerate(doc.paragraphs)
            if paragraph.text.strip() == "Приложение 1"
        ),
        None,
    )
    if russian_appendix_match is not None:
        appendix_index, russian_appendix = russian_appendix_match
        paragraphs = doc.paragraphs
        empty_before = []
        for paragraph in reversed(paragraphs[:appendix_index]):
            if paragraph.text.strip():
                break
            empty_before.append(paragraph)
        for paragraph in empty_before[:4]:
            paragraph._element.getparent().remove(paragraph._element)

        russian_appendix.paragraph_format.space_before = Pt(0)
        russian_appendix.paragraph_format.space_after = Pt(3)
        russian_appendix.paragraph_format.keep_with_next = True

    # Replace the long run of spacer paragraphs between the two language
    # sections with a deterministic page break. This keeps the Kazakh heading
    # at the top of a fresh page in both Word and LibreOffice.
    kazakh_heading_match = next(
        (
            (index, paragraph)
            for index, paragraph in enumerate(doc.paragraphs)
            if paragraph.text.strip() == "ТЕХНИКАЛЫҚ СИПАТТАМА"
        ),
        None,
    )
    if kazakh_heading_match is not None:
        heading_index, kazakh_heading = kazakh_heading_match
        paragraphs = doc.paragraphs
        empty_before = []
        for paragraph in reversed(paragraphs[:heading_index]):
            if paragraph.text.strip():
                break
            empty_before.append(paragraph)
        for paragraph in empty_before:
            paragraph._element.getparent().remove(paragraph._element)
        kazakh_heading.paragraph_format.page_break_before = True
        kazakh_heading.paragraph_format.space_before = Pt(0)

    kazakh_appendix_match = next(
        (
            (index, paragraph)
            for index, paragraph in enumerate(doc.paragraphs)
            if paragraph.text.strip() == "Қосымша 1"
        ),
        None,
    )
    if kazakh_appendix_match is not None:
        appendix_index, kazakh_appendix = kazakh_appendix_match
        paragraphs = doc.paragraphs
        empty_before = []
        for paragraph in reversed(paragraphs[:appendix_index]):
            if paragraph.text.strip():
                break
            empty_before.append(paragraph)
        for paragraph in empty_before:
            paragraph._element.getparent().remove(paragraph._element)
        kazakh_appendix.paragraph_format.space_before = Pt(0)
        kazakh_appendix.paragraph_format.space_after = Pt(3)
        kazakh_appendix.paragraph_format.keep_with_next = True

    for specification_table in doc.tables:
        # The reference stores both appendix tables as floating objects. Word
        # may then place the caption and table on different pages. Inline
        # tables paginate predictably and preserve the same visible geometry.
        floating_position = specification_table._tbl.tblPr.find(qn("w:tblpPr"))
        if floating_position is not None:
            specification_table._tbl.tblPr.remove(floating_position)

        for row_index, row in enumerate(specification_table.rows):
            row.height = Pt(28)
            row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

            row_properties = row._tr.get_or_add_trPr()
            if row_properties.find(qn("w:cantSplit")) is None:
                row_properties.append(OxmlElement("w:cantSplit"))

            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)
                    paragraph.paragraph_format.line_spacing = 1.0
                    paragraph.paragraph_format.keep_together = True
                    if row_index == 0:
                        paragraph.paragraph_format.keep_with_next = True
                    for run in paragraph.runs:
                        run.font.size = Pt(12)
    doc.save(destination)


def arial_font(bold: bool = False) -> Path:
    candidates = [
        Path("/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf"),
        Path("/Library/Fonts/Arial Bold.ttf" if bold else "/Library/Fonts/Arial.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise RuntimeError("Не найден шрифт с поддержкой кириллицы (Arial или DejaVu Sans)")


def erase(page, *rects: fitz.Rect) -> None:
    for rect in rects:
        page.add_redact_annot(rect, fill=(1, 1, 1), cross_out=False)
    page.apply_redactions(
        images=fitz.PDF_REDACT_IMAGE_NONE,
        graphics=fitz.PDF_REDACT_LINE_ART_NONE,
        text=fitz.PDF_REDACT_TEXT_REMOVE,
    )


def textbox(page, rect: fitz.Rect, text: str, max_size: float, min_size: float = 5.0,
            bold: bool = False, align: int = fitz.TEXT_ALIGN_LEFT) -> None:
    font_path = arial_font(bold)
    font_name = "DocArialBold" if bold else "DocArial"
    page.insert_font(fontname=font_name, fontfile=str(font_path))
    size = max_size
    while size >= min_size:
        shape = page.new_shape()
        result = shape.insert_textbox(
            rect,
            text,
            fontname=font_name,
            fontfile=str(font_path),
            fontsize=size,
            lineheight=1.05,
            color=(0, 0, 0),
            align=align,
        )
        if result >= 0:
            shape.commit(overlay=True)
            return
        size -= 0.4
    raise RuntimeError(f"Текст не помещается в шаблон: {text}")


def save_flattened_a4(source: fitz.Document, destination: Path, dpi: int = 144) -> None:
    """Save one edited supplier page as a simple, fast A4 image page."""
    page = source[0]
    a4 = fitz.paper_rect("a4")
    page_ratio = page.rect.width / page.rect.height
    a4_ratio = a4.width / a4.height
    if page_ratio >= a4_ratio:
        content_width = a4.width
        content_height = content_width / page_ratio
        left = 0
        top = (a4.height - content_height) / 2
    else:
        content_height = a4.height
        content_width = content_height * page_ratio
        left = (a4.width - content_width) / 2
        top = 0

    target_width = round(content_width * dpi / 72)
    render_scale = target_width / page.rect.width
    pixmap = page.get_pixmap(matrix=fitz.Matrix(render_scale, render_scale), alpha=False)
    flattened_page = pixmap.tobytes("jpeg", jpg_quality=90)

    normalized = fitz.open()
    out_page = normalized.new_page(width=a4.width, height=a4.height)
    out_page.insert_image(
        fitz.Rect(left, top, left + content_width, top + content_height),
        stream=flattened_page,
        keep_proportion=True,
    )
    normalized.save(
        destination,
        garbage=4,
        deflate=True,
        deflate_images=True,
        use_objstms=1,
    )
    normalized.close()


def create_offer_abdicul(destination: Path, area: Decimal, total: Decimal, kind: dict) -> None:
    doc = fitz.open(TEMPLATE_DIR / "offer-abdicul.pdf")
    page = doc[0]

    erase(
        page,
        fitz.Rect(86, 184, 542, 257),
        fitz.Rect(117.5, 457.75, 369.5, 472.7),
        fitz.Rect(410.5, 457.75, 518.5, 472.7),
    )
    title = f"Коммерческое предложение\nПо {kind['ru_genitive']}"
    textbox(page, fitz.Rect(94, 192, 535, 252), title, 11.0, 8.0, bold=True, align=fitz.TEXT_ALIGN_CENTER)

    textbox(page, fitz.Rect(118, 458, 369, 472), f"{kind['abdicul']} — {decimal_label(area)} м2", 7.2, 5.0)
    textbox(page, fitz.Rect(411, 458, 516, 472), f"{decimal_label(total)} тенге", 7.2, 5.5)
    page.draw_line(fitz.Point(85.58, 457.27), fitz.Point(552.45, 457.27), color=(0, 0, 0), width=0.48, overlay=True)
    save_flattened_a4(doc, destination)
    doc.close()


def create_offer_abdykhalykov(destination: Path, area: Decimal, total: Decimal, kind: dict) -> None:
    doc = fitz.open(TEMPLATE_DIR / "offer-abdykhalykov.pdf")
    page = doc[0]
    erase(
        page,
        fitz.Rect(127.0, 461.6, 327.2, 489.5),
        fitz.Rect(332.5, 461.6, 399.5, 476.5),
        fitz.Rect(432.5, 461.6, 521.5, 476.5),
    )
    textbox(page, fitz.Rect(131, 461, 327, 491), kind["abdykhalykov"], 8.0, 5.2)
    textbox(page, fitz.Rect(334, 462, 398, 476), f"{decimal_label(area)} м2", 8.0, 5.5, align=fitz.TEXT_ALIGN_CENTER)
    textbox(page, fitz.Rect(433, 462, 520, 476), decimal_label(total), 8.0, 5.5, align=fitz.TEXT_ALIGN_CENTER)
    page.draw_line(fitz.Point(99.74, 461.11), fitz.Point(527.25, 461.11), color=(0, 0, 0), width=0.48, overlay=True)
    erase(page, fitz.Rect(0, 0, 40, page.rect.height))
    save_flattened_a4(doc, destination)
    doc.close()


def create_offer_archi(destination: Path, area: Decimal, total: Decimal, kind: dict) -> None:
    source = fitz.open(TEMPLATE_DIR / "offer-archi.pdf")
    page = source[0]
    erase(
        page,
        fitz.Rect(410, 1748, 2250, 1828),
        fitz.Rect(650, 1828, 2300, 1907),
        fitz.Rect(620, 1943, 1900, 2022),
    )
    textbox(page, fitz.Rect(430, 1760, 2210, 1820), kind["archi"], 48.5, 30.0, bold=True)
    textbox(
        page,
        fitz.Rect(685, 1837, 2260, 1898),
        f"29 000 тенге без НДС за 1 м2. Общая площадь: {decimal_label(area)} м2",
        48.0,
        28.0,
    )
    textbox(page, fitz.Rect(631, 1953, 1870, 2013), f"{decimal_label(total)} тенге без НДС", 48.0, 28.0, bold=True)

    # The supplier scan is 2604×3368 points and contains a 33-megapixel
    # image. Flattening it keeps PDF viewers from decoding the original image
    # and nested form resources on every zoom or scroll.
    save_flattened_a4(source, destination)
    source.close()


def add_labeled_line(paragraph, label: str, value: str, break_before: bool = False) -> None:
    if break_before:
        paragraph.add_run("\n")
    label_run = paragraph.add_run(label)
    label_run.bold = True
    label_run.font.name = "Times New Roman"
    value_run = paragraph.add_run(value)
    value_run.font.name = "Times New Roman"


def update_combined_specification(source: Path, destination: Path, area: Decimal, stand_count: int) -> None:
    doc = Document(source)
    for table in doc.tables:
        if len(table.rows) < 3 or len(table.columns) < 4:
            continue
        set_paragraph_text(table.cell(1, 3).paragraphs[0], decimal_label(area))
        set_paragraph_text(table.cell(2, 3).paragraphs[0], str(stand_count))
    doc.save(destination)


def update_combined_justification(source: Path, destination: Path, minimum_total: Decimal) -> None:
    doc = Document(source)
    value = decimal_label(minimum_total)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    text = paragraph.text
                    if "________" not in text:
                        continue
                    updated = re.sub(r"_{4,}", value, text)
                    if minimum_total > CAP_TOTAL:
                        updated = updated.replace("не превышает", "превышает")
                        updated = updated.replace("аспайды", "асады")
                    set_paragraph_text(paragraph, updated)
    doc.save(destination)


def update_combined_offer_best(
    source: Path,
    destination: Path,
    area: Decimal,
    stand_count: int,
    stand_rate: Decimal,
    total: Decimal,
) -> None:
    doc = Document(source)
    for paragraph_index, paragraph in enumerate(doc.paragraphs):
        if paragraph.text.startswith("Общая площадь:"):
            paragraph.clear()
            add_labeled_line(paragraph, "Общая площадь: ", f"{decimal_label(area)} м²")
            add_labeled_line(
                paragraph,
                "Стоимость работ за 1 м²: ",
                f"{decimal_label(COMBINED_AREA_RATES['best'])} тенге",
                break_before=True,
            )
            for run in paragraph.runs:
                run.font.size = Pt(11)
        elif paragraph.text.startswith("Стоимость стенда"):
            paragraph.clear()
            add_labeled_line(
                paragraph,
                "Стенды: ",
                f"{stand_count} шт. × {decimal_label(stand_rate)} = {decimal_label(stand_rate * stand_count)} тенге",
            )
            add_labeled_line(
                paragraph,
                "Общая стоимость: ",
                f"{decimal_label(total)} тенге",
                break_before=True,
            )
            for run in paragraph.runs:
                run.font.size = Pt(11)
        elif paragraph.text.strip().startswith("BEST PRODUCTION") or 19 <= paragraph_index <= 24:
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = Pt(8.5)
            for run in paragraph.runs:
                run.font.size = Pt(7.5)
    for paragraph in reversed(doc.paragraphs):
        if paragraph.text.strip():
            break
        paragraph._element.getparent().remove(paragraph._element)
    doc.save(destination)


def update_combined_offer_nalibai(
    source: Path,
    destination: Path,
    area: Decimal,
    stand_count: int,
    stand_rate: Decimal,
    total: Decimal,
) -> None:
    doc = Document(source)
    table = doc.tables[0]
    values = {
        2: (
            f"{decimal_label(area)} м² × {decimal_label(COMBINED_AREA_RATES['nalibai'])} "
            f"тенге = {decimal_label(COMBINED_AREA_RATES['nalibai'] * area)} тенге"
        ),
        3: (
            f"{stand_count} шт. × {decimal_label(stand_rate)} тенге = "
            f"{decimal_label(stand_rate * stand_count)} тенге"
        ),
        4: f"{decimal_label(total)} тенге без НДС",
    }
    for row_index, value in values.items():
        paragraph = table.cell(row_index, 1).paragraphs[0]
        set_paragraph_text(paragraph, value)
        for run in paragraph.runs:
            run.font.size = Pt(10)
    for paragraph in reversed(doc.paragraphs):
        if paragraph.text.strip():
            break
        paragraph._element.getparent().remove(paragraph._element)
    doc.save(destination)


def create_combined_offer_archi(
    destination: Path,
    area: Decimal,
    stand_count: int,
    stand_rate: Decimal,
    total: Decimal,
) -> None:
    source = fitz.open(TEMPLATE_DIR / "combined-offer-archi.pdf")
    page = source[0]
    erase(
        page,
        fitz.Rect(400, 1892, 2180, 2034),
        fitz.Rect(400, 2060, 1950, 2150),
    )
    textbox(
        page,
        fitz.Rect(420, 1900, 2160, 1968),
        f"Объем отделочных работ: {decimal_label(area)} м²",
        48.0,
        32.0,
    )
    textbox(
        page,
        fitz.Rect(420, 1960, 2210, 2035),
        (
            f"Стенды: {stand_count} шт. × {decimal_label(stand_rate)} = "
            f"{decimal_label(stand_rate * stand_count)} тенге"
        ),
        45.0,
        29.0,
    )
    textbox(
        page,
        fitz.Rect(420, 2070, 1900, 2142),
        f"Общая стоимость работ: {decimal_label(total)} тенге",
        47.0,
        30.0,
        bold=True,
    )
    save_flattened_a4(source, destination)
    source.close()


def convert_docx_to_pdf(docx_path: Path, output_dir: Path) -> Path:
    soffice = shutil.which("soffice") or "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if not Path(soffice).exists():
        raise RuntimeError("LibreOffice не найден: без него нельзя собрать итоговый PDF")
    profile = output_dir / f".lo-profile-{uuid.uuid4().hex}"
    home = output_dir / f".lo-home-{uuid.uuid4().hex}"
    profile.mkdir()
    home.mkdir()
    env = os.environ.copy()
    env["HOME"] = str(home)
    command = [
        soffice,
        "--headless",
        f"-env:UserInstallation={profile.resolve().as_uri()}",
        "--convert-to",
        "pdf",
        "--outdir",
        str(output_dir),
        str(docx_path),
    ]
    completed = subprocess.run(command, env=env, capture_output=True, text=True, timeout=90)
    pdf_path = output_dir / f"{docx_path.stem}.pdf"
    if completed.returncode != 0 or not pdf_path.exists():
        detail = completed.stderr.strip() or completed.stdout.strip() or "неизвестная ошибка"
        raise RuntimeError(f"Не удалось преобразовать обоснование в PDF: {detail}")
    shutil.rmtree(profile, ignore_errors=True)
    shutil.rmtree(home, ignore_errors=True)
    return pdf_path


def merge_pdfs(destination: Path, sources: list[Path]) -> None:
    result = fitz.open()
    try:
        for source_path in sources:
            source = fitz.open(source_path)
            result.insert_pdf(source)
            source.close()
        result.set_metadata({"title": "Обоснование и коммерческие предложения", "author": "Генератор документов"})
        result.save(
            destination,
            garbage=4,
            deflate=True,
            deflate_images=True,
            deflate_fonts=True,
            use_objstms=1,
            compression_effort=50,
        )
    finally:
        result.close()


def generate_combined_package(area: Decimal, stand_count: int, stand_rates: dict[str, Decimal]) -> dict:
    job_id = uuid.uuid4().hex[:12]
    job_dir = GENERATED_DIR / job_id
    job_dir.mkdir(parents=True)
    area_slug = slug_area(area)
    spec_name = f"Техническая_спецификация_{area_slug}м2_{stand_count}шт.docx"
    bundle_name = f"Обоснование_и_3_КП_{area_slug}м2_{stand_count}шт.pdf"
    spec_path = job_dir / spec_name
    bundle_path = job_dir / bundle_name

    totals = {
        provider: COMBINED_AREA_RATES[provider] * area + stand_rates[provider] * stand_count
        for provider in COMBINED_AREA_RATES
    }

    justification_docx = job_dir / "Обоснование.docx"
    update_combined_justification(
        TEMPLATE_DIR / "combined-justification.docx",
        justification_docx,
        totals["best"],
    )
    update_combined_specification(
        TEMPLATE_DIR / "combined-specification.docx",
        spec_path,
        area,
        stand_count,
    )

    best_docx = job_dir / "КП_Best_Production.docx"
    nalibai_docx = job_dir / "КП_Налибай.docx"
    update_combined_offer_best(
        TEMPLATE_DIR / "combined-offer-best.docx",
        best_docx,
        area,
        stand_count,
        stand_rates["best"],
        totals["best"],
    )
    update_combined_offer_nalibai(
        TEMPLATE_DIR / "combined-offer-nalibai.docx",
        nalibai_docx,
        area,
        stand_count,
        stand_rates["nalibai"],
        totals["nalibai"],
    )

    justification_pdf = convert_docx_to_pdf(justification_docx, job_dir)
    best_pdf = convert_docx_to_pdf(best_docx, job_dir)
    nalibai_pdf = convert_docx_to_pdf(nalibai_docx, job_dir)
    archi_pdf = job_dir / "КП_Архи.pdf"
    create_combined_offer_archi(
        archi_pdf,
        area,
        stand_count,
        stand_rates["archi"],
        totals["archi"],
    )
    merge_pdfs(bundle_path, [justification_pdf, best_pdf, nalibai_pdf, archi_pdf])

    metadata = {
        "jobId": job_id,
        "ktru": KTRU["combined"]["label"],
        "area": decimal_label(area),
        "standCount": stand_count,
        "files": [
            {
                "name": spec_name,
                "kind": "DOCX",
                "description": "Техспецификация с площадью и стендами",
                "url": f"/download/{job_id}/{quote(spec_name)}",
            },
            {
                "name": bundle_name,
                "kind": "PDF",
                "description": "Обоснование и 3 коммерческих предложения",
                "url": f"/download/{job_id}/{quote(bundle_name)}",
            },
        ],
    }
    (job_dir / "metadata.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    return metadata


def generate_package(
    ktru_key: str,
    area: Decimal,
    stand_count: int | None = None,
    stand_rates: dict[str, Decimal] | None = None,
) -> dict:
    if ktru_key not in KTRU:
        raise ValueError("Выберите вид закупки из списка")
    if ktru_key == "combined":
        if stand_count is None or stand_rates is None:
            raise ValueError("Укажите количество и цены стендов")
        return generate_combined_package(area, stand_count, stand_rates)
    kind = KTRU[ktru_key]
    job_id = uuid.uuid4().hex[:12]
    job_dir = GENERATED_DIR / job_id
    job_dir.mkdir(parents=True)
    area_slug = slug_area(area)
    spec_name = f"Техническая_спецификация_{area_slug}м2.docx"
    bundle_name = f"Обоснование_и_3_КП_{area_slug}м2.pdf"
    spec_path = job_dir / spec_name
    bundle_path = job_dir / bundle_name

    totals = {key: rate * area for key, rate in RATES.items()}
    justification_docx = job_dir / "Обоснование.docx"
    update_justification(TEMPLATE_DIR / "justification.docx", justification_docx, area, kind)
    update_specification(TEMPLATE_DIR / "specification.docx", spec_path, area, kind)
    justification_pdf = convert_docx_to_pdf(justification_docx, job_dir)

    abdicul_pdf = job_dir / "КП_Абдикул.pdf"
    abdy_pdf = job_dir / "КП_Абдыхалыков.pdf"
    archi_pdf = job_dir / "КП_Архиспецснабмонтаж.pdf"
    create_offer_abdicul(abdicul_pdf, area, totals["abdicul"], kind)
    create_offer_abdykhalykov(abdy_pdf, area, totals["abdykhalykov"], kind)
    create_offer_archi(archi_pdf, area, totals["archi"], kind)
    merge_pdfs(bundle_path, [justification_pdf, abdy_pdf, abdicul_pdf, archi_pdf])

    metadata = {
        "jobId": job_id,
        "ktru": kind["label"],
        "area": decimal_label(area),
        "files": [
            {
                "name": spec_name,
                "kind": "DOCX",
                "description": "Редактируемая техническая спецификация",
                "url": f"/download/{job_id}/{quote(spec_name)}",
            },
            {
                "name": bundle_name,
                "kind": "PDF",
                "description": "Обоснование и 3 коммерческих предложения",
                "url": f"/download/{job_id}/{quote(bundle_name)}",
            },
        ],
    }
    (job_dir / "metadata.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    return metadata


class AppHandler(BaseHTTPRequestHandler):
    server_version = "DocumentGenerator/1.0"

    def log_message(self, format: str, *args) -> None:
        print(f"[{self.log_date_time_string()}] {format % args}")

    def send_bytes(self, body: bytes, content_type: str, status: int = 200, headers: dict | None = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        if headers:
            for key, value in headers.items():
                self.send_header(key, value)
        self.end_headers()
        self.wfile.write(body)

    def send_json(self, payload: dict, status: int = 200) -> None:
        self.send_bytes(json.dumps(payload, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8", status)

    def is_authenticated(self) -> bool:
        cookie = SimpleCookie()
        try:
            cookie.load(self.headers.get("Cookie", ""))
        except Exception:
            return False
        value = cookie.get(AUTH_COOKIE_NAME)
        return value is not None and hmac.compare_digest(value.value, AUTH_SESSION_TOKEN)

    def redirect(self, location: str, cookie: str | None = None) -> None:
        headers = {"Location": location, "Cache-Control": "no-store"}
        if cookie:
            headers["Set-Cookie"] = cookie
        self.send_bytes(b"", "text/plain; charset=utf-8", HTTPStatus.SEE_OTHER, headers)

    def serve_login(self, show_error: bool = False) -> None:
        path = STATIC_DIR / "login.html"
        if not path.exists():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        body = path.read_text(encoding="utf-8")
        if show_error:
            body = body.replace('id="login-error" class="login-error" hidden', 'id="login-error" class="login-error"')
        self.send_bytes(
            body.encode("utf-8"),
            "text/html; charset=utf-8",
            headers={"Cache-Control": "no-store"},
        )

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/health":
            self.send_json({"status": "ok"})
            return
        if parsed.path == "/login":
            if self.is_authenticated():
                self.redirect("/")
                return
            show_error = parse_qs(parsed.query).get("error") == ["1"]
            self.serve_login(show_error)
            return
        if parsed.path == "/":
            if not self.is_authenticated():
                self.redirect("/login")
                return
            self.serve_file(
                STATIC_DIR / "index.html",
                "text/html; charset=utf-8",
                {"Cache-Control": "no-store"},
            )
            return
        if parsed.path.startswith("/static/"):
            filename = parsed.path.removeprefix("/static/")
            if filename not in {"styles.css", "app.js"}:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            content_types = {".css": "text/css; charset=utf-8", ".js": "application/javascript; charset=utf-8"}
            self.serve_file(STATIC_DIR / filename, content_types.get(Path(filename).suffix, "application/octet-stream"))
            return
        if parsed.path.startswith("/download/"):
            if not self.is_authenticated():
                self.redirect("/login")
                return
            parts = parsed.path.split("/", 3)
            if len(parts) != 4 or not re.fullmatch(r"[a-f0-9]{12}", parts[2]):
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            job_dir = GENERATED_DIR / parts[2]
            metadata_path = job_dir / "metadata.json"
            if not metadata_path.exists():
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            filename = unquote(parts[3])
            metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
            allowed = {entry["name"] for entry in metadata["files"]}
            if filename not in allowed:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            path = job_dir / filename
            mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document" if path.suffix == ".docx" else "application/pdf"
            disposition = f"attachment; filename*=UTF-8''{quote(filename)}"
            self.serve_file(path, mime, {"Content-Disposition": disposition})
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def serve_file(self, path: Path, content_type: str, headers: dict | None = None) -> None:
        if not path.exists() or not path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        self.send_bytes(path.read_bytes(), content_type, headers=headers)

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        if path == "/login":
            self.handle_login()
            return
        if path == "/logout":
            self.redirect(
                "/login",
                f"{AUTH_COOKIE_NAME}=; Path=/; HttpOnly; SameSite=Strict; Max-Age=0",
            )
            return
        if path != "/api/generate":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        if not self.is_authenticated():
            self.send_json({"error": "Требуется авторизация"}, HTTPStatus.UNAUTHORIZED)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length <= 0 or length > 1_000_000:
                raise ValueError("Некорректный запрос")
            payload = json.loads(self.rfile.read(length))
            area = parse_area(payload.get("area", ""))
            ktru_key = str(payload.get("ktru", ""))
            stand_count = None
            stand_rates = None
            if ktru_key == "combined":
                stand_count = parse_count(payload.get("standCount", ""))
                stand_rates = {
                    "best": parse_money(payload.get("standPriceBest", ""), "Best Production"),
                    "nalibai": parse_money(payload.get("standPriceNalibai", ""), "ИП Нәлібай"),
                    "archi": parse_money(payload.get("standPriceArchi", ""), "ТОО Архи"),
                }
            result = generate_package(ktru_key, area, stand_count, stand_rates)
            self.send_json(result, HTTPStatus.CREATED)
        except (ValueError, json.JSONDecodeError) as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except Exception as exc:
            self.send_json({"error": f"Не удалось собрать документы: {exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def handle_login(self) -> None:
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length <= 0 or length > 10_000:
                raise ValueError
            fields = parse_qs(self.rfile.read(length).decode("utf-8"), keep_blank_values=True)
            username = fields.get("username", [""])[0]
            password = fields.get("password", [""])[0]
        except (ValueError, UnicodeDecodeError):
            self.redirect("/login?error=1")
            return

        valid_username = hmac.compare_digest(username.encode("utf-8"), AUTH_USERNAME.encode("utf-8"))
        valid_password = hmac.compare_digest(password.encode("utf-8"), AUTH_PASSWORD.encode("utf-8"))
        if not (valid_username and valid_password):
            self.redirect("/login?error=1")
            return

        cookie = (
            f"{AUTH_COOKIE_NAME}={AUTH_SESSION_TOKEN}; Path=/; HttpOnly; "
            f"SameSite=Strict; Max-Age={AUTH_SESSION_SECONDS}"
        )
        self.redirect("/", cookie)


def main() -> None:
    if not AUTH_USERNAME or not AUTH_PASSWORD:
        raise RuntimeError("Задайте переменные DOCGEN_USERNAME и DOCGEN_PASSWORD")
    railway_port = os.environ.get("PORT")
    host = os.environ.get("DOCGEN_HOST", "0.0.0.0" if railway_port else "127.0.0.1")
    port = int(railway_port or os.environ.get("DOCGEN_PORT", "8765"))
    server = ThreadingHTTPServer((host, port), AppHandler)
    print(f"Генератор документов запущен: http://{host}:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
