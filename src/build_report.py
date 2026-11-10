"""Шаг 0. Сборка отчёта `reports/report.docx` из `reports/sections/*.md`.

Почему так, а не правка Word напрямую: DOCX — бинарный формат, в git он
не даёт нормального diff. Поэтому исходники отчёта лежат текстом (`.md`),
а готовый документ собирается скриптом. Аналогия для Java-разработчика:
`.md` — это исходники `.java`, а `report.docx` — собранный `.jar`.

Поддерживаемый мини-синтаксис Markdown:
- заголовки `#`, `##`, `###` (заголовок 1-го уровня начинает новую страницу);
- абзацы, **жирный**, *курсив*, `моноширинный`;
- маркированные (`- `) и нумерованные (`1. `) списки;
- цитаты (`> `);
- таблицы в стиле GFM (строки с `|`, вторая строка-разделитель);
- картинки `![подпись](путь/к/картинке.png)`.

Файл `00_title.md` особый: это титульный лист в формате `ключ: значение`.

Запуск:  python src/build_report.py
"""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent.parent
SECTIONS_DIR = ROOT / "reports" / "sections"
REPORT_PATH = ROOT / "reports" / "report.docx"
TITLE_FILE = "00_title.md"

FONT = "Times New Roman"
MONO_FONT = "Consolas"
PAGE_WIDTH_CM = 16.0  # A4 (21 см) минус поля 2,5 + 2,5 см

# Инлайн-разметка: **жирный**, *курсив*, `код`.
INLINE_RE = re.compile(r"(\*\*.+?\*\*|\*[^*]+?\*|`[^`]+?`)")


# ---------------------------------------------------------------------------
# Стили и оформление
# ---------------------------------------------------------------------------

def _set_run_font(run, size: int | None = None, name: str = FONT) -> None:
    """Явно задаёт шрифт запуска, включая кириллицу (eastAsia)."""
    run.font.name = name
    if size is not None:
        run.font.size = Pt(size)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), name)


def setup_document(doc: Document) -> None:
    """Поля страницы, базовый шрифт и стили заголовков."""
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(12)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.first_line_indent = Cm(1.25)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # У стиля тоже нужно прописать eastAsia, иначе Word подменит шрифт.
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), FONT)

    headings = {"Heading 1": (16, True), "Heading 2": (14, True), "Heading 3": (12, True)}
    for style_name, (size, bold) in headings.items():
        style = doc.styles[style_name]
        style.font.name = FONT
        style.font.size = Pt(size)
        style.font.bold = bold
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.first_line_indent = Cm(0)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        style.paragraph_format.space_before = Pt(18 if size >= 16 else 12)
        style.paragraph_format.space_after = Pt(8)
        style.paragraph_format.keep_with_next = True


def add_page_number_footer(doc: Document) -> None:
    """Номер страницы по центру нижнего колонтитула."""
    footer = doc.sections[0].footer
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.first_line_indent = Cm(0)

    run = paragraph.add_run()
    _set_run_font(run, size=10)

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")

    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)


# ---------------------------------------------------------------------------
# Титульный лист
# ---------------------------------------------------------------------------

def read_title_meta(path: Path) -> dict[str, str]:
    """Читает `00_title.md` как набор строк `ключ: значение`."""
    meta: dict[str, str] = {}
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or ":" not in line:
            continue
        key, value = line.split(":", 1)
        meta[key.strip().lower()] = value.strip()
    return meta


def add_title_page(doc: Document, meta: dict[str, str]) -> None:
    """Титульный лист: тема, подзаголовок, автор и дата."""
    for _ in range(7):
        spacer = doc.add_paragraph()
        spacer.paragraph_format.first_line_indent = Cm(0)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.first_line_indent = Cm(0)
    title.paragraph_format.space_after = Pt(18)
    run = title.add_run(meta.get("title", "Без названия"))
    _set_run_font(run, size=20)
    run.bold = True

    subtitle = meta.get("subtitle")
    if subtitle:
        paragraph = doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.first_line_indent = Cm(0)
        run = paragraph.add_run(subtitle)
        _set_run_font(run, size=14)

    for _ in range(6):
        doc.add_paragraph()

    for key in ("author", "group", "date"):
        value = meta.get(key)
        if not value:
            continue
        paragraph = doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.first_line_indent = Cm(0)
        paragraph.paragraph_format.space_after = Pt(2)
        run = paragraph.add_run(value)
        _set_run_font(run, size=13)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


# ---------------------------------------------------------------------------
# Markdown -> DOCX
# ---------------------------------------------------------------------------

def add_inline(paragraph, text: str, size: int | None = None) -> None:
    """Добавляет в абзац текст с поддержкой **жирного**, *курсива* и `кода`."""
    for token in INLINE_RE.split(text):
        if not token:
            continue
        if token.startswith("**") and token.endswith("**") and len(token) > 4:
            run = paragraph.add_run(token[2:-2])
            run.bold = True
            _set_run_font(run, size=size)
        elif token.startswith("*") and token.endswith("*") and len(token) > 2:
            run = paragraph.add_run(token[1:-1])
            run.italic = True
            _set_run_font(run, size=size)
        elif token.startswith("`") and token.endswith("`") and len(token) > 2:
            run = paragraph.add_run(token[1:-1])
            _set_run_font(run, size=(size or 12) - 1, name=MONO_FONT)
        else:
            run = paragraph.add_run(token)
            _set_run_font(run, size=size)


def add_image(doc: Document, alt: str, target: str) -> None:
    """Вставляет картинку и подпись под ней."""
    image_path = Path(target)
    if not image_path.is_absolute():
        candidates = [ROOT / target, SECTIONS_DIR / target]
        image_path = next((p for p in candidates if p.exists()), candidates[0])

    if not image_path.exists():
        # Не роняем сборку из-за отсутствующей картинки — пишем заметное место.
        paragraph = doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.first_line_indent = Cm(0)
        run = paragraph.add_run(f"[нет файла изображения: {target}]")
        run.italic = True
        return

    doc.add_picture(str(image_path), width=Cm(PAGE_WIDTH_CM))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.paragraphs[-1].paragraph_format.first_line_indent = Cm(0)

    if alt:
        caption = doc.add_paragraph()
        caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption.paragraph_format.first_line_indent = Cm(0)
        caption.paragraph_format.space_after = Pt(10)
        run = caption.add_run(alt)
        _set_run_font(run, size=10)
        run.italic = True


def add_table(doc: Document, rows: list[str]) -> None:
    """Собирает таблицу GFM из строк вида `| a | b |`."""
    parsed: list[list[str]] = []
    for line in rows:
        cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
        parsed.append(cells)

    # Вторая строка GFM-таблицы — разделитель `|---|---|`, её пропускаем.
    if len(parsed) >= 2 and all(re.fullmatch(r":?-{2,}:?", c) for c in parsed[1]):
        header, body = parsed[0], parsed[2:]
    else:
        header, body = parsed[0], parsed[1:]

    columns = max(len(header), *(len(row) for row in body)) if body else len(header)
    table = doc.add_table(rows=1, cols=columns)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    for index in range(columns):
        cell = table.rows[0].cells[index]
        cell.text = ""
        paragraph = cell.paragraphs[0]
        paragraph.paragraph_format.first_line_indent = Cm(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.0
        text = header[index] if index < len(header) else ""
        add_inline(paragraph, f"**{text}**", size=10)

    for row in body:
        cells = table.add_row().cells
        for index in range(columns):
            cell = cells[index]
            cell.text = ""
            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0
            add_inline(paragraph, row[index] if index < len(row) else "", size=10)

    # Небольшой отступ после таблицы.
    spacer = doc.add_paragraph()
    spacer.paragraph_format.first_line_indent = Cm(0)
    spacer.paragraph_format.space_after = Pt(0)


def add_markdown(doc: Document, text: str, first_section: bool) -> None:
    """Переносит один `.md`-файл раздела в документ."""
    lines = text.splitlines()
    buffer: list[str] = []
    index = 0
    seen_h1 = False

    def flush_paragraph() -> None:
        if not buffer:
            return
        paragraph = doc.add_paragraph()
        add_inline(paragraph, " ".join(buffer).strip())
        buffer.clear()

    while index < len(lines):
        line = lines[index]
        stripped = line.strip()

        if not stripped:
            flush_paragraph()
            index += 1
            continue

        if stripped.startswith("|"):
            flush_paragraph()
            table_lines: list[str] = []
            while index < len(lines) and lines[index].strip().startswith("|"):
                table_lines.append(lines[index])
                index += 1
            add_table(doc, table_lines)
            continue

        image_match = re.fullmatch(r"!\[(.*?)]\((.+?)\)", stripped)
        if image_match:
            flush_paragraph()
            add_image(doc, image_match.group(1).strip(), image_match.group(2).strip())
            index += 1
            continue

        heading_match = re.match(r"^(#{1,3})\s+(.*)$", stripped)
        if heading_match:
            flush_paragraph()
            level = len(heading_match.group(1))
            if level == 1:
                if seen_h1 and not first_section:
                    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
                seen_h1 = True
            heading = doc.add_heading(level=level)
            add_inline(heading, heading_match.group(2))
            index += 1
            continue

        if stripped.startswith(">"):
            flush_paragraph()
            paragraph = doc.add_paragraph()
            paragraph.paragraph_format.left_indent = Cm(1)
            paragraph.paragraph_format.first_line_indent = Cm(0)
            run = paragraph.add_run(stripped.lstrip("> ").strip())
            run.italic = True
            index += 1
            continue

        bullet_match = re.match(r"^[-*]\s+(.*)$", stripped)
        if bullet_match:
            flush_paragraph()
            paragraph = doc.add_paragraph(style="List Bullet")
            paragraph.paragraph_format.first_line_indent = Cm(0)
            add_inline(paragraph, bullet_match.group(1))
            index += 1
            continue

        number_match = re.match(r"^(\d+)\.\s+(.*)$", stripped)
        if number_match:
            flush_paragraph()
            paragraph = doc.add_paragraph(style="List Number")
            paragraph.paragraph_format.first_line_indent = Cm(0)
            add_inline(paragraph, number_match.group(2))
            index += 1
            continue

        if stripped in {"---", "***"}:
            flush_paragraph()
            index += 1
            continue

        buffer.append(stripped)
        index += 1

    flush_paragraph()


# ---------------------------------------------------------------------------
# Точка входа
# ---------------------------------------------------------------------------

def build() -> Path:
    doc = Document()
    setup_document(doc)
    add_page_number_footer(doc)

    title_path = SECTIONS_DIR / TITLE_FILE
    if not title_path.exists():
        raise SystemExit(f"Не найден титульный лист: {title_path}")
    add_title_page(doc, read_title_meta(title_path))

    section_files = sorted(
        path for path in SECTIONS_DIR.glob("*.md") if path.name != TITLE_FILE
    )
    if not section_files:
        doc.add_paragraph("_(Разделы отчёта ещё не написаны.)_")

    for number, path in enumerate(section_files):
        if number > 0:
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        add_markdown(doc, path.read_text(encoding="utf-8"), first_section=(number == 0))

    doc.save(REPORT_PATH)
    return REPORT_PATH, [path.name for path in section_files]


def main() -> None:
    path, sections = build()
    print(f"Собрано: {path.relative_to(ROOT)}")
    print(f"Разделов: {len(sections)}" + (f" -> {', '.join(sections)}" if sections else ""))


if __name__ == "__main__":
    main()
