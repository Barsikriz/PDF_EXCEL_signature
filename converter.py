"""Windows Excel export + non-rasterizing PDF signature insertion."""
from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pymupdf

MM = 72 / 25.4
SIGNATURE_HEIGHT = 22 * MM
BOTTOM_OFFSET = 16 * MM
RESERVED_BOTTOM = SIGNATURE_HEIGHT + BOTTOM_OFFSET + 6 * MM


def export_excel(source: Path, target: Path, sheet_name: str = "",
                 landscape: bool = False, use_print_area: bool = False) -> str:
    """Open a separate Excel instance. Never save changes to the workbook."""
    if os.name != "nt":
        raise RuntimeError("Для экспорта нужны Windows и установленный Microsoft Excel.")
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = book = sheet = setup = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3  # Disable VBA macros before opening.
        book = excel.Workbooks.Open(
            Filename=str(source.resolve()), UpdateLinks=0, ReadOnly=True,
            Password="", WriteResPassword="", IgnoreReadOnlyRecommended=True,
            AddToMru=False, Notify=False,
        )
        if sheet_name.strip():
            try:
                sheet = book.Worksheets(sheet_name.strip())
            except Exception as exc:
                raise ValueError(f"Лист «{sheet_name}» не найден.") from exc
        else:
            for index in range(1, book.Worksheets.Count + 1):
                candidate = book.Worksheets(index)
                if candidate.Visible == -1:
                    sheet = candidate
                    break
        if sheet is None or sheet.Visible != -1:
            raise ValueError("Нужен видимый лист с таблицей.")
        # CountA includes formulas, unlike a check of displayed PDF text.
        if excel.WorksheetFunction.CountA(sheet.UsedRange) == 0:
            raise ValueError("Выбранный лист пуст.")
        setup = sheet.PageSetup
        setup.Zoom = False
        setup.FitToPagesWide = 1
        setup.FitToPagesTall = False  # Any number of pages vertically.
        setup.Orientation = 2 if landscape else 1
        setup.PaperSize = 9  # A4
        setup.CenterVertically = False
        setup.BottomMargin = max(float(setup.BottomMargin), RESERVED_BOTTOM)
        # Footer is retained; collision detection below catches unusual footers.
        name = str(sheet.Name)
        sheet.ExportAsFixedFormat(
            Type=0, Filename=str(target.resolve()), Quality=0,
            IncludeDocProperties=True, IgnorePrintAreas=not use_print_area,
            OpenAfterPublish=False,
        )
        if not target.is_file():
            raise RuntimeError("Excel не создал PDF.")
        return name
    finally:
        setup = sheet = None
        try:
            if book is not None:
                book.Close(SaveChanges=False)
        finally:
            book = None
            try:
                if excel is not None:
                    excel.Quit()
            finally:
                excel = None
                pythoncom.CoUninitialize()


def signature_box(page: pymupdf.Page, width_mm: float) -> pymupdf.Rect:
    if not 20 <= width_mm <= 90:
        raise ValueError("Ширина подписи должна быть от 20 до 90 мм.")
    right = page.rect.width - 15 * MM
    bottom = page.rect.height - BOTTOM_OFFSET
    return pymupdf.Rect(right - width_mm * MM, bottom - SIGNATURE_HEIGHT,
                        right, bottom)


def intersects_content(page: pymupdf.Page, rect: pymupdf.Rect) -> bool:
    """Conservative check: do not cover text, images or vector graphics."""
    for block in page.get_text("blocks"):
        if pymupdf.Rect(block[:4]).intersects(rect):
            return True
    for item in page.get_image_info():
        if pymupdf.Rect(item["bbox"]).intersects(rect):
            return True
    for item in page.get_drawings():
        bounds = item["rect"]
        # Expand zero-height/width lines too.
        bounds = pymupdf.Rect(bounds.x0 - 0.5, bounds.y0 - 0.5,
                              bounds.x1 + 0.5, bounds.y1 + 0.5)
        if bounds.intersects(rect):
            return True
    return False


def add_signature(source: Path, signature: Path, target: Path,
                  width_mm: float = 45) -> int:
    """Insert only an image. Keep every original page and its text objects."""
    with pymupdf.open(source) as doc:
        if not len(doc):
            raise ValueError("В PDF нет страниц.")
        original_text = [p.get_text() for p in doc]
        if not any(t.strip() for t in original_text):
            raise ValueError("Excel создал PDF без извлекаемого текста. Проверьте таблицу.")
        page = doc[-1]
        box = signature_box(page, width_mm)
        if intersects_content(page, box):
            raise ValueError(
                "Место для подписи занято (возможно, колонтитулом или рисунком). "
                "Уберите этот элемент из нижней части листа и повторите. "
                "PDF не сохранён, чтобы не перекрыть содержимое."
            )
        pix = pymupdf.Pixmap(str(signature))
        scale = min(box.width / pix.width, box.height / pix.height)
        # Right/bottom aligned, with original proportions.
        actual = pymupdf.Rect(box.x1 - pix.width * scale,
                              box.y1 - pix.height * scale, box.x1, box.y1)
        page.insert_image(actual, pixmap=pix)
        doc.save(target, garbage=3, deflate=True)
    with pymupdf.open(target) as check:
        if [p.get_text() for p in check] != original_text:
            target.unlink(missing_ok=True)
            raise RuntimeError("Проверка сохранности текста PDF не пройдена.")
        return len(check)


def convert(source: Path, signature: Path, target: Path, sheet_name: str = "",
            width_mm: float = 45, landscape: bool = False,
            use_print_area: bool = False) -> tuple[int, str]:
    source, signature, target = map(lambda p: Path(p).resolve(),
                                    (source, signature, target))
    if source.suffix.lower() != ".xlsx" or not source.is_file():
        raise ValueError("Выберите существующий Excel-файл .xlsx.")
    if signature.suffix.lower() not in {".png", ".jpg", ".jpeg"} or not signature.is_file():
        raise ValueError("Выберите изображение подписи PNG или JPEG.")
    if target.suffix.lower() != ".pdf":
        raise ValueError("Выходной файл должен иметь расширение .pdf.")
    if target in (source, signature):
        raise ValueError("Результат нельзя записывать поверх исходного файла.")
    if not 20 <= width_mm <= 90:
        raise ValueError("Ширина подписи должна быть от 20 до 90 мм.")
    # Validate image before starting Excel.
    pymupdf.Pixmap(str(signature))
    # Same filesystem: atomic replacement only after successful validation.
    with tempfile.TemporaryDirectory(prefix="excel-pdf-", dir=target.parent) as temp:
        temp = Path(temp)
        name = export_excel(source, temp / "table.pdf", sheet_name,
                            landscape, use_print_area)
        pages = add_signature(temp / "table.pdf", signature,
                              temp / "signed.pdf", width_mm)
        os.replace(temp / "signed.pdf", target)
    return pages, name
