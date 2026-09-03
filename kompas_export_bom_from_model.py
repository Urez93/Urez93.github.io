# -*- coding: utf-8 -*-
"""
Макрос для Компас-3D: выгружает состав изделия (спецификацию) из
активно открытой модели и сохраняет уже СТРУКТУРИРОВАННУЮ таблицу
(с отступами по уровням вложенности и группировкой строк) рядом с
открытым файлом модели — в ту же папку.

Что делает:
  1. Подключается к запущенному Компас-3D и находит активный документ.
  2. По активному документу определяет папку, где лежит файл модели —
     туда же будет сохранён результат.
  3. Находит документ спецификации: либо сам активный документ уже
     является спецификацией, либо ищет открытую спецификацию среди
     остальных открытых в Компасе документов.
  4. Экспортирует спецификацию во временный плоский excel-файл штатным
     сохранением Компаса и сразу строит из него наглядную
     структурированную таблицу через kompas_bom_structure.py.
  5. Сохраняет итоговый файл "<имя_модели>_состав_изделия.xlsx" в папку
     с открытой моделью, временный плоский файл удаляет.

Требования:
  - pywin32 (win32com) — связь с Компасом. Больше НИЧЕГО стороннего не
    требуется: чтение и запись .xlsx сделаны на чистой стандартной
    библиотеке Python (zipfile + xml), потому что во встроенном
    интерпретаторе Компаса нет pip/openpyxl, и ставить их туда не нужно.

Файл полностью самостоятельный (логика построения структуры из
kompas_bom_structure.py включена сюда же) — это специально сделано,
потому что встроенный Python в Компасе не всегда видит соседние .py
файлы через обычный импорт (другая рабочая директория/__file__).
Для ручной пост-обработки уже сохранённых файлов вне Компаса можно
по-прежнему пользоваться отдельным kompas_bom_structure.py (он
использует openpyxl, если он есть в обычном Python на компьютере).

ВАЖНО про надёжность:
  Точные названия свойств/методов COM-объектов и числовые коды формата
  для экспорта спецификации в Excel могут отличаться между версиями
  Компас-3D. Поэтому для каждого шага сделано несколько попыток разными
  способами (как и в остальных макросах в этом репозитории) — в
  консоли будет видно, какой именно способ сработал, а какой нет. Если
  ни один из способов экспорта не сработает — пришлите вывод консоли,
  чтобы подобрать точный вызов для вашей версии Компаса.
"""

import os
import re
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape as _xml_escape


DOCUMENT_TYPE_SPECIFICATION = 5  # ksDocumentSpecification

# ---------------------------------------------------------------------------
# Построение структурированной таблицы (та же логика, что в
# kompas_bom_structure.py, включена сюда, чтобы макрос не зависел от
# импорта соседнего файла внутри встроенного Python Компаса).
#
# Чтение и запись .xlsx реализованы вручную поверх zipfile/xml — во
# встроенном Python Компаса нет openpyxl/xlrd и обычно нет pip, чтобы их
# поставить, а .xlsx — это просто zip-архив с XML-файлами внутри.
# ---------------------------------------------------------------------------

BOM_OUTPUT_COLUMNS = [
    ("Уровень", "level"),
    ("Позиция", "Позиция"),
    ("Обозначение", "Обозначение"),
    ("Наименование", "Наименование"),
    ("Количество", "Количество"),
    ("Масса", "Масса"),
    ("Материал", "Материал"),
    ("Раздел спецификации", "Раздел спецификации"),
    ("Форматы листов документа", "Форматы листов документа"),
    ("Вид изделия", "Вид изделия"),
    ("Обозначение стандарта", "Обозначение стандарта"),
    ("Типоразмер", "Типоразмер"),
    ("Примечание", "Примечание"),
]

_SS_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_COL_REF_RE = re.compile(r"([A-Z]+)(\d+)")


def _col_letters_to_index(letters):
    """'A' -> 0, 'B' -> 1, ... 'AA' -> 26 ..."""
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _col_index_to_letters(index):
    """0 -> 'A', 1 -> 'B', ... 26 -> 'AA' ..."""
    index += 1
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def _read_xlsx_rows(path):
    """Читает первый лист .xlsx-файла без сторонних библиотек."""
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()

        shared_strings = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall(_SS_NS + "si"):
                text = "".join(t.text or "" for t in si.iter(_SS_NS + "t"))
                shared_strings.append(text)

        sheet_candidates = sorted(
            n for n in names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")
        )
        if not sheet_candidates:
            raise RuntimeError("В файле {} не найден лист с данными.".format(path))
        sheet_name = "xl/worksheets/sheet1.xml" if "xl/worksheets/sheet1.xml" in sheet_candidates else sheet_candidates[0]

        root = ET.fromstring(zf.read(sheet_name))
        sheet_data = root.find(_SS_NS + "sheetData")

        rows = []
        if sheet_data is None:
            return rows

        for row_el in sheet_data.findall(_SS_NS + "row"):
            row_values = {}
            max_col = -1
            next_col = 0
            for c_el in row_el.findall(_SS_NS + "c"):
                ref = c_el.get("r", "")
                match = _COL_REF_RE.match(ref)
                col_idx = _col_letters_to_index(match.group(1)) if match else next_col
                next_col = col_idx + 1

                cell_type = c_el.get("t")
                value = None

                if cell_type == "s":
                    v_el = c_el.find(_SS_NS + "v")
                    if v_el is not None and v_el.text is not None:
                        value = shared_strings[int(v_el.text)]
                elif cell_type == "inlineStr":
                    is_el = c_el.find(_SS_NS + "is")
                    if is_el is not None:
                        value = "".join(t.text or "" for t in is_el.iter(_SS_NS + "t"))
                elif cell_type == "str":
                    v_el = c_el.find(_SS_NS + "v")
                    value = v_el.text if v_el is not None else None
                elif cell_type == "b":
                    v_el = c_el.find(_SS_NS + "v")
                    value = bool(int(v_el.text)) if v_el is not None else None
                else:
                    v_el = c_el.find(_SS_NS + "v")
                    if v_el is not None and v_el.text is not None:
                        text = v_el.text
                        try:
                            value = float(text)
                        except ValueError:
                            value = text

                row_values[col_idx] = value
                if col_idx > max_col:
                    max_col = col_idx

            rows.append([row_values.get(i) for i in range(max_col + 1)])
        return rows


def _read_bom_source_rows(path):
    """Читает исходную плоскую выгрузку Компаса (.xlsx) построчно, без сторонних библиотек."""
    ext = os.path.splitext(path)[1].lower()
    if ext != ".xlsx":
        raise RuntimeError(
            "Ожидается временный файл в формате .xlsx, получено: {}. "
            "Если Компас сохраняет спецификацию только в старом .xls, "
            "сообщите об этом — потребуется другой способ чтения.".format(ext)
        )

    all_rows = _read_xlsx_rows(path)
    if not all_rows:
        return []

    header = [str(v).strip() if v is not None else "" for v in all_rows[0]]
    rows = []
    for values in all_rows[1:]:
        if not any(v not in (None, "") for v in values):
            continue
        row = dict(zip(header, values))
        rows.append(row)
    return rows


def _bom_clean(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _bom_level_from_n(n_value):
    """Уровень вложенности по коду вида '1.2.3' -> 3. Пустой/битый код -> 1."""
    text = _bom_clean(n_value)
    if not text:
        return 1
    return text.count(".") + 1


def _build_bom_tree_rows(source_rows):
    """Дополняет строки уровнем вложенности и признаком 'это узел (есть дети)'."""
    codes = [_bom_clean(r.get("N")) for r in source_rows]

    for row, code in zip(source_rows, codes):
        row["level"] = _bom_level_from_n(code)
        row["_code"] = code
        row["_is_node"] = any(
            other and other != code and other.startswith(code + ".")
            for other in codes
        )
    return source_rows


class _StyleRegistry(object):
    """Собирает уникальные комбинации шрифт+заливка+отступ в компактный список cellXfs."""

    def __init__(self):
        self.fonts = [{"bold": False, "color": None}]
        self.fills = [{"color": None}, {"color": None}]  # 0, 1 зарезервированы конвенцией OOXML
        self.xfs = [{"font": 0, "fill": 0, "indent": 0, "align": None, "wrap": False}]  # 0 — стиль по умолчанию
        self._font_cache = {(False, None): 0}
        self._fill_cache = {None: 0}
        self._xf_cache = {}

    def _font_id(self, bold, color=None):
        key = (bold, color)
        if key not in self._font_cache:
            self.fonts.append({"bold": bold, "color": color})
            self._font_cache[key] = len(self.fonts) - 1
        return self._font_cache[key]

    def _fill_id(self, color):
        if color not in self._fill_cache:
            self.fills.append({"color": color})
            self._fill_cache[color] = len(self.fills) - 1
        return self._fill_cache[color]

    def xf_id(self, bold=False, color=None, fill=None, indent=0, align=None, wrap=False):
        font_id = self._font_id(bold, color)
        fill_id = self._fill_id(fill)
        key = (font_id, fill_id, indent, align, wrap)
        if key not in self._xf_cache:
            self.xfs.append({"font": font_id, "fill": fill_id, "indent": indent, "align": align, "wrap": wrap})
            self._xf_cache[key] = len(self.xfs) - 1
        return self._xf_cache[key]

    def to_xml(self):
        parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
        parts.append(
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        )

        parts.append('<fonts count="{}">'.format(len(self.fonts)))
        for f in self.fonts:
            bold_tag = "<b/>" if f["bold"] else ""
            color_tag = '<color rgb="{}"/>'.format(f["color"]) if f["color"] else ""
            parts.append('<font>{}<sz val="10"/><name val="Calibri"/>{}</font>'.format(bold_tag, color_tag))
        parts.append("</fonts>")

        parts.append('<fills count="{}">'.format(len(self.fills)))
        parts.append('<fill><patternFill patternType="none"/></fill>')
        parts.append('<fill><patternFill patternType="gray125"/></fill>')
        for f in self.fills[2:]:
            parts.append(
                '<fill><patternFill patternType="solid">'
                '<fgColor rgb="{0}"/><bgColor rgb="{0}"/></patternFill></fill>'.format(f["color"])
            )
        parts.append("</fills>")

        parts.append('<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>')
        parts.append('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')

        parts.append('<cellXfs count="{}">'.format(len(self.xfs)))
        for xf in self.xfs:
            align_bits = []
            if xf["align"]:
                align_bits.append('horizontal="{}"'.format(xf["align"]))
            align_bits.append('vertical="center"')
            if xf["indent"]:
                align_bits.append('indent="{}"'.format(xf["indent"]))
            if xf["wrap"]:
                align_bits.append('wrapText="1"')
            parts.append(
                '<xf numFmtId="0" fontId="{}" fillId="{}" borderId="0" xfId="0" applyFont="1" '
                'applyFill="1" applyAlignment="1"><alignment {}/></xf>'.format(
                    xf["font"], xf["fill"], " ".join(align_bits)
                )
            )
        parts.append("</cellXfs>")
        parts.append('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')

        parts.append("</styleSheet>")
        return "".join(parts)


def _xlsx_cell_xml(col_index, row_index, value, style_id):
    ref = "{}{}".format(_col_index_to_letters(col_index), row_index)
    if value is None or value == "":
        return '<c r="{}" s="{}"/>'.format(ref, style_id)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return '<c r="{}" s="{}"><v>{}</v></c>'.format(ref, style_id, value)
    return '<c r="{}" s="{}" t="inlineStr"><is><t xml:space="preserve">{}</t></is></c>'.format(
        ref, style_id, _xml_escape(str(value))
    )


def _write_xlsx(output_path, header, data_rows, name_col_index):
    """
    Пишет .xlsx напрямую (zip + XML), без openpyxl. data_rows — список
    словарей {"level": int, "is_node": bool, "cells": [значения по столбцам]}.
    """
    styles = _StyleRegistry()
    header_style = styles.xf_id(bold=True, color="FFFFFFFF", fill="FF4472C4", align="center", wrap=True)
    normal_style = styles.xf_id()
    node_style = styles.xf_id(bold=True, fill="FFDCE6F1")

    ncols = len(header)
    sheet_xml_rows = []

    header_cells = "".join(
        _xlsx_cell_xml(c, 1, header[c], header_style) for c in range(ncols)
    )
    sheet_xml_rows.append('<row r="1" ht="30" customHeight="1">{}</row>'.format(header_cells))

    last_row_num = 1
    for row_idx, row in enumerate(data_rows, start=2):
        level = row["level"]
        is_node = row["is_node"]
        cells = row["cells"]

        cells_xml = []
        for c in range(ncols):
            value = cells[c] if c < len(cells) else None
            if c == name_col_index:
                indent = min(max(level - 1, 0), 14)
                style_id = styles.xf_id(bold=is_node, fill="FFDCE6F1" if is_node else None, indent=indent)
            else:
                style_id = node_style if is_node else normal_style
            cells_xml.append(_xlsx_cell_xml(c, row_idx, value, style_id))

        outline_level = min(max(level - 1, 0), 7)
        sheet_xml_rows.append(
            '<row r="{}" outlineLevel="{}">{}</row>'.format(row_idx, outline_level, "".join(cells_xml))
        )
        last_row_num = row_idx

    last_col_letters = _col_index_to_letters(ncols - 1)
    dimension_ref = "A1:{}{}".format(last_col_letters, last_row_num)

    col_widths = [10, 10, 20, 55, 10, 10, 30, 20, 12, 20, 22, 22, 25]
    cols_xml = "".join(
        '<col min="{0}" max="{0}" width="{1}" customWidth="1"/>'.format(
            i + 1, col_widths[i] if i < len(col_widths) else 15
        )
        for i in range(ncols)
    )

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetPr><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr>'
        '<dimension ref="{dimension}"/>'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
        'activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<cols>{cols}</cols>'
        '<sheetData>{rows}</sheetData>'
        '<autoFilter ref="{dimension}"/>'
        "</worksheet>"
    ).format(dimension=dimension_ref, cols=cols_xml, rows="".join(sheet_xml_rows))

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        "</Types>"
    )

    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Состав изделия" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )

    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
        "</Relationships>"
    )

    if os.path.exists(output_path):
        os.remove(output_path)

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/styles.xml", styles.to_xml())
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)

    return output_path


def build_structured_workbook(source_path, output_path):
    """Читает плоскую выгрузку Компаса и сохраняет наглядный xlsx с деревом (без сторонних библиотек)."""
    rows = _read_bom_source_rows(source_path)
    rows = _build_bom_tree_rows(rows)

    header = [title for title, _ in BOM_OUTPUT_COLUMNS]
    name_col_index = [i for i, (_, key) in enumerate(BOM_OUTPUT_COLUMNS) if key == "Наименование"][0]

    data_rows = []
    for row in rows:
        level = row.get("level", 1)
        cells = []
        for _, key in BOM_OUTPUT_COLUMNS:
            if key == "level":
                cells.append(level)
            else:
                value = _bom_clean(row.get(key, ""))
                cells.append(value if value != "" else None)
        data_rows.append({"level": level, "is_node": row.get("_is_node", False), "cells": cells})

    return _write_xlsx(output_path, header, data_rows, name_col_index)


def _log(msg):
    print(msg)


def connect_kompas():
    """Подключение к уже запущенному Компас-3D (COM)."""
    import win32com.client

    last_error = None
    for prog_id in ("Kompas.Application.7", "Kompas.Application.5", "KOMPAS.Application.5"):
        try:
            _log("Подключаемся через {}...".format(prog_id))
            app = win32com.client.Dispatch(prog_id)
            _log("  ✓ Подключено через {}".format(prog_id))
            return app
        except Exception as e:
            last_error = e
            _log("  ✗ Не удалось: {}".format(e))

    raise RuntimeError("Не удалось подключиться к Компас-3D: {}".format(last_error))


def get_active_document(app):
    """Получение активного документа."""
    try:
        doc = app.ActiveDocument
        if doc:
            return doc
    except Exception as e:
        _log("ActiveDocument недоступен: {}".format(e))

    try:
        docs = app.Documents
        if docs and docs.Count > 0:
            return docs.Item(0)
    except Exception as e:
        _log("Documents недоступен: {}".format(e))

    raise RuntimeError("Активный документ не найден. Откройте модель/сборку в Компасе.")


def get_document_path(doc):
    """Полный путь к файлу документа — несколько вариантов на случай разных версий API."""
    for attr in ("PathName", "Path", "FullFileName", "FileName"):
        try:
            value = getattr(doc, attr)
            if value:
                _log("  Путь получен через {}: {}".format(attr, value))
                return value
        except Exception:
            continue
    raise RuntimeError(
        "Не удалось получить путь к файлу активного документа. "
        "Сохраните файл на диск (он должен иметь путь) и повторите."
    )


def find_specification_document(app, active_doc):
    """Возвращает документ спецификации: сам активный документ либо один из открытых."""
    try:
        if getattr(active_doc, "DocumentType", None) == DOCUMENT_TYPE_SPECIFICATION:
            _log("Активный документ уже является спецификацией.")
            return active_doc
    except Exception as e:
        _log("Не удалось определить тип активного документа: {}".format(e))

    try:
        docs = app.Documents
        for i in range(docs.Count):
            d = docs.Item(i)
            try:
                if getattr(d, "DocumentType", None) == DOCUMENT_TYPE_SPECIFICATION:
                    _log("Найдена открытая спецификация: {}".format(getattr(d, "Name", "")))
                    return d
            except Exception:
                continue
    except Exception as e:
        _log("Не удалось перебрать открытые документы: {}".format(e))

    return None


def export_specification_to_flat_file(spec_doc, temp_path):
    """Сохраняет спецификацию во временный excel-файл штатными средствами Компаса."""
    attempts = [
        ("SaveAs(путь)", lambda: spec_doc.SaveAs(temp_path)),
        ("SaveAs(путь, 47)", lambda: spec_doc.SaveAs(temp_path, 47)),
        ("Export(путь, 47)", lambda: spec_doc.Export(temp_path, 47)),
        ("SaveAsToFormat(путь)", lambda: spec_doc.SaveAsToFormat(temp_path)),
    ]

    for name, action in attempts:
        try:
            _log("  Попытка экспорта: {}...".format(name))
            action()
            if os.path.exists(temp_path):
                _log("  ✓ Экспорт удался: {}".format(name))
                return True
        except Exception as e:
            _log("  ✗ {} не удалась: {}".format(name, e))

    return False


def run_macro():
    print()
    print("=" * 70)
    print("КОМПАС-3D — Выгрузка структурированного состава изделия")
    print("=" * 70)
    print()

    try:
        app = connect_kompas()
        active_doc = get_active_document(app)

        print()
        _log("Определяем папку активной модели...")
        model_path = get_document_path(active_doc)
        folder = os.path.dirname(model_path)
        base_name = os.path.splitext(os.path.basename(model_path))[0]

        print()
        _log("Ищем документ спецификации...")
        spec_doc = find_specification_document(app, active_doc)
        if spec_doc is None:
            raise RuntimeError(
                "Документ спецификации не найден среди открытых. "
                "Откройте спецификацию сборки в Компасе (она должна быть "
                "среди открытых документов) и запустите макрос ещё раз."
            )

        temp_flat_path = os.path.join(folder, "~{}_temp_export.xlsx".format(base_name))
        output_path = os.path.join(folder, "{}_состав_изделия.xlsx".format(base_name))

        print()
        _log("Экспортируем спецификацию во временный файл...")
        ok = export_specification_to_flat_file(spec_doc, temp_flat_path)
        if not ok:
            raise RuntimeError(
                "Не удалось экспортировать спецификацию автоматически. "
                "Сохраните её вручную через меню Компаса (Файл -> Сохранить как -> MS Excel) "
                "и запустите kompas_bom_structure.py на полученном файле."
            )

        print()
        _log("Строим структурированную таблицу...")
        build_structured_workbook(temp_flat_path, output_path)

        try:
            os.remove(temp_flat_path)
        except Exception as e:
            _log("Не удалось удалить временный файл {}: {}".format(temp_flat_path, e))

        print()
        print("=" * 70)
        print("✓ ГОТОВО")
        print("Файл сохранён: {}".format(output_path))
        print("=" * 70)
        print()

        return output_path

    except Exception as e:
        print()
        print("=" * 70)
        print("✗ ОШИБКА")
        print("=" * 70)
        print(str(e))
        import traceback
        traceback.print_exc()
        return None


if __name__ == "__main__":
    run_macro()
