# -*- coding: utf-8 -*-
"""
Макрос для Компас-3D: выгружает состав изделия из активно открытой
3D-модели (сборки) и сохраняет наглядную СТРУКТУРИРОВАННУЮ таблицу
(с отступами по уровням вложенности и группировкой строк) рядом с
открытым файлом модели — в ту же папку.

Что делает:
  1. Подключается к запущенному Компас-3D и находит активный документ.
  2. Берёт его TopPart (верхнюю деталь/сборку) и рекурсивно обходит
     дерево состава через IPart7.PartsEx — то есть структуру берёт
     напрямую из 3D-модели, БЕЗ файла спецификации (.spw) и без каких-
     либо команд меню.
  3. Для каждой позиции записывает: уровень вложенности, Наименование
     (Name), Обозначение (Marking), Материал (Material) — свойства
     детали, подтверждённые разведкой API как реально работающие в
     этой версии Компаса. Количество считается САМИ: сколько раз
     одинаковая деталь (по обозначению+наименованию+материалу)
     встречается среди прямых потомков одного узла — метод
     InstanceCount оказался методом с неясной сигнатурой аргументов,
     вызов которого не срабатывал и оставлял столбец пустым.
  4. Строит таблицу с отступами по вложенности и группировкой строк
     (Excel: можно сворачивать/разворачивать узлы) и сохраняет как
     "<имя_модели>_состав_изделия.xlsx" в папку с открытой моделью.

ВАЖНОЕ ОГРАНИЧЕНИЕ: раздел спецификации ("Стандартные изделия" /
"Прочие изделия"), формат листа, позиция по спецификации, обозначение
стандарта и типоразмер крепежа в дереве 3D-модели не хранятся — эти
данные есть только в документе спецификации. Здесь их нет.

Требования:
  - pywin32 (win32com) — связь с Компасом. Больше НИЧЕГО стороннего не
    требуется: запись .xlsx сделана на чистой стандартной библиотеке
    Python (zipfile + xml), потому что во встроенном интерпретаторе
    Компаса нет pip/openpyxl.

БЕЗОПАСНОСТЬ: макрос только ЧИТАЕТ свойства объектов модели (Name,
Marking, Material, PartsEx) — он ни разу не вызывает
SaveAs/Export/Save ни на одном документе, поэтому не может случайно
переключить или испортить открытый у вас файл (в отличие от прежних
версий, где применялся SaveAs документа спецификации).

ВАЖНО про надёжность: точное количество и порядок аргументов метода
PartsEx в разных версиях API Компаса может отличаться, поэтому макрос
пробует несколько сигнатур вызова по очереди и запоминает первую
рабочую — в консоли будет видно, какая именно сработала. Если ни одна
не сработает или числа получатся подозрительными (пусто там, где
должны быть детали, или наоборот — тысячи строк) — пришлите вывод
консоли.
"""

import os
import re
import zipfile
from xml.sax.saxutils import escape as _xml_escape


def _log(msg):
    print(msg)


# ---------------------------------------------------------------------------
# Подключение к Компасу и активному документу
# ---------------------------------------------------------------------------

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


def get_top_part(doc):
    """Верхняя деталь/сборка документа (IPart7) через IKompasDocument3D."""
    import win32com.client

    try:
        doc3d = win32com.client.CastTo(doc, "IKompasDocument3D")
    except Exception as e:
        raise RuntimeError("Не удалось привести документ к IKompasDocument3D: {}".format(e))

    try:
        top_part = doc3d.TopPart
    except Exception as e:
        raise RuntimeError("Не удалось получить TopPart: {}".format(e))

    if top_part is None:
        raise RuntimeError("TopPart пуст — похоже, активный документ не является 3D-сборкой/деталью.")

    return top_part


# ---------------------------------------------------------------------------
# Обход дерева состава изделия через IPart7.PartsEx
# ---------------------------------------------------------------------------

# Предпочтение отдаём (False, False) — по смыслу имён аргументов
# (bAllPartsInPart, bAllPartsInAssembly) это должно означать "только
# непосредственные дочерние детали", что и нужно для ручной рекурсии
# (иначе при "плоском" режиме получим одинаковые данные на каждом
# уровне рекурсии и задвоим/растиражируем позиции). Остальные варианты
# — на случай, если в вашей версии API сигнатура иная.
_PARTSEX_SIGNATURES = [
    (False, False),
    (False,),
    (),
    (True, False),
    (False, True),
    (True, True),
    (True,),
    (0, 0),
    (1,),
]

_working_signature = {"sig": None}


def _normalize_parts_result(result):
    """Приводит результат PartsEx к обычному python-списку."""
    if result is None:
        return []
    if isinstance(result, (list, tuple)):
        return list(result)
    # COM-коллекция вида Count/Item(i)
    try:
        count = result.Count
        return [result.Item(i) for i in range(count)]
    except Exception:
        pass
    try:
        return list(result)
    except Exception:
        return []


def get_child_parts(part):
    """
    Возвращает непосредственные дочерние детали/подсборки данной детали
    через PartsEx, перебирая рабочие сигнатуры вызова (см. выше) и
    запоминая первую удачную для последующих вызовов.
    """
    sig = _working_signature["sig"]
    if sig is not None:
        try:
            return _normalize_parts_result(part.PartsEx(*sig))
        except Exception as e:
            _log("  Запомненная сигнатура PartsEx{} перестала работать: {} — подбираем заново.".format(sig, e))
            _working_signature["sig"] = None

    for sig in _PARTSEX_SIGNATURES:
        try:
            result = part.PartsEx(*sig)
            normalized = _normalize_parts_result(result)
            _working_signature["sig"] = sig
            _log("  ✓ Рабочая сигнатура PartsEx{}, дочерних элементов: {}".format(sig, len(normalized)))
            return normalized
        except Exception:
            continue

    _log("  ✗ Ни одна сигнатура PartsEx не сработала для этой детали.")
    return []


def _safe_attr(obj, name, default=""):
    try:
        value = getattr(obj, name)
        if callable(value):
            value = value()
        return value
    except Exception:
        return default


MAX_DEPTH = 15
MAX_ROWS = 2000


def _group_key(child):
    """
    Ключ группировки одинаковых деталей среди прямых потомков одного узла:
    обозначение + наименование + материал. Используется вместо InstanceCount
    (тоже метод с неясной сигнатурой аргументов, вызов которого не
    срабатывал и оставлял столбец "Количество" пустым) — количество
    считаем сами: сколько раз одна и та же деталь встречается среди
    прямых потомков. Это и есть локальное количество "в этом узле",
    которое и нужно в спецификации (например, "Втулка ... Количество: 2"
    одной строкой, а не двумя одинаковыми строками).
    """
    marking = _bom_clean(_safe_attr(child, "Marking", ""))
    name = _bom_clean(_safe_attr(child, "Name", ""))
    material = _bom_clean(_safe_attr(child, "Material", ""))
    return (marking, name, material)


_STANDARD_MARKING_RE = re.compile(r"(ГОСТ|ОСТ|ТУ|DIN|ISO|EN\s?\d|ANSI)", re.IGNORECASE)

# Порядок категорий внутри узла (чем меньше число, тем выше в таблице):
# сначала сборочные единицы/подузлы, потом собственные детали, потом
# стандартные изделия (крепёж и т.п. со ссылкой на ГОСТ/DIN/...), и в
# самом низу — ПКИ (покупные комплектующие без собственного обозначения).
CATEGORY_NODE = 0
CATEGORY_DETAIL = 1
CATEGORY_STANDARD = 2
CATEGORY_PKI = 3


def _classify(marking, has_children):
    if has_children:
        return CATEGORY_NODE
    if not marking:
        return CATEGORY_PKI
    if _STANDARD_MARKING_RE.search(marking):
        return CATEGORY_STANDARD
    return CATEGORY_DETAIL


def _group_children(children):
    """
    Группирует список деталей по _group_key (получая заодно детей
    каждого представителя — пригодится и для классификации, и чтобы не
    вызывать PartsEx на них повторно при рекурсии), затем сортирует:
    сначала по категории (узлы/подузлы -> детали -> стандартные изделия
    -> ПКИ, как в спецификации), внутри категории — по обозначению
    (А-Я), а если оно пустое (обычно у ПКИ) — по наименованию.

    Возвращает список (представитель, количество, дочерние_детали).
    """
    order = []
    groups = {}
    for child in children:
        key = _group_key(child)
        if key not in groups:
            groups[key] = {"representative": child, "count": 0}
            order.append(key)
        groups[key]["count"] += 1

    result = []
    for key in order:
        representative = groups[key]["representative"]
        count = groups[key]["count"]
        grandchildren = get_child_parts(representative)
        marking = _bom_clean(_safe_attr(representative, "Marking", ""))
        name = _bom_clean(_safe_attr(representative, "Name", ""))
        category = _classify(marking, bool(grandchildren))
        result.append((representative, count, grandchildren, category, marking, name))

    result.sort(key=lambda item: (item[3], item[4].lower(), item[5].lower()))
    return [(r, c, gc) for r, c, gc, _cat, _mk, _nm in result]


def collect_bom_rows(part, level=1, rows=None, visited=None, quantity=1, children=None):
    """
    Рекурсивно собирает строки состава изделия из дерева PartsEx.
    `children`, если передан, — уже полученный список дочерних деталей
    этой детали (чтобы не запрашивать PartsEx повторно: он уже был
    вызван в _group_children родителя при классификации/сортировке).
    """
    if rows is None:
        rows = []
    if visited is None:
        visited = set()

    unique_num = _safe_attr(part, "UniqueNum", None)
    if unique_num is not None:
        if unique_num in visited:
            _log("  Пропускаем повтор (защита от циклов), UniqueNum={}".format(unique_num))
            return rows
        visited.add(unique_num)

    name = _bom_clean(_safe_attr(part, "Name", ""))
    marking = _bom_clean(_safe_attr(part, "Marking", ""))
    material = _bom_clean(_safe_attr(part, "Material", ""))

    rows.append({
        "level": level,
        "Наименование": name,
        "Обозначение": marking,
        "Материал": material,
        "Количество": quantity,
    })

    if len(rows) >= MAX_ROWS:
        _log("  ! Достигнут предел в {} строк — останавливаем обход (защита от зацикливания).".format(MAX_ROWS))
        return rows

    if level >= MAX_DEPTH:
        _log("  ! Достигнута максимальная глубина {} — дальше не углубляемся.".format(MAX_DEPTH))
        return rows

    if children is None:
        children = get_child_parts(part)

    for representative, count, grandchildren in _group_children(children):
        collect_bom_rows(representative, level + 1, rows, visited, quantity=count, children=grandchildren)

    return rows


def _build_tree_flags(rows):
    """Помечает строки, у которых есть дочерние (для стиля 'узел')."""
    for i, row in enumerate(rows):
        level = row["level"]
        is_node = i + 1 < len(rows) and rows[i + 1]["level"] > level
        row["_is_node"] = is_node
    return rows


def _bom_clean(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


# ---------------------------------------------------------------------------
# Запись .xlsx без сторонних библиотек (zipfile + xml)
# ---------------------------------------------------------------------------

BOM_OUTPUT_COLUMNS = [
    ("Уровень", "level"),
    ("Обозначение", "Обозначение"),
    ("Наименование", "Наименование"),
    ("Количество", "Количество"),
    ("Материал", "Материал"),
]


def _col_index_to_letters(index):
    """0 -> 'A', 1 -> 'B', ... 26 -> 'AA' ..."""
    index += 1
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


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

    col_widths = [10, 20, 55, 12, 30]
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


def build_structured_workbook_from_rows(rows, output_path):
    """Пишет наглядный xlsx с деревом прямо из уже собранных строк состава изделия."""
    rows = _build_tree_flags(rows)

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


# ---------------------------------------------------------------------------
# Основной сценарий
# ---------------------------------------------------------------------------

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
        _log("Получаем верхнюю деталь (TopPart)...")
        top_part = get_top_part(active_doc)
        top_name = _safe_attr(top_part, "Name", "?")
        _log("  Верхняя деталь: {}".format(top_name))

        print()
        _log("Обходим дерево состава изделия (PartsEx)...")
        rows = collect_bom_rows(top_part)
        _log("Собрано строк: {}".format(len(rows)))

        output_path = os.path.join(folder, "{}_состав_изделия.xlsx".format(base_name))

        print()
        _log("Строим структурированную таблицу...")
        build_structured_workbook_from_rows(rows, output_path)

        print()
        print("=" * 70)
        print("✓ ГОТОВО")
        print("Файл сохранён: {}".format(output_path))
        print("Всего строк: {}".format(len(rows)))
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
