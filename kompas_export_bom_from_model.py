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


_working_mass_accessor = {"name": None, "logged": False}
_MASS_ACCESSOR_CANDIDATES = ("Mass", "GetMassCentrProperties", "MassInertiaParams")


def _try_mass_accessor(part, name):
    try:
        attr = getattr(part, name)
        result = attr() if callable(attr) else attr
    except Exception:
        return None
    if isinstance(result, bool):
        return None
    if isinstance(result, (int, float)):
        return float(result)
    mass_attr = getattr(result, "Mass", None)
    if isinstance(mass_attr, (int, float)) and not isinstance(mass_attr, bool):
        return float(mass_attr)
    return None


def get_mass(part):
    """
    Пытается получить массу ОДНОЙ детали. В отличие от Name/Marking/
    Material/PartsEx это свойство НЕ подтверждено разведкой API —
    пробуем несколько вероятных вариантов и запоминаем первый рабочий,
    как и для PartsEx. Если ни один не сработает — возвращаем None
    (столбцы "Масса"/"Масса общая" останутся пустыми) и один раз
    печатаем предупреждение в консоль, а не на каждую деталь.
    """
    name = _working_mass_accessor["name"]
    if name is not None:
        value = _try_mass_accessor(part, name)
        if value is not None:
            return value
        _working_mass_accessor["name"] = None

    for candidate in _MASS_ACCESSOR_CANDIDATES:
        value = _try_mass_accessor(part, candidate)
        if value is not None:
            _working_mass_accessor["name"] = candidate
            _log("  ✓ Масса получена через {}".format(candidate))
            return value

    if not _working_mass_accessor["logged"]:
        _log(
            "  ! Не удалось получить массу ни одним из способов {} — "
            "столбцы Масса останутся пустыми для деталей без неё. Это "
            "непроверенная часть API — если масса нужна, пришлите "
            "вывод консоли, подберём точный вызов.".format(_MASS_ACCESSOR_CANDIDATES)
        )
        _working_mass_accessor["logged"] = True
    return None


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


_PROFILE_PIPE_KEYWORDS = ("профил", "квадрат", "прямоугольн")


def _is_turning_material(material_clean):
    """
    Материал идёт на токарную обработку: "Круг..." — всегда; "Труба..."
    — только круглая (не профильная/квадратная/прямоугольная — такая
    труба идёт в обычный металл, см. согласованное правило).
    """
    if not material_clean:
        return False
    text = material_clean.lower()
    if text.startswith("круг"):
        return True
    if text.startswith("труба"):
        return not any(kw in text for kw in _PROFILE_PIPE_KEYWORDS)
    return False


def _clean_material_display(raw_material):
    """
    Приводит текст материала к читаемому виду:
      - "$d" в середине строки меняем на пробел (служебный артефакт
        формата материала/заготовки в Компасе);
      - висящий "$" в конце строки просто убираем;
      - "Без указания материала", "Сталь", "Сталь 10" считаем
        отсутствием полезной информации — возвращаем пустую строку.
    """
    text = _bom_clean(raw_material)
    if not text:
        return ""
    text = text.replace("$d", " ")
    if text.endswith("$"):
        text = text[:-1]
    text = re.sub(r"\s+", " ", text).strip()
    if text in ("Без указания материала", "Сталь", "Сталь 10"):
        return ""
    return text


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


def collect_bom_rows(part, level=1, rows=None, visited=None, quantity=1, children=None, total_quantity=1):
    """
    Рекурсивно собирает строки состава изделия из дерева PartsEx.

    `children`, если передан, — уже полученный список дочерних деталей
    этой детали (чтобы не запрашивать PartsEx повторно: он уже был
    вызван в _group_children родителя при классификации/сортировке).

    `total_quantity` — сквозное количество этой позиции по всему
    изделию (произведение количеств вдоль всей цепочки родителей), а
    не только "локально" внутри своего непосредственного узла. Нужно
    для сводной ведомости на втором листе (см. build_summary_rows).
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

    if children is None:
        children = get_child_parts(part)

    name = _bom_clean(_safe_attr(part, "Name", ""))
    marking = _bom_clean(_safe_attr(part, "Marking", ""))
    material = _clean_material_display(_safe_attr(part, "Material", ""))
    mass = get_mass(part)

    rows.append({
        "level": level,
        "Наименование": name,
        "Обозначение": marking,
        "Материал": material,
        "Количество": quantity,
        "Масса": mass,
        "_is_node": bool(children),
        "_total_quantity": total_quantity,
    })

    if len(rows) >= MAX_ROWS:
        _log("  ! Достигнут предел в {} строк — останавливаем обход (защита от зацикливания).".format(MAX_ROWS))
        return rows

    if level >= MAX_DEPTH:
        _log("  ! Достигнута максимальная глубина {} — дальше не углубляемся.".format(MAX_DEPTH))
        return rows

    for representative, count, grandchildren in _group_children(children):
        collect_bom_rows(
            representative, level + 1, rows, visited,
            quantity=count, children=grandchildren, total_quantity=total_quantity * count,
        )

    return rows


def build_summary_rows(rows):
    """
    Сводная ведомость: одна строка на каждую уникальную позицию
    (обозначение+наименование+материал), собранная СО ВСЕХ узлов и
    подузлов — количество суммируется по всему изделию (across всех
    мест, где эта позиция встречается), с учётом умножения количества
    вверх по цепочке вложенности. Верхний узел (само изделие,
    level == 1) в сводку не включается.

    Одинаковые позиции не дублируются строками — суммируются в одну.
    Порядок — тот же, что и на листе "Состав изделия": сначала
    подузлы, затем собственные детали, стандартные изделия и ПКИ
    внизу, внутри категории — по обозначению (А-Я).
    """
    order = []
    groups = {}
    for row in rows:
        if row["level"] == 1:
            continue
        key = (row["Обозначение"], row["Наименование"], row["Материал"])
        if key not in groups:
            groups[key] = {
                "Обозначение": row["Обозначение"],
                "Наименование": row["Наименование"],
                "Материал": row["Материал"],
                "Количество": 0,
                "Масса": row.get("Масса"),
                "_is_node": row["_is_node"],
            }
            order.append(key)
        groups[key]["Количество"] += row["_total_quantity"]

    result = [groups[key] for key in order]
    result.sort(key=lambda item: (
        _classify(item["Обозначение"], item["_is_node"]),
        item["Обозначение"].lower(),
        item["Наименование"].lower(),
    ))
    for item in result:
        item["Масса общая"] = item["Масса"] * item["Количество"] if item["Масса"] is not None else None
    return result


def build_pki_rows(summary_items):
    """Позиции без собственного обозначения (ПКИ) — для листа 'Прочие изделия'."""
    return [
        item for item in summary_items
        if not item["_is_node"] and _classify(item["Обозначение"], False) == CATEGORY_PKI
    ]


def build_standard_rows(summary_items):
    """Стандартные изделия (крепёж и т.п. со ссылкой на ГОСТ/DIN/...) — для листа 'Метиз'."""
    return [
        item for item in summary_items
        if not item["_is_node"] and _classify(item["Обозначение"], False) == CATEGORY_STANDARD
    ]


def build_turning_parts_rows(summary_items):
    """Собственные детали, изготавливаемые из круглого проката/трубы — для листа 'Токарка'."""
    return [
        item for item in summary_items
        if not item["_is_node"]
        and _classify(item["Обозначение"], False) == CATEGORY_DETAIL
        and _is_turning_material(item["Материал"])
    ]


def build_material_rows(summary_items, turning):
    """
    Сводка по материалам собственных деталей (без узлов/крепежа/ПКИ):
    turning=True — материалы для токарной обработки (лист "Токарка
    материал"), turning=False — все остальные (лист "Металл").
    Позиции без известного материала (пустое после очистки) в сводку
    по материалам не попадают — там нечего суммировать.
    """
    order = []
    groups = {}
    for item in summary_items:
        if item["_is_node"]:
            continue
        if _classify(item["Обозначение"], False) != CATEGORY_DETAIL:
            continue
        material = item["Материал"]
        if not material:
            continue
        if _is_turning_material(material) != turning:
            continue
        if material not in groups:
            groups[material] = {"Обозначение": material, "Масса": 0.0, "_has_mass": False}
            order.append(material)
        if item["Масса общая"] is not None:
            groups[material]["Масса"] += item["Масса общая"]
            groups[material]["_has_mass"] = True

    result = [groups[key] for key in order]
    result.sort(key=lambda r: r["Обозначение"].lower())
    for r in result:
        if r["_has_mass"]:
            r["Масса+30%"] = r["Масса"] * 1.3
        else:
            r["Масса"] = None
            r["Масса+30%"] = None
    return result


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

# Столбцы, которые должны попадать в ячейку как настоящее число Excel
# (а не текст) — чтобы с ними можно было считать/суммировать.
_NUMERIC_KEYS = {"level", "Количество", "Масса", "Масса+30%", "Масса общая"}


def _cell_value(value, key):
    """Готовит значение к записи в ячейку: число — как есть (округлённое), текст — через _bom_clean."""
    if key in _NUMERIC_KEYS:
        if value is None or value == "":
            return None
        if isinstance(value, float):
            return round(value, 3)
        return value
    text = _bom_clean(value)
    return text if text != "" else None


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


def _build_sheet_xml(header, data_rows, name_col_index, col_widths, styles, tree_style):
    """
    Строит XML одного листа (<worksheet>...). data_rows — список словарей
    {"level": int, "is_node": bool, "cells": [...]}. Если tree_style
    выключен — пишутся обычные плоские строки, без отступов и
    группировки (для сводной ведомости).
    """
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
        level = row.get("level", 1)
        is_node = row.get("is_node", False)
        cells = row["cells"]

        cells_xml = []
        for c in range(ncols):
            value = cells[c] if c < len(cells) else None
            if tree_style and c == name_col_index:
                indent = min(max(level - 1, 0), 14)
                style_id = styles.xf_id(bold=is_node, fill="FFDCE6F1" if is_node else None, indent=indent)
            else:
                style_id = node_style if is_node else normal_style
            cells_xml.append(_xlsx_cell_xml(c, row_idx, value, style_id))

        if tree_style:
            outline_level = min(max(level - 1, 0), 7)
            sheet_xml_rows.append(
                '<row r="{}" outlineLevel="{}">{}</row>'.format(row_idx, outline_level, "".join(cells_xml))
            )
        else:
            sheet_xml_rows.append('<row r="{}">{}</row>'.format(row_idx, "".join(cells_xml)))
        last_row_num = row_idx

    last_col_letters = _col_index_to_letters(ncols - 1)
    dimension_ref = "A1:{}{}".format(last_col_letters, last_row_num)

    cols_xml = "".join(
        '<col min="{0}" max="{0}" width="{1}" customWidth="1"/>'.format(
            i + 1, col_widths[i] if i < len(col_widths) else 15
        )
        for i in range(ncols)
    )

    sheet_pr = '<sheetPr><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr>' if tree_style else ""

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + sheet_pr +
        '<dimension ref="{dimension}"/>'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
        'activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<cols>{cols}</cols>'
        '<sheetData>{rows}</sheetData>'
        '<autoFilter ref="{dimension}"/>'
        "</worksheet>"
    ).format(dimension=dimension_ref, cols=cols_xml, rows="".join(sheet_xml_rows))

    return sheet_xml


def _write_workbook(output_path, sheets):
    """
    Пишет .xlsx с одним или несколькими листами напрямую (zip + XML),
    без openpyxl. `sheets` — список словарей:
    {"name": str, "header": [...], "data_rows": [...], "name_col_index": int,
     "col_widths": [...], "tree_style": bool}.
    """
    styles = _StyleRegistry()

    sheet_xmls = []
    for sheet in sheets:
        sheet_xmls.append(_build_sheet_xml(
            sheet["header"], sheet["data_rows"], sheet["name_col_index"],
            sheet["col_widths"], styles, sheet["tree_style"],
        ))

    content_type_overrides = "".join(
        '<Override PartName="/xl/worksheets/sheet{0}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.format(i + 1)
        for i in range(len(sheets))
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        + content_type_overrides +
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

    sheet_entries = "".join(
        '<sheet name="{}" sheetId="{}" r:id="rId{}"/>'.format(_xml_escape(sheet["name"]), i + 1, i + 1)
        for i, sheet in enumerate(sheets)
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets>{}</sheets>'
        "</workbook>"
    ).format(sheet_entries)

    styles_rid = len(sheets) + 1
    workbook_rels_entries = "".join(
        '<Relationship Id="rId{0}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet{0}.xml"/>'.format(i + 1)
        for i in range(len(sheets))
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + workbook_rels_entries +
        '<Relationship Id="rId{}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'.format(styles_rid) +
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
        for i, sheet_xml in enumerate(sheet_xmls):
            zf.writestr("xl/worksheets/sheet{}.xml".format(i + 1), sheet_xml)

    return output_path


SUMMARY_COLUMNS = [
    ("Обозначение", "Обозначение"),
    ("Наименование", "Наименование"),
    ("Количество", "Количество"),
    ("Материал", "Материал"),
    ("Масса", "Масса"),
    ("Масса общая", "Масса общая"),
]

# "Прочие изделия" / "Гидравлика" / "Метиз" / "Токарка" — списки позиций.
# "Код 1С ЕРП" и "Артикул" в модели Компаса не хранятся (нет источника
# данных) — столбцы есть, но всегда пустые.
ITEMS_COLUMNS = [
    ("Код 1С ЕРП", None),
    ("Артикул", None),
    ("Обозначение", "Обозначение"),
    ("Наименование", "Наименование"),
    ("Количество", "Количество"),
]
ITEMS_COL_WIDTHS = [15, 15, 20, 55, 12]

# "Металл" / "Токарка материал" — сводки по материалу. "Код 1С ЕРП" —
# по той же причине всегда пусто.
MATERIAL_COLUMNS = [
    ("Код 1С ЕРП", None),
    ("Обозначение", "Обозначение"),
    ("Масса", "Масса"),
    ("Масса+30%", "Масса+30%"),
]
MATERIAL_COL_WIDTHS = [15, 30, 12, 14]


def _items_to_data_rows(items, columns):
    """Готовит data_rows для _write_workbook из списка словарей-позиций."""
    data_rows = []
    for item in items:
        cells = [None if key is None else _cell_value(item.get(key), key) for _title, key in columns]
        data_rows.append({"is_node": item.get("_is_node", False), "cells": cells})
    return data_rows


def _empty_sheet(name):
    """Полностью пустой лист (для "Аутсорсинг"/"Наклейки" — пока без структуры)."""
    return {
        "name": name,
        "header": [""],
        "data_rows": [],
        "name_col_index": 0,
        "col_widths": [20],
        "tree_style": False,
    }


def build_structured_workbook_from_rows(rows, output_path):
    """
    Пишет итоговый xlsx с несколькими листами:
      - "Состав изделия" — наглядное дерево с отступами и группировкой;
      - "Сводная ведомость" — плоский список всех позиций по всему
        изделию (одинаковые суммированы), с массой на деталь и общей;
      - "Прочие изделия" / "Метиз" — позиции без обозначения (ПКИ) и
        стандартные изделия соответственно;
      - "Гидравлика" — та же структура, что и "Прочие изделия", но
        пока без данных (переносить в неё пока нечего);
      - "Металл" / "Токарка материал" — сводки по материалу для
        обычных и токарных деталей (масса + масса с запасом 30%);
      - "Токарка" — детали, изготавливаемые из круглого проката/трубы;
      - "Аутсорсинг" / "Наклейки" — оставлены пустыми.
    """
    header = [title for title, _ in BOM_OUTPUT_COLUMNS]
    name_col_index = [i for i, (_, key) in enumerate(BOM_OUTPUT_COLUMNS) if key == "Наименование"][0]

    tree_data_rows = []
    for row in rows:
        level = row.get("level", 1)
        cells = [_cell_value(row.get(key), key) for _title, key in BOM_OUTPUT_COLUMNS]
        tree_data_rows.append({"level": level, "is_node": row.get("_is_node", False), "cells": cells})

    summary_items = build_summary_rows(rows)
    summary_header = [title for title, _ in SUMMARY_COLUMNS]
    summary_data_rows = _items_to_data_rows(summary_items, SUMMARY_COLUMNS)

    pki_data_rows = _items_to_data_rows(build_pki_rows(summary_items), ITEMS_COLUMNS)
    standard_data_rows = _items_to_data_rows(build_standard_rows(summary_items), ITEMS_COLUMNS)
    turning_parts_data_rows = _items_to_data_rows(build_turning_parts_rows(summary_items), ITEMS_COLUMNS)

    metal_data_rows = _items_to_data_rows(build_material_rows(summary_items, turning=False), MATERIAL_COLUMNS)
    turning_material_data_rows = _items_to_data_rows(build_material_rows(summary_items, turning=True), MATERIAL_COLUMNS)

    items_header = [title for title, _ in ITEMS_COLUMNS]
    material_header = [title for title, _ in MATERIAL_COLUMNS]

    sheets = [
        {
            "name": "Состав изделия",
            "header": header,
            "data_rows": tree_data_rows,
            "name_col_index": name_col_index,
            "col_widths": [10, 20, 55, 12, 30],
            "tree_style": True,
        },
        {
            "name": "Сводная ведомость",
            "header": summary_header,
            "data_rows": summary_data_rows,
            "name_col_index": 1,
            "col_widths": [20, 55, 12, 30, 12, 14],
            "tree_style": False,
        },
        {
            "name": "Прочие изделия",
            "header": items_header,
            "data_rows": pki_data_rows,
            "name_col_index": 3,
            "col_widths": ITEMS_COL_WIDTHS,
            "tree_style": False,
        },
        {
            "name": "Гидравлика",
            "header": items_header,
            "data_rows": [],
            "name_col_index": 3,
            "col_widths": ITEMS_COL_WIDTHS,
            "tree_style": False,
        },
        {
            "name": "Метиз",
            "header": items_header,
            "data_rows": standard_data_rows,
            "name_col_index": 3,
            "col_widths": ITEMS_COL_WIDTHS,
            "tree_style": False,
        },
        {
            "name": "Металл",
            "header": material_header,
            "data_rows": metal_data_rows,
            "name_col_index": 1,
            "col_widths": MATERIAL_COL_WIDTHS,
            "tree_style": False,
        },
        {
            "name": "Токарка материал",
            "header": material_header,
            "data_rows": turning_material_data_rows,
            "name_col_index": 1,
            "col_widths": MATERIAL_COL_WIDTHS,
            "tree_style": False,
        },
        {
            "name": "Токарка",
            "header": items_header,
            "data_rows": turning_parts_data_rows,
            "name_col_index": 3,
            "col_widths": ITEMS_COL_WIDTHS,
            "tree_style": False,
        },
        _empty_sheet("Аутсорсинг"),
        _empty_sheet("Наклейки"),
    ]

    return _write_workbook(output_path, sheets)


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
