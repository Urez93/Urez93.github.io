# -*- coding: utf-8 -*-
"""
КОМПАС-3D -> DXF: развёртка активной детали для лазерного раскроя.

Что делает:
  1. Подключается к запущенному КОМПАС-3D (API7; работает и из внешнего
     python.exe, и из встроенного PyScripter).
  2. Берёт АКТИВНЫЙ 3D-документ и путь к нему на диске.
  3. Определяет тип детали:
       - листовое тело -> переводит модель в развёрнутое состояние;
       - твердотельная  -> работает как с плоской пластиной.
  4. Находит наибольшую плоскую грань, снимает её контуры
     (внешний + отверстия) и определяет толщину материала.
  5. Пишет DXF (R12, мм) РЯДОМ С МОДЕЛЬЮ и С ТЕМ ЖЕ ИМЕНЕМ.

DXF пишется собственным писателем, поэтому в файл попадают только контуры
реза (слой CUT) - без рамок, основных надписей, размеров и осевых линий.
Подпись артикула отключена; включается флагом CONFIG["add_article"], текст
тогда идёт отдельным слоем MARK и слоя реза не касается.

Запуск:
    python kompas_unfold_macro.py          - экспорт
    python kompas_unfold_macro.py probe    - только диагностика API

Модель макрос НЕ сохраняет: изменённое состояние развёртки возвращается назад.
"""

from __future__ import print_function

import math
import os
import sys
import traceback


# ============================================================
# НАСТРОЙКИ
# ============================================================

CONFIG = {
    # Слои DXF
    "cut_layer": "CUT",           # контуры реза
    "mark_layer": "MARK",         # артикул детали

    # Артикул (по умолчанию выключен — в DXF идут только контуры реза)
    "add_article": False,         # писать артикул в DXF
    "article_height": 6.0,        # высота текста, мм
    "article_gap": 6.0,           # отступ текста от контура, мм

    # Геометрия
    "tolerance": 0.02,            # точность аппроксимации кривых, мм
    "plane_tolerance": 0.01,      # допуск на плоскостность грани, мм
    "weld_tolerance": 0.05,       # допуск на стыковку рёбер в контур, мм
    "min_segment": 1e-4,          # короче этого сегменты выбрасываются, мм

    # Файл
    "encoding": "cp1251",         # кодировка DXF (кириллица в артикуле)
    "zero_origin": True,          # сдвинуть контур в первый квадрант от (0,0)
    "overwrite": True,            # перезаписывать существующий DXF

    # Поведение
    "unfold_sheet_metal": True,   # разворачивать листовое тело автоматически
    "probe_on_error": True,       # при сбое печатать дамп API
    "mode": "",                   # "probe" — только диагностика, без экспорта
}


# ============================================================
# ЛОГИРОВАНИЕ
# ============================================================

def log(msg=""):
    print(msg)


def step(msg):
    print("\n>>> {}".format(msg))


def ok(msg):
    print("  [OK] {}".format(msg))


def warn(msg):
    print("  [!] {}".format(msg))


def err(msg):
    print("  [ОШИБКА] {}".format(msg))


# ============================================================
# ПОДКЛЮЧЕНИЕ К КОМПАС-3D
# ============================================================

API7_GUID = "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}"

_API7_MODULE = [None]


def api7_module():
    """Сгенерированный модуль API7 (нужен для приведения интерфейсов)."""
    if _API7_MODULE[0] is None:
        try:
            from win32com.client import gencache
            _API7_MODULE[0] = gencache.EnsureModule(API7_GUID, 0, 1, 0)
        except Exception:
            _API7_MODULE[0] = False
    return _API7_MODULE[0] or None


def qi(obj, *names, **kwargs):
    """
    Приведение COM-объекта к конкретному интерфейсу API7.

    В API7 многое доступно не через свойства, а через QueryInterface:
    ActiveDocument отдаёт IKompasDocument без TopPart, у IPart7 нет тел —
    они лежат в IModelContainer, а листовые операции в ISheetMetalContainer.

    strict=False (по умолчанию) — при неудаче вернуть исходный объект,
    strict=True — вернуть None (нужно, когда сам факт приведения является
    признаком, например наличия листового тела).
    """
    strict = kwargs.get("strict", False)
    fallback = None if strict else obj

    if obj is None:
        return None
    module = api7_module()
    if module is None:
        return fallback

    try:
        import pythoncom
    except ImportError:
        return fallback

    source = getattr(obj, "_oleobj_", obj)
    for name in names:
        interface = getattr(module, name, None)
        if interface is None:
            continue
        try:
            return interface(
                source.QueryInterface(interface.CLSID, pythoncom.IID_IDispatch)
            )
        except Exception:
            continue
    return fallback


def connect_kompas():
    """
    Возвращает (application, api7_module) или (None, None).

    Сначала пробуем объекты, которые КОМПАС кладёт в глобальное пространство
    встроенного Python, затем обычное COM-подключение к запущенной копии.
    """
    step("Подключение к КОМПАС-3D")

    # 1. Встроенный Python КОМПАСа
    for holder in (globals(), getattr(sys.modules.get("__main__"), "__dict__", {})):
        for name in ("Application", "Kompas"):
            obj = holder.get(name)
            if obj is not None and hasattr(obj, "ActiveDocument"):
                ok("используется объект '{}' встроенного Python".format(name))
                return obj, None

    # 2. Внешнее COM-подключение
    try:
        import pythoncom
        from win32com.client import Dispatch, gencache
    except ImportError:
        err("не установлен пакет pywin32 (win32com).")
        log("    Установка:  pip install pywin32")
        return None, None

    try:
        module7 = gencache.EnsureModule(API7_GUID, 0, 1, 0)
        dispatch = Dispatch("Kompas.Application.7")
        application = module7.IApplication(
            dispatch._oleobj_.QueryInterface(
                module7.IApplication.CLSID, pythoncom.IID_IDispatch
            )
        )
        ok("подключение к API7 установлено")
        return application, module7
    except Exception as exc:
        err("не удалось подключиться: {}".format(exc))
        log("    Проверьте, что КОМПАС-3D запущен и разрядность Python")
        log("    совпадает с разрядностью КОМПАСа (обычно 64 бита).")
        return None, None


# ============================================================
# УНИВЕРСАЛЬНЫЙ ДОСТУП К COM-ОБЪЕКТАМ
# ============================================================
#
# Имена части интерфейсов API7 отличаются от версии к версии, поэтому
# доступ к геометрии идёт через перебор кандидатов. Актуальный набор имён
# для конкретной сборки КОМПАСа показывает режим `probe`.

def fetch(obj, names, *args):
    """
    Первое доступное свойство/метод из списка имён.
    Возвращает (имя, значение) или (None, None).
    """
    if obj is None:
        return None, None
    for name in names:
        try:
            value = getattr(obj, name)
        except Exception:
            continue
        try:
            if callable(value):
                value = value(*args)
        except Exception:
            continue
        if value is not None:
            return name, value
    return None, None


def value_of(obj, names, default=None, *args):
    """Значение первого доступного свойства/метода."""
    _, value = fetch(obj, names, *args)
    return default if value is None else value


def as_list(collection):
    """
    Приводит COM-коллекцию КОМПАСа к обычному списку.
    Поддержаны Count+Item, Count+GetByIndex, кортежи и одиночные объекты.
    """
    if collection is None:
        return []
    if isinstance(collection, (list, tuple)):
        return [item for item in collection if item is not None]

    count = None
    for name in ("Count", "GetCount"):
        try:
            raw = getattr(collection, name)
            count = raw() if callable(raw) else raw
            break
        except Exception:
            continue
    if count is None:
        return [collection]

    items = []
    for index in range(int(count)):
        item = None
        for getter in ("Item", "GetByIndex", "GetItem"):
            try:
                accessor = getattr(collection, getter)
            except Exception:
                continue
            try:
                item = accessor(index)
            except Exception:
                item = None
            if item is not None:
                break
        if item is not None:
            items.append(item)
    return items


def describe(obj, title):
    """Печатает доступные члены COM-объекта — для режима probe."""
    log("\n--- {} ---".format(title))
    if obj is None:
        log("    объект отсутствует")
        return
    log("    тип: {}".format(type(obj).__name__))
    names = set()
    for attr in ("_prop_map_get_", "_prop_map_put_"):
        table = getattr(type(obj), attr, None)
        if isinstance(table, dict):
            names.update(table.keys())
    for name in dir(obj):
        if not name.startswith("_"):
            names.add(name)
    if not names:
        log("    члены не определены (поздняя привязка)")
        return
    for name in sorted(names):
        log("    {}".format(name))


# ============================================================
# ВЕКТОРНАЯ АРИФМЕТИКА
# ============================================================

def sub(a, b):
    return (a[0] - b[0], a[1] - b[1], a[2] - b[2])


def add(a, b):
    return (a[0] + b[0], a[1] + b[1], a[2] + b[2])


def scale(a, k):
    return (a[0] * k, a[1] * k, a[2] * k)


def dot(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def cross(a, b):
    return (a[1] * b[2] - a[2] * b[1],
            a[2] * b[0] - a[0] * b[2],
            a[0] * b[1] - a[1] * b[0])


def norm(a):
    return math.sqrt(dot(a, a))


def unit(a):
    length = norm(a)
    if length < 1e-12:
        return None
    return scale(a, 1.0 / length)


# ============================================================
# ДОКУМЕНТ
# ============================================================

def get_active_part(application):
    """Возвращает (документ, деталь, путь_к_файлу) активного 3D-документа."""
    step("Активный документ")

    document = value_of(application, ["ActiveDocument"])
    if document is None:
        err("в КОМПАС-3D нет открытых документов.")
        return None, None, None

    path = value_of(document, ["PathName", "Path", "Name"], "")
    if not path or not os.path.isabs(path):
        err("документ не сохранён на диск, путь неизвестен: '{}'".format(path))
        log("    Сохраните деталь в файл .m3d и повторите.")
        return None, None, None
    ok("файл: {}".format(path))

    # ActiveDocument отдаёт обобщённый IKompasDocument — приводим к 3D.
    document_3d = qi(document, "IKompasDocument3D", "IKompasDocument3D1")
    part = value_of(document_3d, ["TopPart"]) or value_of(document, ["TopPart"])
    if part is None:
        err("не удалось получить деталь из активного документа.")
        log("    Тип документа: {}".format(type(document_3d).__name__))
        log("    Макрос работает с деталями (.m3d); сборки не поддерживаются.")
        return None, None, None
    part = qi(part, "IPart7")

    name = value_of(part, ["Name"], "")
    marking = value_of(part, ["Marking"], "")
    ok("деталь: '{}' обозначение: '{}'".format(name, marking))
    return document_3d, part, path


def get_article(part, path):
    """Артикул детали: Обозначение -> Наименование -> имя файла."""
    for names in (["Marking"], ["Name"]):
        value = value_of(part, names, "")
        if isinstance(value, str) and value.strip():
            return value.strip()
    return os.path.splitext(os.path.basename(path))[0]


# ============================================================
# ЛИСТОВОЕ ТЕЛО И РАЗВЁРТКА
# ============================================================

SHEET_BODY_NAMES = ["SheetMetalBodies", "Bodies", "SheetMetalBody"]
UNFOLD_FLAG_NAMES = ["Unfolded", "Unfold", "IsUnfolded", "UnfoldState",
                     "Unfolding", "IsUnfold"]
UNFOLD_HOLDER_NAMES = ["UnfoldParam", "UnfoldParameters", "Unfolds",
                       "UnfoldParams", "Unfold"]
UNFOLD_PARAMETER_NAMES = ["SheetMetalBendUnfoldParameters",
                          "BendUnfoldParameters", "UnfoldParameters"]


def sheet_container(part):
    """
    Контейнер листовых операций. В API7 это не свойство детали, а отдельный
    интерфейс, к которому деталь приводится через QueryInterface.
    """
    return qi(part, "ISheetMetalContainer", strict=True)


def sheet_bodies(part):
    """Листовые тела детали."""
    container = sheet_container(part)
    if container is None:
        return []
    _, bodies = fetch(container, SHEET_BODY_NAMES)
    return [qi(body, "ISheetMetalBody") for body in as_list(bodies)]


def find_unfold_flag(part):
    """
    Ищет объект и имя свойства, управляющего состоянием развёртки.
    Возвращает (владелец, имя_свойства) или (None, None).
    """
    holders = []
    container = sheet_container(part)
    if container is not None:
        holders.append(container)
        # Параметры развёртки хранятся отдельным объектом контейнера.
        _, parameters = fetch(container, UNFOLD_PARAMETER_NAMES)
        if parameters is not None:
            holders.extend(as_list(parameters))
    for body in sheet_bodies(part):
        holders.append(body)
        _, nested = fetch(body, UNFOLD_HOLDER_NAMES)
        if nested is not None:
            holders.extend(as_list(nested))

    for holder in holders:
        for name in UNFOLD_FLAG_NAMES:
            try:
                current = getattr(holder, name)
            except Exception:
                continue
            if isinstance(current, bool):
                return holder, name
    return None, None


def is_sheet_metal(part):
    """Есть ли у детали листовое тело."""
    return bool(sheet_bodies(part))


def ensure_unfolded(part):
    """
    Переводит листовую деталь в развёрнутое состояние.
    Возвращает (получилось_ли, функция_восстановления).
    """
    step("Тип детали")

    if not is_sheet_metal(part):
        ok("деталь твердотельная — работаем как с плоской пластиной")
        return True, None

    ok("деталь листовая")

    if not CONFIG["unfold_sheet_metal"]:
        warn("автоматическая развёртка отключена в CONFIG")
        return True, None

    holder, flag = find_unfold_flag(part)
    if holder is None:
        warn("не удалось найти в API признак развёртки.")
        log("    Разверните деталь вручную (Листовое тело -> Развернуть)")
        log("    и запустите макрос снова: контур будет снят с развёртки.")
        return True, None

    try:
        previous = bool(getattr(holder, flag))
    except Exception:
        previous = False

    if previous:
        ok("модель уже развёрнута")
        return True, None

    created = value_of(holder, ["IsCreated"])
    if created is False:
        warn("параметры развёртки в модели не заданы (нет неподвижной грани).")

    try:
        apply_unfold(holder, flag, True, part)
        ok("развёртка включена (свойство '{}')".format(flag))
    except Exception as exc:
        err("не удалось включить развёртку: {}".format(exc))
        log("    Разверните деталь вручную и запустите макрос снова.")
        return False, None

    def restore():
        try:
            apply_unfold(holder, flag, previous, part)
            ok("состояние развёртки возвращено к исходному")
        except Exception as exc:
            warn("не удалось вернуть состояние развёртки: {}".format(exc))

    return True, restore


def apply_unfold(holder, flag, state, part):
    """Меняет состояние развёртки и пересчитывает модель."""
    setattr(holder, flag, state)
    # У параметров развёртки изменения применяет отдельный вызов.
    fetch(holder, ["UpdateParam", "Update", "Apply"])
    rebuild(part)


def rebuild(part):
    """Перестроение детали после смены состояния развёртки."""
    for name in ("Update", "Rebuild", "RebuildModel"):
        try:
            method = getattr(part, name)
        except Exception:
            continue
        try:
            method()
            return
        except Exception:
            continue


# ============================================================
# ГЕОМЕТРИЯ: ГРАНИ, РЁБРА, КРИВЫЕ
# ============================================================

BODY_NAMES = ["Bodies", "BodyCollection", "Body"]
FACE_NAMES = ["FaceCollection", "Faces", "Face"]
LOOP_NAMES = ["LoopCollection", "Loops", "Loop"]
EDGE_NAMES = ["EdgeCollection", "Edges", "OrientedEdges", "GetEdges"]
ORIENTED_EDGE_NAMES = ["Edge", "GetEdge", "BaseEdge"]
CURVE_NAMES = ["GetCurve3D", "Curve3D", "GetCurve"]
DEFINITION_NAMES = ["GetDefinition", "Definition"]


SHEET_OPERATION_NAMES = [
    "SheetMetalBodies", "SheetMetalPlates", "SheetMetalRuledShells",
    "SheetMetalLinearRuledShells", "ConvertsToSheetMetals", "SheetMetalBends",
    "SheetMetalSketchBends", "SheetMetalLineBends", "SheetMetalCuts",
]

MODEL_OPERATION_NAMES = [
    "Extrusions", "Rotateds", "Lofts", "Evolutions", "Booleans", "Cuts",
    "Shells", "SurfaceThickenings", "SplitSolids", "CopiesGeometry",
    "CollectionsGeometry", "FeaturePatterns", "Fillets", "Chamfers",
    "Holes3D", "Ribs", "Inclines", "MacroObjects3D", "UserObjects",
]

BODY_RESULT_NAMES = ["OperationResult", "ResultBody", "ResultBodies", "Body"]


def interface_names(collection_name):
    """
    Имя интерфейса элемента коллекции: SheetMetalBodies -> ISheetMetalBody,
    Extrusions -> IExtrusion, Holes3D -> IHole3D.
    """
    base = collection_name
    if base.endswith("ies"):
        base = base[:-3] + "y"
    elif base.endswith("s3D"):
        base = base[:-3] + "3D"
    elif base.endswith("s"):
        base = base[:-1]
    return ["I" + base, "I" + base + "7"]


def operations(part):
    """
    Операции модели с именами их интерфейсов.

    Коллекции отдают элементы обобщённым IModelObject, у которого нет ни
    OperationResult, ни параметров операции, поэтому каждый элемент нужно
    приводить к интерфейсу своей операции.
    """
    result = []
    sources = (
        (sheet_container(part), SHEET_OPERATION_NAMES),
        (qi(part, "IModelContainer", strict=True), MODEL_OPERATION_NAMES),
    )
    for container, names in sources:
        if container is None:
            continue
        for name in names:
            try:
                collection = getattr(container, name)
            except Exception:
                continue
            for item in as_list(collection):
                result.append(qi(item, *interface_names(name)))
    return result


def get_bodies(part):
    """
    Тела детали.

    У IPart7 коллекции тел нет, а IModelContainer хранит операции, а не тела.
    Поэтому тело берётся из результата любой операции модели: OperationResult
    возвращает то тело, которому операция принадлежит.
    """
    found = []
    keys = set()

    def remember(obj):
        body = qi(obj, "IBody7", strict=True)
        if body is None:
            return
        key = value_of(body, ["Reference"]) or value_of(body, ["Name"], "")
        if key in keys:
            return
        keys.add(key)
        found.append(body)

    for operation in operations(part):
        _, result = fetch(operation, BODY_RESULT_NAMES)
        for item in as_list(result):
            remember(item)

    if not found:
        found = bodies_via_objects(part)

    if found:
        ok("тел в детали: {}".format(len(found)))
    return found


def bodies_via_objects(part):
    """
    Запасной путь: перебор объектов модели через IModelContainer.Objects.
    Свойство может быть и методом с кодом типа объекта — пробуем оба вида.
    """
    container = qi(part, "IModelContainer", strict=True)
    if container is None:
        return []

    candidates = []
    try:
        raw = getattr(container, "Objects")
    except Exception:
        return []

    if callable(raw):
        for code in range(0, 40):
            try:
                collection = raw(code)
            except Exception:
                continue
            items = as_list(collection)
            if items and qi(items[0], "IBody7", strict=True) is not None:
                ok("тела найдены через Objects({})".format(code))
                candidates = items
                break
    else:
        candidates = as_list(raw)

    bodies = [qi(item, "IBody7", strict=True) for item in candidates]
    return [body for body in bodies if body is not None]


def get_faces(body):
    _, faces = fetch(body, FACE_NAMES)
    return [qi(item, "IFace") for item in as_list(faces)]


def get_loops(face):
    """Циклы грани: внешний контур и отверстия, каждый уже упорядочен."""
    _, loops = fetch(face, LOOP_NAMES)
    return [qi(item, "ILoop7") for item in as_list(loops)]


def unwrap_edge(item):
    """IOrientedEdge7 -> IEdge; обычное ребро возвращается как есть."""
    _, inner = fetch(item, ORIENTED_EDGE_NAMES)
    if inner is not None:
        item = inner
    return qi(item, "IEdge")


def get_edges(owner):
    """Рёбра грани или цикла."""
    _, edges = fetch(owner, EDGE_NAMES)
    items = as_list(edges)
    if not items:
        # API5-подобная схема: сначала определение объекта, потом его рёбра.
        _, definition = fetch(owner, DEFINITION_NAMES)
        if definition is not None:
            _, edges = fetch(definition, EDGE_NAMES)
            items = as_list(edges)
    return [unwrap_edge(item) for item in items]


def face_contours(face):
    """
    Контуры грани в 3D. Если API отдаёт циклы — берём их (внешний контур и
    отверстия разделены самим КОМПАСом), иначе сшиваем рёбра по концам.
    """
    contours = []
    for loop in get_loops(face):
        contours.extend(build_loops(get_edges(loop)))
    if contours:
        return contours
    return build_loops(get_edges(face))


def curve_of(edge):
    """Кривая ребра — напрямую или через определение ребра."""
    _, curve = fetch(edge, CURVE_NAMES)
    if curve is not None:
        return curve
    _, definition = fetch(edge, DEFINITION_NAMES)
    if definition is not None:
        _, curve = fetch(definition, CURVE_NAMES)
        return curve
    return None


def point_tuple(value):
    """Приводит результат COM-вызова к (x, y, z)."""
    if value is None:
        return None
    if isinstance(value, (list, tuple)):
        numbers = [v for v in value if isinstance(v, (int, float))]
        if len(numbers) >= 3:
            return (float(numbers[0]), float(numbers[1]), float(numbers[2]))
        return None
    coords = []
    for axis in ("X", "Y", "Z"):
        try:
            raw = getattr(value, axis)
            coords.append(float(raw() if callable(raw) else raw))
        except Exception:
            return None
    return tuple(coords)


def curve_points(curve, count):
    """Равномерная выборка точек кривой; None, если API недоступен."""
    if curve is None:
        return None

    bounds = None
    for name in ("GetParamRange", "ParamRange", "GetParamsRange"):
        try:
            raw = getattr(curve, name)
            bounds = raw() if callable(raw) else raw
        except Exception:
            continue
        if isinstance(bounds, (list, tuple)):
            numbers = [v for v in bounds if isinstance(v, (int, float))]
            if len(numbers) >= 2:
                bounds = (float(numbers[0]), float(numbers[1]))
                break
        bounds = None
    if bounds is None:
        bounds = (0.0, 1.0)

    start, finish = bounds
    if abs(finish - start) < 1e-12:
        return None

    points = []
    for index in range(count + 1):
        parameter = start + (finish - start) * index / float(count)
        sample = None
        for name in ("GetPointOn", "GetPoint", "PointOn"):
            try:
                method = getattr(curve, name)
            except Exception:
                continue
            try:
                sample = point_tuple(method(parameter))
            except Exception:
                sample = None
            if sample is not None:
                break
        if sample is None:
            return None
        points.append(sample)
    return points


def edge_vertices(edge):
    """Концевые точки ребра — запасной путь для прямых отрезков."""
    holders = [edge]
    _, definition = fetch(edge, DEFINITION_NAMES)
    if definition is not None:
        holders.append(definition)

    for holder in holders:
        for name in ("GetVertex", "Vertex"):
            try:
                method = getattr(holder, name)
            except Exception:
                continue
            points = []
            for flag in (True, False):
                try:
                    vertex = method(flag)
                except Exception:
                    vertex = None
                if vertex is None:
                    break
                sample = point_tuple(value_of(vertex, ["GetPoint", "Point"]))
                if sample is None:
                    sample = point_tuple(vertex)
                if sample is None:
                    break
                points.append(sample)
            if len(points) == 2:
                return points
    return None


def edge_polyline(edge, samples=64):
    """
    Ребро как ломаная в 3D. Кривые аппроксимируются точками, прямые
    остаются двумя точками. Возвращает список точек или None.
    """
    points = curve_points(curve_of(edge), samples)
    if points:
        return dedupe(points)
    points = edge_vertices(edge)
    if points:
        return dedupe(points)
    return None


def dedupe(points):
    """Убирает совпадающие соседние точки."""
    result = []
    for point in points:
        if not result or norm(sub(point, result[-1])) > CONFIG["min_segment"]:
            result.append(point)
    return result


# ============================================================
# ПЛОСКОСТЬ ГРАНИ И ПОИСК НАИБОЛЬШЕЙ
# ============================================================

class PlanarFace(object):
    """Плоская грань: точки контуров, плоскость, площадь."""

    def __init__(self, loops, origin, normal, area):
        self.loops = loops        # список списков 3D-точек
        self.origin = origin
        self.normal = normal
        self.area = area


def fit_plane(points):
    """Плоскость по облаку точек: (центр, нормаль) или None."""
    if len(points) < 3:
        return None

    center = scale(
        reduce_sum(points), 1.0 / len(points)
    )

    # Нормаль по методу Ньюэлла — устойчива к почти вырожденным контурам.
    normal = (0.0, 0.0, 0.0)
    for index in range(len(points)):
        current = points[index]
        following = points[(index + 1) % len(points)]
        normal = add(normal, (
            (current[1] - following[1]) * (current[2] + following[2]),
            (current[2] - following[2]) * (current[0] + following[0]),
            (current[0] - following[0]) * (current[1] + following[1]),
        ))
    normal = unit(normal)
    if normal is None:
        return None
    return center, normal


def reduce_sum(points):
    total = (0.0, 0.0, 0.0)
    for point in points:
        total = add(total, point)
    return total


def plane_residual(points, center, normal):
    """Максимальное отклонение точек от плоскости."""
    return max(abs(dot(sub(point, center), normal)) for point in points)


def build_loops(edges):
    """
    Собирает рёбра в замкнутые контуры по совпадению концов.
    Возвращает список списков 3D-точек.
    """
    chains = []
    for edge in edges:
        points = edge_polyline(edge)
        if points and len(points) >= 2:
            chains.append(points)
    if not chains:
        return []

    tolerance = CONFIG["weld_tolerance"]
    loops = []
    pending = list(chains)

    def closed(chain):
        return len(chain) > 2 and norm(sub(chain[0], chain[-1])) <= tolerance

    while pending:
        loop = pending.pop(0)
        changed = True
        while changed and pending and not closed(loop):
            changed = False
            for index, chain in enumerate(pending):
                if norm(sub(loop[-1], chain[0])) <= tolerance:
                    loop = loop + chain[1:]
                elif norm(sub(loop[-1], chain[-1])) <= tolerance:
                    loop = loop + list(reversed(chain))[1:]
                elif norm(sub(loop[0], chain[-1])) <= tolerance:
                    loop = chain[:-1] + loop
                elif norm(sub(loop[0], chain[0])) <= tolerance:
                    loop = list(reversed(chain))[:-1] + loop
                else:
                    continue
                pending.pop(index)
                changed = True
                break
        loops.append(loop)
    return loops


def polygon_area_2d(points):
    """Площадь замкнутого многоугольника (со знаком)."""
    total = 0.0
    for index in range(len(points)):
        x1, y1 = points[index]
        x2, y2 = points[(index + 1) % len(points)]
        total += x1 * y2 - x2 * y1
    return total / 2.0


def project(point, origin, axis_x, axis_y):
    offset = sub(point, origin)
    return (dot(offset, axis_x), dot(offset, axis_y))


def plane_axes(normal):
    """Пара ортов в плоскости; (ex, ey, normal) — правая тройка."""
    reference = min(
        [(1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0)],
        key=lambda axis: abs(dot(axis, normal)),
    )
    axis_x = unit(cross(reference, normal))
    if axis_x is None:
        axis_x = (1.0, 0.0, 0.0)
    axis_y = cross(normal, axis_x)
    return axis_x, axis_y


def analyse_face(face):
    """Плоская грань -> PlanarFace; None, если грань не плоская."""
    return analyse_contours(face_contours(face))


def analyse_contours(loops):
    """Контуры грани -> PlanarFace; None, если они не лежат в одной плоскости."""
    if not loops:
        return None

    points = [point for loop in loops for point in loop]
    plane = fit_plane(max(loops, key=len))
    if plane is None:
        return None
    center, normal = plane

    if plane_residual(points, center, normal) > CONFIG["plane_tolerance"]:
        return None

    axis_x, axis_y = plane_axes(normal)
    areas = []
    for loop in loops:
        flat = [project(point, center, axis_x, axis_y) for point in loop]
        areas.append(abs(polygon_area_2d(flat)))
    if not areas:
        return None

    # Площадь грани = внешний контур минус отверстия.
    area = max(areas) - (sum(areas) - max(areas))
    return PlanarFace(loops, center, normal, area)


def find_plate_faces(bodies):
    """
    Наибольшая плоская грань и параллельная ей грань с другой стороны листа.
    Возвращает (грань, толщина, габарит_по_нормали) или (None, None, None).
    """
    step("Поиск наибольшей плоской грани")

    planar = []
    all_points = []
    total_faces = 0
    for body in bodies:
        faces = get_faces(body)
        total_faces += len(faces)
        for face in faces:
            try:
                loops = face_contours(face)
            except Exception:
                continue
            if not loops:
                continue
            for loop in loops:
                all_points.extend(loop)
            try:
                result = analyse_contours(loops)
            except Exception:
                result = None
            if result is not None and result.area > 0:
                planar.append(result)

    log("  граней просмотрено: {}, плоских: {}".format(total_faces, len(planar)))
    if not planar:
        err("плоских граней не найдено.")
        return None, None, None

    planar.sort(key=lambda item: item.area, reverse=True)
    best = planar[0]
    ok("наибольшая грань: площадь {:.2f} мм2".format(best.area))

    # Противоположная грань листа: параллельная, той же площади.
    opposite = None
    for candidate in planar[1:]:
        if abs(abs(dot(candidate.normal, best.normal)) - 1.0) > 1e-3:
            continue
        if abs(candidate.area - best.area) > best.area * 0.05:
            continue
        opposite = candidate
        break

    thickness = None
    if opposite is not None:
        thickness = abs(dot(sub(opposite.origin, best.origin), best.normal))
        ok("толщина материала: {:.2f} мм".format(thickness))
        # Наружу — в сторону, противоположную материалу: контур не зеркалится.
        material = sub(opposite.origin, best.origin)
        if dot(material, best.normal) > 0:
            best.normal = scale(best.normal, -1.0)
    else:
        warn("парная грань не найдена, толщина не определена.")
        log("      Возможен зеркальный контур — проверьте деталь перед резкой.")

    # Габарит детали поперёк наибольшей грани: у плоской заготовки он равен
    # толщине листа, у согнутой — заметно больше.
    extent = 0.0
    for point in all_points:
        extent = max(extent, abs(dot(sub(point, best.origin), best.normal)))
    ok("габарит поперёк грани: {:.2f} мм".format(extent))

    return best, thickness, extent


# ============================================================
# ПОДГОТОВКА КОНТУРОВ В 2D
# ============================================================

def flatten(face):
    """PlanarFace -> список плоских контуров, сдвинутых к началу координат."""
    axis_x, axis_y = plane_axes(face.normal)
    loops = []
    for loop in face.loops:
        flat = [project(point, face.origin, axis_x, axis_y) for point in loop]
        flat = simplify(flat)
        if len(flat) >= 2:
            loops.append(flat)

    if CONFIG["zero_origin"] and loops:
        min_x = min(x for loop in loops for x, _ in loop)
        min_y = min(y for loop in loops for _, y in loop)
        loops = [[(x - min_x, y - min_y) for x, y in loop] for loop in loops]
    return loops


def simplify(points):
    """
    Выбрасывает точки, лежащие на прямой в пределах допуска.
    Дуги и сплайны при этом сохраняют форму, а прямые участки
    не превращаются в частокол вершин.
    """
    if len(points) < 3:
        return points

    tolerance = CONFIG["tolerance"]
    result = [points[0]]
    for index in range(1, len(points) - 1):
        previous = result[-1]
        current = points[index]
        following = points[index + 1]
        if point_line_distance(current, previous, following) > tolerance:
            result.append(current)
    result.append(points[-1])
    return result


def point_line_distance(point, start, finish):
    dx = finish[0] - start[0]
    dy = finish[1] - start[1]
    length = math.hypot(dx, dy)
    if length < 1e-12:
        return math.hypot(point[0] - start[0], point[1] - start[1])
    area = abs(dx * (start[1] - point[1]) - (start[0] - point[0]) * dy)
    return area / length


def is_closed(loop):
    return math.hypot(loop[0][0] - loop[-1][0],
                      loop[0][1] - loop[-1][1]) <= CONFIG["weld_tolerance"]


# ============================================================
# ЗАПИСЬ DXF (R12)
# ============================================================

class DxfWriter(object):
    """
    Минимальный писатель DXF R12: только то, что нужно раскрою.
    Формат R12 понимают все CAM-системы лазерной резки.
    """

    def __init__(self):
        self.entities = []
        self.layers = []
        self.bounds = [None, None, None, None]

    def layer(self, name, color=7):
        if name not in [item[0] for item in self.layers]:
            self.layers.append((name, color))

    def _track(self, x, y):
        min_x, min_y, max_x, max_y = self.bounds
        self.bounds = [
            x if min_x is None else min(min_x, x),
            y if min_y is None else min(min_y, y),
            x if max_x is None else max(max_x, x),
            y if max_y is None else max(max_y, y),
        ]

    def polyline(self, layer, points, closed):
        self.layer(layer)
        body = ["0", "POLYLINE", "8", layer, "66", "1", "70",
                "1" if closed else "0",
                "10", "0.0", "20", "0.0", "30", "0.0"]
        for x, y in points:
            self._track(x, y)
            body += ["0", "VERTEX", "8", layer,
                     "10", fmt(x), "20", fmt(y), "30", "0.0"]
        body += ["0", "SEQEND", "8", layer]
        self.entities += body

    def text(self, layer, x, y, height, value):
        self.layer(layer, color=3)
        self._track(x, y)
        self.entities += ["0", "TEXT", "8", layer,
                          "10", fmt(x), "20", fmt(y), "30", "0.0",
                          "40", fmt(height), "1", value, "7", "STANDARD"]

    def dumps(self):
        min_x, min_y, max_x, max_y = self.bounds
        min_x = 0.0 if min_x is None else min_x
        min_y = 0.0 if min_y is None else min_y
        max_x = 0.0 if max_x is None else max_x
        max_y = 0.0 if max_y is None else max_y

        lines = [
            "0", "SECTION", "2", "HEADER",
            "9", "$ACADVER", "1", "AC1009",
            "9", "$INSUNITS", "70", "4",
            "9", "$MEASUREMENT", "70", "1",
            "9", "$EXTMIN", "10", fmt(min_x), "20", fmt(min_y), "30", "0.0",
            "9", "$EXTMAX", "10", fmt(max_x), "20", fmt(max_y), "30", "0.0",
            "0", "ENDSEC",
            "0", "SECTION", "2", "TABLES",
            "0", "TABLE", "2", "LTYPE", "70", "1",
            "0", "LTYPE", "2", "CONTINUOUS", "70", "0",
            "3", "Solid line", "72", "65", "73", "0", "40", "0.0",
            "0", "ENDTAB",
            "0", "TABLE", "2", "LAYER", "70", str(len(self.layers)),
        ]
        for name, color in self.layers:
            lines += ["0", "LAYER", "2", name, "70", "0",
                      "62", str(color), "6", "CONTINUOUS"]
        lines += [
            "0", "ENDTAB",
            "0", "TABLE", "2", "STYLE", "70", "1",
            "0", "STYLE", "2", "STANDARD", "70", "0", "40", "0.0",
            "41", "1.0", "50", "0.0", "71", "0", "42", "2.5",
            "3", "txt", "4", "",
            "0", "ENDTAB",
            "0", "ENDSEC",
            "0", "SECTION", "2", "ENTITIES",
        ]
        lines += self.entities
        lines += ["0", "ENDSEC", "0", "EOF"]
        return "\r\n".join(lines) + "\r\n"

    def save(self, path):
        data = self.dumps()
        encoding = CONFIG["encoding"]
        with open(path, "wb") as handle:
            handle.write(data.encode(encoding, "replace"))


def fmt(value):
    return "{:.6f}".format(float(value))


def write_dxf(path, loops, article):
    """Пишет контуры и артикул в DXF."""
    step("Запись DXF")

    writer = DxfWriter()
    cut_layer = CONFIG["cut_layer"]
    open_loops = 0

    for loop in loops:
        closed = is_closed(loop)
        if not closed:
            open_loops += 1
        points = loop[:-1] if closed and len(loop) > 2 else loop
        writer.polyline(cut_layer, points, closed)

    if open_loops:
        warn("незамкнутых контуров: {} — проверьте файл перед резкой".format(open_loops))

    if CONFIG["add_article"] and article:
        min_x, min_y = writer.bounds[0], writer.bounds[1]
        writer.text(
            CONFIG["mark_layer"],
            min_x,
            min_y - CONFIG["article_gap"] - CONFIG["article_height"],
            CONFIG["article_height"],
            article,
        )
        ok("артикул '{}' записан в слой {}".format(article, CONFIG["mark_layer"]))

    writer.save(path)
    ok("контуров: {}, файл: {}".format(len(loops), path))


# ============================================================
# ДИАГНОСТИКА API
# ============================================================

def list_interfaces():
    """Интерфейсы API7, доступные в этой сборке КОМПАСа."""
    module = api7_module()
    log("\n--- Интерфейсы API7 (геометрия и листовое тело) ---")
    if module is None:
        log("    модуль API7 недоступен")
        return
    keywords = ("Body", "Face", "Edge", "Vertex", "Loop",
                "Container", "Unfold")
    names = sorted(
        name for name in dir(module)
        if name.startswith("I")
        and "_vtables_" not in name
        and any(key in name for key in keywords)
    )
    for name in names:
        log("    {}".format(name))
    if not names:
        log("    подходящих имён не найдено")


def probe(application):
    """Печатает реальный набор членов COM-объектов текущей сборки КОМПАСа."""
    log("\n" + "=" * 70)
    log("ДИАГНОСТИКА API")
    log("=" * 70)

    list_interfaces()

    document = value_of(application, ["ActiveDocument"])
    describe(document, "Активный документ (как вернул ActiveDocument)")

    document_3d = qi(document, "IKompasDocument3D", "IKompasDocument3D1")
    describe(document_3d, "Документ после приведения к 3D")

    part = value_of(document_3d, ["TopPart"]) or value_of(document, ["TopPart"])
    part = qi(part, "IPart7")
    if part is None:
        log("\nДеталь получить не удалось — дальше идти не с чем.")
        return

    describe(part, "Деталь (IPart7)")

    model_container = qi(part, "IModelContainer", strict=True)
    describe(model_container, "IModelContainer (тела и объекты модели)")

    sheet = sheet_container(part)
    describe(sheet, "ISheetMetalContainer (листовые операции)")
    bodies_sm = sheet_bodies(part)
    log("\nЛистовых тел: {}".format(len(bodies_sm)))
    if bodies_sm:
        describe(bodies_sm[0], "ISheetMetalBody")
        log("\nТолщина листа: {}".format(sheet_thickness(part)))

    _, parameters = fetch(sheet, UNFOLD_PARAMETER_NAMES)
    describe(parameters, "Параметры развёртки (объект контейнера)")
    for index, item in enumerate(as_list(parameters)):
        describe(item, "Параметры развёртки [{}]".format(index))
        if index >= 1:
            break

    holder, flag = find_unfold_flag(part)
    log("\nПризнак развёртки: {}".format(flag or "не найден"))

    found = operations(part)
    log("\nОпераций в модели: {}".format(len(found)))
    for index, operation in enumerate(found):
        describe(operation, "Операция [{}]".format(index))
        name, result = fetch(operation, BODY_RESULT_NAMES)
        log("\nРезультат операции ({}): {}".format(name, type(result).__name__))
        if index >= 1:
            break

    bodies = get_bodies(part)
    if not bodies:
        log("\nТела детали не получены — проверьте имена из списка BODY_NAMES.")
        return
    describe(bodies[0], "Тело (IBody7)")

    faces = get_faces(bodies[0])
    log("\nГраней в первом теле: {}".format(len(faces)))
    if not faces:
        return
    describe(faces[0], "Грань (IFace)")

    loops = get_loops(faces[0])
    log("\nЦиклов в первой грани: {}".format(len(loops)))
    if loops:
        describe(loops[0], "Цикл (ILoop7)")

    owner = loops[0] if loops else faces[0]
    _, raw_edges = fetch(owner, EDGE_NAMES)
    raw_items = as_list(raw_edges)
    log("\nРёбер получено: {}".format(len(raw_items)))
    if not raw_items:
        return
    describe(raw_items[0], "Ребро как отдало API")

    edge = unwrap_edge(raw_items[0])
    describe(edge, "Ребро после приведения (IEdge)")

    describe(curve_of(edge), "Кривая ребра")

    log("\nТочки первого ребра: {}".format(edge_polyline(edge, samples=4)))


# ============================================================
# ГЛАВНАЯ ПРОЦЕДУРА
# ============================================================

def sheet_thickness(part):
    """Толщина листа из параметров листового тела."""
    for body in sheet_bodies(part):
        value = value_of(body, ["Thickness"])
        if isinstance(value, (int, float)) and value > 0:
            return float(value)
    return None


def check_flat(part, thickness, extent):
    """
    Проверяет, что деталь действительно плоская заготовка.

    У развёрнутого листа и у пластины габарит поперёк наибольшей грани равен
    толщине материала. Если он больше — деталь согнута, и контур наибольшей
    грани развёрткой не является: молча отдавать такой DXF нельзя.
    """
    reference = sheet_thickness(part) or thickness
    if reference is None:
        warn("толщину определить не удалось — проверка на плоскостность пропущена.")
        return True

    limit = reference * 1.5 + 0.5
    if extent <= limit:
        return True

    err("деталь не плоская: габарит поперёк грани {:.2f} мм при толщине {:.2f} мм."
        .format(extent, reference))
    if is_sheet_metal(part):
        log("    Разверните листовое тело (Листовое тело -> Развернуть)")
        log("    и запустите макрос снова.")
    else:
        log("    Развёртка возможна только для листового тела. Преобразуйте")
        log("    деталь командой 'Распознать листовое тело' и повторите.")
    return False


def export_active_part(application):
    document, part, path = get_active_part(application)
    if part is None:
        return False

    unfolded, restore = ensure_unfolded(part)
    try:
        if not unfolded:
            return False

        bodies = get_bodies(part)
        if not bodies:
            err("не удалось получить тела детали.")
            return False

        face, thickness, extent = find_plate_faces(bodies)
        if face is None:
            return False

        if not check_flat(part, thickness, extent):
            return False

        loops = flatten(face)
        if not loops:
            err("контуры грани не получены.")
            return False

        output = os.path.join(
            os.path.dirname(path),
            os.path.splitext(os.path.basename(path))[0] + ".dxf",
        )
        if os.path.exists(output) and not CONFIG["overwrite"]:
            err("файл уже существует: {}".format(output))
            return False

        write_dxf(output, loops, get_article(part, path))
        return True
    finally:
        if restore is not None:
            restore()


def command_line_mode():
    """Режим из аргументов запуска; во встроенном Python аргументов нет."""
    try:
        return sys.argv[1].lower() if len(sys.argv) > 1 else ""
    except Exception:
        return ""


def probe_api():
    """Диагностика API. Вызывается из консоли PyScripter: probe_api()"""
    application, _ = connect_kompas()
    if application is not None:
        probe(application)


def run_macro(mode=None):
    log("=" * 70)
    log("КОМПАС-3D -> DXF: развёртка для лазерного раскроя")
    log("=" * 70)

    application, _ = connect_kompas()
    if application is None:
        return False

    if mode is None:
        mode = CONFIG["mode"] or command_line_mode()
    if mode in ("probe", "diag", "-p"):
        probe(application)
        return True

    success = False
    try:
        success = export_active_part(application)
    except Exception as exc:
        err("сбой: {}".format(exc))
        traceback.print_exc()

    if not success and CONFIG["probe_on_error"]:
        log("\nЭкспорт не выполнен. Ниже — состав API текущей сборки КОМПАСа;")
        log("пришлите этот вывод, чтобы уточнить имена интерфейсов.")
        try:
            probe(application)
        except Exception:
            traceback.print_exc()

    log("\n" + "=" * 70)
    log("ГОТОВО" if success else "ЭКСПОРТ НЕ ВЫПОЛНЕН")
    log("=" * 70)
    return success


if __name__ == "__main__":
    run_macro()
