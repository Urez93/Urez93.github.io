# -*- coding: utf-8 -*-
"""
Разведочный скрипт v4: НИЧЕГО не сохраняет и не изменяет.

Задача: найти, откуда панель "Свойства" в Компасе (со "Списком свойств":
Обозначение/Наименование/Количество/Материал/Масса/Позиция/Раздел
спецификации/Форматы листов/Примечание) берёт значение "Раздел
спецификации" для детали — и получить это же значение из Python.

6 прямых свойств IPart7 (SpecSectionType/SectionType/...) не сработали.
Скорее всего эти данные лежат в отдельном связанном объекте — что-то
вроде "Article" (изделие/статья спецификации) — а не прямо на IPart7.
Пробуем несколько вероятных названий такого объекта и, если найдём,
смотрим его реальные члены через dir() и пробуем достать значение по
нескольким вероятным именам параметра.

Запустите на детали, у которой вы точно видели "Раздел спецификации" =
"Стандартные изделия" в панели свойств (например, тот самый болт с
скриншота), и пришлите весь вывод.
"""

import win32com.client


def _print_members(title, obj):
    print()
    print("-" * 70)
    print(title)
    print("-" * 70)
    if obj is None:
        print("  (объект отсутствует)")
        return
    print("  Python-тип объекта:", type(obj))
    try:
        members = [m for m in dir(obj) if not m.startswith("_")]
    except Exception as e:
        print("  dir() не сработал:", e)
        return
    print("  Члены ({}):".format(len(members)))
    for m in sorted(members):
        print("    ", m)


def _try_get(obj, name, *args):
    try:
        attr = getattr(obj, name)
        result = attr(*args) if callable(attr) else attr
        print("    ✓ {}{} -> {!r}".format(name, args, result))
        return result
    except Exception as e:
        print("    ✗ {}{} недоступен: {}".format(name, args, e))
        return None


def main():
    print("=" * 70)
    print("РАЗВЕДКА API КОМПАСА v4 — источник 'Раздел спецификации'")
    print("=" * 70)

    app = win32com.client.Dispatch("Kompas.Application.7")
    doc = app.ActiveDocument
    print("Активный документ:", doc.PathName)

    doc3d = win32com.client.CastTo(doc, "IKompasDocument3D")
    top_part = doc3d.TopPart

    # Найдём деталь, у которой заведомо есть "Раздел спецификации" —
    # пользователь запускает скрипт на выделенной детали или на
    # top_part напрямую (если top_part сам такая деталь). Дальше
    # достаточно первого прямого потомка, если у top_part'а самого
    # свойство пустое — просто для проверки механизма доступа.
    target = top_part
    print()
    print("Целевая деталь (TopPart):", getattr(target, "Name", "?"), "/", getattr(target, "Marking", "?"))

    print()
    print("-" * 70)
    print("Пробуем найти связанный объект-контейнер свойств спецификации")
    print("-" * 70)
    container = None
    for attr_name in ("Article", "SpecificationInfo", "PropertyMng", "Properties",
                       "AttributeMng", "SpecArticle", "SpcArticle", "BomInfo"):
        try:
            value = getattr(target, attr_name)
            result = value() if callable(value) else value
            if result is not None:
                print("  ✓ target.{} -> {}".format(attr_name, type(result)))
                container = (attr_name, result)
                break
            else:
                print("  ✗ target.{} вернул None".format(attr_name))
        except Exception as e:
            print("  ✗ target.{} недоступен: {}".format(attr_name, e))

    if container:
        _print_members("Объект '{}'".format(container[0]), container[1])

        obj = container[1]
        print()
        print("-" * 70)
        print("Пробуем получить значение параметра 'Раздел спецификации' из найденного объекта")
        print("-" * 70)
        for method_name in ("GetParamValue", "GetValue", "GetParameterValue", "Value", "get_Value"):
            _try_get(obj, method_name, "Раздел спецификации")
            _try_get(obj, method_name, "SpecSection")
            _try_get(obj, method_name, "SectionName")
    else:
        print()
        print("  Ни один из вероятных контейнеров не найден на самой детали.")

    # На случай, если это не деталь, а СПЕЦИФИКАЦИЯ-ФРАГМЕНТ на уровне
    # документа: проверим и app.ActiveDocument напрямую теми же именами.
    print()
    print("-" * 70)
    print("То же самое, но на активном документе (doc), а не на детали")
    print("-" * 70)
    for attr_name in ("Article", "SpecificationInfo", "PropertyMng", "Properties", "AttributeMng"):
        try:
            value = getattr(doc, attr_name)
            result = value() if callable(value) else value
            print("  {} doc.{} -> {}".format("✓" if result is not None else "✗(None)", attr_name, type(result) if result is not None else ""))
        except Exception as e:
            print("  ✗ doc.{} недоступен: {}".format(attr_name, e))

    print()
    print("=" * 70)
    print("ГОТОВО. Пришлите весь вывод выше.")
    print("=" * 70)


if __name__ == "__main__":
    main()
