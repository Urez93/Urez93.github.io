# -*- coding: utf-8 -*-
"""
Разведочный скрипт v2: НИЧЕГО не сохраняет и не изменяет.

Первая разведка показала, что в объектах IKompasDocument/IPart7 (API7)
нет метода, похожего на команду меню "Управление -> Отчёты -> Создать
отчёт" — похоже, это чисто UI-команда без прямого COM-метода. Но она же
показала, что dir() ненадёжен: например, doc.PathName прекрасно
работает, хотя в dir(doc) его не было. Значит для реальных свойств
нужно не смотреть dir(), а пробовать их вызвать напрямую (getattr) —
COM в pywin32 обычно позволяет обращаться и к тому, что не попало в
статически сгенерированный список.

Зато у IPart7 есть настоящий метод PartsEx — это прямой доступ к
дочерним деталям/подсборкам верхней детали. План: вместо экспорта
спецификации/отчёта строить структуру состава изделия напрямую обходом
дерева сборки через PartsEx — так вообще не нужен ни файл
спецификации, ни команда "Создать отчёт".

Этот скрипт проверяет:
  - как из PartsEx получить список дочерних деталей (Count/Item и т.п.);
  - какие у детали реально работают свойства: Name, обозначение
    (Marking/Designation/...), количество экземпляров, материал;
  - работает ли то же самое рекурсивно на дочерней детали (есть ли у
    неё тоже PartsEx, то есть можно ли идти вглубь).

Запустите на той же сборке верхнего уровня, что и раньше, и пришлите
весь вывод.
"""

import win32com.client


def _try_attrs(obj, names, label):
    print("  Проверяем атрибуты объекта [{}]:".format(label))
    found = {}
    for name in names:
        try:
            value = getattr(obj, name)
            # Если это метод/COM-объект, а не простое значение — не печатаем целиком.
            printable = value
            if callable(value):
                printable = "<callable>"
            print("    ✓ {} = {!r}".format(name, printable))
            found[name] = value
        except Exception as e:
            print("    ✗ {} недоступен: {}".format(name, e))
    return found


def _try_collection_access(obj, label):
    print("  Пробуем получить количество элементов и первый элемент [{}]:".format(label))

    if isinstance(obj, (list, tuple)):
        print("    Результат — обычный Python {} длиной {}".format(type(obj).__name__, len(obj)))
        first = obj[0] if obj else None
        if first is not None:
            print("    Первый элемент, тип: {}".format(type(first)))
        return len(obj), first

    count = None
    for name in ("Count", "GetCount", "Length"):
        try:
            value = getattr(obj, name)
            count = value() if callable(value) else value
            print("    ✓ {} -> {}".format(name, count))
            break
        except Exception as e:
            print("    ✗ {} недоступен: {}".format(name, e))

    first_item = None
    if count:
        for name in ("Item", "GetItem", "GetPart"):
            try:
                method = getattr(obj, name)
                first_item = method(0)
                print("    ✓ {}(0) -> {}".format(name, type(first_item)))
                break
            except Exception as e:
                print("    ✗ {}(0) недоступен: {}".format(name, e))

    try:
        print("    Пробуем for-итерацию по объекту напрямую...")
        items = list(obj)
        print("    ✓ Итерация сработала, элементов: {}".format(len(items)))
        if items and first_item is None:
            first_item = items[0]
    except Exception as e:
        print("    ✗ Итерация не сработала: {}".format(e))

    return count, first_item


CANDIDATE_NAMES = [
    "Name", "Marking", "Designation", "Symbol", "FullMarking",
    "Material", "MaterialName",
    "Count", "InstanceCount", "UniqueNum",
    "Comment", "Note",
]


def main():
    print("=" * 70)
    print("РАЗВЕДКА API КОМПАСА v2 — структура сборки через PartsEx")
    print("=" * 70)

    app = win32com.client.Dispatch("Kompas.Application.7")
    doc = app.ActiveDocument
    print("Активный документ:", doc.PathName)

    doc3d = win32com.client.CastTo(doc, "IKompasDocument3D")
    top_part = doc3d.TopPart
    print()
    print("TopPart получен:", type(top_part))

    print()
    _try_attrs(top_part, CANDIDATE_NAMES, "TopPart")

    print()
    print("-" * 70)
    print("top_part.PartsEx — это МЕТОД (нужны аргументы), пробуем разные сигнатуры")
    print("-" * 70)

    call_attempts = [
        ("PartsEx()", lambda: top_part.PartsEx()),
        ("PartsEx(True)", lambda: top_part.PartsEx(True)),
        ("PartsEx(False)", lambda: top_part.PartsEx(False)),
        ("PartsEx(True, True)", lambda: top_part.PartsEx(True, True)),
        ("PartsEx(False, False)", lambda: top_part.PartsEx(False, False)),
        ("PartsEx(True, False)", lambda: top_part.PartsEx(True, False)),
        ("PartsEx(False, True)", lambda: top_part.PartsEx(False, True)),
        ("PartsEx(0)", lambda: top_part.PartsEx(0)),
        ("PartsEx(1)", lambda: top_part.PartsEx(1)),
    ]

    parts_ex = None
    for name, action in call_attempts:
        try:
            result = action()
            print("  ✓ {} сработал, тип результата: {}, значение (если короткое): {!r}".format(
                name, type(result), result if not hasattr(result, "__len__") or len(str(result)) < 200 else "<длинное>"
            ))
            if result:
                parts_ex = result
                print("    -> будем использовать этот результат для дальнейшего разбора")
                break
        except Exception as e:
            print("  ✗ {} не сработал: {}".format(name, e))

    try:
        if parts_ex is not None:
            print()
            print("Тип parts_ex:", type(parts_ex))
        count, first_child = _try_collection_access(parts_ex, "PartsEx(...)") if parts_ex is not None else (None, None)

        if first_child is not None:
            print()
            print("-" * 70)
            print("ПЕРВАЯ ДОЧЕРНЯЯ ДЕТАЛЬ/ПОДСБОРКА")
            print("-" * 70)
            print("Тип:", type(first_child))
            _try_attrs(first_child, CANDIDATE_NAMES, "первый child")

            print()
            print("  Проверяем, есть ли у дочерней детали свой PartsEx (можно ли идти глубже)...")
            child_parts_ex = None
            for name, action in (
                ("child.PartsEx()", lambda: first_child.PartsEx()),
                ("child.PartsEx(True)", lambda: first_child.PartsEx(True)),
                ("child.PartsEx(True, True)", lambda: first_child.PartsEx(True, True)),
            ):
                try:
                    child_parts_ex = action()
                    print("    ✓ {} сработал, тип: {}".format(name, type(child_parts_ex)))
                    break
                except Exception as e:
                    print("    ✗ {} не сработал: {}".format(name, e))
            if child_parts_ex is not None:
                _try_collection_access(child_parts_ex, "child.PartsEx(...)")
        else:
            print("Не удалось получить первый дочерний элемент — сборка либо пуста, "
                  "либо нужен другой способ доступа (пришлите этот вывод).")

    except Exception as e:
        print("top_part.PartsEx недоступен:", e)

    print()
    print("=" * 70)
    print("ГОТОВО. Пришлите весь вывод выше.")
    print("=" * 70)


if __name__ == "__main__":
    main()
