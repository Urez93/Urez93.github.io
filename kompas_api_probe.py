# -*- coding: utf-8 -*-
"""
Разведочный скрипт: НИЧЕГО не сохраняет и не изменяет — только
подключается к Компасу и печатает реальные имена свойств/методов
активного документа и его детали/сборки, чтобы найти правильный вызов
для меню "Управление -> Отчёты -> Создать отчёт" (и соседних пунктов
"Создать таблицу исполнений", "Создать отчёт по массиву").

Почему так: два предыдущих макроса дважды промахнулись с угаданными
именами методов (Export/SaveAsToFormat/числовые коды формата), а один
раз это даже привело к тому, что реальный открытый документ оказался
переключён на временный файл. Чтобы больше не гадать вслепую, этот
скрипт использует dir() по объектам, полученным через win32com —
судя по прошлому выводу консоли ("win32com.gen_py...IKompasDocument
instance"), pywin32 сгенерировал полноценную обёртку с реальными
именами (в отличие от "голого" IDispatch, где dir() ничего не покажет).

Запустите этот скрипт в Компасе на активной модели/сборке (для которой
нужно получить состав изделия) и пришлите весь вывод консоли — по нему
будет видно точное имя метода/интерфейса для генерации отчёта.
"""

import win32com.client


def _print_members(title, obj, keywords=None):
    print()
    print("-" * 70)
    print(title)
    print("-" * 70)
    if obj is None:
        print("  (объект отсутствует)")
        return

    print("  Python-тип объекта: {}".format(type(obj)))

    try:
        members = [m for m in dir(obj) if not m.startswith("_")]
    except Exception as e:
        print("  dir() не сработал: {}".format(e))
        return

    if keywords:
        filtered = [m for m in members if any(k.lower() in m.lower() for k in keywords)]
        print("  Всего членов: {}. Из них похожих на {}: {}".format(len(members), keywords, len(filtered)))
        for m in sorted(filtered):
            print("    ", m)
        print("  (полный список ниже)")

    print("  Полный список членов:")
    for m in sorted(members):
        print("    ", m)


def main():
    print("=" * 70)
    print("РАЗВЕДКА API КОМПАСА — ничего не сохраняет и не меняет")
    print("=" * 70)

    app = win32com.client.Dispatch("Kompas.Application.7")
    print("Приложение получено:", type(app))

    doc = app.ActiveDocument
    _print_members("АКТИВНЫЙ ДОКУМЕНТ (doc)", doc, keywords=["report", "спец", "spec", "struct", "compos", "bom"])

    try:
        print()
        print("doc.PathName =", doc.PathName)
    except Exception as e:
        print("doc.PathName недоступен:", e)

    try:
        doc7 = win32com.client.CastTo(doc, "IKompasDocument3D")
        print()
        print("Успешно приведён к IKompasDocument3D:", type(doc7))
    except Exception as e:
        print()
        print("CastTo(IKompasDocument3D) не сработал:", e)
        doc7 = None

    top_part = None
    for attr in ("TopPart",):
        try:
            top_part = getattr(doc7 if doc7 else doc, attr)
            print("{} получен: {}".format(attr, type(top_part)))
            break
        except Exception as e:
            print("{} недоступен: {}".format(attr, e))

    _print_members("ВЕРХНЯЯ ДЕТАЛЬ/СБОРКА (TopPart)", top_part, keywords=["report", "спец", "spec", "struct", "compos", "bom"])

    if top_part is not None:
        for candidate_attr in ("ReportManager", "Report", "Reports"):
            try:
                sub_obj = getattr(top_part, candidate_attr)
                _print_members("top_part.{}".format(candidate_attr), sub_obj)
            except Exception as e:
                print()
                print("top_part.{} недоступен: {}".format(candidate_attr, e))

    for candidate_attr in ("ReportManager", "Report", "Reports"):
        try:
            sub_obj = getattr(doc, candidate_attr)
            _print_members("doc.{}".format(candidate_attr), sub_obj)
        except Exception as e:
            print()
            print("doc.{} недоступен: {}".format(candidate_attr, e))

    print()
    print("=" * 70)
    print("ГОТОВО. Пришлите весь вывод выше.")
    print("=" * 70)


if __name__ == "__main__":
    main()
