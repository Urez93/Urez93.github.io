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
  - pywin32 (win32com) — связь с Компасом;
  - openpyxl (и xlrd, если промежуточный файл получится в старом .xls);
  - kompas_bom_structure.py должен лежать в той же папке, что и этот файл.

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
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from kompas_bom_structure import build_structured_workbook


DOCUMENT_TYPE_SPECIFICATION = 5  # ksDocumentSpecification


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
