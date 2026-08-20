# -*- coding: utf-8 -*-
"""
Макрос для создания разверток в DXF из 3D модели Компас-3D 22.
Используется встроенный Python в Компас-3D (PyScripter).
"""

import os
import sys


def get_kompas_app():
    """Получение объекта приложения Компас-3D"""
    print("Пытаемся получить приложение Компас...")
    
    # Попытка 1: через встроенный объект Компаса
    try:
        print("  Попытка 1: ищем Kompas в globals...")
        kompas = globals().get('Kompas')
        if kompas:
            print("  ✓ Найден Kompas через globals()")
            return kompas
    except Exception as e:
        print("  ✗ Попытка 1 не удалась: {}".format(str(e)))
    
    # Попытка 2: через sys.modules
    try:
        print("  Попытка 2: ищем в sys.modules...")
        if 'Kompas' in sys.modules:
            kompas = sys.modules['Kompas']
            print("  ✓ Найден Kompas в sys.modules")
            return kompas
    except Exception as e:
        print("  ✗ Попытка 2 не удалась: {}".format(str(e)))
    
    # Попытка 3: через __main__
    try:
        print("  Попытка 3: ищем в __main__...")
        main = sys.modules.get('__main__')
        if hasattr(main, 'Kompas'):
            kompas = getattr(main, 'Kompas')
            print("  ✓ Найден Kompas в __main__")
            return kompas
    except Exception as e:
        print("  ✗ Попытка 3 не удалась: {}".format(str(e)))
    
    # Попытка 4: через встроенный объект приложения
    try:
        print("  Попытка 4: ищем Application...")
        app = globals().get('Application')
        if app:
            print("  ✓ Найден Application через globals()")
            return app
    except Exception as e:
        print("  ✗ Попытка 4 не удалась: {}".format(str(e)))
    
    # Попытка 5: прямой импорт
    try:
        print("  Попытка 5: прямой импорт...")
        from Kompas import Application
        print("  ✓ Успешно импортирован Application")
        return Application
    except Exception as e:
        print("  ✗ Попытка 5 не удалась: {}".format(str(e)))
    
    print("  ✗ Не удалось получить приложение.")
    return None


def get_active_document(app):
    """Получение активного документа"""
    try:
        print("Получаем активный документ...")
        
        # Попытка 1: прямой доступ
        try:
            doc = app.ActiveDocument
            if doc:
                print("✓ Активный документ получен (способ 1)")
                return doc
        except Exception as e:
            print("  ✗ Способ 1 не удался: {}".format(str(e)))
        
        # Попытка 2: через Documents
        try:
            docs = app.Documents
            if docs and docs.Count > 0:
                doc = docs.Item(0)
                print("✓ Активный документ получен (способ 2)")
                return doc
        except Exception as e:
            print("  ✗ Способ 2 не удался: {}".format(str(e)))
        
    except Exception as e:
        print("Ошибка получения документа: {}".format(str(e)))
    
    return None


def get_top_part(doc):
    """Получение верхней детали"""
    try:
        print("Получаем верхнюю деталь...")
        
        # Попытка 1: прямой доступ
        try:
            part = doc.TopPart
            if part:
                print("✓ TopPart получена (способ 1)")
                print("  Имя: {}".format(part.Name))
                return part
        except Exception as e:
            print("  ✗ Способ 1 не удался: {}".format(str(e)))
        
        # Попытка 2: через Parts
        try:
            parts = doc.Parts
            if parts and parts.Count > 0:
                part = parts.Item(0)
                print("✓ TopPart получена (способ 2)")
                print("  Имя: {}".format(part.Name))
                return part
        except Exception as e:
            print("  ✗ Способ 2 не удался: {}".format(str(e)))
        
    except Exception as e:
        print("Ошибка получения верхней детали: {}".format(str(e)))
    
    return None


def is_sheet_metal_part(part):
    """Проверка является ли деталь листовой"""
    print("Проверяем тип детали...")
    
    try:
        # Попытка 1: SheetMetalContainer
        try:
            container = part.SheetMetalContainer
            if container:
                print("✓ Деталь листовая (SheetMetalContainer)")
                return True, container
        except Exception as e:
            print("  SheetMetalContainer недоступен: {}".format(str(e)))
        
        # Попытка 2: проверка типа через свойства
        try:
            # Если есть свойство IsSheetMetal
            is_sheet = part.IsSheetMetal
            if is_sheet:
                print("✓ Деталь листовая (IsSheetMetal = True)")
                return True, part
        except Exception as e:
            print("  IsSheetMetal недоступен: {}".format(str(e)))
        
        # Если ничего не сработало - твердотельная
        print("✗ Деталь твердотельная")
        return False, None
        
    except Exception as e:
        print("Ошибка проверки типа: {}".format(str(e)))
        return False, None


def get_largest_face(part):
    """Получение самой большой грани твердотельной детали"""
    print("Определяем самую большую грань...")
    
    try:
        bodies = part.Bodies
        if not bodies or bodies.Count == 0:
            print("✗ В детали нет тел.")
            return None
        
        print("  Тел найдено: {}".format(bodies.Count))
        
        body = bodies.Item(0)
        faces = body.Faces
        
        if not faces or faces.Count == 0:
            print("✗ В теле нет граней.")
            return None
        
        print("  Граней найдено: {}".format(faces.Count))
        
        max_area = 0
        largest_face = None
        largest_face_index = 0
        
        # Ищем грань с максимальной площадью
        for i in range(faces.Count):
            try:
                face = faces.Item(i)
                area = face.Area
                print("    Грань {}: площадь {}".format(i, area))
                
                if area > max_area:
                    max_area = area
                    largest_face = face
                    largest_face_index = i
            except Exception as e:
                print("    Грань {} — ошибка: {}".format(i, str(e)))
                continue
        
        if largest_face is None:
            largest_face = faces.Item(0)
            largest_face_index = 0
        
        print("✓ Самая большая грань: индекс {}, площадь {}".format(
            largest_face_index, max_area))
        return largest_face
        
    except Exception as e:
        print("Ошибка при получении грани: {}".format(str(e)))
        import traceback
        traceback.print_exc()
        return None


def export_to_dxf(doc, output_path):
    """Экспорт документа в DXF"""
    print("Экспортируем в DXF: {}".format(output_path))
    
    try:
        # Убедимся, что папка существует
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # Попытка 1: SaveAs с форматом DXF
        try:
            print("  Попытка 1: SaveAs с форматом DXF...")
            doc.SaveAs(output_path, 31)  # 31 = DXF
            print("✓ Файл сохранен через SaveAs")
            return True
        except Exception as e:
            print("  ✗ SaveAs не удалась: {}".format(str(e)))
        
        # Попытка 2: Export
        try:
            print("  Попытка 2: Export...")
            doc.Export(output_path, 31)
            print("✓ Файл экспортирован через Export")
            return True
        except Exception as e:
            print("  ✗ Export не удалась: {}".format(str(e)))
        
        # Попытка 3: через StdCmd
        try:
            print("  Попытка 3: StdCmd.ExportToDXF...")
            # Попробуем через команду Компаса
            app = globals().get('Kompas')
            if app and hasattr(app, 'StdCmd'):
                app.StdCmd.ExportToDXF(output_path)
                print("✓ Файл экспортирован через StdCmd")
                return True
        except Exception as e:
            print("  ✗ StdCmd не удалась: {}".format(str(e)))
        
        print("✗ Экспорт не удалась всеми методами.")
        return False
        
    except Exception as e:
        print("Ошибка при экспорте: {}".format(str(e)))
        import traceback
        traceback.print_exc()
        return False


def run_macro():
    """Основной макрос"""
    print()
    print("=" * 70)
    print("КОМПАС-3D — Создание разверток в DXF")
    print("=" * 70)
    print()

    try:
        # Получаем приложение
        app = get_kompas_app()
        if app is None:
            raise RuntimeError("Не удалось подключиться к Компас-3D.")
        print()

        # Получаем активный документ
        doc = get_active_document(app)
        if doc is None:
            raise RuntimeError("Активный документ не найден.")
        print()

        # Получаем верхнюю деталь
        part = get_top_part(doc)
        if part is None:
            raise RuntimeError("Верхняя деталь не найдена.")
        print()

        # Определяем тип детали
        is_sheet, sheet_container = is_sheet_metal_part(part)
        print()

        # Обработка детали
        if is_sheet:
            print("Обработка листовой детали...")
        else:
            print("Обработка твердотельной детали...")
            largest_face = get_largest_face(part)
            if largest_face is None:
                raise RuntimeError("Не удалось получить самую большую грань.")
        
        print()

        # Экспортируем в DXF
        output_dxf = os.path.expanduser("~\\Desktop\\unfold_output.dxf")
        
        success = export_to_dxf(doc, output_dxf)
        
        print()
        print("=" * 70)
        if success:
            print("✓ УСПЕШНО ЗАВЕРШЕНО")
            print("Файл сохранен: {}".format(output_dxf))
        else:
            print("✗ ЭКСПОРТ НЕ УДАЛАСЬ")
            print("Попробуйте экспортировать вручную через меню Компаса.")
        print("=" * 70)
        print()

        return success

    except Exception as e:
        print()
        print("=" * 70)
        print("✗ ОШИБКА")
        print("=" * 70)
        print(str(e))
        print()
        import traceback
        traceback.print_exc()
        return False


# ============================================================
# ЗАПУСК МАКРОСА
# ============================================================

if __name__ == "__main__":
    run_macro()
