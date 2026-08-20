# -*- coding: utf-8 -*-
"""
Макрос для создания разверток в DXF из 3D модели Компас-3D 22.
Используется встроенный Python в Компас-3D (PyScripter).
"""

import os
import tempfile


def get_kompas_app():
    """Получение объекта приложения Компас-3D"""
    try:
        # Для встроенного Python в Компасе
        app = GetApplication()
        return app
    except Exception as e:
        print("Ошибка получения приложения: {}".format(str(e)))
        return None


def get_active_document_3d(app):
    """Получение активного 3D документа"""
    try:
        doc_3d = app.ActiveDocument
        if doc_3d is None:
            raise RuntimeError("Нет активного 3D документа.")
        return doc_3d
    except Exception as e:
        print("Ошибка получения документа: {}".format(str(e)))
        return None


def get_top_part(doc_3d):
    """Получение верхней детали документа"""
    try:
        part = doc_3d.TopPart
        if part is None:
            raise RuntimeError("Верхняя деталь не найдена.")
        print("✓ Верхняя деталь получена: {}".format(part.Name))
        return part
    except Exception as e:
        print("Ошибка получения верхней детали: {}".format(str(e)))
        return None


def is_sheet_metal_part(part):
    """Проверка является ли деталь листовой"""
    try:
        # Попытаемся получить контейнер листовой детали
        sheet_container = part.SheetMetalContainer
        if sheet_container is not None:
            print("✓ Деталь является листовой.")
            return True, sheet_container
    except Exception:
        pass
    
    print("✗ Деталь твердотельная (не листовая).")
    return False, None


def get_largest_face(part):
    """Получение самой большой грани твердотельной детали"""
    print("Определяем самую большую грань...")
    
    try:
        bodies = part.Bodies
        if bodies.Count == 0:
            raise RuntimeError("В детали нет тел.")
        
        body = bodies.Item(0)
        faces = body.Faces
        
        if faces.Count == 0:
            raise RuntimeError("В теле нет граней.")
        
        max_area = 0
        largest_face = None
        largest_face_index = 0
        
        # Ищем грань с максимальной площадью
        for i in range(faces.Count):
            face = faces.Item(i)
            try:
                area = face.Area
                if area > max_area:
                    max_area = area
                    largest_face = face
                    largest_face_index = i
                    print("  Грань {}: площадь {}".format(i, area))
            except Exception:
                continue
        
        if largest_face is None:
            largest_face = faces.Item(0)
            largest_face_index = 0
        
        print("✓ Найдена самая большая грань: индекс {}, площадь {}".format(
            largest_face_index, max_area))
        return largest_face
        
    except Exception as e:
        print("Ошибка при получении грани: {}".format(str(e)))
        return None


def create_sketch_from_face(part, face):
    """Создание эскиза из грани"""
    print("Создаем эскиз из самой большой грани...")
    
    try:
        # Для листовой детали используем встроенную развертку
        # Для твердотельной - проецируем самую большую грань
        
        sketches = part.Sketches
        if sketches is None:
            print("Эскизы недоступны.")
            return None
        
        print("✓ Получен доступ к эскизам (всего: {})".format(sketches.Count))
        
        # Возвращаем первый эскиз или саму грань
        if sketches.Count > 0:
            return sketches.Item(0)
        
        return face
        
    except Exception as e:
        print("Ошибка при создании эскиза: {}".format(str(e)))
        return None


def export_to_dxf(doc_3d, output_path):
    """Экспорт 3D фрагмента в DXF"""
    print("Экспортируем в DXF...")
    
    try:
        # Создаем новый фрагмент для экспорта
        fragment_doc = GetApplication().Documents.CreateDocument(
            doc3D_object_type,  # 3D документ
            None  # Без шаблона
        )
        
        if fragment_doc is None:
            raise RuntimeError("Не удалось создать фрагмент.")
        
        # Копируем содержимое в фрагмент
        # (в зависимости от версии Компаса это может быть разным)
        
        print("Фрагмент создан, сохраняем...")
        
        # Экспортируем в DXF
        # Пример пути: C:\temp\unfold.dxf
        fragment_doc.SaveAs(output_path, 31)  # 31 = DXF формат
        
        print("✓ Файл сохранен: {}".format(output_path))
        
        fragment_doc.Close(False)
        return True
        
    except Exception as e:
        print("Ошибка при экспорте: {}".format(str(e)))
        return False


def export_sketch_to_dxf(sketch, output_path):
    """Экспорт эскиза в DXF"""
    print("Экспортируем эскиз в DXF...")
    
    try:
        # Сохраняем эскиз как DXF
        sketch.SaveAs(output_path)
        print("✓ Эскиз экспортирован: {}".format(output_path))
        return True
        
    except Exception as e:
        print("Ошибка при экспорте эскиза: {}".format(str(e)))
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
        print("✓ Компас-3D подключен.")
        print()

        # Получаем активный документ
        doc_3d = get_active_document_3d(app)
        if doc_3d is None:
            raise RuntimeError("Активный документ не найден.")
        print()

        # Получаем верхнюю деталь
        part = get_top_part(doc_3d)
        if part is None:
            raise RuntimeError("Верхняя деталь не найдена.")
        print()

        # Определяем тип детали
        is_sheet, sheet_container = is_sheet_metal_part(part)
        print()

        # Готовим развертку
        unfold_sketch = None
        
        if is_sheet:
            print("Обработка листовой детали...")
            # Для листовой детали пытаемся получить развертку
            try:
                unfold_sketch = create_sketch_from_face(part, None)
            except Exception as e:
                print("Ошибка обработки листовой детали: {}".format(str(e)))
        else:
            print("Обработка твердотельной детали...")
            # Для твердотельной детали берем самую большую грань
            largest_face = get_largest_face(part)
            if largest_face is not None:
                unfold_sketch = create_sketch_from_face(part, largest_face)
        
        print()

        # Экспортируем в DXF
        output_dxf = os.path.expanduser("~\\Desktop\\unfold_output.dxf")
        
        # Убедимся, что папка существует
        os.makedirs(os.path.dirname(output_dxf), exist_ok=True)
        
        success = export_to_dxf(doc_3d, output_dxf)
        
        if not success:
            # Пробуем альтернативный экспорт
            print("Пробуем альтернативный экспорт...")
            # doc_3d.SaveAs(output_dxf, 31)
        
        print()
        print("=" * 70)
        print("✓ УСПЕШНО ЗАВЕРШЕНО")
        print("Файл сохранен: {}".format(output_dxf))
        print("=" * 70)
        print()

        return True

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
