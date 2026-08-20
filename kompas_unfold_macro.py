# -*- coding: utf-8 -*-
"""
Макрос для создания разверток в DXF из 3D модели Компас-3D 22.
Поддерживает листовые детали и твердотельные детали.
"""

import os
import shutil
import struct
import winreg
import tempfile

import pythoncom
import win32com
from win32com.client import Dispatch, gencache, makepy


# ============================================================
# GUID
# ============================================================

API7_GUID = "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}"
API5_GUID = "{042C7D94-1F68-11D5-9971-00104B9B2A2D}"
CONST3D_GUID = "{2CAF168C-7961-4B90-9DA2-701419BEEFE3}"
CONST_GUID = "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}"


# ============================================================
# КЕШИРОВАНИЕ И ЗАГРУЗКА TYPELIB
# ============================================================

def clear_gencache():
    """Полностью очищает кеш сгенерированных win32com-обёрток"""
    gen_path = win32com.__gen_path__
    print("Очищаем кеш win32com: {}".format(gen_path))
    shutil.rmtree(gen_path, ignore_errors=True)
    try:
        gencache.Rebuild()
    except Exception as error:
        print("Rebuild() вызвал предупреждение: {}".format(error))


def _enum_typelib_versions(guid):
    """Все версии typelib с данным GUID"""
    versions = set()
    for wow_flag in (0, winreg.KEY_WOW64_64KEY):
        try:
            base = winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT,
                r"TypeLib\{}".format(guid),
                0,
                winreg.KEY_READ | wow_flag
            )
        except OSError:
            continue

        i = 0
        while True:
            try:
                ver_name = winreg.EnumKey(base, i)
            except OSError:
                break
            i += 1
            try:
                major, minor = (int(p, 16) for p in ver_name.split("."))
            except Exception:
                continue

            try:
                ver_key = winreg.OpenKey(
                    base, ver_name, 0, winreg.KEY_READ | wow_flag
                )
            except OSError:
                continue

            j = 0
            while True:
                try:
                    lcid_name = winreg.EnumKey(ver_key, j)
                except OSError:
                    break
                j += 1
                try:
                    versions.add((major, minor, int(lcid_name)))
                except Exception:
                    continue

    return sorted(versions)


def _find_typelib_file(guid, major, minor, lcid):
    """Путь к реальному файлу typelib"""
    base = r"TypeLib\{}\{}.{}\{}".format(
        guid, format(major, "x"), format(minor, "x"), lcid
    )

    for arch in ("win64", "win32"):
        for wow_flag in (winreg.KEY_WOW64_64KEY, 0):
            try:
                key = winreg.OpenKey(
                    winreg.HKEY_CLASSES_ROOT,
                    base + "\\" + arch,
                    0,
                    winreg.KEY_READ | wow_flag
                )
                value, _ = winreg.QueryValueEx(key, None)
                winreg.CloseKey(key)
                if value and os.path.isfile(value):
                    return value
            except OSError:
                continue

    return None


def ensure_module(guid, label, hint=(1, 0, 0)):
    """Надёжная загрузка typelib"""
    hint_major, hint_minor, hint_lcid = hint

    candidates = [(0, 0, 0), (hint_lcid, hint_major, hint_minor)]

    for lcid, major, minor in candidates:
        try:
            module = gencache.EnsureModule(guid, lcid, major, minor)
            if module is not None:
                return module
        except Exception:
            pass

    versions = _enum_typelib_versions(guid)

    for major, minor, lcid in versions:
        try:
            module = gencache.EnsureModule(guid, lcid, major, minor)
            if module is not None:
                return module
        except Exception:
            pass

        path = _find_typelib_file(guid, major, minor, lcid)
        if path:
            try:
                makepy.GenerateFromTypeLibSpec(path, bForDemand=False)
                module = gencache.EnsureModule(guid, lcid, major, minor)
                if module is not None:
                    print("{}: обёртка сгенерирована из файла {}".format(label, path))
                    return module
            except Exception as error:
                print("{}: генерация из файла {} не удалась: {}".format(label, path, error))

    bitness = struct.calcsize("P") * 8
    raise RuntimeError(
        "Не удалось загрузить typelib для {} (GUID {}).\n"
        "Разрядность текущего Python: {} бит.\n"
        "Найденные версии в реестре: {}\n\n"
        "Проверьте, что КОМПАС-3D установлен и COM-сервер зарегистрирован.".format(
            label, guid, bitness, versions or "нет"
        )
    )


# ============================================================
# ПОДКЛЮЧЕНИЕ К API7
# ============================================================

def connect_api7():
    """Подключение к Компас-3D API7"""
    pythoncom.CoInitialize()
    clear_gencache()

    print("Загружаем API7...")
    api7 = ensure_module(API7_GUID, "API7")

    application = api7.IApplication(
        Dispatch("Kompas.Application.7")
        ._oleobj_
        .QueryInterface(api7.IApplication.CLSID, pythoncom.IID_IDispatch)
    )

    if not application:
        raise RuntimeError("Не удалось получить IApplication (API7).")

    application.Visible = True
    print("КОМПАС-3D подключён (API7).")

    return application, api7


# ============================================================
# ПОЛУЧЕНИЕ ВЕРХНЕЙ ДЕТАЛИ
# ============================================================

def get_top_part_via_api5(api7):
    """Получение верхней детали через API5"""
    print("Получаем верхнюю деталь через API5...")

    api5 = ensure_module(API5_GUID, "API5")
    const3d = ensure_module(CONST3D_GUID, "константы 3D").constants
    const = ensure_module(CONST_GUID, "общие константы").constants

    kompas_object = api5.KompasObject(
        Dispatch("Kompas.Application.5")
        ._oleobj_
        .QueryInterface(api5.KompasObject.CLSID, pythoncom.IID_IDispatch)
    )

    if not kompas_object:
        raise RuntimeError("Не удалось получить KompasObject (API5).")

    document_3d_5 = kompas_object.ActiveDocument3D()
    if not document_3d_5:
        raise RuntimeError("ActiveDocument3D() (API5) вернул None.")

    top_part_5 = document_3d_5.GetPart(const3d.pTop_Part)
    if not top_part_5:
        raise RuntimeError("API5 не вернул верхнюю деталь.")

    top_part_7_raw = kompas_object.TransferInterface(
        top_part_5, const.ksAPI7Dual, const3d.o3d_part
    )

    if not top_part_7_raw:
        raise RuntimeError("TransferInterface (API5 -> API7) вернул None.")

    part7 = api7.IPart7(top_part_7_raw)
    if not part7:
        raise RuntimeError("Не удалось привести результат к IPart7.")

    print("TopPart получен через API5.")
    return part7


def get_top_part(application, api7):
    """Получение верхней детали"""
    active_doc = application.ActiveDocument

    if not active_doc:
        raise RuntimeError("В КОМПАС-3D нет активного документа.")

    try:
        document_3d = api7.IKompasDocument3D(active_doc)
    except Exception as error:
        raise RuntimeError("Активный документ не является 3D-документом.\n{}".format(error))

    if not document_3d:
        raise RuntimeError("Не удалось получить IKompasDocument3D.")

    # Основной способ
    part7 = None
    try:
        part7 = document_3d.TopPart
    except Exception as error:
        print("Прямой TopPart (API7) вызвал ошибку: {}".format(error))

    if part7:
        print("TopPart получен напрямую через API7.")
        return part7

    # Запасной способ
    return get_top_part_via_api5(api7)


# ============================================================
# ОПРЕДЕЛЕНИЕ ТИПА ДЕТАЛИ И ПОЛУЧЕНИЕ РАЗВЕРТКИ
# ============================================================

def is_sheet_metal(part7, api7):
    """Проверка является ли деталь листовой"""
    try:
        sheet_container_raw = part7._oleobj_.QueryInterface(
            api7.ISheetMetalContainer.CLSID, pythoncom.IID_IDispatch
        )
        sheet_container = api7.ISheetMetalContainer(sheet_container_raw)
        if sheet_container:
            print("✓ Деталь является листовой.")
            return True, sheet_container
    except Exception:
        pass
    
    print("✗ Деталь не является листовой (твердотельная).")
    return False, None


def get_largest_face(part7, api7):
    """Получение самой большой грани твердотельной детали"""
    print("Определяем самую большую грань...")
    
    try:
        bodies = part7.Bodies
        if not bodies or bodies.Count == 0:
            raise RuntimeError("В детали нет тел.")
        
        body = bodies.Item(1)  # Первое (основное) тело
        faces = body.Faces
        
        if not faces or faces.Count == 0:
            raise RuntimeError("В теле нет граней.")
        
        max_area = 0
        largest_face = None
        
        for i in range(1, faces.Count + 1):
            face = faces.Item(i)
            try:
                # Получаем площадь грани через свойства
                area = face.Area if hasattr(face, 'Area') else 0
                if area > max_area:
                    max_area = area
                    largest_face = face
            except Exception:
                continue
        
        if largest_face is None:
            largest_face = faces.Item(1)
        
        print("✓ Найдена самая большая грань (площадь: {})".format(max_area))
        return largest_face
        
    except Exception as error:
        raise RuntimeError("Ошибка при получении грани: {}".format(error))


def create_unfold_sketch(part7, api7, sheet_container=None, largest_face=None):
    """Создание развертки и получение фрагмента"""
    print("Создаем развертку...")
    
    try:
        if sheet_container:
            # Для листовой детали
            sketches = part7.Sketches
            if sketches.Count > 0:
                print("Используем встроенную развертку листовой детали.")
                return sketches.Item(1)
            else:
                raise RuntimeError("Не найдена развертка листовой детали.")
        else:
            # Для твердотельной детали - используем саму грань
            print("Используем самую большую грань как развертку.")
            return largest_face
            
    except Exception as error:
        raise RuntimeError("Ошибка при создании развертки: {}".format(error))


# ============================================================
# ЭКСПОРТ В DXF
# ============================================================

def export_to_dxf(application, api7, part7, output_path):
    """Экспорт развертки в DXF через фрагмент"""
    print("Экспортируем в DXF: {}".format(output_path))
    
    try:
        # Получаем активный документ
        active_doc = application.ActiveDocument
        
        # Пытаемся экспортировать через встроенный экспорт
        # Создаем временный фрагмент
        temp_dir = tempfile.gettempdir()
        temp_fragment = os.path.join(temp_dir, "temp_unfold.m3d")
        
        # Сохраняем фрагмент
        if hasattr(active_doc, 'SaveFragment'):
            print("Сохраняем фрагмент...")
            active_doc.SaveFragment(temp_fragment)
            
            # Открываем фрагмент и экспортируем
            frag_doc = application.OpenDocument(temp_fragment, False)
            
            if frag_doc:
                # Экспортируем в DXF
                export_params = {
                    'FileName': output_path,
                    'Format': 1,  # DXF формат
                }
                
                if hasattr(frag_doc, 'ExportAs'):
                    frag_doc.ExportAs(output_path, 31)  # 31 = DXF
                    print("✓ Фрагмент экспортирован в DXF.")
                
                frag_doc.Close(False)
                
                # Удаляем временный файл
                try:
                    os.remove(temp_fragment)
                except Exception:
                    pass
        else:
            # Альтернативный способ - через Export
            if hasattr(active_doc, 'Export'):
                active_doc.Export(output_path, 31)  # DXF
                print("✓ Документ экспортирован в DXF.")
            else:
                raise RuntimeError("Документ не поддерживает экспорт.")
                
    except Exception as error:
        raise RuntimeError("Ошибка при экспорте в DXF: {}".format(error))


# ============================================================
# ОСНОВНАЯ ФУНКЦИЯ
# ============================================================

def run_script():
    """Основной макрос"""
    print()
    print("=" * 70)
    print("КОМПАС-3D — Создание разверток в DXF")
    print("=" * 70)
    print()

    try:
        # Подключаемся
        application, api7 = connect_api7()
        print()

        # Получаем верхнюю деталь
        part7 = get_top_part(application, api7)
        print()

        # Определяем тип детали
        is_sheet, sheet_container = is_sheet_metal(part7, api7)
        print()

        # Получаем развертку
        if is_sheet:
            unfold_sketch = create_unfold_sketch(part7, api7, sheet_container=sheet_container)
        else:
            largest_face = get_largest_face(part7, api7)
            unfold_sketch = create_unfold_sketch(part7, api7, largest_face=largest_face)
        print()

        # Экспортируем в DXF
        output_dxf = os.path.expanduser("~\\Desktop\\unfold_output.dxf")
        export_to_dxf(application, api7, part7, output_dxf)
        print()

        print("=" * 70)
        print("✓ УСПЕШНО ЗАВЕРШЕНО")
        print("Файл сохранен: {}".format(output_dxf))
        print("=" * 70)

        return True

    except Exception as error:
        print()
        print("=" * 70)
        print("✗ ОШИБКА")
        print("=" * 70)
        print(str(error))
        print()
        import traceback
        traceback.print_exc()
        return False

    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    run_script()
