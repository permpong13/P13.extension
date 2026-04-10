# -*- coding: utf-8 -*-
__title__ = "Numbering\nby Manual"
__author__ = "เพิ่มพงษ์ ทวีกุล"
__doc__ = "พิมพ์ชื่อ Parameter เองเพื่อความแม่นยำ (รองรับ Shared / Built-in / Type)"

from pyrevit import revit, DB, forms, script

doc = revit.doc
cfg = script.get_config()
BIC = DB.BuiltInCategory

# --- ระบบจดจำการตั้งค่า ---
def get_saved_setting(category_name, key, default):
    return cfg.get_option("{}_{}".format(category_name, key), default)

def save_setting(category_name, key, value):
    cfg.set_option("{}_{}".format(category_name, key), str(value))
    script.save_config()

# --- ฟังก์ชันเขียนค่าลง Parameter (พยายามหาทั้ง Instance และ Type) ---
def set_param_value(elem, param_name, value):
    # 1. ลองหาใน Instance ก่อน
    p = elem.LookupParameter(param_name)
    
    # 2. กรณีพิเศษถ้าพิมพ์ว่า Mark ให้ลองใช้ Built-in ID
    if not p and param_name.lower() == "mark":
        p = elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
        
    # 3. ถ้าหาที่ Instance ไม่เจอ ให้ลองไปหาที่ Type
    if not p:
        t_id = elem.GetTypeId()
        if t_id != DB.ElementId.InvalidElementId:
            t_type = doc.GetElement(t_id)
            p = t_type.LookupParameter(param_name)
            # กรณีพิมพ์ว่า Type Mark
            if not p and param_name.lower() == "type mark":
                p = t_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK)

    # เมื่อเจอ Parameter แล้วทำการเขียนค่า
    if p and not p.IsReadOnly:
        try:
            if p.StorageType == DB.StorageType.String:
                p.Set(str(value))
            elif p.StorageType == DB.StorageType.Integer:
                p.Set(int(value))
            elif p.StorageType == DB.StorageType.Double:
                p.Set(float(value))
            return True
        except:
            return False
    return False

# --- Main Execution ---
def run_tool():
    category_mapping = {
        "Doors": BIC.OST_Doors,
        "Windows": BIC.OST_Windows,
        "Furniture": BIC.OST_Furniture,
        "Rooms": BIC.OST_Rooms,
        "Generic Models": BIC.OST_GenericModel,
        "Viewports": BIC.OST_Viewports
    }

    # 1. เลือกหมวดหมู่ก่อน (เพื่อให้โปรแกรมรู้ว่าจะคลิกเลือกอะไรได้บ้าง)
    sel_cat_name = forms.CommandSwitchWindow.show(category_mapping.keys(), message="เลือกหมวดหมู่ที่ต้องการทำงาน:")
    if not sel_cat_name: return
    bicat_enum = category_mapping[sel_cat_name]

    # 2. พิมพ์ชื่อ Parameter เอง (ไม่ต้องเลือกจาก List แล้ว)
    last_p_name = get_saved_setting(sel_cat_name, "manual_param", "Mark")
    selected_param = forms.ask_for_string(
        default=last_p_name,
        title="ระบุชื่อ Parameter",
        prompt="พิมพ์ชื่อ Parameter ที่ต้องการใส่ค่า (Shared / Built-in / Instance / Type):"
    )

    if not selected_param: return
    save_setting(sel_cat_name, "manual_param", selected_param)

    # 3. ตั้งค่าตัวเลข
    digits = forms.ask_for_string(default=get_saved_setting(sel_cat_name, "digits", "3"), prompt="จำนวนหลัก (เช่น 3 คือ 001):")
    if not digits: digits = "1"
    save_setting(sel_cat_name, "digits", digits)

    prefix = forms.ask_for_string(default=get_saved_setting(sel_cat_name, "prefix", ""), prompt="คำนำหน้า (Prefix):")
    if prefix is None: prefix = ""
    save_setting(sel_cat_name, "prefix", prefix)

    # 4. เริ่มการคลิกเลือกชิ้นงาน
    with revit.TransactionGroup("Renumber " + sel_cat_name):
        counter = 1
        while True:
            try:
                pick = revit.pick_element_by_category(bicat_enum, "คลิกเลือก Object (กด ESC เพื่อจบการทำงาน)")
                if not pick: break

                new_val = "{}{}".format(prefix, str(counter).zfill(int(digits)))

                with revit.Transaction("Set " + selected_param):
                    if set_param_value(pick, selected_param, new_val):
                        counter += 1
                    else:
                        forms.alert("ไม่พบ Parameter ชื่อ '{}' ในชิ้นงานที่เลือก หรือเป็นค่าที่แก้ไขไม่ได้".format(selected_param), exits=False)
            except:
                # กรณี User กด ESC
                break

if __name__ == "__main__":
    run_tool()