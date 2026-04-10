# -*- coding: utf-8 -*-
"""Set Base Level - Length Parameter"""

__title__ = "Base_Level\nParameter"

from pyrevit import revit, DB, script

doc = revit.doc
output = script.get_output()

output.print_md("## **Base Level - งานโครงสร้าง สถาปัตย์ และงานระบบ MEP**")

# =====================================================
# ฟังก์ชันหาระดับขององค์ประกอบ
# =====================================================
def get_element_level(element):
    try:
        # วิธีที่ 1: Level จาก LevelId ของ element
        if hasattr(element, 'LevelId') and element.LevelId != DB.ElementId.InvalidElementId:
            level = doc.GetElement(element.LevelId)
            if isinstance(level, DB.Level):
                return level
        
        # วิธีที่ 2: LEVEL_PARAM ขององค์ประกอบ
        level_param = element.get_Parameter(DB.BuiltInParameter.LEVEL_PARAM)
        if level_param and level_param.AsElementId() != DB.ElementId.InvalidElementId:
            level = doc.GetElement(level_param.AsElementId())
            if isinstance(level, DB.Level):
                return level
                
        # วิธีที่ 3: Host ที่มี Level
        if hasattr(element, 'Host') and element.Host:
            host = element.Host
            if hasattr(host, 'LevelId') and host.LevelId != DB.ElementId.InvalidElementId:
                level = doc.GetElement(host.LevelId)
                if isinstance(level, DB.Level):
                    return level
                    
    except:
        pass
    return None


# =====================================================
# รายการ BuiltInCategory
# =====================================================

# งานโครงสร้าง
structural_categories = [
    "OST_StructuralColumns",
    "OST_StructuralFraming",
    "OST_StructuralFoundation",
    "OST_StructuralRebar",
    "OST_StructuralConnections",
    "OST_StructuralStiffener",
    "OST_StructuralTrusses",
    "OST_StructuralBracing",
    "OST_AreaRein",
    "OST_PathRein",
    "OST_FabricReinforcement",
    "OST_StructuralAnchor",
]

# งานสถาปัตย์
architectural_categories = [
    "OST_Walls",
    "OST_Floors",
    "OST_Doors",
    "OST_Windows",
    "OST_Stairs",
    "OST_Railings",
    "OST_Ramps",
    "OST_Ceilings",
    "OST_Roofs",
    "OST_Furniture",
    "OST_GenericModel",
    "OST_Columns",
    "OST_CurtainWallMullions",
    "OST_CurtainPanels",
    "OST_StairsRailing",
]

# งานระบบ (MEP)
MEP_categories = [
    "OST_MechanicalEquipment",
    "OST_ElectricalEquipment",
    "OST_PlumbingFixtures",
    "OST_LightingFixtures",
    "OST_DataDevices",
    "OST_FireAlarmDevices",
    "OST_CommunicationDevices",
    "OST_Conduit",
    "OST_ConduitFitting",
    "OST_CableTray",
    "OST_CableTrayFitting",
    "OST_DuctCurves",
    "OST_DuctFitting",
    "OST_DuctAccessories",
    "OST_PipeCurves",
    "OST_PipeFitting",
    "OST_PipeAccessories",
    "OST_Furniture",
    "OST_SpecialityEquipment",
    "OST_ElectricalFixtures",
    "OST_Sprinklers",
]

# รวมทุกหมวดหมู่
all_categories = structural_categories + architectural_categories + MEP_categories


# =====================================================
# ค้นหาองค์ประกอบทั้งหมด
# =====================================================
output.print_md("### **ค้นหาองค์ประกอบในทุกหมวดหมู่**")

all_elements = []
category_results = {}

structural_results = {}
architectural_results = {}
MEP_results = {}

for cat_name in all_categories:
    try:
        enum_cat = getattr(DB.BuiltInCategory, cat_name)
        collector = DB.FilteredElementCollector(doc).OfCategory(enum_cat).WhereElementIsNotElementType()
        elems = list(collector)

        if elems:
            category_results[cat_name] = len(elems)
            all_elements.extend(elems)

            if cat_name in structural_categories:
                structural_results[cat_name] = len(elems)
            elif cat_name in architectural_categories:
                architectural_results[cat_name] = len(elems)
            elif cat_name in MEP_categories:
                MEP_results[cat_name] = len(elems)

    except:
        continue


# =====================================================
# แสดงผลแยกกลุ่ม
# =====================================================
output.print_md("#### **งานโครงสร้าง (Structural)**")
for name, count in sorted(structural_results.items()):
    output.print_md("- **{}**: {} elements".format(name, count))

output.print_md("#### **งานสถาปัตย์ (Architectural)**")
for name, count in sorted(architectural_results.items()):
    output.print_md("- **{}**: {} elements".format(name, count))

output.print_md("#### **งานระบบ MEP**")
for name, count in sorted(MEP_results.items()):
    output.print_md("- **{}**: {} elements".format(name, count))


# =====================================================
# ลบองค์ประกอบซ้ำ
# =====================================================
unique_elements = []
seen_ids = set()

for e in all_elements:
    if e.Id not in seen_ids:
        unique_elements.append(e)
        seen_ids.add(e.Id)


# =====================================================
# สรุปจำนวน
# =====================================================
output.print_md("### **สรุป**")
output.print_md("**รวมงานโครงสร้าง:** {} รายการ".format(sum(structural_results.values())))
output.print_md("**รวมงานสถาปัตย์:** {} รายการ".format(sum(architectural_results.values())))
output.print_md("**รวมงานระบบ MEP:** {} รายการ".format(sum(MEP_results.values())))
output.print_md("**องค์ประกอบทั้งหมด (ไม่ซ้ำ):** {} รายการ".format(len(unique_elements)))

if not unique_elements:
    script.exit()


# =====================================================
# เริ่มตั้งค่า Base_Level
# =====================================================
output.print_md("### **กำลังตั้งค่า Base_Level ให้ทุกรายการ...**")

t = DB.Transaction(doc, "Set Base Level - ST, AR & MEP")
t.Start()

success_count = 0
level_count = 0
read_only_count = 0

struct_success = 0
arch_success = 0
MEP_success = 0

# เพื่อเทียบชื่อหมวดหมู่จาก Category.Name
MEP_names = []
for c in MEP_categories:
    try:
        MEP_names.append(doc.Settings.Categories.get_Item(getattr(DB.BuiltInCategory, c)).Name)
    except:
        pass

for e in unique_elements:
    try:
        level = get_element_level(e)
        if not level:
            continue

        level_count += 1

        p = e.LookupParameter("Base_Level")
        if not p:
            continue
        
        if p.StorageType == DB.StorageType.Double and not p.IsReadOnly:

            # ตั้งค่า Base_Level ด้วยค่าฟุต (Revit ใช้ฟุตภายใน)
            p.Set(level.Elevation)

            success_count += 1

            cat_name = e.Category.Name

            # จัดเข้ากลุ่ม
            if "Structural" in cat_name or "Foundation" in cat_name or "Rebar" in cat_name:
                struct_success += 1
            elif cat_name in MEP_names:
                MEP_success += 1
            else:
                arch_success += 1

        else:
            if p.IsReadOnly:
                read_only_count += 1

    except:
        continue

t.Commit()


# =====================================================
# แสดงผลลัพธ์
# =====================================================
output.print_md("### **ผลการตั้ง Base_Level**")
output.print_md("- จำนวนองค์ประกอบทั้งหมด: {}".format(len(unique_elements)))
output.print_md("- องค์ประกอบที่มี Level: {}".format(level_count))
output.print_md("---")
output.print_md("#### **ตั้งค่า Base_Level สำเร็จ:**")
output.print_md("- งานโครงสร้าง: **{}** รายการ".format(struct_success))
output.print_md("- งานสถาปัตย์: **{}** รายการ".format(arch_success))
output.print_md("- งานระบบ MEP: **{}** รายการ".format(MEP_success))
output.print_md("- รวมทั้งหมด: **{}** รายการ".format(success_count))

if read_only_count > 0:
    output.print_md("⚠️ Base_Level เป็น Read-only: {} รายการ".format(read_only_count))


# =====================================================
# ตัวอย่างค่าที่ตั้ง (หน่วย "เมตร")
# =====================================================
output.print_md("### **ตัวอย่างค่าที่ตั้ง Base_Level (แสดงเป็นเมตร)**")

shown = 0
for e in unique_elements:
    if shown >= 3:
        break

    level = get_element_level(e)
    p = e.LookupParameter("Base_Level")

    if not (level and p):
        continue

    elev_m = level.Elevation * 0.3048
    base_m = p.AsDouble() * 0.3048

    output.print_md("{}. **{}**".format(shown+1, e.Category.Name))
    output.print_md("   - Level: {}".format(level.Name))
    output.print_md("   - Elevation: {:.3f} m".format(elev_m))
    output.print_md("   - Base_Level: {:.3f} m".format(base_m))

    shown += 1


output.print_md("---")
output.print_md("**เสร็จสิ้น — แสดงผลเป็นหน่วยเมตรเรียบร้อย**")
