# -*- coding: utf-8 -*-
__title__ = "Check Untagged Elements\nby Category"
__author__ = "Tee_เพิ่มพงษ์"
__doc__ = "ตรวจสอบองค์ประกอบที่ยังไม่ได้ทำ Tag ใน View ปัจจุบัน โดยเลือกหมวดหมู่ที่ต้องการ"

from pyrevit import revit, DB, script
from pyrevit import forms
from System.Collections.Generic import List
import sys

doc = revit.doc
uidoc = revit.uidoc
view = doc.ActiveView
output = script.get_output()
output.title = "ตรวจสอบองค์ประกอบที่ยังไม่ได้ทำ Tag"

# --------------------------------------------------------
# 📋 เลือกหมวดหมู่ที่ต้องการตรวจสอบ
# --------------------------------------------------------

def get_all_categories():
    """ดึงรายการหมวดหมู่ทั้งหมดที่มีในโปรเจค"""
    categories = {}
    
    # รายการ BuiltInCategory ที่ถูกต้อง
    common_categories = [
        # งานระบบท่อ (Plumbing)
        DB.BuiltInCategory.OST_PipeCurves,
        DB.BuiltInCategory.OST_PipeFitting,
        DB.BuiltInCategory.OST_PipeAccessory,
        DB.BuiltInCategory.OST_PlumbingFixtures,
        
        # งานระบบปรับอากาศ (HVAC)
        DB.BuiltInCategory.OST_DuctCurves,
        DB.BuiltInCategory.OST_DuctFitting,
        DB.BuiltInCategory.OST_DuctAccessory,
        DB.BuiltInCategory.OST_DuctTerminal,
        DB.BuiltInCategory.OST_MechanicalEquipment,
        
        # งานระบบไฟฟ้า (Electrical)
        DB.BuiltInCategory.OST_ElectricalEquipment,
        DB.BuiltInCategory.OST_ElectricalFixtures,
        DB.BuiltInCategory.OST_LightingFixtures,
        DB.BuiltInCategory.OST_LightingDevices,
        DB.BuiltInCategory.OST_CableTray,
        DB.BuiltInCategory.OST_CableTrayFitting,
        DB.BuiltInCategory.OST_Conduit,
        DB.BuiltInCategory.OST_ConduitFitting,
        
        # งานสถาปัตยกรรม (Architecture)
        DB.BuiltInCategory.OST_Doors,
        DB.BuiltInCategory.OST_Windows,
        DB.BuiltInCategory.OST_Walls,
        DB.BuiltInCategory.OST_Floors,
        DB.BuiltInCategory.OST_Ceilings,
        DB.BuiltInCategory.OST_Roofs,
        DB.BuiltInCategory.OST_Stairs,
        DB.BuiltInCategory.OST_Ramps,
        DB.BuiltInCategory.OST_Furniture,
        DB.BuiltInCategory.OST_GenericModel,
        
        # งานโครงสร้าง (Structure)
        DB.BuiltInCategory.OST_StructuralFraming,
        DB.BuiltInCategory.OST_StructuralColumns,
        DB.BuiltInCategory.OST_StructuralFoundation,
        DB.BuiltInCategory.OST_StructuralStiffener,
        
        # งานอื่นๆ
        DB.BuiltInCategory.OST_Columns,
        DB.BuiltInCategory.OST_Planting,
        DB.BuiltInCategory.OST_Site,
        DB.BuiltInCategory.OST_Parking
    ]
    
    # เพิ่มหมวดหมู่ที่มีองค์ประกอบในโปรเจค
    for builtin_cat in common_categories:
        try:
            collector = DB.FilteredElementCollector(doc, view.Id)\
                          .OfCategory(builtin_cat)\
                          .WhereElementIsNotElementType()\
                          .ToElements()
            
            if len(collector) > 0:
                cat = DB.Category.GetCategory(doc, builtin_cat)
                if cat and cat.Name:
                    # ใช้ภาษาอังกฤษใน string เพื่อป้องกันปัญหา encoding
                    categories[builtin_cat] = "{} ({} items)".format(cat.Name, len(collector))
        except Exception as e:
            continue
    
    return categories

def select_categories():
    """ให้ผู้ใช้เลือกหมวดหมู่ที่ต้องการตรวจสอบ"""
    categories = get_all_categories()
    
    if not categories:
        forms.alert("No elements found in current view", exitscript=True)
        return []
    
    # สร้างรายการสำหรับแสดงในหน้าต่างเลือก
    category_items = []
    for cat, display_text in sorted(categories.items(), key=lambda x: x[1]):
        category_items.append({
            'builtin_cat': cat,
            'display': display_text
        })
    
    # แสดงหน้าต่างให้เลือกหมวดหมู่
    display_list = [item['display'] for item in category_items]
    
    selected_displays = forms.SelectFromList.show(
        display_list,
        title="Select categories to check (Ctrl+click for multiple)",
        button_name="Select Categories",
        multiselect=True,
        width=600,
        height=600
    )
    
    if not selected_displays:
        forms.alert("No categories selected", exitscript=True)
        return []
    
    # แปลงกลับเป็น BuiltInCategory
    selected_categories = []
    for item in category_items:
        if item['display'] in selected_displays:
            selected_categories.append(item['builtin_cat'])
    
    return selected_categories

# --------------------------------------------------------
# 🔍 ดึงองค์ประกอบจากหมวดหมู่ที่เลือก
# --------------------------------------------------------

def safe_collect_elements(selected_categories):
    """ดึงองค์ประกอบจากหมวดหมู่ที่เลือกอย่างปลอดภัย"""
    all_elements = []
    category_counts = {}
    
    for category in selected_categories:
        try:
            collector = DB.FilteredElementCollector(doc, view.Id)\
                          .OfCategory(category)\
                          .WhereElementIsNotElementType()\
                          .ToElements()
            elements = list(collector)
            all_elements.extend(elements)
            
            # นับจำนวนตามหมวดหมู่
            cat = DB.Category.GetCategory(doc, category)
            cat_name = cat.Name if cat else "Unknown"
            category_counts[cat_name] = len(elements)
            
        except Exception as e:
            print("Error collecting category {}: {}".format(category, e))
            continue
    
    return all_elements, category_counts

def safe_collect_tags():
    """ดึง Tag อย่างปลอดภัย"""
    try:
        collector = DB.FilteredElementCollector(doc, view.Id)\
                      .OfClass(DB.IndependentTag)\
                      .WhereElementIsNotElementType()\
                      .ToElements()
        return list(collector)
    except Exception as e:
        print("Error collecting tags: {}".format(e))
        return []

# --------------------------------------------------------
# 🔧 ฟังก์ชันช่วย
# --------------------------------------------------------

def get_tagged_elements(tags):
    """ดึงรายการ element ที่ถูก tag แล้ว"""
    tagged_element_ids = set()
    
    for tag in tags:
        try:
            # วิธีที่ 1: ใช้ GetTaggedLocalElementIds
            if hasattr(tag, 'GetTaggedLocalElementIds'):
                try:
                    tagged_ids = tag.GetTaggedLocalElementIds()
                    for elem_id in tagged_ids:
                        if elem_id and elem_id != DB.ElementId.InvalidElementId:
                            tagged_element_ids.add(elem_id)
                except:
                    pass
            
            # วิธีที่ 2: ใช้ TaggedLocalElementId
            if hasattr(tag, 'TaggedLocalElementId'):
                try:
                    elem_id = tag.TaggedLocalElementId
                    if elem_id and elem_id != DB.ElementId.InvalidElementId:
                        tagged_element_ids.add(elem_id)
                except:
                    pass
            
            # วิธีที่ 3: ใช้ GetTaggedLocalElements
            if hasattr(tag, 'GetTaggedLocalElements'):
                try:
                    referenced_ids = tag.GetTaggedLocalElements()
                    for elem in referenced_ids:
                        if elem and elem.Id:
                            tagged_element_ids.add(elem.Id)
                except:
                    pass
                    
        except Exception as e:
            continue
    
    return tagged_element_ids

def get_element_info(element):
    """ดึงข้อมูลของ element แบบปลอดภัย"""
    try:
        if not element.IsValidObject:
            return None
            
        info = {
            'id': element.Id,
            'name': element.Name if hasattr(element, 'Name') else "No Name",
            'category': element.Category.Name if element.Category and element.Category.Name else "No Category"
        }
        
        # ดึงข้อมูลประเภท
        try:
            type_param = element.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM)
            if type_param and type_param.HasValue:
                type_element = doc.GetElement(type_param.AsElementId())
                info['type'] = type_element.Name if type_element else "Unknown"
            else:
                # ลองดึงจาก Family Name
                family_param = element.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
                if family_param and family_param.HasValue:
                    family_element = doc.GetElement(family_param.AsElementId())
                    info['type'] = family_element.Name if family_element else "Unknown"
                else:
                    info['type'] = "Unknown"
        except:
            info['type'] = "Unknown"
            
        # ดึงระดับ
        try:
            level_id = None
            level_param = element.get_Parameter(DB.BuiltInParameter.LEVEL_PARAM)
            if level_param:
                level_id = level_param.AsElementId()
            
            if level_id and level_id != DB.ElementId.InvalidElementId:
                level = doc.GetElement(level_id)
                info['level'] = level.Name if level and hasattr(level, 'Name') else "N/A"
            else:
                info['level'] = "N/A"
        except:
            info['level'] = "N/A"
            
        return info
        
    except Exception as e:
        return None

# --------------------------------------------------------
# 🎯 เริ่มการทำงานหลัก
# --------------------------------------------------------

try:
    # ให้ผู้ใช้เลือกหมวดหมู่
    output.print_md("## 📋 Step 1: Select Categories to Check")
    output.print_md("Loading categories...")
    
    selected_categories = select_categories()
    
    if not selected_categories:
        script.exit()
    
    # แสดงหมวดหมู่ที่เลือก
    output.print_md("\n### Selected Categories:")
    for cat in selected_categories:
        try:
            cat_name = DB.Category.GetCategory(doc, cat).Name
            output.print_md("- {}".format(cat_name))
        except:
            output.print_md("- {}".format(str(cat)))
    
    output.print_md("---")
    output.print_md("## 🔍 Step 2: Checking Untagged Elements")
    
    # ดึงข้อมูลทั้งหมด
    all_elements, category_counts = safe_collect_elements(selected_categories)
    all_tags = safe_collect_tags()
    
    if not all_elements:
        output.print_md("\n⚠️ **No elements found in selected categories**")
        script.exit()
    
    output.print_md("\n### 📊 Elements by Category:")
    for cat_name, count in sorted(category_counts.items()):
        output.print_md("- **{}**: {} items".format(cat_name, count))
    
    output.print_md("\n### 📊 Summary:")
    output.print_md("- Total elements: **{}**".format(len(all_elements)))
    output.print_md("- Total tags: **{}**".format(len(all_tags)))
    
    # ดึง element ที่ถูก tag แล้ว
    tagged_element_ids = get_tagged_elements(all_tags)
    output.print_md("- Tagged elements: **{}**".format(len(tagged_element_ids)))
    
    # แยกองค์ประกอบที่ยังไม่ได้ทำ Tag
    untagged_elements = []
    valid_element_ids = []
    untagged_by_category = {}
    
    for element in all_elements:
        try:
            if element.IsValidObject:
                if element.Id not in tagged_element_ids:
                    untagged_elements.append(element)
                    valid_element_ids.append(element.Id)
                    
                    # นับตามหมวดหมู่
                    cat_name = element.Category.Name if element.Category else "Unknown"
                    untagged_by_category[cat_name] = untagged_by_category.get(cat_name, 0) + 1
        except Exception as e:
            continue
    
    # --------------------------------------------------------
    # 📋 แสดงผล
    # --------------------------------------------------------
    
    output.print_md("---")
    
    if not untagged_elements:
        output.print_md("## ✅ **No untagged elements found in selected categories**")
    else:
        output.print_md("## ⚠️ **Found {} untagged elements**".format(len(untagged_elements)))
        
        # แสดงสรุปตามหมวดหมู่ (แยกส่วนเพื่อป้องกัน error)
        if untagged_by_category:
            output.print_md("\n### 📊 Untagged Elements by Category:")
            for cat_name, count in sorted(untagged_by_category.items()):
                total = category_counts.get(cat_name, 0)
                if total > 0:
                    percentage = (float(count) / float(total)) * 100.0
                    output.print_md("- **{}**: {} items ({:.1f}%)".format(cat_name, count, percentage))
                else:
                    output.print_md("- **{}**: {} items".format(cat_name, count))
        
        # สร้างตารางแสดงผล
        table_data = []
        display_count = min(len(untagged_elements), 200)
        
        for i, element in enumerate(untagged_elements[:display_count]):
            element_info = get_element_info(element)
            if element_info:
                element_id_link = output.linkify(element.Id)
                table_data.append([
                    element_id_link,
                    element_info['name'],
                    element_info['category'],
                    element_info['type'],
                    element_info['level']
                ])
        
        # แสดงตาราง
        if table_data:
            output.print_md("\n### 📋 Untagged Elements Details")
            output.print_table(
                table_data=table_data,
                columns=["Element ID", "Name", "Category", "Type", "Level"],
                title="Untagged Elements (showing {} items)".format(len(table_data))
            )
            
            if len(untagged_elements) > display_count:
                output.print_md("\n📝 **Note:** Showing first {} of {} items".format(display_count, len(untagged_elements)))
        
        # ไฮไลต์ element
        try:
            if valid_element_ids:
                highlight_limit = min(len(valid_element_ids), 100)
                highlight_ids = valid_element_ids[:highlight_limit]
                
                if highlight_ids:
                    uidoc.Selection.SetElementIds(List[DB.ElementId](highlight_ids))
                    output.print_md("\n🔍 **Highlighted first {} items in view**".format(highlight_limit))
                    
                    if len(valid_element_ids) > highlight_limit:
                        output.print_md("📝 **Note:** Highlighted first {} of {} items".format(highlight_limit, len(valid_element_ids)))
        except Exception as e:
            output.print_md("\n⚠️ **Could not highlight elements**")
        
        # สรุปสถิติรวม
        output.print_md("\n### 📊 Statistics:")
        output.print_md("- Total elements in view: **{}**".format(len(all_elements)))
        output.print_md("- Tagged elements: **{}**".format(len(tagged_element_ids)))
        output.print_md("- Untagged elements: **{}**".format(len(untagged_elements)))
        
        if len(all_elements) > 0:
            percentage = (float(len(untagged_elements)) / float(len(all_elements))) * 100.0
            output.print_md("- **Untagged percentage: {:.1f}%**".format(percentage))
        
        output.print_md("\n🔗 **Click on Element ID to navigate**")
        
except Exception as e:
    output.print_md("## ❌ Error")
    output.print_md("**Error details:** {}".format(str(e)))
    output.print_md("**Error type:** {}".format(type(e).__name__))
    
    import traceback
    error_details = traceback.format_exc()
    output.print_md("**Traceback:**")
    output.print_code(error_details)

# --------------------------------------------------------
# 💡 Tips
# --------------------------------------------------------

output.print_md("---")
output.print_md("## 💡 Tips:")
output.print_md("1. **Click on Element ID** to navigate to the element")
output.print_md("2. **Select multiple categories** using Ctrl key")
output.print_md("3. Use **Tag All** to tag selected elements")
output.print_md("4. Make sure **appropriate tag types** are loaded in project")
output.print_md("5. Display is limited to 200 items for performance")

print("\n✅ Complete")