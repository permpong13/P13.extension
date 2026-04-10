# -*- coding: utf-8 -*-
"""
Script Name: Create Pipe Isometric
Description: สร้าง 3D View แบบ Isometric จากท่อที่เลือก, ตัด Section Box ให้พอดี, 
             ล็อควิวเพื่อใส่ Tag และทำการติด Tag ให้ท่อโดยอัตโนมัติ
"""

from pyrevit import revit, DB, UI, script, forms
from System.Collections.Generic import List
import math

# ตัวแปรพื้นฐาน
doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# =================================================================
#                         HELPER FUNCTIONS
# =================================================================

def get_selected_elements():
    """ดึง Element ที่ user เลือกอยู่ปัจจุบัน"""
    selection = uidoc.Selection.GetElementIds()
    if not selection:
        forms.alert("กรุณาเลือกท่อ (Pipes) หรืออุปกรณ์ที่ต้องการก่อนรันคำสั่ง", exitscript=True)
    return [doc.GetElement(id) for id in selection]

def get_3d_view_family_type():
    """ฟังก์ชันช่วยหา ViewFamilyType สำหรับ 3D"""
    types = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).ToElements()
    
    for t in types:
        if t.ViewFamily == DB.ViewFamily.ThreeDimensional:
            return t
    return None

def find_existing_view(view_name):
    """ฟังก์ชันช่วยหา View ที่มีชื่อซ้ำ"""
    views = DB.FilteredElementCollector(doc).OfClass(DB.View3D).ToElements()
    for v in views:
        if v.Name == view_name and not v.IsTemplate:
            return v
    return None

def find_safe_view():
    """ค้นหา View ที่ปลอดภัยสำหรับการเปลี่ยนไปเปิดใช้งานก่อนทำการลบ View"""
    views = DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements()
    for v in views:
        # หากพบ View 3D ชื่อ "{3D}" หรือ View Plan ใดๆ ที่ไม่ใช่ Template
        if v.Name == "{3D}" or isinstance(v, DB.ViewPlan) and not v.IsTemplate:
            return v
    return None

# =================================================================
#                         MAIN PROCESS
# =================================================================

def create_isometric_view(elements, view_name="ISO_Generated_View"):
    # 1. หา ViewFamilyType ของ 3D View
    view_family_type = get_3d_view_family_type()
    if not view_family_type:
        forms.alert("ไม่พบ ViewFamilyType สำหรับ 3D View ในโปรเจกต์นี้", exitscript=True)

    # --- START: Active View Management (นอก Transaction) ---
    existing_view = find_existing_view(view_name)

    if existing_view:
        # ตรวจสอบว่า View เก่าเป็น Active View หรือไม่
        if existing_view.Id == uidoc.ActiveView.Id:
            safe_view_to_switch = find_safe_view()
            
            if safe_view_to_switch:
                # เปลี่ยน Active View ไปที่ View ที่ปลอดภัยก่อน (นอก Transaction)
                uidoc.RequestViewChange(safe_view_to_switch)
            else:
                # หากหา View อื่นเปิดไม่ได้ ให้แจ้งเตือนและยกเลิก
                forms.alert("ไม่สามารถลบ View เก่าได้เนื่องจาก View นั้นกำลังถูกเปิดใช้งาน (Active View). กรุณาตรวจสอบว่ามี View อื่นที่สามารถเปิดใช้งานได้ใน Project", exitscript=True)
                return None
    # --- END: Active View Management ---

    # 2. เริ่ม Transaction เพื่อสร้างวิว (ตอนนี้ปลอดภัยแล้วที่จะลบ)
    t = DB.Transaction(doc, "Create Isometric View")
    t.Start()
    
    new_view = None

    try:
        # ลบ View เก่าทิ้ง (ถ้ามี)
        if existing_view:
            doc.Delete(existing_view.Id)

        # สร้าง View ใหม่
        new_view = DB.View3D.CreateIsometric(doc, view_family_type.Id)
        
        doc.Regenerate() 
        new_view.Name = view_name

        # 3. คำนวณ Bounding Box รวมของ Elements ที่เลือก
        min_pt = [float('inf'), float('inf'), float('inf')]
        max_pt = [float('-inf'), float('-inf'), float('-inf')]
        has_geometry = False

        for el in elements:
            bbox = el.get_BoundingBox(None)
            if bbox:
                has_geometry = True
                min_pt[0] = min(min_pt[0], bbox.Min.X)
                min_pt[1] = min(min_pt[1], bbox.Min.Y)
                min_pt[2] = min(min_pt[2], bbox.Min.Z)
                
                max_pt[0] = max(max_pt[0], bbox.Max.X)
                max_pt[1] = max(max_pt[1], bbox.Max.Y)
                max_pt[2] = max(max_pt[2], bbox.Max.Z)

        if has_geometry:
            # ตรวจสอบว่าโปรเจกต์ใช้หน่วยเมตริกหรือไม่เพื่อกำหนด Offset
            is_metric = False
            try:
                metric_param = doc.ProjectInformation.get_Parameter(DB.BuiltInParameter.PROJECT_MEASURED_IN_METRIC)
                if metric_param and metric_param.AsInteger() == 1:
                    is_metric = True
            except:
                pass 

            offset = 1.0 # Default: 1 ฟุต (สำหรับ Imperial)
            if is_metric:
                offset = 0.3048 # 0.3048 เมตร (ประมาณ 1 ฟุต) สำหรับ Metric
            
            new_bbox = DB.BoundingBoxXYZ()
            new_bbox.Min = DB.XYZ(min_pt[0] - offset, min_pt[1] - offset, min_pt[2] - offset)
            new_bbox.Max = DB.XYZ(max_pt[0] + offset, max_pt[1] + offset, max_pt[2] + offset)

            # เปิดใช้งาน Section Box และตั้งค่าขนาด
            new_view.IsSectionBoxActive = True
            new_view.SetSectionBox(new_bbox)

            # 4. Isolate เฉพาะ Elements ที่เลือก (Temporary Isolate)
            ids_to_isolate = List[DB.ElementId]([el.Id for el in elements])
            new_view.IsolateElementsTemporary(ids_to_isolate)

            # 5. Lock View (Save Orientation) เพื่อให้ใส่ Tag ได้
            new_view.SaveOrientationAndLock()
            
            # 6. ปรับการแสดงผล
            new_view.DetailLevel = DB.ViewDetailLevel.Fine
            # ตั้งค่า Visual Style เป็น Hidden Line (ค่า 4)
            visual_style_param = new_view.get_Parameter(DB.BuiltInParameter.MODEL_GRAPHICS_STYLE)
            if visual_style_param:
                visual_style_param.Set(4) 

        t.Commit()
        
        # สลับหน้าจอไปที่ View ใหม่ (นอก Transaction แล้ว)
        # เนื่องจากเราเปลี่ยน Active View ไปก่อนหน้านี้ (ถ้ามี) การ RequestViewChange ครั้งนี้จะทำงานต่อเมื่อรันสคริปต์เสร็จ
        # แต่เราสามารถ RequestViewChange ที่นี่ได้เลยเพราะ Transaction ได้ Commit ไปแล้ว
        uidoc.RequestViewChange(new_view)
        return new_view

    except Exception as e:
        if t.GetStatus() == DB.TransactionStatus.Started:
            t.RollBack()
        output.print_md("## Error Occurred During View Creation")
        output.print_md(str(e))
        return None

# =================================================================
#                         AUTO-TAG FUNCTION (แก้ไข 7 Arguments แล้ว)
# =================================================================

def auto_tag_pipes(iso_view, elements):
    """ทำการติด Tag ให้กับท่อและอุปกรณ์ใน View ที่กำหนด"""
    
    PIPE_CAT_ID = DB.ElementId(DB.BuiltInCategory.OST_PipeCurves) 
    FITTING_CAT_ID = DB.ElementId(DB.BuiltInCategory.OST_PipeFitting) 
    ACCESSORY_CAT_ID = DB.ElementId(DB.BuiltInCategory.OST_PipeAccessory) 
    
    t_tag = DB.Transaction(doc, "Auto Tag Pipes")
    t_tag.Start()
    
    tagged_count = 0

    try:
        relevant_elements = [
            el for el in elements 
            if el.Category and el.Category.Id in [PIPE_CAT_ID, FITTING_CAT_ID, ACCESSORY_CAT_ID]
        ]

        for el in relevant_elements:
            location = el.Location
            tag_point = None
            
            if isinstance(location, DB.LocationCurve):
                curve = location.Curve
                tag_point = curve.Evaluate(0.5, True)
                
            elif isinstance(location, DB.LocationPoint):
                tag_point = location.Point
            
            if tag_point:
                offset = 1.5 
                offset_point = tag_point + DB.XYZ(offset, offset, 0)
                
                # --- แก้ไข: เพิ่ม TagMode (Argument ที่ 7) ---
                DB.IndependentTag.Create(
                    doc, 
                    iso_view.Id, 
                    el.Id, 
                    True, # isLeaderVisible 
                    DB.TagOrientation.Horizontal, 
                    offset_point,
                    DB.TagMode.TM_ADDBY_ELEMENT # <<< นี่คือ Argument ที่ 7
                )
                tagged_count += 1
                
        t_tag.Commit()
        
        output.print_md("---")
        output.print_md("✅ **Auto-Tagging Complete:** ใส่ Tag ให้กับ {} Elements".format(tagged_count))
        
    except Exception as e:
        if t_tag.GetStatus() == DB.TransactionStatus.Started:
            t_tag.RollBack()
        output.print_md("## ⚠️ Warning: Auto-Tagging Failed")
        output.print_md("Error: {}".format(e))
        output.print_md("สาเหตุ: **โปรดตรวจสอบว่ามี Tag Family (เช่น Multi-Category Tag) โหลดอยู่ใน Project**")

# =================================================================
#                         EXECUTION
# =================================================================

if __name__ == '__main__':
    selected_elements = get_selected_elements()
    isometric_view = create_isometric_view(selected_elements)
    
    if isometric_view:
        auto_tag_pipes(isometric_view, selected_elements)