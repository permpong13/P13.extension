# -*- coding: utf-8 -*-
__title__ = "Update\nCoordinates"
__doc__ = "คำนวณพิกัด (เมตร) ลงในพารามิเตอร์ N/E_Coordinate\n- แก้ไข Error ตอนสรุปผล (Rollback Fix)\n- รองรับ Revit 2024/2025/2026+"
__author__ = "เพิ่มพงษ์ & Gemini"

from pyrevit import forms, script, DB, HOST_APP
import math
import sys

# ================================================================
# 1. MATHEMATICAL LOGIC
# ================================================================

def rotate(x, y, theta):
    """หมุนพิกัดตามมุม (Radians)"""
    rotated_x = math.cos(theta) * x + math.sin(theta) * y
    rotated_y = -math.sin(theta) * x + math.cos(theta) * y
    return (rotated_x, rotated_y)

def find_cord(x, y, theta, bp_ew, bp_ns):
    """
    คำนวณพิกัดจริง (Logic เดียวกับต้นฉบับ)
    """
    rotated_x, rotated_y = rotate(x, y, theta)
    actual_east = rotated_x + bp_ew
    actual_north = rotated_y + bp_ns
    # Return (North, East)
    return (actual_north, actual_east)

def feet_to_meters(value):
    return value * 0.3048

class WarningSwallower(DB.IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        fails = failuresAccessor.GetFailureMessages()
        for f in fails:
            if f.GetSeverity() == DB.FailureSeverity.Warning:
                failuresAccessor.DeleteWarning(f)
        return DB.FailureProcessingResult.Continue

# ================================================================
# 2. MAIN SCRIPT
# ================================================================

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

# --- STEP 1: UI Selection ---
options_category = {
    '❄️ Mechanical Equipment': DB.BuiltInCategory.OST_MechanicalEquipment,
    '⚡ Electrical Equipment': DB.BuiltInCategory.OST_ElectricalEquipment,
    '🚽 Plumbing Fixtures': DB.BuiltInCategory.OST_PlumbingFixtures,
    '💡 Lighting Fixtures': DB.BuiltInCategory.OST_LightingFixtures,
    '🔌 Data Devices': DB.BuiltInCategory.OST_DataDevices,
    '🔔 Fire Alarm Devices': DB.BuiltInCategory.OST_FireAlarmDevices,
    '📡 Communication Devices': DB.BuiltInCategory.OST_CommunicationDevices,
    '💨 Air Terminals': DB.BuiltInCategory.OST_DuctTerminal,
    '🔧 Pipe Fittings': DB.BuiltInCategory.OST_PipeFitting,
    '⚙️ Pipe Accessories': DB.BuiltInCategory.OST_PipeAccessory,
    '🌬️ Duct Fittings': DB.BuiltInCategory.OST_DuctFitting,
    '🚪 Doors': DB.BuiltInCategory.OST_Doors,
    '🪟 Windows': DB.BuiltInCategory.OST_Windows,
    '🏛️ Columns': DB.BuiltInCategory.OST_Columns,
    '🏗️ Structural Columns': DB.BuiltInCategory.OST_StructuralColumns,
    '🧱 Structural Foundations': DB.BuiltInCategory.OST_StructuralFoundation,
    '🪑 Furniture': DB.BuiltInCategory.OST_Furniture,
    '🧩 Generic Models': DB.BuiltInCategory.OST_GenericModel,
    '📝 Detail Components': DB.BuiltInCategory.OST_DetailComponents,
}

selected_keys = forms.SelectFromList.show(
    sorted(options_category.keys()),
    multiselect=True,
    title="เลือกหมวดหมู่ (Fix Rollback Error)",
    button_name="🚀 เริ่มคำนวณ"
)

if not selected_keys:
    sys.exit()

# รวบรวม Elements
elements = []
for key in selected_keys:
    bic = options_category[key]
    col = DB.FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElements()
    elements.extend(list(col))

if not elements:
    forms.alert("❌ ไม่พบ Element ในหมวดหมู่ที่เลือก")
    sys.exit()

# --- STEP 2: Get Base Point Info ---
angle = 0.0
bp_ewest_m = 0.0
bp_nsouth_m = 0.0
base_point_info = "Default"

base_points = DB.FilteredElementCollector(doc).OfClass(DB.BasePoint).ToElements()
for bp in base_points:
    if not bp.IsShared:
        angle_param = bp.get_Parameter(DB.BuiltInParameter.BASEPOINT_ANGLETON_PARAM)
        if angle_param:
            angle = angle_param.AsDouble()
        
        raw_ns = bp.get_Parameter(DB.BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble()
        raw_ew = bp.get_Parameter(DB.BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble()
        
        rot_x, rot_y = rotate(bp.Position.X, bp.Position.Y, angle)
        
        bp_nsouth_ft = raw_ns - rot_y
        bp_ewest_ft = raw_ew - rot_x
        
        bp_nsouth_m = feet_to_meters(bp_nsouth_ft)
        bp_ewest_m = feet_to_meters(bp_ewest_ft)
        
        base_point_info = "N: {:.3f}m, E: {:.3f}m, Angle: {:.2f}°".format(
            bp_nsouth_m, bp_ewest_m, math.degrees(angle))
        break

# --- STEP 3: Processing Loop ---
stats = {
    "success": 0,
    "skipped_group": 0,
    "error": 0,
    "total": len(elements)
}
skipped_ids = []

with forms.ProgressBar(title="Calculating... {value}/{max_value}", cancellable=True) as pb:
    t = DB.Transaction(doc, "Update Coordinates")
    t.Start()
    
    t.SetFailureHandlingOptions(t.GetFailureHandlingOptions().SetFailuresPreprocessor(WarningSwallower()))

    count = 0
    for el in elements:
        if pb.cancelled: break
        count += 1
        pb.update_progress(count, stats["total"])

        # Check Group
        if el.GroupId != DB.ElementId.InvalidElementId:
            stats["skipped_group"] += 1
            skipped_ids.append(el.Id.ToString())
            continue

        # Get Location
        x_ft, y_ft = None, None
        try:
            loc = el.Location
            if loc is None:
                bbox = el.get_BoundingBox(None)
                if bbox: x_ft, y_ft = (bbox.Min.X + bbox.Max.X)*0.5, (bbox.Min.Y + bbox.Max.Y)*0.5
            elif isinstance(loc, DB.LocationPoint):
                x_ft, y_ft = loc.Point.X, loc.Point.Y
            elif isinstance(loc, DB.LocationCurve):
                mid = loc.Curve.Evaluate(0.5, True)
                x_ft, y_ft = mid.X, mid.Y
            else:
                bbox = el.get_BoundingBox(None)
                if bbox: x_ft, y_ft = (bbox.Min.X + bbox.Max.X)*0.5, (bbox.Min.Y + bbox.Max.Y)*0.5
                
            if x_ft is None and hasattr(el, 'GetTransform'):
                trf = el.GetTransform()
                if trf: x_ft, y_ft = trf.Origin.X, trf.Origin.Y
        except:
            stats["error"] += 1; continue

        if x_ft is None: stats["error"] += 1; continue

        # Calculate & Set
        try:
            x_m = feet_to_meters(x_ft)
            y_m = feet_to_meters(y_ft)
            north_m, east_m = find_cord(x_m, y_m, angle, bp_ewest_m, bp_nsouth_m)
            
            # ปัดเศษ
            north_m = round(north_m, 3)
            east_m = round(east_m, 3)

            params_set = False
            
            # Loop Set Value (N & E)
            for p_name, val_m in [("N_Coordinate", north_m), ("E_Coordinate", east_m)]:
                param = el.LookupParameter(p_name)
                if param and not param.IsReadOnly:
                    # ตรวจสอบ Parameter Type
                    is_length = False
                    try:
                        if HOST_APP.is_newer_than(2021):
                            spec = param.Definition.GetSpecTypeId()
                            is_length = (spec == DB.SpecTypeId.Length)
                        else:
                            is_length = (param.Definition.ParameterType == DB.ParameterType.Length)
                    except: is_length = True # Default safe

                    if param.StorageType == DB.StorageType.Double:
                        if is_length:
                            # ถ้าเป็น Length ให้แปลงเมตรกลับเป็นฟุต (Revit เก็บเป็นฟุต)
                            param.Set(val_m / 0.3048)
                        else:
                            # ถ้าเป็น Number ให้ใส่ค่าเมตรตรงๆ
                            param.Set(val_m)
                    elif param.StorageType == DB.StorageType.String:
                        param.Set("{:.3f}".format(val_m))
                    
                    params_set = True
            
            if params_set: stats["success"] += 1
            else: stats["error"] += 1

        except:
            stats["error"] += 1

    t.Commit()

# ================================================================
# 3. REPORT UI (Fixed Crash)
# ================================================================

output.print_md("# 📊 Coordinate Update Report")
output.print_md("---\n**Project Base Point Info:** `{}`".format(base_point_info))
output.print_md("**Total Elements Checked:** {}".format(stats["total"]))

table_data = [
    ["✅ Success", str(stats["success"]), "Updated successfully"],
    ["⚠️ Skipped", str(stats["skipped_group"]), "Inside Model Group"],
    ["❌ Error", str(stats["error"]), "Issue"]
]

output.print_table(table_data=table_data, columns=["Status", "Count", "Desc"], formats=["", "", ""])

if skipped_ids:
    output.print_md("---\n### ⚠️ Skipped Elements")
    output.print_md("`IDs: {}`".format(", ".join(skipped_ids[:50]) + ("..." if len(skipped_ids)>50 else "")))

# Summary (FIXED LINE)
output.print_md("---")
if stats["success"] > 0:
    # แก้ไข: แปลงเป็น float ก่อนคำนวณและแสดงผล เพื่อป้องกัน Integer Formatting Error
    success_rate = (float(stats["success"]) / float(stats["total"])) * 100.0
    output.print_md("## 🎉 Success Rate: {:.1f}%".format(success_rate))

if stats["error"] == 0 and stats["skipped_group"] == 0:
    output.print_md("## 🎉 All Clear! 100% Success.")