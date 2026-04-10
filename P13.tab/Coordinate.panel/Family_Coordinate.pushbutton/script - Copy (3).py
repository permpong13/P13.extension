# -*- coding: utf-8 -*-
__title__ = "Update\nCoordinates"
__doc__ = "คำนวณพิกัด (เมตร) ลงในพารามิเตอร์ N/E_Coordinate\n- ข้าม Element ที่อยู่ใน Group\n- รองรับ Revit 2024/2025/2026+"
__author__ = "เพิ่มพงษ์ & Gemini"

from pyrevit import forms, script, DB, HOST_APP
import math
import sys

# ================================================================
# 1. MATHEMATICAL LOGIC
# ================================================================

def feet_to_meters(value):
    return value * 0.3048

def rotate(x, y, theta):
    """หมุนพิกัดตามมุม (Radians)"""
    rotated_x = math.cos(theta) * x + math.sin(theta) * y
    rotated_y = -math.sin(theta) * x + math.cos(theta) * y
    return (rotated_x, rotated_y)

def find_cord(x_m, y_m, theta, bp_ew_m, bp_ns_m):
    """คำนวณพิกัดจริงเทียบกับ Project Base Point"""
    rotated_x, rotated_y = rotate(x_m, y_m, theta)
    actual_east = rotated_x + bp_ew_m
    actual_north = rotated_y + bp_ns_m
    return (actual_north, actual_east)

class WarningSwallower(DB.IFailuresPreprocessor):
    """Class สำหรับจัดการ Warning ของ Revit ไม่ให้เด้งขัดจังหวะ"""
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
    title="เลือกหมวดหมู่ที่ต้องการ (Revit 2026 Compatible)",
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
base_point_info = "Default (0,0)"

base_points = DB.FilteredElementCollector(doc).OfClass(DB.BasePoint).ToElements()
for bp in base_points:
    if not bp.IsShared: # Project Base Point
        angle = bp.get_Parameter(DB.BuiltInParameter.BASEPOINT_ANGLETON_PARAM).AsDouble()
        
        # คำนวณ Offset
        rot_x, rot_y = rotate(bp.Position.X, bp.Position.Y, angle)
        raw_ns = bp.get_Parameter(DB.BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble()
        raw_ew = bp.get_Parameter(DB.BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble()
        
        bp_nsouth_m = feet_to_meters(raw_ns - rot_y)
        bp_ewest_m = feet_to_meters(raw_ew - rot_x)
        
        base_point_info = "N: {:.3f}m, E: {:.3f}m, Angle: {:.2f}°".format(bp_nsouth_m, bp_ewest_m, math.degrees(angle))
        break

# --- STEP 3: Processing Loop ---
stats = {
    "success": 0,
    "skipped_group": 0,
    "error": 0,
    "total": len(elements)
}

skipped_ids = []

with forms.ProgressBar(title="กำลังคำนวณพิกัด... {value}/{max_value}", cancellable=True) as pb:
    t = DB.Transaction(doc, "Update Coordinates")
    t.Start()
    
    # Setup Warning Swallower
    fail_opt = t.GetFailureHandlingOptions()
    fail_opt.SetFailuresPreprocessor(WarningSwallower())
    t.SetFailureHandlingOptions(fail_opt)

    count = 0
    for el in elements:
        if pb.cancelled: break
        count += 1
        pb.update_progress(count, stats["total"])

        # 1. Check Group (ข้ามถ้าอยู่ใน Group)
        if el.GroupId != DB.ElementId.InvalidElementId:
            stats["skipped_group"] += 1
            # --- FIX FOR REVIT 2026: Use .ToString() instead of .IntegerValue ---
            skipped_ids.append(el.Id.ToString()) 
            continue

        # 2. Get Location
        x_ft, y_ft = None, None
        try:
            loc = el.Location
            if isinstance(loc, DB.LocationPoint):
                x_ft, y_ft = loc.Point.X, loc.Point.Y
            elif isinstance(loc, DB.LocationCurve):
                mid = loc.Curve.Evaluate(0.5, True)
                x_ft, y_ft = mid.X, mid.Y
            elif not x_ft:
                # Fallback to Bounding Box
                bbox = el.get_BoundingBox(None)
                if bbox:
                    c = (bbox.Min + bbox.Max) * 0.5
                    x_ft, y_ft = c.X, c.Y
            
            # Method 3: Transform (for Detail Items/Families)
            if x_ft is None and hasattr(el, 'GetTransform'):
                 trf = el.GetTransform()
                 if trf:
                     x_ft, y_ft = trf.Origin.X, trf.Origin.Y

        except:
            stats["error"] += 1
            continue

        if x_ft is None:
            stats["error"] += 1
            continue

        # 3. Calculate
        try:
            north, east = find_cord(feet_to_meters(x_ft), feet_to_meters(y_ft), angle, bp_ewest_m, bp_nsouth_m)

            # 4. Set Parameters
            params_set = False
            for p_name, val in [("N_Coordinate", north), ("E_Coordinate", east)]:
                param = el.LookupParameter(p_name)
                if param and not param.IsReadOnly:
                    # Check Storage Type
                    if param.StorageType == DB.StorageType.Double:
                        # Check Parameter Type (Length vs Number) - Compatible with 2026
                        is_length = False
                        try:
                            # Revit 2022+ uses SpecTypeId
                            if HOST_APP.is_newer_than(2021):
                                spec_type = param.Definition.GetSpecTypeId()
                                is_length = (spec_type == DB.SpecTypeId.Length)
                            else:
                                is_length = (param.Definition.ParameterType == DB.ParameterType.Length)
                        except:
                            pass # Fallback safe
                        
                        if is_length:
                            param.Set(val / 0.3048) # Convert back to feet
                        else:
                            param.Set(val)
                    elif param.StorageType == DB.StorageType.String:
                        param.Set("{:.3f}".format(val))
                    params_set = True
            
            if params_set:
                stats["success"] += 1
            else:
                stats["error"] += 1 

        except Exception as e:
            stats["error"] += 1

    t.Commit()

# ================================================================
# 3. REPORT UI
# ================================================================

output.print_md("# 📊 Coordinate Update Report")
output.print_md("---")
output.print_md("**Project Base Point Info:** `{}`".format(base_point_info))
output.print_md("**Total Elements Checked:** {}".format(stats["total"]))

# Create Status Table
table_data = [
    ["✅ Success", str(stats["success"]), "Updated successfully"],
    ["⚠️ Skipped (Group)", str(stats["skipped_group"]), "Element is inside a Model Group"],
    ["❌ Error/No Param", str(stats["error"]), "No location or Parameter missing"]
]

output.print_table(
    table_data=table_data,
    columns=["Status", "Count", "Description"],
    formats=["", "", ""]
)

# Show Skipped IDs
if skipped_ids:
    output.print_md("---")
    output.print_md("### ⚠️ Skipped Elements (Inside Group)")
    output.print_md("To update these, please ungroup them or edit inside the group editor.")
    
    display_ids = skipped_ids[:50]
    id_str = ", ".join(display_ids)
    if len(skipped_ids) > 50:
        id_str += "... and {} more".format(len(skipped_ids) - 50)
    
    output.print_md("`IDs: {}`".format(id_str))

if stats["error"] == 0 and stats["skipped_group"] == 0:
    output.print_md("## 🎉 All Clear! 100% Success.")