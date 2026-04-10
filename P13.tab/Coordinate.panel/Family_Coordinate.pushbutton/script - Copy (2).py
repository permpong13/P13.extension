# -*- coding: utf-8 -*-
__doc__ = "Find coordinates of families (meters) + Ignore warnings + Progress Bar + Auto Create Parameters"
__title__ = "Family \ncoordinates"
__author__ = "เพิ่มพงษ์"

from pyrevit import forms
from pyrevit import DB, HOST_APP
import math
import sys
import time
import threading
import System

from Autodesk.Revit.DB import IFailuresPreprocessor, FailureProcessingResult

# ================================================================
# IGNORE ALL REVIT FAILURES / WARNINGS - แก้ไขเฉพาะ warnings
# ================================================================
class WarningSwallower(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        fails = failuresAccessor.GetFailureMessages()
        for f in fails:
            severity = f.GetSeverity()
            # ลบเฉพาะ warnings, เก็บ errors
            if severity == DB.FailureSeverity.Warning:
                failuresAccessor.DeleteWarning(f)
            else:
                # แสดง errors
                print("Error: {}".format(f.GetDescriptionText()))
        return FailureProcessingResult.Continue

# ================================================================
# COORDINATE CALCULATION - USING ORIGINAL FORMULA
# ================================================================
def rotate(x, y, theta):
    """หมุนพิกัด (x, y) ด้วยมุม theta (radian) - ตามสูตรเดิม"""
    rotated = [
        math.cos(theta) * x + math.sin(theta) * y,
        -math.sin(theta) * x + math.cos(theta) * y
    ]
    return rotated

def find_cord(x, y, theta, bp_ewest, bp_nsouth):
    """คำนวณพิกัดตามสูตรต้นฉบับ - หมุน + อ้างอิง Base Point"""
    # หมุนพิกัด
    rotated = rotate(x, y, theta)
    
    # บวกกับพิกัด Base Point
    result = [
        rotated[0] + bp_ewest,  # East
        rotated[1] + bp_nsouth  # North
    ]
    
    # แปลงหน่วยจาก feet เป็น meters
    result_meters = [coord * 0.3048 for coord in result]
    
    return (result_meters[1], result_meters[0])  # (North, East) ในหน่วยเมตร

# ================================================================
# GET ELEMENT LOCATION
# ================================================================
def get_element_location(element):
    """Get the location of an element using various methods"""
    try:
        # Method 1: Direct location point
        if hasattr(element, 'Location') and element.Location:
            location = element.Location
            if isinstance(location, DB.LocationPoint):
                point = location.Point
                return point.X, point.Y
            elif isinstance(location, DB.LocationCurve):
                curve = location.Curve
                midpoint = curve.Evaluate(0.5, True)
                return midpoint.X, midpoint.Y

        # Method 2: Bounding box
        bbox = element.get_BoundingBox(None)
        if bbox:
            center = (bbox.Min + bbox.Max) * 0.5
            return center.X, center.Y

        # Method 3: For specific element types
        if hasattr(element, 'GetTransform'):
            transform = element.GetTransform()
            if transform:
                origin = transform.Origin
                return origin.X, origin.Y

    except Exception as e:
        print("เกิดข้อผิดพลาดในการหาตำแหน่งองค์ประกอบ {}: {}".format(element.Id, str(e)))

    return None, None

# ================================================================
# GET BASE POINT - USING ORIGINAL FORMULA
# ================================================================
def get_base_point(doc):
    """อ่านค่า Project Base Point ตามสูตรต้นฉบับ"""
    locations = DB.FilteredElementCollector(doc).OfClass(DB.BasePoint).ToElements()
    bp_nsouth = bp_ewest = angle = 0.0
    basepoint_found = False

    for loc in locations:
        try:
            if not loc.IsShared:  # base point
                angle_param = loc.get_Parameter(DB.BuiltInParameter.BASEPOINT_ANGLETON_PARAM)
                if angle_param and angle_param.AsDouble() is not None:
                    angle = angle_param.AsDouble()
                    bp_nsouth_param = loc.get_Parameter(DB.BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM)
                    bp_ewest_param = loc.get_Parameter(DB.BuiltInParameter.BASEPOINT_EASTWEST_PARAM)

                    if bp_nsouth_param and bp_ewest_param:
                        bp_nsouth_val = bp_nsouth_param.AsDouble()
                        bp_ewest_val = bp_ewest_param.AsDouble()

                        # คำนวณพิกัด base point ที่ปรับแล้วตามสูตรต้นฉบับ
                        rotated_pos = rotate(loc.Position.X, loc.Position.Y, angle)
                        bp_nsouth = bp_nsouth_val - rotated_pos[1]
                        bp_ewest = bp_ewest_val - rotated_pos[0]
                        basepoint_found = True
                        
                        print("พบ base point:")
                        print("- Base Point N/S: {} ft".format(round(bp_nsouth_val, 3)))
                        print("- Base Point E/W: {} ft".format(round(bp_ewest_val, 3)))
                        print("- Position X: {} ft".format(round(loc.Position.X, 3)))
                        print("- Position Y: {} ft".format(round(loc.Position.Y, 3)))
                        print("- Rotated X: {} ft".format(round(rotated_pos[0], 3)))
                        print("- Rotated Y: {} ft".format(round(rotated_pos[1], 3)))
                        print("- Adjusted N: {} ft".format(round(bp_nsouth, 3)))
                        print("- Adjusted E: {} ft".format(round(bp_ewest, 3)))
                        print("- Angle: {} องศา".format(round(math.degrees(angle), 1)))
                        break
        except Exception as e:
            print("เกิดข้อผิดพลาดในการรับข้อมูล base point: {}".format(str(e)))

    if not basepoint_found:
        print("ไม่พบข้อมูล base point ที่ถูกต้อง ใช้พิกัดเริ่มต้น")
        return 0.0, 0.0, 0.0

    return angle, bp_ewest, bp_nsouth

# ================================================================
# PARAMETER MANAGEMENT
# ================================================================
def find_existing_parameter(element, param_name):
    """ค้นหาพารามิเตอร์ที่สามารถ Set ค่าได้จริง"""
    try:
        # 1) Search with LookupParameter
        param = element.LookupParameter(param_name)
        if param and not param.IsReadOnly:
            return param

        # 2) Search with GetParameters
        params = element.GetParameters(param_name)
        if params:
            for p in params:
                if p and not p.IsReadOnly:
                    return p

        # 3) Search in parameters
        for param in element.Parameters:
            if param.Definition.Name == param_name and not param.IsReadOnly:
                return param

        return None
    except:
        return None

def try_set_parameter_value(element, param_name, value):
    """พยายาม Set พารามิเตอร์ - แก้ไขให้บันทึกค่าในหน่วย feet"""
    param = find_existing_parameter(element, param_name)
    if not param:
        return False

    try:
        # แปลงค่าเมตรเป็นฟุต (หน่วยภายในของ Revit)
        value_feet = float(value) / 0.3048
        
        st = param.StorageType
        
        if st == DB.StorageType.Double:
            param.Set(value_feet)
            return True
        elif st == DB.StorageType.Integer:
            param.Set(int(round(value_feet)))
            return True
        elif st == DB.StorageType.String:
            param.Set(str(value))  # เก็บเป็น string
            return True
        else:
            return False
    except Exception as e:
        print("⚠ ERROR setting parameter '{}': {}".format(param_name, e))
        return False

# ================================================================
# CREATE PARAMETER IF NEEDED
# ================================================================
def create_parameter_if_needed(doc, param_name, categories_bics):
    """สร้าง Shared Parameter ถ้ายังไม่มี"""
    try:
        # ตรวจสอบว่ามี parameter นี้อยู่แล้วหรือไม่
        binding_map = doc.ParameterBindings
        iterator = binding_map.ForwardIterator()
        while iterator.MoveNext():
            if iterator.Key.Name == param_name:
                print("✓ พบ parameter '{}' อยู่แล้ว".format(param_name))
                return True
                
        print("⚠ ไม่พบ parameter '{}' จะพยายามสร้างใหม่".format(param_name))
        
        # สร้าง Shared Parameter
        app = doc.Application
        shared_param_file = app.OpenSharedParameterFile()
        
        if not shared_param_file:
            print("⚠ ไม่สามารถเปิด Shared Parameter File ได้")
            return False
            
        # สร้าง parameter definition
        group = shared_param_file.Groups.Create("Coordinates")
        
        # กำหนดชนิดข้อมูล
        external_def = app.Create.NewExternalDefinition(
            param_name, 
            DB.SpecTypeId.Length
        )
        
        # สร้าง binding
        categories = doc.Settings.Categories
        cat_set = app.Create.NewCategorySet()
        
        # เพิ่ม category ที่เลือก
        for bic in categories_bics:
            category = categories.get_Item(bic)
            if category:
                cat_set.Insert(category)
        
        # สร้าง binding เป็น Instance Parameter
        binding = app.Create.NewInstanceBinding(cat_set)
            
        # เพิ่ม parameter
        doc.ParameterBindings.Insert(external_def, binding)
        
        print("✓ สร้าง parameter '{}' สำเร็จ".format(param_name))
        return True
        
    except Exception as e:
        print("⚠ ไม่สามารถสร้าง parameter '{}': {}".format(param_name, e))
        return False

# ================================================================
# SIMPLE PARAMETER CHECK
# ================================================================
def check_parameters_exist(doc, categories_bics):
    """ตรวจสอบว่ามีพารามิเตอร์อยู่แล้วหรือไม่"""
    print("กำลังตรวจสอบพารามิเตอร์...")
    
    # หาองค์ประกอบตัวอย่างเพื่อทดสอบ
    test_element = None
    for bic in categories_bics:
        collector = DB.FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
        elements = list(collector)
        if elements:
            test_element = elements[0]
            break
    
    if test_element:
        n_param = find_existing_parameter(test_element, "N_Coordinate")
        e_param = find_existing_parameter(test_element, "E_Coordinate")
        
        if n_param and e_param:
            print("✓ พบพารามิเตอร์ N_Coordinate และ E_Coordinate อยู่แล้ว")
            return True
        else:
            print("⚠ ไม่พบพารามิเตอร์ N_Coordinate และ E_Coordinate")
            
            # สร้างพารามิเตอร์ใหม่
            create_n = create_parameter_if_needed(doc, "N_Coordinate", categories_bics)
            create_e = create_parameter_if_needed(doc, "E_Coordinate", categories_bics)
            
            if create_n and create_e:
                print("✓ สร้างพารามิเตอร์สำเร็จ")
                return True
            else:
                print("⚠ ไม่สามารถสร้างพารามิเตอร์ได้ จะพยายามใช้พารามิเตอร์อื่น")
                
                # แสดงพารามิเตอร์ที่มีอยู่
                all_params = []
                for param in test_element.Parameters:
                    if not param.IsReadOnly and param.StorageType != DB.StorageType.ElementId:
                        all_params.append(param.Definition.Name)
                
                if all_params:
                    print("พารามิเตอร์ที่มีอยู่ในองค์ประกอบ: {}".format(", ".join(all_params[:10])))
                
                return True
    
    return True

# ================================================================
# FIND ALTERNATIVE PARAMETERS
# ================================================================
def find_alternative_parameters(element):
    """ค้นหาพารามิเตอร์ทางเลือกที่สามารถใช้บันทึกค่าพิกัดได้"""
    alternative_params = []
    
    # รายการพารามิเตอร์ที่อาจใช้แทนได้
    possible_params = [
        "N_Coordinate", "E_Coordinate",
        "North", "East", 
        "N", "E",
        "Northing", "Easting",
        "X", "Y",
        "Coordinate_X", "Coordinate_Y",
        "Location_N", "Location_E",
        "Mark", "Comments", "Description",
        "Type Comments", "Assembly Description",
        "Text", "Note", "Remarks"
    ]
    
    for param_name in possible_params:
        param = find_existing_parameter(element, param_name)
        if param and not param.IsReadOnly:
            alternative_params.append(param_name)
    
    return alternative_params

# ================================================================
# VALIDATE COORDINATE CALCULATION
# ================================================================
def validate_calculation(doc, elements, angle, bp_ewest, bp_nsouth):
    """ตรวจสอบการคำนวณพิกัดด้วยตัวอย่าง"""
    print("\n" + "="*50)
    print("VALIDATION - ตรวจสอบการคำนวณพิกัด")
    print("="*50)
    
    if len(elements) > 0:
        sample_element = elements[0]
        element_id = sample_element.Id
        element_category = sample_element.Category.Name if sample_element.Category else "Unknown"
        
        print("องค์ประกอบตัวอย่าง:")
        print("- ID: {}".format(element_id))
        print("- หมวดหมู่: {}".format(element_category))
        
        # อ่านตำแหน่งองค์ประกอบ
        x, y = get_element_location(sample_element)
        if x is not None and y is not None:
            print("- ตำแหน่งใน Revit: X={:.3f} ft, Y={:.3f} ft".format(x, y))
            
            # คำนวณพิกัด
            north, east = find_cord(x, y, angle, bp_ewest, bp_nsouth)
            
            print("\nการคำนวณพิกัด:")
            print("- มุมหมุน (theta): {:.3f} radians ({:.1f} องศา)".format(angle, math.degrees(angle)))
            print("- Base Point E/W: {:.3f} ft".format(bp_ewest))
            print("- Base Point N/S: {:.3f} ft".format(bp_nsouth))
            print("- พิกัด North: {:.3f} m".format(north))
            print("- พิกัด East: {:.3f} m".format(east))
            
            # ตรวจสอบด้วยการคำนวณแบบทีละขั้นตอน
            print("\nการคำนวณแบบทีละขั้นตอน:")
            rotated = rotate(x, y, angle)
            print("1. หลังจากหมุน: X={:.3f}, Y={:.3f}".format(rotated[0], rotated[1]))
            
            east_ft = rotated[0] + bp_ewest
            north_ft = rotated[1] + bp_nsouth
            print("2. หลังจากบวก Base Point: E={:.3f} ft, N={:.3f} ft".format(east_ft, north_ft))
            
            east_m = east_ft * 0.3048
            north_m = north_ft * 0.3048
            print("3. หลังจากแปลงเป็นเมตร: E={:.3f} m, N={:.3f} m".format(east_m, north_m))
            
            return True
    
    print("⚠ ไม่สามารถตรวจสอบการคำนวณได้")
    return False

# ================================================================
# MAIN SCRIPT
# ================================================================
doc = __revit__.ActiveUIDocument.Document
app = doc.Application

if not doc:
    forms.alert("ไม่พบเอกสาร Revit ที่เปิดอยู่", exitscript=True)

# Category list
options_category = {
    'Mechanical Equipment': DB.BuiltInCategory.OST_MechanicalEquipment,
    'Electrical Equipment': DB.BuiltInCategory.OST_ElectricalEquipment,
    'Plumbing Fixtures': DB.BuiltInCategory.OST_PlumbingFixtures,
    'Lighting Fixtures': DB.BuiltInCategory.OST_LightingFixtures,
    'Data Devices': DB.BuiltInCategory.OST_DataDevices,
    'Fire Alarm Devices': DB.BuiltInCategory.OST_FireAlarmDevices,
    'Communication Devices': DB.BuiltInCategory.OST_CommunicationDevices,
    'Air Terminals': DB.BuiltInCategory.OST_DuctTerminal,
    'Pipe Fittings': DB.BuiltInCategory.OST_PipeFitting,
    'Pipe Accessories': DB.BuiltInCategory.OST_PipeAccessory,
    'Duct Fittings': DB.BuiltInCategory.OST_DuctFitting,
    'Doors': DB.BuiltInCategory.OST_Doors,
    'Windows': DB.BuiltInCategory.OST_Windows,
    'Columns': DB.BuiltInCategory.OST_Columns,
    'Structural Columns': DB.BuiltInCategory.OST_StructuralColumns,
    'Structural Foundations': DB.BuiltInCategory.OST_StructuralFoundation,
    'Furniture': DB.BuiltInCategory.OST_Furniture,
    'Generic Models': DB.BuiltInCategory.OST_GenericModel,
    'Detail Components': DB.BuiltInCategory.OST_DetailComponents,
}

# ------------------------------------------------
# เลือก Category ด้วยมือ
# ------------------------------------------------
selected = forms.SelectFromList.show(
    sorted(options_category.keys()),
    multiselect=True,
    title="เลือกหมวดหมู่สำหรับการคำนวณพิกัด"
)
if not selected:
    sys.exit()

categories_bics = [options_category[s] for s in selected]

# ------------------------------------------------
# ตรวจสอบพารามิเตอร์
# ------------------------------------------------
parameters_ready = check_parameters_exist(doc, categories_bics)

if not parameters_ready:
    forms.alert("ไม่สามารถเข้าถึงพารามิเตอร์ได้", exitscript=True)

print("✓ พารามิเตอร์พร้อมใช้งาน")

# ------------------------------------------------
# รวบรวม Element ทั้งหมดจากหมวดหมู่ที่เลือก
# ------------------------------------------------
elements = []
for name in selected:
    bic = options_category[name]
    col = DB.FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
    elems = list(col)
    elements.extend(elems)
    print("พบองค์ประกอบ {0} รายการในหมวดหมู่: {1}".format(len(elems), name))

total = len(elements)
if total == 0:
    forms.alert("ไม่พบองค์ประกอบในหมวดหมู่ที่เลือก", exitscript=True)

print("รวมทั้งหมด {0} องค์ประกอบ".format(total))

if total > 1000:
    res = forms.alert(
        "พบองค์ประกอบทั้งหมด {0} รายการ\nอาจใช้เวลาสักครู่ ต้องการดำเนินการต่อหรือไม่?".format(total),
        ok=False, yes=True, no=True
    )
    if not res:
        sys.exit()

# ------------------------------------------------
# อ่านข้อมูล Base Point
# ------------------------------------------------
print("\nกำลังอ่านข้อมูล Base Point...")
angle, bp_ewest, bp_nsouth = get_base_point(doc)

# ------------------------------------------------
# ตรวจสอบการคำนวณ
# ------------------------------------------------
validation_ok = validate_calculation(doc, elements, angle, bp_ewest, bp_nsouth)

if not validation_ok:
    forms.alert("พบปัญหาในการคำนวณพิกัด", exitscript=True)

# ------------------------------------------------
# ตรวจสอบพารามิเตอร์ที่จะใช้
# ------------------------------------------------
primary_n_param = "N_Coordinate"
primary_e_param = "E_Coordinate"

# ตรวจสอบว่าพารามิเตอร์หลักมีอยู่จริงในองค์ประกอบตัวอย่าง
if len(elements) > 0:
    sample_element = elements[0]
    if not find_existing_parameter(sample_element, primary_n_param) or not find_existing_parameter(sample_element, primary_e_param):
        alt_params = find_alternative_parameters(sample_element)
        if len(alt_params) >= 2:
            # ให้ผู้ใช้เลือกพารามิเตอร์
            param_options = []
            for i, param in enumerate(alt_params[:4]):  # ใช้ 4 พารามิเตอร์แรก
                if i % 2 == 0:
                    param_options.append("{} → North Coordinate".format(param))
                else:
                    param_options.append("{} → East Coordinate".format(param))
            
            if len(param_options) >= 2:
                result = forms.SelectFromList.show(
                    param_options,
                    title="เลือกพารามิเตอร์สำหรับบันทึกค่าพิกัด\n(เลือก 2 พารามิเตอร์สำหรับ North และ East)",
                    multiselect=True
                )
                
                if result and len(result) == 2:
                    primary_n_param = result[0].split(" → ")[0]
                    primary_e_param = result[1].split(" → ")[0]
                    print("จะใช้พารามิเตอร์: {} สำหรับ North, {} สำหรับ East".format(primary_n_param, primary_e_param))
                else:
                    forms.alert("ต้องเลือกพารามิเตอร์ 2 ตัวสำหรับบันทึกค่าพิกัด", exitscript=True)

# ------------------------------------------------
# Main Transaction + Progress Bar
# ------------------------------------------------
start_time = time.time()
success_count = 0
error_count = 0
no_parameters_count = 0

with forms.ProgressBar(title="Calculating Coordinates... ({value} of {max_value})", cancellable=True) as pb:
    # ใช้ Transaction แบบ Manual
    transaction = DB.Transaction(doc, "Assign Coordinates")
    transaction.Start()

    try:
        # Ignore warnings
        failure_options = transaction.GetFailureHandlingOptions()
        failure_options.SetFailuresPreprocessor(WarningSwallower())
        transaction.SetFailureHandlingOptions(failure_options)

        for i, element in enumerate(elements):
            if pb.cancelled:
                print("ผู้ใช้ยกเลิกการทำงาน")
                transaction.RollBack()
                sys.exit(0)

            pb.update_progress(i + 1, total)

            try:
                x, y = get_element_location(element)
                if x is None or y is None:
                    error_count += 1
                    continue

                # คำนวณพิกัดตามสูตรต้นฉบับ
                north, east = find_cord(x, y, angle, bp_ewest, bp_nsouth)

                # พยายามบันทึกค่าลงพารามิเตอร์หลัก
                n_success = try_set_parameter_value(element, primary_n_param, round(north, 3))
                e_success = try_set_parameter_value(element, primary_e_param, round(east, 3))

                if n_success and e_success:
                    success_count += 1
                else:
                    # ถ้าไม่สำเร็จ ให้พยายามหาพารามิเตอร์ทางเลือก
                    alt_params = find_alternative_parameters(element)
                    if len(alt_params) >= 2:
                        # พยายามใช้พารามิเตอร์สองตัวแรกที่หาได้
                        alt_n_success = try_set_parameter_value(element, alt_params[0], "N:" + str(round(north, 3)))
                        alt_e_success = try_set_parameter_value(element, alt_params[1], "E:" + str(round(east, 3)))
                        
                        if alt_n_success or alt_e_success:
                            success_count += 1
                        else:
                            no_parameters_count += 1
                    else:
                        no_parameters_count += 1

            except Exception as e:
                error_count += 1
                if error_count <= 5:  # แสดงเพียง 5 ข้อผิดพลาดแรก
                    print("Error processing element {0}: {1}".format(element.Id, e))

        # Commit transaction
        transaction.Commit()
        print("✓ Transaction committed สำเร็จ")
        
    except Exception as e:
        print("✗ Transaction failed: {}".format(e))
        transaction.RollBack()
        raise

end_time = time.time()
elapsed = end_time - start_time

# ------------------------------------------------
# ตรวจสอบว่าค่าถูกบันทึกจริงหรือไม่
# ------------------------------------------------
if len(elements) > 0:
    sample_element = elements[0]
    n_param = find_existing_parameter(sample_element, primary_n_param)
    e_param = find_existing_parameter(sample_element, primary_e_param)
    
    print("\nตรวจสอบการบันทึกค่า:")
    if n_param and n_param.HasValue:
        print("✓ ค่า North หลังจาก Commit: {}".format(n_param.AsDouble()))
    else:
        print("⚠ ไม่พบค่า North หลัง Commit")
        
    if e_param and e_param.HasValue:
        print("✓ ค่า East หลังจาก Commit: {}".format(e_param.AsDouble()))
    else:
        print("⚠ ไม่พบค่า East หลัง Commit")

# ------------------------------------------------
# สรุปผล
# ------------------------------------------------
result_message = (
    "Completed!\n\n"
    "Total elements: {0}\n"
    "Success: {1}\n"
    "No parameters available: {2}\n"
    "Errors: {3}\n"
    "Processing time: {4} วินาที"
).format(total, success_count, no_parameters_count, error_count, round(elapsed, 1))

forms.alert(result_message, title="Element Coordinates - Completed")

print("\n" + "="*50)
print("FINAL SUMMARY")
print("="*50)
print("Base Point Information:")
print("- Angle: {0} degrees".format(round(math.degrees(angle), 1)))
print("- E/W Offset: {0} ft".format(round(bp_ewest, 3)))
print("- N/S Offset: {0} ft".format(round(bp_nsouth, 3)))
print("Processing Results:")
print("- Total elements: {0}".format(total))
print("- Success: {0}".format(success_count))
print("- No parameters available: {0}".format(no_parameters_count))
print("- Errors: {0}".format(error_count))
print("Parameters used:")
print("- North coordinate: {0}".format(primary_n_param))
print("- East coordinate: {0}".format(primary_e_param))

if no_parameters_count > 0:
    print("\n⚠ คำแนะนำ:")
    print("- โปรดสร้างพารามิเตอร์ N_Coordinate และ E_Coordinate ด้วยมือ")
    print("- หรือตรวจสอบว่าพารามิเตอร์เหล่านี้ถูกผูกกับหมวดหมู่ที่เลือกแล้ว")