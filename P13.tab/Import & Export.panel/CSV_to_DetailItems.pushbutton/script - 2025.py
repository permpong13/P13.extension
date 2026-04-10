# -*- coding: utf-8 -*-
__title__ = "Detail Items\nfrom CSV"
__author__ = "Tee_เพิ่มพงษ์"
__doc__ = "อ่านพิกัดจากไฟล์ CSV แล้วสร้าง Detail Items ใน Plan View โดยอัตโนมัติ"

import clr
import csv
import math
import os
import sys

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Windows.Forms import *
from System.Drawing import *

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# ===== CONFIGURATION =====
SHOW_DETAILED_LOGS = False
PROGRESS_UPDATE_INTERVAL = 10

def get_project_base_point_coordinates():
    """ดึงค่าพิกัดจาก Project Base Point โดยอัตโนมัติ"""
    try:
        # วิธีที่ 1: ใช้ BasePoint.GetProjectBasePoint()
        project_base_point = BasePoint.GetProjectBasePoint(doc)
        if project_base_point:
            # ดึงค่าพิกัดจากพารามิเตอร์
            east_param = project_base_point.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM)
            north_param = project_base_point.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM)
            elevation_param = project_base_point.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)
            angle_param = project_base_point.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM)
            
            if east_param and east_param.HasValue and north_param and north_param.HasValue:
                # แปลงจากฟุตเป็นเมตร
                base_east = east_param.AsDouble() * 0.3048
                base_north = north_param.AsDouble() * 0.3048
                
                # ดึงมุม True North
                angle_degrees = 0.0
                if angle_param and angle_param.HasValue:
                    angle_degrees = angle_param.AsDouble() * (180.0 / math.pi)
                
                print("✅ ดึงค่าพิกัดจาก Project Base Point สำเร็จ:")
                print("   📍 E: {:.6f} m".format(base_east))
                print("   📍 N: {:.6f} m".format(base_north))
                print("   📐 Angle: {:.6f}°".format(angle_degrees))
                
                return base_east, base_north, angle_degrees
        
        # วิธีที่ 2: ใช้ FilteredElementCollector
        collector = FilteredElementCollector(doc).OfClass(BasePoint)
        for bp in collector:
            if not bp.IsShared:  # Project Base Point
                east_param = bp.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM)
                north_param = bp.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM)
                angle_param = bp.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM)
                
                if east_param and east_param.HasValue and north_param and north_param.HasValue:
                    base_east = east_param.AsDouble() * 0.3048
                    base_north = north_param.AsDouble() * 0.3048
                    
                    angle_degrees = 0.0
                    if angle_param and angle_param.HasValue:
                        angle_degrees = angle_param.AsDouble() * (180.0 / math.pi)
                    
                    print("✅ ดึงค่าพิกัดจาก Project Base Point สำเร็จ (วิธีที่ 2):")
                    print("   📍 E: {:.6f} m".format(base_east))
                    print("   📍 N: {:.6f} m".format(base_north))
                    print("   📐 Angle: {:.6f}°".format(angle_degrees))
                    
                    return base_east, base_north, angle_degrees
        
        # วิธีที่ 3: ใช้ ProjectLocation
        try:
            project_location = doc.ActiveProjectLocation
            if project_location:
                position = project_location.GetProjectPosition(XYZ.Zero)
                if position:
                    base_east = position.EastWest * 0.3048  # ฟุต → เมตร
                    base_north = position.NorthSouth * 0.3048
                    angle_degrees = position.Angle * (180.0 / math.pi)
                    
                    print("✅ ดึงค่าพิกัดจาก Project Location สำเร็จ:")
                    print("   📍 E: {:.6f} m".format(base_east))
                    print("   📍 N: {:.6f} m".format(base_north))
                    print("   📐 Angle: {:.6f}°".format(angle_degrees))
                    
                    return base_east, base_north, angle_degrees
        except:
            pass
                
    except Exception as e:
        print("❌ ไม่สามารถดึงค่าพิกัดจาก Project Base Point: {}".format(e))
    
    print("⚠️ ไม่สามารถดึงค่าพิกัดได้ ใช้ค่าประมาณ")
    return 748053.651, 1449973.325, 263.53

def get_actual_project_base_point_position():
    """ดึงตำแหน่งจริงของ Project Base Point ในแบบจำลอง Revit"""
    try:
        project_base_point = BasePoint.GetProjectBasePoint(doc)
        if project_base_point:
            position = project_base_point.Position
            return position.X, position.Y
        
        collector = FilteredElementCollector(doc).OfClass(BasePoint)
        for bp in collector:
            if not bp.IsShared:
                position = bp.Position
                return position.X, position.Y
                
    except Exception as e:
        print("❌ ไม่สามารถดึงตำแหน่ง Project Base Point: {}".format(e))
    
    return 0.0, 0.0

class BasePointInputForm(Form):
    def __init__(self, default_east, default_north, default_angle):
        self.east_value = default_east
        self.north_value = default_north
        self.angle_value = default_angle
        self.InitializeComponent()
    
    def InitializeComponent(self):
        self.Text = "ตรวจสอบค่า Base Point"
        self.Size = Size(450, 350)
        self.StartPosition = FormStartPosition.CenterScreen
        self.Font = Font("Microsoft Sans Serif", 10)
        
        lbl_title = Label()
        lbl_title.Text = "ค่า Base Point ที่ดึงได้จากโครงการ"
        lbl_title.Location = Point(20, 20)
        lbl_title.Size = Size(400, 25)
        lbl_title.Font = Font(lbl_title.Font, FontStyle.Bold)
        self.Controls.Add(lbl_title)
        
        lbl_desc = Label()
        lbl_desc.Text = "กรุณาตรวจสอบและแก้ไขค่าถ้าจำเป็น:"
        lbl_desc.Location = Point(20, 50)
        lbl_desc.Size = Size(400, 20)
        self.Controls.Add(lbl_desc)
        
        lbl_east = Label()
        lbl_east.Text = "พิกัด East (E):"
        lbl_east.Location = Point(20, 80)
        lbl_east.Size = Size(120, 20)
        self.Controls.Add(lbl_east)
        
        self.txt_east = TextBox()
        self.txt_east.Text = "{:.6f}".format(self.east_value)
        self.txt_east.Location = Point(150, 80)
        self.txt_east.Size = Size(250, 25)
        self.txt_east.Font = Font("Microsoft Sans Serif", 10)
        self.Controls.Add(self.txt_east)
        
        lbl_north = Label()
        lbl_north.Text = "พิกัด North (N):"
        lbl_north.Location = Point(20, 110)
        lbl_north.Size = Size(120, 20)
        self.Controls.Add(lbl_north)
        
        self.txt_north = TextBox()
        self.txt_north.Text = "{:.6f}".format(self.north_value)
        self.txt_north.Location = Point(150, 110)
        self.txt_north.Size = Size(250, 25)
        self.txt_north.Font = Font("Microsoft Sans Serif", 10)
        self.Controls.Add(self.txt_north)
        
        lbl_angle = Label()
        lbl_angle.Text = "มุม True North:"
        lbl_angle.Location = Point(20, 140)
        lbl_angle.Size = Size(120, 20)
        self.Controls.Add(lbl_angle)
        
        self.txt_angle = TextBox()
        self.txt_angle.Text = "{:.6f}".format(self.angle_value)
        self.txt_angle.Location = Point(150, 140)
        self.txt_angle.Size = Size(250, 25)
        self.txt_angle.Font = Font("Microsoft Sans Serif", 10)
        self.Controls.Add(self.txt_angle)
        
        lbl_info = Label()
        lbl_info.Text = "💡 ค่าเหล่านี้ดึงมาจาก Project Base Point ในโครงการ\nสามารถแก้ไขได้หากไม่ถูกต้อง"
        lbl_info.Location = Point(20, 175)
        lbl_info.Size = Size(400, 40)
        lbl_info.ForeColor = Color.Blue
        self.Controls.Add(lbl_info)
        
        self.btn_ok = Button()
        self.btn_ok.Text = "ตกลง"
        self.btn_ok.Location = Point(250, 230)
        self.btn_ok.Size = Size(80, 30)
        self.btn_ok.Click += self.on_ok_click
        self.Controls.Add(self.btn_ok)
        
        self.btn_cancel = Button()
        self.btn_cancel.Text = "ยกเลิก"
        self.btn_cancel.Location = Point(340, 230)
        self.btn_cancel.Size = Size(80, 30)
        self.btn_cancel.Click += self.on_cancel_click
        self.Controls.Add(self.btn_cancel)
    
    def on_ok_click(self, sender, event):
        try:
            self.east_value = float(self.txt_east.Text)
            self.north_value = float(self.txt_north.Text)
            self.angle_value = float(self.txt_angle.Text)
            self.DialogResult = DialogResult.OK
            self.Close()
        except:
            MessageBox.Show("กรุณาป้อนตัวเลขที่ถูกต้อง", "ข้อผิดพลาด")
    
    def on_cancel_click(self, sender, event):
        self.DialogResult = DialogResult.Cancel
        self.Close()

def correct_coordinates_for_base_point_offset(x_feet, y_feet, base_point_x, base_point_y):
    """แก้ไขพิกัดโดยคำนึงถึงการย้ายตำแหน่งของ Project Base Point"""
    corrected_x = x_feet + base_point_x
    corrected_y = y_feet + base_point_y
    return corrected_x, corrected_y

def transform_coordinates_corrected(survey_e, survey_n, base_e, base_n, angle_radians, base_point_offset_x, base_point_offset_y):
    """การแปลงพิกัดที่แก้ไขแล้ว"""
    delta_e = survey_e - base_e
    delta_n = survey_n - base_n
    
    cos_angle = math.cos(angle_radians)
    sin_angle = math.sin(angle_radians)
    
    revit_x = delta_e * cos_angle - delta_n * sin_angle
    revit_y = delta_e * sin_angle + delta_n * cos_angle
    
    revit_x_feet = revit_x * 3.28084
    revit_y_feet = revit_y * 3.28084
    
    corrected_x, corrected_y = correct_coordinates_for_base_point_offset(
        revit_x_feet, revit_y_feet, 
        base_point_offset_x, base_point_offset_y
    )
    
    return corrected_x, corrected_y

def select_csv_file():
    dialog = OpenFileDialog()
    dialog.Title = "เลือกไฟล์ CSV ที่มีพิกัดเสาเข็ม"
    dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    dialog.Multiselect = False
    
    if dialog.ShowDialog() == DialogResult.OK:
        return dialog.FileName
    else:
        print("❌ ไม่ได้เลือกไฟล์")
        return None

def safe_float(value):
    if value is None: return None
    try:
        return float(str(value).strip())
    except:
        return None

def read_csv_file(csv_path):
    pile_data = []
    
    try:
        with open(csv_path, 'rb') as f:
            raw_content = f.read()
            
        encodings = ['utf-8-sig', 'utf-8', 'tis-620', 'cp874', 'latin-1']
        decoded_content = None
        
        for encoding in encodings:
            try:
                decoded_content = raw_content.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
                
        if decoded_content is None:
            decoded_content = raw_content.decode('latin-1')
        
        try:
            from StringIO import StringIO
        except ImportError:
            from io import StringIO
            
        f_io = StringIO(decoded_content)
        reader = csv.DictReader(f_io)
        
        for row in reader:
            pile_no = None
            e = None
            n = None
            
            for key in row.keys():
                if 'pile' in key.lower() or 'no' in key.lower() or 'number' in key.lower():
                    pile_no = str(row.get(key, "")).strip()
                    break
            
            if not pile_no and row.keys():
                first_key = list(row.keys())[0]
                pile_no = str(row.get(first_key, "")).strip()
                        
            if not pile_no: 
                continue
                
            for key in row.keys():
                key_upper = key.upper()
                if key_upper == 'E' or 'east' in key.lower() or 'x' in key_upper:
                    e = safe_float(row.get(key))
                elif key_upper == 'N' or 'north' in key.lower() or 'y' in key_upper:
                    n = safe_float(row.get(key))
            
            if e is None or n is None:
                keys = list(row.keys())
                if len(keys) >= 3:
                    e = safe_float(row.get(keys[1]))
                    n = safe_float(row.get(keys[2]))
                        
            if e is None or n is None:
                continue
                
            pile_data.append({"PileNo": pile_no, "E": e, "N": n})
                
    except Exception as e:
        print("❌ เกิดข้อผิดพลาดในการอ่านไฟล์ CSV: {}".format(e))
        return []
    
    return pile_data

def get_all_detail_family_symbols(doc):
    all_symbols_info = []
    
    try:
        collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
        
        for symbol in collector:
            try:
                if symbol.Category is None:
                    continue
                    
                category_name = ""
                try:
                    category_name = symbol.Category.Name
                except:
                    continue
                
                valid_categories = ['detail', 'annotation', 'symbol', 'mark', 'generic']
                if any(keyword in category_name.lower() for keyword in valid_categories):
                    class SymbolInfo:
                        def __init__(self, symbol, family_name, symbol_name, category_name):
                            self.symbol = symbol
                            self.family_name = family_name
                            self.symbol_name = symbol_name
                            self.category_name = category_name
                    
                    family_name = "Unknown"
                    try:
                        if symbol.Family:
                            family_name = symbol.Family.Name
                    except:
                        pass
                    
                    symbol_name = "Unknown"
                    try:
                        symbol_name = symbol.Name
                    except:
                        pass
                    
                    symbol_info = SymbolInfo(
                        symbol=symbol,
                        family_name=family_name,
                        symbol_name=symbol_name,
                        category_name=category_name
                    )
                    all_symbols_info.append(symbol_info)
                    
            except Exception:
                continue
        
        family_collector = FilteredElementCollector(doc).OfClass(Family)
        
        for family in family_collector:
            try:
                if family.FamilyCategory is None:
                    continue
                    
                category_name = ""
                try:
                    category_name = family.FamilyCategory.Name
                except:
                    continue
                
                valid_categories = ['detail', 'annotation', 'symbol', 'mark', 'generic']
                if any(keyword in category_name.lower() for keyword in valid_categories):
                    symbol_ids = family.GetFamilySymbolIds()
                    
                    for symbol_id in symbol_ids:
                        symbol = doc.GetElement(symbol_id)
                        if symbol is not None:
                            symbol_exists = False
                            for existing_symbol in all_symbols_info:
                                if existing_symbol.symbol.Id == symbol.Id:
                                    symbol_exists = True
                                    break
                            
                            if not symbol_exists:
                                class SymbolInfo:
                                    def __init__(self, symbol, family_name, symbol_name, category_name):
                                        self.symbol = symbol
                                        self.family_name = family_name
                                        self.symbol_name = symbol_name
                                        self.category_name = category_name
                                
                                family_name = "Unknown"
                                try:
                                    family_name = family.Name
                                except:
                                    pass
                                
                                symbol_name = "Unknown"
                                try:
                                    symbol_name = symbol.Name
                                except:
                                    pass
                                
                                symbol_info = SymbolInfo(
                                    symbol=symbol,
                                    family_name=family_name,
                                    symbol_name=symbol_name,
                                    category_name=category_name
                                )
                                all_symbols_info.append(symbol_info)
                            
            except Exception:
                continue
                
    except Exception as e:
        print("❌ ข้อผิดพลาดในการค้นหา Families: {}".format(e))
    
    return all_symbols_info

class FamilySelectionForm(Form):
    def __init__(self, families_data):
        self.families_data = families_data
        self.selected_family_symbol = None
        self.InitializeComponent()
    
    def InitializeComponent(self):
        self.Text = "เลือก Family และ Type สำหรับ Detail Items"
        self.Size = Size(500, 400)
        self.StartPosition = FormStartPosition.CenterScreen
        self.Font = Font("Microsoft Sans Serif", 9)
        
        self.lbl_family = Label()
        self.lbl_family.Text = "เลือก Family:"
        self.lbl_family.Location = Point(20, 20)
        self.lbl_family.Size = Size(100, 20)
        self.Controls.Add(self.lbl_family)
        
        self.cb_family = ComboBox()
        self.cb_family.Location = Point(120, 20)
        self.cb_family.Size = Size(350, 20)
        self.cb_family.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb_family.SelectedIndexChanged += self.on_family_selected
        self.Controls.Add(self.cb_family)
        
        self.lbl_type = Label()
        self.lbl_type.Text = "เลือก Type:"
        self.lbl_type.Location = Point(20, 60)
        self.lbl_type.Size = Size(100, 20)
        self.Controls.Add(self.lbl_type)
        
        self.cb_type = ComboBox()
        self.cb_type.Location = Point(120, 60)
        self.cb_type.Size = Size(350, 20)
        self.cb_type.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb_type.SelectedIndexChanged += self.on_type_selected
        self.Controls.Add(self.cb_type)
        
        self.lbl_info = Label()
        self.lbl_info.Text = "ข้อมูล Family:"
        self.lbl_info.Location = Point(20, 100)
        self.lbl_info.Size = Size(450, 60)
        self.lbl_info.BorderStyle = BorderStyle.FixedSingle
        self.Controls.Add(self.lbl_info)
        
        self.btn_ok = Button()
        self.btn_ok.Text = "ตกลง"
        self.btn_ok.Location = Point(300, 320)
        self.btn_ok.Size = Size(80, 30)
        self.btn_ok.Click += self.on_ok_click
        self.Controls.Add(self.btn_ok)
        
        self.btn_cancel = Button()
        self.btn_cancel.Text = "ยกเลิก"
        self.btn_cancel.Location = Point(390, 320)
        self.btn_cancel.Size = Size(80, 30)
        self.btn_cancel.Click += self.on_cancel_click
        self.Controls.Add(self.btn_cancel)
        
        self.load_families()
    
    def load_families(self):
        family_names = []
        for symbol_info in self.families_data:
            family_name = symbol_info.family_name
            if family_name not in family_names:
                family_names.append(family_name)
        
        family_names.sort()
        
        for name in family_names:
            self.cb_family.Items.Add(name)
        
        if self.cb_family.Items.Count > 0:
            self.cb_family.SelectedIndex = 0
    
    def on_family_selected(self, sender, event):
        if self.cb_family.SelectedIndex < 0:
            return
        
        selected_family = self.cb_family.SelectedItem.ToString()
        self.cb_type.Items.Clear()
        
        for symbol_info in self.families_data:
            if symbol_info.family_name == selected_family:
                display_name = "{} - {}".format(symbol_info.symbol_name, symbol_info.category_name)
                self.cb_type.Items.Add((display_name, symbol_info))
        
        if self.cb_type.Items.Count > 0:
            self.cb_type.SelectedIndex = 0
            
        self.update_info()
    
    def on_type_selected(self, sender, event):
        self.update_info()
    
    def update_info(self):
        if (self.cb_family.SelectedIndex >= 0 and 
            self.cb_type.SelectedIndex >= 0):
            
            family_name = self.cb_family.SelectedItem.ToString()
            display_name, symbol_info = self.cb_type.SelectedItem
            
            info_text = "Family: {}\nType: {}\nCategory: {}\nSymbol Name: {}".format(
                family_name,
                symbol_info.symbol_name,
                symbol_info.category_name,
                symbol_info.symbol_name
            )
            self.lbl_info.Text = info_text
    
    def on_ok_click(self, sender, event):
        if (self.cb_family.SelectedIndex < 0 or 
            self.cb_type.SelectedIndex < 0):
            MessageBox.Show("กรุณาเลือก Family และ Type", "ข้อผิดพลาด")
            return
        
        display_name, symbol_info = self.cb_type.SelectedItem
        self.selected_family_symbol = symbol_info.symbol
        self.DialogResult = DialogResult.OK
        self.Close()
    
    def on_cancel_click(self, sender, event):
        self.DialogResult = DialogResult.Cancel
        self.Close()

def let_user_select_family_symbol():
    all_symbols_info = get_all_detail_family_symbols(doc)
    
    if not all_symbols_info:
        print("❌ ไม่พบ Detail Item Families ในโครงการ")
        return None
    
    form = FamilySelectionForm(all_symbols_info)
    result = form.ShowDialog()
    
    if result == DialogResult.OK and form.selected_family_symbol:
        return form.selected_family_symbol
    else:
        print("❌ ผู้ใช้ยกเลิกการเลือก")
        return None

def set_parameter_value(element, param_name, value):
    try:
        param_names_to_try = [
            param_name,
            param_name.replace(" ", ""),
            param_name.upper(),
            param_name.lower()
        ]
        
        for name in param_names_to_try:
            param = element.LookupParameter(name)
            if param and not param.IsReadOnly:
                if param.StorageType == StorageType.String:
                    param.Set(str(value))
                    return True
                elif param.StorageType == StorageType.Double:
                    try:
                        float_value = float(value)
                        param.Set(float_value)
                        return True
                    except:
                        pass
                elif param.StorageType == StorageType.Integer:
                    try:
                        int_value = int(float(value))
                        param.Set(int_value)
                        return True
                    except:
                        pass
        return False
    except Exception as e:
        return False

def show_task_dialog(title, message):
    try:
        MessageBox.Show(message, title)
    except:
        print("{}: {}".format(title, message))

def create_custom_detail_item(view, location, pile_no, e, n):
    try:
        radius = 0.5
        center = location
        
        normal = XYZ.BasisZ
        plane = Plane.CreateByNormalAndOrigin(normal, center)
        
        arcs = []
        for i in range(4):
            start_angle = i * math.pi / 2
            end_angle = (i + 1) * math.pi / 2
            arc = Arc.Create(plane, radius, start_angle, end_angle)
            arcs.append(arc)
        
        curves = []
        for arc in arcs:
            detail_curve = doc.Create.NewDetailCurve(view, arc)
            if detail_curve:
                curves.append(detail_curve)
                set_parameter_value(detail_curve, "Comments", "Pile: {}".format(pile_no))
        
        return curves[0] if curves else None
        
    except Exception:
        try:
            start_point = location
            end_point = XYZ(location.X + 1.0, location.Y, location.Z)
            line = Line.CreateBound(start_point, end_point)
            detail_line = doc.Create.NewDetailCurve(view, line)
            set_parameter_value(detail_line, "Comments", "Pile: {}".format(pile_no))
            return detail_line
        except Exception:
            return None

def main():
    csv_path = select_csv_file()
    if not csv_path:
        return

    print("🚀 เริ่มสร้าง Detail Items จาก CSV")

    # ดึงค่าพิกัดจาก Project Base Point โดยอัตโนมัติ
    print("\n🔍 กำลังดึงค่าพิกัดจาก Project Base Point...")
    BASE_E, BASE_N, ANGLE_DEGREES = get_project_base_point_coordinates()
    
    # แสดงฟอร์มให้ผู้ใช้ตรวจสอบและแก้ไขค่าถ้าจำเป็น
    form = BasePointInputForm(BASE_E, BASE_N, ANGLE_DEGREES)
    result = form.ShowDialog()
    
    if result == DialogResult.OK:
        BASE_E = form.east_value
        BASE_N = form.north_value
        ANGLE_DEGREES = form.angle_value
        ANGLE_RADIANS = math.radians(ANGLE_DEGREES)
        
        print("📍 ใช้ค่า Base Point: E={:.3f} m, N={:.3f} m".format(BASE_E, BASE_N))
        print("📐 มุม True North: {:.2f}°".format(ANGLE_DEGREES))
    else:
        print("❌ ผู้ใช้ยกเลิก")
        return

    # ดึงตำแหน่ง offset ของ Project Base Point
    BASE_POINT_OFFSET_X, BASE_POINT_OFFSET_Y = get_actual_project_base_point_position()
    print("📍 Base Point Offset: X={:.3f} m, Y={:.3f} m".format(
        BASE_POINT_OFFSET_X * 0.3048, BASE_POINT_OFFSET_Y * 0.3048))

    active_view = doc.ActiveView

    if not isinstance(active_view, ViewPlan):
        show_task_dialog("Error", "❌ กรุณาเปิด Plan View ก่อนรันสคริปต์")
        return

    pile_data = read_csv_file(csv_path)

    if not pile_data:
        show_task_dialog("Error", "❌ ไม่พบข้อมูลในไฟล์ CSV หรือรูปแบบไฟล์ไม่ถูกต้อง")
        return

    print("📖 อ่านข้อมูล {} ตำแหน่งจาก CSV".format(len(pile_data)))

    fam_type = let_user_select_family_symbol()

    use_custom_elements = False
    if not fam_type:
        result = MessageBox.Show(
            "ไม่พบ Detail Item Families ในโครงการ\n\n" +
            "ต้องการสร้างองค์ประกอบแบบง่ายๆ แทนหรือไม่?",
            "ไม่พบ Families",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )
        
        if result == DialogResult.Yes:
            use_custom_elements = True
        else:
            return

    t = Transaction(doc, "สร้าง Detail Items จาก CSV")
    t.Start()
    
    created_count = 0
    failed_count = 0
    failed_items = []
    
    try:
        if fam_type and not fam_type.IsActive:
            fam_type.Activate()
            doc.Regenerate()

        view_elev = active_view.GenLevel.Elevation if active_view.GenLevel else 0.0

        # ทดสอบกับจุดแรก
        if pile_data:
            test_pile = pile_data[0]
            test_x, test_y = transform_coordinates_corrected(
                test_pile["E"], test_pile["N"], 
                BASE_E, BASE_N, 
                ANGLE_RADIANS,
                BASE_POINT_OFFSET_X, 
                BASE_POINT_OFFSET_Y
            )
            print("🧪 ทดสอบจุดแรก: X={:.2f}, Y={:.2f} ฟุต".format(test_x, test_y))

        for i, pile in enumerate(pile_data):
            pile_no = pile["PileNo"]
            e, n = pile["E"], pile["N"]
            
            try:
                x, y = transform_coordinates_corrected(
                    e, n, 
                    BASE_E, BASE_N, 
                    ANGLE_RADIANS,
                    BASE_POINT_OFFSET_X, 
                    BASE_POINT_OFFSET_Y
                )
                
                point = XYZ(x, y, view_elev)
                
                if use_custom_elements:
                    element = create_custom_detail_item(active_view, point, pile_no, e, n)
                else:
                    element = doc.Create.NewFamilyInstance(point, fam_type, active_view)
                    
                    if element:
                        set_parameter_value(element, "Mark", pile_no)
                        set_parameter_value(element, "Comments", pile_no)
                        set_parameter_value(element, "Survey_E", e)
                        set_parameter_value(element, "Survey_N", n)

                if element:
                    created_count += 1
                else:
                    failed_count += 1
                    failed_items.append(pile_no)
                
                if (i + 1) % PROGRESS_UPDATE_INTERVAL == 0:
                    print("   📊 ความคืบหน้า: {}/{}".format(i + 1, len(pile_data)))

            except Exception as ex:
                failed_count += 1
                failed_items.append(pile_no)
                print("❌ ล้มเหลว {}: {}".format(pile_no, str(ex)))

        t.Commit()
        
        print("\n✅ สร้างเสร็จสิ้น: {} ตำแหน่ง, ❌ ล้มเหลว: {} ตำแหน่ง".format(created_count, failed_count))
        
        if failed_items:
            print("📋 รายการที่ล้มเหลว: {}".format(", ".join(failed_items[:10])))
            if len(failed_items) > 10:
                print("   ... และอีก {} รายการ".format(len(failed_items) - 10))

    except Exception as ex:
        t.RollBack()
        print("❌ Transaction ล้มเหลว: {}".format(str(ex)))
        show_task_dialog("Error", "❌ เกิดข้อผิดพลาดในการสร้าง Detail Items")

if __name__ == "__main__":
    main()