# -*- coding: utf-8 -*-
__title__ = "Family Coord\nfrom CSV"
__author__ = "Tee_เพิ่มพงษ์"
__doc__ = "อ่านพิกัดจากไฟล์ CSV แล้วสร้าง Structural Columns หรือ Structural Foundations ใน Plan View โดยอัตโนมัติ"

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
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI import *
from System.Windows.Forms import *
from System.Drawing import *

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# ===== CONFIGURATION =====
SHOW_DETAILED_LOGS = True
PROGRESS_UPDATE_INTERVAL = 50

# ===== LEVEL SELECTION FORM =====
class LevelSelectionForm(Form):
    def __init__(self, levels, is_base_level=True):
        self.levels = levels
        self.selected_level = None
        self.is_base_level = is_base_level
        self.InitializeComponent()
    
    def InitializeComponent(self):
        if self.is_base_level:
            self.Text = "เลือก Level ฐาน (ระดับ 0.00) สำหรับองค์ประกอบโครงสร้าง"
        else:
            self.Text = "เลือก Level สำหรับองค์ประกอบโครงสร้าง"
            
        self.Size = Size(500, 200)
        self.StartPosition = FormStartPosition.CenterScreen
        self.Font = Font("Microsoft Sans Serif", 9)
        
        # Label สำหรับ Level
        self.lbl_level = Label()
        self.lbl_level.Text = "เลือก Level:"
        self.lbl_level.Location = Point(20, 20)
        self.lbl_level.Size = Size(100, 20)
        self.Controls.Add(self.lbl_level)
        
        # ComboBox สำหรับ Level
        self.cb_level = ComboBox()
        self.cb_level.Location = Point(120, 20)
        self.cb_level.Size = Size(350, 20)
        self.cb_level.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb_level.SelectedIndexChanged += self.on_level_selected
        self.Controls.Add(self.cb_level)
        
        # Label แสดงข้อมูล Level
        self.lbl_info = Label()
        self.lbl_info.Text = "ข้อมูล Level:"
        self.lbl_info.Location = Point(20, 60)
        self.lbl_info.Size = Size(450, 40)
        self.lbl_info.BorderStyle = BorderStyle.FixedSingle
        self.Controls.Add(self.lbl_info)
        
        # ปุ่ม OK
        self.btn_ok = Button()
        self.btn_ok.Text = "ตกลง"
        self.btn_ok.Location = Point(300, 120)
        self.btn_ok.Size = Size(80, 30)
        self.btn_ok.Click += self.on_ok_click
        self.Controls.Add(self.btn_ok)
        
        # ปุ่ม Cancel
        self.btn_cancel = Button()
        self.btn_cancel.Text = "ยกเลิก"
        self.btn_cancel.Location = Point(390, 120)
        self.btn_cancel.Size = Size(80, 30)
        self.btn_cancel.Click += self.on_cancel_click
        self.Controls.Add(self.btn_cancel)
        
        # โหลดข้อมูล Level ลงใน ComboBox
        self.load_levels()
    
    def load_levels(self):
        """โหลด Level ลงใน ComboBox"""
        level_names = []
        for level in self.levels:
            if hasattr(level, 'Name'):
                level_names.append(level.Name)
        
        level_names.sort()
        
        for name in level_names:
            self.cb_level.Items.Add(name)
        
        if self.cb_level.Items.Count > 0:
            self.cb_level.SelectedIndex = 0
    
    def on_level_selected(self, sender, event):
        """เมื่อเลือก Level ให้อัปเดตข้อมูล"""
        self.update_info()
    
    def update_info(self):
        """อัปเดตข้อมูลที่แสดง"""
        if self.cb_level.SelectedIndex >= 0:
            selected_level_name = self.cb_level.SelectedItem.ToString()
            
            # ค้นหา Level object ที่ตรงกับชื่อ
            for level in self.levels:
                if hasattr(level, 'Name') and level.Name == selected_level_name:
                    elevation_feet = level.Elevation
                    elevation_meters = elevation_feet * 0.3048
                    
                    if self.is_base_level:
                        info_text = "Level ฐาน: {}\nElevation: {:.3f} ft ({:.3f} m)\n\nLevel นี้จะใช้เป็นระดับอ้างอิง 0.00\nค่า Pile Cut Off จาก CSV จะถูกใช้เป็น Height Offset จาก Level นี้".format(
                            selected_level_name,
                            elevation_feet,
                            elevation_meters
                        )
                    else:
                        info_text = "Level: {}\nElevation: {:.3f} ft ({:.3f} m)".format(
                            selected_level_name,
                            elevation_feet,
                            elevation_meters
                        )
                    self.lbl_info.Text = info_text
                    break
    
    def on_ok_click(self, sender, event):
        """เมื่อกดปุ่ม OK"""
        if self.cb_level.SelectedIndex < 0:
            MessageBox.Show("กรุณาเลือก Level", "ข้อผิดพลาด")
            return
        
        selected_level_name = self.cb_level.SelectedItem.ToString()
        
        # ค้นหา Level object ที่ตรงกับชื่อ
        for level in self.levels:
            if hasattr(level, 'Name') and level.Name == selected_level_name:
                self.selected_level = level
                break
        
        if self.selected_level:
            self.DialogResult = DialogResult.OK
            self.Close()
        else:
            MessageBox.Show("ไม่พบ Level ที่เลือก", "ข้อผิดพลาด")
    
    def on_cancel_click(self, sender, event):
        """เมื่อกดปุ่ม Cancel"""
        self.DialogResult = DialogResult.Cancel
        self.Close()

# ===== FAMILY SELECTION FORM =====
class FamilySelectionForm(Form):
    def __init__(self, families_data):
        self.families_data = families_data
        self.selected_family_symbol = None
        self.InitializeComponent()
    
    def InitializeComponent(self):
        self.Text = "เลือก Structural Family และ Type"
        self.Size = Size(500, 400)
        self.StartPosition = FormStartPosition.CenterScreen
        self.Font = Font("Microsoft Sans Serif", 9)
        
        # Label สำหรับ Family
        self.lbl_family = Label()
        self.lbl_family.Text = "เลือก Family:"
        self.lbl_family.Location = Point(20, 20)
        self.lbl_family.Size = Size(100, 20)
        self.Controls.Add(self.lbl_family)
        
        # ComboBox สำหรับ Family
        self.cb_family = ComboBox()
        self.cb_family.Location = Point(120, 20)
        self.cb_family.Size = Size(350, 20)
        self.cb_family.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb_family.SelectedIndexChanged += self.on_family_selected
        self.Controls.Add(self.cb_family)
        
        # Label สำหรับ Type
        self.lbl_type = Label()
        self.lbl_type.Text = "เลือก Type:"
        self.lbl_type.Location = Point(20, 60)
        self.lbl_type.Size = Size(100, 20)
        self.Controls.Add(self.lbl_type)
        
        # ComboBox สำหรับ Type
        self.cb_type = ComboBox()
        self.cb_type.Location = Point(120, 60)
        self.cb_type.Size = Size(350, 20)
        self.cb_type.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb_type.SelectedIndexChanged += self.on_type_selected
        self.Controls.Add(self.cb_type)
        
        # Label แสดงข้อมูล
        self.lbl_info = Label()
        self.lbl_info.Text = "ข้อมูล Family:"
        self.lbl_info.Location = Point(20, 100)
        self.lbl_info.Size = Size(450, 60)
        self.lbl_info.BorderStyle = BorderStyle.FixedSingle
        self.Controls.Add(self.lbl_info)
        
        # ปุ่ม OK
        self.btn_ok = Button()
        self.btn_ok.Text = "ตกลง"
        self.btn_ok.Location = Point(300, 320)
        self.btn_ok.Size = Size(80, 30)
        self.btn_ok.Click += self.on_ok_click
        self.Controls.Add(self.btn_ok)
        
        # ปุ่ม Cancel
        self.btn_cancel = Button()
        self.btn_cancel.Text = "ยกเลิก"
        self.btn_cancel.Location = Point(390, 320)
        self.btn_cancel.Size = Size(80, 30)
        self.btn_cancel.Click += self.on_cancel_click
        self.Controls.Add(self.btn_cancel)
        
        # โหลดข้อมูล Family ลงใน ComboBox
        self.load_families()
    
    def load_families(self):
        """โหลด Family ลงใน ComboBox"""
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
        """เมื่อเลือก Family ให้โหลด Type ที่เกี่ยวข้อง"""
        if self.cb_family.SelectedIndex < 0:
            return
        
        selected_family = self.cb_family.SelectedItem.ToString()
        self.cb_type.Items.Clear()
        
        for symbol_info in self.families_data:
            if symbol_info.family_name == selected_family:
                # แสดงชื่อ Type ที่แท้จริงจาก Revit
                display_name = "{} - {}".format(symbol_info.symbol_name, symbol_info.category_name)
                self.cb_type.Items.Add((display_name, symbol_info))
        
        if self.cb_type.Items.Count > 0:
            self.cb_type.SelectedIndex = 0
            
        self.update_info()
    
    def on_type_selected(self, sender, event):
        """เมื่อเลือก Type ให้อัปเดตข้อมูล"""
        self.update_info()
    
    def update_info(self):
        """อัปเดตข้อมูลที่แสดง"""
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
        """เมื่อกดปุ่ม OK"""
        if (self.cb_family.SelectedIndex < 0 or 
            self.cb_type.SelectedIndex < 0):
            MessageBox.Show("กรุณาเลือก Family และ Type", "ข้อผิดพลาด")
            return
        
        display_name, symbol_info = self.cb_type.SelectedItem
        self.selected_family_symbol = symbol_info.symbol
        self.DialogResult = DialogResult.OK
        self.Close()
    
    def on_cancel_click(self, sender, event):
        """เมื่อกดปุ่ม Cancel"""
        self.DialogResult = DialogResult.Cancel
        self.Close()

def get_all_levels(doc):
    """ดึง Level ทั้งหมดจาก Revit Model"""
    levels = []
    collector = FilteredElementCollector(doc).OfClass(Level)
    for level in collector:
        if level is not None:
            levels.append(level)
    
    # เรียงลำดับ Level ตามค่า Elevation
    levels.sort(key=lambda x: x.Elevation)
    return levels

def let_user_select_base_level():
    """ให้ผู้ใช้เลือก Level ฐาน (ระดับ 0.00) จากฟอร์ม"""
    print("\n🔍 กำลังค้นหา Levels ในโครงการ...")
    
    levels = get_all_levels(doc)
    
    if not levels:
        print("❌ ไม่พบ Levels ในโครงการ")
        return None
    
    print("📋 พบ {} Levels ในโครงการ".format(len(levels)))
    
    form = LevelSelectionForm(levels, is_base_level=True)
    result = form.ShowDialog()
    
    if result == DialogResult.OK and form.selected_level:
        selected_level = form.selected_level
        elevation_feet = selected_level.Elevation
        elevation_meters = elevation_feet * 0.3048
        
        print("✅ ผู้ใช้เลือก Level ฐาน: {} (Elevation: {:.3f} ft, {:.3f} m)".format(
            selected_level.Name, elevation_feet, elevation_meters))
        return selected_level
    else:
        print("❌ ผู้ใช้ยกเลิกการเลือก Level ฐาน")
        return None

def get_all_families_by_category(doc, category_names):
    """ค้นหา Families ตามหมวดหมู่ที่กำหนด"""
    families = []
    collector = FilteredElementCollector(doc).OfClass(Family)
    for family in collector:
        try:
            if (family is not None and 
                hasattr(family, 'FamilyCategory') and 
                family.FamilyCategory is not None and
                hasattr(family.FamilyCategory, 'Name')):
                
                category_name = family.FamilyCategory.Name
                if category_name in category_names:
                    families.append(family)
        except:
            continue
    return families

def get_family_symbols_info(family):
    """ดึงข้อมูล Family Symbols ทั้งหมดจาก Family"""
    symbols_info = []
    try:
        symbol_ids = family.GetFamilySymbolIds()
        for symbol_id in symbol_ids:
            symbol = doc.GetElement(symbol_id)
            if symbol is not None:
                class SymbolInfo:
                    def __init__(self, symbol, family_name, symbol_name, category_name):
                        self.symbol = symbol
                        self.family_name = family_name
                        self.symbol_name = symbol_name
                        self.category_name = category_name
                
                category_name = "Unknown"
                if (hasattr(symbol, 'Category') and 
                    symbol.Category is not None and
                    hasattr(symbol.Category, 'Name')):
                    category_name = symbol.Category.Name
                
                symbol_info = SymbolInfo(
                    symbol=symbol,
                    family_name=family.Name if hasattr(family, 'Name') else "Unknown",
                    symbol_name=symbol.Name if hasattr(symbol, 'Name') else "Unknown",
                    category_name=category_name
                )
                symbols_info.append(symbol_info)
    except:
        pass
    return symbols_info

def let_user_select_family_symbol():
    """ให้ผู้ใช้เลือก Family และ Type"""
    print("\n🔍 กำลังค้นหา Structural Families ในโครงการ...")
    
    # ค้นหาเฉพาะหมวดหมู่ที่เกี่ยวข้องกับโครงสร้าง
    relevant_categories = [
        "Structural Foundations", 
        "Structural Columns"
    ]
    
    families = get_all_families_by_category(doc, relevant_categories)
    
    if not families:
        print("❌ ไม่พบ Structural Families ในโครงการ")
        return None
    
    all_symbols_info = []
    for family in families:
        symbols_info = get_family_symbols_info(family)
        all_symbols_info.extend(symbols_info)
    
    if not all_symbols_info:
        print("❌ ไม่พบ Structural Family Symbols ในโครงการ")
        return None
    
    print("📋 พบ {} Structural Families และ {} Symbols".format(
        len(families), len(all_symbols_info)))
    
    form = FamilySelectionForm(all_symbols_info)
    result = form.ShowDialog()
    
    if result == DialogResult.OK and form.selected_family_symbol:
        selected_symbol = form.selected_family_symbol
        family_name = selected_symbol.Family.Name if hasattr(selected_symbol.Family, 'Name') else "Unknown"
        symbol_name = selected_symbol.Name if hasattr(selected_symbol, 'Name') else "Unknown"
        category_name = selected_symbol.Category.Name if hasattr(selected_symbol.Category, 'Name') else "Unknown"
        print("✅ ผู้ใช้เลือก: {} - {} ({})".format(family_name, symbol_name, category_name))
        return selected_symbol
    else:
        print("❌ ผู้ใช้ยกเลิกการเลือก")
        return None

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
    dialog.Title = "เลือกไฟล์ CSV ที่มีพิกัด"
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
    data = []
    has_cutoff_data = False
    
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
            element_no = None
            e = None
            n = None
            pile_cut_off = None
            
            for key in row.keys():
                if 'pile' in key.lower() or 'no' in key.lower() or 'number' in key.lower() or 'mark' in key.lower():
                    element_no = str(row.get(key, "")).strip()
                    break
            
            if not element_no and row.keys():
                first_key = list(row.keys())[0]
                element_no = str(row.get(first_key, "")).strip()
                        
            if not element_no: 
                continue
                
            for key in row.keys():
                key_upper = key.upper()
                if key_upper == 'E' or 'east' in key.lower() or 'x' in key_upper:
                    e = safe_float(row.get(key))
                elif key_upper == 'N' or 'north' in key.lower() or 'y' in key_upper:
                    n = safe_float(row.get(key))
                elif 'cut' in key.lower() or 'top' in key.lower() or 'elev' in key.lower():
                    pile_cut_off = safe_float(row.get(key))
            
            if e is None or n is None:
                keys = list(row.keys())
                if len(keys) >= 3:
                    e = safe_float(row.get(keys[1]))
                    n = safe_float(row.get(keys[2]))
                    if len(keys) >= 4:
                        pile_cut_off = safe_float(row.get(keys[3]))
                        
            if e is None or n is None:
                continue
            
            # ตรวจสอบว่ามีข้อมูล Pile Cut Off หรือไม่
            if pile_cut_off is not None:
                has_cutoff_data = True
                
            data.append({
                "ElementNo": element_no, 
                "E": e, 
                "N": n, 
                "PileCutOff": pile_cut_off
            })
                
    except Exception as e:
        print("❌ เกิดข้อผิดพลาดในการอ่านไฟล์ CSV: {}".format(e))
        return [], False
    
    return data, has_cutoff_data

def set_parameter_value(element, param_name, value):
    """ตั้งค่าพารามิเตอร์ให้กับองค์ประกอบ"""
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
    """แสดงข้อความแจ้งเตือน"""
    try:
        MessageBox.Show(message, title)
    except:
        print("{}: {}".format(title, message))

def create_structural_element_with_cutoff(view, location, family_symbol, element_no, e, n, base_level, pile_cut_off):
    """สร้าง Structural Element และตั้งค่า Height Offset จาก Level ฐานตามค่า Pile Cut Off"""
    try:
        # ตรวจสอบประเภทของ Family
        category_name = family_symbol.Category.Name.lower()
        
        # เลือก StructuralType ตามประเภทขององค์ประกอบ
        if 'column' in category_name:
            structural_type = StructuralType.Column
        elif 'foundation' in category_name:
            structural_type = StructuralType.Footing
        else:
            structural_type = StructuralType.NonStructural
        
        # สร้าง Structural Element
        structural_element = doc.Create.NewFamilyInstance(location, family_symbol, base_level, structural_type)
        
        if structural_element:
            # ตั้งค่าพารามิเตอร์พื้นฐาน
            set_parameter_value(structural_element, "Mark", element_no)
            set_parameter_value(structural_element, "Comments", "Element: {}".format(element_no))
            set_parameter_value(structural_element, "Survey_E", e)
            set_parameter_value(structural_element, "Survey_N", n)
            set_parameter_value(structural_element, "Element Number", element_no)
            
            # ตั้งค่า Height Offset จาก Level ฐานตามค่า Pile Cut Off
            if pile_cut_off is not None:
                # คำนวณค่า Offset จาก Level ฐาน
                # สมมติว่า Pile Cut Off เป็นค่าระดับสัมบูรณ์ (Absolute Elevation)
                # และ Level ฐานมีค่าระดับ base_level.Elevation (ฟุต)
                
                # แปลงค่า Pile Cut Off จากเมตรเป็นฟุต
                pile_cut_off_feet = pile_cut_off * 3.28084
                
                # คำนวณ Height Offset = Pile Cut Off - Base Level Elevation
                base_level_elevation = base_level.Elevation
                height_offset = pile_cut_off_feet - base_level_elevation
                
                print("📐 คำนวณ Height Offset สำหรับ {}: {:.3f} ม. (Pile Cut Off) - {:.3f} ฟุต (Base Level) = {:.3f} ฟุต".format(
                    element_no, pile_cut_off, base_level_elevation, height_offset))
                
                # พยายามตั้งค่า Height Offset ในพารามิเตอร์ต่างๆ
                offset_set = False
                
                # ลองพารามิเตอร์มาตรฐานของ Revit
                offset_params = [
                    "Height Offset From Level",
                    "Offset From Level", 
                    "Base Offset",
                    "Offset",
                    "Top Offset",
                    "Start Level Offset",
                    "Base Level Offset"
                ]
                
                for param_name in offset_params:
                    if set_parameter_value(structural_element, param_name, height_offset):
                        offset_set = True
                        print("✅ ตั้งค่า {} เป็น {:.3f} ฟุต สำหรับ {}".format(param_name, height_offset, element_no))
                        break
                
                if not offset_set:
                    # ถ้าไม่พบพารามิเตอร์มาตรฐาน ให้ลองใช้ BuiltInParameter
                    try:
                        # พารามิเตอร์ BuiltIn สำหรับ Structural Foundations
                        if 'foundation' in category_name:
                            param = structural_element.get_Parameter(BuiltInParameter.STRUCTURAL_BOTTOM_LEVEL_OFFSET_PARAM)
                            if param and not param.IsReadOnly:
                                param.Set(height_offset)
                                offset_set = True
                                print("✅ ตั้งค่า STRUCTURAL_BOTTOM_LEVEL_OFFSET_PARAM เป็น {:.3f} ฟุต สำหรับ {}".format(height_offset, element_no))
                    except:
                        pass
                
                # บันทึกค่า Pile Cut Off ดิบไว้ในพารามิเตอร์ (หน่วยเมตร)
                set_parameter_value(structural_element, "Pile Cut Off", pile_cut_off)
                set_parameter_value(structural_element, "Cut Off Level", pile_cut_off)
                
                if not offset_set:
                    print("⚠️ ไม่สามารถตั้งค่า Height Offset สำหรับ {} (ค่า Cut Off: {:.3f} ม.)".format(element_no, pile_cut_off))
            
            # พารามิเตอร์เพิ่มเติมสำหรับ Structural Elements
            set_parameter_value(structural_element, "Reference", element_no)
            set_parameter_value(structural_element, "Type Mark", element_no)
            
            print("✅ สร้าง {} สำเร็จสำหรับ {}".format(family_symbol.Category.Name, element_no))
            return structural_element
        else:
            return None
            
    except Exception as ex:
        print("❌ เกิดข้อผิดพลาดในการสร้าง Structural Element: {}".format(ex))
        return None

def main():
    """ฟังก์ชันหลัก"""
    csv_path = select_csv_file()
    if not csv_path:
        return

    print("🚀 เริ่มสร้าง Structural Elements จาก CSV")

    # ดึงค่าพิกัดจาก Project Base Point โดยอัตโนมัติ
    print("\n🔍 กำลังดึงค่าพิกัดจาก Project Base Point...")
    BASE_E, BASE_N, ANGLE_DEGREES = get_project_base_point_coordinates()
    
    ANGLE_RADIANS = math.radians(ANGLE_DEGREES)
    
    print("📍 ใช้ค่า Base Point: E={:.3f} m, N={:.3f} m".format(BASE_E, BASE_N))
    print("📐 มุม True North: {:.2f}°".format(ANGLE_DEGREES))

    # ดึงตำแหน่ง offset ของ Project Base Point
    BASE_POINT_OFFSET_X, BASE_POINT_OFFSET_Y = get_actual_project_base_point_position()
    print("📍 Base Point Offset: X={:.3f} m, Y={:.3f} m".format(
        BASE_POINT_OFFSET_X * 0.3048, BASE_POINT_OFFSET_Y * 0.3048))

    active_view = doc.ActiveView

    if not isinstance(active_view, ViewPlan):
        show_task_dialog("Error", "❌ กรุณาเปิด Plan View ก่อนรันสคริปต์")
        return

    # ให้ผู้ใช้เลือก Level ฐาน (ระดับ 0.00)
    print("\n📋 ขั้นตอนที่ 1: เลือก Level ฐาน (ระดับ 0.00)")
    base_level = let_user_select_base_level()
    if not base_level:
        show_task_dialog("Error", "❌ ไม่ได้เลือก Level ฐาน")
        return

    # อ่านข้อมูลจาก CSV
    data, has_cutoff_data = read_csv_file(csv_path)

    if not data:
        show_task_dialog("Error", "❌ ไม่พบข้อมูลในไฟล์ CSV หรือรูปแบบไฟล์ไม่ถูกต้อง")
        return

    print("📖 อ่านข้อมูล {} ตำแหน่งจาก CSV".format(len(data)))
    if has_cutoff_data:
        print("📏 พบข้อมูล Pile Cut Off ในไฟล์ CSV - จะใช้ค่าเหล่านี้เพื่อกำหนด Height Offset จาก Level ฐาน")
    else:
        print("⚠️ ไม่พบข้อมูล Pile Cut Off ในไฟล์ CSV - องค์ประกอบจะถูกสร้างที่ระดับ Level ฐาน")

    # ให้ผู้ใช้เลือก Structural Family Symbol
    print("\n📋 ขั้นตอนที่ 2: เลือก Family และ Type")
    family_symbol = let_user_select_family_symbol()

    if not family_symbol:
        show_task_dialog("Error", "❌ ไม่ได้เลือก Structural Family Symbol")
        return

    t = Transaction(doc, "สร้าง Structural Elements จาก CSV ด้วย Pile Cut Off")
    t.Start()
    
    created_count = 0
    failed_count = 0
    failed_items = []
    elements_with_cutoff = 0
    
    try:
        # เปิดใช้งาน Family Symbol ถ้ายังไม่ถูกเปิดใช้งาน
        if not family_symbol.IsActive:
            family_symbol.Activate()
            doc.Regenerate()

        # ทดสอบกับจุดแรก
        if data:
            test_element = data[0]
            test_x, test_y = transform_coordinates_corrected(
                test_element["E"], test_element["N"], 
                BASE_E, BASE_N, 
                ANGLE_RADIANS,
                BASE_POINT_OFFSET_X, 
                BASE_POINT_OFFSET_Y
            )
            print("🧪 ทดสอบจุดแรก: X={:.2f}, Y={:.2f} ฟุต".format(test_x, test_y))

        for i, element_data in enumerate(data):
            element_no = element_data["ElementNo"]
            e, n = element_data["E"], element_data["N"]
            pile_cut_off = element_data.get("PileCutOff")
            
            try:
                # แปลงพิกัด
                x, y = transform_coordinates_corrected(
                    e, n, 
                    BASE_E, BASE_N, 
                    ANGLE_RADIANS,
                    BASE_POINT_OFFSET_X, 
                    BASE_POINT_OFFSET_Y
                )
                
                # สร้างตำแหน่ง XYZ (ใช้ Z = 0 สำหรับระดับของ Level)
                point = XYZ(x, y, 0)
                
                # สร้าง Structural Element พร้อมตั้งค่า Height Offset จากค่า Pile Cut Off
                element = create_structural_element_with_cutoff(
                    active_view, 
                    point, 
                    family_symbol, 
                    element_no, 
                    e, n, 
                    base_level, 
                    pile_cut_off
                )

                if element:
                    created_count += 1
                    if pile_cut_off is not None:
                        elements_with_cutoff += 1
                else:
                    failed_count += 1
                    failed_items.append(element_no)
                
                if (i + 1) % PROGRESS_UPDATE_INTERVAL == 0:
                    print("   📊 ความคืบหน้า: {}/{} (สำเร็จ: {}, ล้มเหลว: {})".format(
                        i + 1, len(data), created_count, failed_count))

            except Exception as ex:
                failed_count += 1
                failed_items.append(element_no)
                print("❌ ล้มเหลว {}: {}".format(element_no, str(ex)))

        t.Commit()
        
        print("\n✅ สร้างเสร็จสิ้น: {} ตำแหน่ง, ❌ ล้มเหลว: {} ตำแหน่ง".format(created_count, failed_count))
        print("📏 ตั้งค่า Height Offset จาก Level ฐานสำหรับ {} องค์ประกอบ".format(elements_with_cutoff))
        print("🏗️ Level ฐาน: {} (Elevation: {:.3f} ฟุต)".format(base_level.Name, base_level.Elevation))
        
        if failed_items:
            print("📋 รายการที่ล้มเหลว: {}".format(", ".join(failed_items[:10])))
            if len(failed_items) > 10:
                print("   ... และอีก {} รายการ".format(len(failed_items) - 10))

        # แสดงสรุปผล
        category_name = family_symbol.Category.Name
        show_task_dialog("สำเร็จ", 
            "✅ สร้าง {} สำเร็จ!\n\nสร้าง: {} ตำแหน่ง\nล้มเหลว: {} ตำแหน่ง\nตั้งค่า Height Offset: {} องค์ประกอบ\nLevel ฐาน: {}\nประเภท: {}".format(
                category_name, created_count, failed_count, elements_with_cutoff, base_level.Name, category_name))

    except Exception as ex:
        t.RollBack()
        print("❌ Transaction ล้มเหลว: {}".format(str(ex)))
        show_task_dialog("Error", "❌ เกิดข้อผิดพลาดในการสร้าง Structural Elements")

if __name__ == "__main__":
    main()