# -*- coding: utf-8 -*-
__title__ = "Pipes Flow\nDirections"
__author__ = "Tee_เพิ่มพงษ์"
__doc__ = "Create flow direction arrows on all pipes in active view with family selection"

import clr
import sys
import math
import os

# Revit API References
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import Pipe
from System.Windows.Forms import *
from System.Drawing import *

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# -----------------------------
# Helper Functions
# -----------------------------

def meters_to_feet(m):
    return m * 3.28084

def get_all_detail_family_symbols(doc):
    """ดึง Family Symbols ทั้งหมดที่เป็น Detail Items หรือ Annotation"""
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
                
                # รับเฉพาะ Detail Items และ Annotation
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
        
        # ตรวจสอบจาก Family ด้วย
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

class PipeFlowArrowForm(Form):
    def __init__(self, families_data):
        self.families_data = families_data
        self.selected_family_symbol = None
        self.spacing_meters = 5.0  # ค่าเริ่มต้น
        self.InitializeComponent()
    
    def InitializeComponent(self):
        self.Text = "Pipe Flow Arrows - เลือก Family และระยะห่าง"
        self.Size = Size(500, 350)
        self.StartPosition = FormStartPosition.CenterScreen
        self.Font = Font("Microsoft Sans Serif", 9)
        
        # Label สำหรับเลือก Family
        self.lbl_family = Label()
        self.lbl_family.Text = "เลือก Family ลูกศร:"
        self.lbl_family.Location = Point(20, 20)
        self.lbl_family.Size = Size(150, 20)
        self.Controls.Add(self.lbl_family)
        
        # ComboBox สำหรับเลือก Family
        self.cb_family = ComboBox()
        self.cb_family.Location = Point(180, 20)
        self.cb_family.Size = Size(280, 20)
        self.cb_family.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb_family.SelectedIndexChanged += self.on_family_selected
        self.Controls.Add(self.cb_family)
        
        # Label สำหรับเลือก Type
        self.lbl_type = Label()
        self.lbl_type.Text = "เลือก Type:"
        self.lbl_type.Location = Point(20, 60)
        self.lbl_type.Size = Size(150, 20)
        self.Controls.Add(self.lbl_type)
        
        # ComboBox สำหรับเลือก Type
        self.cb_type = ComboBox()
        self.cb_type.Location = Point(180, 60)
        self.cb_type.Size = Size(280, 20)
        self.cb_type.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb_type.SelectedIndexChanged += self.on_type_selected
        self.Controls.Add(self.cb_type)
        
        # Label สำหรับระยะห่าง
        self.lbl_spacing = Label()
        self.lbl_spacing.Text = "ระยะห่างระหว่างลูกศร (เมตร):"
        self.lbl_spacing.Location = Point(20, 100)
        self.lbl_spacing.Size = Size(200, 20)
        self.Controls.Add(self.lbl_spacing)
        
        # TextBox สำหรับป้อนระยะห่าง
        self.txt_spacing = TextBox()
        self.txt_spacing.Text = "5.0"
        self.txt_spacing.Location = Point(220, 100)
        self.txt_spacing.Size = Size(100, 20)
        self.txt_spacing.TextChanged += self.on_spacing_changed
        self.Controls.Add(self.txt_spacing)
        
        # Radio buttons สำหรับระยะห่างที่ใช้บ่อย
        self.radio_5m = RadioButton()
        self.radio_5m.Text = "5 เมตร"
        self.radio_5m.Location = Point(20, 130)
        self.radio_5m.Size = Size(80, 20)
        self.radio_5m.Checked = True
        self.radio_5m.CheckedChanged += self.on_radio_5m_changed
        self.Controls.Add(self.radio_5m)
        
        self.radio_10m = RadioButton()
        self.radio_10m.Text = "10 เมตร"
        self.radio_10m.Location = Point(110, 130)
        self.radio_10m.Size = Size(80, 20)
        self.radio_10m.CheckedChanged += self.on_radio_10m_changed
        self.Controls.Add(self.radio_10m)
        
        self.radio_custom = RadioButton()
        self.radio_custom.Text = "กำหนดเอง"
        self.radio_custom.Location = Point(200, 130)
        self.radio_custom.Size = Size(100, 20)
        self.radio_custom.CheckedChanged += self.on_radio_custom_changed
        self.Controls.Add(self.radio_custom)
        
        # Label แสดงข้อมูล Family
        self.lbl_info = Label()
        self.lbl_info.Text = "ข้อมูล Family:"
        self.lbl_info.Location = Point(20, 160)
        self.lbl_info.Size = Size(440, 80)
        self.lbl_info.BorderStyle = BorderStyle.FixedSingle
        self.Controls.Add(self.lbl_info)
        
        # ปุ่มตกลง
        self.btn_ok = Button()
        self.btn_ok.Text = "ตกลง"
        self.btn_ok.Location = Point(300, 260)
        self.btn_ok.Size = Size(80, 30)
        self.btn_ok.Click += self.on_ok_click
        self.Controls.Add(self.btn_ok)
        
        # ปุ่มยกเลิก
        self.btn_cancel = Button()
        self.btn_cancel.Text = "ยกเลิก"
        self.btn_cancel.Location = Point(390, 260)
        self.btn_cancel.Size = Size(80, 30)
        self.btn_cancel.Click += self.on_cancel_click
        self.Controls.Add(self.btn_cancel)
        
        self.load_families()
    
    def load_families(self):
        """โหลดรายการ Family ลงใน ComboBox"""
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
                display_name = "{} - {}".format(symbol_info.symbol_name, symbol_info.category_name)
                self.cb_type.Items.Add((display_name, symbol_info))
        
        if self.cb_type.Items.Count > 0:
            self.cb_type.SelectedIndex = 0
            
        self.update_info()
    
    def on_type_selected(self, sender, event):
        """เมื่อเลือก Type ให้อัพเดทข้อมูล"""
        self.update_info()
    
    def on_spacing_changed(self, sender, event):
        """เมื่อแก้ไขระยะห่าง"""
        try:
            self.spacing_meters = float(self.txt_spacing.Text)
            self.radio_custom.Checked = True
        except:
            pass
    
    def on_radio_5m_changed(self, sender, event):
        if self.radio_5m.Checked:
            self.txt_spacing.Text = "5.0"
            self.spacing_meters = 5.0
    
    def on_radio_10m_changed(self, sender, event):
        if self.radio_10m.Checked:
            self.txt_spacing.Text = "10.0"
            self.spacing_meters = 10.0
    
    def on_radio_custom_changed(self, sender, event):
        if self.radio_custom.Checked:
            try:
                self.spacing_meters = float(self.txt_spacing.Text)
            except:
                self.spacing_meters = 5.0
    
    def update_info(self):
        """อัพเดทข้อมูล Family ที่เลือก"""
        if (self.cb_family.SelectedIndex >= 0 and 
            self.cb_type.SelectedIndex >= 0):
            
            family_name = self.cb_family.SelectedItem.ToString()
            display_name, symbol_info = self.cb_type.SelectedItem
            
            info_text = "Family: {}\nType: {}\nCategory: {}\nSymbol Name: {}\n\nระยะห่าง: {} เมตร".format(
                family_name,
                symbol_info.symbol_name,
                symbol_info.category_name,
                symbol_info.symbol_name,
                self.spacing_meters
            )
            self.lbl_info.Text = info_text
    
    def on_ok_click(self, sender, event):
        """เมื่อกดปุ่มตกลง"""
        if (self.cb_family.SelectedIndex < 0 or 
            self.cb_type.SelectedIndex < 0):
            MessageBox.Show("กรุณาเลือก Family และ Type", "ข้อผิดพลาด")
            return
        
        try:
            self.spacing_meters = float(self.txt_spacing.Text)
            if self.spacing_meters <= 0:
                MessageBox.Show("กรุณาป้อนระยะห่างที่มากกว่า 0", "ข้อผิดพลาด")
                return
        except:
            MessageBox.Show("กรุณาป้อนตัวเลขที่ถูกต้องสำหรับระยะห่าง", "ข้อผิดพลาด")
            return
        
        display_name, symbol_info = self.cb_type.SelectedItem
        self.selected_family_symbol = symbol_info.symbol
        self.DialogResult = DialogResult.OK
        self.Close()
    
    def on_cancel_click(self, sender, event):
        """เมื่อกดปุ่มยกเลิก"""
        self.DialogResult = DialogResult.Cancel
        self.Close()

def place_arrow(doc, point, direction, symbol, view):
    """วางลูกศรและหมุนตามทิศทาง"""
    try:
        inst = doc.Create.NewFamilyInstance(point, symbol, view)
        angle = math.atan2(direction.Y, direction.X)
        axis = Line.CreateBound(point, point + XYZ.BasisZ)
        ElementTransformUtils.RotateElement(doc, inst.Id, axis, angle)
        return True
    except Exception as e:
        print("⚠️ Failed to place arrow: " + str(e))
        return False

# -----------------------------
# Main Function
# -----------------------------
def create_pipe_flow_arrows():
    # ใช้ Revit TaskDialog โดยตรงในฟังก์ชัน
    from Autodesk.Revit.UI import TaskDialog
    
    try:
        view = doc.ActiveView
        if view.ViewType != ViewType.FloorPlan:
            TaskDialog.Show("Error", "Please open a Plan View before running this script.")
            return

        pipes = FilteredElementCollector(doc, view.Id).OfClass(Pipe).ToElements()
        if not pipes:
            TaskDialog.Show("Error", "No pipes found in the active view.")
            return

        # ดึงรายการ Family Symbols
        all_symbols_info = get_all_detail_family_symbols(doc)
        if not all_symbols_info:
            TaskDialog.Show("Error", "No Detail Item or Annotation Families found in project.")
            return

        # แสดงฟอร์มให้ผู้ใช้เลือก Family และระยะห่าง
        form = PipeFlowArrowForm(all_symbols_info)
        result = form.ShowDialog()
        
        if result != DialogResult.OK or not form.selected_family_symbol:
            return
        
        selected_symbol = form.selected_family_symbol
        spacing_meters = form.spacing_meters
        spacing_feet = meters_to_feet(spacing_meters)

        # Activate symbol (separate transaction)
        if not selected_symbol.IsActive:
            t_act = Transaction(doc, "Activate Symbol")
            t_act.Start()
            selected_symbol.Activate()
            t_act.Commit()

        # === Main Transaction for placing arrows ===
        t = Transaction(doc, "Create Pipe Flow Arrows")
        t.Start()

        arrow_count = 0
        processed = 0

        for pipe in pipes:
            processed += 1
            loc = pipe.Location
            if not hasattr(loc, 'Curve'):
                continue

            curve = loc.Curve
            length = curve.Length
            if length < 0.1:
                continue

            # Pipe direction
            try:
                direction = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
            except:
                direction = XYZ(1, 0, 0)

            # Arrow placement
            if length < spacing_feet:
                distances = [length / 2.0]
            else:
                num_arrows = max(1, int(math.floor(length / spacing_feet)))
                distances = [(i + 1) * spacing_feet for i in range(num_arrows)]

            for dist in distances:
                if dist >= length:
                    dist = length - 0.5
                try:
                    param = dist / length
                    point = curve.Evaluate(param, True)
                    if place_arrow(doc, point, direction, selected_symbol, view):
                        arrow_count += 1
                except Exception as e:
                    print("⚠️ Error placing arrow: " + str(e))
                    continue

        t.Commit()

        TaskDialog.Show("Flow Arrows Created",
                        "Processed {} pipes\nPlaced {} flow arrows.\nSpacing: {} meters".format(
                            processed, arrow_count, spacing_meters))

    except Exception as e:
        # Ensure rollback if any crash
        try:
            if 't' in locals() and t.GetStatus() == TransactionStatus.Started:
                t.RollBack()
        except:
            pass
        error_msg = "Error: {}\nLine: {}".format(e, sys.exc_info()[2].tb_lineno)
        TaskDialog.Show("Script Error", error_msg)

# -----------------------------
# Run
# -----------------------------
if __name__ == "__main__":
    create_pipe_flow_arrows()