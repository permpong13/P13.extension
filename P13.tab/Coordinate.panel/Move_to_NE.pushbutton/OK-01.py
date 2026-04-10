# -*- coding: utf-8 -*-
__doc__ = "ย้าย Family ที่เลือกไปยังพิกัด North/East ที่ต้องการ (หน่วย: เมตร)"
__title__ = "Move to\nN/E"
__author__ = "เพิ่มพงษ์"

import math
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from pyrevit import forms, DB
from Autodesk.Revit.UI import Selection
from Autodesk.Revit.UI.Selection import ObjectType

from System.Windows.Forms import Form, Label, TextBox, Button, FormStartPosition, DialogResult
from System.Drawing import Size, Point, Font

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


# ============================================================
# ฟอร์มกรอกพิกัด N / E
# ============================================================
class NEInputForm(Form):
    def __init__(self):
        self.north = None
        self.east = None
        self.InitializeComponent()

    def InitializeComponent(self):
        self.Text = "กรอกพิกัด N / E (เมตร)"
        self.Size = Size(350, 200)
        self.StartPosition = FormStartPosition.CenterScreen
        self.Font = Font("Microsoft Sans Serif", 9)

        # Label N
        self.lbl_n = Label()
        self.lbl_n.Text = "Northing (N) [m]:"
        self.lbl_n.Location = Point(20, 20)
        self.lbl_n.Size = Size(120, 20)
        self.Controls.Add(self.lbl_n)

        # TextBox N
        self.tb_n = TextBox()
        self.tb_n.Location = Point(150, 20)
        self.tb_n.Size = Size(150, 20)
        self.Controls.Add(self.tb_n)

        # Label E
        self.lbl_e = Label()
        self.lbl_e.Text = "Easting (E) [m]:"
        self.lbl_e.Location = Point(20, 60)
        self.lbl_e.Size = Size(120, 20)
        self.Controls.Add(self.lbl_e)

        # TextBox E
        self.tb_e = TextBox()
        self.tb_e.Location = Point(150, 60)
        self.tb_e.Size = Size(150, 20)
        self.Controls.Add(self.tb_e)

        # ปุ่ม OK
        self.btn_ok = Button()
        self.btn_ok.Text = "ตกลง"
        self.btn_ok.Location = Point(140, 110)
        self.btn_ok.Size = Size(80, 30)
        self.btn_ok.Click += self.on_ok_click
        self.Controls.Add(self.btn_ok)

        # ปุ่ม Cancel
        self.btn_cancel = Button()
        self.btn_cancel.Text = "ยกเลิก"
        self.btn_cancel.Location = Point(230, 110)
        self.btn_cancel.Size = Size(80, 30)
        self.btn_cancel.Click += self.on_cancel_click
        self.Controls.Add(self.btn_cancel)

    def on_ok_click(self, sender, args):
        try:
            n_val = float(self.tb_n.Text.replace(",", "").strip())
            e_val = float(self.tb_e.Text.replace(",", "").strip())
        except:
            forms.alert("กรุณากรอกค่าพิกัดเป็นตัวเลข (เมตร)", title="ข้อมูลไม่ถูกต้อง")
            return

        self.north = n_val
        self.east = e_val
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel_click(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()


# ============================================================
# ฟังก์ชันดึงข้อมูล Base Point - ใช้สูตรเดียวกับสคริปต์คำนวณพิกัด
# ============================================================
def get_base_point_info(doc):
    """ดึงข้อมูล Base Point ตามสูตรเดียวกับสคริปต์คำนวณพิกัด"""
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
                        break
        except Exception as e:
            # ไม่แสดง error ใน console
            pass

    if not basepoint_found:
        # ไม่แสดง warning ใน console
        return 0.0, 0.0, 0.0

    return angle, bp_ewest, bp_nsouth


def rotate(x, y, theta):
    """หมุนพิกัดตามสูตรเดียวกับสคริปต์คำนวณพิกัด"""
    return [
        math.cos(theta) * x + math.sin(theta) * y,
        -math.sin(theta) * x + math.cos(theta) * y
    ]


def find_coord(x, y, theta, bp_ewest, bp_nsouth):
    """คำนวณพิกัดตามสูตรเดียวกับสคริปต์คำนวณพิกัด"""
    # หมุนพิกัด
    rotated = rotate(x, y, theta)
    
    # บวกกับพิกัด Base Point
    result = [
        rotated[0] + bp_ewest,  # East (feet)
        rotated[1] + bp_nsouth  # North (feet)
    ]
    
    # แปลงหน่วยจาก feet เป็น meters
    result_meters = [coord * 0.3048 for coord in result]
    
    return (result_meters[1], result_meters[0])  # (North, East) ในหน่วยเมตร


def reverse_transform(target_north_m, target_east_m, angle, bp_ewest, bp_nsouth):
    """
    แปลงพิกัดจากระบบ Survey (N,E ในเมตร) กลับเป็นระบบ Revit (X,Y ในฟุต)
    """
    # แปลงจากเมตรเป็นฟุต
    target_north_ft = target_north_m / 0.3048
    target_east_ft = target_east_m / 0.3048
    
    # ลบ Base Point offset
    relative_north = target_north_ft - bp_nsouth
    relative_east = target_east_ft - bp_ewest
    
    # หมุนกลับ (inverse rotation)
    # ใช้มุมลบเพื่อหมุนกลับ
    inverse_angle = -angle
    revit_coords = rotate(relative_east, relative_north, inverse_angle)
    
    return revit_coords[0], revit_coords[1]  # X, Y ในระบบ Revit (ฟุต)


# ============================================================
# ฟังก์ชันช่วยหาตำแหน่ง Element (X,Y,Z)
# ============================================================
def get_element_xy(element):
    """คืนค่า (X, Y, Z) ของตำแหน่ง element"""
    loc = element.Location
    if isinstance(loc, DB.LocationPoint):
        p = loc.Point
        return p.X, p.Y, p.Z
    elif isinstance(loc, DB.LocationCurve):
        mid = loc.Curve.Evaluate(0.5, True)
        return mid.X, mid.Y, mid.Z
    elif hasattr(element, 'GetTransform'):
        # สำหรับ Family Instance
        transform = element.GetTransform()
        if transform:
            origin = transform.Origin
            return origin.X, origin.Y, origin.Z
    else:
        # พยายามใช้ BoundingBox
        bbox = element.get_BoundingBox(None)
        if bbox:
            center = (bbox.Min + bbox.Max) * 0.5
            return center.X, center.Y, center.Z
    
    return None, None, None


# ============================================================
# MAIN
# ============================================================
def main():
    # ตรวจสอบว่ามีเอกสารถูกเปิดหรือไม่
    if doc is None:
        forms.alert("ไม่พบไฟล์ Revit ที่เปิดอยู่", exitscript=True)

    # ดึง selection ปัจจุบัน
    sel_ids = list(uidoc.Selection.GetElementIds())

    # ถ้าไม่ได้เลือกอะไร ให้บังคับให้เลือก
    if not sel_ids:
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, "เลือก Family หรือ Element ที่ต้องการย้าย")
            if ref:
                sel_ids = [ref.ElementId]
        except:
            forms.alert("ไม่ได้เลือก Element ใด ๆ", exitscript=True)

    if not sel_ids:
        forms.alert("ไม่ได้เลือก Element ใด ๆ", exitscript=True)

    elements = [doc.GetElement(eid) for eid in sel_ids]
    elements = [e for e in elements if e is not None]

    if not elements:
        forms.alert("ไม่พบ Element ที่สามารถใช้งานได้", exitscript=True)

    # อ่านข้อมูล Base Point ตามสูตรเดียวกับสคริปต์คำนวณพิกัด
    angle, bp_ewest, bp_nsouth = get_base_point_info(doc)

    # แสดงฟอร์มกรอกพิกัด N/E
    form = NEInputForm()
    result = form.ShowDialog()

    if result != DialogResult.OK:
        return

    target_n = form.north  # เมตร
    target_e = form.east   # เมตร

    # ใช้ element ตัวแรกเป็น reference ในการหา delta move
    first_elem = elements[0]
    cur_x, cur_y, cur_z = get_element_xy(first_elem)
    if cur_x is None:
        forms.alert("Element แรกไม่มี Location ที่สามารถย้ายได้", exitscript=True)

    # คำนวณตำแหน่งใหม่จาก N/E (เมตร) → X,Y (ฟุต)
    new_x, new_y = reverse_transform(target_n, target_e, angle, bp_ewest, bp_nsouth)

    # สร้างจุดใหม่
    new_point = DB.XYZ(new_x, new_y, cur_z)
    cur_point = DB.XYZ(cur_x, cur_y, cur_z)
    move_vec = new_point - cur_point

    # ย้ายทุก element ตาม move_vec เดียวกัน
    with DB.Transaction(doc, "Move Elements to N/E") as t:
        t.Start()
        try:
            moved_count = 0
            for e in elements:
                try:
                    DB.ElementTransformUtils.MoveElement(doc, e.Id, move_vec)
                    moved_count += 1
                except Exception as e_move:
                    # ไม่แสดง error ใน console
                    pass
            
            t.Commit()
            
        except Exception as ex:
            t.RollBack()
            forms.alert("เกิดข้อผิดพลาดในการย้าย Element:\n{}".format(ex), title="Error", exitscript=True)

    forms.alert(
        "ย้าย Element {} ตัว ไปยังพิกัด:\nN = {:.3f} m\nE = {:.3f} m\n\nสำเร็จแล้ว 🎉".format(
            len(elements), target_n, target_e
        ),
        title="สำเร็จ"
    )


if __name__ == "__main__":
    main()