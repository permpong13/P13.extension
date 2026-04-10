# -*- coding: utf-8 -*-
__title__ = "F-Corner\nCoordinates"
__author__ = "Tee_เพิ่มพงษ์"
__doc__ = "Foundation Corner Coordinates with Leader Line to TextNote Center + Smart Offset, ทศนิยม 3 ตำแหน่ง, Pure IronPython"

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
from Autodesk.Revit.DB import *
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc = __revit__.ActiveUIDocument.Document
view = __revit__.ActiveUIDocument.ActiveView

print("เริ่มสร้าง Foundation Coordinate Marker (TextNote + Leader Line)")

# -------------------------------
# Project Base Point
# -------------------------------
def get_project_basepoint():
    try:
        bps = list(FilteredElementCollector(doc)
                   .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                   .WhereElementIsNotElementType())
        if bps:
            bp = bps[0]
            loc = bp.Location.Point
            n_param = bp.LookupParameter("N/S") or bp.LookupParameter("Northing")
            e_param = bp.LookupParameter("E/W") or bp.LookupParameter("Easting")
            n = n_param.AsDouble() if n_param else 1459042.584
            e = e_param.AsDouble() if e_param else 749356.687
            return loc, n, e
        return XYZ(0,0,0), 1459042.584, 749356.687
    except:
        return XYZ(0,0,0), 1459042.584, 749356.687

# -------------------------------
# คำนวณ Northing/Easting
# -------------------------------
def calculate_northing_easting(pt, base_loc, base_n, base_e):
    delta_e = (pt.X - base_loc.X) * 0.3048
    delta_n = (pt.Y - base_loc.Y) * 0.3048
    return base_n + delta_n, base_e + delta_e

# -------------------------------
# ลบจุดซ้ำ
# -------------------------------
def unique_points(points, tol=0.001):
    result = []
    for pt in points:
        if not any(pt.DistanceTo(r) < tol for r in result):
            result.append(pt)
    return result

# -------------------------------
# Simple Convex Hull 2D
# -------------------------------
def convex_hull(points):
    if len(points) <= 3:
        return points[:]
    pts = sorted(points, key=lambda p: (p.X, p.Y))
    def cross(o,a,b):
        return (a.X - o.X)*(b.Y - o.Y) - (a.Y - o.Y)*(b.X - o.X)
    lower = []
    for p in pts:
        while len(lower) >=2 and cross(lower[-2], lower[-1], p) <=0:
            lower.pop()
        lower.append(p)
    upper = []
    for p in reversed(pts):
        while len(upper) >=2 and cross(upper[-2], upper[-1], p) <=0:
            upper.pop()
        upper.append(p)
    hull = lower[:-1] + upper[:-1]
    return hull

# -------------------------------
# กำหนด TextNote ตำแหน่ง + Alignment
# -------------------------------
def get_text_position(pt, corner, leader_len=0.5):
    dx = leader_len if "Right" in corner else -leader_len
    dy = leader_len if "Top" in corner else -leader_len
    text_pt = XYZ(pt.X + dx, pt.Y + dy, pt.Z)
    alignment = HorizontalTextAlignment.Left if dx>0 else HorizontalTextAlignment.Right
    return text_pt, alignment

# -------------------------------
# Smart Offset ป้องกันข้อความซ้อนกัน
# -------------------------------
def smart_offset(new_pt, existing_pts, corner, step=0.3, min_dist=0.6, max_iterations=10):
    """
    ป้องกันการทับกันของ TextNote
    """
    x_direction = 1 if "Right" in corner else -1
    y_direction = 1 if "Top" in corner else -1
    
    adjusted_pt = new_pt
    iteration = 0
    
    while iteration < max_iterations:
        too_close = False
        for existing_pt in existing_pts:
            if adjusted_pt.DistanceTo(existing_pt) < min_dist:
                too_close = True
                break
        
        if not too_close:
            return adjusted_pt, False  # ไม่ต้องเพิ่มความยาวเส้น
        
        # ขยับตำแหน่ง
        adjusted_pt = XYZ(
            adjusted_pt.X + (x_direction * step),
            adjusted_pt.Y + (y_direction * step),
            adjusted_pt.Z
        )
        iteration += 1
    
    # ถ้าขยับจนเกินจำนวนครั้งที่กำหนดแล้วยังใกล้อยู่
    # ให้เพิ่มความยาวเส้น Leader แทน
    return new_pt, True

# -------------------------------
# สร้างเส้น Leader ที่ทำมุม 45 องศาและยาวถึง TextNote
# -------------------------------
def create_extended_45_degree_leader(view, foundation_pt, text_pt, corner, extend_leader=False):
    """สร้างเส้น Leader ที่ทำมุม 45 องศาจากมุม Foundation และยาวจนถึง TextNote"""
    try:
        # กำหนดทิศทางตามตำแหน่งของ TextNote
        dx = text_pt.X - foundation_pt.X
        dy = text_pt.Y - foundation_pt.Y
        
        # คำนวณจุดกึ่งกลางที่ทำมุม 45 องศา
        min_dist = min(abs(dx), abs(dy))
        
        # ถ้าต้องการเพิ่มความยาวเส้น Leader
        if extend_leader:
            min_dist = max(min_dist, 1.5)  # เพิ่มความยาวขั้นต่ำเป็น 1.5 ฟุต
        
        # กำหนดทิศทางของจุดกึ่งกลางตามตำแหน่ง
        if "Right" in corner and "Top" in corner:
            mid_pt = XYZ(foundation_pt.X + min_dist, foundation_pt.Y + min_dist, foundation_pt.Z)
        elif "Right" in corner and "Bottom" in corner:
            mid_pt = XYZ(foundation_pt.X + min_dist, foundation_pt.Y - min_dist, foundation_pt.Z)
        elif "Left" in corner and "Top" in corner:
            mid_pt = XYZ(foundation_pt.X - min_dist, foundation_pt.Y + min_dist, foundation_pt.Z)
        else:  # Left-Bottom
            mid_pt = XYZ(foundation_pt.X - min_dist, foundation_pt.Y - min_dist, foundation_pt.Z)
        
        # สร้างเส้นจาก Foundation ถึงจุดกึ่งกลาง (มุม 45 องศา)
        if foundation_pt.DistanceTo(mid_pt) > 0.001:
            leader_line1 = Line.CreateBound(foundation_pt, mid_pt)
            doc.Create.NewDetailCurve(view, leader_line1)
        
        # สร้างเส้นจากจุดกึ่งกลางถึง TextNote (แนวนอนหรือแนวตั้ง)
        if mid_pt.DistanceTo(text_pt) > 0.001:
            # ตรวจสอบว่าควรลากในแนวตั้งหรือแนวนอน
            if abs(mid_pt.X - text_pt.X) > abs(mid_pt.Y - text_pt.Y):
                # ลากแนวนอนก่อนแล้วค่อยแนวตั้ง
                horizontal_pt = XYZ(text_pt.X, mid_pt.Y, text_pt.Z)
                if mid_pt.DistanceTo(horizontal_pt) > 0.001:
                    leader_line2 = Line.CreateBound(mid_pt, horizontal_pt)
                    doc.Create.NewDetailCurve(view, leader_line2)
                if horizontal_pt.DistanceTo(text_pt) > 0.001:
                    leader_line3 = Line.CreateBound(horizontal_pt, text_pt)
                    doc.Create.NewDetailCurve(view, leader_line3)
            else:
                # ลากแนวตั้งก่อนแล้วค่อยแนวนอน
                vertical_pt = XYZ(mid_pt.X, text_pt.Y, text_pt.Z)
                if mid_pt.DistanceTo(vertical_pt) > 0.001:
                    leader_line2 = Line.CreateBound(mid_pt, vertical_pt)
                    doc.Create.NewDetailCurve(view, leader_line2)
                if vertical_pt.DistanceTo(text_pt) > 0.001:
                    leader_line3 = Line.CreateBound(vertical_pt, text_pt)
                    doc.Create.NewDetailCurve(view, leader_line3)
                    
    except Exception as e:
        print("⚠️ ไม่สามารถสร้าง Leader ได้: " + str(e))

# -------------------------------
# Main execution
# -------------------------------
def safe_execute():
    try:
        TransactionManager.Instance.EnsureInTransaction(doc)
        base_loc, base_n, base_e = get_project_basepoint()

        text_types = list(FilteredElementCollector(doc).OfClass(TextNoteType).ToElements())
        if not text_types:
            print("❌ ไม่พบ Text Note Type")
            return 0
        text_type = text_types[0]

        fnds = list(FilteredElementCollector(doc, view.Id)
                    .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                    .WhereElementIsNotElementType())
        non_pile = [f for f in fnds if "PILE" not in (f.Name or "").upper() and "PILE" not in (f.Symbol.FamilyName or "").upper()]
        total_foundations = len(non_pile)
        print("พบ Foundation ทั้งหมด: " + str(total_foundations) + " ชิ้น")

        opt = Options(); opt.View = view
        total = 0
        all_text_pts = []
        leader_len = 0.8  # ความยาวเริ่มต้น
        problem_count = 0
        success_count = 0

        # กำหนดเปอร์เซ็นต์ที่จะแสดงความคืบหน้า
        last_reported_percent = -1

        for i, fnd in enumerate(non_pile):
            # คำนวณเปอร์เซ็นต์ความคืบหน้า
            current_percent = int((float(i) / float(total_foundations)) * 100.0)
            
            # แสดงความคืบหน้าเมื่อเปอร์เซ็นต์เปลี่ยน (ทุก 10% หรือเมื่อจบ)
            if current_percent != last_reported_percent and (current_percent % 10 == 0 or i == total_foundations - 1):
                print("กำลังประมวลผล... " + str(current_percent) + "% (" + str(i) + "/" + str(total_foundations) + ")")
                last_reported_percent = current_percent
                
            try:
                geom = fnd.get_Geometry(opt)
                pts = []
                for g in geom:
                    if isinstance(g, Solid) and g.Faces.Size > 0:
                        for e in g.Edges:
                            c = e.AsCurve()
                            if c:
                                pts.extend([c.GetEndPoint(0), c.GetEndPoint(1)])
                pts = unique_points(pts)
                if not pts:
                    continue

                hull_pts = convex_hull(pts)
                cx = sum(p.X for p in hull_pts)/len(hull_pts)
                cy = sum(p.Y for p in hull_pts)/len(hull_pts)

                for pt in hull_pts:
                    n, e = calculate_northing_easting(pt, base_loc, base_n, base_e)
                    label = "N {0:.3f}\nE {1:.3f}".format(n, e)

                    # กำหนด corner
                    if pt.X >= cx and pt.Y >= cy:
                        corner = "Top-Right"
                    elif pt.X < cx and pt.Y >= cy:
                        corner = "Top-Left"
                    elif pt.X >= cx and pt.Y < cy:
                        corner = "Bottom-Right"
                    else:
                        corner = "Bottom-Left"

                    text_pt, alignment = get_text_position(pt, corner, leader_len)
                    
                    # ตรวจสอบและป้องกันการทับกัน
                    extend_leader = False
                    text_pt, extend_leader = smart_offset(text_pt, all_text_pts, corner)
                    all_text_pts.append(text_pt)

                    # สร้างเส้น Leader ที่ทำมุม 45 องศาและยาวถึง TextNote
                    create_extended_45_degree_leader(view, pt, text_pt, corner, extend_leader)

                    # สร้าง TextNote
                    text_note = TextNote.Create(doc, view.Id, text_pt, label, text_type.Id)
                    param_size = text_note.get_Parameter(BuiltInParameter.TEXT_SIZE)
                    if param_size:
                        param_size.Set(0.00656)  # 2 mm
                    text_note.HorizontalAlignment = alignment
                    total += 1
                    success_count += 1
                    
                    # แจ้งเตือนเฉพาะจุดที่ต้องเพิ่มความยาวเส้น Leader
                    if extend_leader:
                        print("⚠️ เพิ่มความยาวเส้น Leader สำหรับจุด N {0:.3f}, E {1:.3f}".format(n, e))
                        problem_count += 1

            except Exception as e:
                print("⚠️ ข้าม Foundation " + str(fnd.Id) + ": " + str(e))
                problem_count += 1
                continue

        TransactionManager.Instance.TransactionTaskDone()
        print("\n🎯 สร้าง Marker ทั้งหมด " + str(success_count) + " จุด จาก " + str(total) + " จุดที่พยายามสร้าง")
        if problem_count > 0:
            print("⚠️ พบปัญหา " + str(problem_count) + " จุด")
        return success_count

    except Exception as ex:
        TransactionManager.Instance.ForceCloseTransaction()
        print("❌ เกิดข้อผิดพลาด: " + str(ex))
        return 0

# -------------------------------
# Main
# -------------------------------
if not isinstance(view, ViewPlan):
    print("⚠️ โปรดเปิด Plan View ก่อนใช้งาน")
else:
    result = safe_execute()
    try:
        TransactionManager.Instance.ForceCloseTransaction()
    except:
        pass
    if result > 0:
        print("🎉 กระบวนการทำงานเสร็จสมบูรณ์")
    else:
        print("ℹ️ ไม่สามารถสร้าง Marker ได้")