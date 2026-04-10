# -*- coding: utf-8 -*-
__title__ = "Pipe\nAutoTag"
__doc__ = "Auto Tag ท่อทุกเส้นใน View"

from pyrevit import revit, DB, forms
import math
import sys

doc = revit.doc
uidoc = revit.uidoc
view = doc.ActiveView

# --------------------------------------------
# Configuration
# --------------------------------------------
def get_scale_config():
    VIEW_SCALE = float(view.Scale) if hasattr(view, "Scale") and view.Scale else 100.0
    scale_factor = VIEW_SCALE / 100.0
    return {
        "view_scale": VIEW_SCALE,
        "scale_factor": scale_factor,
        "row_spacing": max(0.2, 0.4 * scale_factor),
        "column_spacing": max(0.15, 0.3 * scale_factor),
        "tag_distance": max(0.5, 1.0 * scale_factor),
        "leader_length": max(2.0, 3.0 * scale_factor),
    }


# --------------------------------------------
# Visibility
# --------------------------------------------
def is_element_visible_in_view(element, view):
    try:
        if element.IsHidden(view):
            return False
        
        # สำหรับ Revit 2021+
        if hasattr(DB, 'ElementVisibility'):
            if not view.IsElementVisibleInTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate, element.Id):
                return False
        
        bbox = element.get_BoundingBox(view)
        if not bbox:
            return False

        if bbox.Min.X == bbox.Max.X and bbox.Min.Y == bbox.Max.Y:
            return False

        return True
    except:
        return True


# --------------------------------------------
# Collect Pipes Only
# --------------------------------------------
def get_all_pipes_in_view():
    try:
        collector = DB.FilteredElementCollector(doc, view.Id) \
            .OfCategory(DB.BuiltInCategory.OST_PipeCurves) \
            .WhereElementIsNotElementType()
        pipes = list(collector)
        
        # กรองเฉพาะท่อที่อยู่ในระดับความสูงของ view
        filtered_pipes = []
        for pipe in pipes:
            if can_element_be_tagged(pipe):
                filtered_pipes.append(pipe)
        
        return filtered_pipes
    except Exception as e:
        print("Error getting pipes: {}".format(str(e)))
        return []


# --------------------------------------------
# Basic Validation
# --------------------------------------------
def can_element_be_tagged(e):
    try:
        if e.IsElementType:
            return False
        if not e.Location:
            return False
        if e.ViewSpecific:  # ตรวจสอบว่าเป็น element เฉพาะ view หรือไม่
            return False
        return True
    except:
        return False


# --------------------------------------------
# Selection Modes
# --------------------------------------------
def filter_pipes_by_selection(pipes):
    if not pipes:
        return []

    choice = forms.SelectFromList.show(
        [
            "Tag All Pipes",
            "Tag Only Visible Pipes", 
            "Select Visible Pipes (Manual)",
            "Select Any Pipes (Manual)",
            "Select From List",
            "Cancel"
        ],
        title="Pipe Selection Mode",
        multiselect=False
    )

    if not choice or choice == "Cancel":
        return []

    if choice == "Tag All Pipes":
        return [p for p in pipes if can_element_be_tagged(p)]

    if choice == "Tag Only Visible Pipes":
        return [p for p in pipes if is_element_visible_in_view(p, view) and can_element_be_tagged(p)]

    if choice == "Select Visible Pipes (Manual)":
        visible_ids = {p.Id for p in pipes if is_element_visible_in_view(p, view)}
        try:
            picked = uidoc.Selection.PickObjects(
                DB.UI.Selection.ObjectType.Element,
                "Pick visible pipes"
            )
            result = []
            for ref in picked:
                element = doc.GetElement(ref.ElementId)
                if element.Id in visible_ids:
                    result.append(element)
            return result
        except Exception as e:
            print("Selection cancelled: {}".format(str(e)))
            return []

    if choice == "Select Any Pipes (Manual)":
        try:
            picked = uidoc.Selection.PickObjects(
                DB.UI.Selection.ObjectType.Element,
                "Pick pipes"
            )
            return [doc.GetElement(ref.ElementId) for ref in picked]
        except Exception as e:
            print("Selection cancelled: {}".format(str(e)))
            return []

    if choice == "Select From List":
        names = []
        pipe_dict = {}
        for i, p in enumerate(pipes):
            if not can_element_be_tagged(p):
                continue
            pid = p.Id.IntegerValue
            dia = ""
            dp = p.get_Parameter(DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
            if dp and dp.AsDouble() > 0:
                dia = " - {} mm".format(int(dp.AsDouble() * 304.8))
            name = "[{}] Pipe{}".format(pid, dia)
            names.append(name)
            pipe_dict[name] = p

        if not names:
            forms.alert("No taggable pipes found.")
            return []

        selected_names = forms.SelectFromList.show(
            names, 
            title="Select Pipes to Tag",
            multiselect=True
        )
        
        if not selected_names:
            return []
            
        return [pipe_dict[name] for name in selected_names if name in pipe_dict]

    return []


# --------------------------------------------
# Geometry Helpers
# --------------------------------------------
def get_center(e):
    try:
        if hasattr(e.Location, "Curve"):
            curve = e.Location.Curve
            if curve:
                return curve.Evaluate(0.5, True)
        
        bb = e.get_BoundingBox(view)
        if bb:
            return DB.XYZ(
                (bb.Min.X + bb.Max.X) / 2,
                (bb.Min.Y + bb.Max.Y) / 2, 
                (bb.Min.Z + bb.Max.Z) / 2
            )
    except Exception as ex:
        print("Error getting center: {}".format(str(ex)))
    
    return DB.XYZ.Zero


def calculate_distance(a, b):
    return ((a.X - b.X)**2 + (a.Y - b.Y)**2 + (a.Z - b.Z)**2) ** 0.5


def get_pipe_direction(p):
    try:
        if hasattr(p.Location, "Curve"):
            curve = p.Location.Curve
            if curve:
                d = curve.Direction
                L = math.sqrt(d.X**2 + d.Y**2 + d.Z**2)
                if L > 0:
                    return DB.XYZ(d.X/L, d.Y/L, d.Z/L)
    except:
        pass
    return DB.XYZ.BasisX


# --------------------------------------------
# Grouping
# --------------------------------------------
def group_pipes(pipes, grid=2.0):  # ลด grid size เพื่อการจัดกลุ่มที่ดีขึ้น
    groups = {}
    for p in pipes:
        c = get_center(p)
        if c != DB.XYZ.Zero:
            key = (int(c.X / grid), int(c.Y / grid))
            groups.setdefault(key, []).append(p)
    return groups


# --------------------------------------------
# Tag Positioning
# --------------------------------------------
def calculate_tag_positions(pipes, config):
    positions = []
    groups = group_pipes(pipes)

    for key, items in groups.items():
        if not items:
            continue
            
        # เรียงลำดับท่อตามตำแหน่ง
        items.sort(key=lambda x: (get_center(x).Z, get_center(x).Y, get_center(x).X))
        
        # กำหนดจำนวนแถวและคอลัมน์
        row_cap = min(6, max(2, len(items) // 3 + 1))
        rows = [items[i:i + row_cap] for i in range(0, len(items), row_cap)]

        for ridx, row in enumerate(rows):
            if not row:
                continue
                
            # คำนวณความสูงของแถว
            row_centers = [get_center(p) for p in row if get_center(p) != DB.XYZ.Zero]
            if not row_centers:
                continue
                
            avg_z = sum(c.Z for c in row_centers) / len(row_centers)
            row_z = avg_z - ridx * config["row_spacing"]

            for cidx, p in enumerate(row):
                pc = get_center(p)
                if pc == DB.XYZ.Zero:
                    continue
                    
                d = get_pipe_direction(p)
                perp = DB.XYZ(-d.Y, d.X, 0)
                L = math.sqrt(perp.X**2 + perp.Y**2)
                if L > 0:
                    perp = DB.XYZ(perp.X/L, perp.Y/L, 0)
                else:
                    perp = DB.XYZ.BasisY
                
                # ตำแหน่งฐานสำหรับแท็ก
                base = DB.XYZ(
                    pc.X + perp.X * config["tag_distance"],
                    pc.Y + perp.Y * config["tag_distance"], 
                    row_z
                )

                # offset ตามคอลัมน์
                col_offset = (cidx - len(row) / 2) * config["column_spacing"]
                pos = DB.XYZ(
                    base.X + perp.X * col_offset,
                    base.Y + perp.Y * col_offset,
                    base.Z
                )

                positions.append({
                    "pipe": p,
                    "position": pos,
                    "pipe_center": pc,
                    "group": key
                })

    return positions


# --------------------------------------------
# Align tags per group
# --------------------------------------------
def align_per_group(positions):
    if not positions:
        return positions
        
    groups = {}
    for p in positions:
        groups.setdefault(p["group"], []).append(p)

    result = []
    is_plan = view.ViewType in (DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.EngineeringPlan)

    for key, items in groups.items():
        if not items:
            continue
            
        if is_plan:
            # จัดแนวในแผนผัง - ใช้ค่า Z เดียวกัน
            zs = [it["position"].Z for it in items]
            base_z = sorted(zs)[len(zs) // 2]
            for it in items:
                pos = it["position"]
                it["position"] = DB.XYZ(pos.X, pos.Y, base_z)
                result.append(it)
        else:
            # จัดแนวในภาพด้าน - ใช้ค่า X เดียวกัน
            xs = [it["position"].X for it in items]
            base_x = sorted(xs)[len(xs) // 2]
            for it in items:
                pos = it["position"]
                it["position"] = DB.XYZ(base_x, pos.Y, pos.Z)
                result.append(it)

    return result


# --------------------------------------------
# Auto-split tiers (overlap handling)
# --------------------------------------------
def auto_split(positions, min_dist, offset):
    if not positions:
        return positions
        
    final = []
    for pos in positions:
        p = pos["position"]
        tier = 0
        max_tiers = 5  # จำกัดจำนวนชั้นเพื่อป้องกันการลูปไม่สิ้นสุด
        
        while tier < max_tiers:
            hit = False
            for other in final:
                if calculate_distance(p, other["position"]) < min_dist:
                    hit = True
                    break
            if not hit:
                break

            p = DB.XYZ(p.X, p.Y, p.Z - offset)  # เคลื่อนลงด้านล่าง
            tier += 1

        if tier < max_tiers:
            pos["position"] = p
            final.append(pos)

    return final


# --------------------------------------------
# Check existing tags
# --------------------------------------------
def has_existing_tag(pipe):
    try:
        existing_tags = DB.FilteredElementCollector(doc, view.Id) \
            .OfClass(DB.IndependentTag) \
            .ToElements()
            
        for tag in existing_tags:
            try:
                tagged_elements = tag.GetTaggedElements()
                for ref in tagged_elements:
                    if ref.ElementId == pipe.Id:
                        return True
            except:
                continue
                
        return False
    except:
        return False


# --------------------------------------------
# Get proper reference for tagging
# --------------------------------------------
def get_pipe_reference(pipe):
    try:
        # ลองใช้ Reference ที่มีอยู่
        if hasattr(pipe, 'GetReferences'):
            refs = pipe.GetReferences(DB.FamilyInstanceReferenceType.Center)
            if refs and len(refs) > 0:
                return refs[0]
        
        # สำหรับ Curve Elements เช่น ท่อ
        if hasattr(pipe, 'Location') and hasattr(pipe.Location, 'Curve'):
            curve = pipe.Location.Curve
            if curve:
                # สร้าง reference จาก curve
                return DB.Reference(pipe)
                
    except Exception as e:
        print("Error getting reference: {}".format(str(e)))
    
    return DB.Reference(pipe)


# --------------------------------------------
# Tag Creation
# --------------------------------------------
def create_pipe_tags(positions):
    created = []
    failed = []

    with revit.Transaction("Create Pipe Tags"):
        for pos in positions:
            pipe = pos["pipe"]
            location = pos["position"]

            if has_existing_tag(pipe):
                failed.append({"pipe": pipe.Id, "error": "Already Tagged"})
                continue

            try:
                # Get reference for tagging
                reference = get_pipe_reference(pipe)
                if not reference:
                    failed.append({"pipe": pipe.Id, "error": "No valid reference"})
                    continue

                # สร้างแท็ก
                tag = DB.IndependentTag.Create(
                    doc,
                    view.Id,
                    reference,
                    False,
                    DB.TagMode.TM_ADDBY_CATEGORY,
                    DB.TagOrientation.Horizontal,
                    location
                )

                if tag:
                    created.append(tag)
                    # พยายามตั้งค่า leader
                    try:
                        tag.HasLeader = True
                        # พยายามตั้งค่า leader end point
                        tag.LeaderEndCondition = DB.LeaderEndCondition.Free
                        tag.LeaderEnd = pos["pipe_center"]
                    except Exception as leader_error:
                        print("Leader setup failed: {}".format(str(leader_error)))
                else:
                    failed.append({"pipe": pipe.Id, "error": "Tag creation failed"})

            except Exception as e:
                failed.append({"pipe": pipe.Id, "error": str(e)})
                print("Tag creation error: {}".format(str(e)))

    return created, failed


# --------------------------------------------
# MAIN
# --------------------------------------------
def main():
    # ตรวจสอบว่าเป็น view ที่เหมาะสมหรือไม่
    if view.ViewType not in [DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, 
                            DB.ViewType.EngineeringPlan, DB.ViewType.Section, 
                            DB.ViewType.Elevation, DB.ViewType.ThreeD]:
        forms.alert("This tool works only in Plan, Section, Elevation, or 3D views.")
        return

    pipes = get_all_pipes_in_view()
    if not pipes:
        forms.alert("No pipes found in current view.")
        return

    selected = filter_pipes_by_selection(pipes)
    if not selected:
        return

    config = get_scale_config()

    positions = calculate_tag_positions(selected, config)
    if not positions:
        forms.alert("Could not calculate tag positions.")
        return

    positions = align_per_group(positions)
    positions = auto_split(positions, config["tag_distance"] * 0.8, config["row_spacing"])

    created, failed = create_pipe_tags(positions)

    # แสดงผลลัพธ์
    msg = "Created {} tags".format(len(created))
    if failed:
        msg += "\nFailed {} pipes:".format(len(failed))
        for f in failed[:10]:  # แสดงเพียง 10 อันแรก
            msg += "\n• {} - {}".format(f["pipe"], f["error"])
        if len(failed) > 10:
            msg += "\n• ... and {} more".format(len(failed) - 10)

    forms.alert(msg)


if __name__ == "__main__":
    main()