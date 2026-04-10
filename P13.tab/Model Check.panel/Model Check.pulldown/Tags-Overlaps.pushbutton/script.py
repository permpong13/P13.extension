# -*- coding: utf-8 -*-
__title__ = "Tags\nOverlaps"
__author__ = "Tee_เพิ่มพงษ์"
__doc__ = "ตรวจสอบ Tags ที่ทับกันโดยใช้ Bounding Box ของ Tag และลดขนาด"

from pyrevit import revit, DB, script
from System.Collections.Generic import List
import math

doc = revit.doc
uidoc = revit.uidoc
view = doc.ActiveView
output = script.get_output()
output.title = "ตรวจสอบ Tags ที่ทับกัน"

# --------------------------------------------------------
# 🔧 ฟังก์ชันช่วย - ใช้ Bounding Box ของ Tag โดยตรง
# --------------------------------------------------------
def get_tag_text(tag):
    """ ดึงข้อความจาก Tag """
    try:
        if hasattr(tag, 'TagText'):
            text = tag.TagText
            if text and str(text).strip():
                return text
        
        if hasattr(tag, 'Text'):
            text = tag.Text
            if text and str(text).strip():
                return text
        
        # ตรวจสอบพารามิเตอร์
        text_params = ["Text", "Tag Text", "Contents", "Value", "Mark", "Type Mark"]
        for param_name in text_params:
            param = tag.LookupParameter(param_name)
            if param and param.HasValue:
                if param.StorageType == DB.StorageType.String:
                    text = param.AsString()
                else:
                    text = param.AsValueString()
                
                if text and str(text).strip():
                    return text
        
        return None
    except:
        return None

def get_reduced_tag_bounds(tag, reduction_ratio=0.3):
    """ สร้าง Bounding Box เล็กๆ สำหรับ Tag โดยลดขนาดลง 70% """
    try:
        tag_bbox = tag.get_BoundingBox(view)
        if not tag_bbox:
            return None
        
        # คำนวณศูนย์กลาง
        center_x = (tag_bbox.Min.X + tag_bbox.Max.X) / 2
        center_y = (tag_bbox.Min.Y + tag_bbox.Max.Y) / 2
        center_z = (tag_bbox.Min.Z + tag_bbox.Max.Z) / 2
        
        # คำนวณขนาดใหม่ (ลดลงเหลือ 30% ของขนาดเดิม)
        original_width = tag_bbox.Max.X - tag_bbox.Min.X
        original_height = tag_bbox.Max.Y - tag_bbox.Min.Y
        
        new_width = original_width * reduction_ratio
        new_height = original_height * reduction_ratio
        
        # กำหนดค่าขั้นต่ำเพื่อป้องกันขนาดเล็กเกินไป
        min_size = 0.3
        new_width = max(new_width, min_size)
        new_height = max(new_height, min_size)
        
        # สร้าง Bounding Box ใหม่ที่เล็กกว่า
        min_point = DB.XYZ(
            center_x - new_width / 2,
            center_y - new_height / 2,
            center_z
        )
        max_point = DB.XYZ(
            center_x + new_width / 2,
            center_y + new_height / 2,
            center_z
        )
        
        class ReducedTagBBox:
            def __init__(self, min_pt, max_pt):
                self.Min = min_pt
                self.Max = max_pt
        
        return ReducedTagBBox(min_point, max_point)
        
    except Exception as e:
        print("ข้อผิดพลาดในการสร้างขอบเขต Tag ขนาดเล็ก: {}".format(e))
        return None

def get_very_small_tag_bounds(tag, reduction_ratio=0.15):
    """ สร้าง Bounding Box เล็กมากๆ สำหรับ Tag (เหลือ 15% ของขนาดเดิม) """
    try:
        tag_bbox = tag.get_BoundingBox(view)
        if not tag_bbox:
            return None
        
        # คำนวณศูนย์กลาง
        center_x = (tag_bbox.Min.X + tag_bbox.Max.X) / 2
        center_y = (tag_bbox.Min.Y + tag_bbox.Max.Y) / 2
        center_z = (tag_bbox.Min.Z + tag_bbox.Max.Z) / 2
        
        # คำนวณขนาดใหม่ (ลดลงเหลือ 15% ของขนาดเดิม)
        original_width = tag_bbox.Max.X - tag_bbox.Min.X
        original_height = tag_bbox.Max.Y - tag_bbox.Min.Y
        
        new_width = original_width * reduction_ratio
        new_height = original_height * reduction_ratio
        
        # กำหนดค่าขั้นต่ำ
        min_size = 0.2
        new_width = max(new_width, min_size)
        new_height = max(new_height, min_size)
        
        # สร้าง Bounding Box ใหม่ที่เล็กมาก
        min_point = DB.XYZ(
            center_x - new_width / 2,
            center_y - new_height / 2,
            center_z
        )
        max_point = DB.XYZ(
            center_x + new_width / 2,
            center_y + new_height / 2,
            center_z
        )
        
        class VerySmallTagBBox:
            def __init__(self, min_pt, max_pt):
                self.Min = min_pt
                self.Max = max_pt
        
        return VerySmallTagBBox(min_point, max_point)
        
    except Exception as e:
        print("ข้อผิดพลาดในการสร้างขอบเขต Tag ขนาดเล็กมาก: {}".format(e))
        return None

def is_any_overlap(bbox1, bbox2, tolerance=0.01):
    """ ตรวจสอบการทับกันใดๆ ด้วย tolerance เล็กน้อย """
    if not bbox1 or not bbox2:
        return False
    
    # ตรวจสอบการทับกันในแนว X และ Y
    x_overlap = (bbox1.Min.X <= bbox2.Max.X + tolerance and 
                 bbox1.Max.X >= bbox2.Min.X - tolerance)
    
    y_overlap = (bbox1.Min.Y <= bbox2.Max.Y + tolerance and 
                 bbox1.Max.Y >= bbox2.Min.Y - tolerance)
    
    return x_overlap and y_overlap

def calculate_overlap_area(bbox1, bbox2):
    """ คำนวณพื้นที่ทับกัน """
    if not bbox1 or not bbox2:
        return 0
    
    overlap_min_x = max(bbox1.Min.X, bbox2.Min.X)
    overlap_min_y = max(bbox1.Min.Y, bbox2.Min.Y)
    overlap_max_x = min(bbox1.Max.X, bbox2.Max.X)
    overlap_max_y = min(bbox1.Max.Y, bbox2.Max.Y)
    
    if overlap_max_x <= overlap_min_x or overlap_max_y <= overlap_min_y:
        return 0
    
    return (overlap_max_x - overlap_min_x) * (overlap_max_y - overlap_min_y)

def calculate_overlap_ratio(bbox1, bbox2):
    """ คำนวณอัตราส่วนการทับกัน """
    overlap_area = calculate_overlap_area(bbox1, bbox2)
    if overlap_area == 0:
        return 0
    
    # คำนวณพื้นที่ของ Bounding Box ที่เล็กกว่า
    area1 = (bbox1.Max.X - bbox1.Min.X) * (bbox1.Max.Y - bbox1.Min.Y)
    area2 = (bbox2.Max.X - bbox2.Min.X) * (bbox2.Max.Y - bbox2.Min.Y)
    min_area = min(area1, area2)
    
    if min_area > 0:
        return overlap_area / min_area
    
    return 0

# --------------------------------------------------------
# 🎯 เริ่มการประมวลผล
# --------------------------------------------------------
try:
    output.print_md("## 🔍 ตรวจสอบ Tags ที่ทับกัน")
    output.print_md("**View:** {}".format(view.Name))
    output.print_md("**View Scale:** {}".format(view.Scale if hasattr(view, 'Scale') else "N/A"))
    output.print_md("💡 **อนุญาตให้เส้น Leader ทับกันได้**")
    output.print_md("🎯 **ใช้ Bounding Box ของ Tag ที่ลดขนาดลง**")
    
    # ดึง Tags ทั้งหมด
    collector = DB.FilteredElementCollector(doc, view.Id)\
                  .OfClass(DB.IndependentTag)\
                  .WhereElementIsNotElementType()
    
    all_tags = list(collector)
    output.print_md("📊 **พบ Tags ทั้งหมด: {}**".format(len(all_tags)))
    
    if len(all_tags) == 0:
        output.print_md("ℹ️ **ไม่พบ Tags ใน View นี้**")
        script.exit()
    
    # วิธีที่ 1: ใช้ Bounding Box ที่ลดขนาดลง 70% (เหลือ 30%)
    output.print_md("## 🔧 วิธีที่ 1: Bounding Box ขนาดเล็ก (30% ของขนาดเดิม)")
    
    reduced_bounds = []
    for tag in all_tags:
        bounds = get_reduced_tag_bounds(tag, 0.3)
        if bounds:
            reduced_bounds.append({
                'tag': tag,
                'bounds': bounds,
                'text': get_tag_text(tag)
            })
    
    output.print_md("📐 **คำนวณขอบเขตขนาดเล็กได้: {} Tags**".format(len(reduced_bounds)))
    
    reduced_overlaps = []
    if len(reduced_bounds) >= 2:
        for i in range(len(reduced_bounds)):
            for j in range(i + 1, len(reduced_bounds)):
                if is_any_overlap(reduced_bounds[i]['bounds'], reduced_bounds[j]['bounds']):
                    overlap_area = calculate_overlap_area(
                        reduced_bounds[i]['bounds'], 
                        reduced_bounds[j]['bounds']
                    )
                    overlap_ratio = calculate_overlap_ratio(
                        reduced_bounds[i]['bounds'], 
                        reduced_bounds[j]['bounds']
                    )
                    
                    # คำนวณจุดศูนย์กลาง
                    b1 = reduced_bounds[i]['bounds']
                    b2 = reduced_bounds[j]['bounds']
                    center_x = (min(b1.Max.X, b2.Max.X) + max(b1.Min.X, b2.Min.X)) / 2
                    center_y = (min(b1.Max.Y, b2.Max.Y) + max(b1.Min.Y, b2.Min.Y)) / 2
                    
                    reduced_overlaps.append({
                        'tag1': reduced_bounds[i]['tag'],
                        'tag2': reduced_bounds[j]['tag'],
                        'text1': reduced_bounds[i]['text'],
                        'text2': reduced_bounds[j]['text'],
                        'center': (center_x, center_y),
                        'overlap_area': overlap_area,
                        'overlap_ratio': overlap_ratio,
                        'method': 'Reduced Bounds (30%)'
                    })
        
        output.print_md("📊 **พบการทับกันด้วยวิธีขอบเขตขนาดเล็ก: {} คู่**".format(len(reduced_overlaps)))
    
    # วิธีที่ 2: ใช้ Bounding Box ที่ลดขนาดลงมาก (เหลือ 15%)
    output.print_md("## 🔧 วิธีที่ 2: Bounding Box ขนาดเล็กมาก (15% ของขนาดเดิม)")
    
    very_small_bounds = []
    for tag in all_tags:
        bounds = get_very_small_tag_bounds(tag, 0.15)
        if bounds:
            very_small_bounds.append({
                'tag': tag,
                'bounds': bounds,
                'text': get_tag_text(tag)
            })
    
    output.print_md("📐 **คำนวณขอบเขตขนาดเล็กมากได้: {} Tags**".format(len(very_small_bounds)))
    
    very_small_overlaps = []
    if len(very_small_bounds) >= 2:
        for i in range(len(very_small_bounds)):
            for j in range(i + 1, len(very_small_bounds)):
                if is_any_overlap(very_small_bounds[i]['bounds'], very_small_bounds[j]['bounds']):
                    overlap_area = calculate_overlap_area(
                        very_small_bounds[i]['bounds'], 
                        very_small_bounds[j]['bounds']
                    )
                    overlap_ratio = calculate_overlap_ratio(
                        very_small_bounds[i]['bounds'], 
                        very_small_bounds[j]['bounds']
                    )
                    
                    # คำนวณจุดศูนย์กลาง
                    b1 = very_small_bounds[i]['bounds']
                    b2 = very_small_bounds[j]['bounds']
                    center_x = (min(b1.Max.X, b2.Max.X) + max(b1.Min.X, b2.Min.X)) / 2
                    center_y = (min(b1.Max.Y, b2.Max.Y) + max(b1.Min.Y, b2.Min.Y)) / 2
                    
                    very_small_overlaps.append({
                        'tag1': very_small_bounds[i]['tag'],
                        'tag2': very_small_bounds[j]['tag'],
                        'text1': very_small_bounds[i]['text'],
                        'text2': very_small_bounds[j]['text'],
                        'center': (center_x, center_y),
                        'overlap_area': overlap_area,
                        'overlap_ratio': overlap_ratio,
                        'method': 'Very Small Bounds (15%)'
                    })
        
        output.print_md("📊 **พบการทับกันด้วยวิธีขอบเขตขนาดเล็กมาก: {} คู่**".format(len(very_small_overlaps)))
    
    # รวมผลลัพธ์จากทุกวิธี
    all_overlaps = reduced_overlaps + very_small_overlaps
    
    # ลบคู่ที่ซ้ำกัน
    unique_overlaps = []
    seen_pairs = set()
    
    for overlap in all_overlaps:
        pair_id = tuple(sorted([overlap['tag1'].Id.IntegerValue, overlap['tag2'].Id.IntegerValue]))
        if pair_id not in seen_pairs:
            seen_pairs.add(pair_id)
            unique_overlaps.append(overlap)
    
    # แสดงผลรวม
    if not unique_overlaps:
        output.print_md("## ❌ ผลลัพธ์: ไม่พบ Tags ที่ทับกัน")
        output.print_md("💡 **ข้อสังเกต:**")
        output.print_md("- Tags ใน View นี้มีการจัดวางที่ดีมาก")
        output.print_md("- ไม่มีข้อความหรือพื้นที่ของ Tags ทับกันเลย")
        output.print_md("- แม้จะใช้ Bounding Box ที่ลดขนาดลงเหลือ 15% แล้วก็ตาม")
    else:
        output.print_md("## ✅ ผลลัพธ์: พบ Tags ที่ทับกัน {} คู่".format(len(unique_overlaps)))
        
        # แสดงสถิติตามวิธี
        output.print_md("### 📊 สถิติตามวิธีการตรวจจับ")
        output.print_md("- **ขอบเขตขนาดเล็ก (30%):** {} คู่".format(len(reduced_overlaps)))
        output.print_md("- **ขอบเขตขนาดเล็กมาก (15%):** {} คู่".format(len(very_small_overlaps)))
        
        # เรียงลำดับตามอัตราส่วนการทับกัน
        unique_overlaps.sort(key=lambda x: x['overlap_ratio'], reverse=True)
        
        # แสดงตารางผลลัพธ์
        table = []
        highlight_ids = []
        
        for i, overlap in enumerate(unique_overlaps):
            highlight_ids.extend([overlap['tag1'].Id, overlap['tag2'].Id])
            
            # รับชื่อประเภท Tag
            try:
                t1_type = overlap['tag1'].GetType().Name
            except:
                t1_type = "Unknown"
                
            try:
                t2_type = overlap['tag2'].GetType().Name
            except:
                t2_type = "Unknown"
            
            tag1_id_link = output.linkify(overlap['tag1'].Id)
            tag2_id_link = output.linkify(overlap['tag2'].Id)
            
            # แสดงข้อความ (ตัดให้สั้น)
            text1 = str(overlap['text1']) if overlap['text1'] else "N/A"
            text2 = str(overlap['text2']) if overlap['text2'] else "N/A"
            
            text1_display = text1[:20] + "..." if len(text1) > 20 else text1
            text2_display = text2[:20] + "..." if len(text2) > 20 else text2
            
            table.append([
                tag1_id_link,
                t1_type,
                text1_display,
                tag2_id_link,
                t2_type,
                text2_display,
                "({:.2f}, {:.2f})".format(overlap['center'][0], overlap['center'][1]),
                "{:.6f}".format(overlap['overlap_area']),
                "{:.1%}".format(overlap['overlap_ratio']),
                overlap['method']
            ])
        
        output.print_table(
            table_data=table,
            columns=[
                "Tag1 ID", "Tag1 Type", "Tag1 Text", 
                "Tag2 ID", "Tag2 Type", "Tag2 Text",
                "Overlap Center", "Overlap Area", "Overlap Ratio", "Method"
            ],
            title="รายการ Tags ที่ทับกัน"
        )
        
        if len(unique_overlaps) > 50:
            output.print_md("💡 **แสดงเฉพาะ 50 คู่แรกจากทั้งหมด {} คู่**".format(len(unique_overlaps)))
        
        # ไฮไลต์ใน View
        if highlight_ids:
            unique_ids = list(set(highlight_ids))
            output.print_md("🔦 **กำลังไฮไลต์ {} Tags ใน View**".format(len(unique_ids)))
            
            uidoc.Selection.SetElementIds(List[DB.ElementId](unique_ids))
        
        output.print_md("💡 **หมายเหตุ:**")
        output.print_md("- **อนุญาตให้เส้น Leader ทับกันได้**")
        output.print_md("- **ใช้ Bounding Box ของ Tag ที่ลดขนาดลง**")
        output.print_md("- **ขอบเขตขนาดเล็ก:** 30% ของขนาด Tag จริง")
        output.print_md("- **ขอบเขตขนาดเล็กมาก:** 15% ของขนาด Tag จริง")
        output.print_md("- พื้นที่ทับกันมีหน่วยเป็นตารางฟุต")

except Exception as e:
    output.print_md("❌ **เกิดข้อผิดพลาด:**")
    output.print_md("```{}```".format(str(e)))
    import traceback
    output.print_md("**รายละเอียดข้อผิดพลาด:**")
    output.print_md("```{}```".format(traceback.format_exc()))