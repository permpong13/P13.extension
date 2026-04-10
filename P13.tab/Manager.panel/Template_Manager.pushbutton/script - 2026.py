# -*- coding: utf-8 -*-
"""View Template Manager - Ultimate Pro (Fixed)"""
__title__ = "จัดการ View Template (All-in-One)"
__author__ = "เพิ่มพงษ์ & Gemini"

import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from pyrevit import forms, script

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

# --------------------------------------------------------
# ⚙️ CORE ENGINE (วิเคราะห์ข้อมูล)
# --------------------------------------------------------

def get_template_data():
    """วิเคราะห์ Template ทั้งหมดและแยกกลุ่มตามการใช้งานจริง"""
    # ดึง View ทั้งหมดที่เป็น Template
    all_tpls = [v for v in FilteredElementCollector(doc).OfClass(View) if v.IsTemplate]
    
    used_list = []   
    unused_list = [] 

    if not all_tpls:
        return [], []

    with forms.ProgressBar(title="กำลังวิเคราะห์การใช้งาน...", cancellable=True) as pb:
        for i, tpl in enumerate(all_tpls):
            if pb.cancelled: break
            pb.update_progress(i + 1, len(all_tpls))

            # ใช้ Filter หา View ที่ผูกกับ Template นี้
            rule = ParameterFilterRuleFactory.CreateEqualsRule(ElementId(BuiltInParameter.VIEW_TEMPLATE), tpl.Id)
            count = FilteredElementCollector(doc).OfClass(View).WherePasses(ElementParameterFilter(rule)).GetElementCount()

            if count > 0:
                used_list.append({'el': tpl, 'name': tpl.Name, 'count': count})
            else:
                unused_list.append({'el': tpl, 'name': tpl.Name})
    
    return used_list, unused_list

# --------------------------------------------------------
# 🎨 UI & DASHBOARD
# --------------------------------------------------------

def print_dashboard(used, unused):
    """สร้างตารางรายงานในหน้าต่าง Output"""
    # ลบ output.clear() ออกเพื่อป้องกัน Error
    output.print_md("# 🏢 View Template Pro Dashboard")
    output.insert_divider()
    
    header = ["สถานะการใช้งาน", "จำนวน", "ข้อแนะนำ"]
    table_data = [
        ["✅ ถูกใช้งานอยู่ (Active)", str(len(used)), "ข้อมูลสำคัญ ห้ามลบ"],
        ["⚠️ ไม่ได้ใช้งาน (Unused)", str(len(unused)), "ลบเพื่อลดขนาดไฟล์ได้"]
    ]
    output.print_table(table_data, columns=header)
    output.print_md("> **รวมทั้งหมด:** {} รายการ".format(len(used) + len(unused)))
    output.insert_divider()

# --------------------------------------------------------
# 🚀 MAIN COMMANDS
# --------------------------------------------------------

def run_manager():
    # 1. วิเคราะห์ข้อมูลก่อนเริ่ม
    used, unused = get_template_data()
    
    if not used and not unused:
        forms.alert("ไม่พบ View Template ในโปรเจคนี้", exitscript=True)

    # 2. โชว์หน้า Dashboard
    print_dashboard(used, unused)

    # 3. เมนูหลักแบบปุ่มกด
    ops = {
        "1. ดูรายละเอียด Template ที่ใช้งานอยู่": "SHOW_USED",
        "2. เลือกและลบ Template ที่ไม่ได้ใช้ (Cleanup)": "DELETE_UNUSED",
        "3. ปิดโปรแกรม": "EXIT"
    }
    
    choice = forms.CommandSwitchWindow.show(ops.keys(), title="เลือกการดำเนินการ")
    
    if not choice or ops[choice] == "EXIT":
        return

    # --- ACTION: SHOW USED ---
    if ops[choice] == "SHOW_USED":
        display = sorted(["🔹 {} (ใช้ใน {} views)".format(i['name'], i['count']) for i in used])
        forms.SelectFromList.show(display, title="รายการ Template ที่ใช้งานอยู่", button_name="รับทราบ")
        run_manager() # วนกลับไปหน้า Dashboard

    # --- ACTION: DELETE UNUSED ---
    elif ops[choice] == "DELETE_UNUSED":
        if not unused:
            forms.alert("ไม่มีรายการที่ไม่ได้ใช้งาน")
            run_manager()
            return

        # สร้าง List ให้เลือกติ๊กถูก
        unused_names = sorted([i['name'] for i in unused])
        selected = forms.SelectFromList.show(
            unused_names, 
            title="เลือก Template ที่ต้องการลบ (Unused Only)",
            button_name="ลบรายการที่เลือก",
            multiselect=True
        )

        if selected:
            confirm = forms.alert("ยืนยันการลบที่เลือก {} รายการ?".format(len(selected)), yes=True, no=True)
            if confirm:
                # ดึง Element ID จากชื่อที่เลือก
                to_del = [i['el'] for i in unused if i['name'] in selected]
                
                success = 0
                with Transaction(doc, "Cleanup Unused View Templates") as t:
                    t.Start()
                    for tpl_el in to_del:
                        try:
                            doc.Delete(tpl_el.Id)
                            success += 1
                        except: pass
                    t.Commit()
                
                forms.alert("ลบสำเร็จ {} รายการ".format(success))
                run_manager() # อัปเดตข้อมูลใหม่

if __name__ == "__main__":
    run_manager()