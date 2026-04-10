# -*- coding: utf-8 -*-
"""View Template Manager (Apply + Remove)"""
__title__ = "จัดการ View Template"
__author__ = "เพิ่มพงษ์"

import clr
import sys
import os

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List

from pyrevit import forms
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# --------------------------------------------------------
#  ฟังก์ชันร่วม
# --------------------------------------------------------
def get_all_templates():
    collector = FilteredElementCollector(doc)
    views = collector.OfClass(View).ToElements()

    result = {}
    for v in views:
        if v and v.IsValidObject and v.IsTemplate:
            result[v.Name] = v

    return result


# ========================================================
#  โหมด 1 : Apply Template
# ========================================================
def get_selected_views():
    ids = uidoc.Selection.GetElementIds()
    result = []

    for i in ids:
        e = doc.GetElement(i)
        if isinstance(e, View) and not e.IsTemplate:
            result.append(e)

    if not result:
        collector = FilteredElementCollector(doc)
        for v in collector.OfClass(View).ToElements():
            if not v.IsTemplate:
                result.append(v)

    return result


def apply_template_mode():
    templates = get_all_templates()
    if not templates:
        forms.alert("ไม่พบ View Template ในโปรเจค", exitscript=True)
        return

    template_names = sorted(templates.keys())

    selected_tpl_name = forms.SelectFromList.show(
        template_names,
        title="เลือก View Template",
        button_name="เลือก Template"
    )

    if not selected_tpl_name:
        return

    tpl = templates[selected_tpl_name]
    views = get_selected_views()

    if not views:
        forms.alert("ไม่พบ Views ที่สามารถใช้ Template ได้", exitscript=True)
        return

    display_map = {}
    display_list = []

    for v in views:
        current_name = "None"

        if v.ViewTemplateId != ElementId.InvalidElementId:
            t = doc.GetElement(v.ViewTemplateId)
            if t:
                current_name = t.Name

        label = "{} (ปัจจุบัน: {})".format(v.Name, current_name)
        display_map[label] = v
        display_list.append(label)

    selected_views = forms.SelectFromList.show(
        sorted(display_list),
        title="เลือก Views ที่ต้องการ Apply (ทั้งหมด {} view)".format(len(display_list)),
        button_name="Apply",
        multiselect=True
    )

    if not selected_views:
        return

    t = Transaction(doc, "Apply View Template")
    t.Start()

    success = 0
    fail = 0
    logs = []

    for label in selected_views:
        v = display_map[label]
        try:
            v.ViewTemplateId = tpl.Id
            success += 1
            logs.append("✔ {} : Apply สำเร็จ".format(v.Name))
        except Exception as e:
            fail += 1
            logs.append("✘ {} : {}".format(v.Name, str(e)))

    t.Commit()

    msg = "ผลลัพธ์การ Apply Template '{}'\n".format(selected_tpl_name)
    msg += "สำเร็จ: {}\nล้มเหลว: {}\n\nรายละเอียด:\n".format(success, fail)
    msg += "\n".join(logs)

    forms.alert(msg, title="Apply เสร็จสิ้น")


# ========================================================
#  โหมด 2 : ลบ Template
# ========================================================
def get_views_using_template(tpl):
    collector = FilteredElementCollector(doc)
    views = collector.OfClass(View).ToElements()

    result = []
    for v in views:
        if not v.IsTemplate and v.ViewTemplateId == tpl.Id:
            result.append(v)
    return result


def remove_template_mode():
    templates = get_all_templates()
    if not templates:
        forms.alert("ไม่พบ Template", exitscript=True)
        return

    options = []
    option_map = {}

    for name, tpl in templates.items():
        used_views = get_views_using_template(tpl)
        count = len(used_views)

        label = "{} (ถูกใช้ {} view)".format(name, count)
        options.append(label)

        option_map[label] = {
            "template": tpl,
            "count": count,
            "views": used_views
        }

    selected = forms.SelectFromList.show(
        sorted(options),
        title="เลือก Template ที่จะลบ",
        button_name="ลบ",
        multiselect=True
    )

    if not selected:
        return

    preview = ["รายการ Template ที่เลือกจะลบ:\n"]

    for opt in selected:
        info = option_map[opt]
        preview.append("- {} ({} views ใช้งานอยู่)".format(info["template"].Name, info["count"]))

    ans = forms.alert(
        "\n".join(preview),
        title="ยืนยันการลบ",
        options=["ยืนยัน", "ยกเลิก"]
    )

    if ans != "ยืนยัน":
        return

    t = Transaction(doc, "Delete Templates")
    t.Start()

    success = 0
    fail = 0
    logs = []

    for opt in selected:
        info = option_map[opt]
        tpl = info["template"]

        try:
            doc.Delete(tpl.Id)
            success += 1
            logs.append("✔ ลบ '{}' แล้ว".format(tpl.Name))
        except Exception as e:
            fail += 1
            logs.append("✘ ลบ '{}' ไม่สำเร็จ: {}".format(tpl.Name, str(e)))

    t.Commit()

    msg = "ลบสำเร็จ {} รายการ / ล้มเหลว {}\n\n".format(success, fail)
    msg += "\n".join(logs)

    forms.alert(msg, title="การลบเสร็จสมบูรณ์")


# ========================================================
#  MAIN MENU
# ========================================================
def main():
    choice = forms.SelectFromList.show(
        ["Apply Template ให้ Views", "ลบ View Templates"],
        title="จัดการ View Template",
        button_name="เริ่มทำงาน"
    )

    if not choice:
        return

    if choice == "Apply Template ให้ Views":
        apply_template_mode()
    else:
        remove_template_mode()


if __name__ == "__main__":
    main()
