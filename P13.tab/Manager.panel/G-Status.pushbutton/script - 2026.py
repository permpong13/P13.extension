# -*- coding: utf-8 -*-
__title__ = 'G-Element\nStatus'
__author__ = 'เพิ่มพงษ์ ทวีกุล'
__doc__ = 'Updates g_Element Status parameter (0, 1, 2) for visible elements'

import clr
import System
import os
from System.Collections.Generic import List

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
import Autodesk.Revit.UI as UI
from System.Windows.Forms import *
from System.Drawing import *

# --------------------------
# TARGET CATEGORIES
# --------------------------
TARGET_CATEGORIES = {
    "Walls": BuiltInCategory.OST_Walls,
    "Floors": BuiltInCategory.OST_Floors,
    "Ceilings": BuiltInCategory.OST_Ceilings,
    "Roofs": BuiltInCategory.OST_Roofs,
    "Doors": BuiltInCategory.OST_Doors,
    "Windows": BuiltInCategory.OST_Windows,
    "Columns": BuiltInCategory.OST_StructuralColumns,
    "Beams": BuiltInCategory.OST_StructuralFraming,
    "Foundations": BuiltInCategory.OST_StructuralFoundation,
    "Stairs": BuiltInCategory.OST_Stairs,
    "Ramps": BuiltInCategory.OST_Ramps,
    "Railings": BuiltInCategory.OST_Railings,
    "Furniture": BuiltInCategory.OST_Furniture,
    "Generic Models": BuiltInCategory.OST_GenericModel,
    "Mechanical Equipment": BuiltInCategory.OST_MechanicalEquipment,
    "Electrical Equipment": BuiltInCategory.OST_ElectricalEquipment,
    "Lighting Fixtures": BuiltInCategory.OST_LightingFixtures,
    "Plumbing Fixtures": BuiltInCategory.OST_PlumbingFixtures,
    "Sprinklers": BuiltInCategory.OST_Sprinklers,
    "Ducts": BuiltInCategory.OST_DuctCurves,
    "Pipes": BuiltInCategory.OST_PipeCurves,
    "Cable Trays": BuiltInCategory.OST_CableTray,
    "Conduits": BuiltInCategory.OST_Conduit,
    "Duct Terminals": BuiltInCategory.OST_DuctTerminal,
    "Site": BuiltInCategory.OST_Site,
    "Topography": BuiltInCategory.OST_Topography,
    "Planting": BuiltInCategory.OST_Planting
}

# --------------------------
# Helper for Revit 2026 (64-bit ID)
# --------------------------
def get_id_int(element_id):
    if hasattr(element_id, "Value"):
        return element_id.Value
    return element_id.IntegerValue

# --------------------------
# Form: Select rows
# --------------------------
class ElementStatusForm(Form):
    def __init__(self, categories_with_elements):
        self.categories_with_elements = categories_with_elements
        self.InitializeComponent()

    def InitializeComponent(self):
        self.Text = "Update Element Status (0, 1, 2)"
        self.Width = 800
        self.Height = 650
        self.StartPosition = FormStartPosition.CenterScreen

        layout = TableLayoutPanel()
        layout.Dock = DockStyle.Fill
        layout.RowCount = 5
        layout.ColumnCount = 1
        layout.RowStyles.Add(RowStyle(SizeType.Absolute, 40))
        layout.RowStyles.Add(RowStyle(SizeType.Absolute, 30))
        layout.RowStyles.Add(RowStyle(SizeType.Percent, 100))
        layout.RowStyles.Add(RowStyle(SizeType.Absolute, 40))
        layout.RowStyles.Add(RowStyle(SizeType.Absolute, 60))

        lbl_title = Label()
        lbl_title.Text = "Select groups to update. You can manually edit 'New Status' column."
        lbl_title.Font = Font(lbl_title.Font, FontStyle.Bold)
        lbl_title.Dock = DockStyle.Fill
        layout.Controls.Add(lbl_title, 0, 0)

        lbl_desc = Label()
        lbl_desc.Text = "Suggested: 0->1, 1->2, 2->1. Double click 'New Status' to change to any value."
        lbl_desc.ForeColor = Color.DarkBlue
        lbl_desc.Dock = DockStyle.Fill
        layout.Controls.Add(lbl_desc, 0, 1)

        self.dgv = DataGridView()
        self.dgv.Dock = DockStyle.Fill
        self.dgv.RowHeadersVisible = False
        self.dgv.AllowUserToAddRows = False
        self.dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        col_sel = DataGridViewCheckBoxColumn()
        col_sel.HeaderText = "Select"
        col_sel.Name = "Selected"
        col_sel.Width = 60
        self.dgv.Columns.Add(col_sel)

        self.dgv.Columns.Add("Category", "Category")
        self.dgv.Columns["Category"].ReadOnly = True
        self.dgv.Columns.Add("Status", "Current Value")
        self.dgv.Columns["Status"].ReadOnly = True
        self.dgv.Columns.Add("Count", "Element Count")
        self.dgv.Columns["Count"].ReadOnly = True
        
        col_new = DataGridViewTextBoxColumn()
        col_new.HeaderText = "New Status"
        col_new.Name = "NewStatus"
        col_new.Width = 100
        # อนุญาตให้ผู้ใช้แก้ไขค่าในช่องนี้ได้เอง
        col_new.ReadOnly = False 
        self.dgv.Columns.Add(col_new)

        for cname in sorted(self.categories_with_elements):
            elements = self.categories_with_elements[cname]["elements"]
            value_groups = self.split_by_status(elements)
            for val_key, el_list in value_groups.items():
                if val_key is None:
                    display_val, default_new = "Empty", 0
                elif str(val_key) == "0":
                    display_val, default_new = "0", 1
                elif str(val_key) == "1":
                    display_val, default_new = "1", 2 # 1 -> 2 ตามโจทย์
                elif str(val_key) == "2":
                    display_val, default_new = "2", 1 # 2 -> 1 ตามโจทย์
                else:
                    display_val, default_new = str(val_key), 1
                
                self.dgv.Rows.Add(False, cname, display_val, len(el_list), default_new)

        layout.Controls.Add(self.dgv, 0, 2)

        btn_panel = FlowLayoutPanel()
        btn_panel.Dock = DockStyle.Fill
        btn_select_all = Button(Text="Select All", Width=120)
        btn_unselect_all = Button(Text="Unselect All", Width=120)
        btn_select_all.Click += self.select_all
        btn_unselect_all.Click += self.unselect_all
        btn_panel.Controls.Add(btn_select_all)
        btn_panel.Controls.Add(btn_unselect_all)
        layout.Controls.Add(btn_panel, 0, 3)

        act_panel = FlowLayoutPanel()
        act_panel.Dock = DockStyle.Fill
        act_panel.FlowDirection = FlowDirection.RightToLeft
        btn_ok = Button(Text="Update Selected", Width=150, BackColor=Color.LightGreen)
        btn_ok.Click += self.ok_click
        btn_cancel = Button(Text="Cancel", Width=80)
        btn_cancel.Click += self.cancel_click
        act_panel.Controls.Add(btn_ok)
        act_panel.Controls.Add(btn_cancel)
        layout.Controls.Add(act_panel, 0, 4)

        self.Controls.Add(layout)

    def split_by_status(self, elements):
        groups = {}
        for el in elements:
            p = el.LookupParameter("g_Element Status")
            key = None
            if p and p.HasValue:
                if p.StorageType == StorageType.Integer:
                    key = p.AsInteger()
                elif p.StorageType == StorageType.String:
                    val = p.AsString()
                    try: key = int(val)
                    except: key = val
            if key not in groups: groups[key] = []
            groups[key].append(el)
        return groups

    def select_all(self, s, e):
        for r in self.dgv.Rows: r.Cells["Selected"].Value = True
    def unselect_all(self, s, e):
        for r in self.dgv.Rows: r.Cells["Selected"].Value = False
    def ok_click(self, s, e):
        self.DialogResult = DialogResult.OK
        self.Close()
    def cancel_click(self, s, e):
        self.DialogResult = DialogResult.Cancel
        self.Close()

    def get_selected_rows(self):
        selected = []
        for r in self.dgv.Rows:
            if r.Cells["Selected"].Value:
                try:
                    # อ่านค่าจากช่อง NewStatus ที่ผู้ใช้อาจจะแก้เป็นเลข 2 หรือเลขอื่นๆ
                    new_val_raw = r.Cells["NewStatus"].Value
                    new_val = int(new_val_raw)
                except:
                    new_val = 1
                selected.append({
                    "Category": str(r.Cells["Category"].Value),
                    "CurrentValue": r.Cells["Status"].Value,
                    "NewValue": new_val
                })
        return selected

# --------------------------
# Update Logic
# --------------------------
def update_elements_with_progress(doc, selected_rows, category_elements):
    total_elements = sum(category_elements[row["Category"]]["count"] for row in selected_rows)
    pf = ProgressForm(total_elements)
    pf.Show()

    tx = Transaction(doc, "Update Element Status")
    tx.Start()
    updated = processed = 0
    try:
        for row in selected_rows:
            cname, cur_val, new_val = row["Category"], row["CurrentValue"], row["NewValue"]
            elements = category_elements[cname]["elements"]
            
            for el in elements:
                p = el.LookupParameter("g_Element Status")
                if p and not p.IsReadOnly:
                    # ตรวจสอบค่าปัจจุบันเพื่อให้แน่ใจว่าตรงกับกลุ่มที่เลือก
                    val = None
                    if p.HasValue:
                        if p.StorageType == StorageType.Integer: val = p.AsInteger()
                        else:
                            try: val = int(p.AsString())
                            except: val = p.AsString()
                    
                    is_match = False
                    if cur_val == "Empty" and val is None: is_match = True
                    elif str(val) == str(cur_val): is_match = True

                    if is_match:
                        if p.StorageType == StorageType.Integer: p.Set(new_val)
                        else: p.Set(str(new_val))
                        updated += 1
                
                processed += 1
                pf.progress.Value = processed
                pf.label.Text = "Updating {0}... ({1}/{2})".format(cname, processed, total_elements)
                pf.Refresh()
        tx.Commit()
        pf.Close()
        return updated
    except:
        tx.RollBack()
        pf.Close()
        raise

# --------------------------
# Main
# --------------------------
def main():
    doc = __revit__.ActiveUIDocument.Document
    category_elements = get_elements_by_category(doc, doc.ActiveView)
    if not category_elements: return

    form = ElementStatusForm(category_elements)
    if form.ShowDialog() == DialogResult.OK:
        selected = form.get_selected_rows()
        if selected:
            count = update_elements_with_progress(doc, selected, category_elements)
            UI.TaskDialog.Show("Complete", "Updated {0} elements successfully.".format(count))

def get_elements_by_category(doc, view):
    all_elements = FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().ToElements()
    out = {}
    for cname, bic in TARGET_CATEGORIES.items():
        bic_int = int(bic)
        elems = [el for el in all_elements if el.Category and get_id_int(el.Category.Id) == bic_int]
        if elems: out[cname] = {"elements": elems, "count": len(elems)}
    return out

class ProgressForm(Form):
    def __init__(self, total):
        self.Width, self.Height = 500, 150
        self.Text = "Processing..."
        self.StartPosition = FormStartPosition.CenterScreen
        self.label = Label(Text="Starting...", Dock=DockStyle.Top, Height=30)
        self.progress = ProgressBar(Dock=DockStyle.Fill, Minimum=0, Maximum=total)
        self.Controls.Add(self.progress)
        self.Controls.Add(self.label)

if __name__ == "__main__":
    main()