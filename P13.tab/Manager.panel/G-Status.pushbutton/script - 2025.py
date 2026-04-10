# -*- coding: utf-8 -*-
__title__ = 'G-Element\nStatus'
__author__ = 'เพิ่มพงษ์ ทวีกุล'
__doc__ = 'Updates g_Element Status parameter for visible elements in current view'

import clr
import System
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
# Form: Select rows
# --------------------------
class ElementStatusForm(Form):
    def __init__(self, categories_with_elements):
        self.categories_with_elements = categories_with_elements
        self.InitializeComponent()

    def InitializeComponent(self):
        self.Text = "Update Element Status"
        self.Width = 780
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
        lbl_title.Text = "Select category-value groups to update g_Element Status:"
        lbl_title.Font = Font(lbl_title.Font, FontStyle.Bold)
        lbl_title.Dock = DockStyle.Fill
        layout.Controls.Add(lbl_title, 0, 0)

        lbl_desc = Label()
        lbl_desc.Text = "Empty → 0, 0 → 1, 1 → 1, Other → 1. Select rows to update."
        lbl_desc.ForeColor = Color.DarkBlue
        lbl_desc.Dock = DockStyle.Fill
        layout.Controls.Add(lbl_desc, 0, 1)

        # DataGrid
        self.dgv = DataGridView()
        self.dgv.Dock = DockStyle.Fill
        self.dgv.RowHeadersVisible = False
        self.dgv.AllowUserToAddRows = False
        self.dgv.AllowUserToDeleteRows = False
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
        self.dgv.Columns["Status"].Width = 120

        self.dgv.Columns.Add("Count", "Element Count")
        self.dgv.Columns["Count"].ReadOnly = True
        self.dgv.Columns["Count"].Width = 80

        col_new = DataGridViewTextBoxColumn()
        col_new.HeaderText = "New Status"
        col_new.Name = "NewStatus"
        col_new.Width = 80
        self.dgv.Columns.Add(col_new)

        # Populate rows with default New Status
        for cname in sorted(self.categories_with_elements):
            elements = self.categories_with_elements[cname]["elements"]
            value_groups = self.split_by_status(elements)
            for val_key, el_list in value_groups.items():
                display_name = cname
                if val_key is None:
                    display_val = "Empty"
                    default_new = 0
                elif val_key == 0:
                    display_val = "0"
                    default_new = 1
                elif val_key == 1:
                    display_val = "1"
                    default_new = 1
                else:
                    display_val = str(val_key)
                    default_new = 1
                self.dgv.Rows.Add(False, display_name, display_val, len(el_list), default_new)

        layout.Controls.Add(self.dgv, 0, 2)

        # Buttons
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
        btn_ok = Button(Text="Update (0/1)", Width=120)
        btn_ok.Click += self.ok_click
        btn_cancel = Button(Text="Cancel", Width=80)
        btn_cancel.Click += self.cancel_click
        act_panel.Controls.Add(btn_ok)
        act_panel.Controls.Add(btn_cancel)
        layout.Controls.Add(act_panel, 0, 4)

        self.Controls.Add(layout)
        self.DialogResult = DialogResult.Cancel

    def split_by_status(self, elements):
        groups = {}
        for el in elements:
            p = el.LookupParameter("g_Element Status")
            key = None
            if p:
                if p.StorageType == StorageType.Integer:
                    if p.HasValue:
                        key = p.AsInteger()
                elif p.StorageType == StorageType.String:
                    val = p.AsString()
                    if val is not None and val != "":
                        try:
                            key = int(val)
                        except:
                            key = val
            if key not in groups:
                groups[key] = []
            groups[key].append(el)
        return groups

    def select_all(self, s, e):
        for r in self.dgv.Rows:
            r.Cells["Selected"].Value = True

    def unselect_all(self, s, e):
        for r in self.dgv.Rows:
            r.Cells["Selected"].Value = False

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
                cname = str(r.Cells["Category"].Value)
                cur_val = r.Cells["Status"].Value
                try:
                    new_val = int(r.Cells["NewStatus"].Value)
                except:
                    new_val = 1
                selected.append({"Category": cname, "CurrentValue": cur_val, "NewValue": new_val})
        return selected

# --------------------------
# Progress Form
# --------------------------
class ProgressForm(Form):
    def __init__(self, total):
        self.Width = 500
        self.Height = 150
        self.StartPosition = FormStartPosition.CenterScreen
        self.Text = "Updating Elements..."
        self.label = Label(Text="Starting...", Dock=DockStyle.Top, Height=30)
        self.progress = ProgressBar(Dock=DockStyle.Fill, Minimum=0, Maximum=total)
        layout = TableLayoutPanel(Dock=DockStyle.Fill)
        layout.RowCount = 2
        layout.ColumnCount = 1
        layout.RowStyles.Add(RowStyle(SizeType.Absolute, 30))
        layout.RowStyles.Add(RowStyle(SizeType.Percent, 100))
        layout.Controls.Add(self.label, 0, 0)
        layout.Controls.Add(self.progress, 0, 1)
        self.Controls.Add(layout)

# --------------------------
# Collect visible elements
# --------------------------
def get_visible_elements(doc, view):
    col = FilteredElementCollector(doc, view.Id)
    return col.WhereElementIsNotElementType().ToElements()

def get_elements_by_category(doc, view):
    all_elements = get_visible_elements(doc, view)
    out = {}
    for cname, bic in TARGET_CATEGORIES.items():
        bic_int = int(bic)
        elems = [el for el in all_elements if el.Category and el.Category.Id.IntegerValue == bic_int]
        if elems:
            out[cname] = {"elements": elems, "count": len(elems)}
    return out

# --------------------------
# Update with progress (IronPython compatible)
# --------------------------
def update_elements_with_progress(doc, selected_rows, category_elements):
    total_elements = 0
    for row in selected_rows:
        cname = row["Category"]
        total_elements += category_elements[cname]["count"]

    pf = ProgressForm(total_elements)
    pf.Show()
    pf.Refresh()

    tx = Transaction(doc, "Update Element Status")
    tx.Start()
    updated = zero_cnt = one_cnt = err_cnt = 0
    processed = 0
    try:
        for row in selected_rows:
            cname = row["Category"]
            cur_val = row["CurrentValue"]
            new_val = row["NewValue"]
            elements = category_elements[cname]["elements"]
            matching_elements = []
            for el in elements:
                p = el.LookupParameter("g_Element Status")
                val = None
                if p:
                    if p.StorageType == StorageType.Integer:
                        val = p.AsInteger() if p.HasValue else None
                    elif p.StorageType == StorageType.String:
                        s = p.AsString()
                        if s is None or s == "":
                            val = None
                        else:
                            try:
                                val = int(s)
                            except:
                                val = s
                if cur_val == "Empty" and val is None:
                    matching_elements.append(el)
                elif str(val) == str(cur_val):
                    matching_elements.append(el)

            for el in matching_elements:
                try:
                    p = el.LookupParameter("g_Element Status")
                    if p and not p.IsReadOnly:
                        if p.StorageType == StorageType.Integer:
                            p.Set(new_val)
                            if new_val == 0:
                                zero_cnt += 1
                            elif new_val == 1:
                                one_cnt += 1
                            updated += 1
                        elif p.StorageType == StorageType.String:
                            p.Set(str(new_val))
                            if new_val == 0:
                                zero_cnt += 1
                            elif new_val == 1:
                                one_cnt += 1
                            updated += 1
                except:
                    err_cnt += 1
                processed += 1
                pf.progress.Value = processed
                pf.label.Text = "Updating {0} (Current: {1}) {2}/{3}".format(cname, cur_val, processed, total_elements)
                pf.Refresh()

        tx.Commit()
        pf.Close()
        return updated, zero_cnt, one_cnt, err_cnt
    except:
        tx.RollBack()
        pf.Close()
        raise

# --------------------------
# MAIN
# --------------------------
def main():
    doc = __revit__.ActiveUIDocument.Document
    view = doc.ActiveView
    category_elements = get_elements_by_category(doc, view)
    if not category_elements:
        UI.TaskDialog.Show("Info", "No elements found in this view.")
        return

    form = ElementStatusForm(category_elements)
    if form.ShowDialog() != DialogResult.OK:
        return

    selected_rows = form.get_selected_rows()
    if not selected_rows:
        UI.TaskDialog.Show("Warning", "No rows selected.")
        return

    updated, zero_cnt, one_cnt, err_cnt = update_elements_with_progress(doc, selected_rows, category_elements)

    msg = "Update completed:\nTotal elements updated: {0}\nSet to 0: {1}\nSet to 1: {2}".format(updated, zero_cnt, one_cnt)
    if err_cnt:
        msg += "\nErrors: {0}".format(err_cnt)
    UI.TaskDialog.Show("Complete", msg)

if __name__ == "__main__":
    main()
