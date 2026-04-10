# -*- coding: utf-8 -*-
__title__ = 'G-Element\nStatus'
__author__ = 'เพิ่มพงษ์ ทวีกุล'
__doc__ = 'Updates g_Element Status parameter (0, 0.1, 0.5, 1, 2) for visible elements grouped by Workset'

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
# TARGET CATEGORIES (กรองเฉพาะวัตถุหลัก)
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
# Helper for Workset
# --------------------------
def get_workset_name(doc, elem):
    try:
        if doc.IsWorkshared:
            ws_id = elem.WorksetId
            if ws_id != WorksetId.InvalidWorksetId:
                ws = doc.GetWorksetTable().GetWorkset(ws_id)
                if ws:
                    return ws.Name
    except Exception:
        pass
    return "Non-Shared"

# --------------------------
# Form: Select rows
# --------------------------
class ElementStatusForm(Form):
    def __init__(self, grouped_elements):
        self.grouped_elements = grouped_elements
        self.InitializeComponent()

    def InitializeComponent(self):
        self.Text = "Update Element Status by Workset"
        self.Width = 850 
        self.Height = 650
        self.StartPosition = FormStartPosition.CenterScreen
        self.Font = Font("Segoe UI", 9.5)
        self.BackColor = Color.White

        layout = TableLayoutPanel()
        layout.Dock = DockStyle.Fill
        layout.RowCount = 5
        layout.ColumnCount = 1
        layout.RowStyles.Add(RowStyle(SizeType.Absolute, 45))
        layout.RowStyles.Add(RowStyle(SizeType.Absolute, 35))
        layout.RowStyles.Add(RowStyle(SizeType.Percent, 100))
        layout.RowStyles.Add(RowStyle(SizeType.Absolute, 50))
        layout.RowStyles.Add(RowStyle(SizeType.Absolute, 60))
        layout.Padding = Padding(10)

        # Title Label
        lbl_title = Label()
        lbl_title.Text = "Select groups to update. You can manually edit the 'New Status' column."
        lbl_title.Font = Font("Segoe UI", 11, FontStyle.Bold)
        lbl_title.ForeColor = Color.FromArgb(40, 40, 40)
        lbl_title.Dock = DockStyle.Fill
        lbl_title.TextAlign = ContentAlignment.BottomLeft
        layout.Controls.Add(lbl_title, 0, 0)

        # Description Label
        lbl_desc = Label()
        lbl_desc.Text = "Sequence: Empty ➔ 0 ➔ 0.1 ➔ 0.5 ➔ 1 ➔ 2 | Double-click 'New Status' cell to customize value."
        lbl_desc.ForeColor = Color.DimGray
        lbl_desc.Dock = DockStyle.Fill
        lbl_desc.TextAlign = ContentAlignment.MiddleLeft
        layout.Controls.Add(lbl_desc, 0, 1)

        # DataGridView Setup
        self.dgv = DataGridView()
        self.dgv.Dock = DockStyle.Fill
        self.dgv.RowHeadersVisible = False
        self.dgv.AllowUserToAddRows = False
        self.dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        self.dgv.BackgroundColor = Color.White
        self.dgv.BorderStyle = BorderStyle.FixedSingle
        self.dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        self.dgv.GridColor = Color.LightGray
        self.dgv.RowTemplate.Height = 35
        self.dgv.EnableHeadersVisualStyles = False
        
        # Header Styling
        header_style = DataGridViewCellStyle()
        header_style.BackColor = Color.FromArgb(240, 240, 240)
        header_style.Font = Font("Segoe UI", 10, FontStyle.Bold)
        header_style.Alignment = DataGridViewContentAlignment.MiddleCenter
        self.dgv.ColumnHeadersDefaultCellStyle = header_style
        self.dgv.ColumnHeadersHeight = 40

        # Columns Configuration
        col_sel = DataGridViewCheckBoxColumn()
        col_sel.HeaderText = "Select"
        col_sel.Name = "Selected"
        col_sel.Width = 60
        self.dgv.Columns.Add(col_sel)

        col_ws = DataGridViewTextBoxColumn()
        col_ws.HeaderText = "Workset"
        col_ws.Name = "Workset"
        col_ws.ReadOnly = True
        col_ws.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        self.dgv.Columns.Add(col_ws)

        col_status = DataGridViewTextBoxColumn()
        col_status.HeaderText = "Current Status"
        col_status.Name = "Status"
        col_status.ReadOnly = True
        col_status.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        col_status.Width = 120
        self.dgv.Columns.Add(col_status)
        
        col_count = DataGridViewTextBoxColumn()
        col_count.HeaderText = "Element Count"
        col_count.Name = "Count"
        col_count.ReadOnly = True
        col_count.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        col_count.Width = 120
        self.dgv.Columns.Add(col_count)
        
        col_new = DataGridViewTextBoxColumn()
        col_new.HeaderText = "✎ New Status"
        col_new.Name = "NewStatus"
        col_new.Width = 120
        col_new.ReadOnly = False 
        col_new.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        col_new.DefaultCellStyle.BackColor = Color.AliceBlue
        col_new.DefaultCellStyle.ForeColor = Color.DarkBlue
        col_new.DefaultCellStyle.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.dgv.Columns.Add(col_new)

        # Populate Rows with Status Color Coding
        for ws_name in sorted(self.grouped_elements.keys()):
            elements = self.grouped_elements[ws_name]["elements"]
            value_groups = self.split_by_status(elements)
            
            for val_key, el_list in value_groups.items():
                val_str = str(val_key).strip() if val_key is not None else "Empty"
                
                # ลำดับสถานะ: Empty -> 0 -> 0.1 -> 0.5 -> 1 -> 2
                if val_str == "Empty":
                    display_val, default_new = "Empty", "0"
                    bg_color = Color.WhiteSmoke
                elif val_str == "0" or val_str == "0.0":
                    display_val, default_new = "0", "0.1"
                    bg_color = Color.WhiteSmoke
                elif val_str == "0.1":
                    display_val, default_new = "0.1", "0.5"
                    bg_color = Color.AliceBlue # ฟ้าอ่อน
                elif val_str == "0.5":
                    display_val, default_new = "0.5", "1"
                    bg_color = Color.MistyRose # ส้ม/ชมพูอ่อน
                elif val_str == "1" or val_str == "1.0":
                    display_val, default_new = "1", "2"
                    bg_color = Color.FromArgb(255, 250, 205) # เหลืองอ่อน
                elif val_str == "2" or val_str == "2.0":
                    display_val, default_new = "2", "0" # เสร็จแล้ว ให้วนกลับเป็น 0 ได้หากต้องการรีเซ็ต
                    bg_color = Color.FromArgb(240, 255, 240) # เขียวอ่อน
                else:
                    display_val, default_new = val_str, "0"
                    bg_color = Color.White
                
                idx = self.dgv.Rows.Add(False, ws_name, display_val, len(el_list), default_new)
                self.dgv.Rows[idx].DefaultCellStyle.BackColor = bg_color

        layout.Controls.Add(self.dgv, 0, 2)

        # Selection Buttons Panel
        btn_panel = FlowLayoutPanel()
        btn_panel.Dock = DockStyle.Fill
        btn_panel.Padding = Padding(0, 10, 0, 0)
        
        btn_select_all = Button(Text="Select All", Width=120, Height=35)
        btn_select_all.FlatStyle = FlatStyle.System
        btn_select_all.Click += self.select_all
        
        btn_unselect_all = Button(Text="Unselect All", Width=120, Height=35)
        btn_unselect_all.FlatStyle = FlatStyle.System
        btn_unselect_all.Click += self.unselect_all
        
        btn_panel.Controls.Add(btn_select_all)
        btn_panel.Controls.Add(btn_unselect_all)
        layout.Controls.Add(btn_panel, 0, 3)

        # Action Buttons Panel
        act_panel = FlowLayoutPanel()
        act_panel.Dock = DockStyle.Fill
        act_panel.FlowDirection = FlowDirection.RightToLeft
        act_panel.Padding = Padding(0, 10, 0, 0)
        
        btn_ok = Button(Text="Update Selected", Width=160, Height=40)
        btn_ok.BackColor = Color.FromArgb(76, 175, 80) # Material Green
        btn_ok.ForeColor = Color.White
        btn_ok.Font = Font("Segoe UI", 10, FontStyle.Bold)
        btn_ok.FlatStyle = FlatStyle.Flat
        btn_ok.FlatAppearance.BorderSize = 0
        btn_ok.Click += self.ok_click
        
        btn_cancel = Button(Text="Cancel", Width=100, Height=40)
        btn_cancel.BackColor = Color.FromArgb(220, 220, 220)
        btn_cancel.FlatStyle = FlatStyle.Flat
        btn_cancel.FlatAppearance.BorderSize = 0
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
                    key = str(p.AsInteger())
                elif p.StorageType == StorageType.Double:
                    key = str(p.AsDouble())
                elif p.StorageType == StorageType.String:
                    key = p.AsString()
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
                    new_val_raw = r.Cells["NewStatus"].Value
                    new_val = str(new_val_raw).strip()
                except:
                    new_val = "1"
                
                group_key = str(r.Cells["Workset"].Value)
                
                selected.append({
                    "GroupKey": group_key,
                    "CurrentValue": r.Cells["Status"].Value,
                    "NewValue": new_val,
                    "Count": int(r.Cells["Count"].Value)
                })
        return selected

# --------------------------
# Update Logic
# --------------------------
def update_elements_with_progress(doc, selected_rows, grouped_elements):
    total_elements = sum(row["Count"] for row in selected_rows)
    pf = ProgressForm(total_elements)
    pf.Show()

    tx = Transaction(doc, "Update Element Status by Workset")
    tx.Start()
    updated = processed = 0
    
    try:
        for row in selected_rows:
            ws_name = row["GroupKey"]
            cur_val = row["CurrentValue"]
            new_val = row["NewValue"]
            
            elements = grouped_elements[ws_name]["elements"]
            
            for el in elements:
                p = el.LookupParameter("g_Element Status")
                if p and not p.IsReadOnly:
                    # แปลงสถานะปัจจุบันเพื่อตรวจสอบว่าตรงกับกลุ่มหรือไม่
                    val_str = None
                    if p.HasValue:
                        if p.StorageType == StorageType.Integer: val_str = str(p.AsInteger())
                        elif p.StorageType == StorageType.Double: val_str = str(p.AsDouble())
                        else: val_str = p.AsString()
                    
                    is_match = False
                    if cur_val == "Empty" and val_str is None: is_match = True
                    elif str(val_str).strip() == str(cur_val).strip(): is_match = True
                    # ช่วยจับกรณีเปรียบเทียบ "0" กับ "0.0" หรือ "1" กับ "1.0"
                    elif val_str is not None and cur_val != "Empty":
                        try:
                            if float(val_str) == float(cur_val): is_match = True
                        except: pass

                    if is_match:
                        # พยายามเขียนค่ากลับใน Type ที่ถูกต้อง
                        if p.StorageType == StorageType.Integer:
                            try: p.Set(int(float(new_val))) # หากใส่ 0.5 ใน Int จะปัดหรือตัดเศษตาม Revit
                            except: pass
                        elif p.StorageType == StorageType.Double:
                            try: p.Set(float(new_val))
                            except: pass
                        else:
                            p.Set(str(new_val))
                        updated += 1
                
                processed += 1
                pf.progress.Value = processed
                pf.label.Text = "Updating Workset: {0}... ({1}/{2})".format(ws_name, processed, total_elements)
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
    grouped_elements = get_elements_by_workset(doc, doc.ActiveView)
    if not grouped_elements: return

    form = ElementStatusForm(grouped_elements)
    if form.ShowDialog() == DialogResult.OK:
        selected = form.get_selected_rows()
        if selected:
            count = update_elements_with_progress(doc, selected, grouped_elements)
            UI.TaskDialog.Show("Complete", "Updated {0} elements successfully.".format(count))

def get_elements_by_workset(doc, view):
    all_elements = FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().ToElements()
    out = {}
    
    target_bic_ints = [int(bic) for bic in TARGET_CATEGORIES.values()]

    for el in all_elements:
        if not getattr(el, "Category", None):
            continue
            
        bic_int = get_id_int(el.Category.Id)
        
        if bic_int in target_bic_ints:
            ws_name = get_workset_name(doc, el)
            
            if ws_name not in out:
                out[ws_name] = {"elements": [], "count": 0}
            
            out[ws_name]["elements"].append(el)
            out[ws_name]["count"] += 1
            
    return out

class ProgressForm(Form):
    def __init__(self, total):
        self.Width, self.Height = 500, 150
        self.Text = "Processing..."
        self.StartPosition = FormStartPosition.CenterScreen
        self.Font = Font("Segoe UI", 9)
        self.BackColor = Color.White
        
        layout = TableLayoutPanel()
        layout.Dock = DockStyle.Fill
        layout.RowCount = 2
        layout.Padding = Padding(15)
        
        self.label = Label()
        self.label.Text = "Starting..."
        self.label.Dock = DockStyle.Fill
        self.label.TextAlign = ContentAlignment.BottomLeft
        
        safe_total = total if total > 0 else 1
        self.progress = ProgressBar()
        self.progress.Dock = DockStyle.Fill
        self.progress.Minimum = 0
        self.progress.Maximum = safe_total
        self.progress.Height = 30
        
        layout.Controls.Add(self.label, 0, 0)
        layout.Controls.Add(self.progress, 0, 1)
        self.Controls.Add(layout)

if __name__ == "__main__":
    main()