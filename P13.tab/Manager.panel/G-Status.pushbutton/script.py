# -*- coding: utf-8 -*-
__title__ = 'G-Element\nStatus'
__author__ = 'เพิ่มพงษ์ ทวีกุล'

import clr
import System
from System.Collections.Generic import List

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import Autodesk.Revit.DB as DB
import Autodesk.Revit.UI as UI
from System.Windows.Forms import *
from System.Drawing import *

# --------------------------
# CONFIG
# --------------------------
PARAM_NAME = "g_Element Status"

STATUS_MAP = [
    ("0", "0 : ยังไม่ได้ส่ง SHOP DRAWING"),
    ("0.1", "0.1 : ส่ง SHOP DRAWING แล้ว"),
    ("0.5", "0.5 : SHOPDRAWING ตอบกลับ AN หรือ RR"),
    ("1", "1 : SHOPDRAWING ตอบกลับ AP"),
    ("2", "2 : ส่ง AS-BUILT")
]

# รายการ Categories ทั้งหมดตามรูปภาพที่คุณพี่ส่งมา
CAT_LIST = [
    "OST_AirTerminals", "OST_CableTrayFitting", "OST_CableTray", "OST_Ceilings",
    "OST_Columns", "OST_ConduitFitting", "OST_Conduits", "OST_CurtainWallPanels",
    "OST_CurtainWallMullions", "OST_Doors", "OST_DuctAccessory", "OST_DuctFitting",
    "OST_DuctInsulations", "OST_DuctLinings", "OST_DuctPlaceHolders", "OST_DuctCurves",
    "OST_ElectricalEquipment", "OST_ElectricalFixtures", "OST_FireAlarmDevices",
    "OST_FireProtection", "OST_FlexDuctCurves", "OST_FlexPipeCurves", "OST_Floors",
    "OST_FoodServiceEquipment", "OST_Furniture", "OST_GenericModel", "OST_LightingDevices",
    "OST_LightingFixtures", "OST_MechanicalControlDevices", "OST_MechanicalEquipment",
    "OST_MechanicalEquipmentSet", "OST_MedicalEquipment", "OST_PipeAccessory",
    "OST_PipeFitting", "OST_PipeInsulations", "OST_PipePlaceHolders", "OST_PipeCurves",
    "OST_Planting", "OST_PlumbingEquipment", "OST_PlumbingFixtures", "OST_Railings",
    "OST_Ramps", "OST_Roads", "OST_Roofs", "OST_SecurityDevices", "OST_Signage",
    "OST_Site", "OST_Sprinklers", "OST_Stairs", "OST_BeamSystem", "OST_StructuralColumns",
    "OST_StructuralConnections", "OST_StructuralFoundation", "OST_StructuralFraming",
    "OST_StructuralTruss", "OST_TelephoneDevices", "OST_Walls", "OST_Windows", "OST_Wires",
    "OST_Rebar", "OST_StructuralRebar"
]

def get_target_categories():
    valid_cats = []
    for name in CAT_LIST:
        try:
            if hasattr(DB.BuiltInCategory, name):
                valid_cats.append(getattr(DB.BuiltInCategory, name))
        except: pass
    return valid_cats

def get_identity_group_id():
    # แก้ไข Error สำหรับ Revit 2026 โดยเฉพาะ
    try:
        # ใช้ GroupTypeId สำหรับ Revit 2024, 2025, 2026
        return DB.GroupTypeId.IdentityData
    except:
        # กรณีรันในเครื่องรุ่นเก่ากว่า 2024
        return DB.BuiltInParameterGroup.PG_IDENTITY_DATA

def get_workset_name(doc, elem):
    if doc.IsWorkshared:
        ws_id = elem.WorksetId
        if ws_id != DB.WorksetId.InvalidWorksetId:
            ws = doc.GetWorksetTable().GetWorkset(ws_id)
            return ws.Name if ws else "Shared"
    return "Non-Shared"

class ElementStatusForm(Form):
    def __init__(self, doc, grouped_elements):
        self.doc = doc
        self.grouped_elements = grouped_elements
        self.final_data = []
        self.param_exists = self.check_param_exists()
        self.InitializeComponent()

    def check_param_exists(self):
        it = self.doc.ParameterBindings.ForwardIterator()
        while it.MoveNext():
            if it.Key.Name == PARAM_NAME: return True
        return False

    def InitializeComponent(self):
        self.Text = "G-Element Status Manager (Revit 2026 Compatible)"
        self.Width, self.Height = 920, 880
        self.StartPosition = FormStartPosition.CenterScreen
        self.Font = Font("Segoe UI", 9)
        self.BackColor = Color.White

        container = TableLayoutPanel(Dock=DockStyle.Fill, RowCount=7, ColumnCount=1)
        container.RowStyles.Add(RowStyle(SizeType.Absolute, 80))
        container.RowStyles.Add(RowStyle(SizeType.Absolute, 50))
        container.RowStyles.Add(RowStyle(SizeType.Absolute, 50))
        container.RowStyles.Add(RowStyle(SizeType.Percent, 100))
        container.RowStyles.Add(RowStyle(SizeType.Absolute, 180))
        container.RowStyles.Add(RowStyle(SizeType.Absolute, 85))
        
        # 1. Header
        header_panel = Panel(Dock=DockStyle.Fill, BackColor=Color.FromArgb(0, 70, 140))
        lbl_title = Label(Text="Update Element Status (Revit 2026)", ForeColor=Color.White, 
                          Font=Font("Segoe UI", 14, FontStyle.Bold), AutoSize=True, Location=Point(15, 12))
        
        status_txt = "READY" if self.param_exists else "NOT FOUND"
        status_clr = Color.LimeGreen if self.param_exists else Color.Yellow
        self.lbl_status = Label(Text="Parameter Status: " + status_txt, ForeColor=status_clr,
                               Font=Font("Segoe UI", 10, FontStyle.Bold), AutoSize=True, Location=Point(17, 42))
        
        btn_text = "ADD MISSING CATEGORIES" if self.param_exists else "CREATE PARAMETER"
        btn_color = Color.MediumSeaGreen if self.param_exists else Color.OrangeRed
        
        btn_config = Button(Text=btn_text, Location=Point(650, 20), 
                             Size=Size(230, 40), BackColor=btn_color, ForeColor=Color.White, 
                             FlatStyle=FlatStyle.Flat, Font=Font("Segoe UI", 9, FontStyle.Bold))
        btn_config.Click += self.append_categories_safely
        
        header_panel.Controls.Add(lbl_title); header_panel.Controls.Add(self.lbl_status); header_panel.Controls.Add(btn_config)

        # Selection & DataGrid (เหมือนเวอร์ชันก่อนหน้า)
        sel_panel = FlowLayoutPanel(Dock=DockStyle.Fill, Padding=Padding(10, 8, 0, 0))
        for t, f in [("Select All", self.select_all), ("Unselect All", self.unselect_all), 
                     ("Select Highlight", self.select_highlight), ("Unselect Highlight", self.unselect_highlight)]:
            b = Button(Text=t, Width=130, Height=32, BackColor=Color.WhiteSmoke)
            b.Click += f; sel_panel.Controls.Add(b)

        hi_panel = FlowLayoutPanel(Dock=DockStyle.Fill, Padding=Padding(10, 8, 0, 0))
        hi_panel.Controls.Add(Label(Text="Set Status for Highlight:", AutoSize=True, Margin=Padding(0, 7, 5, 0)))
        self.cmb_hi = ComboBox(Width=120, DataSource=[s[0] for s in STATUS_MAP], DropDownStyle=ComboBoxStyle.DropDownList)
        btn_apply = Button(Text="Apply to List", Width=120, Height=28, BackColor=Color.AliceBlue)
        btn_apply.Click += self.apply_highlight_status
        hi_panel.Controls.Add(self.cmb_hi); hi_panel.Controls.Add(btn_apply)
        
        self.dgv = DataGridView(Dock=DockStyle.Fill, RowHeadersVisible=False, AllowUserToAddRows=False,
                                SelectionMode=DataGridViewSelectionMode.FullRowSelect, BackgroundColor=Color.White)
        self.dgv.Columns.Add(DataGridViewCheckBoxColumn(Name="Selected", HeaderText="Update?", Width=70))
        self.dgv.Columns.Add(DataGridViewTextBoxColumn(Name="Workset", HeaderText="Workset", Width=280, ReadOnly=True))
        self.dgv.Columns.Add(DataGridViewTextBoxColumn(Name="Current", HeaderText="Current Status", Width=120, ReadOnly=True))
        self.dgv.Columns.Add(DataGridViewTextBoxColumn(Name="Count", HeaderText="Count", Width=80, ReadOnly=True))
        self.dgv.Columns.Add(DataGridViewComboBoxColumn(Name="NewStatus", HeaderText="New Status", Width=180, DataSource=[s[0] for s in STATUS_MAP]))
        self.populate_data()

        leg_box = GroupBox(Text="Status Definitions", Dock=DockStyle.Fill, Margin=Padding(10))
        lbl_leg = Label(Text="\n".join([s[1] for s in STATUS_MAP]), Dock=DockStyle.Fill, Padding=Padding(10), Font=Font("Segoe UI", 10))
        leg_box.Controls.Add(lbl_leg)

        btn_pnl = FlowLayoutPanel(Dock=DockStyle.Fill, FlowDirection=FlowDirection.RightToLeft, Padding=Padding(0, 10, 20, 0))
        self.btn_ok = Button(Text="UPDATE SELECTED", Size=Size(250, 55), 
                            BackColor=Color.FromArgb(0, 0, 128), ForeColor=Color.White, 
                            Font=Font("Segoe UI", 13, FontStyle.Bold),
                            FlatStyle=FlatStyle.Flat, Enabled=self.param_exists)
        self.btn_ok.Click += self.ok_click
        btn_pnl.Controls.Add(self.btn_ok)

        container.Controls.Add(header_panel, 0, 0); container.Controls.Add(sel_panel, 0, 1)
        container.Controls.Add(hi_panel, 0, 2); container.Controls.Add(self.dgv, 0, 3)
        container.Controls.Add(leg_box, 0, 4); container.Controls.Add(btn_pnl, 0, 5)
        self.Controls.Add(container)

    def append_categories_safely(self, s, e):
        app = self.doc.Application
        sp_file = app.OpenSharedParameterFile()
        if not sp_file:
            UI.TaskDialog.Show("Error", "กรุณาเชื่อมต่อ Shared Parameter File ก่อนครับ")
            return

        ext_def = next((d for g in sp_file.Groups for d in g.Definitions if d.Name == PARAM_NAME), None)
        if not ext_def:
            UI.TaskDialog.Show("Error", "ไม่พบพารามิเตอร์ '{}' ในไฟล์ Shared Parameter".format(PARAM_NAME))
            return

        with DB.Transaction(self.doc, "Safe Category Update") as tx:
            tx.Start()
            try:
                # ดึง Binding ปัจจุบัน
                it = self.doc.ParameterBindings.ForwardIterator()
                binding = None
                while it.MoveNext():
                    if it.Key.Name == PARAM_NAME:
                        binding = it.Current
                        break
                
                # เตรียม CategorySet ใหม่
                new_cat_set = app.Create.NewCategorySet()
                if binding:
                    for old_cat in binding.Categories:
                        new_cat_set.Insert(old_cat)
                
                # เพิ่มหมวดงานจากรูปภาพที่ยังขาดอยู่
                target_cats = get_target_categories()
                for bic in target_cats:
                    try:
                        cat = self.doc.Settings.Categories.get_Item(bic)
                        if cat and not new_cat_set.Contains(cat):
                            new_cat_set.Insert(cat)
                    except: pass
                
                # สร้าง Instance Binding และใช้ Group ID ที่แก้ Error แล้ว
                new_binding = app.Create.NewInstanceBinding(new_cat_set)
                group_id = get_identity_group_id()
                
                # ใช้ ReInsert เพื่อเจาะเข้าไปติ๊กถูกเพิ่ม โดยไม่ลบข้อมูลเดิม
                self.doc.ParameterBindings.ReInsert(ext_def, new_binding, group_id)
                
                tx.Commit()
                UI.TaskDialog.Show("Success", "ติ๊กถูกเพิ่มหมวดงานเรียบร้อยแล้วครับ! (ข้อมูลเดิมยังอยู่ครบ)")
                self.Close()
            except Exception as ex:
                tx.RollBack()
                UI.TaskDialog.Show("Error", str(ex))

    def select_all(self, s, e):
        for r in self.dgv.Rows: r.Cells["Selected"].Value = True
    def unselect_all(self, s, e):
        for r in self.dgv.Rows: r.Cells["Selected"].Value = False
    def select_highlight(self, s, e):
        for r in self.dgv.SelectedRows: r.Cells["Selected"].Value = True
    def unselect_highlight(self, s, e):
        for r in self.dgv.SelectedRows: r.Cells["Selected"].Value = False
    def apply_highlight_status(self, s, e):
        val = self.cmb_hi.SelectedItem
        for r in self.dgv.SelectedRows:
            r.Cells["NewStatus"].Value = val; r.Cells["Selected"].Value = True

    def populate_data(self):
        for ws in sorted(self.grouped_elements.keys()):
            for status, el_list in self.split_by_status(self.grouped_elements[ws]["elements"]).items():
                self.dgv.Rows.Add(False, ws, status, len(el_list), "0")

    def split_by_status(self, elements):
        gs = {}
        for el in elements:
            val = "Not Found"
            if self.param_exists:
                p = el.LookupParameter(PARAM_NAME)
                val = "Empty"
                if p and p.HasValue:
                    val = p.AsValueString() or str(p.AsDouble()) if p.StorageType == DB.StorageType.Double else str(p.AsInteger())
                    val = val.replace(".0","")
            if val not in gs: gs[val] = []
            gs[val].append(el)
        return gs

    def ok_click(self, s, e):
        self.final_data = []
        for r in self.dgv.Rows:
            if r.Cells["Selected"].Value:
                self.final_data.append({"WS": r.Cells["Workset"].Value, "Old": r.Cells["Current"].Value, "New": r.Cells["NewStatus"].Value})
        if self.final_data: self.DialogResult = DialogResult.OK; self.Close()

def main():
    doc = __revit__.ActiveUIDocument.Document
    target_cats = get_target_categories()
    cats = List[DB.ElementId]([DB.ElementId(c) for c in target_cats])
    elems = DB.FilteredElementCollector(doc, doc.ActiveView.Id).WherePasses(DB.ElementMulticategoryFilter(cats)).WhereElementIsNotElementType().ToElements()
    if not elems: return
    
    grouped = {}
    for el in elems:
        ws = get_workset_name(doc, el)
        if ws not in grouped: grouped[ws] = {"elements": []}
        grouped[ws]["elements"].append(el)

    form = ElementStatusForm(doc, grouped)
    if form.ShowDialog() == DialogResult.OK:
        with DB.Transaction(doc, "Update G-Status") as tx:
            tx.Start()
            count = 0
            for task in form.final_data:
                for el in grouped[task["WS"]]["elements"]:
                    p = el.LookupParameter(PARAM_NAME)
                    if p and not p.IsReadOnly:
                        cur = "Empty"
                        if p.HasValue:
                            cur = p.AsValueString() or str(p.AsDouble()) if p.StorageType == DB.StorageType.Double else str(p.AsInteger())
                            cur = cur.replace(".0","")
                        
                        if cur == task["Old"]:
                            new_val = task["New"]
                            try:
                                if p.StorageType == DB.StorageType.String: p.Set(str(new_val))
                                elif p.StorageType == DB.StorageType.Double: p.Set(float(new_val))
                                elif p.StorageType == DB.StorageType.Integer: p.Set(int(float(new_val)))
                                count += 1
                            except: pass
            tx.Commit()
            UI.TaskDialog.Show("G-Status", "อัปเดตสถานะเรียบร้อยแล้วจำนวน {} ชิ้นครับ".format(count))

if __name__ == "__main__":
    main()