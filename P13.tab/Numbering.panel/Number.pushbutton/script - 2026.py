# -*- coding: utf-8 -*-
__title__ = "Numbering\nCategory"
__author__ = "เพิ่มพงษ์ ทวีกุล"
__doc__ = "Numbering + Preview + Predicted Preview (Revit 2026, pyRevit)"

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('System')
clr.AddReference('Microsoft.VisualBasic')  # InputBox

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System import Array
from System.Windows.Forms import *
from System.Drawing import *
from System.Collections.Generic import List
from Microsoft.VisualBasic import Interaction

# Revit 2026 compatible doc/uidoc access
try:
    # Try multiple methods to get doc/uidoc
    doc = None
    uidoc = None
    
    # Method 1: pyRevit 2026+
    try:
        from pyrevit import revit
        doc = revit.doc
        uidoc = revit.uidoc
    except:
        pass
    
    # Method 2: Standard pyRevit
    if not doc:
        try:
            doc = __revit__.ActiveUIDocument.Document
            uidoc = __revit__.ActiveUIDocument
        except:
            pass
    
    # Method 3: Direct from command data
    if not doc:
        try:
            from Autodesk.Revit.UI import TaskDialog
            TaskDialog.Show("Error", "Cannot access Revit document. Please run this script from pyRevit.")
            raise Exception("Cannot access Revit document")
        except:
            pass
    
    if not doc:
        MessageBox.Show("Run this inside Revit (pyRevit / RPS).", "Error")
        raise Exception("Cannot access Revit document")
        
except Exception as ex:
    MessageBox.Show("Run this inside Revit (pyRevit / RPS).\n" + str(ex), "Error")
    raise

# ---------------- Helpers ----------------
def get_element_id_value(element_id):
    """Helper function to get integer value from ElementId (compatible with Revit 2025 and 2026)"""
    try:
        # Try multiple methods to get integer value
        if hasattr(element_id, 'IntegerValue'):
            return element_id.IntegerValue
        elif hasattr(element_id, 'Value'):
            return element_id.Value
        elif isinstance(element_id, int):
            return element_id
        else:
            # Try to cast to int
            try:
                return int(element_id)
            except:
                return 0
    except:
        return 0

def get_element_location(element):
    try:
        loc = element.Location
        if loc:
            if hasattr(loc, "Point"):
                return loc.Point
            elif hasattr(loc, "Curve"):
                try:
                    return loc.Curve.GetEndPoint(0)
                except:
                    pass
    except:
        pass
    try:
        bbox = element.get_BoundingBox(None)
        if bbox:
            return XYZ((bbox.Min.X + bbox.Max.X)/2.0,
                       (bbox.Min.Y + bbox.Max.Y)/2.0,
                       (bbox.Min.Z + bbox.Max.Z)/2.0)
    except:
        pass
    return XYZ(0,0,0)

def has_pile_in_family(element):
    try:
        if hasattr(element, "Symbol") and element.Symbol:
            fam = element.Symbol.Family
            if fam and hasattr(fam, "Name"):
                return "PILE" in fam.Name.upper()
    except:
        pass
    return False

def get_element_category(el):
    try:
        if el and el.Category:
            return el.Category.Name
    except:
        pass
    return "Unknown"

def get_basic_parameters():
    """Auto detect writable params (string/int/double) from sample elements."""
    params = set()
    try:
        coll = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        sample = list(coll)[:40]
        for el in sample:
            try:
                for p in el.Parameters:
                    try:
                        if p and hasattr(p, "Definition") and not p.IsReadOnly:
                            if p.StorageType in (StorageType.String, StorageType.Integer, StorageType.Double):
                                params.add(p.Definition.Name)
                    except:
                        continue
            except:
                continue
    except:
        pass
    preferred = ["Mark", "Type Mark", "Comments", "Type Comments", "Description"]
    out = []
    for pr in preferred:
        if pr in params:
            out.append(pr)
    for p in sorted(params):
        if p not in out:
            out.append(p)
    if not out:
        out = ["Mark", "Type Mark", "Comments", "Description"]
    return out

def collect_elements_for_category(display_name, active_view=None):
    elems = []
    try:
        if display_name == "เสาเข็ม":
            bic = BuiltInCategory.OST_StructuralFoundation
            if active_view:
                col = FilteredElementCollector(doc, active_view.Id).OfCategory(bic).WhereElementIsNotElementType()
            else:
                col = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
            elems = [e for e in list(col) if has_pile_in_family(e)]
        else:
            bic = BUILTIN_MAP.get(display_name, None)
            if bic:
                if active_view:
                    col = FilteredElementCollector(doc, active_view.Id).OfCategory(bic).WhereElementIsNotElementType()
                else:
                    col = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
                elems = list(col)
    except:
        elems = []
    return elems

def is_writable_param(el, pname):
    try:
        p = el.LookupParameter(pname)
        if not p: return False
        if p.IsReadOnly: return False
        if p.StorageType not in (StorageType.String, StorageType.Integer, StorageType.Double): return False
        return True
    except:
        return False

def is_element_visible_in_view(element, view):
    """Check if element is visible in the given view"""
    try:
        # Check if element is in the view
        if hasattr(element, 'IsHidden'):
            if view.IsElementHidden(element.Id):
                return False
        
        # Check if element category is visible in view
        category = element.Category
        if category and not view.GetCategoryHidden(category.Id):
            return True
            
        # Additional visibility checks can be added here
        return True
    except:
        # If we can't determine visibility, assume it's visible
        return True

def filter_elements_by_view_visibility(elements, active_view):
    """Filter elements to only those visible in the active view"""
    if not active_view:
        return elements
    
    visible_elements = []
    for element in elements:
        if is_element_visible_in_view(element, active_view):
            visible_elements.append(element)
    
    return visible_elements

# ---------------- Defaults (Option A prefixes) ----------------
DEFAULT_PREFIX_MAP = {
    "Structural Columns": "C",
    "Structural Framing": "B",
    "Structural Foundations": "FDN",
    "Walls": "W",
    "Floors": "FL",
    "Roofs": "RF",
    "Doors": "D",
    "Windows": "WN",
    "Stairs": "ST",
    "Generic Models": "GM",
    "Furniture": "FN",
    "Plumbing Fixtures": "PF",
    "Plumbing Equipment": "PE",
    "Electrical Fixtures": "EF",
    "Specialty Equipment": "SE",
    "Pipes": "PP",
    "Pipe Fittings": "PFIT",
    "Mechanical Equipment": "ME",
    "Lighting Fixtures": "LG",
    "Electrical Equipment": "EE",
    "Ceilings": "CL",
    "Railings": "RL",
    "Site": "SITE",
    "Detail Items": "DT",
    "เสาเข็ม": "P"
}

# Map display name -> BuiltInCategory where possible
BUILTIN_MAP = {
    "Structural Columns": BuiltInCategory.OST_StructuralColumns,
    "Structural Framing": BuiltInCategory.OST_StructuralFraming,
    "Structural Foundations": BuiltInCategory.OST_StructuralFoundation,
    "Walls": BuiltInCategory.OST_Walls,
    "Floors": BuiltInCategory.OST_Floors,
    "Roofs": BuiltInCategory.OST_Roofs,
    "Doors": BuiltInCategory.OST_Doors,
    "Windows": BuiltInCategory.OST_Windows,
    "Stairs": BuiltInCategory.OST_Stairs,
    "Generic Models": BuiltInCategory.OST_GenericModel,
    "Furniture": BuiltInCategory.OST_Furniture,
    "Plumbing Fixtures": BuiltInCategory.OST_PlumbingFixtures,
    "Plumbing Equipment": BuiltInCategory.OST_PlumbingEquipment,
    "Electrical Fixtures": BuiltInCategory.OST_ElectricalFixtures,
    "Specialty Equipment": getattr(BuiltInCategory, "OST_SpecialtyEquipment", None) or getattr(BuiltInCategory, "OST_SpecialityEquipment", None),
    "Pipes": BuiltInCategory.OST_PipeCurves,
    "Pipe Fittings": BuiltInCategory.OST_PipeFitting,
    "Mechanical Equipment": BuiltInCategory.OST_MechanicalEquipment,
    "Lighting Fixtures": BuiltInCategory.OST_LightingFixtures,
    "Electrical Equipment": BuiltInCategory.OST_ElectricalEquipment,
    "Ceilings": BuiltInCategory.OST_Ceilings,
    "Railings": BuiltInCategory.OST_Railings,
    "Site": BuiltInCategory.OST_Site,
    "Detail Items": getattr(BuiltInCategory, "OST_DetailComponents", None) or getattr(BuiltInCategory, "OST_DetailItems", None),
    "เสาเข็ม": BuiltInCategory.OST_StructuralFoundation
}

# ---------------- ResultForm ----------------
class ResultForm(Form):
    def __init__(self, title, message):
        self.Text = title
        self.Size = Size(600, 500)
        self.StartPosition = FormStartPosition.CenterScreen
        self.Font = Font("Segoe UI", 10)
        self.BackColor = Color.White

        # Main panel with padding
        main_panel = Panel()
        main_panel.Dock = DockStyle.Fill
        main_panel.Padding = Padding(10)
        main_panel.BackColor = Color.White

        # Title label
        title_label = Label()
        title_label.Text = title
        title_label.Font = Font("Segoe UI", 12, FontStyle.Bold)
        title_label.ForeColor = Color.FromArgb(0, 51, 102)
        title_label.Dock = DockStyle.Top
        title_label.Height = 30
        title_label.TextAlign = ContentAlignment.MiddleLeft

        # Text box for results
        txt = TextBox()
        txt.Multiline = True
        txt.ReadOnly = True
        txt.ScrollBars = ScrollBars.Vertical
        txt.Dock = DockStyle.Fill
        txt.Text = message
        txt.Font = Font("Consolas", 9)
        txt.BackColor = Color.FromArgb(248, 248, 248)
        txt.BorderStyle = BorderStyle.FixedSingle
        txt.Margin = Padding(0, 10, 0, 10)

        # OK button
        btn = Button()
        btn.Text = "OK"
        btn.Dock = DockStyle.Bottom
        btn.Height = 40
        btn.BackColor = Color.FromArgb(0, 102, 204)
        btn.ForeColor = Color.White
        btn.FlatStyle = FlatStyle.Flat
        btn.Font = Font("Segoe UI", 10, FontStyle.Bold)
        btn.Click += self._close

        main_panel.Controls.Add(title_label)
        main_panel.Controls.Add(txt)
        main_panel.Controls.Add(btn)
        self.Controls.Add(main_panel)

    def _close(self, s, e):
        self.Close()

# ---------------- Toggle Switch Control ----------------
class ToggleSwitch(Control):
    def __init__(self):
        self.Width = 60
        self.Height = 30
        self.BackColor = Color.LightGray
        self.Checked = False
        self.Click += self._toggle
        self.Paint += self._on_paint
        
    def _toggle(self, sender, args):
        self.Checked = not self.Checked
        self.Invalidate()
        
    def _on_paint(self, sender, args):
        g = args.Graphics
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        
        # Draw background
        if self.Checked:
            g.FillRectangle(Brushes.DodgerBlue, 0, 0, self.Width, self.Height)
        else:
            g.FillRectangle(Brushes.LightGray, 0, 0, self.Width, self.Height)
            
        # Draw toggle circle
        circle_x = self.Width - 15 if self.Checked else 15
        g.FillEllipse(Brushes.White, circle_x - 10, 5, 20, 20)
        g.DrawEllipse(Pens.DarkGray, circle_x - 10, 5, 20, 20)
        
        # Draw labels
        if self.Checked:
            g.DrawString("ON", Font("Segoe UI", 7), Brushes.White, 8, 8)
        else:
            g.DrawString("OFF", Font("Segoe UI", 7), Brushes.DarkGray, self.Width - 22, 8)

# ---------------- NumberingForm ----------------
class NumberingForm(Form):
    def __init__(self):
        self.Text = "Element Numbering Tool"
        self.Size = Size(1080, 900)
        self.StartPosition = FormStartPosition.CenterScreen
        self.Font = Font("Segoe UI", 9)
        self.BackColor = Color.FromArgb(240, 240, 240)
        self.MinimumSize = Size(980, 800)

        self.preview_elements = []
        self.active_view = doc.ActiveView if doc else None  # Get current active view
        self.use_view_filter = False  # Default to project mode

        self._build_ui()
        self._load_categories()
        self._load_parameters()

    def _build_ui(self):
        # Main container with padding
        main_container = Panel()
        main_container.Dock = DockStyle.Fill
        main_container.Padding = Padding(15)
        main_container.BackColor = Color.FromArgb(240, 240, 240)

        # Create two main panels side by side
        left_panel = self._create_left_panel()
        right_panel = self._create_right_panel()

        # Add panels to main container
        main_container.Controls.Add(left_panel)
        main_container.Controls.Add(right_panel)

        self.Controls.Add(main_container)

    def _create_left_panel(self):
        # Left panel - Category & Prefix
        panel = Panel()
        panel.Size = Size(500, 850)
        panel.Location = Point(0, 0)
        panel.BackColor = Color.White
        panel.BorderStyle = BorderStyle.FixedSingle
        panel.Padding = Padding(10)

        # Title
        title = Label()
        title.Text = "หมวดหมู่และคำนำหน้า"
        title.Location = Point(10, 10)
        title.Size = Size(400, 25)
        title.Font = Font("Segoe UI", 11, FontStyle.Bold)
        title.ForeColor = Color.FromArgb(0, 51, 102)
        panel.Controls.Add(title)

        # Category grid
        self.grid_cat = DataGridView()
        self.grid_cat.Location = Point(10, 45)
        self.grid_cat.Size = Size(460, 500)
        self.grid_cat.AllowUserToAddRows = False
        self.grid_cat.AllowUserToDeleteRows = False
        self.grid_cat.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        self.grid_cat.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        self.grid_cat.BackgroundColor = Color.White
        self.grid_cat.BorderStyle = BorderStyle.FixedSingle
        self.grid_cat.DefaultCellStyle.Font = Font("Segoe UI", 9)
        self.grid_cat.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 102, 204)
        self.grid_cat.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        self.grid_cat.ColumnHeadersDefaultCellStyle.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.grid_cat.EnableHeadersVisualStyles = False

        # Configure columns
        col_use = DataGridViewCheckBoxColumn()
        col_use.HeaderText = "เลือก"
        col_use.Name = "Use"
        col_use.Width = 45
        
        col_name = DataGridViewTextBoxColumn()
        col_name.HeaderText = "หมวดหมู่"
        col_name.Name = "Category"
        col_name.ReadOnly = True
        
        col_pref = DataGridViewTextBoxColumn()
        col_pref.HeaderText = "คำนำหน้า"
        col_pref.Name = "Prefix"
        col_pref.Width = 70
        
        col_cnt = DataGridViewTextBoxColumn()
        col_cnt.HeaderText = "จำนวน"
        col_cnt.Name = "Count"
        col_cnt.ReadOnly = True
        col_cnt.Width = 60

        self.grid_cat.Columns.AddRange(Array[DataGridViewColumn]([col_use, col_name, col_pref, col_cnt]))
        panel.Controls.Add(self.grid_cat)

        # Button container - เพิ่มปุ่มรีเฟรชพารามิเตอร์
        button_panel = Panel()
        button_panel.Location = Point(10, 560)
        button_panel.Size = Size(460, 40)
        
        btn_select = self._create_button("เลือกทั้งหมด", Point(0, 0), self._select_all)
        btn_deselect = self._create_button("ยกเลิกทั้งหมด", Point(120, 0), self._deselect_all)
        btn_refresh = self._create_button("รีเฟรชจำนวน", Point(240, 0), self._refresh_counts)
        
        button_panel.Controls.Add(btn_select)
        button_panel.Controls.Add(btn_deselect)
        button_panel.Controls.Add(btn_refresh)
        panel.Controls.Add(button_panel)

        # ปุ่มรีเฟรชพารามิเตอร์ทั้งโครงการ
        refresh_param_panel = Panel()
        refresh_param_panel.Location = Point(10, 610)
        refresh_param_panel.Size = Size(460, 40)
        
        btn_refresh_param = self._create_button("รีเฟรชพารามิเตอร์ทั้งโครงการ", Point(0, 0), self._refresh_all_parameters)
        btn_refresh_param.BackColor = Color.FromArgb(156, 39, 176)  # สีม่วงเพื่อให้แตกต่าง
        refresh_param_panel.Controls.Add(btn_refresh_param)
        panel.Controls.Add(refresh_param_panel)

        # Info label
        info_label = Label()
        info_label.Text = "เลือกหมวดหมู่และคลิก 'โหลดตัวอย่าง' เพื่อดำเนินการต่อ"
        info_label.Location = Point(10, 660)
        info_label.Size = Size(460, 40)
        info_label.ForeColor = Color.FromArgb(0, 102, 0)
        info_label.Font = Font("Segoe UI", 9, FontStyle.Italic)
        info_label.TextAlign = ContentAlignment.MiddleLeft
        panel.Controls.Add(info_label)

        return panel

    def _create_right_panel(self):
        # Right panel - Numbering Options
        panel = Panel()
        panel.Size = Size(540, 850)
        panel.Location = Point(520, 0)
        panel.BackColor = Color.White
        panel.BorderStyle = BorderStyle.FixedSingle
        panel.Padding = Padding(10)

        # Title
        title = Label()
        title.Text = "ตัวเลือกการกำหนดหมายเลข"
        title.Location = Point(10, 10)
        title.Size = Size(450, 25)
        title.Font = Font("Segoe UI", 11, FontStyle.Bold)
        title.ForeColor = Color.FromArgb(0, 51, 102)
        panel.Controls.Add(title)

        # Current View Info
        view_info = Label()
        view_name = self.active_view.Name if self.active_view else "ไม่มี"
        view_info.Text = "วิวปัจจุบัน: " + view_name
        view_info.Location = Point(10, 40)
        view_info.Size = Size(450, 20)
        view_info.Font = Font("Segoe UI", 9, FontStyle.Bold)
        view_info.ForeColor = Color.FromArgb(0, 102, 0)
        panel.Controls.Add(view_info)

        # View Filter Toggle
        toggle_panel = Panel()
        toggle_panel.Location = Point(10, 70)
        toggle_panel.Size = Size(500, 40)
        
        toggle_label = Label()
        toggle_label.Text = "โหมดการค้นหา:"
        toggle_label.Location = Point(0, 8)
        toggle_label.Size = Size(120, 20)
        toggle_label.Font = Font("Segoe UI", 9, FontStyle.Bold)
        toggle_panel.Controls.Add(toggle_label)
        
        # Toggle switch
        self.toggle_switch = ToggleSwitch()
        self.toggle_switch.Location = Point(120, 0)
        self.toggle_switch.Checked = self.use_view_filter
        self.toggle_switch.Click += self._toggle_search_mode
        toggle_panel.Controls.Add(self.toggle_switch)
        
        # Mode labels
        self.lbl_project_mode = Label()
        self.lbl_project_mode.Text = "ทั้งโครงการ"
        self.lbl_project_mode.Location = Point(190, 8)
        self.lbl_project_mode.Size = Size(120, 20)
        self.lbl_project_mode.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.lbl_project_mode.ForeColor = Color.FromArgb(0, 102, 0)
        toggle_panel.Controls.Add(self.lbl_project_mode)
        
        self.lbl_view_mode = Label()
        self.lbl_view_mode.Text = "เฉพาะวิวปัจจุบัน"
        self.lbl_view_mode.Location = Point(190, 8)
        self.lbl_view_mode.Size = Size(150, 20)
        self.lbl_view_mode.Font = Font("Segoe UI", 9)
        self.lbl_view_mode.ForeColor = Color.Gray
        toggle_panel.Controls.Add(self.lbl_view_mode)
        
        panel.Controls.Add(toggle_panel)

        # Parameter selection
        param_label = Label()
        param_label.Text = "พารามิเตอร์เป้าหมาย (ตรวจสอบอัตโนมัติ):"
        param_label.Location = Point(10, 120)
        param_label.Size = Size(350, 20)
        param_label.Font = Font("Segoe UI", 9, FontStyle.Bold)
        panel.Controls.Add(param_label)
        
        self.cmb_param = ComboBox()
        self.cmb_param.Location = Point(10, 145)
        self.cmb_param.Size = Size(350, 26)
        self.cmb_param.DropDownStyle = ComboBoxStyle.DropDownList
        self.cmb_param.Font = Font("Segoe UI", 9)
        panel.Controls.Add(self.cmb_param)
        
        self.chk_custom = CheckBox()
        self.chk_custom.Text = "ใช้พารามิเตอร์ที่กำหนดเอง"
        self.chk_custom.Location = Point(370, 147)
        self.chk_custom.Size = Size(150, 20)
        self.chk_custom.Font = Font("Segoe UI", 9)
        self.chk_custom.CheckedChanged += self._toggle_custom
        panel.Controls.Add(self.chk_custom)
        
        self.txt_custom = TextBox()
        self.txt_custom.Text = "พิมพ์ชื่อพารามิเตอร์ที่ต้องการ..."
        self.txt_custom.Location = Point(10, 180)
        self.txt_custom.Size = Size(480, 24)
        self.txt_custom.ForeColor = Color.Gray
        self.txt_custom.Font = Font("Segoe UI", 9)
        self.txt_custom.Enter += self._custom_enter
        self.txt_custom.Leave += self._custom_leave
        self.txt_custom.Enabled = False
        panel.Controls.Add(self.txt_custom)

        # Numbering options group
        grp_numbering = GroupBox()
        grp_numbering.Text = "การตั้งค่าการกำหนดหมายเลข"
        grp_numbering.Location = Point(10, 220)
        grp_numbering.Size = Size(500, 120)
        grp_numbering.Font = Font("Segoe UI", 9, FontStyle.Bold)
        grp_numbering.ForeColor = Color.FromArgb(0, 51, 102)
        
        # Start number
        start_label = Label()
        start_label.Text = "หมายเลขเริ่มต้น:"
        start_label.Location = Point(15, 30)
        start_label.Size = Size(100, 20)
        start_label.Font = Font("Segoe UI", 9)
        start_label.ForeColor = Color.Black
        grp_numbering.Controls.Add(start_label)
        
        self.txt_start = TextBox()
        self.txt_start.Text = "1"
        self.txt_start.Location = Point(120, 27)
        self.txt_start.Size = Size(60, 24)
        self.txt_start.Font = Font("Segoe UI", 9)
        grp_numbering.Controls.Add(self.txt_start)

        # Digit selection
        digit_label = Label()
        digit_label.Text = "จำนวนหลัก:"
        digit_label.Location = Point(200, 30)
        digit_label.Size = Size(80, 20)
        digit_label.Font = Font("Segoe UI", 9)
        digit_label.ForeColor = Color.Black
        grp_numbering.Controls.Add(digit_label)
        
        # Radio buttons for digit selection
        self.digit_radios = []
        digit_options = ["1", "2", "3", "4"]
        x_pos = 280
        for i, digits in enumerate(digit_options):
            rb = RadioButton()
            rb.Text = digits
            rb.Location = Point(x_pos, 28)
            rb.Size = Size(40, 20)
            rb.Tag = int(digits)
            rb.Font = Font("Segoe UI", 9)
            rb.ForeColor = Color.Black
            if i == 1:  # Default to 2 digits
                rb.Checked = True
            self.digit_radios.append(rb)
            grp_numbering.Controls.Add(rb)
            x_pos += 45

        # Numbering mode
        mode_label = Label()
        mode_label.Text = "โหมดการกำหนดหมายเลข:"
        mode_label.Location = Point(15, 65)
        mode_label.Size = Size(140, 20)
        mode_label.Font = Font("Segoe UI", 9)
        mode_label.ForeColor = Color.Black
        grp_numbering.Controls.Add(mode_label)
        
        self.rb_cont = RadioButton()
        self.rb_cont.Text = "ต่อเนื่อง"
        self.rb_cont.Location = Point(160, 63)
        self.rb_cont.Font = Font("Segoe UI", 9)
        self.rb_cont.ForeColor = Color.Black
        self.rb_cont.Checked = True
        
        self.rb_percat = RadioButton()
        self.rb_percat.Text = "ตามหมวดหมู่"
        self.rb_percat.Location = Point(230, 63)
        self.rb_percat.Font = Font("Segoe UI", 9)
        self.rb_percat.ForeColor = Color.Black
        
        grp_numbering.Controls.Add(self.rb_cont)
        grp_numbering.Controls.Add(self.rb_percat)

        panel.Controls.Add(grp_numbering)

        # Direction options
        dir_label = Label()
        dir_label.Text = "ทิศทางการเรียงลำดับ:"
        dir_label.Location = Point(10, 350)
        dir_label.Size = Size(150, 20)
        dir_label.Font = Font("Segoe UI", 9, FontStyle.Bold)
        panel.Controls.Add(dir_label)
        
        opts = ["ซ้ายไปขวา", "ขวาไปซ้าย", "ล่างไปบน", "บนไปล่าง", "ใกล้ไปไกล (Z)", "ไกลไปใกล้ (Z)"]
        self.dir_radios = []
        x = 10
        y = 375
        for i, txt in enumerate(opts):
            r = RadioButton()
            r.Text = txt
            r.Location = Point(x, y)
            r.Size = Size(150, 22)
            r.Font = Font("Segoe UI", 9)
            if i == 0: 
                r.Checked = True
            panel.Controls.Add(r)
            self.dir_radios.append(r)
            x += 160
            if (i + 1) % 3 == 0:
                x = 10
                y += 30

        # Preview section
        preview_label = Label()
        preview_label.Text = "ตัวอย่าง (ดับเบิลคลิกที่แถวเพื่อกำหนดหมายเลขด้วยตนเอง):"
        preview_label.Location = Point(10, 460)
        preview_label.Size = Size(500, 20)
        preview_label.Font = Font("Segoe UI", 9, FontStyle.Bold)
        panel.Controls.Add(preview_label)
        
        self.grid_preview = DataGridView()
        self.grid_preview.Location = Point(10, 485)
        self.grid_preview.Size = Size(500, 250)
        self.grid_preview.ReadOnly = True
        self.grid_preview.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        self.grid_preview.MultiSelect = False
        self.grid_preview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        self.grid_preview.BackgroundColor = Color.White
        self.grid_preview.BorderStyle = BorderStyle.FixedSingle
        self.grid_preview.DefaultCellStyle.Font = Font("Segoe UI", 9)
        self.grid_preview.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 102, 204)
        self.grid_preview.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        self.grid_preview.ColumnHeadersDefaultCellStyle.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.grid_preview.EnableHeadersVisualStyles = False
        
        # Configure preview columns
        c_id = DataGridViewTextBoxColumn()
        c_id.HeaderText = "รหัส"
        c_id.Name = "Id"
        c_id.Width = 70
        
        c_name = DataGridViewTextBoxColumn()
        c_name.HeaderText = "ชื่อ"
        c_name.Name = "Name"
        
        c_val = DataGridViewTextBoxColumn()
        c_val.HeaderText = "ค่า"
        c_val.Name = "Value"
        
        c_num = DataGridViewTextBoxColumn()
        c_num.HeaderText = "กำหนดหมายเลขแล้ว"
        c_num.Name = "Numbered"
        c_num.Width = 100
        
        self.grid_preview.Columns.AddRange(Array[DataGridViewColumn]([c_id, c_name, c_val, c_num]))
        self.grid_preview.CellDoubleClick += self._preview_double_click
        panel.Controls.Add(self.grid_preview)

        # Action buttons - Two rows for better organization
        action_panel1 = Panel()
        action_panel1.Location = Point(10, 745)
        action_panel1.Size = Size(500, 40)
        
        btn_load = self._create_action_button("โหลดตัวอย่าง", Point(0, 0), self._load_preview, Color.FromArgb(0, 102, 204))
        btn_step = self._create_action_button("ขั้นตอน", Point(130, 0), self._step, Color.FromArgb(76, 175, 80))
        btn_run = self._create_action_button("รันทั้งหมด", Point(240, 0), self._run_all, Color.FromArgb(46, 125, 50))
        btn_results = self._create_action_button("แสดงผลลัพธ์", Point(350, 0), self._show_results, Color.FromArgb(255, 152, 0))

        action_panel1.Controls.Add(btn_load)
        action_panel1.Controls.Add(btn_step)
        action_panel1.Controls.Add(btn_run)
        action_panel1.Controls.Add(btn_results)
        panel.Controls.Add(action_panel1)

        # Second row of action buttons
        action_panel2 = Panel()
        action_panel2.Location = Point(10, 790)
        action_panel2.Size = Size(500, 40)
        
        btn_high = self._create_action_button("ไฮไลต์+ซูม", Point(0, 0), self._highlight_zoom, Color.FromArgb(156, 39, 176))
        btn_clear = self._create_action_button("ล้างไฮไลต์", Point(130, 0), self._clear_highlight, Color.FromArgb(121, 85, 72))
        btn_close = self._create_action_button("ปิด", Point(260, 0), self._close_form, Color.FromArgb(200, 50, 50))

        action_panel2.Controls.Add(btn_high)
        action_panel2.Controls.Add(btn_clear)
        action_panel2.Controls.Add(btn_close)
        panel.Controls.Add(action_panel2)

        # Status label
        self.lbl_status = Label()
        self.lbl_status.Text = "พร้อมทำงาน - โหมดทั้งโครงการ"
        self.lbl_status.Location = Point(10, 835)
        self.lbl_status.Size = Size(500, 20)
        self.lbl_status.Font = Font("Segoe UI", 9)
        self.lbl_status.ForeColor = Color.FromArgb(0, 102, 0)
        panel.Controls.Add(self.lbl_status)

        return panel

    def _create_button(self, text, location, click_handler):
        btn = Button()
        btn.Text = text
        btn.Location = location
        btn.Size = Size(110, 35)
        btn.Font = Font("Segoe UI", 9)
        btn.BackColor = Color.FromArgb(0, 102, 204)
        btn.ForeColor = Color.White
        btn.FlatStyle = FlatStyle.Flat
        btn.Click += click_handler
        return btn

    def _create_action_button(self, text, location, click_handler, color):
        btn = Button()
        btn.Text = text
        btn.Location = location
        btn.Size = Size(120, 35)
        btn.Font = Font("Segoe UI", 9)
        btn.BackColor = color
        btn.ForeColor = Color.White
        btn.FlatStyle = FlatStyle.Flat
        btn.Click += click_handler
        return btn

    def _toggle_search_mode(self, sender, args):
        self.use_view_filter = self.toggle_switch.Checked
        
        # Update mode labels appearance
        if self.use_view_filter:
            self.lbl_view_mode.ForeColor = Color.FromArgb(0, 102, 0)
            self.lbl_view_mode.Font = Font("Segoe UI", 9, FontStyle.Bold)
            self.lbl_project_mode.ForeColor = Color.Gray
            self.lbl_project_mode.Font = Font("Segoe UI", 9)
            self._set_status("เปลี่ยนเป็นโหมดวิว - เฉพาะองค์ประกอบที่มองเห็นในวิวปัจจุบัน")
        else:
            self.lbl_project_mode.ForeColor = Color.FromArgb(0, 102, 0)
            self.lbl_project_mode.Font = Font("Segoe UI", 9, FontStyle.Bold)
            self.lbl_view_mode.ForeColor = Color.Gray
            self.lbl_view_mode.Font = Font("Segoe UI", 9)
            self._set_status("เปลี่ยนเป็นโหมดโครงการ - องค์ประกอบทั้งหมดในโครงการ")
        
        # Refresh counts to reflect the new mode
        self._refresh_counts(None, None)

    def _close_form(self, sender, args):
        self.Close()

    # ---------- category helpers ----------
    def _load_categories(self):
        try:
            self.grid_cat.Rows.Clear()
            for k in DEFAULT_PREFIX_MAP.keys():
                cnt = 0
                try:
                    # Use active view for counting if view filter is enabled
                    active_view = self.active_view if self.use_view_filter else None
                    cnt = len(collect_elements_for_category(k, active_view))
                except:
                    cnt = 0
                idx = self.grid_cat.Rows.Add()
                row = self.grid_cat.Rows[idx]
                row.Cells["Category"].Value = k
                row.Cells["Prefix"].Value = DEFAULT_PREFIX_MAP.get(k, "X")
                row.Cells["Count"].Value = cnt
                row.Cells["Use"].Value = False
            self._set_status("โหลดหมวดหมู่เรียบร้อยแล้ว")
        except Exception as ex:
            self._set_status("โหลดหมวดหมู่ล้มเหลว: " + str(ex), True)

    def _refresh_counts(self, s, e):
        try:
            for r in self.grid_cat.Rows:
                try:
                    name = r.Cells["Category"].Value
                    # Use active view for counting if view filter is enabled
                    active_view = self.active_view if self.use_view_filter else None
                    cnt = len(collect_elements_for_category(name, active_view))
                    r.Cells["Count"].Value = cnt
                except:
                    r.Cells["Count"].Value = 0
            mode = "โหมดวิว" if self.use_view_filter else "โหมดโครงการ"
            self._set_status("รีเฟรชจำนวนเรียบร้อยแล้ว ({})".format(mode))
        except Exception as ex:
            self._set_status("รีเฟรชจำนวนล้มเหลว: " + str(ex), True)

    def _select_all(self, s, e):
        for r in self.grid_cat.Rows:
            r.Cells["Use"].Value = True
        self._set_status("เลือกหมวดหมู่ทั้งหมดแล้ว")

    def _deselect_all(self, s, e):
        for r in self.grid_cat.Rows:
            r.Cells["Use"].Value = False
        self._set_status("ยกเลิกการเลือกหมวดหมู่ทั้งหมดแล้ว")

    # ---------- parameters ----------
    def _load_parameters(self):
        try:
            self.cmb_param.Items.Clear()
            for p in get_basic_parameters():
                self.cmb_param.Items.Add(p)
            if self.cmb_param.Items.Count > 0:
                self.cmb_param.SelectedIndex = 0
            self._set_status("โหลดพารามิเตอร์เรียบร้อยแล้ว")
        except Exception as ex:
            self._set_status("โหลดพารามิเตอร์ล้มเหลว: " + str(ex), True)

    def _toggle_custom(self, s, e):
        self.txt_custom.Enabled = self.chk_custom.Checked
        self.cmb_param.Enabled = not self.chk_custom.Checked

    def _custom_enter(self, s, e):
        if self.txt_custom.Text == "พิมพ์ชื่อพารามิเตอร์ที่ต้องการ...":
            self.txt_custom.Text = ""
            self.txt_custom.ForeColor = Color.Black

    def _custom_leave(self, s, e):
        if not self.txt_custom.Text:
            self.txt_custom.Text = "พิมพ์ชื่อพารามิเตอร์ที่ต้องการ..."
            self.txt_custom.ForeColor = Color.Gray

    # ---------- Helper function to get selected digit count ----------
    def _get_digit_count(self):
        for rb in self.digit_radios:
            if rb.Checked:
                return int(rb.Tag)
        return 2  # Default to 2 digits

    # ---------- Helper function to format number with leading zeros ----------
    def _format_number(self, number, digit_count):
        """Format number with leading zeros to match digit count"""
        if digit_count <= 0:
            return str(number)
        return str(number).zfill(digit_count)

    # ---------- Function to refresh all parameters ----------
    def _refresh_all_parameters(self, s, e):
        """รีเฟรชพารามิเตอร์ทั้งหมดจากทั้งโครงการ"""
        try:
            # แสดงกล่องข้อความยืนยัน
            result = MessageBox.Show(
                "คุณต้องการรีเฟรชพารามิเตอร์ทั้งโครงการหรือไม่?\n\nการดำเนินการนี้อาจใช้เวลาสักครู่หากโครงการมีขนาดใหญ่",
                "ยืนยันการรีเฟรชพารามิเตอร์",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            )
            
            if result == DialogResult.Yes:
                self._set_status("กำลังรวบรวมพารามิเตอร์จากทั้งโครงการ...")
                
                # รวบรวมพารามิเตอร์จากองค์ประกอบทั้งหมดในโครงการ
                all_params = set()
                try:
                    # ใช้ FilteredElementCollector เพื่อรวบรวมองค์ประกอบทั้งหมด
                    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
                    elements = list(collector)
                    
                    # จำกัดจำนวนองค์ประกอบที่ตรวจสอบเพื่อป้องกันการทำงานช้า
                    max_elements = 1000
                    sample_elements = elements[:max_elements] if len(elements) > max_elements else elements
                    
                    for el in sample_elements:
                        try:
                            for param in el.Parameters:
                                try:
                                    if (param and 
                                        hasattr(param, "Definition") and 
                                        not param.IsReadOnly and
                                        param.StorageType in (StorageType.String, StorageType.Integer, StorageType.Double)):
                                        all_params.add(param.Definition.Name)
                                except:
                                    continue
                        except:
                            continue
                            
                    # เรียงลำดับพารามิเตอร์ตามลำดับที่ต้องการ
                    preferred = ["Mark", "Type Mark", "Comments", "Type Comments", "Description"]
                    sorted_params = []
                    
                    # เพิ่มพารามิเตอร์ที่ต้องการก่อน
                    for pr in preferred:
                        if pr in all_params:
                            sorted_params.append(pr)
                            
                    # เพิ่มพารามิเตอร์อื่นๆ ที่เหลือ
                    for param in sorted(all_params):
                        if param not in sorted_params:
                            sorted_params.append(param)
                            
                    # อัปเดต ComboBox
                    self.cmb_param.Items.Clear()
                    for param in sorted_params:
                        self.cmb_param.Items.Add(param)
                        
                    if self.cmb_param.Items.Count > 0:
                        self.cmb_param.SelectedIndex = 0
                        
                    self._set_status("รีเฟรชพารามิเตอร์เรียบร้อยแล้ว: พบ {} พารามิเตอร์".format(len(sorted_params)))
                    
                except Exception as ex:
                    self._set_status("รีเฟรชพารามิเตอร์ล้มเหลว: " + str(ex), True)
                    
        except Exception as ex:
            self._set_status("รีเฟรชพารามิเตอร์ล้มเหลว: " + str(ex), True)

    # ---------- preview ----------
    def _get_param(self):
        if self.chk_custom.Checked and self.txt_custom.Text and self.txt_custom.Text != "พิมพ์ชื่อพารามิเตอร์ที่ต้องการ...":
            return self.txt_custom.Text
        else:
            return str(self.cmb_param.SelectedItem) if self.cmb_param.SelectedItem is not None else ""

    def _load_preview(self, s, e):
        try:
            elems = []
            prefix_map = {}
            
            # ใช้เฉพาะการเลือกจากหมวดหมู่ในกริด
            selcats = []
            for r in self.grid_cat.Rows:
                try:
                    if r.Cells["Use"].Value:
                        selcats.append((r.Cells["Category"].Value, r.Cells["Prefix"].Value))
                except:
                    pass
                    
            if not selcats:
                MessageBox.Show("กรุณาเลือกอย่างน้อยหนึ่งหมวดหมู่ก่อน", "ข้อมูล")
                return
                
            # Determine if we should filter by view
            active_view = self.active_view if self.use_view_filter else None
            
            for name,pref in selcats:
                found = collect_elements_for_category(name, active_view)
                
                # Additional visibility filtering for view mode
                if self.use_view_filter and active_view:
                    found = filter_elements_by_view_visibility(found, active_view)
                
                for el in found:
                    elems.append(el)
                    prefix_map[el.Id] = pref if pref else DEFAULT_PREFIX_MAP.get(name, "X")

            param = self._get_param()
            if not param:
                MessageBox.Show("กรุณาเลือกพารามิเตอร์หรือใช้พารามิเตอร์ที่กำหนดเอง", "ข้อมูล")
                return

            filtered = []
            for el in elems:
                try:
                    if el.LookupParameter(param) is not None:
                        filtered.append(el)
                except:
                    continue
                    
            if not filtered:
                MessageBox.Show("ไม่พบองค์ประกอบที่มีพารามิเตอร์เป้าหมาย", "ข้อมูล")
                return

            # sort
            dir_text = next((r.Text for r in self.dir_radios if r.Checked), "ซ้ายไปขวา")
            if dir_text == "ซ้ายไปขวา":
                keyf = lambda x: get_element_location(x).X
                rev=False
            elif dir_text == "ขวาไปซ้าย":
                keyf = lambda x: get_element_location(x).X
                rev=True
            elif dir_text == "ล่างไปบน":
                keyf = lambda x: get_element_location(x).Y
                rev=False
            elif dir_text == "บนไปล่าง":
                keyf = lambda x: get_element_location(x).Y
                rev=True
            elif dir_text == "ใกล้ไปไกล (Z)":
                keyf = lambda x: get_element_location(x).Z
                rev=False
            else:
                keyf = lambda x: get_element_location(x).Z
                rev=True

            try:
                filtered = sorted(filtered, key=keyf, reverse=rev)
            except:
                pass

            self.preview_elements = filtered
            self.grid_preview.Rows.Clear()
            
            # Get numbering settings for preview
            try:
                start = int(self.txt_start.Text)
            except:
                start = 1
            digit_count = self._get_digit_count()
            
            for i, el in enumerate(filtered):
                # Use the helper function to get element id value
                pid = get_element_id_value(el.Id)
                name = el.Name if hasattr(el, "Name") else str(el.Id)
                p = el.LookupParameter(param)
                val = ""
                try:
                    if p.StorageType == StorageType.String:
                        val = p.AsString() or ""
                    elif p.StorageType == StorageType.Integer:
                        val = str(p.AsInteger())
                    elif p.StorageType == StorageType.Double:
                        val = str(p.AsDouble())
                except:
                    try:
                        val = p.AsValueString() or ""
                    except:
                        val = ""
                
                # Show predicted numbering in preview
                current_number = start + i
                formatted_number = self._format_number(current_number, digit_count)
                pref = prefix_map.get(el.Id, "")
                predicted_value = pref + formatted_number if pref else formatted_number
                
                numbered = "ใช่" if val else "ไม่"
                self.grid_preview.Rows.Add(pid, name, predicted_value, numbered)
                
            mode = "โหมดวิว" if self.use_view_filter else "โหมดโครงการ"
            self._set_status("โหลดตัวอย่างเรียบร้อยแล้ว: {} องค์ประกอบ ({}, {} หลัก)".format(len(filtered), mode, digit_count))
        except Exception as ex:
            self._set_status("โหลดตัวอย่างล้มเหลว: " + str(ex), True)

    # ---------- highlight/zoom/select ----------
    def _highlight_zoom(self, s, e):
        try:
            if not self.preview_elements:
                MessageBox.Show("กรุณาโหลดตัวอย่างก่อน", "ข้อมูล")
                return
            ids = List[ElementId]()
            for el in self.preview_elements:
                ids.Add(el.Id)
            
            # For Revit 2026, use SelectElements instead of SetElementIds for better compatibility
            try:
                # Try the new method first
                uidoc.Selection.SetElementIds(ids)
            except:
                # Fallback to older method
                selection_ids = List[ElementId]([el.Id for el in self.preview_elements])
                uidoc.Selection.SetElementIds(selection_ids)
            
            # try ShowElements with collection
            try:
                uidoc.ShowElements(ids)
            except:
                # fallback: call ShowElements per element
                for el in self.preview_elements[:40]:
                    try:
                        uidoc.ShowElements(el)
                    except:
                        pass
            self._set_status("เลือกและซูม {} องค์ประกอบ".format(len(self.preview_elements)))
        except Exception as ex:
            self._set_status("ไฮไลต์/ซูมล้มเหลว: " + str(ex), True)

    def _clear_highlight(self, s, e):
        try:
            uidoc.Selection.SetElementIds(List[ElementId]())
            self._set_status("ล้างไฮไลต์เรียบร้อยแล้ว")
        except Exception as ex:
            self._set_status("ล้างไฮไลต์ล้มเหลว: " + str(ex), True)

    # ---------- preview double click => manual edit ----------
    def _preview_double_click(self, sender, args):
        try:
            if not self.grid_preview.SelectedRows or self.grid_preview.SelectedRows.Count == 0:
                return
            row = self.grid_preview.SelectedRows[0]
            eid = int(row.Cells["Id"].Value)
            el = doc.GetElement(ElementId(eid))
            param = self._get_param()
            if not is_writable_param(el, param):
                MessageBox.Show("พารามิเตอร์นี้ไม่สามารถเขียนได้ในองค์ประกอบนี้", "ข้อมูล")
                return
            val = Interaction.InputBox("ป้อนค่าสำหรับองค์ประกอบ {}:".format(eid), "กำหนดหมายเลขด้วยตนเอง", "")
            if val is None: 
                return
            t = Transaction(doc, "กำหนดหมายเลขด้วยตนเอง")
            t.Start()
            try:
                p = el.LookupParameter(param)
                if p.StorageType == StorageType.String: 
                    p.Set(val)
                elif p.StorageType == StorageType.Integer:
                    try: 
                        p.Set(int(val))
                    except: 
                        try: 
                            p.Set(val)
                        except: 
                            pass
                elif p.StorageType == StorageType.Double:
                    try: 
                        p.Set(float(val))
                    except:
                        try: 
                            p.Set(val)
                        except: 
                            pass
                row.Cells["Value"].Value = val
                row.Cells["Numbered"].Value = "ใช่"
                self._set_status("กำหนดด้วยตนเองสำหรับองค์ประกอบ {}".format(eid))
            finally:
                t.Commit()
        except Exception as ex:
            self._set_status("แก้ไขด้วยตนเองล้มเหลว: " + str(ex), True)

    # ---------- Step & Run All ----------
    def _step(self, s, e):
        try:
            if not self.preview_elements:
                MessageBox.Show("กรุณาโหลดตัวอย่างก่อน", "ข้อมูล")
                return
            if self.grid_preview.SelectedRows.Count == 0:
                MessageBox.Show("เลือกแถวตัวอย่างเพื่อเริ่มขั้นตอน", "ข้อมูล")
                return
            idx = self.grid_preview.SelectedRows[0].Index
            try:
                start = int(self.txt_start.Text)
            except:
                start = 1
            current = start + idx
            digit_count = self._get_digit_count()
            formatted_number = self._format_number(current, digit_count)
            
            param = self._get_param()
            el = self.preview_elements[idx]
            if not is_writable_param(el, param):
                MessageBox.Show("พารามิเตอร์นี้ไม่สามารถเขียนได้สำหรับองค์ประกอบนี้", "ข้อมูล")
                return
            
            # Get prefix for this element
            prefix_map = {}
            for r in self.grid_cat.Rows:
                try:
                    if r.Cells["Use"].Value:
                        name = r.Cells["Category"].Value
                        pref = r.Cells["Prefix"].Value if r.Cells["Prefix"].Value else DEFAULT_PREFIX_MAP.get(name,"X")
                        # Use active view for collecting elements if view filter is enabled
                        active_view = self.active_view if self.use_view_filter else None
                        for el_cat in collect_elements_for_category(name, active_view):
                            prefix_map[el_cat.Id] = pref
                except:
                    continue
            
            pref = prefix_map.get(el.Id, "")
            value = pref + formatted_number if pref else formatted_number
            
            t = Transaction(doc, "ขั้นตอนการกำหนดหมายเลข")
            t.Start()
            try:
                p = el.LookupParameter(param)
                if p.StorageType == StorageType.String: 
                    p.Set(value)
                elif p.StorageType == StorageType.Integer: 
                    try:
                        p.Set(int(current))
                    except:
                        p.Set(value)
                else:
                    try: 
                        p.Set(value)
                    except: 
                        pass
                row = self.grid_preview.Rows[idx]
                row.Cells["Value"].Value = value
                row.Cells["Numbered"].Value = "ใช่"
                self._set_status("ขั้นตอนองค์ประกอบ {} = {} ({} หลัก)".format(get_element_id_value(el.Id), value, digit_count))
            finally:
                t.Commit()
        except Exception as ex:
            self._set_status("ขั้นตอนล้มเหลว: " + str(ex), True)

    def _run_all(self, s, e):
        try:
            if not self.preview_elements:
                MessageBox.Show("กรุณาโหลดตัวอย่างก่อน", "ข้อมูล")
                return
            param = self._get_param()
            try:
                start = int(self.txt_start.Text)
            except:
                start = 1
            digit_count = self._get_digit_count()
            mode = "CONTINUOUS" if self.rb_cont.Checked else "PER_CATEGORY"

            prefix_map = {}
            for r in self.grid_cat.Rows:
                try:
                    if r.Cells["Use"].Value:
                        name = r.Cells["Category"].Value
                        pref = r.Cells["Prefix"].Value if r.Cells["Prefix"].Value else DEFAULT_PREFIX_MAP.get(name,"X")
                        # Use active view for collecting elements if view filter is enabled
                        active_view = self.active_view if self.use_view_filter else None
                        for el in collect_elements_for_category(name, active_view):
                            prefix_map[el.Id] = pref
                except:
                    continue
            for el in self.preview_elements:
                if el.Id not in prefix_map:
                    cn = get_element_category(el)
                    prefix_map[el.Id] = (cn[0].upper() if cn else "X")

            t = Transaction(doc, "กำหนดหมายเลขอัตโนมัติ")
            t.Start()
            try:
                if mode == "CONTINUOUS":
                    counter = start
                    for el in self.preview_elements:
                        p = el.LookupParameter(param)
                        if p and not p.IsReadOnly:
                            pref = prefix_map.get(el.Id, "")
                            formatted_number = self._format_number(counter, digit_count)
                            value = pref + formatted_number if pref else formatted_number
                            if p.StorageType == StorageType.String:
                                p.Set(value)
                            elif p.StorageType == StorageType.Integer:
                                try: 
                                    p.Set(int(counter))
                                except: 
                                    p.Set(value)
                            else:
                                try: 
                                    p.Set(value)
                                except: 
                                    pass
                            counter += 1
                else:
                    counters = {}
                    for el in self.preview_elements:
                        p = el.LookupParameter(param)
                        if p and not p.IsReadOnly:
                            pref = prefix_map.get(el.Id, "")
                            if pref not in counters:
                                counters[pref] = start
                            formatted_number = self._format_number(counters[pref], digit_count)
                            value = pref + formatted_number if pref else formatted_number
                            if p.StorageType == StorageType.String:
                                p.Set(value)
                            elif p.StorageType == StorageType.Integer:
                                try: 
                                    p.Set(int(counters[pref]))
                                except: 
                                    p.Set(value)
                            else:
                                try: 
                                    p.Set(value)
                                except: 
                                    pass
                            counters[pref] += 1

                # update preview grid
                for i, el in enumerate(self.preview_elements):
                    row = self.grid_preview.Rows[i]
                    p = el.LookupParameter(param)
                    try:
                        if p.StorageType == StorageType.String:
                            row.Cells["Value"].Value = p.AsString() or ""
                        else:
                            row.Cells["Value"].Value = p.AsValueString() or ""
                    except:
                        try: 
                            row.Cells["Value"].Value = p.AsValueString()
                        except: 
                            row.Cells["Value"].Value = ""
                    row.Cells["Numbered"].Value = "ใช่" if row.Cells["Value"].Value else "ไม่"
                
                view_mode = "โหมดวิว" if self.use_view_filter else "โหมดโครงการ"
                self._set_status("รันทั้งหมดเสร็จสิ้น ({} องค์ประกอบ, {} หลัก, {})".format(len(self.preview_elements), digit_count, view_mode))
            finally:
                t.Commit()
        except Exception as ex:
            self._set_status("รันทั้งหมดล้มเหลว: " + str(ex), True)

    def _show_results(self, s, e):
        try:
            msg = "ผลลัพธ์ตัวอย่าง:\n\n"
            for r in self.grid_preview.Rows:
                try:
                    eid = r.Cells["Id"].Value
                    name = r.Cells["Name"].Value
                    val = r.Cells["Value"].Value or ""
                    msg += "{} ({}): {}\n".format(eid, name, val)
                except:
                    continue
            rf = ResultForm("ผลลัพธ์ตัวอย่าง", msg)
            rf.ShowDialog()
        except Exception as ex:
            self._set_status("แสดงผลลัพธ์ล้มเหลว: " + str(ex), True)

    def _set_status(self, txt, is_error=False):
        self.lbl_status.Text = txt
        self.lbl_status.ForeColor = Color.Red if is_error else Color.FromArgb(0, 102, 0)
        self.Refresh()

# ---------------- main ----------------
def main():
    try:
        form = NumberingForm()
        form.ShowDialog()
    except Exception as ex:
        MessageBox.Show("ข้อผิดพลาดร้ายแรง: " + str(ex), "ข้อผิดพลาด")

if __name__ == "__main__":
    main()