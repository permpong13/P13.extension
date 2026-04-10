# -*- coding: utf-8 -*-
"""Workset Manager Tool"""
__title__ = 'Workset\nManager'
__author__ = 'เพิ่มพงษ์'

import clr
import traceback
from collections import defaultdict
from datetime import datetime

# CLR references
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('System')

from System import EventArgs
from System.ComponentModel import BackgroundWorker
from System.Windows.Forms import (
    DialogResult, SaveFileDialog,
    Form, Button, Label, ComboBox, GroupBox, CheckBox,
    ListView, View, ColumnHeader, TabControl, TabPage,
    MessageBox, MessageBoxButtons, MessageBoxIcon,
    BorderStyle, AnchorStyles, SelectionMode,
    TextBox, FormStartPosition, ComboBoxStyle,
    ColumnHeaderStyle, HorizontalAlignment, ProgressBar,
    ProgressBarStyle
)
from System.Drawing import (
    Point, Size, Font, FontStyle, Color, SystemColors,
    ContentAlignment
)

# pyRevit / Revit API
from pyrevit import revit, DB, UI
from pyrevit import forms
from pyrevit import script

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
output = script.get_output()

class WorksetManager:
    """Workset Manager Core Logic"""

    @staticmethod
    def get_all_worksets():
        """Get all worksets in the project"""
        try:
            if not doc.IsWorkshared:
                return []

            worksets = []
            workset_table = doc.GetWorksetTable()

            # Get all workset IDs
            workset_ids = DB.FilteredWorksetCollector(doc).ToWorksets()

            for workset in workset_ids:
                if workset.Kind == DB.WorksetKind.UserWorkset:
                    worksets.append(workset)

            return worksets
        except Exception as e:
            logger.error("Error getting worksets: {}".format(e))
            return []

    @staticmethod
    def get_element_count_by_workset():
        """Count elements in each workset"""
        element_count = defaultdict(int)
        try:
            collector = DB.FilteredElementCollector(doc)
            elements = collector.WhereElementIsNotElementType().ToElements()

            for element in elements:
                try:
                    workset_id = element.WorksetId
                    if workset_id != DB.WorksetId.InvalidWorksetId:
                        element_count[workset_id.IntegerValue] += 1
                except:
                    continue

        except Exception as e:
            logger.error("Error counting elements: {}".format(e))

        return element_count

    @staticmethod
    def get_elements_in_workset(workset):
        """Get all elements in a specific workset"""
        elements = []
        try:
            collector = DB.FilteredElementCollector(doc)
            all_elements = collector.WhereElementIsNotElementType().ToElements()

            for element in all_elements:
                try:
                    if element.WorksetId.IntegerValue == workset.Id.IntegerValue:
                        elements.append(element)
                except:
                    continue

        except Exception as e:
            logger.error("Error getting elements in workset: {}".format(e))

        return elements

    @staticmethod
    def move_elements_to_workset(elements, target_workset):
        """Move elements to target workset"""
        success_count = 0
        failed_count = 0
        failed_elements = []

        try:
            with revit.Transaction("Move Elements to Workset"):
                for element in elements:
                    try:
                        # Check if element can change workset
                        if element.WorksetId != target_workset.Id:
                            param = element.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
                            if param and not param.IsReadOnly:
                                param.Set(target_workset.Id.IntegerValue)
                                success_count += 1
                            else:
                                failed_count += 1
                                failed_elements.append(element)
                    except Exception as e:
                        failed_count += 1
                        failed_elements.append(element)
                        logger.error("Error moving element: {}".format(e))

        except Exception as e:
            logger.error("Error in move transaction: {}".format(e))

        return success_count, failed_count, failed_elements

    @staticmethod
    def create_workset(name):
        """Create new workset"""
        try:
            with revit.Transaction("Create Workset"):
                new_workset = DB.Workset.Create(doc, name)
                return new_workset is not None
        except Exception as e:
            logger.error("Error creating workset: {}".format(e))
            return False

    @staticmethod
    def delete_workset(workset):
        """Delete workset - UPDATED VERSION WITH DIFFERENT ATTRIBUTE NAMES"""
        try:
            with revit.Transaction("Delete Workset"):
                # สร้าง DeleteWorksetSettings object
                settings = DB.DeleteWorksetSettings()
                
                # ลองใช้ชื่อ attribute ที่ต่างกัน
                # สำหรับ Revit 2020+ อาจใช้ชื่อเหล่านี้:
                if hasattr(settings, 'AllowDeletingLastWorkset'):
                    settings.AllowDeletingLastWorkset = False  # ไม่ลบ Workset สุดท้าย
                elif hasattr(settings, 'DeleteLastWorkset'):
                    settings.DeleteLastWorkset = False  # ไม่ลบ Workset สุดท้าย
                
                if hasattr(settings, 'AllowDeletingWorksetWithElements'):
                    settings.AllowDeletingWorksetWithElements = False  # ไม่ลบ Workset ที่มีองค์ประกอบ
                elif hasattr(settings, 'DeleteWorksetWithElements'):
                    settings.DeleteWorksetWithElements = False  # ไม่ลบ Workset ที่มีองค์ประกอบ
                
                # ใช้ static method ของ WorksetTable ด้วย DeleteWorksetSettings
                DB.WorksetTable.DeleteWorkset(doc, workset.Id, settings)
                return True
        except Exception as e:
            logger.error("Error deleting workset: {}".format(e))
            return False

    @staticmethod
    def can_delete_workset(workset):
        """Check if workset can be deleted - SIMPLIFIED VERSION"""
        try:
            # วิธีที่ง่ายที่สุด - ตรวจสอบว่า workset ว่างและไม่ใช่ workset เริ่มต้น
            workset_table = doc.GetWorksetTable()
            element_counts = WorksetManager.get_element_count_by_workset()
            default_workset = WorksetManager.get_default_workset()

            # ตรวจสอบจำนวนองค์ประกอบ
            element_count = element_counts.get(workset.Id.IntegerValue, 0)
            if element_count > 0:
                return False, "Workset มีองค์ประกอบอยู่"

            # ตรวจสอบว่าเป็น workset เริ่มต้นหรือไม่
            if default_workset and workset.Id.IntegerValue == default_workset.Id.IntegerValue:
                return False, "ไม่สามารถลบ Workset เริ่มต้นได้"

            # ตรวจสอบว่าเป็น workset สุดท้ายหรือไม่
            all_worksets = WorksetManager.get_all_worksets()
            user_worksets = [w for w in all_worksets if w.Kind == DB.WorksetKind.UserWorkset]
            if len(user_worksets) <= 1:
                return False, "ไม่สามารถลบ Workset สุดท้ายได้"

            return True, "สามารถลบได้"

        except Exception as e:
            logger.error("Error checking if workset can be deleted: {}".format(e))
            return False, "เกิดข้อผิดพลาดในการตรวจสอบ"

    @staticmethod
    def rename_workset(workset, new_name):
        """Rename workset"""
        try:
            with revit.Transaction("Rename Workset"):
                workset_table = doc.GetWorksetTable()
                workset_table.RenameWorkset(workset.Id, new_name)
                return True
        except Exception as e:
            logger.error("Error renaming workset: {}".format(e))
            return False

    @staticmethod
    def set_default_workset(workset):
        """Set default workset for new elements"""
        try:
            workset_table = doc.GetWorksetTable()
            workset_table.SetActiveWorksetId(workset.Id)
            return True
        except Exception as e:
            logger.error("Error setting default workset: {}".format(e))
            return False

    @staticmethod
    def get_default_workset():
        """Get current default workset"""
        try:
            workset_table = doc.GetWorksetTable()
            active_workset_id = workset_table.GetActiveWorksetId()
            return workset_table.GetWorkset(active_workset_id)
        except Exception as e:
            logger.error("Error getting default workset: {}".format(e))
            return None

    @staticmethod
    def select_elements_in_workset(workset):
        """Select all elements in workset"""
        try:
            elements = WorksetManager.get_elements_in_workset(workset)
            if elements:
                element_ids = [element.Id for element in elements]
                uidoc.Selection.SetElementIds(DB.List[DB.ElementId](element_ids))
                return len(elements)
            return 0
        except Exception as e:
            logger.error("Error selecting elements in workset: {}".format(e))
            return 0

    @staticmethod
    def move_selected_elements_to_workset(workset):
        """Move currently selected elements to workset"""
        try:
            selected_elements = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]
            if not selected_elements:
                return 0, 0, []

            return WorksetManager.move_elements_to_workset(selected_elements, workset)
        except Exception as e:
            logger.error("Error moving selected elements: {}".format(e))
            return 0, 0, []

    @staticmethod
    def set_workset_visibility(view, workset, visible):
        """Set workset visibility in view"""
        try:
            workset_id = workset.Id
            workset_visibility = view.GetWorksetVisibility(workset_id)

            if visible:
                if workset_visibility == DB.WorksetVisibility.Hidden:
                    view.SetWorksetVisibility(workset_id, DB.WorksetVisibility.Visible)
            else:
                if workset_visibility == DB.WorksetVisibility.Visible:
                    view.SetWorksetVisibility(workset_id, DB.WorksetVisibility.Hidden)

            return True
        except Exception as e:
            logger.error("Error setting workset visibility: {}".format(e))
            return False

    @staticmethod
    def get_unused_worksets():
        """Get worksets that are not used (no elements and not default)"""
        try:
            all_worksets = WorksetManager.get_all_worksets()
            element_counts = WorksetManager.get_element_count_by_workset()
            default_workset = WorksetManager.get_default_workset()

            unused_worksets = []
            for workset in all_worksets:
                element_count = element_counts.get(workset.Id.IntegerValue, 0)
                is_default = default_workset and workset.Id.IntegerValue == default_workset.Id.IntegerValue

                if element_count == 0 and not is_default:
                    unused_worksets.append(workset)

            return unused_worksets
        except Exception as e:
            logger.error("Error getting unused worksets: {}".format(e))
            return []

class WorksetManagerForm(Form):
    def __init__(self):
        self.Text = "Workset Manager"
        self.Size = Size(1100, 800)
        self.StartPosition = FormStartPosition.CenterScreen
        self.Font = Font("Microsoft Sans Serif", 9)
        self.BackColor = SystemColors.Control
        self.MinimumSize = Size(1000, 700)

        # Data
        self.all_worksets = []
        self.worksets_with_elements = []
        self.empty_worksets = []
        self.element_counts = {}

        # Search boxes
        self.search_box_all = None
        self.search_box_with = None
        self.search_box_empty = None

        # Initialize UI
        self.InitializeComponents()

        # Load data
        self.LoadWorksets()

    def InitializeComponents(self):
        # Title
        title_label = Label()
        title_label.Text = "Workset Manager - จัดการ Worksets"
        title_label.Location = Point(20, 15)
        title_label.Size = Size(1060, 30)
        title_label.Font = Font("Microsoft Sans Serif", 14, FontStyle.Bold)
        title_label.ForeColor = Color.DarkBlue
        title_label.TextAlign = ContentAlignment.MiddleCenter
        self.Controls.Add(title_label)

        # Main Tab Control
        self.tab_control = TabControl()
        self.tab_control.Location = Point(20, 50)
        self.tab_control.Size = Size(1040, 400)

        # Tab 1: All Worksets
        self.tab_all = TabPage()
        self.tab_all.Text = "Worksets ทั้งหมด"

        # Tab 2: Worksets with Elements
        self.tab_with_elements = TabPage()
        self.tab_with_elements.Text = "Worksets ที่มีองค์ประกอบ"

        # Tab 3: Empty Worksets
        self.tab_empty = TabPage()
        self.tab_empty.Text = "Worksets ว่าง"

        self.tab_control.Controls.Add(self.tab_all)
        self.tab_control.Controls.Add(self.tab_with_elements)
        self.tab_control.Controls.Add(self.tab_empty)
        self.Controls.Add(self.tab_control)

        # Initialize tabs
        self.InitializeAllWorksetsTab()
        self.InitializeWithElementsTab()
        self.InitializeEmptyWorksetsTab()

        # Action buttons
        self.InitializeActionButtons()

    def InitializeAllWorksetsTab(self):
        # Filter section
        filter_group = GroupBox()
        filter_group.Text = "ตัวกรอง"
        filter_group.Location = Point(20, 20)
        filter_group.Size = Size(990, 60)
        self.tab_all.Controls.Add(filter_group)

        # Search box
        search_label = Label()
        search_label.Text = "ค้นหา:"
        search_label.Location = Point(20, 25)
        search_label.Size = Size(50, 20)
        filter_group.Controls.Add(search_label)

        self.search_box_all = TextBox()
        self.search_box_all.Location = Point(80, 22)
        self.search_box_all.Size = Size(200, 25)
        self.search_box_all.Tag = "all"
        self.search_box_all.TextChanged += self.OnSearchTextChanged
        filter_group.Controls.Add(self.search_box_all)

        # Refresh button
        refresh_btn = Button()
        refresh_btn.Text = "รีเฟรช"
        refresh_btn.Location = Point(300, 20)
        refresh_btn.Size = Size(80, 30)
        refresh_btn.Click += self.RefreshWorksets
        filter_group.Controls.Add(refresh_btn)

        # Clear search button
        clear_btn = Button()
        clear_btn.Text = "ล้าง"
        clear_btn.Location = Point(390, 20)
        clear_btn.Size = Size(80, 30)
        clear_btn.Tag = "all"
        clear_btn.Click += self.OnClearSearchClicked
        filter_group.Controls.Add(clear_btn)

        # Default workset info
        self.default_workset_label = Label()
        self.default_workset_label.Location = Point(490, 25)
        self.default_workset_label.Size = Size(300, 20)
        self.default_workset_label.Text = "กำลังโหลด..."
        filter_group.Controls.Add(self.default_workset_label)

        # Worksets list
        self.all_worksets_list = ListView()
        self.all_worksets_list.Location = Point(20, 90)
        self.all_worksets_list.Size = Size(990, 250)
        self.all_worksets_list.View = View.Details
        self.all_worksets_list.FullRowSelect = True
        self.all_worksets_list.GridLines = True
        self.all_worksets_list.MultiSelect = True
        self.all_worksets_list.CheckBoxes = True

        # Add columns
        self.all_worksets_list.Columns.Add("เลือก", 50)
        self.all_worksets_list.Columns.Add("ชื่อ Workset", 250)
        self.all_worksets_list.Columns.Add("จำนวนองค์ประกอบ", 120)
        self.all_worksets_list.Columns.Add("สถานะ", 100)
        self.all_worksets_list.Columns.Add("เจ้าของ", 150)
        self.all_worksets_list.Columns.Add("ID", 80)

        self.tab_all.Controls.Add(self.all_worksets_list)

        # Selection buttons
        select_all_btn = Button()
        select_all_btn.Text = "เลือกทั้งหมด"
        select_all_btn.Location = Point(20, 350)
        select_all_btn.Size = Size(100, 30)
        select_all_btn.Click += lambda s, e: self.SelectAllItems(self.all_worksets_list)
        self.tab_all.Controls.Add(select_all_btn)

        select_none_btn = Button()
        select_none_btn.Text = "ไม่เลือกเลย"
        select_none_btn.Location = Point(130, 350)
        select_none_btn.Size = Size(100, 30)
        select_none_btn.Click += lambda s, e: self.SelectNoneItems(self.all_worksets_list)
        self.tab_all.Controls.Add(select_none_btn)

    def InitializeWithElementsTab(self):
        # Filter section
        filter_group = GroupBox()
        filter_group.Text = "ตัวกรอง"
        filter_group.Location = Point(20, 20)
        filter_group.Size = Size(990, 60)
        self.tab_with_elements.Controls.Add(filter_group)

        search_label = Label()
        search_label.Text = "ค้นหา:"
        search_label.Location = Point(20, 25)
        search_label.Size = Size(50, 20)
        filter_group.Controls.Add(search_label)

        self.search_box_with = TextBox()
        self.search_box_with.Location = Point(80, 22)
        self.search_box_with.Size = Size(200, 25)
        self.search_box_with.Tag = "with"
        self.search_box_with.TextChanged += self.OnSearchTextChanged
        filter_group.Controls.Add(self.search_box_with)

        clear_btn = Button()
        clear_btn.Text = "ล้าง"
        clear_btn.Location = Point(300, 20)
        clear_btn.Size = Size(80, 30)
        clear_btn.Tag = "with"
        clear_btn.Click += self.OnClearSearchClicked
        filter_group.Controls.Add(clear_btn)

        # Worksets with elements list
        self.with_elements_list = ListView()
        self.with_elements_list.Location = Point(20, 90)
        self.with_elements_list.Size = Size(990, 250)
        self.with_elements_list.View = View.Details
        self.with_elements_list.FullRowSelect = True
        self.with_elements_list.GridLines = True
        self.with_elements_list.MultiSelect = True
        self.with_elements_list.CheckBoxes = True

        self.with_elements_list.Columns.Add("เลือก", 50)
        self.with_elements_list.Columns.Add("ชื่อ Workset", 250)
        self.with_elements_list.Columns.Add("จำนวนองค์ประกอบ", 120)
        self.with_elements_list.Columns.Add("สถานะ", 100)
        self.with_elements_list.Columns.Add("เจ้าของ", 150)
        self.with_elements_list.Columns.Add("ID", 80)

        self.tab_with_elements.Controls.Add(self.with_elements_list)

        # Selection buttons
        select_all_btn = Button()
        select_all_btn.Text = "เลือกทั้งหมด"
        select_all_btn.Location = Point(20, 350)
        select_all_btn.Size = Size(100, 30)
        select_all_btn.Click += lambda s, e: self.SelectAllItems(self.with_elements_list)
        self.tab_with_elements.Controls.Add(select_all_btn)

        select_none_btn = Button()
        select_none_btn.Text = "ไม่เลือกเลย"
        select_none_btn.Location = Point(130, 350)
        select_none_btn.Size = Size(100, 30)
        select_none_btn.Click += lambda s, e: self.SelectNoneItems(self.with_elements_list)
        self.tab_with_elements.Controls.Add(select_none_btn)

    def InitializeEmptyWorksetsTab(self):
        # Filter section
        filter_group = GroupBox()
        filter_group.Text = "ตัวกรอง"
        filter_group.Location = Point(20, 20)
        filter_group.Size = Size(990, 60)
        self.tab_empty.Controls.Add(filter_group)

        search_label = Label()
        search_label.Text = "ค้นหา:"
        search_label.Location = Point(20, 25)
        search_label.Size = Size(50, 20)
        filter_group.Controls.Add(search_label)

        self.search_box_empty = TextBox()
        self.search_box_empty.Location = Point(80, 22)
        self.search_box_empty.Size = Size(200, 25)
        self.search_box_empty.Tag = "empty"
        self.search_box_empty.TextChanged += self.OnSearchTextChanged
        filter_group.Controls.Add(self.search_box_empty)

        clear_btn = Button()
        clear_btn.Text = "ล้าง"
        clear_btn.Location = Point(300, 20)
        clear_btn.Size = Size(80, 30)
        clear_btn.Tag = "empty"
        clear_btn.Click += self.OnClearSearchClicked
        filter_group.Controls.Add(clear_btn)

        # Empty worksets list
        self.empty_worksets_list = ListView()
        self.empty_worksets_list.Location = Point(20, 90)
        self.empty_worksets_list.Size = Size(990, 250)
        self.empty_worksets_list.View = View.Details
        self.empty_worksets_list.FullRowSelect = True
        self.empty_worksets_list.GridLines = True
        self.empty_worksets_list.MultiSelect = True
        self.empty_worksets_list.CheckBoxes = True

        self.empty_worksets_list.Columns.Add("เลือก", 50)
        self.empty_worksets_list.Columns.Add("ชื่อ Workset", 300)
        self.empty_worksets_list.Columns.Add("สถานะ", 100)
        self.empty_worksets_list.Columns.Add("เจ้าของ", 200)
        self.empty_worksets_list.Columns.Add("ID", 80)

        self.tab_empty.Controls.Add(self.empty_worksets_list)

        # Selection buttons
        select_all_btn = Button()
        select_all_btn.Text = "เลือกทั้งหมด"
        select_all_btn.Location = Point(20, 350)
        select_all_btn.Size = Size(100, 30)
        select_all_btn.Click += lambda s, e: self.SelectAllItems(self.empty_worksets_list)
        self.tab_empty.Controls.Add(select_all_btn)

        select_none_btn = Button()
        select_none_btn.Text = "ไม่เลือกเลย"
        select_none_btn.Location = Point(130, 350)
        select_none_btn.Size = Size(100, 30)
        select_none_btn.Click += lambda s, e: self.SelectNoneItems(self.empty_worksets_list)
        self.tab_empty.Controls.Add(select_none_btn)

    def InitializeActionButtons(self):
        # Action group
        action_group = GroupBox()
        action_group.Text = "การดำเนินการ"
        action_group.Location = Point(20, 460)
        action_group.Size = Size(1040, 300)
        self.Controls.Add(action_group)

        # Row 1 - Basic Actions
        self.create_btn = Button()
        self.create_btn.Text = "สร้าง Workset ใหม่"
        self.create_btn.Location = Point(20, 30)
        self.create_btn.Size = Size(120, 35)
        self.create_btn.Click += self.CreateWorkset
        action_group.Controls.Add(self.create_btn)

        self.delete_btn = Button()
        self.delete_btn.Text = "ลบ Workset ที่เลือก"
        self.delete_btn.Location = Point(150, 30)
        self.delete_btn.Size = Size(120, 35)
        self.delete_btn.Click += self.DeleteWorksets
        action_group.Controls.Add(self.delete_btn)

        self.rename_btn = Button()
        self.rename_btn.Text = "เปลี่ยนชื่อ Workset"
        self.rename_btn.Location = Point(280, 30)
        self.rename_btn.Size = Size(120, 35)
        self.rename_btn.Click += self.RenameWorkset
        action_group.Controls.Add(self.rename_btn)

        # Row 2 - Element Actions
        self.select_elements_btn = Button()
        self.select_elements_btn.Text = "เลือกองค์ประกอบใน Workset"
        self.select_elements_btn.Location = Point(410, 30)
        self.select_elements_btn.Size = Size(150, 35)
        self.select_elements_btn.BackColor = Color.LightBlue
        self.select_elements_btn.Click += self.SelectElementsInWorkset
        action_group.Controls.Add(self.select_elements_btn)

        self.move_selected_btn = Button()
        self.move_selected_btn.Text = "ย้ายองค์ประกอบที่เลือก"
        self.move_selected_btn.Location = Point(570, 30)
        self.move_selected_btn.Size = Size(150, 35)
        self.move_selected_btn.BackColor = Color.LightGreen
        self.move_selected_btn.Click += self.MoveSelectedElementsToWorkset
        action_group.Controls.Add(self.move_selected_btn)

        # Row 3 - Settings
        self.set_default_btn = Button()
        self.set_default_btn.Text = "กำหนดเป็น Workset เริ่มต้น"
        self.set_default_btn.Location = Point(20, 75)
        self.set_default_btn.Size = Size(160, 35)
        self.set_default_btn.Click += self.SetDefaultWorkset
        action_group.Controls.Add(self.set_default_btn)

        self.cleanup_btn = Button()
        self.cleanup_btn.Text = "ทำความสะอาด Worksets"
        self.cleanup_btn.Location = Point(190, 75)
        self.cleanup_btn.Size = Size(140, 35)
        self.cleanup_btn.BackColor = Color.LightYellow
        self.cleanup_btn.Click += self.CleanupUnusedWorksets
        action_group.Controls.Add(self.cleanup_btn)

        # Row 4 - Move Elements
        self.move_btn = Button()
        self.move_btn.Text = "ย้ายองค์ประกอบระหว่าง Workset"
        self.move_btn.Location = Point(340, 75)
        self.move_btn.Size = Size(180, 35)
        self.move_btn.BackColor = Color.LightGreen
        self.move_btn.Click += self.MoveElementsBetweenWorksets
        action_group.Controls.Add(self.move_btn)

        # Close button
        close_btn = Button()
        close_btn.Text = "ปิด"
        close_btn.Location = Point(900, 30)
        close_btn.Size = Size(100, 80)
        close_btn.BackColor = Color.LightCoral
        close_btn.Click += self.CloseForm
        action_group.Controls.Add(close_btn)

        # Status label
        self.status_label = Label()
        self.status_label.Location = Point(20, 165)
        self.status_label.Size = Size(980, 25)
        self.status_label.BorderStyle = BorderStyle.FixedSingle
        self.status_label.Text = "พร้อมทำงาน"
        self.status_label.TextAlign = ContentAlignment.MiddleLeft
        action_group.Controls.Add(self.status_label)

        # Progress bar for move operation
        self.progress_bar = ProgressBar()
        self.progress_bar.Location = Point(20, 195)
        self.progress_bar.Size = Size(980, 20)
        self.progress_bar.Visible = False
        action_group.Controls.Add(self.progress_bar)

        # Progress label
        self.progress_label = Label()
        self.progress_label.Location = Point(20, 220)
        self.progress_label.Size = Size(980, 20)
        self.progress_label.Text = ""
        self.progress_label.TextAlign = ContentAlignment.MiddleLeft
        action_group.Controls.Add(self.progress_label)

    def LoadWorksets(self):
        """Load all worksets and their element counts"""
        try:
            self.status_label.Text = "กำลังโหลด Worksets..."

            # Get all worksets
            self.all_worksets = WorksetManager.get_all_worksets()

            # Get element counts
            self.element_counts = WorksetManager.get_element_count_by_workset()

            # Categorize worksets
            self.worksets_with_elements = []
            self.empty_worksets = []

            for workset in self.all_worksets:
                element_count = self.element_counts.get(workset.Id.IntegerValue, 0)
                if element_count > 0:
                    self.worksets_with_elements.append(workset)
                else:
                    self.empty_worksets.append(workset)

            # Update UI
            self.UpdateAllLists()

            # Update default workset info
            self.UpdateDefaultWorksetInfo()

            self.status_label.Text = "โหลด Worksets สำเร็จ: {} worksets".format(len(self.all_worksets))

        except Exception as e:
            logger.error("Error loading worksets: {}".format(e))
            self.status_label.Text = "ข้อผิดพลาดในการโหลด Worksets"

    def UpdateDefaultWorksetInfo(self):
        """Update default workset information"""
        try:
            default_workset = WorksetManager.get_default_workset()
            if default_workset:
                self.default_workset_label.Text = "Workset เริ่มต้น: {}".format(default_workset.Name)
                self.default_workset_label.ForeColor = Color.DarkGreen
            else:
                self.default_workset_label.Text = "ไม่พบ Workset เริ่มต้น"
                self.default_workset_label.ForeColor = Color.DarkRed
        except Exception as e:
            logger.error("Error updating default workset info: {}".format(e))
            self.default_workset_label.Text = "ข้อผิดพลาดในการโหลด Workset เริ่มต้น"

    def UpdateAllLists(self):
        """Update all list views"""
        self.UpdateAllWorksetsList()
        self.UpdateWithElementsList()
        self.UpdateEmptyWorksetsList()

    def _match_search(self, workset, search_text):
        """Return True if workset matches search_text (by name, owner, or ID)"""
        if not search_text:
            return True
        search_text = search_text.lower()
        owner = workset.Owner if workset.Owner else ""
        ws_id = str(workset.Id.IntegerValue)
        combined = u"{} {} {}".format(workset.Name, owner, ws_id).lower()
        return search_text in combined

    def UpdateAllWorksetsList(self):
        """Update the all worksets list view"""
        self.all_worksets_list.Items.Clear()

        default_workset = WorksetManager.get_default_workset()
        search_text = ""
        if self.search_box_all and self.search_box_all.Text:
            search_text = self.search_box_all.Text.strip()

        for workset in self.all_worksets:
            if not self._match_search(workset, search_text):
                continue

            element_count = self.element_counts.get(workset.Id.IntegerValue, 0)
            status = "แก้ไขได้" if workset.IsEditable else "ล็อกแล้ว"
            is_default = default_workset and workset.Id.IntegerValue == default_workset.Id.IntegerValue

            item = self.all_worksets_list.Items.Add("")
            item.SubItems.Add(workset.Name)
            item.SubItems.Add(str(element_count))

            if is_default:
                item.SubItems.Add("เริ่มต้น")
                item.BackColor = Color.LightGreen
            else:
                item.SubItems.Add(status)

            item.SubItems.Add(workset.Owner if workset.Owner else "N/A")
            item.SubItems.Add(str(workset.Id.IntegerValue))
            item.Tag = workset

    def UpdateWithElementsList(self):
        """Update the worksets with elements list view"""
        self.with_elements_list.Items.Clear()

        search_text = ""
        if self.search_box_with and self.search_box_with.Text:
            search_text = self.search_box_with.Text.strip()

        for workset in self.worksets_with_elements:
            if not self._match_search(workset, search_text):
                continue

            element_count = self.element_counts.get(workset.Id.IntegerValue, 0)
            status = "แก้ไขได้" if workset.IsEditable else "ล็อกแล้ว"

            item = self.with_elements_list.Items.Add("")
            item.SubItems.Add(workset.Name)
            item.SubItems.Add(str(element_count))
            item.SubItems.Add(status)
            item.SubItems.Add(workset.Owner if workset.Owner else "N/A")
            item.SubItems.Add(str(workset.Id.IntegerValue))
            item.Tag = workset

    def UpdateEmptyWorksetsList(self):
        """Update the empty worksets list view"""
        self.empty_worksets_list.Items.Clear()

        search_text = ""
        if self.search_box_empty and self.search_box_empty.Text:
            search_text = self.search_box_empty.Text.strip()

        for workset in self.empty_worksets:
            if not self._match_search(workset, search_text):
                continue

            status = "แก้ไขได้" if workset.IsEditable else "ล็อกแล้ว"

            item = self.empty_worksets_list.Items.Add("")
            item.SubItems.Add(workset.Name)
            item.SubItems.Add(status)
            item.SubItems.Add(workset.Owner if workset.Owner else "N/A")
            item.SubItems.Add(str(workset.Id.IntegerValue))
            item.Tag = workset

    def GetSelectedWorksets(self, list_view):
        """Get selected worksets from a list view"""
        selected = []
        for item in list_view.CheckedItems:
            if item.Tag:
                selected.append(item.Tag)
        return selected

    def GetCurrentListView(self):
        """Get the currently active list view based on selected tab"""
        current_tab = self.tab_control.SelectedTab
        if current_tab == self.tab_all:
            return self.all_worksets_list
        elif current_tab == self.tab_with_elements:
            return self.with_elements_list
        elif current_tab == self.tab_empty:
            return self.empty_worksets_list
        return None

    def SelectAllItems(self, list_view):
        """Select all items in a list view"""
        for item in list_view.Items:
            item.Checked = True

    def SelectNoneItems(self, list_view):
        """Deselect all items in a list view"""
        for item in list_view.Items:
            item.Checked = False

    # Event handlers
    def OnSearchTextChanged(self, sender, args):
        """Handle search text changes - shared for all tabs"""
        tag = getattr(sender, "Tag", None)
        if tag == "all":
            self.UpdateAllWorksetsList()
        elif tag == "with":
            self.UpdateWithElementsList()
        elif tag == "empty":
            self.UpdateEmptyWorksetsList()
        else:
            self.UpdateAllLists()

    def OnClearSearchClicked(self, sender, args):
        """Clear search text for the corresponding tab"""
        tag = getattr(sender, "Tag", None)
        if tag == "all" and self.search_box_all:
            self.search_box_all.Text = ""
        elif tag == "with" and self.search_box_with:
            self.search_box_with.Text = ""
        elif tag == "empty" and self.search_box_empty:
            self.search_box_empty.Text = ""

    def RefreshWorksets(self, sender, args):
        """Refresh worksets data"""
        self.LoadWorksets()

    def CreateWorkset(self, sender, args):
        """Create a new workset"""
        try:
            workset_name = forms.ask_for_string(
                prompt="ระบุชื่อ Workset ใหม่:",
                title="สร้าง Workset ใหม่"
            )

            if workset_name:
                if WorksetManager.create_workset(workset_name):
                    self.status_label.Text = "สร้าง Workset '{}' สำเร็จ".format(workset_name)
                    self.LoadWorksets()
                else:
                    self.status_label.Text = "ไม่สามารถสร้าง Workset '{}' ได้".format(workset_name)
        except Exception as e:
            logger.error("Error in CreateWorkset: {}".format(e))
            forms.alert("เกิดข้อผิดพลาดในการสร้าง Workset", title="ข้อผิดพลาด")

    def DeleteWorksets(self, sender, args):
        """Delete selected worksets - CORRECTED VERSION"""
        list_view = self.GetCurrentListView()
        if not list_view:
            return

        selected_worksets = self.GetSelectedWorksets(list_view)
        if not selected_worksets:
            forms.alert("กรุณาเลือก Workset ที่ต้องการลบ", title="ไม่มีการเลือก")
            return

        # Filter out worksets that have elements or are default
        deletable_worksets = []
        non_deletable_worksets = []

        default_workset = WorksetManager.get_default_workset()

        for workset in selected_worksets:
            element_count = self.element_counts.get(workset.Id.IntegerValue, 0)
            is_default = default_workset and workset.Id.IntegerValue == default_workset.Id.IntegerValue

            if element_count == 0 and not is_default:
                deletable_worksets.append(workset)
            else:
                if element_count > 0:
                    non_deletable_worksets.append("{} (มีองค์ประกอบ {} ชิ้น)".format(workset.Name, element_count))
                elif is_default:
                    non_deletable_worksets.append("{} (เป็น Workset เริ่มต้น)".format(workset.Name))

        if non_deletable_worksets:
            forms.alert("ไม่สามารถลบ Workset ต่อไปนี้:\n{}".format(
                "\n".join(non_deletable_worksets)), title="ข้อผิดพลาด")

        if not deletable_worksets:
            return

        # ใช้ options แทน yes/no
        result = forms.alert(
            "คุณแน่ใจหรือไม่ที่จะลบ {} Workset?".format(len(deletable_worksets)), 
            options=["ใช่", "ไม่"],
            title="ยืนยันการลบ"
        )

        if result == "ใช่":
            success_count = 0
            failed_worksets = []

            for workset in deletable_worksets:
                try:
                    # ตรวจสอบว่าสามารถลบได้ก่อน
                    can_delete, reason = WorksetManager.can_delete_workset(workset)
                    if can_delete:
                        if WorksetManager.delete_workset(workset):
                            success_count += 1
                        else:
                            failed_worksets.append("{} (ลบไม่สำเร็จ)".format(workset.Name))
                    else:
                        failed_worksets.append("{} ({})".format(workset.Name, reason))
                except Exception as e:
                    logger.error("Error deleting workset {}: {}".format(workset.Name, e))
                    failed_worksets.append("{} (ข้อผิดพลาด: {})".format(workset.Name, str(e)))

            if failed_worksets:
                forms.alert("ลบ Workset สำเร็จ {} รายการ\nลบไม่สำเร็จ {} รายการ:\n{}".format(
                    success_count, len(failed_worksets), "\n".join(failed_worksets)), title="ผลการลบ")
            else:
                forms.alert("ลบ Workset สำเร็จ {} รายการ".format(success_count), title="สำเร็จ")
                self.status_label.Text = "ลบ Workset สำเร็จ {} รายการ".format(success_count)

            self.LoadWorksets()

    def RenameWorkset(self, sender, args):
        """Rename selected workset"""
        list_view = self.GetCurrentListView()
        if not list_view:
            return

        selected_worksets = self.GetSelectedWorksets(list_view)
        if not selected_worksets:
            forms.alert("กรุณาเลือก Workset ที่ต้องการเปลี่ยนชื่อ", title="ไม่มีการเลือก")
            return

        if len(selected_worksets) > 1:
            forms.alert("กรุณาเลือกเพียงหนึ่ง Workset เท่านั้น", title="ข้อผิดพลาด")
            return

        workset = selected_worksets[0]

        # แก้ไขการเรียกใช้ ask_for_string
        try:
            # ใช้ prompt อย่างเดียว
            new_name = forms.ask_for_string(
                prompt="ระบุชื่อใหม่:\n(ชื่อเดิม: {})".format(workset.Name),
                title="เปลี่ยนชื่อ Workset"
            )

            if not new_name:
                return

        except Exception as e:
            logger.error("Error getting new name: {}".format(e))
            forms.alert("เกิดข้อผิดพลาดในการรับชื่อใหม่", title="ข้อผิดพลาด")
            return

        if new_name and new_name != workset.Name:
            if WorksetManager.rename_workset(workset, new_name):
                self.status_label.Text = "เปลี่ยนชื่อ Workset เป็น '{}' สำเร็จ".format(new_name)
                self.LoadWorksets()
            else:
                self.status_label.Text = "ไม่สามารถเปลี่ยนชื่อ Workset ได้"

    def SelectElementsInWorkset(self, sender, args):
        """Select all elements in selected workset"""
        list_view = self.GetCurrentListView()
        if not list_view:
            return

        selected_worksets = self.GetSelectedWorksets(list_view)
        if not selected_worksets:
            forms.alert("กรุณาเลือก Workset ที่ต้องการเลือกองค์ประกอบ", title="ไม่มีการเลือก")
            return

        if len(selected_worksets) > 1:
            forms.alert("กรุณาเลือกเพียงหนึ่ง Workset เท่านั้น", title="ข้อผิดพลาด")
            return

        workset = selected_worksets[0]
        element_count = WorksetManager.select_elements_in_workset(workset)

        if element_count > 0:
            self.status_label.Text = "เลือกองค์ประกอบใน Workset '{}' สำเร็จ: {} องค์ประกอบ".format(
                workset.Name, element_count)
            forms.alert("เลือกองค์ประกอบ {} รายการใน Workset '{}'".format(
                element_count, workset.Name), title="สำเร็จ")
        else:
            self.status_label.Text = "ไม่พบองค์ประกอบใน Workset '{}'".format(workset.Name)
            forms.alert("ไม่พบองค์ประกอบใน Workset '{}'".format(workset.Name), title="ไม่พบข้อมูล")

    def MoveSelectedElementsToWorkset(self, sender, args):
        """Move currently selected elements to selected workset"""
        list_view = self.GetCurrentListView()
        if not list_view:
            return

        selected_worksets = self.GetSelectedWorksets(list_view)
        if not selected_worksets:
            forms.alert("กรุณาเลือก Workset ปลายทาง", title="ไม่มีการเลือก")
            return

        if len(selected_worksets) > 1:
            forms.alert("กรุณาเลือกเพียงหนึ่ง Workset เท่านั้น", title="ข้อผิดพลาด")
            return

        target_workset = selected_worksets[0]

        # Get selected elements count
        selected_element_ids = uidoc.Selection.GetElementIds()
        if not selected_element_ids or selected_element_ids.Count == 0:
            forms.alert("กรุณาเลือกองค์ประกอบใน Revit ก่อน", title="ไม่มีการเลือก")
            return

        # ใช้ options แทน yes/no  
        result = forms.alert(
            "ย้ายองค์ประกอบที่เลือก {} รายการไปยัง Workset '{}'?".format(
                selected_element_ids.Count, target_workset.Name), 
            options=["ใช่", "ไม่"]
        )

        if result == "ใช่":
            success_count, failed_count, failed_elements = WorksetManager.move_selected_elements_to_workset(target_workset)

            result_msg = "ย้ายองค์ประกอบสำเร็จ: {} รายการ\nย้ายไม่สำเร็จ: {} รายการ".format(
                success_count, failed_count)

            if failed_count > 0:
                result_msg += "\n\nองค์ประกอบที่ย้ายไม่สำเร็จอาจเป็นประเภทที่ไม่สามารถเปลี่ยน Workset ได้"

            forms.alert(result_msg, title="ผลการย้ายองค์ประกอบ")
            self.status_label.Text = "ย้ายองค์ประกอบ {} รายการไปยัง Workset '{}'".format(
                success_count, target_workset.Name)

            self.LoadWorksets()

    def SetDefaultWorkset(self, sender, args):
        """Set selected workset as default"""
        list_view = self.GetCurrentListView()
        if not list_view:
            return

        selected_worksets = self.GetSelectedWorksets(list_view)
        if not selected_worksets:
            forms.alert("กรุณาเลือก Workset ที่ต้องการกำหนดเป็นค่าเริ่มต้น", title="ไม่มีการเลือก")
            return

        if len(selected_worksets) > 1:
            forms.alert("กรุณาเลือกเพียงหนึ่ง Workset เท่านั้น", title="ข้อผิดพลาด")
            return

        workset = selected_worksets[0]

        if WorksetManager.set_default_workset(workset):
            self.status_label.Text = "กำหนด Workset '{}' เป็นค่าเริ่มต้นสำเร็จ".format(workset.Name)
            self.UpdateDefaultWorksetInfo()
            forms.alert("กำหนด Workset '{}' เป็นค่าเริ่มต้นสำหรับองค์ประกอบใหม่สำเร็จ".format(
                workset.Name), title="สำเร็จ")
        else:
            self.status_label.Text = "ไม่สามารถกำหนด Workset เป็นค่าเริ่มต้นได้"

    def CleanupUnusedWorksets(self, sender, args):
        """Clean up unused worksets"""
        unused_worksets = WorksetManager.get_unused_worksets()

        if not unused_worksets:
            forms.alert("ไม่พบ Workset ที่ไม่ได้ใช้งาน", title="ผลการตรวจสอบ")
            return

        # Show unused worksets
        output.print_md("# **Workset ที่ไม่ได้ใช้งาน**")
        output.print_md("**พบ {} Workset ที่ไม่ได้ใช้งาน**".format(len(unused_worksets)))

        for i, workset in enumerate(unused_worksets, 1):
            output.print_md("{}. **{}** (ID: {})".format(i, workset.Name, workset.Id.IntegerValue))

        # ใช้ options แทน yes/no
        result = forms.alert(
            "ต้องการลบ Workset ที่ไม่ได้ใช้งาน {} รายการหรือไม่?".format(len(unused_worksets)), 
            options=["ใช่", "ไม่"], 
            title="ยืนยันการทำความสะอาด"
        )

        if result == "ใช่":
            success_count = 0
            failed_worksets = []

            for workset in unused_worksets:
                try:
                    # ตรวจสอบว่าสามารถลบได้ก่อน
                    can_delete, reason = WorksetManager.can_delete_workset(workset)
                    if can_delete:
                        if WorksetManager.delete_workset(workset):
                            success_count += 1
                        else:
                            failed_worksets.append("{} (ลบไม่สำเร็จ)".format(workset.Name))
                    else:
                        failed_worksets.append("{} ({})".format(workset.Name, reason))
                except Exception as e:
                    logger.error("Error deleting unused workset {}: {}".format(workset.Name, e))
                    failed_worksets.append("{} (ข้อผิดพลาด: {})".format(workset.Name, str(e)))

            if failed_worksets:
                forms.alert("ลบ Workset ที่ไม่ได้ใช้งานสำเร็จ {} รายการ\nลบไม่สำเร็จ {} รายการ:\n{}".format(
                    success_count, len(failed_worksets), "\n".join(failed_worksets)), title="ผลการทำความสะอาด")
            else:
                forms.alert("ลบ Workset ที่ไม่ได้ใช้งานสำเร็จ {} รายการ".format(success_count), title="สำเร็จ")

            self.status_label.Text = "ทำความสะอาด Workset ที่ไม่ได้ใช้งานสำเร็จ {} รายการ".format(success_count)
            self.LoadWorksets()

    def MoveElementsBetweenWorksets(self, sender, args):
        """Move elements from one workset to another"""
        try:
            # ถ้าไม่มี worksets
            if not self.all_worksets:
                forms.alert("ไม่พบ Workset ในโปรเจค", title="ข้อผิดพลาด")
                return

            # ----- เลือก Workset ต้นทาง -----
            # สร้างรายการชื่อ worksets สำหรับ dropdown
            source_names = []
            source_dict = {}
            for ws in self.all_worksets:
                element_count = self.element_counts.get(ws.Id.IntegerValue, 0)
                display_name = "{} ({} elements)".format(ws.Name, element_count)
                source_names.append(display_name)
                source_dict[display_name] = ws

            source_display = forms.ask_for_one_item(
                source_names,
                title="เลือก Workset ต้นทาง",
                prompt="เลือก Workset ต้นทาง:"
            )

            if not source_display:
                return
                
            source_ws = source_dict[source_display]

            # ----- เลือก Workset ปลายทาง -----
            # กรองออก workset ต้นทาง
            target_names = []
            target_dict = {}
            for ws in self.all_worksets:
                if ws.Id.IntegerValue != source_ws.Id.IntegerValue:
                    element_count = self.element_counts.get(ws.Id.IntegerValue, 0)
                    display_name = "{} ({} elements)".format(ws.Name, element_count)
                    target_names.append(display_name)
                    target_dict[display_name] = ws

            if not target_names:
                forms.alert("ไม่มี Workset อื่นให้เลือกเป็นปลายทาง", title="ข้อผิดพลาด")
                return

            target_display = forms.ask_for_one_item(
                target_names,
                title="เลือก Workset ปลายทาง",
                prompt="เลือก Workset ปลายทาง:"
            )

            if not target_display:
                return
                
            target_ws = target_dict[target_display]

            # ----- ดึงรายการ Element ใน Workset ต้นทาง -----
            elements = WorksetManager.get_elements_in_workset(source_ws)
            if not elements:
                forms.alert("ไม่พบองค์ประกอบใน Workset ต้นทาง '{}'".format(source_ws.Name))
                return

            # ----- คอนเฟิร์ม -----
            confirm = forms.alert(
                "ต้องการย้ายองค์ประกอบ {} รายการ\nจาก '{}' ไปยัง '{}' ?".format(
                    len(elements), source_ws.Name, target_ws.Name),
                options=["ย้าย", "ยกเลิก"]
            )

            if confirm != "ย้าย":
                return

            # ----- เริ่มย้าย -----
            with revit.Transaction("Move Elements Between Worksets"):
                success = 0
                fail = 0

                for element in elements:
                    try:
                        param = element.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
                        if param and not param.IsReadOnly:
                            param.Set(target_ws.Id.IntegerValue)
                            success += 1
                        else:
                            fail += 1
                    except:
                        fail += 1
                        continue

            # ----- แสดงผล -----
            msg = ("ย้ายองค์ประกอบเสร็จสิ้น\n"
                   "จาก Workset: {}\n"
                   "ไปยัง: {}\n\n"
                   "สำเร็จ: {}\n"
                   "ไม่สำเร็จ: {}").format(
                source_ws.Name, target_ws.Name, success, fail
            )

            forms.alert(msg, title="ผลการย้ายองค์ประกอบ")

            # รีเฟรชข้อมูล
            self.LoadWorksets()

        except Exception as e:
            logger.error("Error in MoveElementsBetweenWorksets: {}".format(e))
            forms.alert("เกิดข้อผิดพลาดระหว่างย้ายองค์ประกอบ:\n{}".format(str(e)), title="ข้อผิดพลาด")

    def CloseForm(self, sender, args):
        """Close the form"""
        self.DialogResult = DialogResult.OK
        self.Close()

def main():
    try:
        # Check if worksharing is enabled
        if not doc.IsWorkshared:
            forms.alert("โปรเจคนี้ไม่ได้เปิดใช้งาน Worksharing\nไม่สามารถใช้ Workset Manager ได้", 
                       title="ข้อผิดพลาด")
            return

        # Check if there are any worksets
        worksets = WorksetManager.get_all_worksets()
        if not worksets:
            forms.alert("ไม่พบ Workset ในโปรเจคนี้", title="ข้อผิดพลาด")
            return

        form = WorksetManagerForm()
        form.ShowDialog()

    except Exception as e:
        logger.error("Error in Workset Manager: {}".format(traceback.format_exc()))
        forms.alert("ข้อผิดพลาดในการเปิด Workset Manager:\n{}".format(str(e)), title="ข้อผิดพลาด")

if __name__ == "__main__":
    main()