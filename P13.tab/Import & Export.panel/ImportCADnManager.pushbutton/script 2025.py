# -*- coding: utf-8 -*-
__title__ = "ImportCAD\nManager"
__author__ = "เพิ่มพงษ์ ทวีกุล"
__doc__ = "Manage CAD Imports - Explode, Delete, Open View"

import clr
clr.AddReference("System")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

import System
from pyrevit import revit, DB
from System.Windows import Window, Thickness, RoutedEventHandler
from System.Windows.Controls import (
    Grid, RowDefinition, ColumnDefinition, DataGrid, DataGridTextColumn, 
    Button, StackPanel, Label, CheckBox, ProgressBar, ToolTip, TextBox
)
from System.Windows.Data import Binding
from System.Windows.Controls import DataGridLength, DataGridLengthUnitType
from System.Windows.Media import Brushes, FontFamily
from System.Windows.Input import MouseButtonEventHandler, KeyEventHandler, Key
from System.Collections.ObjectModel import ObservableCollection
import time

doc = revit.doc
uidoc = revit.uidoc

# ---------------------------------------------------------------------------
# CAD ITEM CLASS
# ---------------------------------------------------------------------------
class CADItem(object):
    def __init__(self, element_id, cad_name, view_name, view_type, status, element_obj, cad_path=""):
        self.element_id = element_id
        self.cad_name = cad_name
        self.view_name = view_name
        self.view_type = view_type
        self.status = status
        self.selected = False
        self.element_obj = element_obj
        self.cad_path = cad_path

# ---------------------------------------------------------------------------
# MAIN WINDOW
# ---------------------------------------------------------------------------
class CADImportWindow(Window):
    def __init__(self):
        Window.__init__(self)
        self.cad_data = ObservableCollection[CADItem]()
        self.all_cad_data = []  # เก็บข้อมูลทั้งหมดสำหรับการค้นหา
        self.Title = "CAD Import Manager - pyRevit"
        self.Width = 1100
        self.Height = 700
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.last_search_time = 0
        self.search_delay = 0.3  # หน่วงเวลา 0.3 วินาที

        self.InitializeComponent()
        self.find_all_cad_imports()
        self.load_data()

    # =========================================================================
    # UI INITIALIZATION
    # =========================================================================
    def InitializeComponent(self):
        main_grid = Grid()
        self.Content = main_grid

        # ROW DEFINITIONS
        row_heights = [
            (40, 0),   # Title
            (40, 0),   # Search Panel
            (1, 1),    # DataGrid
            (35, 0),   # Summary
            (30, 0),   # ProgressBar
            (60, 0)    # Buttons
        ]
        for h, t in row_heights:
            rd = RowDefinition()
            rd.Height = System.Windows.GridLength(h, System.Windows.GridUnitType.Pixel if t==0 else System.Windows.GridUnitType.Star)
            main_grid.RowDefinitions.Add(rd)

        # COLUMN DEFINITIONS สำหรับ Search Panel
        for i in range(3):
            cd = ColumnDefinition()
            if i == 0:  # Column สำหรับป้ายค้นหา
                cd.Width = System.Windows.GridLength(80, System.Windows.GridUnitType.Pixel)
            elif i == 1:  # Column สำหรับช่องค้นหา
                cd.Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
            else:  # Column สำหรับปุ่มค้นหา
                cd.Width = System.Windows.GridLength(100, System.Windows.GridUnitType.Pixel)
            main_grid.ColumnDefinitions.Add(cd)

        # TITLE LABEL
        title_label = Label()
        title_label.Content = "📁 CAD Import Manager - pyRevit"
        title_label.FontSize = 16
        title_label.FontWeight = System.Windows.FontWeights.Bold
        title_label.Foreground = Brushes.DarkBlue
        title_label.HorizontalContentAlignment = System.Windows.HorizontalAlignment.Center
        title_label.VerticalContentAlignment = System.Windows.VerticalAlignment.Center
        title_label.Background = Brushes.LightSkyBlue
        Grid.SetRow(title_label, 0)
        Grid.SetColumnSpan(title_label, 3)
        main_grid.Children.Add(title_label)

        # SEARCH PANEL
        # Search Label
        search_label = Label()
        search_label.Content = "ค้นหา:"
        search_label.FontWeight = System.Windows.FontWeights.Bold
        search_label.HorizontalContentAlignment = System.Windows.HorizontalAlignment.Right
        search_label.VerticalContentAlignment = System.Windows.VerticalAlignment.Center
        Grid.SetRow(search_label, 1)
        Grid.SetColumn(search_label, 0)
        main_grid.Children.Add(search_label)

        # Search TextBox
        self.search_textbox = TextBox()
        self.search_textbox.Margin = Thickness(5, 5, 5, 5)
        self.search_textbox.FontSize = 14
        self.search_textbox.VerticalContentAlignment = System.Windows.VerticalAlignment.Center
        self.search_textbox.ToolTip = "ค้นหาตามชื่อ CAD, View, View Type, Status, หรือ Path\nกด Enter เพื่อค้นหา, Esc เพื่อล้าง"
        self.search_textbox.TextChanged += self.on_search_text_changed
        self.search_textbox.KeyDown += self.on_search_key_down
        Grid.SetRow(self.search_textbox, 1)
        Grid.SetColumn(self.search_textbox, 1)
        main_grid.Children.Add(self.search_textbox)

        # Search Buttons Panel
        search_button_panel = StackPanel()
        search_button_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        search_button_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Left
        search_button_panel.Margin = Thickness(5, 5, 5, 5)

        # Search Button
        search_button = Button()
        search_button.Content = "ค้นหา"
        search_button.Width = 80
        search_button.Height = 25
        search_button.Background = Brushes.LightBlue
        search_button.Click += self.search_click
        search_button_panel.Children.Add(search_button)

        # Clear Search Button
        clear_search_button = Button()
        clear_search_button.Content = "ล้าง"
        clear_search_button.Width = 80
        clear_search_button.Height = 25
        clear_search_button.Margin = Thickness(5, 0, 0, 0)
        clear_search_button.Background = Brushes.LightYellow
        clear_search_button.Click += self.clear_search_click
        search_button_panel.Children.Add(clear_search_button)

        Grid.SetRow(search_button_panel, 1)
        Grid.SetColumn(search_button_panel, 2)
        main_grid.Children.Add(search_button_panel)

        # DATA GRID
        self.data_grid = DataGrid()
        self.data_grid.AutoGenerateColumns = False
        self.data_grid.IsReadOnly = True
        self.data_grid.CanUserAddRows = False
        # ตั้งค่าให้เลือกหลายแถวได้ด้วย Shift และ Ctrl
        self.data_grid.SelectionMode = System.Windows.Controls.DataGridSelectionMode.Extended
        self.data_grid.SelectionUnit = System.Windows.Controls.DataGridSelectionUnit.FullRow
        self.data_grid.Background = Brushes.White
        self.data_grid.RowBackground = Brushes.White
        self.data_grid.AlternatingRowBackground = Brushes.WhiteSmoke
        self.data_grid.Foreground = Brushes.Black
        self.data_grid.FontSize = 12
        self.data_grid.FontFamily = FontFamily("Segoe UI")
        self.data_grid.MouseDoubleClick += MouseButtonEventHandler(self.open_cad_view)

        # Columns
        columns = [
            ("Element ID", "element_id", 80),
            ("CAD Name", "cad_name", 200),
            ("View", "view_name", 150),
            ("View Type", "view_type", 120),
            ("Status", "status", 100),
            ("Path", "cad_path", 250),
        ]
        for hdr, bind, w in columns:
            col = DataGridTextColumn()
            col.Header = hdr
            col.Binding = Binding(bind)
            col.Width = DataGridLength(w, DataGridLengthUnitType.Pixel)
            self.data_grid.Columns.Add(col)

        Grid.SetRow(self.data_grid, 2)
        Grid.SetColumnSpan(self.data_grid, 3)
        main_grid.Children.Add(self.data_grid)

        # SUMMARY LABEL
        self.summary_label = Label()
        self.summary_label.Content = "Found 0 CAD files"
        self.summary_label.FontWeight = System.Windows.FontWeights.Bold
        self.summary_label.HorizontalContentAlignment = System.Windows.HorizontalAlignment.Center
        self.summary_label.VerticalContentAlignment = System.Windows.VerticalAlignment.Center
        self.summary_label.Background = Brushes.LightGray
        Grid.SetRow(self.summary_label, 3)
        Grid.SetColumnSpan(self.summary_label, 3)
        main_grid.Children.Add(self.summary_label)

        # PROGRESS BAR
        self.progress_bar = ProgressBar()
        self.progress_bar.Minimum = 0
        self.progress_bar.Maximum = 100
        self.progress_bar.Value = 0
        self.progress_bar.Height = 20
        self.progress_bar.Margin = Thickness(5)
        Grid.SetRow(self.progress_bar, 4)
        Grid.SetColumnSpan(self.progress_bar, 3)
        main_grid.Children.Add(self.progress_bar)

        # BUTTON PANEL
        button_panel = StackPanel()
        button_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        button_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
        button_panel.Margin = Thickness(5)

        buttons = [
            ("Select All", Brushes.LightGreen, self.select_all_click),
            ("Clear Selection", Brushes.LightYellow, self.clear_selection_click),
            ("Explode Selected", Brushes.Orange, self.explode_selected_click),
            ("Delete Selected", Brushes.LightCoral, self.delete_selected_click),
            ("Explode All", Brushes.Gold, self.explode_all_click),
            ("Delete All", Brushes.Red, self.delete_all_click),
            ("Refresh", Brushes.LightBlue, self.refresh_click),
        ]
        for text, color, handler in buttons:
            btn = Button()
            btn.Content = text
            btn.Width = 120
            btn.Height = 35
            btn.Margin = Thickness(5)
            btn.Background = color
            btn.Click += handler
            button_panel.Children.Add(btn)

        Grid.SetRow(button_panel, 5)
        Grid.SetColumnSpan(button_panel, 3)
        main_grid.Children.Add(button_panel)

    # =========================================================================
    # SEARCH FUNCTIONS - FIXED REAL-TIME SEARCH
    # =========================================================================
    def on_search_text_changed(self, sender, e):
        # Real-time search ด้วยการหน่วงเวลา
        current_time = time.time()
        if current_time - self.last_search_time > self.search_delay:
            self.last_search_time = current_time
            self.perform_search()

    def on_search_key_down(self, sender, e):
        if e.Key == Key.Enter:
            self.perform_search()
        elif e.Key == Key.Escape:
            self.clear_search_click(None, None)
            self.search_textbox.Text = ""

    def search_click(self, sender, e):
        self.perform_search()

    def clear_search_click(self, sender, e):
        self.search_textbox.Text = ""
        self.load_full_data()
        self.update_summary()
        self.update_status("Cleared search")

    def perform_search(self):
        search_text = self.search_textbox.Text.Trim()
        if not search_text:
            self.clear_search_click(None, None)
            return

        # สร้าง filtered collection ใหม่
        filtered_items = ObservableCollection[CADItem]()
        search_lower = search_text.lower()
        
        for item in self.all_cad_data:
            # ค้นหาในทุกฟิลด์ที่เกี่ยวข้อง (แปลงเป็น lowercase ก่อนค้นหา)
            if (search_lower in (item.cad_name or "").lower() or 
                search_lower in (item.view_name or "").lower() or 
                search_lower in (item.view_type or "").lower() or 
                search_lower in (item.status or "").lower() or 
                search_lower in (item.cad_path or "").lower() or
                search_lower in (item.element_id or "").lower()):
                filtered_items.Add(item)

        # อัพเดท DataGrid ด้วยข้อมูลที่กรองแล้ว
        self.cad_data.Clear()
        for item in filtered_items:
            self.cad_data.Add(item)
            
        self.data_grid.ItemsSource = self.cad_data
        self.update_summary()
        self.update_status("Search completed: {} items found".format(self.cad_data.Count))

    def load_full_data(self):
        """โหลดข้อมูลทั้งหมดกลับเข้าไปใน cad_data"""
        self.cad_data.Clear()
        for item in self.all_cad_data:
            self.cad_data.Add(item)
        self.data_grid.ItemsSource = self.cad_data

    # =========================================================================
    # GET CAD NAME & PATH
    # =========================================================================
    def get_cad_name_path(self, cad):
        cad_name = "N/A"
        cad_path = ""
        try:
            elem_type = doc.GetElement(cad.GetTypeId())
            if elem_type and hasattr(elem_type, "Name") and elem_type.Name:
                cad_name = elem_type.Name
            else:
                param = cad.get_Parameter(DB.BuiltInParameter.IMPORT_SYMBOL_NAME)
                if param and param.HasValue:
                    cad_name = param.AsString()
            # CAD Path
            param_path = cad.get_Parameter(DB.BuiltInParameter.IMPORT_SYMBOL_FILENAME)
            if param_path and param_path.HasValue:
                cad_path = param_path.AsString()
            if cad_name == "N/A" and cad.Category:
                cad_name = cad.Category.Name
        except:
            pass
        return cad_name, cad_path

    # =========================================================================
    # DATA COLLECTION
    # =========================================================================
    def find_all_cad_imports(self):
        self.all_cad_data = []
        self.cad_data.Clear()
        
        cad_links = list(DB.FilteredElementCollector(doc)
                         .OfClass(DB.ImportInstance)
                         .WhereElementIsNotElementType())
        for cad in cad_links:
            try:
                cad_name, cad_path = self.get_cad_name_path(cad)
                view_name, view_type = "Project-wide", "Project-wide"
                owner_id = cad.OwnerViewId
                if owner_id != DB.ElementId.InvalidElementId:
                    owner_view = doc.GetElement(owner_id)
                    if owner_view:
                        view_name = owner_view.Name
                        view_type = str(owner_view.ViewType)
                item = CADItem(
                    element_id=str(cad.Id.IntegerValue),
                    cad_name=cad_name,
                    view_name=view_name,
                    view_type=view_type,
                    status="Ready",
                    element_obj=cad,
                    cad_path=cad_path
                )
                self.all_cad_data.append(item)
                self.cad_data.Add(item)
            except Exception as e:
                print("Error processing CAD: {}".format(str(e)))
                pass

    # =========================================================================
    # LOAD DATA
    # =========================================================================
    def load_data(self):
        self.data_grid.ItemsSource = self.cad_data
        self.update_summary()

    def update_summary(self):
        total_all = len(self.all_cad_data)
        current_count = self.cad_data.Count
        remaining = len([i for i in self.all_cad_data if i.status == "Ready"])
        selected_count = len(self.data_grid.SelectedItems) if self.data_grid.SelectedItems else 0
        
        if current_count == total_all:
            self.summary_label.Content = "Found {0} CAD files   |   {1} remaining   |   {2} selected".format(total_all, remaining, selected_count)
        else:
            self.summary_label.Content = "Found {0} CAD files   |   {1} filtered   |   {2} remaining   |   {3} selected".format(total_all, current_count, remaining, selected_count)

    # =========================================================================
    # BUTTON HANDLERS
    # =========================================================================
    def select_all_click(self, sender, e):
        self.data_grid.SelectAll()
        self.update_summary()
        self.update_status("Selected all CAD files")

    def clear_selection_click(self, sender, e):
        self.data_grid.UnselectAll()
        self.update_summary()
        self.update_status("Cleared selection")

    def explode_selected_click(self, sender, e):
        self._process_selected_items(self.explode_item, "Exploded Selected")

    def delete_selected_click(self, sender, e):
        self._process_selected_items(self.delete_item, "Deleted Selected")

    def explode_all_click(self, sender, e):
        self._process_all_items(self.explode_item, "Exploded All")

    def delete_all_click(self, sender, e):
        self._process_all_items(self.delete_item, "Deleted All")

    def refresh_click(self, sender, e):
        self.find_all_cad_imports()
        self.load_data()
        self.update_status("List refreshed")

    # =========================================================================
    # CORE PROCESSING
    # =========================================================================
    def _process_selected_items(self, func, action_name):
        selected = list(self.data_grid.SelectedItems)
        total = len(selected)
        if total == 0:
            self.update_status("No selection")
            return
        self.progress_bar.Minimum = 0
        self.progress_bar.Maximum = total
        self.progress_bar.Value = 0
        for idx, item in enumerate(selected):
            func(item)
            self.progress_bar.Value = idx + 1
        self.load_data()
        self.update_status("{0}: {1}/{2}".format(action_name, total, total))

    def _process_all_items(self, func, action_name):
        # ทำงานเฉพาะข้อมูลที่แสดงอยู่ (filtered)
        items_to_process = list(self.cad_data)
        total = len(items_to_process)
        self.progress_bar.Minimum = 0
        self.progress_bar.Maximum = total
        self.progress_bar.Value = 0
        for idx, item in enumerate(items_to_process):
            func(item)
            self.progress_bar.Value = idx + 1
        self.load_data()
        self.update_status("{0}: {1}/{2}".format(action_name, total, total))

    # =========================================================================
    # OPERATIONS
    # =========================================================================
    def explode_item(self, item):
        try:
            cad = doc.GetElement(DB.ElementId(int(item.element_id)))
            if cad:
                t = DB.Transaction(doc, "Explode CAD")
                t.Start()
                try:
                    cad.Explode()
                    t.Commit()
                    item.status = "Exploded"
                except Exception as e:
                    t.RollBack()
                    item.status = "Explode Failed: {}".format(str(e))
        except Exception as e:
            item.status = "Error: {}".format(str(e))

    def delete_item(self, item):
        try:
            eid = DB.ElementId(int(item.element_id))
            t = DB.Transaction(doc, "Delete CAD")
            t.Start()
            try:
                doc.Delete(eid)
                t.Commit()
                item.status = "Deleted"
            except Exception as e:
                t.RollBack()
                item.status = "Delete Failed: {}".format(str(e))
        except Exception as e:
            item.status = "Error: {}".format(str(e))

    # =========================================================================
    # OPEN VIEW ON DOUBLE CLICK
    # =========================================================================
    def open_cad_view(self, sender, e):
        if self.data_grid.SelectedItems.Count > 0:
            row = self.data_grid.SelectedItems[0]
            if row and hasattr(row, "element_obj"):
                cad = row.element_obj
                owner_id = cad.OwnerViewId
                if owner_id != DB.ElementId.InvalidElementId:
                    view = doc.GetElement(owner_id)
                    if view:
                        try:
                            uidoc.ActiveView = view
                            self.update_status("Opened view: {}".format(view.Name))
                        except:
                            self.update_status("Cannot open view")
                else:
                    self.update_status("Project-wide CAD cannot open a view")

    def update_status(self, msg):
        # แยกส่วน summary ออกและแสดงเฉพาะ status ล่าสุด
        base_content = self.summary_label.Content
        if "   |   " in str(base_content):
            parts = str(base_content).split("   |   ")
            # เก็บเฉพาะส่วนที่เป็นสถิติ (ส่วนแรก)
            new_content = parts[0] + "   |   {0}".format(msg)
            self.summary_label.Content = new_content
        else:
            self.summary_label.Content = str(base_content) + "   |   {0}".format(msg)

# ---------------------------------------------------------------------------
# RUN WINDOW
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    cad_links = list(DB.FilteredElementCollector(doc)
                     .OfClass(DB.ImportInstance)
                     .WhereElementIsNotElementType())
    if cad_links:
        window = CADImportWindow()
        window.ShowDialog()
    else:
        from System.Windows import MessageBox
        MessageBox.Show("No CAD imports found in the project.", "CAD Import Manager")