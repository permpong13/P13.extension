# -*- coding: utf-8 -*-
from __future__ import unicode_literals, print_function

import os
import re
import json
import traceback
import clr

clr.AddReference('System.Windows.Forms')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from pyrevit import revit, DB, forms
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List
# ลบบรรทัดที่มี ListSortDirection ออกเพื่อแก้ปัญหา Import Error
from System.Windows import (
    Window, WindowStartupLocation, GridLength, GridUnitType, 
    Thickness, VerticalAlignment, FontWeight, FontWeights, Visibility
)
from System.Windows.Data import Binding
from System.Windows.Controls import (
    Grid, RowDefinition, ColumnDefinition, GroupBox, StackPanel, 
    TextBlock, TextBox, Button, DataGrid, DataGridCheckBoxColumn, 
    DataGridTextColumn, Orientation, DataGridLength, DataGridLengthUnitType, 
    ComboBox, ListBox, DataGridSelectionMode, CheckBox
)
from System.Windows.Media import Brushes

doc = revit.doc
THIS_DIR = os.path.dirname(__file__)
CONFIG_FILE = os.path.join(THIS_DIR, 'p13_supersheet_config.json')
LAST_SETTING_FILE = os.path.join(THIS_DIR, 'p13_last_settings.json')

class ExportItem(object):
    def __init__(self, obj):
        self.Include = False
        self.item_obj = obj
        self.Id = obj.Id
        self.Number = obj.SheetNumber or ""
        self.Name = obj.Name or ""
        self.Revision = ""
        try:
            rev_ids = obj.GetAllRevisionIds()
            if rev_ids:
                last_rev_id = list(rev_ids)[-1]
                self.Revision = obj.GetRevisionNumber(last_rev_id) or ""
        except: pass

class SuperSheetsUltimate(Window):
    def __init__(self):
        self.Title = "P13 SuperSheets Pro Max 2026 | Fixed Edition"
        self.Width = 1200; self.Height = 950
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = Brushes.WhiteSmoke
        self._all_items = []
        self._profiles = {}
        
        self._setup_ui()
        self._load_profiles_from_disk()
        self._load_project_params()
        self._refresh_data()
        self._load_last_settings()
        
        self.Closing += lambda s, e: self._save_current_as_last()

    def _setup_ui(self):
        main_layout = Grid(); main_layout.Margin = Thickness(15)
        for h in ["Auto", "Auto", "Auto", "Auto", "*", "Auto"]:
            rd = RowDefinition()
            rd.Height = GridLength.Auto if h == "Auto" else GridLength(1, GridUnitType.Star)
            main_layout.RowDefinitions.Add(rd)
        self.Content = main_layout

        # 1. PROFILE & FORMAT
        gb_prof = GroupBox(); gb_prof.Header = "Profile & Format Settings"; gb_prof.Margin = Thickness(0,0,0,10)
        sp_prof = StackPanel(); sp_prof.Orientation = Orientation.Horizontal; sp_prof.Margin = Thickness(10)
        sp_prof.Children.Add(TextBlock(Text="Profile: ", VerticalAlignment=VerticalAlignment.Center))
        self.cboProfile = ComboBox(); self.cboProfile.Width = 120; self.cboProfile.SelectionChanged += self._on_profile_select
        sp_prof.Children.Add(self.cboProfile)
        self.txtProfileName = TextBox(); self.txtProfileName.Width = 100; self.txtProfileName.Margin = Thickness(5,0,5,0)
        sp_prof.Children.Add(self.txtProfileName)
        btn_save = Button(); btn_save.Content = " Save Profile "; btn_save.Click += self._on_save_profile; sp_prof.Children.Add(btn_save)
        
        sp_prof.Children.Add(TextBlock(Text=" | Format: ", VerticalAlignment=VerticalAlignment.Center, Margin=Thickness(15,0,0,0)))
        self.cboType = ComboBox(); self.cboType.Width = 60; self.cboType.Items.Add("PDF"); self.cboType.Items.Add("DWG"); self.cboType.SelectedIndex = 0
        sp_prof.Children.Add(self.cboType)
        
        sp_prof.Children.Add(TextBlock(Text=" | Mode: ", VerticalAlignment=VerticalAlignment.Center, Margin=Thickness(10,0,0,0)))
        self.cboColor = ComboBox(); self.cboColor.Width = 100; self.cboColor.Items.Add("Color"); self.cboColor.Items.Add("Black & White"); self.cboColor.SelectedIndex = 0
        sp_prof.Children.Add(self.cboColor)

        self.chkCombine = CheckBox(); self.chkCombine.Content = "Combine PDF"; self.chkCombine.VerticalAlignment = VerticalAlignment.Center
        self.chkCombine.Margin = Thickness(15,0,0,0); self.chkCombine.Checked += self._toggle_combine_ui; self.chkCombine.Unchecked += self._toggle_combine_ui
        sp_prof.Children.Add(self.chkCombine)
        gb_prof.Content = sp_prof; Grid.SetRow(gb_prof, 0); main_layout.Children.Add(gb_prof)

        # 2. NAMING & PATH
        gb_name = GroupBox(); gb_name.Header = "Naming & Path Settings"; gb_name.Margin = Thickness(0,0,0,10)
        gs = Grid(); gs.Margin = Thickness(10)
        gs.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1.3, GridUnitType.Star)))
        gs.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        
        sp_naming = StackPanel(); sp_naming.Margin = Thickness(0,0,10,0)
        sp_path = StackPanel(); sp_path.Orientation = Orientation.Horizontal; sp_path.Margin = Thickness(0,0,0,10)
        sp_path.Children.Add(TextBlock(Text="Export Path: ", VerticalAlignment=VerticalAlignment.Center))
        self.txtPath = TextBox(); self.txtPath.Width = 350; self.txtPath.Margin = Thickness(5,0,5,0)
        btn_br = Button(); btn_br.Content = " Browse "; btn_br.Click += self._on_browse
        sp_path.Children.Add(self.txtPath); sp_path.Children.Add(btn_br); sp_naming.Children.Add(sp_path)
        
        # Single Naming Setup
        self.spSingleName = StackPanel()
        sp_pre = StackPanel(); sp_pre.Orientation = Orientation.Horizontal
        self.txtPrefix = TextBox(); self.txtPrefix.Width = 80; self.txtSuffix = TextBox(); self.txtSuffix.Width = 80
        sp_pre.Children.Add(TextBlock(Text="Prefix: ")); sp_pre.Children.Add(self.txtPrefix); sp_pre.Children.Add(TextBlock(Text=" Suffix: ")); sp_pre.Children.Add(self.txtSuffix)
        self.spSingleName.Children.Add(sp_pre)
        self.txtPattern = TextBox(); self.txtPattern.Height = 25; self.txtPattern.Margin = Thickness(0,5,0,5); self.txtPattern.Text = "{SheetNumber}_{SheetName}"
        self.spSingleName.Children.Add(self.txtPattern)
        sp_tok = StackPanel(); sp_tok.Orientation = Orientation.Horizontal
        for t in ["{SheetNumber}", "{SheetName}", "{Revision}"]:
            b = Button(); b.Content = t; b.Margin = Thickness(0,0,3,0); b.Click += lambda s,e,v=t: self.txtPattern.AppendText(v); sp_tok.Children.Add(b)
        self.spSingleName.Children.Add(sp_tok); sp_naming.Children.Add(self.spSingleName)

        # Combined Naming Setup
        self.spCombineName = StackPanel(); self.spCombineName.Visibility = Visibility.Collapsed
        self.spCombineName.Children.Add(TextBlock(Text="Combined Filename:", FontWeight=FontWeights.Bold))
        self.txtCombineName = TextBox(); self.txtCombineName.Height = 25; self.txtCombineName.Margin = Thickness(0,5,0,0); self.txtCombineName.Text = "Combined_Project_Set"
        self.spCombineName.Children.Add(self.txtCombineName); sp_naming.Children.Add(self.spCombineName)
        
        Grid.SetColumn(sp_naming, 0); gs.Children.Add(sp_naming)
        
        # Project Parameters Search
        sp_p = StackPanel(); self.txtPSearch = TextBox(); self.txtPSearch.TextChanged += self._filter_params; self.lstP = ListBox(); self.lstP.Height = 75
        btn_ap = Button(); btn_ap.Content = " + Insert Param "; btn_ap.Click += self._add_param_token
        sp_p.Children.Add(self.txtPSearch); sp_p.Children.Add(self.lstP); sp_p.Children.Add(btn_ap); Grid.SetColumn(sp_p, 1); gs.Children.Add(sp_p)
        gb_name.Content = gs; Grid.SetRow(gb_name, 1); main_layout.Children.Add(gb_name)

        # 3. FILTER & SELECTION
        gb_filter = GroupBox(); gb_filter.Header = "Filter & Selection"; gb_filter.Margin = Thickness(0,0,0,10)
        sp_f = StackPanel(); sp_f.Orientation = Orientation.Horizontal; sp_f.Margin = Thickness(10)
        sp_f.Children.Add(TextBlock(Text="Search Sheets: ", VerticalAlignment=VerticalAlignment.Center, FontWeight=FontWeights.Bold))
        self.txtSearchData = TextBox(); self.txtSearchData.Width = 250; self.txtSearchData.TextChanged += self._filter_data_grid
        sp_f.Children.Add(self.txtSearchData)
        b_all = Button(); b_all.Content = " Check All "; b_all.Margin = Thickness(20,0,5,0); b_all.Click += lambda s,e: self._set_all(True)
        b_none = Button(); b_none.Content = " Uncheck All "; b_none.Click += lambda s,e: self._set_all(False)
        b_h = Button(); b_h.Content = " Check Hilighted "; b_h.Background = Brushes.LightBlue; b_h.Margin = Thickness(5,0,0,0); b_h.Click += self._on_check_highlighted
        sp_f.Children.Add(b_all); sp_f.Children.Add(b_none); sp_f.Children.Add(b_h)
        gb_filter.Content = sp_f; Grid.SetRow(gb_filter, 2); main_layout.Children.Add(gb_filter)

        # 4. DATA GRID
        self.dg = DataGrid(); self.dg.AutoGenerateColumns = False; self.dg.CanUserAddRows = False; self.dg.SelectionMode = DataGridSelectionMode.Extended
        col_chk = DataGridCheckBoxColumn(); col_chk.Header = "Export"; col_chk.Binding = Binding("Include"); self.dg.Columns.Add(col_chk)
        self.dg.Columns.Add(DataGridTextColumn(Header="Number", Binding=Binding("Number"), IsReadOnly=True))
        self.dg.Columns.Add(DataGridTextColumn(Header="Name", Binding=Binding("Name"), Width=DataGridLength(1, DataGridLengthUnitType.Star), IsReadOnly=True))
        self.dg.Columns.Add(DataGridTextColumn(Header="Rev", Binding=Binding("Revision"), IsReadOnly=True))
        Grid.SetRow(self.dg, 4); main_layout.Children.Add(self.dg)

        # 5. START BUTTON
        btn_run = Button(); btn_run.Content = " START EXPORT "; btn_run.Height = 45; btn_run.Margin = Thickness(0,10,0,0); btn_run.Background = Brushes.ForestGreen; btn_run.Foreground = Brushes.White; btn_run.FontWeight = FontWeights.Bold; btn_run.Click += self._on_export
        Grid.SetRow(btn_run, 5); main_layout.Children.Add(btn_run)

    def _toggle_combine_ui(self, s, e):
        if self.chkCombine.IsChecked:
            self.spCombineName.Visibility = Visibility.Visible
            self.spSingleName.Visibility = Visibility.Collapsed
        else:
            self.spCombineName.Visibility = Visibility.Collapsed
            self.spSingleName.Visibility = Visibility.Visible

    def _filter_data_grid(self, s, e):
        txt = self.txtSearchData.Text.lower()
        filtered = [i for i in self._all_items if txt in i.Number.lower() or txt in i.Name.lower()]
        self.dg.ItemsSource = ObservableCollection[ExportItem](filtered)

    def _save_current_as_last(self):
        data = {
            'prefix': self.txtPrefix.Text, 'pattern': self.txtPattern.Text, 'suffix': self.txtSuffix.Text,
            'path': self.txtPath.Text, 'type': self.cboType.SelectedIndex, 'color': self.cboColor.SelectedIndex,
            'combine': self.chkCombine.IsChecked, 'combine_name': self.txtCombineName.Text
        }
        with open(LAST_SETTING_FILE, 'w') as f: json.dump(data, f)

    def _load_last_settings(self):
        if os.path.exists(LAST_SETTING_FILE):
            try:
                with open(LAST_SETTING_FILE, 'r') as f:
                    p = json.load(f)
                    self.txtPrefix.Text = p.get('prefix',''); self.txtPattern.Text = p.get('pattern','{SheetNumber}_{SheetName}')
                    self.txtSuffix.Text = p.get('suffix',''); self.txtPath.Text = p.get('path','')
                    self.cboType.SelectedIndex = p.get('type', 0); self.cboColor.SelectedIndex = p.get('color', 0)
                    self.chkCombine.IsChecked = p.get('combine', False)
                    self.txtCombineName.Text = p.get('combine_name', 'Combined_Project_Set')
                    self._toggle_combine_ui(None, None)
            except: pass

    def _on_export(self, s, e):
        current_items = list(self.dg.ItemsSource)
        selected = [i for i in current_items if i.Include]
        if not selected: forms.alert("No sheets selected!"); return
        
        if not self.txtPath.Text or not os.path.exists(self.txtPath.Text):
            forms.alert("Please select Export Folder Path first.")
            self._on_browse(None, None)
            if not self.txtPath.Text: return

        self._save_current_as_last()
        folder = self.txtPath.Text
        is_dwg = self.cboType.SelectedIndex == 1
        is_bw = self.cboColor.SelectedIndex == 1
        combine = self.chkCombine.IsChecked if not is_dwg else False

        try:
            if combine:
                opt_pdf = DB.PDFExportOptions()
                opt_pdf.ColorDepth = DB.ColorDepthType.BlackLine if is_bw else DB.ColorDepthType.Color
                opt_pdf.Combine = True
                opt_pdf.FileName = re.sub(r'[\\/*?:"<>|]', "_", self.txtCombineName.Text).strip()
                doc.Export(folder, List[DB.ElementId]([i.Id for i in selected]), opt_pdf)
            else:
                with forms.ProgressBar(title="Exporting...", total=len(selected)) as pb:
                    for idx, item in enumerate(selected):
                        fn = self.txtPrefix.Text + self.txtPattern.Text + self.txtSuffix.Text
                        fn = fn.replace("{SheetNumber}", item.Number).replace("{SheetName}", item.Name).replace("{Revision}", item.Revision)
                        fn = re.sub(r'[\\/*?:"<>|]', "_", fn).strip()
                        if is_dwg:
                            opt = DB.DWGExportOptions(); opt.MergedViews = True
                            doc.Export(folder, fn, List[DB.ElementId]([item.Id]), opt)
                        else:
                            opt_pdf = DB.PDFExportOptions(); opt_pdf.FileName = fn
                            opt_pdf.ColorDepth = DB.ColorDepthType.BlackLine if is_bw else DB.ColorDepthType.Color
                            doc.Export(folder, List[DB.ElementId]([item.Id]), opt_pdf)
                        pb.update_progress(idx+1, len(selected))
            forms.alert("Export Successful!"); self.Close()
        except Exception as ex: print(traceback.format_exc())

    def _on_save_profile(self, s, e):
        n = self.txtProfileName.Text
        if n:
            self._profiles[n] = {
                'prefix': self.txtPrefix.Text, 'pattern': self.txtPattern.Text, 'suffix': self.txtSuffix.Text,
                'path': self.txtPath.Text, 'type': self.cboType.SelectedIndex, 'color': self.cboColor.SelectedIndex,
                'combine': self.chkCombine.IsChecked, 'combine_name': self.txtCombineName.Text
            }
            with open(CONFIG_FILE, 'w') as f: json.dump(self._profiles, f)
            self._load_profiles_from_disk(); self.cboProfile.SelectedItem = n

    def _on_profile_select(self, s, e):
        name = self.cboProfile.SelectedItem
        if name in self._profiles:
            p = self._profiles[name]
            self.txtPrefix.Text = p.get('prefix',''); self.txtPattern.Text = p.get('pattern','')
            self.txtSuffix.Text = p.get('suffix',''); self.txtPath.Text = p.get('path','')
            self.cboType.SelectedIndex = p.get('type', 0); self.cboColor.SelectedIndex = p.get('color', 0)
            self.chkCombine.IsChecked = p.get('combine', False); self.txtCombineName.Text = p.get('combine_name', '')
            self._toggle_combine_ui(None, None); self.txtProfileName.Text = name

    def _on_browse(self, s, e):
        p = forms.pick_folder()
        if p: self.txtPath.Text = p

    def _on_check_highlighted(self, s, e):
        for item in self.dg.SelectedItems: item.Include = True
        self.dg.Items.Refresh()

    def _set_all(self, val):
        for i in self.dg.ItemsSource: i.Include = val
        self.dg.Items.Refresh()

    def _load_profiles_from_disk(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    self._profiles = json.load(f); self.cboProfile.Items.Clear()
                    for k in sorted(self._profiles.keys()): self.cboProfile.Items.Add(k)
            except: pass

    def _load_project_params(self):
        self._all_params = sorted([p.Definition.Name for p in doc.ProjectInformation.Parameters if p.Definition])
        for p in self._all_params: self.lstP.Items.Add(p)

    def _filter_params(self, s, e):
        t = self.txtPSearch.Text.lower(); self.lstP.Items.Clear()
        for p in self._all_params: 
            if t in p.lower(): self.lstP.Items.Add(p)

    def _add_param_token(self, s, e):
        if self.lstP.SelectedItem:
            target = self.txtCombineName if self.chkCombine.IsChecked else self.txtPattern
            target.AppendText("<" + self.lstP.SelectedItem + ">")

    def _refresh_data(self):
        self._all_items = [ExportItem(s) for s in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet) if not s.IsPlaceholder]
        # ระบบ Sort พื้นฐานโดยใช้ Python Sort (เสถียรกว่า)
        self._all_items.sort(key=lambda x: x.Number)
        self.dg.ItemsSource = ObservableCollection[ExportItem](self._all_items)

if __name__ == '__main__':
    SuperSheetsUltimate().ShowDialog()