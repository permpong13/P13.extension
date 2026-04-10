# -*- coding: utf-8 -*-
from __future__ import unicode_literals, print_function

import os
import re
import json
import traceback
import clr
import datetime

# Import Libraries
clr.AddReference('System.Windows.Forms')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from pyrevit import revit, DB, forms, output
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List
from System.Windows import (
    Window, WindowStartupLocation, GridLength, GridUnitType, 
    Thickness, VerticalAlignment, FontWeight, FontWeights, Visibility
)
from System.Windows.Data import Binding
from System.Windows.Controls import (
    Grid, RowDefinition, ColumnDefinition, GroupBox, StackPanel, 
    TextBlock, TextBox, Button, DataGrid, DataGridCheckBoxColumn, 
    DataGridTextColumn, Orientation, DataGridLength, DataGridLengthUnitType, 
    ComboBox, ListBox, DataGridSelectionMode, CheckBox, VirtualizingStackPanel
)
from System.Windows.Media import Brushes

doc = revit.doc
THIS_DIR = os.path.dirname(__file__)
CONFIG_FILE = os.path.join(THIS_DIR, 'p13_supersheet_config.json')
LAST_SETTING_FILE = os.path.join(THIS_DIR, 'p13_last_settings.json')

# --- DATA CLASS ---
class ExportItem(object):
    def __init__(self, obj):
        self.Include = False
        self.item_obj = obj
        self.Id = obj.Id
        self.Number = obj.SheetNumber or ""
        self.Name = obj.Name or ""
        self.Revision = self._get_rev()
        self.PreviewName = "" 

    def _get_rev(self):
        try:
            p = self.item_obj.get_Parameter(DB.BuiltInParameter.SHEET_CURRENT_REVISION)
            return p.AsString() or p.AsValueString() or ""
        except: return ""
    
    def get_param_val(self, p_name):
        mapping = {"SheetNumber": DB.BuiltInParameter.SHEET_NUMBER, "SheetName": DB.BuiltInParameter.SHEET_NAME, "Current Revision": DB.BuiltInParameter.SHEET_CURRENT_REVISION}
        p = self.item_obj.get_Parameter(mapping[p_name]) if p_name in mapping else self.item_obj.LookupParameter(p_name.strip("{} "))
        if not p: p = doc.ProjectInformation.LookupParameter(p_name.strip("{} "))
        return p.AsValueString() or p.AsString() or "" if p else ""

# --- MAIN APP ---
class SuperSheetsUltimate(Window):
    def __init__(self):
        self.Title = "P13 SuperSheets 2026 | Profile Fixed & Silent"
        self.Width = 1350; self.Height = 920
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True
        self.Background = Brushes.WhiteSmoke
        
        self._all_items = []
        self._profiles = {}
        self._all_params = []
        self._ignore_events = False # ตัวแปรสำคัญ: ป้องกัน Event ตีกัน

        self._setup_ui()
        self.ContentRendered += self._initial_load
        self.Closing += lambda s, e: self._save_current_as_last()

    def _initial_load(self, s, e):
        try:
            self._refresh_data()
            self._load_profiles_from_disk()
            self._load_project_params_list()
            self._load_all_params()
            self._load_print_sets()
            self._load_last_settings()
            self._update_all_previews()
            self._toggle_ui(None, None)
        except: print("Init Error: " + traceback.format_exc())

    def _setup_ui(self):
        main_layout = Grid(); main_layout.Margin = Thickness(15)
        for h in ["Auto", "Auto", "Auto", "Auto", "*", "Auto"]:
            rd = RowDefinition(); rd.Height = GridLength.Auto if h == "Auto" else GridLength(1, GridUnitType.Star)
            main_layout.RowDefinitions.Add(rd)
        self.Content = main_layout

        # 1. PROFILE
        gb_prof = GroupBox(); gb_prof.Header = "Profile & Format"; gb_prof.Margin = Thickness(0,0,0,10)
        sp_prof = StackPanel(); sp_prof.Orientation = Orientation.Horizontal; sp_prof.Margin = Thickness(10)
        
        sp_prof.Children.Add(TextBlock(Text="Profile: ", VerticalAlignment=VerticalAlignment.Center))
        self.cboProfile = ComboBox(); self.cboProfile.Width = 130; self.cboProfile.SelectionChanged += self._on_profile_select
        sp_prof.Children.Add(self.cboProfile)
        
        self.txtProfileName = TextBox(); self.txtProfileName.Width = 100; self.txtProfileName.Margin = Thickness(5,0,5,0)
        sp_prof.Children.Add(self.txtProfileName)
        
        btn_save = Button(); btn_save.Content = " Save/Edit "; btn_save.Click += self._on_save_profile; sp_prof.Children.Add(btn_save)
        btn_del = Button(); btn_del.Content = " Del "; btn_del.Margin = Thickness(5,0,0,0); btn_del.Background = Brushes.MistyRose; btn_del.Click += self._on_delete_profile; sp_prof.Children.Add(btn_del)
        
        btn_imp = Button(); btn_imp.Content = " Imp "; btn_imp.Margin = Thickness(15,0,0,0); btn_imp.Click += self._on_import_profile; sp_prof.Children.Add(btn_imp)
        btn_exp = Button(); btn_exp.Content = " Exp "; btn_exp.Margin = Thickness(5,0,0,0); btn_exp.Click += self._on_export_profile; sp_prof.Children.Add(btn_exp)

        sp_prof.Children.Add(TextBlock(Text=" | Export: ", VerticalAlignment=VerticalAlignment.Center, Margin=Thickness(15,0,0,0)))
        self.chkPDF = CheckBox(); self.chkPDF.Content = "PDF"; self.chkPDF.IsChecked = True; self.chkPDF.VerticalAlignment = VerticalAlignment.Center; sp_prof.Children.Add(self.chkPDF)
        self.chkDWG = CheckBox(); self.chkDWG.Content = "DWG"; self.chkDWG.Margin = Thickness(8,0,0,0); self.chkDWG.VerticalAlignment = VerticalAlignment.Center; sp_prof.Children.Add(self.chkDWG)
        self.chkIFC = CheckBox(); self.chkIFC.Content = "IFC"; self.chkIFC.Margin = Thickness(8,0,0,0); self.chkIFC.VerticalAlignment = VerticalAlignment.Center; sp_prof.Children.Add(self.chkIFC)
        self.chkNWC = CheckBox(); self.chkNWC.Content = "NWC"; self.chkNWC.Margin = Thickness(8,0,0,0); self.chkNWC.VerticalAlignment = VerticalAlignment.Center; sp_prof.Children.Add(self.chkNWC)
        
        self.cboColor = ComboBox(); self.cboColor.Width = 80; self.cboColor.Items.Add("Color"); self.cboColor.Items.Add("B&W"); self.cboColor.SelectedIndex = 0; self.cboColor.Margin = Thickness(15,0,0,0); sp_prof.Children.Add(self.cboColor)
        self.chkCombine = CheckBox(); self.chkCombine.Content = "Combine PDF"; self.chkCombine.Margin = Thickness(15,0,0,0); self.chkCombine.VerticalAlignment = VerticalAlignment.Center; self.chkCombine.Checked += self._toggle_ui; self.chkCombine.Unchecked += self._toggle_ui; sp_prof.Children.Add(self.chkCombine)
        
        gb_prof.Content = sp_prof; Grid.SetRow(gb_prof, 0); main_layout.Children.Add(gb_prof)

        # 2. NAMING
        gb_name = GroupBox(); gb_name.Header = "Path & Naming"; gb_name.Margin = Thickness(0,0,0,10)
        gs = Grid(); gs.Margin = Thickness(10); gs.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1.4, GridUnitType.Star))); gs.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        sp_left = StackPanel(); sp_left.Margin = Thickness(0,0,10,0)
        
        sp_path = StackPanel(); sp_path.Orientation = Orientation.Horizontal
        self.txtPath = TextBox(); self.txtPath.Width = 350; btn_br = Button(); btn_br.Content = " ... "; btn_br.Click += self._on_browse; btn_op = Button(); btn_op.Content = " Open "; btn_op.Click += self._on_open_folder
        sp_path.Children.Add(TextBlock(Text="Path: ", VerticalAlignment=VerticalAlignment.Center))
        sp_path.Children.Add(self.txtPath); sp_path.Children.Add(btn_br); sp_path.Children.Add(btn_op); sp_left.Children.Add(sp_path)
        
        sp_pre = StackPanel(); sp_pre.Orientation = Orientation.Horizontal; sp_pre.Margin = Thickness(0,5,0,5)
        self.txtPrefix = TextBox(); self.txtPrefix.Width = 80; self.txtSuffix = TextBox(); self.txtSuffix.Width = 80
        self.txtPrefix.TextChanged += self._on_text_change; self.txtSuffix.TextChanged += self._on_text_change
        sp_pre.Children.Add(TextBlock(Text="Prefix: ")); sp_pre.Children.Add(self.txtPrefix)
        sp_pre.Children.Add(TextBlock(Text=" Suffix: ")); sp_pre.Children.Add(self.txtSuffix)
        sp_left.Children.Add(sp_pre)
        
        # Single Naming
        self.spSingle = StackPanel()
        self.txtPattern = TextBox(); self.txtPattern.Height = 25; self.txtPattern.Text = "{SheetNumber}_{SheetName}"; self.txtPattern.TextChanged += self._on_text_change
        self.spSingle.Children.Add(self.txtPattern)
        sp_tok = StackPanel(); sp_tok.Orientation = Orientation.Horizontal
        for t in ["{SheetNumber}", "{SheetName}", "{Current Revision}"]:
            b = Button(); b.Content = t; b.Click += lambda s,e,v=t: self.txtPattern.AppendText(v); sp_tok.Children.Add(b)
        self.spSingle.Children.Add(sp_tok); sp_left.Children.Add(self.spSingle)
        
        # Combine Naming
        self.spCombine = StackPanel(); self.spCombine.Visibility = Visibility.Collapsed
        self.spCombine.Children.Add(TextBlock(Text="Combined Filename:", FontWeight=FontWeights.Bold))
        self.txtCombineName = TextBox(); self.txtCombineName.Height = 25; self.txtCombineName.Text = "Combined_Set"
        self.spCombine.Children.Add(self.txtCombineName)
        sp_left.Children.Add(self.spCombine)
        
        self.chkExcel = CheckBox(); self.chkExcel.Content = "Excel Transmittal"; self.chkExcel.Margin = Thickness(0,10,0,0); self.chkExcel.Foreground = Brushes.ForestGreen; sp_left.Children.Add(self.chkExcel)
        self.chkAutoFolder = CheckBox(); self.chkAutoFolder.Content = "Auto Create Folders"; self.chkAutoFolder.IsChecked = True; self.chkAutoFolder.Margin = Thickness(0,5,0,0); sp_left.Children.Add(self.chkAutoFolder)
        
        Grid.SetColumn(sp_left, 0); gs.Children.Add(sp_left)
        
        # Parameter Search
        sp_p = StackPanel(); self.lstP = ListBox(); self.lstP.Height = 100
        btn_ap = Button(); btn_ap.Content = " + Insert Param "; btn_ap.Click += self._add_param
        sp_p.Children.Add(TextBlock(Text="Search Parameter:", FontWeight=FontWeights.Bold))
        self.txtPSearch = TextBox(); self.txtPSearch.TextChanged += self._filter_params; sp_p.Children.Add(self.txtPSearch)
        sp_p.Children.Add(self.lstP); sp_p.Children.Add(btn_ap)
        Grid.SetColumn(sp_p, 1); gs.Children.Add(sp_p)
        
        gb_name.Content = gs; Grid.SetRow(gb_name, 1); main_layout.Children.Add(gb_name)

        # 3. FILTER
        gb_sel = GroupBox(); gb_sel.Header = "Filter"; gb_sel.Margin = Thickness(0,0,0,10)
        sp_sel = StackPanel(); sp_sel.Orientation = Orientation.Horizontal; sp_sel.Margin = Thickness(10)
        self.cboSets = ComboBox(); self.cboSets.Width = 150; self.cboSets.SelectionChanged += self._on_set_select; sp_sel.Children.Add(self.cboSets)
        self.txtSearch = TextBox(); self.txtSearch.Width = 150; self.txtSearch.Margin = Thickness(10,0,0,0); self.txtSearch.TextChanged += self._filter_grid; sp_sel.Children.Add(self.txtSearch)
        
        btn_all = Button(); btn_all.Content = " Check All "; btn_all.Click += lambda s,e: self._set_all(True); sp_sel.Children.Add(btn_all)
        btn_none = Button(); btn_none.Content = " Uncheck All "; btn_none.Click += lambda s,e: self._set_all(False); sp_sel.Children.Add(btn_none)
        btn_hi = Button(); btn_hi.Content = " Check Hilighted "; btn_hi.Background = Brushes.LightBlue; btn_hi.Click += self._on_check_high
        sp_sel.Children.Add(btn_hi)
        
        gb_sel.Content = sp_sel; Grid.SetRow(gb_sel, 2); main_layout.Children.Add(gb_sel)

        # 4. GRID
        self.dg = DataGrid(); self.dg.AutoGenerateColumns = False; self.dg.CanUserAddRows = False; self.dg.SelectionMode = DataGridSelectionMode.Extended
        self.dg.Columns.Add(DataGridCheckBoxColumn(Header="X", Binding=Binding("Include")))
        self.dg.Columns.Add(DataGridTextColumn(Header="Number", Binding=Binding("Number"), IsReadOnly=True))
        self.dg.Columns.Add(DataGridTextColumn(Header="Name", Binding=Binding("Name"), Width=DataGridLength(1, DataGridLengthUnitType.Star), IsReadOnly=True))
        self.dg.Columns.Add(DataGridTextColumn(Header="Rev", Binding=Binding("Revision"), IsReadOnly=True, Width=DataGridLength(50)))
        self.dg.Columns.Add(DataGridTextColumn(Header="Preview", Binding=Binding("PreviewName"), Foreground=Brushes.Blue, IsReadOnly=True, Width=DataGridLength(300)))
        Grid.SetRow(self.dg, 4); main_layout.Children.Add(self.dg)

        # 5. RUN
        btn_run = Button(); btn_run.Content = " START MULTI-EXPORT "; btn_run.Height = 50; btn_run.Background = Brushes.ForestGreen; btn_run.Foreground = Brushes.White; btn_run.FontWeight = FontWeights.Bold; btn_run.Click += self._on_run
        Grid.SetRow(btn_run, 5); main_layout.Children.Add(btn_run)

    # --- PARAM LOGIC ---
    def _load_all_params(self):
        p_names = set()
        if self._all_items:
            for p in self._all_items[0].item_obj.Parameters:
                if p.Definition: p_names.add(p.Definition.Name)
        for p in doc.ProjectInformation.Parameters:
            if p.Definition: p_names.add(p.Definition.Name)
        self._all_params = sorted(list(p_names))
        self._filter_params(None, None)

    def _filter_params(self, s, e):
        t = self.txtPSearch.Text.lower()
        self.lstP.Items.Clear()
        for p in self._all_params:
            if t in p.lower(): self.lstP.Items.Add(p)

    def _add_param(self, s, e):
        if self.lstP.SelectedItem: 
            token = "{" + self.lstP.SelectedItem + "}"
            if self.chkCombine.IsChecked: self.txtCombineName.AppendText(token)
            else: self.txtPattern.AppendText(token); self._update_all_previews()

    # --- PROFILE LOGIC (FIXED & SILENT) ---
    def _on_save_profile(self, s, e):
        n = self.txtProfileName.Text.strip()
        if not n: return
        data = {
            'prefix': self.txtPrefix.Text, 'suffix': self.txtSuffix.Text, 'pattern': self.txtPattern.Text,
            'path': self.txtPath.Text, 'color': self.cboColor.SelectedIndex,
            'pdf': self.chkPDF.IsChecked, 'dwg': self.chkDWG.IsChecked, 'ifc': self.chkIFC.IsChecked, 'nwc': self.chkNWC.IsChecked,
            'combine': self.chkCombine.IsChecked, 'combine_name': self.txtCombineName.Text,
            'excel': self.chkExcel.IsChecked, 'auto_folder': self.chkAutoFolder.IsChecked
        }
        # Save overwrite immediately (No Dialog)
        self._profiles[n] = data
        with open(CONFIG_FILE, 'w') as f: json.dump(self._profiles, f)
        self._load_profiles_from_disk()
        
        # Reset Selection (Prevent locking)
        self._ignore_events = True
        self.cboProfile.SelectedItem = n
        self._ignore_events = False
        self._toast("Saved: " + n)

    def _on_profile_select(self, s, e):
        if self._ignore_events: return # หยุดถ้าถูกสั่งให้ข้าม
        
        n = self.cboProfile.SelectedItem
        if n in self._profiles:
            self._ignore_events = True # ล็อกการทำงานชั่วคราว
            
            p = self._profiles[n]
            self.txtProfileName.Text = n 
            self.txtPrefix.Text = p.get('prefix', '')
            self.txtSuffix.Text = p.get('suffix', '')
            self.txtPattern.Text = p.get('pattern', '')
            self.txtPath.Text = p.get('path', '')
            self.cboColor.SelectedIndex = p.get('color', 0)
            self.chkPDF.IsChecked = p.get('pdf', True)
            self.chkDWG.IsChecked = p.get('dwg', False)
            self.chkIFC.IsChecked = p.get('ifc', False)
            self.chkNWC.IsChecked = p.get('nwc', False)
            self.chkCombine.IsChecked = p.get('combine', False)
            self.txtCombineName.Text = p.get('combine_name', 'Combined_Set')
            self.chkExcel.IsChecked = p.get('excel', False)
            self.chkAutoFolder.IsChecked = p.get('auto_folder', True)
            
            self._ignore_events = False # ปลดล็อก
            
            self._toggle_ui(None, None)
            self._update_all_previews()

    def _on_delete_profile(self, s, e):
        n = self.cboProfile.SelectedItem
        if n in self._profiles:
            # Delete immediately (No Dialog)
            del self._profiles[n]
            with open(CONFIG_FILE, 'w') as f: json.dump(self._profiles, f)
            self._load_profiles_from_disk(); self.txtProfileName.Text = ""
            self._toast("Deleted")

    def _on_import_profile(self, s, e):
        src = forms.pick_file(file_ext='json')
        if src:
            try:
                self._profiles.update(json.load(open(src)))
                with open(CONFIG_FILE, 'w') as f: json.dump(self._profiles, f)
                self._load_profiles_from_disk(); self._toast("Imported")
            except: pass

    def _on_export_profile(self, s, e):
        dest = forms.save_file(file_ext='json')
        if dest:
            with open(dest, 'w') as f: json.dump(self._profiles, f)
            self._toast("Exported")

    # --- CORE EXPORT ---
    def _sanitize(self, name):
        return re.sub(r'[\\/*?:"<>|]', "_", name).strip()

    def _on_run(self, s, e):
        selected = [i for i in self.dg.ItemsSource if i.Include]
        if not selected: return
        if not self.txtPath.Text: forms.alert("Select Path"); return
        self._save_current_as_last()
        folder = self.txtPath.Text
        
        pdf_path = os.path.join(folder, "PDF") if self.chkAutoFolder.IsChecked else folder
        dwg_path = os.path.join(folder, "DWG") if self.chkAutoFolder.IsChecked else folder
        if self.chkAutoFolder.IsChecked:
            for d in [pdf_path, dwg_path]: 
                if not os.path.exists(d): os.makedirs(d)

        try:
            total = len(selected) * (int(self.chkPDF.IsChecked) + int(self.chkDWG.IsChecked) + int(self.chkIFC.IsChecked) + int(self.chkNWC.IsChecked))
            if self.chkCombine.IsChecked: total += 1

            with forms.ProgressBar(title="Working...", total=total) as pb:
                c = 0
                # Combine
                if self.chkCombine.IsChecked and self.chkPDF.IsChecked:
                    try:
                        raw_c_name = self.txtCombineName.Text
                        if "{" in raw_c_name and selected:
                            for p_name in [x.strip("{}") for x in re.findall(r'\{.*?\}', raw_c_name)]:
                                val = selected[0].get_param_val(p_name)
                                raw_c_name = raw_c_name.replace("{" + p_name + "}", val)
                        
                        final_c_name = self._sanitize(raw_c_name)
                        
                        opt = DB.PDFExportOptions(); opt.Combine = True; opt.FileName = final_c_name
                        opt.ColorDepth = DB.ColorDepthType.BlackLine if self.cboColor.SelectedIndex == 1 else DB.ColorDepthType.Color
                        doc.Export(pdf_path, List[DB.ElementId]([i.Id for i in selected]), opt)
                    except Exception as ex: print("Combine Error: " + str(ex))
                    c += 1; pb.update_progress(c, total)

                # Individual
                for item in selected:
                    fn = self._sanitize(item.PreviewName)
                    
                    if self.chkPDF.IsChecked and not self.chkCombine.IsChecked:
                        try:
                            opt = DB.PDFExportOptions(); opt.FileName = fn
                            opt.ColorDepth = DB.ColorDepthType.BlackLine if self.cboColor.SelectedIndex == 1 else DB.ColorDepthType.Color
                            doc.Export(pdf_path, List[DB.ElementId]([item.Id]), opt)
                        except: pass
                        c += 1; pb.update_progress(c, total)
                    
                    if self.chkDWG.IsChecked:
                        try:
                            clean_fn = fn.replace(".dwg", "")
                            opt_dwg = DB.DWGExportOptions(); opt_dwg.MergedViews = True 
                            doc.Export(dwg_path, clean_fn, List[DB.ElementId]([item.Id]), opt_dwg)
                        except Exception as ex: print("DWG Error {}: {}".format(fn, ex))
                        c += 1; pb.update_progress(c, total)

                    if self.chkIFC.IsChecked:
                        try: doc.Export(folder, fn, DB.IFCExportOptions())
                        except: pass
                        c += 1; pb.update_progress(c, total)
                    if self.chkNWC.IsChecked:
                        try: doc.Export(folder, fn, DB.NavisworksExportOptions())
                        except: pass
                        c += 1; pb.update_progress(c, total)
            
            if self.chkExcel.IsChecked: self._create_excel(selected, folder)
            self.Close(); os.startfile(folder); self._toast("Done!")
        except Exception as ex: 
            forms.alert("Critical Error:\n" + str(ex))

    # --- HELPERS ---
    def _toast(self, msg):
        try: output.get_output().toast(msg)
        except: pass

    def _create_excel(self, items, folder):
        try:
            csv_p = os.path.join(folder, "Transmittal_{}.csv".format(datetime.datetime.now().strftime("%Y%m%d")))
            with open(csv_p, 'w') as f:
                f.write("Number,Name,Revision,Date\n")
                for i in items: f.write("{},{},{},{}\n".format(i.Number, i.Name, i.Revision, datetime.datetime.now().date()))
        except: pass

    def _update_all_previews(self):
        if self._ignore_events: return # SKIP UPDATE IF LOCKED
        pat = self.txtPattern.Text; pre = self.txtPrefix.Text; suf = self.txtSuffix.Text
        for item in self._all_items:
            fn = pre + pat + suf
            fn = fn.replace("{SheetNumber}", item.Number).replace("{SheetName}", item.Name).replace("{Current Revision}", item.Revision)
            if "{" in fn:
                for p_name in [x.strip("{}") for x in re.findall(r'\{.*?\}', fn)]:
                    val = item.get_param_val(p_name)
                    fn = fn.replace("{" + p_name + "}", val)
            item.PreviewName = self._sanitize(fn)
        self.dg.Items.Refresh()

    def _on_text_change(self, s, e): self._update_all_previews()
    def _refresh_data(self):
        sheets = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).ToElements()
        self._all_items = [ExportItem(s) for s in sheets if not s.IsPlaceholder]
        self._all_items.sort(key=lambda x: x.Number)
        self.dg.ItemsSource = ObservableCollection[ExportItem](self._all_items)

    def _on_check_high(self, s, e):
        for i in self.dg.SelectedItems: i.Include = True
        self.dg.Items.Refresh()

    def _set_all(self, val):
        for i in self.dg.ItemsSource: i.Include = val
        self.dg.Items.Refresh()

    def _filter_grid(self, s, e):
        t = self.txtSearch.Text.lower()
        self.dg.ItemsSource = ObservableCollection[ExportItem]([i for i in self._all_items if t in i.Number.lower() or t in i.Name.lower()])

    def _toggle_ui(self, s, e):
        self.spCombine.Visibility = Visibility.Visible if self.chkCombine.IsChecked else Visibility.Collapsed
        self.spSingle.Visibility = Visibility.Collapsed if self.chkCombine.IsChecked else Visibility.Visible

    def _on_browse(self, s, e):
        p = forms.pick_folder()
        if p: self.txtPath.Text = p

    def _on_open_folder(self, s, e):
        if os.path.exists(self.txtPath.Text): os.startfile(self.txtPath.Text)

    def _load_print_sets(self):
        self.cboSets.Items.Add("- Sets -")
        for s in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheetSet): self.cboSets.Items.Add(s.Name)
        self.cboSets.SelectedIndex = 0

    def _on_set_select(self, s, e):
        n = self.cboSets.SelectedItem
        if n and n != "- Sets -":
            vset = next((x for x in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheetSet) if x.Name == n), None)
            if vset:
                ids = [v.Id for v in vset.Views]
                for i in self._all_items: i.Include = i.Id in ids
                self.dg.Items.Refresh()

    def _load_profiles_from_disk(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f: self._profiles = json.load(f)
            self.cboProfile.Items.Clear()
            for k in sorted(self._profiles.keys()): self.cboProfile.Items.Add(k)

    def _load_project_params_list(self):
        for p in ["SheetNumber", "SheetName", "Current Revision", "Drawn By", "Checked By", "Approved By"]: self.lstP.Items.Add(p)

    def _save_current_as_last(self):
        json.dump({'path': self.txtPath.Text, 'pattern': self.txtPattern.Text}, open(LAST_SETTING_FILE, 'w'))

    def _load_last_settings(self):
        if os.path.exists(LAST_SETTING_FILE):
            d = json.load(open(LAST_SETTING_FILE))
            self.txtPath.Text = d.get('path', '')
            self.txtPattern.Text = d.get('pattern', '{SheetNumber}_{SheetName}')

if __name__ == '__main__':
    SuperSheetsUltimate().ShowDialog()