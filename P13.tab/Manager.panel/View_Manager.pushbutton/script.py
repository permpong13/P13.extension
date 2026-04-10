# -*- coding: utf-8 -*-
"""View Manager Tool - Final Stable Edition for Revit 2026"""
__title__ = 'View\nManager'
__author__ = 'เพิ่มพงษ์'

import os
import sys
import clr
import json
import traceback
from collections import defaultdict

# CLR references
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('System')

from System import Array
from System.Windows.Forms import (
    DialogResult, SaveFileDialog, Form, Button, Label, 
    ComboBox, GroupBox, ListView, View, TabControl, TabPage, BorderStyle, 
    TextBox, FormStartPosition, ComboBoxStyle, 
    Application, Cursor, Cursors, Control
)
from System.Drawing import Point, Size, Font, FontStyle, Color, SystemColors, ContentAlignment

from pyrevit import revit, DB, forms, script

doc = revit.doc
logger = script.get_logger()

# --- Settings Management ---
SETTINGS_PATH = os.path.join(os.getenv('APPDATA'), 'P13_ViewManager_Settings.json')

def get_saved_path():
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, 'r') as f:
                return json.load(f).get('export_path')
        except: return None
    return None

def save_export_path(path):
    try:
        with open(SETTINGS_PATH, 'w') as f:
            json.dump({'export_path': path}, f)
    except: pass

def sanitize_view_name(name):
    prohibited = '<>?*|:;[]{}'
    for ch in prohibited:
        name = name.replace(ch, '_')
    return name.strip()

# --- Main UI ---
class EnhancedViewManagerForm(Form):
    def __init__(self):
        self.Text = "View Manager - Advanced Filters (Stable v2.5)"
        self.Size = Size(1200, 950)
        self.StartPosition = FormStartPosition.CenterScreen
        self.Font = Font("Segoe UI", 9)
        self.MinimumSize = Size(1100, 850)
        
        self.all_views_data = []
        self.current_tab = "orphan"
        
        self.InitializeComponents()
        self.RefreshData()

    def InitializeComponents(self):
        # Header
        header = Label(Text="ระบบจัดการ View และ Filters ขั้นสูง", 
                     Location=Point(20, 10), Size=Size(1140, 40),
                     Font=Font("Segoe UI", 16, FontStyle.Bold), ForeColor=Color.FromArgb(45, 45, 45), 
                     TextAlign=ContentAlignment.MiddleLeft)
        self.Controls.Add(header)

        # Tab Control
        self.tab_control = TabControl(Location=Point(20, 60), Size=Size(1140, 650))
        self.tab_orphan = TabPage(Text="View ที่ยังไม่ได้ใช้ (Orphan)", BackColor=Color.White)
        self.tab_all = TabPage(Text="View ทั้งหมดในโครงการ", BackColor=Color.White)
        self.tab_control.Controls.AddRange(Array[Control]([self.tab_orphan, self.tab_all]))
        self.tab_control.SelectedIndexChanged += self.OnTabChanged
        self.Controls.Add(self.tab_control)

        self.InitFilterUI(self.tab_orphan, "o")
        self.InitFilterUI(self.tab_all, "a")
        self.InitActionPanel()

    def InitFilterUI(self, parent_tab, prefix):
        # Filter Group
        filter_group = GroupBox(Text="การจัดการตัวกรอง (Filters Management)", 
                              Location=Point(15, 10), Size=Size(1100, 130), Font=Font("Segoe UI", 9, FontStyle.Bold))
        
        lbl_style = Font("Segoe UI", 8.5)
        
        # Row 1: Search & Type
        Label(Text="ค้นหาชื่อ:", Location=Point(15, 30), Size=Size(60, 20), Font=lbl_style, Parent=filter_group)
        search_txt = TextBox(Location=Point(80, 27), Size=Size(400, 25), Parent=filter_group)
        
        Label(Text="ประเภท:", Location=Point(500, 30), Size=Size(60, 20), Font=lbl_style, Parent=filter_group)
        type_cb = ComboBox(Location=Point(560, 27), Size=Size(200, 25), DropDownStyle=ComboBoxStyle.DropDownList, Parent=filter_group)
        
        # Row 2: Discipline, Phase, Detail
        Label(Text="สาขา (Disc):", Location=Point(15, 65), Size=Size(70, 20), Font=lbl_style, Parent=filter_group)
        disc_cb = ComboBox(Location=Point(80, 62), Size=Size(150, 25), DropDownStyle=ComboBoxStyle.DropDownList, Parent=filter_group)
        
        Label(Text="Phase:", Location=Point(245, 65), Size=Size(50, 20), Font=lbl_style, Parent=filter_group)
        phase_cb = ComboBox(Location=Point(300, 62), Size=Size(180, 25), DropDownStyle=ComboBoxStyle.DropDownList, Parent=filter_group)

        Label(Text="Detail:", Location=Point(500, 65), Size=Size(50, 20), Font=lbl_style, Parent=filter_group)
        detail_cb = ComboBox(Location=Point(560, 62), Size=Size(120, 25), DropDownStyle=ComboBoxStyle.DropDownList, Parent=filter_group)

        count_lbl = Label(Text="พบ 0 รายการ", Location=Point(15, 100), Size=Size(200, 20), ForeColor=Color.Blue, Parent=filter_group)
        reset_btn = Button(Text="ล้างตัวกรอง", Location=Point(980, 25), Size=Size(100, 65), BackColor=Color.GhostWhite, Parent=filter_group)

        setattr(self, prefix + "_search_txt", search_txt)
        setattr(self, prefix + "_type_cb", type_cb)
        setattr(self, prefix + "_disc_cb", disc_cb)
        setattr(self, prefix + "_phase_cb", phase_cb)
        setattr(self, prefix + "_detail_cb", detail_cb)
        setattr(self, prefix + "_count_lbl", count_lbl)

        # List View - FIXED: Using .Add() instead of .AddRange() for Stability
        lv = ListView(Location=Point(15, 150), Size=Size(1100, 480), View=View.Details, 
                    FullRowSelect=True, GridLines=True, CheckBoxes=True)
        
        lv.Columns.Add("เลือก", 50)
        lv.Columns.Add("ชื่อ View", 380)
        lv.Columns.Add("ประเภท", 150)
        lv.Columns.Add("Discipline", 120)
        lv.Columns.Add("Phase", 120)
        lv.Columns.Add("Scale", 70)
        lv.Columns.Add("Detail", 100)
        
        setattr(self, prefix + "_lv", lv)
        parent_tab.Controls.AddRange(Array[Control]([filter_group, lv]))

        # Events
        search_txt.TextChanged += self.ApplyFilters
        for cb in [type_cb, disc_cb, phase_cb, detail_cb]:
            cb.SelectedIndexChanged += self.ApplyFilters
        reset_btn.Click += lambda s, e: self.ResetAllFilters(prefix)

    def InitActionPanel(self):
        action_group = GroupBox(Text="การดำเนินการ (Batch Actions)", Location=Point(20, 720), Size=Size(1140, 170))
        
        btn_data = [
            ("ลบ View", self.DeleteViews, Color.MistyRose),
            ("Prefix", self.AddPrefix, Color.White),
            ("Rename {n}", self.BatchRename, Color.White),
            ("Set Detail", self.SetDetailLevel, Color.White),
            ("Export Report", self.ExportToTxt, Color.LightCyan),
            ("Check Dups", self.CheckDups, Color.White),
            ("Stats", self.ShowStats, Color.White),
            ("Refresh Data", lambda s, e: self.RefreshData(), Color.Honeydew)
        ]

        for i, (txt, func, color) in enumerate(btn_data):
            btn = Button(Text=txt, Size=Size(170, 45), BackColor=color,
                        Location=Point(20 + (185 * (i % 4)), 35 + (55 * (i // 4))))
            btn.Click += func
            action_group.Controls.Add(btn)

        self.status_bar = Label(Text="Ready", Location=Point(760, 35), Size=Size(350, 100), 
                              BorderStyle=BorderStyle.FixedSingle, TextAlign=ContentAlignment.MiddleCenter, BackColor=Color.White)
        action_group.Controls.Add(self.status_bar)
        self.Controls.Add(action_group)

    def RefreshData(self):
        Cursor.Current = Cursors.WaitCursor
        self.status_bar.Text = "กำลังอ่านข้อมูล View จาก Revit..."
        Application.DoEvents()

        self.all_views_data = []
        placed_ids = set()
        for s in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet):
            for vp_id in s.GetAllViewports():
                vp = doc.GetElement(vp_id)
                if vp: placed_ids.add(vp.ViewId.Value)

        all_views = DB.FilteredElementCollector(doc).OfClass(DB.View).WhereElementIsNotElementType().ToElements()
        types, discs, phases, details = set(), set(), set(), set()

        for v in all_views:
            if v.IsTemplate or v.ViewType in [DB.ViewType.Internal, DB.ViewType.ProjectBrowser]: continue
            
            try:
                phase_param = v.get_Parameter(DB.BuiltInParameter.VIEW_PHASE)
                phase_name = doc.GetElement(phase_param.AsElementId()).Name if phase_param and phase_param.AsElementId() != DB.ElementId.InvalidElementId else "None"
                
                info = {
                    'view': v,
                    'name': v.Name,
                    'type': v.ViewType.ToString(),
                    'disc': v.Discipline.ToString() if hasattr(v, 'Discipline') else "None",
                    'phase': phase_name,
                    'detail': v.DetailLevel.ToString() if hasattr(v, 'DetailLevel') else "N/A",
                    'scale': "1:{}".format(v.Scale) if v.Scale > 0 else "N/A",
                    'is_orphan': v.Id.Value not in placed_ids
                }
                self.all_views_data.append(info)
                
                types.add(info['type'])
                discs.add(info['disc'])
                phases.add(info['phase'])
                details.add(info['detail'])
            except: continue

        for prefix in ["o", "a"]:
            self.PopulateCB(getattr(self, prefix + "_type_cb"), types)
            self.PopulateCB(getattr(self, prefix + "_disc_cb"), discs)
            self.PopulateCB(getattr(self, prefix + "_phase_cb"), phases)
            self.PopulateCB(getattr(self, prefix + "_detail_cb"), details)

        self.ApplyFilters(None, None)
        self.status_bar.Text = "โหลดเสร็จสิ้น\nพบทั้งหมด {} Views".format(len(self.all_views_data))
        Cursor.Current = Cursors.Default

    def PopulateCB(self, cb, items):
        cb.Items.Clear()
        cb.Items.Add("--- ทั้งหมด ---")
        for item in sorted(list(items)):
            cb.Items.Add(item)
        cb.SelectedIndex = 0

    def ApplyFilters(self, sender, args):
        prefix = "a" if self.tab_control.SelectedIndex == 1 else "o"
        search = getattr(self, prefix + "_search_txt").Text.lower()
        v_type = getattr(self, prefix + "_type_cb").SelectedItem
        v_disc = getattr(self, prefix + "_disc_cb").SelectedItem
        v_phase = getattr(self, prefix + "_phase_cb").SelectedItem
        v_det = getattr(self, prefix + "_detail_cb").SelectedItem
        
        filtered = []
        for info in self.all_views_data:
            if prefix == "o" and not info['is_orphan']: continue
            if search and search not in info['name'].lower(): continue
            if v_type != "--- ทั้งหมด ---" and info['type'] != v_type: continue
            if v_disc != "--- ทั้งหมด ---" and info['disc'] != v_disc: continue
            if v_phase != "--- ทั้งหมด ---" and info['phase'] != v_phase: continue
            if v_det != "--- ทั้งหมด ---" and info['detail'] != v_det: continue
            filtered.append(info)

        lv = getattr(self, prefix + "_lv")
        lv.BeginUpdate()
        lv.Items.Clear()
        for info in filtered:
            item = lv.Items.Add("")
            item.SubItems.AddRange(Array[str]([
                info['name'], info['type'], info['disc'], 
                info['phase'], info['scale'], info['detail']
            ]))
            item.Tag = info['view'].Id
        lv.EndUpdate()
        getattr(self, prefix + "_count_lbl").Text = "แสดงผล {} รายการ".format(len(filtered))

    def ResetAllFilters(self, prefix):
        getattr(self, prefix + "_search_txt").Text = ""
        getattr(self, prefix + "_type_cb").SelectedIndex = 0
        getattr(self, prefix + "_disc_cb").SelectedIndex = 0
        getattr(self, prefix + "_phase_cb").SelectedIndex = 0
        getattr(self, prefix + "_detail_cb").SelectedIndex = 0
        self.ApplyFilters(None, None)

    def GetChecked(self):
        lv = self.all_lv if self.tab_control.SelectedIndex == 1 else self.o_lv
        return [doc.GetElement(item.Tag) for item in lv.CheckedItems if item.Tag]

    # --- Actions ---
    def DeleteViews(self, s, e):
        sel = self.GetChecked()
        if not sel: return
        if forms.alert("ยืนยันการลบ {} รายการ?".format(len(sel)), yes=True, no=True):
            with revit.Transaction("Batch Delete"):
                for v in sel:
                    try: doc.Delete(v.Id)
                    except: pass
            self.RefreshData()

    def AddPrefix(self, s, e):
        sel = self.GetChecked()
        if not sel: return
        prefix = forms.ask_for_string("Prefix:", title="Add Prefix")
        if prefix:
            with revit.Transaction("Batch Prefix"):
                for v in sel:
                    try: v.Name = sanitize_view_name("{} {}".format(prefix, v.Name))
                    except: pass
            self.RefreshData()

    def BatchRename(self, s, e):
        sel = self.GetChecked()
        if not sel: return
        pattern = forms.ask_for_string("ชื่อใหม่ (ใช้ {n} รันเลข):", title="Rename")
        if pattern:
            with revit.Transaction("Batch Rename"):
                for i, v in enumerate(sel, 1):
                    try: v.Name = sanitize_view_name(pattern.replace("{n}", str(i)))
                    except: pass
            self.RefreshData()

    def SetDetailLevel(self, s, e):
        sel = self.GetChecked()
        if not sel: return
        res = forms.SelectFromList.show(["Coarse", "Medium", "Fine"], title="Detail Level")
        if res:
            lvl = {"Coarse": DB.ViewDetailLevel.Coarse, "Medium": DB.ViewDetailLevel.Medium, "Fine": DB.ViewDetailLevel.Fine}[res]
            with revit.Transaction("Set Detail"):
                for v in sel:
                    if hasattr(v, "DetailLevel"): v.DetailLevel = lvl
            self.RefreshData()

    def ExportToTxt(self, s, e):
        last_path = get_saved_path()
        save_file = SaveFileDialog()
        save_file.Filter = "Text Files|*.txt"
        save_file.FileName = "ViewReport_{}.txt".format(doc.Title)
        if last_path and os.path.exists(last_path): save_file.InitialDirectory = last_path
        
        if save_file.ShowDialog() == DialogResult.OK:
            path = save_file.FileName
            save_export_path(os.path.dirname(path))
            with open(path, 'w', encoding='utf-8') as f:
                f.write("View Report: {}\n\n".format(doc.Title))
                for info in self.all_views_data:
                    f.write("[{}] {} | {}\n".format("SHEET" if not info['is_orphan'] else "FREE", info['name'], info['type']))
            os.startfile(path)

    def ShowStats(self, s, e):
        orphans = [v for v in self.all_views_data if v['is_orphan']]
        forms.alert("ทั้งหมด: {}\nไม่ได้ใช้: {}\nอยู่ใน Sheet: {}".format(len(self.all_views_data), len(orphans), len(self.all_views_data)-len(orphans)))

    def CheckDups(self, s, e):
        counts = defaultdict(int)
        for info in self.all_views_data: counts[info['name']] += 1
        dups = [n for n, c in counts.items() if c > 1]
        if dups: print("ชื่อซ้ำ:\n" + "\n".join(dups))
        else: forms.alert("ไม่พบชื่อซ้ำ")

    def OnTabChanged(self, s, e):
        self.current_tab = "all" if self.tab_control.SelectedIndex == 1 else "orphan"
        self.ApplyFilters(None, None)

if __name__ == "__main__":
    try:
        form = EnhancedViewManagerForm()
        form.ShowDialog()
    except Exception:
        print(traceback.format_exc())