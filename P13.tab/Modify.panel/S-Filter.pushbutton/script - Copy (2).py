# -*- coding: utf-8 -*-
from pyrevit import revit, forms, script
import json
import os
import clr

clr.AddReference('PresentationFramework')
from System.Windows.Controls import TreeViewItem, StackPanel, CheckBox, TextBlock, Orientation
from System.Windows import Thickness, VerticalAlignment
from System.Windows.Input import Key

from Autodesk.Revit.DB import *
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc

# ==========================================
# 1. ระบบความจำตั้งค่า (Export Path & Saved Selection)
# ==========================================
CONFIG_FILE = os.path.join(os.path.expanduser("~"), "pyrevit_superfilter_config.json")

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    return {"export_path": "", "saved_selection": []}

def save_config(config):
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)

config = load_config()

if not config.get("export_path"):
    selected_folder = forms.pick_folder(title="ไม่พบ Export Path: กรุณาเลือกโฟลเดอร์สำหรับทำงาน")
    if selected_folder:
        config["export_path"] = selected_folder
        save_config(config)
        forms.alert("ตั้งค่า Export Path เรียบร้อย: {}".format(selected_folder))
    else:
        forms.alert("คุณยังไม่ได้เลือก Export Path ฟังก์ชันบางอย่างอาจทำงานไม่สมบูรณ์", exitscript=False)


# ==========================================
# 2. คลาสควบคุม S-Filter
# ==========================================
class SuperFilterWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name):
        forms.WPFWindow.__init__(self, xaml_file_name)
        
        self.selection_ids = uidoc.Selection.GetElementIds()
        self.node_children = {}
        self.leaf_checkboxes = [] 
        
        if hasattr(self, 'btn_run'): self.btn_run.Click += self.run_select
        if hasattr(self, 'btn_isolate'): self.btn_isolate.Click += self.run_isolate
        if hasattr(self, 'btn_hide'): self.btn_hide.Click += self.run_hide
        if hasattr(self, 'btn_filter_param'): self.btn_filter_param.Click += self.run_parameter_filter
            
        # ✨ ผูกปุ่ม Memory
        if hasattr(self, 'btn_save_mem'): self.btn_save_mem.Click += self.run_save_memory
        if hasattr(self, 'btn_restore_mem'): self.btn_restore_mem.Click += self.run_restore_memory
            
        if hasattr(self, 'btn_check_all'): self.btn_check_all.Click += self.check_all_items
        if hasattr(self, 'btn_uncheck_all'): self.btn_uncheck_all.Click += self.uncheck_all_items
        if hasattr(self, 'btn_expand_all'): self.btn_expand_all.Click += self.expand_all_items
        if hasattr(self, 'btn_collapse_all'): self.btn_collapse_all.Click += self.collapse_all_items
        if hasattr(self, 'btn_invert_all'): self.btn_invert_all.Click += self.invert_all_items
            
        if hasattr(self, 'rad_current_sel'): self.rad_current_sel.Checked += self.refresh_tree
        if hasattr(self, 'rad_all_model'): self.rad_all_model.Checked += self.refresh_tree
        if hasattr(self, 'rad_visible_view'): self.rad_visible_view.Checked += self.refresh_tree
        if hasattr(self, 'rad_belong_view'): self.rad_belong_view.Checked += self.refresh_tree
            
        if hasattr(self, 'txt_search'):
            self.txt_search.KeyDown += self.on_search_enter
            
        self.populate_categories()

    def on_search_enter(self, sender, args):
        if args.Key == Key.Enter:
            self.populate_categories()
            self.txt_search.Focus()
            self.txt_search.CaretIndex = len(self.txt_search.Text)

    def on_node_check(self, sender, args):
        is_checked = sender.IsChecked
        def cascade(parent_chk, state):
            if parent_chk in self.node_children:
                for child in self.node_children[parent_chk]:
                    child.IsChecked = state
                    cascade(child, state)
        cascade(sender, is_checked)

    def check_all_items(self, sender, args):
        if hasattr(self, 'root_chk'):
            self.root_chk.IsChecked = True
            self.on_node_check(self.root_chk, None)

    def uncheck_all_items(self, sender, args):
        if hasattr(self, 'root_chk'):
            self.root_chk.IsChecked = False
            self.on_node_check(self.root_chk, None)
            
    def invert_all_items(self, sender, args):
        for chk, _ in self.leaf_checkboxes:
            chk.IsChecked = not chk.IsChecked
            
    def _expand_nodes(self, items, state):
        for item in items:
            item.IsExpanded = state
            self._expand_nodes(item.Items, state)

    def expand_all_items(self, sender, args):
        self._expand_nodes(self.ui_treeview.Items, True)

    def collapse_all_items(self, sender, args):
        self._expand_nodes(self.ui_treeview.Items, False)
        if hasattr(self, 'root_item'): self.root_item.IsExpanded = True

    def refresh_tree(self, sender, args):
        self.populate_categories()

    def create_tree_node(self, name, count):
        item = TreeViewItem()
        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal
        sp.Margin = Thickness(0, 2, 0, 2)
        
        chk = CheckBox()
        chk.IsChecked = False 
        chk.VerticalAlignment = VerticalAlignment.Center
        chk.Margin = Thickness(0, 0, 5, 0)
        chk.Click += self.on_node_check 
        
        txt = TextBlock()
        txt.Text = "{}   [{}]".format(name, count)
        txt.VerticalAlignment = VerticalAlignment.Center
        
        sp.Children.Add(chk)
        sp.Children.Add(txt)
        item.Header = sp
        return item, chk

    def populate_categories(self):
        if not hasattr(self, 'ui_treeview'): return
            
        self.ui_treeview.Items.Clear()
        self.node_children = {}
        self.leaf_checkboxes = []
        elements = []

        if hasattr(self, 'rad_current_sel') and self.rad_current_sel.IsChecked:
            if self.selection_ids: elements = [doc.GetElement(eid) for eid in self.selection_ids]
        elif hasattr(self, 'rad_visible_view') and self.rad_visible_view.IsChecked:
            elements = [el for el in FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType().ToElements() if el.Category]
        elif hasattr(self, 'rad_belong_view') and self.rad_belong_view.IsChecked:
            elements = [el for el in FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType().ToElements() if el.Category and el.ViewSpecific]
        elif hasattr(self, 'rad_all_model') and self.rad_all_model.IsChecked:
            elements = [el for el in FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements() if el.Category]

        search_text = self.txt_search.Text.lower() if hasattr(self, 'txt_search') else ""
            
        cat_dict = {}
        for el in elements:
            cat_name = el.Category.Name if el.Category else "Uncategorized"
            type_name = "{} : {}".format(el.Symbol.Family.Name, el.Name) if (isinstance(el, FamilyInstance) and el.Symbol) else getattr(el, 'Name', type(el).__name__)
            if not type_name: type_name = "Unknown Type"
            
            if search_text and search_text not in cat_name.lower() and search_text not in type_name.lower():
                continue
                
            if cat_name not in cat_dict: cat_dict[cat_name] = {}
            if type_name not in cat_dict[cat_name]: cat_dict[cat_name][type_name] = []
            cat_dict[cat_name][type_name].append(el.Id)
            
        if hasattr(self, 'lbl_status_count'): self.lbl_status_count.Text = str(len(elements))
        if not cat_dict: return

        self.root_item, self.root_chk = self.create_tree_node("All", sum([len(ids) for c in cat_dict.values() for ids in c.values()]))
        self.ui_treeview.Items.Add(self.root_item)
        self.node_children[self.root_chk] = []
        
        for cat_name in sorted(cat_dict.keys()):
            cat_eids = [eid for t_list in cat_dict[cat_name].values() for eid in t_list]
            cat_item, cat_chk = self.create_tree_node(cat_name, len(cat_eids))
            self.node_children[self.root_chk].append(cat_chk)
            self.node_children[cat_chk] = []
            self.root_item.Items.Add(cat_item)
            
            for type_name in sorted(cat_dict[cat_name].keys()):
                type_eids = cat_dict[cat_name][type_name]
                type_item, type_chk = self.create_tree_node(type_name, len(type_eids))
                self.node_children[cat_chk].append(type_chk)
                self.leaf_checkboxes.append((type_chk, type_eids)) 
                cat_item.Items.Add(type_item)
                
        self.root_item.IsExpanded = True

    def _get_id_val(self, eid):
        if hasattr(eid, "Value"): return eid.Value
        if hasattr(eid, "IntegerValue"): return eid.IntegerValue
        return eid 

    def get_final_elements(self):
        final_eids = {} 
        
        for chk, eids in self.leaf_checkboxes:
            if chk.IsChecked:
                for eid in eids:
                    val = self._get_id_val(eid)
                    final_eids[val] = eid
                
        if not final_eids: return List[ElementId]()
        
        base_values = list(final_eids.keys()) 
        for val in base_values:
            eid = final_eids[val]
            el = doc.GetElement(eid)
            if not el: continue
            
            if self.chk_exp_host.IsChecked and hasattr(el, "Host") and el.Host:
                h_id = el.Host.Id
                h_val = self._get_id_val(h_id)
                final_eids[h_val] = h_id
                
            if self.chk_exp_level.IsChecked and el.LevelId != ElementId.InvalidElementId:
                if el.Category:
                    lv_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategoryId(el.Category.Id).ToElements()
                    for l_el in lv_elements:
                        if l_el.LevelId == el.LevelId:
                            l_id = l_el.Id
                            l_val = self._get_id_val(l_id)
                            final_eids[l_val] = l_id
                            
            if self.chk_exp_workset.IsChecked and doc.IsWorkshared:
                if el.Category:
                    ws_id = el.WorksetId
                    ws_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategoryId(el.Category.Id).ToElements()
                    for w_el in ws_elements:
                        if w_el.WorksetId == ws_id:
                            w_id = w_el.Id
                            w_val = self._get_id_val(w_id)
                            final_eids[w_val] = w_id

        out_list = List[ElementId]()
        for eid in final_eids.values(): out_list.Add(eid)
            
        return out_list

    # --- ✨ ฟังก์ชัน Memory (จำและเรียกคืน) ---
    def run_save_memory(self, sender, args):
        ids = self.get_final_elements()
        if not ids or ids.Count == 0:
            forms.alert("กรุณาติ๊กเลือกวัตถุใน Tree View ก่อนกดบันทึกค่ะ")
            return
            
        # บันทึกเป็น IntegerValue หรือ Value ตามเวอร์ชันของ Revit ลงใน Config
        config["saved_selection"] = [self._get_id_val(eid) for eid in ids]
        save_config(config)
        
        forms.alert("บันทึกวัตถุลงหน่วยความจำแล้วจำนวน {} ชิ้น! (กด Restore เพื่อเรียกคืนได้ตลอดเวลา)".format(ids.Count))

    def run_restore_memory(self, sender, args):
        saved_ids = config.get("saved_selection", [])
        if not saved_ids:
            forms.alert("ยังไม่มีการบันทึกข้อมูลในหน่วยความจำค่ะ")
            return
            
        restore_list = List[ElementId]()
        for val in saved_ids:
            try:
                # รองรับการแปลงกลับเป็น ElementId ในทุกเวอร์ชัน
                eid = ElementId(val)
                if doc.GetElement(eid):
                    restore_list.Add(eid)
            except:
                pass
                
        if restore_list.Count > 0:
            uidoc.Selection.SetElementIds(restore_list)
            forms.alert("S-Filter: เรียกคืนวัตถุจำนวน {} ชิ้น สำเร็จ!".format(restore_list.Count))
            self.Close()
        else:
            forms.alert("วัตถุที่เคยจำไว้ ไม่มีอยู่ในโมเดลนี้แล้วค่ะ")

    # --- ปุ่ม Action ต่างๆ ---
    def run_select(self, sender, args):
        ids = self.get_final_elements()
        if ids and ids.Count > 0: uidoc.Selection.SetElementIds(ids)
        self.Close()

    def run_isolate(self, sender, args):
        ids = self.get_final_elements()
        if ids and ids.Count > 0:
            uidoc.Selection.SetElementIds(ids)
            try:
                t = Transaction(doc, "S-Filter Isolate")
                t.Start()
                doc.ActiveView.IsolateElementsTemporary(ids)
                t.Commit()
            except:
                pass 
        self.Close()

    def run_hide(self, sender, args):
        ids = self.get_final_elements()
        if ids and ids.Count > 0:
            try:
                t = Transaction(doc, "S-Filter Hide")
                t.Start()
                doc.ActiveView.HideElementsTemporary(ids)
                t.Commit()
            except:
                pass
        self.Close()

    def run_parameter_filter(self, sender, args):
        self.Hide() 
        
        ids = self.get_final_elements()
        if not ids or ids.Count == 0:
            forms.alert("กรุณาติ๊กเลือกวัตถุใน Tree View ก่อนใช้คำสั่งนี้ค่ะ")
            self.ShowDialog()
            return
            
        elements = [doc.GetElement(eid) for eid in ids if doc.GetElement(eid)]
        if not elements:
            self.ShowDialog()
            return

        param_names = set()
        for el in elements:
            for p in el.Parameters:
                if p.Definition.Name: param_names.add(p.Definition.Name)
                
        chosen_param = forms.SelectFromList.show(sorted(list(param_names)), title="Step 1: เลือก Parameter ที่ต้องการใช้กรอง", multiselect=False)
        
        if chosen_param:
            value_dict = {}
            for el in elements:
                p = el.LookupParameter(chosen_param)
                val_str = "< ไม่มีค่า >"
                if p and p.HasValue:
                    val_str = p.AsValueString() or p.AsString() or str(p.AsInteger()) or str(p.AsDouble())
                if val_str not in value_dict: value_dict[val_str] = []
                value_dict[val_str].append(el.Id)
                
            selected_values = forms.SelectFromList.show(sorted(value_dict.keys()), title="Step 2: เลือกค่าของ '{}'".format(chosen_param), multiselect=True)
            
            if selected_values:
                new_selection = List[ElementId]()
                for v in selected_values:
                    for eid in value_dict[v]: 
                        new_selection.Add(eid)
                uidoc.Selection.SetElementIds(new_selection)
                self.Close() 
                return
                
        self.ShowDialog() 

if __name__ == '__main__':
    xaml_path = script.get_bundle_file('ui.xaml')
    window = SuperFilterWindow(xaml_path)
    window.ShowDialog()