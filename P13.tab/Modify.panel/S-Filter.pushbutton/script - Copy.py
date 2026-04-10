# -*- coding: utf-8 -*-
from pyrevit import revit, forms, script
import json
import os
import clr

clr.AddReference('PresentationFramework')
from System.Windows.Controls import TreeViewItem, StackPanel, CheckBox, TextBlock, Orientation
from System.Windows import Thickness, VerticalAlignment

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
# 2. คลาสสำหรับควบคุมหน้าต่าง S-Filter
# ==========================================
class SuperFilterWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name):
        forms.WPFWindow.__init__(self, xaml_file_name)
        
        self.selection_ids = uidoc.Selection.GetElementIds()
        
        self.node_children = {}
        self.leaf_checkboxes = [] 
        
        if hasattr(self, 'btn_run'):
            self.btn_run.Click += self.run_filter
            
        if hasattr(self, 'btn_check_all'):
            self.btn_check_all.Click += self.check_all_items
        if hasattr(self, 'btn_uncheck_all'):
            self.btn_uncheck_all.Click += self.uncheck_all_items
        if hasattr(self, 'btn_expand_all'):
            self.btn_expand_all.Click += self.expand_all_items
        if hasattr(self, 'btn_collapse_all'):
            self.btn_collapse_all.Click += self.collapse_all_items
            
        if hasattr(self, 'rad_current_sel'):
            self.rad_current_sel.Checked += self.refresh_tree
        if hasattr(self, 'rad_all_model'):
            self.rad_all_model.Checked += self.refresh_tree
        if hasattr(self, 'rad_visible_view'):
            self.rad_visible_view.Checked += self.refresh_tree
        if hasattr(self, 'rad_belong_view'):
            self.rad_belong_view.Checked += self.refresh_tree
            
        self.populate_categories()

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

    def _expand_nodes(self, items, state):
        for item in items:
            item.IsExpanded = state
            self._expand_nodes(item.Items, state)

    def expand_all_items(self, sender, args):
        self._expand_nodes(self.ui_treeview.Items, True)

    def collapse_all_items(self, sender, args):
        self._expand_nodes(self.ui_treeview.Items, False)
        if hasattr(self, 'root_item'):
            self.root_item.IsExpanded = True

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
        if not hasattr(self, 'ui_treeview'):
            return
            
        self.ui_treeview.Items.Clear()
        self.node_children = {}
        self.leaf_checkboxes = []
        elements = []

        if hasattr(self, 'rad_current_sel') and self.rad_current_sel.IsChecked:
            if self.selection_ids:
                elements = [doc.GetElement(eid) for eid in self.selection_ids]
        elif hasattr(self, 'rad_visible_view') and self.rad_visible_view.IsChecked:
            all_in_view = FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType().ToElements()
            elements = [el for el in all_in_view if el.Category is not None]
        elif hasattr(self, 'rad_belong_view') and self.rad_belong_view.IsChecked:
            all_in_view = FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType().ToElements()
            elements = [el for el in all_in_view if el.Category is not None and el.ViewSpecific]
        elif hasattr(self, 'rad_all_model') and self.rad_all_model.IsChecked:
            all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
            elements = [el for el in all_elements if el.Category is not None]

        if not elements:
            if hasattr(self, 'lbl_status_count'):
                self.lbl_status_count.Text = "0"
            return
            
        cat_dict = {}
        for el in elements:
            cat_name = el.Category.Name if el.Category else "Uncategorized"
            
            if isinstance(el, FamilyInstance) and el.Symbol:
                type_name = "{} : {}".format(el.Symbol.Family.Name, el.Name)
            else:
                try:
                    type_name = el.Name
                except:
                    type_name = type(el).__name__
                    
            if not type_name:
                type_name = "Unknown Type"
                
            if cat_name not in cat_dict:
                cat_dict[cat_name] = {}
            if type_name not in cat_dict[cat_name]:
                cat_dict[cat_name][type_name] = []
                
            cat_dict[cat_name][type_name].append(el.Id)
            
        if hasattr(self, 'lbl_status_count'):
            self.lbl_status_count.Text = str(len(elements))

        self.root_item, self.root_chk = self.create_tree_node("All", len(elements))
        self.ui_treeview.Items.Add(self.root_item)
        self.node_children[self.root_chk] = []
        
        for cat_name in sorted(cat_dict.keys()):
            cat_eids = []
            for t_list in cat_dict[cat_name].values():
                cat_eids.extend(t_list)
                
            cat_item, cat_chk = self.create_tree_node(cat_name, len(cat_eids))
            self.node_children[self.root_chk].append(cat_chk)
            self.node_children[cat_chk] = []
            self.root_item.Items.Add(cat_item)
            
            for type_name in sorted(cat_dict[cat_name].keys()):
                type_eids = cat_dict[cat_name][type_name]
                type_item, type_chk = self.create_tree_node(type_name, len(type_eids))
                
                self.node_children[cat_chk].append(type_chk)
                self.node_children[type_chk] = [] 
                self.leaf_checkboxes.append((type_chk, type_eids)) 
                cat_item.Items.Add(type_item)
                
        self.root_item.IsExpanded = True

    def run_filter(self, sender, args):
        new_selection = List[ElementId]()
        
        for chk, eids in self.leaf_checkboxes:
            if chk.IsChecked:
                for eid in eids:
                    new_selection.Add(eid)
                    
        uidoc.Selection.SetElementIds(new_selection)
        
        # ✨ ลบ Pop-up แจ้งเตือนออกแล้ว สคริปต์จะแค่ปิดหน้าต่างและส่ง selection เข้า Revit ทันที
        self.Close()

if __name__ == '__main__':
    xaml_path = script.get_bundle_file('ui.xaml')
    window = SuperFilterWindow(xaml_path)
    window.ShowDialog()