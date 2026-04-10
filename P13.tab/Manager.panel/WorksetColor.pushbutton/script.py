"""
Workset Commander (Clean Version)
Author: Tee_Permpong
Description: Advanced Workset visualization tool.
             Modes:
             1. Rainbow: Colorize all worksets.
             2. Focus: Highlight one workset.
             (English Only - No Emojis/Special Chars)
"""
# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms, script
import random

# --- SETUP ---
doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView

# --- HELPERS ---
def get_solid_fill_pattern(doc):
    """Finds Solid Fill Pattern."""
    patterns = DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement)
    for p in patterns:
        fp = p.GetFillPattern()
        if fp.IsSolidFill:
            return p.Id
    return None

def generate_color_from_string(name_string):
    """Generates a consistent color based on string hash."""
    seed_val = hash(name_string)
    random.seed(seed_val)
    # Avoid too light colors
    r = random.randint(30, 200)
    g = random.randint(30, 200)
    b = random.randint(30, 200)
    return DB.Color(r, g, b)

def reset_view(view):
    """Clears all overrides."""
    collector = DB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType()
    t = DB.Transaction(doc, "Reset View")
    t.Start()
    try:
        for el in collector:
            view.SetElementOverrides(el.Id, DB.OverrideGraphicSettings())
        t.Commit()
    except:
        t.RollBack()

# --- ANALYZE DATA ---
def analyze_worksets(view):
    collector = DB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType()
    ws_map = {}
    
    for el in collector:
        try:
            ws_id = el.WorksetId
            if ws_id.IntegerValue > 0:
                ws_name = doc.GetWorksetTable().GetWorkset(ws_id).Name
                if ws_name not in ws_map: ws_map[ws_name] = []
                ws_map[ws_name].append(el.Id)
        except: pass
    return ws_map

# --- MODE 1: RAINBOW ---
def mode_rainbow(view, ws_map, exclude_grids=True):
    solid_fill = get_solid_fill_pattern(doc)
    if not solid_fill: return

    t = DB.Transaction(doc, "Rainbow Mode")
    t.Start()
    
    # Reset first
    collector = DB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType()
    for el in collector:
        view.SetElementOverrides(el.Id, DB.OverrideGraphicSettings())

    for ws_name, ids in ws_map.items():
        # Filter Grids
        if exclude_grids and "Shared Levels and Grids" in ws_name:
            continue

        color = generate_color_from_string(ws_name)
        ogs = DB.OverrideGraphicSettings()
        ogs.SetSurfaceForegroundPatternId(solid_fill)
        ogs.SetSurfaceForegroundPatternColor(color)
        ogs.SetProjectionLineColor(color)
        ogs.SetCutForegroundPatternId(solid_fill)
        ogs.SetCutForegroundPatternColor(color)

        for eid in ids:
            try:
                view.SetElementOverrides(eid, ogs)
            except: pass
            
    t.Commit()
    uidoc.RefreshActiveView()
    print("Rainbow Mode Applied.")

# --- MODE 2: FOCUS ---
def mode_focus(view, ws_map, target_ws):
    solid_fill = get_solid_fill_pattern(doc)
    if not solid_fill: return

    # Focus Color (Blue)
    c_focus = DB.Color(0, 120, 255)
    ogs_focus = DB.OverrideGraphicSettings()
    ogs_focus.SetSurfaceForegroundPatternId(solid_fill)
    ogs_focus.SetSurfaceForegroundPatternColor(c_focus)
    ogs_focus.SetProjectionLineColor(c_focus)
    ogs_focus.SetCutForegroundPatternId(solid_fill)
    ogs_focus.SetCutForegroundPatternColor(c_focus)

    # Background (Halftone)
    ogs_bg = DB.OverrideGraphicSettings()
    ogs_bg.SetHalftone(True)
    ogs_bg.SetSurfaceTransparency(70) 

    t = DB.Transaction(doc, "Focus Mode")
    t.Start()

    # Apply Background to ALL
    all_ids = [eid for ids in ws_map.values() for eid in ids]
    for eid in all_ids:
        try:
            view.SetElementOverrides(eid, ogs_bg)
        except: pass

    # Apply Focus to TARGET
    if target_ws in ws_map:
        for eid in ws_map[target_ws]:
            try:
                view.SetElementOverrides(eid, ogs_focus)
            except: pass

    t.Commit()
    uidoc.RefreshActiveView()
    print("Focus Mode: " + target_ws)


# --- MAIN UI ---
if not active_view.AreGraphicsOverridesAllowed():
    forms.alert("View does not support overrides.", exitscript=True)

ws_data = analyze_worksets(active_view)

if not ws_data:
    forms.alert("No worksets found in this view.", exitscript=True)

# Build Menu (English Only)
main_options = [
    "Rainbow Mode (Color All)",
    "Rainbow Mode (Exclude 'Shared Levels/Grids')",
    "Reset View"
]

focus_options = []
sorted_keys = sorted(ws_data.keys())
for ws in sorted_keys:
    count = len(ws_data[ws])
    focus_options.append("Focus: {} ({} items)".format(ws, count))

all_options = main_options + ["--- Focus on Specific Workset ---"] + focus_options

res = forms.SelectFromList.show(
    all_options,
    title="Workset Commander",
    button_name="Execute"
)

if res:
    if "Reset View" in res:
        reset_view(active_view)
        uidoc.RefreshActiveView()
        
    elif "Rainbow Mode" in res:
        exclude = "Exclude" in res
        mode_rainbow(active_view, ws_data, exclude_grids=exclude)
        
    elif "Focus:" in res:
        # Extract name: "Focus: Name (50 items)" -> "Name"
        raw = res.split(": ")[1]
        target = raw.rsplit(" (", 1)[0]
        mode_focus(active_view, ws_data, target)