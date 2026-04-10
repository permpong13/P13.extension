# -*- coding: utf-8 -*-
__title__ = "Numbering\nby Select"
__author__ = "เพิ่มพงษ์ ทวีกุล"
__doc__ = "ReNumber ของ Family ต่างๆ"

# pylint: disable=import-error,invalid-name,broad-except
from collections import OrderedDict
from pyrevit.coreutils import applocales
from pyrevit import revit, DB
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import script

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
output = script.get_output()

# shortcut for DB.BuiltInCategory
BIC = DB.BuiltInCategory

# -----------------------------
# Option Class
# -----------------------------
class RNOpts(object):
    """Renumber tool option"""
    def __init__(self, cat, by_bicat=None):
        self.bicat = cat
        self._cat = revit.query.get_category(self.bicat)
        self.by_bicat = by_bicat
        self._by_cat = revit.query.get_category(self.by_bicat) if self.by_bicat else None
        self.parameter_name = None
        self.digit_count = 1  # Default to 1 digit
        self.prefix = ""  # Default to no prefix

    @property
    def name(self):
        """Renumber option name derived from option categories."""
        try:
            # ตรวจสอบว่า _cat มีค่าไม่เป็น None
            if self._cat is None:
                return "Unknown Category"
                
            if self.by_bicat:
                # ตรวจสอบว่า _by_cat มีค่าไม่เป็น None
                if self._by_cat is None:
                    return "{} by Unknown".format(self._cat.Name)
                    
                applocale = applocales.get_host_applocale()
                if 'english' in applocale.lang_name.lower():
                    return '{} by {}'.format(self._cat.Name, self._by_cat.Name)
                return '{} <- {}'.format(self._cat.Name, self._by_cat.Name)
            return self._cat.Name
        except Exception as e:
            logger.error("Error getting option name: {}".format(str(e)))
            return "Error Category"

# -----------------------------
# Helper Functions
# -----------------------------
def get_open_views():
    """Collect open views in the current document."""
    ui_views = uidoc.GetOpenUIViews()
    views = []
    for ui_View in ui_views:
        viewId = ui_View.ViewId
        view = doc.GetElement(viewId)
        if view.ViewType in (DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.DrawingSheet):
            views.append(view)
    return views

def toggle_element_selection_handles(target_views, bicat, state=True):
    """Toggle handles for spatial elements"""
    with revit.Transaction("Toggle handles"):
        rr_cat = revit.query.get_subcategory(bicat, 'Reference')
        rr_int = revit.query.get_subcategory(bicat, 'Interior Fill')
        if state and bicat != BIC.OST_Viewports:
            for target_view in target_views:
                target_view.EnableTemporaryViewPropertiesMode(target_view.Id)
                try:
                    rr_cat.Visible[target_view] = state
                except Exception as vex:
                    logger.debug(
                        'Failed changing category visibility for \"%s\" to \"%s\" on view \"%s\" | %s',
                        bicat, state, target_view.Name, str(vex)
                    )
                try:
                    rr_int.Visible[target_view] = state
                except Exception as vex:
                    logger.debug(
                        'Failed changing interior fill visibility for \"%s\" to \"%s\" on view \"%s\" | %s',
                        bicat, state, target_view.Name, str(vex)
                    )
        if not state:
            for target_view in target_views:
                target_view.DisableTemporaryViewMode(DB.TemporaryViewMode.TemporaryViewProperties)
                try:
                    rr_int.Visible[target_view] = state
                except Exception as vex:
                    logger.debug(
                        'Failed changing interior fill visibility for \"%s\" to \"%s\" on view \"%s\" | %s',
                        bicat, state, target_view.Name, str(vex)
                    )

class EasilySelectableElements(object):
    """Toggle spatial element handles for easy selection."""
    def __init__(self, target_views, bicat):
        self.supported_categories = [
            BIC.OST_Rooms,
            BIC.OST_Areas,
            BIC.OST_MEPSpaces
        ]
        self.target_views = target_views
        self.bicat = bicat

    def __enter__(self):
        if self.bicat in self.supported_categories:
            toggle_element_selection_handles(self.target_views, self.bicat)
        return self

    def __exit__(self, exception, exception_value, traceback):
        if self.bicat in self.supported_categories:
            toggle_element_selection_handles(self.target_views, self.bicat, state=False)

def increment(number, digit_count=1, prefix=""):
    """Increment given item number by one with leading zeros."""
    try:
        # ถ้ามี prefix ให้ลบ prefix ออกก่อน
        if prefix and number.startswith(prefix):
            number = number[len(prefix):]
        
        # Try to convert to integer and increment
        num_int = int(number)
        num_int += 1
        # Format with leading zeros and add prefix back
        formatted = str(num_int).zfill(digit_count)
        return prefix + formatted
    except (ValueError, TypeError):
        # If not a pure number, use the original increment logic
        incremented = coreutils.increment_str(number, expand=True)
        return prefix + incremented if prefix else incremented

def format_number(number, digit_count=1, prefix=""):
    """Format number with leading zeros and prefix."""
    try:
        # ถ้ามี prefix ให้ลบ prefix ออกก่อน
        if prefix and number.startswith(prefix):
            number = number[len(prefix):]
        
        # Try to convert to integer and format
        num_int = int(number)
        formatted = str(num_int).zfill(digit_count)
        return prefix + formatted
    except (ValueError, TypeError):
        # If not a pure number, return as is with prefix
        return prefix + number if prefix else number

def get_available_parameters(bicat):
    """Get available parameters for the given category that can store text/number values."""
    elements = revit.query.get_elements_by_categories([bicat])
    if not elements:
        return []
    
    # ใช้ element ตัวแรกเพื่อดึง parameters
    sample_element = elements[0]
    parameters = []
    
    for param in sample_element.Parameters:
        if param and param.Definition:
            storage_type = param.StorageType
            # รับเฉพาะ parameters ที่สามารถเก็บค่า string, integer, หรือ double ได้
            if storage_type in [DB.StorageType.String, DB.StorageType.Integer, DB.StorageType.Double]:
                param_name = param.Definition.Name
                # ไม่รวม parameters ที่เป็น system หรือไม่เหมาะสม
                if not any(exclude in param_name.lower() for exclude in ['id', 'guid', 'unique', 'type', 'family']):
                    parameters.append(param_name)
    
    # เอาเฉพาะ unique parameter names และเรียงลำดับ
    return sorted(list(set(parameters)))

def select_parameter_for_category(bicat, category_name):
    """Let user select which parameter to use for numbering."""
    available_params = get_available_parameters(bicat)
    if not available_params:
        forms.alert("No suitable parameters found for {} category.".format(category_name))
        return None
    
    # เพิ่ม default parameters ที่พบบ่อย
    common_params = ["Mark", "Type Mark", "Comments", "Description"]
    
    # จัดเรียงให้ parameters ที่พบบ่อยอยู่ด้านบน
    sorted_params = []
    for param in common_params:
        if param in available_params:
            sorted_params.append(param)
            available_params.remove(param)
    
    sorted_params.extend(available_params)
    
    selected_param = forms.SelectFromList.show(
        sorted_params,
        title="Select Parameter for {}".format(category_name),
        button_name='Use Selected Parameter',
        width=400
    )
    
    return selected_param

def select_digit_count():
    """Let user select number of digits for numbering."""
    digit_options = ["1", "2", "3", "4"]
    
    selected_digit = forms.SelectFromList.show(
        digit_options,
        title="Select Number of Digits",
        button_name='Use Selected Digit Count',
        width=300
    )
    
    if selected_digit:
        return int(selected_digit)
    return 1  # Default to 1 digit

def ask_for_prefix(category_name):
    """Ask user for prefix."""
    prefix = forms.ask_for_string(
        prompt="Enter prefix (leave empty for no prefix)",
        title="Prefix for {}".format(category_name),
        default=""
    )
    return prefix or ""  # Return empty string if None or empty

def get_number(target_element, parameter_name=None):
    """Get target element number from specified parameter."""
    # หากระบุ parameter_name ให้ใช้ parameter นั้น
    if parameter_name:
        param = target_element.LookupParameter(parameter_name)
        if param and param.HasValue:
            storage_type = param.StorageType
            if storage_type == DB.StorageType.String:
                return param.AsString()
            elif storage_type == DB.StorageType.Integer:
                return str(param.AsInteger())
            elif storage_type == DB.StorageType.Double:
                return str(param.AsDouble())
    
    # Fallback ไปใช้ logic เดิม
    if hasattr(target_element, "Number"):
        return target_element.Number
    mark_param = target_element.Parameter[DB.BuiltInParameter.ALL_MODEL_MARK]
    if isinstance(target_element, (DB.Level, DB.Grid)):
        mark_param = target_element.Parameter[DB.BuiltInParameter.DATUM_TEXT]
    if isinstance(target_element, DB.Viewport):
        mark_param = target_element.Parameter[DB.BuiltInParameter.VIEWPORT_DETAIL_NUMBER]
    if mark_param and mark_param.HasValue:
        return mark_param.AsString()
    
    return None

def set_number(target_element, new_number, parameter_name=None):
    """Set target element number to specified parameter."""
    # หากระบุ parameter_name ให้ใช้ parameter นั้น
    if parameter_name:
        param = target_element.LookupParameter(parameter_name)
        if param and not param.IsReadOnly:
            storage_type = param.StorageType
            try:
                if storage_type == DB.StorageType.String:
                    param.Set(new_number)
                    return True
                elif storage_type == DB.StorageType.Integer:
                    param.Set(int(new_number))
                    return True
                elif storage_type == DB.StorageType.Double:
                    param.Set(float(new_number))
                    return True
            except Exception as e:
                logger.error("Error setting parameter {}: {}".format(parameter_name, str(e)))
    
    # Fallback ไปใช้ logic เดิม
    if hasattr(target_element, "Number"):
        target_element.Number = new_number
        return True
    
    mark_param = target_element.Parameter[DB.BuiltInParameter.ALL_MODEL_MARK]
    if isinstance(target_element, (DB.Level, DB.Grid)):
        mark_param = target_element.Parameter[DB.BuiltInParameter.DATUM_TEXT]
    if isinstance(target_element, DB.Viewport):
        mark_param = target_element.Parameter[DB.BuiltInParameter.VIEWPORT_DETAIL_NUMBER]
    if mark_param and not mark_param.IsReadOnly:
        mark_param.Set(new_number)
        return True
    
    return False

def mark_element_as_renumbered(target_view, elem):
    """Override element VG to transparent and halftone."""
    ogs = DB.OverrideGraphicSettings()
    ogs.SetHalftone(True)
    ogs.SetSurfaceTransparency(100)
    target_view.SetElementOverrides(elem.Id, ogs)

def unmark_renamed_elements(target_views, marked_element_ids):
    """Reset element VG to default."""
    for eid in marked_element_ids:
        ogs = DB.OverrideGraphicSettings()
        for view in target_views:
            view.SetElementOverrides(eid, ogs)

def get_elements_dict(views, builtin_cat, parameter_name=None):
    """Collect number:id information about target elements."""
    if BIC.OST_Viewports == builtin_cat:
        for view in views:
            if isinstance(view, DB.ViewSheet):
                return {get_number(doc.GetElement(vpid), parameter_name): vpid for vpid in view.GetAllViewports()}
    all_elements = revit.query.get_elements_by_categories([builtin_cat])
    return {get_number(x, parameter_name): x.Id for x in all_elements if get_number(x, parameter_name) is not None}

def find_replacement_number(existing_number, elements_dict, digit_count=1, prefix=""):
    """Find an appropriate replacement number for conflicting numbers."""
    replaced_number = increment(existing_number, digit_count, prefix)
    while replaced_number in elements_dict:
        replaced_number = increment(replaced_number, digit_count, prefix)
    return replaced_number

def renumber_element(target_element, new_number, elements_dict, parameter_name=None, digit_count=1, prefix=""):
    """Renumber given element."""
    if new_number in elements_dict:
        element_with_same_number = doc.GetElement(elements_dict[new_number])
        if element_with_same_number and element_with_same_number.Id != target_element.Id:
            current_number = get_number(element_with_same_number, parameter_name)
            replaced_number = find_replacement_number(current_number, elements_dict, digit_count, prefix)
            set_number(element_with_same_number, replaced_number, parameter_name)
            elements_dict[replaced_number] = element_with_same_number.Id

    existing_number = get_number(target_element, parameter_name)
    if existing_number in elements_dict:
        elements_dict.pop(existing_number)

    logger.debug('applying %s to parameter %s', new_number, parameter_name)
    success = set_number(target_element, new_number, parameter_name)
    if success:
        elements_dict[new_number] = target_element.Id
        mark_element_as_renumbered(revit.active_view, target_element)
    else:
        logger.error("Failed to set number for element {}".format(target_element.Id))

def ask_for_starting_number(category_name, digit_count=1, prefix=""):
    """Ask user for starting number."""
    default_start = prefix + "1".zfill(digit_count) if prefix else "1".zfill(digit_count)
    starting_number = forms.ask_for_string(
        prompt="Enter starting number (will be formatted to {}{} digits)".format(
            prefix + "-" if prefix else "", digit_count),
        title="ReNumber {}".format(category_name),
        default=default_start
    )
    
    if starting_number:
        # Format the starting number with leading zeros and prefix
        try:
            # ถ้ามี prefix ให้ลบ prefix ออกก่อน
            temp_number = starting_number
            if prefix and temp_number.startswith(prefix):
                temp_number = temp_number[len(prefix):]
            
            # Try to convert to integer and format
            num_int = int(temp_number)
            formatted = str(num_int).zfill(digit_count)
            return prefix + formatted
        except (ValueError, TypeError):
            # If not a pure number, return as is with prefix
            return prefix + starting_number if prefix else starting_number
    return None

def _unmark_collected(category_name, renumbered_element_ids):
    with revit.Transaction("Unmark {}".format(category_name)):
        unmark_renamed_elements(get_open_views(), renumbered_element_ids)

def pick_and_renumber(rnopts, starting_index):
    """Main renumbering routine for elements of given category."""
    open_views = get_open_views() if rnopts.bicat != BIC.OST_Viewports else [revit.active_view]
    with revit.TransactionGroup("Renumber {}".format(rnopts.name)):
        with EasilySelectableElements(open_views, rnopts.bicat):
            index = starting_index
            existing_elements_data = get_elements_dict(open_views, rnopts.bicat, rnopts.parameter_name)
            renumbered_element_ids = []
            for picked_element in revit.get_picked_elements_by_category(
                    rnopts.bicat,
                    message="Select {} in order".format(rnopts.name.lower())):
                with revit.Transaction("Renumber {}".format(rnopts.name)):
                    renumber_element(picked_element, index, existing_elements_data, 
                                   rnopts.parameter_name, rnopts.digit_count, rnopts.prefix)
                    renumbered_element_ids.append(picked_element.Id)
                index = increment(index, rnopts.digit_count, rnopts.prefix)
            _unmark_collected(rnopts.name, renumbered_element_ids)

def door_by_room_renumber(rnopts):
    """Renumber doors based on associated rooms."""
    open_views = get_open_views()
    with revit.TransactionGroup("Renumber Doors by Room"):
        existing_doors_data = get_elements_dict(open_views, rnopts.bicat, rnopts.parameter_name)
        renumbered_door_ids = []
        with EasilySelectableElements(open_views, rnopts.bicat) and EasilySelectableElements(open_views, rnopts.by_bicat):
            while True:
                picked_door = revit.pick_element_by_category(rnopts.bicat, message="Select a door")
                if not picked_door:
                    return _unmark_collected("Doors", renumbered_door_ids)
                from_room, to_room = revit.query.get_door_rooms(picked_door)
                if all([from_room, to_room]) or not any([from_room, to_room]):
                    picked_room = revit.pick_element_by_category(rnopts.by_bicat, message="Select a room")
                    if not picked_room:
                        return _unmark_collected("Rooms", renumbered_door_ids)
                else:
                    picked_room = from_room or to_room
                room_doors = revit.query.get_doors(room_id=picked_room.Id)
                room_number = get_number(picked_room)  # สำหรับ room ยังใช้ logic เดิม
                with revit.Transaction("Renumber Door"):
                    door_count = len(room_doors)
                    if door_count == 1:
                        # ใช้ prefix และ digit_count กับ door numbering
                        formatted_room_number = format_number(room_number, rnopts.digit_count, rnopts.prefix)
                        renumber_element(picked_door, formatted_room_number, existing_doors_data, 
                                       rnopts.parameter_name, rnopts.digit_count, rnopts.prefix)
                        renumbered_door_ids.append(picked_door.Id)
                    elif door_count > 1:
                        room_door_numbers = [get_number(x, rnopts.parameter_name) for x in room_doors]
                        # ใช้ prefix และ digit_count กับ door numbering
                        base_number = format_number(room_number, rnopts.digit_count, rnopts.prefix)
                        new_number = increment(base_number, rnopts.digit_count, rnopts.prefix)
                        while new_number in room_door_numbers:
                            new_number = increment(new_number, rnopts.digit_count, rnopts.prefix)
                        renumber_element(picked_door, new_number, existing_doors_data, 
                                       rnopts.parameter_name, rnopts.digit_count, rnopts.prefix)
                        renumbered_door_ids.append(picked_door.Id)

# -----------------------------
# Main Category Setup
# -----------------------------
if isinstance(revit.active_view, (DB.View3D, DB.ViewPlan, DB.ViewSection, DB.ViewSheet)):
    category_mapping = {
        "Structural Columns": BIC.OST_StructuralColumns,
        "Structural Framing": BIC.OST_StructuralFraming,
        "Structural Foundations": BIC.OST_StructuralFoundation,
        "Walls": BIC.OST_Walls,
        "Floors": BIC.OST_Floors,
        "Roofs": BIC.OST_Roofs,
        "Doors": BIC.OST_Doors,
        "Doors by Room": BIC.OST_Doors,
        "Windows": BIC.OST_Windows,
        "Stairs": BIC.OST_Stairs,
        "Generic Models": BIC.OST_GenericModel,
        "Furniture": BIC.OST_Furniture,
        "Plumbing Fixtures": BIC.OST_PlumbingFixtures,
        "Plumbing Equipment": BIC.OST_PlumbingEquipment,
        "Electrical Fixtures": BIC.OST_ElectricalFixtures,
        "Specialty Equipment": getattr(BIC, "OST_SpecialityEquipment", None),
        "Pipes": BIC.OST_PipeCurves,
        "Pipe Fittings": BIC.OST_PipeFitting,
        "Mechanical Equipment": BIC.OST_MechanicalEquipment,
        "Lighting Fixtures": BIC.OST_LightingFixtures,
        "Electrical Equipment": BIC.OST_ElectricalEquipment,
        "Ceilings": BIC.OST_Ceilings,
        "Railings": BIC.OST_Railings,
        "Site": BIC.OST_Site,
        "Detail Items": getattr(BIC, "OST_DetailItems", None),
        "Rooms": BIC.OST_Rooms,
        "Areas": BIC.OST_Areas,
        "MEP Spaces": BIC.OST_MEPSpaces,
        "Viewports": BIC.OST_Viewports
    }

    # สร้างรายการ category พื้นฐาน
    renumber_options = []
    for name, bicat in category_mapping.items():
        if name == "Doors by Room":
            renumber_options.append(RNOpts(cat=bicat, by_bicat=BIC.OST_Rooms))
        else:
            renumber_options.append(RNOpts(cat=bicat))

    # Filter categories that are None (may not exist in some Revit versions)
    renumber_options = [opt for opt in renumber_options if opt.bicat is not None and opt._cat is not None]

    # หากต้องการกรอง Doors by Room ด้วย (ต้องมีทั้งสอง category)
    renumber_options = [opt for opt in renumber_options 
                       if opt.bicat is not None and opt._cat is not None
                       and (opt.by_bicat is None or (opt.by_bicat is not None and opt._by_cat is not None))]

    # If active view is an AreaPlan, ensure Areas option exists
    if revit.active_view.ViewType == DB.ViewType.AreaPlan:
        if not any(opt.bicat == BIC.OST_Areas for opt in renumber_options):
            areas_opt = RNOpts(cat=BIC.OST_Areas)
            if areas_opt._cat is not None:
                renumber_options.append(areas_opt)

    # Build UI - แสดงเฉพาะชื่อ category เท่านั้น
    options_dict = OrderedDict()
    for renumber_option in renumber_options:
        options_dict[renumber_option.name] = renumber_option

    if options_dict:
        selected_option_name = forms.CommandSwitchWindow.show(
            options_dict,
            message='Pick element type to renumber:',
            width=400
        )

        if selected_option_name:
            selected_option = options_dict[selected_option_name]
            
            # เลือก parameter สำหรับ category นี้
            parameter_name = select_parameter_for_category(selected_option.bicat, selected_option.name)
            if not parameter_name:
                # ผู้ใช้ยกเลิกการเลือก parameter
                script.exit()
            
            selected_option.parameter_name = parameter_name
            
            # เลือกจำนวนหลัก
            digit_count = select_digit_count()
            selected_option.digit_count = digit_count
            
            # ถาม prefix
            prefix = ask_for_prefix(selected_option.name)
            selected_option.prefix = prefix
            
            if selected_option.by_bicat:
                if selected_option.bicat == BIC.OST_Doors and selected_option.by_bicat == BIC.OST_Rooms:
                    with forms.WarningBar(title='Pick Pairs of Door and Room. ESCAPE to end.'):
                        door_by_room_renumber(selected_option)
            else:
                starting_number = ask_for_starting_number(selected_option.name, digit_count, prefix)
                if starting_number:
                    with forms.WarningBar(title='Pick {} One by One. ESCAPE to end.'.format(selected_option.name)):
                        pick_and_renumber(selected_option, starting_number)
    else:
        forms.alert("No valid categories found for renumbering.", exitscript=True)