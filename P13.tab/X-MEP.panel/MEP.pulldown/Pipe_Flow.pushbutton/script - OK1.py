# -*- coding: utf-8 -*-
__title__ = "Pipes Flow\nDirections"
__author__ = "Tee_เพิ่มพงษ์"
__doc__ = "Create flow direction arrows on all pipes in active view"

import clr
import sys
import math

# Revit API References
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')
clr.AddReference('System')

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.UI import *
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# -----------------------------
# Helper Functions
# -----------------------------

def meters_to_feet(m):
    return m * 3.28084

def find_arrow_family(family_name="Flow Arrow"):
    collector = FilteredElementCollector(doc).OfClass(Family)
    for fam in collector:
        if fam.Name.lower() == family_name.lower():
            return fam
    return None

def place_arrow(doc, point, direction, symbol, view):
    try:
        inst = doc.Create.NewFamilyInstance(point, symbol, view)
        angle = math.atan2(direction.Y, direction.X)
        axis = Line.CreateBound(point, point + XYZ.BasisZ)
        ElementTransformUtils.RotateElement(doc, inst.Id, axis, angle)
        return True
    except Exception as e:
        print("⚠️ Failed to place arrow: " + str(e))
        return False

# -----------------------------
# Main Function
# -----------------------------
def create_pipe_flow_arrows():
    try:
        SPACING_M = 5.0
        SPACING_FT = meters_to_feet(SPACING_M)
        ARROW_FAMILY_NAME = "Flow Arrow"

        view = doc.ActiveView
        if view.ViewType != ViewType.FloorPlan:
            TaskDialog.Show("Error", "Please open a Plan View before running this script.")
            return

        pipes = FilteredElementCollector(doc, view.Id).OfClass(Pipe).ToElements()
        if not pipes:
            TaskDialog.Show("Error", "No pipes found in the active view.")
            return

        # === Find and activate family symbol ===
        arrow_family = find_arrow_family(ARROW_FAMILY_NAME)
        if arrow_family is None:
            TaskDialog.Show("Error", "Cannot find family '{}'.".format(ARROW_FAMILY_NAME))
            return

        symbol_ids = list(arrow_family.GetFamilySymbolIds())
        if not symbol_ids:
            TaskDialog.Show("Error", "No symbol types found in family '{}'.".format(ARROW_FAMILY_NAME))
            return

        symbol = doc.GetElement(symbol_ids[0])

        # Activate symbol (separate transaction)
        if not symbol.IsActive:
            t_act = Transaction(doc, "Activate Symbol")
            t_act.Start()
            symbol.Activate()
            t_act.Commit()

        # === Main Transaction for placing arrows ===
        t = Transaction(doc, "Create Pipe Flow Arrows")
        t.Start()

        arrow_count = 0
        processed = 0

        for pipe in pipes:
            processed += 1
            loc = pipe.Location
            if not hasattr(loc, 'Curve'):
                continue

            curve = loc.Curve
            length = curve.Length
            if length < 0.1:
                continue

            # Pipe direction
            try:
                direction = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
            except:
                direction = XYZ(1, 0, 0)

            # Arrow placement
            if length < SPACING_FT:
                distances = [length / 2.0]
            else:
                num_arrows = max(1, int(math.floor(length / SPACING_FT)))
                distances = [(i + 1) * SPACING_FT for i in range(num_arrows)]

            for dist in distances:
                if dist >= length:
                    dist = length - 0.5
                try:
                    param = dist / length
                    point = curve.Evaluate(param, True)
                    if place_arrow(doc, point, direction, symbol, view):
                        arrow_count += 1
                except Exception as e:
                    print("⚠️ Error placing arrow: " + str(e))
                    continue

        t.Commit()  # ✅ Properly closed transaction

        TaskDialog.Show("Flow Arrows Created",
                        "Processed {} pipes\nPlaced {} flow arrows.".format(processed, arrow_count))

    except Exception as e:
        # Ensure rollback if any crash
        try:
            if t.GetStatus() == TransactionStatus.Started:
                t.RollBack()
        except:
            pass
        error_msg = "Error: {}\nLine: {}".format(e, sys.exc_info()[2].tb_lineno)
        TaskDialog.Show("Script Error", error_msg)

# -----------------------------
# Run
# -----------------------------
create_pipe_flow_arrows()
