# -*- coding: utf-8 -*-
"""Pipe Accessory Placement Tool - Material Design UI"""
import clr
import sys
import System
import math
import os

clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import Pipe

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType, ObjectSnapTypes
from Autodesk.Revit.UI import *

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
import System.Windows.Forms
from System.Windows.Forms import *
import System.Drawing
from System.Drawing import *

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager

doc = DocumentManager.Instance.CurrentDBDocument
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

# ─── Selection Filter ─────────────────────────────────────────────────────────
class PipeSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, Pipe)
    def AllowReference(self, ref, pos):
        return False

# ─── UI Form ──────────────────────────────────────────────────────────────────
class MainForm(Form):
    def __init__(self, accs_list, last_symbol_id=None, last_dist_str="0"):
        self.all_accs = accs_list
        self.filtered_accs = list(accs_list)
        self.last_symbol_id = last_symbol_id
        self.last_dist_str = last_dist_str
        self.user_cancelled = True
        self.selected_symbolId = None
        self.distance_mm = 0.0
        self.InitUI()

    def InitUI(self):
        c_primary   = Color.FromArgb(98, 0, 238)
        c_success   = Color.FromArgb(76, 175, 80)
        c_bg        = Color.FromArgb(245, 245, 246)
        c_surface   = Color.White
        c_dark      = Color.FromArgb(33, 33, 33)
        c_cancel_bg = Color.FromArgb(224, 224, 224)

        self.Text = "Place Pipe Accessories"
        self.ClientSize = Size(540, 390)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.BackColor = c_bg
        self.Font = Font("Segoe UI", 10)

        # ── Card Panel ──
        card = Panel()
        card.BackColor = c_surface
        card.Location = Point(20, 20)
        card.Size = Size(500, 290)
        self.Controls.Add(card)

        # ── Title ──
        lbl_title = Label()
        lbl_title.Text = "Accessory Settings"
        lbl_title.Font = Font("Segoe UI", 12, FontStyle.Bold)
        lbl_title.ForeColor = c_primary
        lbl_title.Location = Point(20, 15)
        lbl_title.AutoSize = True
        card.Controls.Add(lbl_title)

        x_lbl = 20; w_lbl = 160
        x_ctrl = 190; w_ctrl = 290
        row1 = 60; row2 = 120; row3 = 185; h_lbl = 34; h_ctrl = 28

        def make_label(text, y):
            lbl = Label()
            lbl.Text = text
            lbl.Font = Font("Segoe UI", 9, FontStyle.Bold)
            lbl.BackColor = c_primary
            lbl.ForeColor = Color.White
            lbl.Location = Point(x_lbl, y)
            lbl.Size = Size(w_lbl, h_lbl)
            lbl.TextAlign = ContentAlignment.MiddleCenter
            return lbl

        card.Controls.Add(make_label("SEARCH FILTER", row1))
        card.Controls.Add(make_label("ACCESSORY", row2))
        card.Controls.Add(make_label("DISTANCE (mm)", row3))

        # Filter textbox
        self.txb_filter = TextBox()
        self.txb_filter.Location = Point(x_ctrl, row1 + 3)
        self.txb_filter.Size = Size(w_ctrl, h_ctrl)
        self.txb_filter.BorderStyle = BorderStyle.FixedSingle
        self.txb_filter.TextChanged += self.on_filter_changed
        card.Controls.Add(self.txb_filter)

        # Accessory combo
        self.cb_acc = ComboBox()
        self.cb_acc.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb_acc.Location = Point(x_ctrl, row2 + 3)
        self.cb_acc.Size = Size(w_ctrl, h_ctrl)
        self.cb_acc.DropDownWidth = 620
        self.cb_acc.BackColor = Color.White
        card.Controls.Add(self.cb_acc)
        self.refresh_combo(self.all_accs)

        # Distance textbox
        self.txb_dist = TextBox()
        self.txb_dist.Location = Point(x_ctrl, row3 + 3)
        self.txb_dist.Size = Size(w_ctrl, h_ctrl)
        self.txb_dist.BorderStyle = BorderStyle.FixedSingle
        self.txb_dist.Text = self.last_dist_str
        card.Controls.Add(self.txb_dist)

        # ── Buttons ──
        btn_run = Button()
        btn_run.Text = "RUN PLACEMENT"
        btn_run.Font = Font("Segoe UI", 10, FontStyle.Bold)
        btn_run.BackColor = c_success
        btn_run.ForeColor = Color.White
        btn_run.FlatStyle = FlatStyle.Flat
        btn_run.FlatAppearance.BorderSize = 0
        btn_run.Location = Point(230, 325)
        btn_run.Size = Size(160, 42)
        btn_run.Click += self.on_run
        self.Controls.Add(btn_run)

        btn_cancel = Button()
        btn_cancel.Text = "CANCEL"
        btn_cancel.Font = Font("Segoe UI", 10, FontStyle.Bold)
        btn_cancel.BackColor = c_cancel_bg
        btn_cancel.ForeColor = c_dark
        btn_cancel.FlatStyle = FlatStyle.Flat
        btn_cancel.FlatAppearance.BorderSize = 0
        btn_cancel.Location = Point(400, 325)
        btn_cancel.Size = Size(110, 42)
        btn_cancel.Click += self.on_cancel
        self.Controls.Add(btn_cancel)

        self.TopMost = True

    def refresh_combo(self, accs):
        self.cb_acc.Items.Clear()
        for a in accs:
            self.cb_acc.Items.Add(a["name"])
        if self.cb_acc.Items.Count > 0:
            self.cb_acc.SelectedIndex = 0
            if self.last_symbol_id:
                for i, a in enumerate(accs):
                    if a["symbol"].Id == self.last_symbol_id:
                        self.cb_acc.SelectedIndex = i
                        break

    def on_filter_changed(self, sender, e):
        q = self.txb_filter.Text.lower()
        self.filtered_accs = [a for a in self.all_accs if q in a["name"].lower()]
        self.refresh_combo(self.filtered_accs)

    def on_run(self, sender, e):
        if self.cb_acc.SelectedIndex < 0:
            MessageBox.Show("Please select an accessory.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        try:
            self.distance_mm = float(self.txb_dist.Text.strip())
        except:
            MessageBox.Show("Invalid distance. Please enter a number.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        sel = self.filtered_accs[self.cb_acc.SelectedIndex]
        self.selected_symbolId = sel["symbol"].Id
        self.user_cancelled = False
        self.Close()

    def on_cancel(self, sender, e):
        self.user_cancelled = True
        self.Close()

# ─── Helpers ──────────────────────────────────────────────────────────────────
def get_all_accessories():
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(FamilySymbol)
    lst = []
    for sym in collector:
        fam_name = "?"
        try:
            p = sym.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
            fam_name = p.AsString() if p and p.AsString() else sym.FamilyName
        except: pass

        dia = ""
        for pname in [BuiltInParameter.RBS_PIPE_DIAMETER_PARAM]:
            try:
                p = sym.get_Parameter(pname)
                if p: dia = p.AsValueString() or ""; break
            except: pass
        if not dia:
            for lname in ["Nominal Diameter", "Size"]:
                try:
                    p = sym.LookupParameter(lname)
                    if p: dia = p.AsValueString() or ""; break
                except: pass
        if not dia:
            try:
                p = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if p: dia = p.AsString() or "?"
            except: pass
        if not dia:
            dia = "?"

        lst.append({"symbol": sym, "name": "%s_%s" % (fam_name, dia)})

    lst.sort(key=lambda x: x["name"])
    return lst

# ─── Main ─────────────────────────────────────────────────────────────────────
def main():
    accs = get_all_accessories()
    if not accs:
        MessageBox.Show("No pipe accessories found in this project.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        return

    last_sym_id = None
    last_dist   = "0"

    while True:
        form = MainForm(accs, last_sym_id, last_dist)
        Application.Run(form)

        if form.user_cancelled or not form.selected_symbolId:
            break

        last_sym_id = form.selected_symbolId
        last_dist   = form.txb_dist.Text

        symbol_id = form.selected_symbolId
        dist_ft   = form.distance_mm / 304.8

        try:
            pipe_ref = uidoc.Selection.PickObject(
                ObjectType.Element, PipeSelectionFilter(),
                "SELECT A PIPE  (ESC = back to settings)")
            pipe = doc.GetElement(pipe_ref.ElementId)

            origin_pt = uidoc.Selection.PickPoint(
                ObjectSnapTypes.Endpoints,
                "PICK ORIGIN POINT")

            # Find nearest connector → c1 (origin side), c2 (far side)
            conn_list = [c for c in pipe.ConnectorManager.Connectors]
            if len(conn_list) < 2:
                MessageBox.Show("Pipe has fewer than 2 connectors.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                continue

            conn_list.sort(key=lambda c: c.Origin.DistanceTo(origin_pt))
            c1 = conn_list[0]
            c2 = conn_list[1]

            direction   = (c2.Origin - c1.Origin).Normalize()
            insert_pt   = c1.Origin + direction * dist_ft
            pipe_length = c1.Origin.DistanceTo(c2.Origin)

            if dist_ft > pipe_length:
                MessageBox.Show("Distance exceeds pipe length!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                continue

            # ── Transaction ──────────────────────────────────────────────────
            t = Transaction(doc, "Auto Place Pipe Accessory")
            t.Start()
            try:
                symbol = doc.GetElement(symbol_id)
                if not symbol.IsActive:
                    symbol.Activate()
                    doc.Regenerate()

                # Place family at insert_pt (floating, not hosted)
                inst = doc.Create.NewFamilyInstance(
                    insert_pt, symbol,
                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural)
                doc.Regenerate()

                # Get accessory connectors
                acc_conns = [c for c in inst.MEPModel.ConnectorManager.Connectors
                             if c.ConnectorType == ConnectorType.End
                             or c.ConnectorType == ConnectorType.Curve]

                if len(acc_conns) >= 2:
                    D = acc_conns[0].Origin.DistanceTo(acc_conns[1].Origin)
                    p1_end   = insert_pt
                    p2_start = insert_pt + direction * D

                    if pipe_length >= dist_ft + D:
                        # ── Rotate (CrossProduct method) ──────────────────
                        try:
                            acc_vec = (acc_conns[0].Origin - acc_conns[1].Origin).Normalize()
                            angle   = direction.AngleTo(acc_vec)
                            rot_ctr = (acc_conns[0].Origin + acc_conns[1].Origin) * 0.5

                            if 0.001 < angle < math.pi - 0.001:
                                ax_dir = acc_vec.CrossProduct(direction).Normalize()
                                ax = Line.CreateBound(rot_ctr, rot_ctr + ax_dir)
                                ElementTransformUtils.RotateElement(doc, inst.Id, ax, angle)
                            elif angle >= math.pi - 0.001:
                                ax_dir = acc_vec.CrossProduct(XYZ.BasisZ)
                                if ax_dir.IsZeroLength():
                                    ax_dir = acc_vec.CrossProduct(XYZ.BasisY)
                                ax = Line.CreateBound(rot_ctr, rot_ctr + ax_dir.Normalize())
                                ElementTransformUtils.RotateElement(doc, inst.Id, ax, math.pi)
                        except: pass

                        doc.Regenerate()

                        # ── Move to gap center ────────────────────────────
                        ac2 = [c for c in inst.MEPModel.ConnectorManager.Connectors
                               if c.ConnectorType == ConnectorType.End
                               or c.ConnectorType == ConnectorType.Curve]
                        if len(ac2) >= 2:
                            ctr_acc = (ac2[0].Origin + ac2[1].Origin) * 0.5
                            ctr_gap = (p1_end + p2_start) * 0.5
                            ElementTransformUtils.MoveElement(doc, inst.Id, ctr_gap - ctr_acc)

                        doc.Regenerate()

                        # ── BreakCurve + trim downstream pipe ─────────────
                        new_pipe = None
                        try:
                            import Autodesk.Revit.DB.Plumbing as Plumbing
                            new_pid  = Plumbing.PlumbingUtils.BreakCurve(doc, pipe.Id, p1_end)
                            new_pipe = doc.GetElement(new_pid)

                            cv1 = pipe.Location.Curve
                            is_ds = (cv1.GetEndPoint(0).IsAlmostEqualTo(c2.Origin) or
                                     cv1.GetEndPoint(1).IsAlmostEqualTo(c2.Origin))

                            tgt  = pipe if is_ds else new_pipe
                            oc   = tgt.Location.Curve
                            ep0  = oc.GetEndPoint(0)
                            ep1  = oc.GetEndPoint(1)

                            # Preserve original draw direction (avoids "Opposite Direction" error)
                            if ep0.DistanceTo(p1_end) < ep1.DistanceTo(p1_end):
                                tgt.Location.Curve = Line.CreateBound(p2_start, ep1)
                            else:
                                tgt.Location.Curve = Line.CreateBound(ep0, p2_start)
                        except: pass

                        doc.Regenerate()

                        # ── Connect by proximity ──────────────────────────
                        try:
                            for ca in inst.MEPModel.ConnectorManager.Connectors:
                                for cp in pipe.ConnectorManager.Connectors:
                                    if ca.Origin.IsAlmostEqualTo(cp.Origin):
                                        try: ca.ConnectTo(cp)
                                        except: pass
                                if new_pipe:
                                    for cn in new_pipe.ConnectorManager.Connectors:
                                        if ca.Origin.IsAlmostEqualTo(cn.Origin):
                                            try: ca.ConnectTo(cn)
                                            except: pass
                        except: pass

                    else:
                        MessageBox.Show("Not enough pipe length to insert this accessory.", "Warning",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning)

                t.Commit()

            except Exception as ex:
                try: t.RollBack()
                except: pass
                MessageBox.Show("Placement error: " + str(ex), "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)

        except Autodesk.Revit.Exceptions.OperationCanceledException:
            pass  # ESC → loop back to UI
        except Exception as e:
            MessageBox.Show("Unexpected error: " + str(e), "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

main()
