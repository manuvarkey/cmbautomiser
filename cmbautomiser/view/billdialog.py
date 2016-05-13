#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# bill_dialog.py
#  
#  Copyright 2014 Manu Varkey <manuvarkey@gmail.com>
#  
#  This program is free software; you can redistribute it and/or modify
#  it under the terms of the GNU General Public License as published by
#  the Free Software Foundation; either version 2 of the License, or
#  (at your option) any later version.
#  
#  This program is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU General Public License for more details.
#  
#  You should have received a copy of the GNU General Public License
#  along with this program; if not, write to the Free Software
#  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston,
#  MA 02110-1301, USA.
#  
#  

import copy, logging

from gi.repository import Gtk, Gdk, GLib

from data.schedule import *
from bill import *
from view.scheduledialog import *
from misc import *

# Setup logger object
log = logging.getLogger(__name__)

class BillDialog:
    
    # Callbacks for GUI elements

    def OnNumberTextChanged(self, entry):
        """Set text in text entry"""
        try:
            val = int(entry.get_text())
            entry.set_text(str(val))
        except ValueError:
            entry.set_text('')

    def onComboChanged(self, combo):
    """For selecting bill item from combobox"""
        tree_iter = combo.get_active_iter()
        if tree_iter is not None or tree_iter != 0:
            model = combo.get_model()
            itemno = model[tree_iter][0]
            self.data.previous_bill = itemno
        else:
            self.data.previous_bill = None

    def onButtonPropertiesPressed(self, button):
        """Create a bill properties dialog window"""
        
        # Variables
        
        toplevel = self.window
        itemnos = []
        item_schedule = self.schedule
        captions = []
        columntypes = []
        populated_items = []
        cellrenderers = []
        
        # Call backs
        
        def callback_agmntno(value, row):
            return populated_items[row][0]

        def callback_description(value, row):
            return populated_items[row][1]

        def callback_unit(value, row):
            return populated_items[row][2]

        def callback_rate(value, row):
            if populated_items[row][3] != 0:
                return str(populated_items[row][3])
            else:
                return ''

        def callback_flag(value, row):
            return str(populated_items[row][5])
        
        # Obtain values to be passed
        
        itemnos = self.data.bills[self.this_bill].item_qty.keys
        # Items specific to normal bill
        if self.data.bill_type == BILL_NORMAL:
            for itemno in itemnos:
                item = self.schedule.[itemno]
                description = item.extended_description_limited
                unit = item.unit
                rate = item.rate
                if self.data.bills[self.this_bill].item_excess_qty[itemno] > 0:
                    flag = 'EXCEEDED'
                else:
                    flag = ''
                populated_items.append([itemno, description, unit, rate, flag])
                captions = ['Agmnt.No.', 'Description', 'Unit', 'Rate', 'Excess Rate', 'P.R.(%)', 'Excess P.R(%)','Excess ?']
                columntypes = [MEAS_CUST, MEAS_CUST, MEAS_CUST, MEAS_CUST, MEAS_L, MEAS_L, MEAS_L, MEAS_CUST]
                cellrenderers = [callback_agmntno] + [callback_description] + [callback_unit] + [callback_rate] + \
                                [None] * 3 + [callback_flag]
        
        # Items specific to custom bill
        elif self.data.bill_type == BILL_CUSTOM:
            for itemno in itemnos:
                item = self.schedule.[itemno]
                description = item.extended_description_limited
                unit = item.unit
                rate = item.rate
                populated_items.append([item_index, itemno, description, unit, rate])
                captions = ['Agmnt.No.', 'Description', 'Unit', 'Rate', 'Total Qty', 'Amount', 'Excess Amount']
                columntypes = [MEAS_CUST, MEAS_CUST, MEAS_CUST, MEAS_CUST, MEAS_L, MEAS_L, MEAS_L]
                cellrenderers = [callback_agmntno] + [callback_description] + [callback_unit] + [callback_rate] + [None] * 3
            
        # Raise Dialog for entry of per item values
        dialog = ScheduleDialog(toplevel, itemnos, captions, columntypes, cellrenderers, item_schedule)

        # Deactivate add and delete buttons, remarks column in dialog
        dialog.builder.get_object("toolbutton_schedule_add").set_sensitive(False)
        dialog.builder.get_object("toolbutton_schedule_add_mult").set_sensitive(False)
        dialog.builder.get_object("toolbutton_schedule_delete").set_sensitive(False)
        dialog.builder.get_object("toolbutton_schedule_copy").set_sensitive(False)
        dialog.builder.get_object("toolbutton_schedule_paste").set_sensitive(False)
        dialog.builder.get_object("filechooserbutton_schedule").set_sensitive(False)
        dialog.builder.get_object("toolbutton_schedule_import").set_sensitive(False)
        dialog.remark_cell.set_sensitive(False)

        # Create data object for current data
        records = []
        for item in populated_items:
            if self.data.bill_type == BILL_NORMAL:
                x = self.data.item_excess_rates[item[0]]
                y = self.data.item_part_percentage[item[0]]
                z = self.data.item_excess_part_percentage[item[0]]
                records.append(['', '', '', '', str(x), str(y), str(z),''])
            elif self.data.bill_type == BILL_CUSTOM:
                x = self.data.item_qty[item[0]][0]
                y = self.data.item_normal_amount[item[0]]
                z = self.data.item_excess_amount[item[0]]
                records.append(['', '', '', '', str(x), str(y), str(z)])
        old_val = [[], records, '', []]

        dialog.set_model(old_val)  # modify dialog with current values of data

        # Run Dialog and get modified values
        data = dialog.run()

        if data is not None:
            records = data[1]
            for count, item in enumerate(populated_items):
                if self.data.bill_type == BILL_NORMAL:
                    self.data.item_excess_rates[item[0]] = float(eval(records[count][4]))
                    self.data.item_part_percentage[item[0]] = float(eval(records[count][5]))
                    self.data.item_excess_part_percentage[item[0]] = float(eval(records[count][6]))
                elif self.data.bill_type == BILL_CUSTOM:
                    self.data.item_qty[item[0]][0] = float(eval(records[count][4]))
                    self.data.item_normal_amount[item[0]] = float(eval(records[count][5]))
                    self.data.item_excess_amount[item[0]] = float(eval(records[count][6]))

    def onToggleCellRendererToggle(self, toggle, path_str):
        """On toggle clicked"""
        path_obj = Gtk.TreePath.new_from_string(path_str)
        path = path_obj.get_indices()
        
        def select_item(item, path, state = None):
            """Change state of item at path if item is toggleable
            
                Arguments:
                    item: Item selected
                    path: Path to item selected
                    state: Final state of item. Toggles on None"""
            if not isinstance(item, data.measurement.MeasurementItemHeading) and self.locked[path] != True:
                if state != None:
                    self.selected[path[0]][path[1]][path[2]] = state
                else:
                    state = self.selected[path[0]][path[1]][path[2]]
                    if state is None:
                        state = False
                    self.selected[path[0]][path[1]][path[2]] = not(state)
            
        if len(path) == 1:
            # Measure all items under CMB
            for p2, meas in enumerate(self.measurements_view.cmbs[path[0]]):
                if not isinstance(meas, Completion):
                    for p3, meas_item in enumerate(meas):
                        select_item(meas_item, path + [p2, p3], True)
        elif len(path) == 2:
            # Measure all items under Measurement
            meas = self.measurements_view.cmbs[path[0]][path[1]]
            if not isinstance(meas, Completion):
                for p3, meas_item in enumerate(meas):
                    select_item(meas_item, path + [p3], True)
        elif len(path) == 3:
            meas_item = self.measurements_view.cmbs[path[0]][path[1]][path[2]]
            select_item(meas_item, path)
            
        self.update_store()

    # General class methods

    def get_model(self):
        self.update_store()
        return self.data

    def clear(self):
        self.data = None

    def update_store(self):
        
        # Update mitems
        self.data.mitems = self.selected.get_paths()
        
        # Update locked states
        self.locked = self.data.get_lock_states() - self.selected
        
        # Update store from lock states
        self.measurements_view.update_store()
        for count_cmb, cmb in enumerate(self.data.cmbs):
            for count_meas, meas in enumerate(cmb):
                if isinstance(meas, Measurement):
                    for count_meas_item, meas_item in enumerate(meas):
                        path = [count_cmb, count_meas, count_meas_item]
                        if self.locked[path]:
                            # Apply color to locked items
                            self.measurements_view.set_colour(path, misc.MEAS_COLOR_LOCKED)
                        elif self.selected[path]:
                            # Apply color to selected items
                            self.measurements_view.set_colour(path, misc.MEAS_COLOR_SELECTED)
                        elif isinstance(meas_item, MeasurementItemHeading):
                            self.measurements_view.set_colour(path, misc.MEAS_COLOR_LOCKED)
                else:
                    path = [count_cmb, count_meas]
                    # Apply color to locked items
                    self.measurements_view.set_colour(path, misc.MEAS_COLOR_LOCKED)
        
        # Update Data elements
        if self.combobox_bill_last_bill.get_active() != 0:
            self.data.prev_bill = self.combobox_bill_last_bill.get_active() - 1
        else:
            self.data.prev_bill = None
        self.data.title = self.entry_bill_title.get_text()
        self.data.cmb_name = self.entry_bill_cmbname.get_text()
        self.data.bill_date = self.entry_bill_bill_date.get_text()
        self.data.starting_page = int(self.entry_bill_starting_page.get_text())

    def __init__(self, parent, data, bill_data, this_bill=None):
        # Setup variables
        self.parent = parent
        self.data = data
        self.billdata = data
        self.this_bill = this_bill  # if in edit mode, use this to filter out this bill entries
        # Derived data
        self.schedule = data.schedule
        self.bills = data.bills
        self.cmbs = data.cmbs
        self.selected = data.datamodel.LockState()
        self.locked = data.get_lock_states()
                
        # Setup dialog window
        self.builder = Gtk.Builder()
        self.builder.add_from_file(abs_path("interface","billdialog.glade"))
        self.window = self.builder.get_object("dialog")
        self.window.set_transient_for(self.parent)
        self.builder.connect_signals(self)
        # Get required objects
        self.treeview_bill = self.builder.get_object("treeview_bill")
        self.tree_bill_text_billed = self.builder.get_object("tree_bill_text_billed")
        self.liststore_previous_bill = self.builder.get_object("liststore_previous_bill")
        self.combobox_bill_last_bill = self.builder.get_object("combobox_bill_last_bill")
        # Text Entries
        self.entry_bill_title = self.builder.get_object("entry_bill_title")
        self.entry_bill_cmbname = self.builder.get_object("entry_bill_cmbname")
        self.entry_bill_bill_date = self.builder.get_object("entry_bill_bill_date")
        self.entry_bill_starting_page = self.builder.get_object("entry_bill_starting_page")

        # Setup cmb tree view
        self.measurements_view = measurement.MeasurementsView(self.schedule, self.data, self.treeview_bill)
        # Setup previous bill combo box list store
        self.liststore_previous_bill.append([0, 'None'])  # Add nil entry
        for row, bill in enumerate(self.bills):
            if row != self.this_bill:
                self.liststore_previous_bill.append([row + 1, bill.get_text()])
            else:
                break  # do not add entries beyond

        # Connect toggled signal of measurement view to callback
        self.tree_bill_text_billed.connect("toggled", self.onToggleCellRendererToggle)
        
        # Load data
        if self.data.prev_bill is not None:
            self.combobox_bill_last_bill.set_active(self.data.prev_bill + 1)
        else:
            self.combobox_bill_last_bill.set_active(0)
        self.entry_bill_title.set_text(self.data.title)
        self.entry_bill_cmbname.set_text(self.data.cmb_name)
        self.entry_bill_bill_date.set_text(self.data.bill_date)
        self.entry_bill_starting_page.set_text(str(self.data.starting_page))
        # Special conditions for custom bill
        if self.data.bill_type == BILL_CUSTOM:
            self.treeview_bill.set_sensitive(False)  # Deactivate measurement item entry
            self.combobox_bill_last_bill.set_sensitive(False)  # Deactivate last bill selection
        
        # Update GUI
        self.update_store()

    def run(self):
        self.window.show_all()
        response = self.window.run()

        if response == 1:
            data = self.get_model()
            self.window.destroy()
            return data
        else:
            self.window.destroy()
            return None
