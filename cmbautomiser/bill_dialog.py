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

import copy

from gi.repository import Gtk, Gdk, GLib

from globalconstants import *
from schedule import *
from bill import *
from schedule_dialog import *
from misc import *


class BillDialog:
    # General signal handler Methods

    def OnNumberTextChanged(self, entry):
        try:
            val = int(entry.get_text())
            entry.set_text(unicode(val))
        except ValueError:
            entry.set_text('')

    def onComboChanged(self, combo):
    # for selecting item bill

        tree_iter = combo.get_active_iter()
        if tree_iter is not None or tree_iter != 0:
            model = combo.get_model()
            itemno = model[tree_iter][0]
            self.data.previous_bill = itemno
        else:
            self.data.previous_bill = None

    def onButtonPropertiesPressed(self, button):
        # Create schedule window

        # variables
        toplevel = self.window
        itemnos = []
        item_schedule = self.schedule
        captions = []
        columntypes = []
        populated_items = []
        cellrenderers = []
        # call backs
        def callback_agmntno(value, row):
            return populated_items[row][1]

        def callback_description(value, row):
            return populated_items[row][2]

        def callback_unit(value, row):
            return populated_items[row][3]

        def callback_rate(value, row):
            if populated_items[row][4] != 0:
                return unicode(populated_items[row][4])
            else:
                return ''

        def callback_flag(value, row):
            return unicode(populated_items[row][5])

        if self.data.bill_type == BILL_NORMAL:
            self.evaluate_qtys()
            for item_index in range(self.schedule.length()):
                item = self.schedule.get_item_by_index(item_index)
                itemno = item.itemno
                description = item.description
                unit = item.unit
                rate = item.rate
                if self.item_excess_qty[item_index] > 0:
                    flag = 'EXCEEDED'
                else:
                    flag = ''
                populated_items.append([item_index, itemno, description, unit, rate, flag])
            captions = ['Agmnt.No.', 'Description', 'Unit', 'Rate', 'Excess Rate', 'P.R.(%)', 'Excess P.R(%)','Excess ?']
            columntypes = [MEAS_CUST, MEAS_CUST, MEAS_CUST, MEAS_CUST, MEAS_L, MEAS_L, MEAS_L,MEAS_CUST]
            cellrenderers = [callback_agmntno] + [callback_description] + [callback_unit] + [callback_rate] + \
                            [None] * 3 + [callback_flag]
        elif self.data.bill_type == BILL_CUSTOM:
            for item_index in range(self.schedule.length()):
                item = self.schedule.get_item_by_index(item_index)
                itemno = item.itemno
                description = item.description
                unit = item.unit
                rate = item.rate
                populated_items.append([item_index, itemno, description, unit, rate])
            captions = ['Agmnt.No.', 'Description', 'Unit', 'Rate', 'Total Qty', 'Amount', 'Excess Amount']
            columntypes = [MEAS_CUST, MEAS_CUST, MEAS_CUST, MEAS_CUST, MEAS_L, MEAS_L, MEAS_L]
            cellrenderers = [callback_agmntno] + [callback_description] + [callback_unit] + [callback_rate] + [None] * 3

        dialog = ScheduleDialog(toplevel, itemnos, captions, columntypes, cellrenderers, item_schedule)

        # deactivate add and delete buttons, remarks column in dialog
        dialog.builder.get_object("toolbutton_schedule_add").set_sensitive(False)
        dialog.builder.get_object("toolbutton_schedule_add_mult").set_sensitive(False)
        dialog.builder.get_object("toolbutton_schedule_delete").set_sensitive(False)
        dialog.builder.get_object("toolbutton_schedule_copy").set_sensitive(False)
        dialog.builder.get_object("toolbutton_schedule_paste").set_sensitive(False)
        dialog.builder.get_object("filechooserbutton_schedule").set_sensitive(False)
        dialog.builder.get_object("toolbutton_schedule_import").set_sensitive(False)
        dialog.remark_cell.set_sensitive(False)

        # create data object for current data
        records = []
        for item in populated_items:
            if self.data.bill_type == BILL_NORMAL:
                x = self.data.item_excess_rates[item[0]]
                y = self.data.item_part_percentage[item[0]]
                z = self.data.item_excess_part_percentage[item[0]]
                records.append(['', '', '', '', unicode(x), unicode(y), unicode(z),''])
            elif self.data.bill_type == BILL_CUSTOM:
                x = self.data.item_qty[item[0]][0]
                y = self.data.item_normal_amount[item[0]]
                z = self.data.item_excess_amount[item[0]]
                records.append(['', '', '', '', unicode(x), unicode(y), unicode(z)])
        old_val = [[], [], records, '', []]

        dialog.set_model(old_val)  # modify dialog with current values of data

        data = dialog.run()  # get modified values of data

        if data is not None:
            records = data[2]
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
        path_obj = Gtk.TreePath.new_from_string(path_str)
        path = path_obj.get_indices()
        if len(path) == 1:
            for meas in self.measurements_view.cmbs[path[0]]:
                if not isinstance(meas, Completion):
                    for meas_item in meas:
                        if not isinstance(meas_item, MeasurementItemHeading) and not (path in self.locked):
                            meas_item.set_billed_flag(True)
        elif len(path) == 2:
            meas = self.measurements_view.cmbs[path[0]][path[1]]
            if not isinstance(meas, Completion):
                for meas_item in meas:
                    if not isinstance(meas_item, MeasurementItemHeading) and not (path in self.locked):
                        meas_item.set_billed_flag(True)
        elif len(path) == 3:
            meas_item = self.measurements_view.cmbs[path[0]][path[1]][path[2]]
            if not isinstance(meas_item, MeasurementItemHeading) and not (path in self.locked):
                meas_item.set_billed_flag(not (meas_item.get_billed_flag()))
        self.update_store()

    # Class methods

    def evaluate_qtys(self):
        self.update_store()
        # Update schedule item vars
        if self.data.prev_bill is not None:
            self.prev_bill = self.bill_view.bills[self.data.prev_bill]  # get prev_bill object
        else:
            self.prev_bill = None

        sch_len = self.schedule.length()
        # initialise variables to schedule length
        self.item_cmb_ref = []
        self.item_qty = []
        for i in range(sch_len):  # make a list of empty lists
            self.item_cmb_ref.append([])
            self.item_qty.append([])
        self.item_normal_qty = [0] * sch_len
        self.item_excess_qty = [0] * sch_len

        # fill in from measurement items
        for mitem in self.data.mitems:
            item = self.cmbs[mitem[0]][mitem[1]][mitem[2]]
            if not isinstance(item, MeasurementItemHeading):
                for itemno, item_qty in zip(item.itemnos, item.get_total()):
                    item_index = self.schedule.get_item_index(itemno)
                    if item_index is not None:
                        self.item_cmb_ref[item_index].append(mitem[0])
                        self.item_qty[item_index].append(item_qty)

        # fill in from Prev Bill
        if self.prev_bill is not None:
            for item_index, item_qty in enumerate(self.item_qty):
                self.item_cmb_ref[item_index].append(-1)  # use -1 as marker for prev abstract
                self.item_qty[item_index].append(sum(self.prev_bill.item_qty[item_index]))  # add total qty from previous bill

        # Evaluate remaining variables from above data
        for item_index in range(sch_len):
            item = self.schedule[item_index]
            # determine total qty
            total_qty = sum(self.item_qty[item_index])
            # determine items above and at normal rates
            if total_qty > item.qty * (1 + 0.01 * item.excess_rate_percent):
                if item.unit.lower() in INT_ITEMS:
                    self.item_normal_qty[item_index] = math.floor(item.qty * (1 + 0.01 * item.excess_rate_percent))
                else:
                    self.item_normal_qty[item_index] = round(item.qty * (1 + 0.01 * item.excess_rate_percent), 2)
                self.item_excess_qty[item_index] = total_qty - self.item_normal_qty[item_index]
            else:
                self.item_normal_qty[item_index] = total_qty

    def get_model(self):
        self.update_store()
        return self.data

    def clear(self):
        self.data = None

    def update_store(self):
        # Update measurements store
        self.measurements_view.update_store()

        # Update Data elements
        if self.combobox_bill_last_bill.get_active() != 0:
            self.data.prev_bill = self.combobox_bill_last_bill.get_active() - 1
        else:
            self.data.prev_bill = None
        self.data.title = self.entry_bill_title.get_text()
        self.data.cmb_name = self.entry_bill_cmbname.get_text()
        self.data.bill_date = self.entry_bill_bill_date.get_text()
        self.data.starting_page = int(self.entry_bill_starting_page.get_text())

        # update data.mitems from self.cmbs
        self.data.mitems = []  # clear mitems
        for count_cmb, cmb in enumerate(self.cmbs):
            for count_meas, meas in enumerate(cmb):
                if isinstance(meas, Measurement):
                    for count_meas_item, meas_item in enumerate(meas):
                        path = Gtk.TreePath.new_from_string(
                            unicode(count_cmb) + ':' + unicode(count_meas) + ':' + unicode(count_meas_item))
                        path_iter = self.measurements_view.store.get_iter(path)
                        if [count_cmb, count_meas, count_meas_item] in self.locked:
                            self.measurements_view.store.set_value(path_iter, 3,
                                                                   MEAS_COLOR_LOCKED)  # apply color to locked items
                        elif meas_item.get_billed_flag() is True:
                            self.data.mitems.append([count_cmb, count_meas, count_meas_item])
                            self.measurements_view.store.set_value(path_iter, 3,
                                                                   MEAS_COLOR_SELECTED)  # apply color to selected items
                        elif isinstance(meas_item, MeasurementItemHeading):
                            self.measurements_view.store.set_value(path_iter, 3, MEAS_COLOR_LOCKED)
                else:
                    path = Gtk.TreePath.new_from_string(unicode(count_cmb) + ':' + unicode(count_meas))
                    path_iter = self.measurements_view.store.get_iter(path)
                    self.measurements_view.store.set_value(path_iter, 3, MEAS_COLOR_LOCKED)  # apply color to locked

    def __init__(self, parent, data, bill_view, this_bill=None):
        # Setup variables

        self.parent = parent
        self.data = data
        self.bill_view = bill_view
        self.schedule = bill_view.schedule
        self.this_bill = this_bill  # if in edit mode, use this to filter out this bill entries
        self.bills = bill_view.bills  # reference to bill
        self.cmbs = copy.deepcopy(self.bill_view.cmbs)  # make copy of cmbs for tinkering
        self.locked = []  # paths to items loacked from changing
        # expand variables to schedule length
        sch_len = self.schedule.length()
        existing_length = len(self.data.item_excess_rates)
        if sch_len > existing_length:
            self.data.item_part_percentage += [100] * (sch_len - existing_length)
            self.data.item_excess_part_percentage += [100] * (sch_len - existing_length)
            self.data.item_excess_rates += [0] * (sch_len - existing_length)
            for i in range(sch_len - existing_length):
                self.data.item_qty.append([0])
            self.data.item_normal_amount += [0] * (sch_len - existing_length)
            self.data.item_excess_amount += [0] * (sch_len - existing_length)
        # Setup dialog window
        self.builder = Gtk.Builder()
        self.builder.add_from_file(abs_path("interface","billdialog.glade"))
        self.window = self.builder.get_object("dialog")
        self.window.set_transient_for(self.parent)
        self.builder.connect_signals(self)
        # Get required objects
        self.treeview_bill = self.builder.get_object("treeview_bill")
        self.liststore_bill = self.builder.get_object("liststore_bill")
        self.tree_bill_text_billed = self.builder.get_object("tree_bill_text_billed")
        self.liststore_previous_bill = self.builder.get_object("liststore_previous_bill")
        self.combobox_bill_last_bill = self.builder.get_object("combobox_bill_last_bill")
        # Text Entries
        self.entry_bill_title = self.builder.get_object("entry_bill_title")
        self.entry_bill_cmbname = self.builder.get_object("entry_bill_cmbname")
        self.entry_bill_bill_date = self.builder.get_object("entry_bill_bill_date")
        self.entry_bill_starting_page = self.builder.get_object("entry_bill_starting_page")

        # setup cmb tree view
        self.measurements_view = MeasurementsView(self.schedule, self.liststore_bill, self.treeview_bill)
        self.measurements_view.set_data_object(self.schedule, self.cmbs)
        # setup previous bill combo box list store
        self.liststore_previous_bill.append([0, 'None'])  # Add nil entry
        for row, bill in enumerate(self.bills):
            if row != self.this_bill:
                self.liststore_previous_bill.append([row + 1, bill.get_text()])
            else:
                break  # do not add entries beyond
                # self.liststore_previous_bill.append([0,'None']) # Add nil entry
        self.combobox_bill_last_bill.set_active(0)

        # Connect signals
        self.tree_bill_text_billed.connect("toggled", self.onToggleCellRendererToggle)
        # Restore UI elements from data
        # Measured Items
        # lock all measured items
        for count_cmb, cmb in enumerate(self.cmbs):
            for count_meas, meas in enumerate(cmb):
                if isinstance(meas, Measurement):
                    for count_meas_item, meas_item in enumerate(meas):
                        if meas_item.get_billed_flag() is True:
                            self.locked.append([count_cmb, count_meas, count_meas_item])
        # unlock any measured items
        for path in self.data.mitems:
            try:
                self.locked.remove([path[0], path[1], path[2]])
            except ValueError:
                print('Unable to release lock on item :' + unicode(path))

        if self.data.prev_bill is not None:
            self.combobox_bill_last_bill.set_active(self.data.prev_bill + 1)
        else:
            self.combobox_bill_last_bill.set_active(0)
        self.entry_bill_title.set_text(self.data.title)
        self.entry_bill_cmbname.set_text(self.data.cmb_name)
        self.entry_bill_bill_date.set_text(self.data.bill_date)
        self.entry_bill_starting_page.set_text(unicode(self.data.starting_page))
        # special conditions for custom bill
        if self.data.bill_type == BILL_CUSTOM:
            self.treeview_bill.set_sensitive(False)  # deactivate measurement item entry
            self.combobox_bill_last_bill.set_sensitive(False)  # deactivate last bill selection
        else:
            self.evaluate_qtys()
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
