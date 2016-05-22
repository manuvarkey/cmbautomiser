#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# bill.py
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

import codecs, pickle, copy, logging
from gi.repository import Gtk, Gdk, GLib

# local files import
from __main__ import data, misc
from . import measurement, scheduledialog

# Setup logger object
log = logging.getLogger(__name__)


class BillView:
    """Implements a view for display and manipulation of bill items over a treeview"""
    
    # Callback functions

    def onKeyPressTreeviewSchedule(self, treeview, event):
        """Unselect view on escape"""
        keyname = Gdk.keyval_name(event.keyval)
        if keyname == "Escape":
            self.tree.get_selection().unselect_all()

    # Bill view methods

    def add_bill(self):
        """Add bill at end"""
        model = data.bill.BillData().get_model()
        bill_dialog = BillDialog(self.parent, self.data, model)
        data_model = bill_dialog.run()
        if data_model is not None:  # if cancel not pressed
            self.data.insert_bill_at_row(data_model, None)
        self.update_store()

    def add_bill_custom(self):
        """Add custom bill at end"""
        model = data.bill.BillData(misc.BILL_CUSTOM).get_model()
        bill_dialog = BillDialog(self.parent, self.data, model)
        data_model = bill_dialog.run()
        if data_model is not None:  # if cancel not pressed
            self.data.insert_bill_at_row(data_model, None)
        self.update_store()

    def insert_bill_at_selection(self, item):
        """Insert bill item at selection"""
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0:  # if selection exists copy at selection
            [model, paths] = selection.get_selected_rows()
            path = paths[0].get_indices()
            row = path[0]
            self.data.insert_bill_at_row(item, row)
        else:  # if no selection append a new item
            self.data.insert_bill_at_row(item, None)
        self.update_store()

    def edit_selected_row(self):
        """Edit selected bill item"""
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0:  # if selection exists, copy at selection
            [model, paths] = selection.get_selected_rows()
            path = paths[0].get_indices()
            row = path[0]
            bill = self.bills[row]
            # Edit bill data
            bill_dialog = BillDialog(self.parent, self.data, bill.get_model(), row)
            new_data = bill_dialog.run()
            if new_data is not None:  # if cancel is not pressed
                self.data.edit_bill_at_row(new_data, row)
        self.update_store()

    def delete_selected_row(self):
        """Delete selected rows"""
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0:
            [model, path] = selection.get_selected_rows()
            row = path[0].get_indices()[0]
            # Delete row
            self.data.delete_bill(row)
            # Modify selection
            if len(self.store) > 0:
                if row == 0:
                    path = Gtk.TreePath.new_from_indices([0])
                    self.tree.set_cursor(path)
                elif row >= len(self.store):
                    path = Gtk.TreePath.new_from_indices([len(self.store)-1])
                    self.tree.set_cursor(path)
                else:
                    path = Gtk.TreePath.new_from_indices([row-1])
                    self.tree.set_cursor(path)
        self.update_store()

    def copy_selection(self):
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0:  # if selection exists
            test_string = "BillView"
            [model, paths] = selection.get_selected_rows()
            path = paths[0]
            row = int(path.get_indices()[0])  # get item row
            item = self.bills[row]
            data = item.get_model()
            text = codecs.encode(pickle.dumps([test_string, data]), "base64").decode() # dump item as text
            self.clipboard.set_text(text, -1)  # push to clipboard
        else:  # if no selection
            log.warning("BillView - copy_selection - No items selected to copy")

    def paste_at_selection(self):
        text = self.clipboard.wait_for_text()  # get text from clipboard
        if text is not None:
            try:
                test_string = "BillView"
                itemlist = pickle.loads(codecs.decode(text.encode(), "base64"))  # recover item from string
                if itemlist[0] == test_string:
                    model = data.bill.BillData()
                    model.set_model(itemlist[1])
                    selection = self.tree.get_selection()
                    if selection.count_selected_rows() != 0:  # if selection exists copy at selection
                        [model_, paths] = selection.get_selected_rows()
                        path = paths[0].get_indices()
                        row = path[0]
                        # Handle different bill types
                        if self.bills[row].data.bill_type == misc.BILL_CUSTOM:
                            if model.bill_type == misc.BILL_NORMAL:
                                # create duplicate bill
                                bill = data.bill.Bill()
                                bill.set_model(model.get_model())
                                bill.update(self.schedule, self.cmbs, self.bills)
                                # Fill in values for custom bill from duplicate bill
                                model.item_normal_amount = bill.item_normal_amount
                                model.item_excess_amount = bill.item_excess_amount
                                model.item_qty = dict()
                                for itemno in bill.item_qty:
                                    model.item_qty[itemno] = [sum(bill.item_qty[itemno])]
                                # clear items for normal bill
                                model.mitems = dict()
                                model.item_part_percentage = dict()  # part rate for exess rate items
                                model.item_excess_part_percentage = dict()  # part rate for exess rate items
                                model.item_excess_rates = dict()  # list of excess rates above excess_percentage
                                # set bill type
                                model.bill_type = misc.BILL_CUSTOM
                        else:
                            model.mitems = dict()  # clear measured items
                            # clear additional elements for custom bill
                            model.item_qty = dict()  # qtys of items b/f
                            model.item_normal_amount = dict()  # total item amount for qty at normal rate
                            model.item_excess_amount = dict()  # amounts for qty at excess rate
                            # set bill type
                            model.bill_type = misc.BILL_NORMAL
                            
                        self.data.edit_bill_at_row(model.get_model(), row)
                    else:  # if selection do not exist paste at end
                        model.mitems = []  # clear measured items
                        self.data.insert_bill_at_row(model.get_model(), None)
                    self.update_store()
            except:
                log.warning('BillView - paste_at_selection - No valid data in clipboard')
        else:
            log.warning('BillView - paste_at_selection - No text on the clipboard')

    def clear(self):
        self.store.clear()

    def update_store(self):
        """Update Bill Store"""
        self.store.clear()
        for count, bill in enumerate(self.data.bills):
            self.store.append()
            self.store[count][0] = str(count + 1)
            self.store[count][1] = bill.get_text()

    def render_selected(self, folder, replacement_dict):
        # get selection
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0 and folder is not None:  # if selection exists
            # get path of selection
            [model, paths] = selection.get_selected_rows()
            path = paths[0].get_indices()
            code = self.data.render_bill(folder, replacement_dict, path)
            return code
        else:
            return (misc.CMB_WARNING, 'Please select a Bill for rendering')

    def __init__(self, parent, data_model, tree):
        self.parent = parent
        self.data = data_model
        self.tree = tree
        
        # Derived data
        self.schedule = self.data.schedule
        self.cmbs = self.data.cmbs
        self.bills = self.data.bills
        
        ## Setup treeview store
        # Item Description, Billed Flag, Tooltip, Colour
        self.store = Gtk.ListStore(str,str)
        # Treeview columns
        self.column_slno = Gtk.TreeViewColumn('Sl.No.')
        self.column_slno.props.fixed_width = 100
        self.column_slno.props.min_width = 100
        self.column_desc = Gtk.TreeViewColumn('Bill Description')
        self.column_desc.props.expand = True
        # Pack Columns
        self.tree.append_column(self.column_slno)
        self.tree.append_column(self.column_desc)
        # Treeview renderers
        self.renderer_slno = Gtk.CellRendererText()
        self.renderer_desc = Gtk.CellRendererText()
        # Pack renderers
        self.column_slno.pack_start(self.renderer_slno, True)
        self.column_desc.pack_start(self.renderer_desc, True)
        # Add attributes
        self.column_slno.add_attribute(self.renderer_slno, "text", 0)
        self.column_desc.add_attribute(self.renderer_desc, "markup", 1)
        # Set model for store
        self.tree.set_model(self.store)

        # Initialise clipboard
        self.clipboard = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD)
        
        # Connect Callbacks
        self.tree.connect("key-press-event", self.onKeyPressTreeviewSchedule)
        
        
class BillDialog:
    """Implements a dialog window for selecting items to be billed"""
    
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
        item_schedule = self.schedule
        captions = []
        columntypes = []
        populated_items = []
        cellrenderers = []
        dimensions = [[],[]]
        
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
            return str(populated_items[row][4])
        
        # Obtain values to be passed
        
        itemnos = self.data.schedule.get_itemnos()
        # Items specific to normal bill
        if self.billdata.bill_type == misc.BILL_NORMAL:
            # Evaluate a bill with current selected items for determining EXCEEDED flag
            bill = data.bill.Bill(self.billdata.get_model())
            bill.update(self.schedule, self.cmbs, self.bills)
            
            # Update model variables to include itemnos
            for itemno in itemnos:
                if itemno not in self.billdata.item_excess_rates:
                    self.billdata.item_excess_rates[itemno] = 0
                    self.billdata.item_part_percentage[itemno] = 100
                    self.billdata.item_excess_part_percentage[itemno] = 100
                
            for itemno in itemnos:
                item = self.schedule[itemno]
                description = item.extended_description_limited
                unit = item.unit
                rate = item.rate
                if bill.item_excess_qty[itemno] > 0:
                    flag = 'EXCEEDED'
                else:
                    flag = ''
                populated_items.append([itemno, description, unit, rate, flag])
            captions = ['AgmtNo', 'Description', 'Unit', 'Rate', 'Ex.Rate', 'PR(%)', 'Ex.PR(%)','Remarks']
            columntypes = [misc.MEAS_CUST, misc.MEAS_CUST, misc.MEAS_CUST, misc.MEAS_CUST, misc.MEAS_L, misc.MEAS_L, misc.MEAS_L, misc.MEAS_CUST]
            cellrenderers = [callback_agmntno] + [callback_description] + [callback_unit] + [callback_rate] + \
                            [None] * 3 + [callback_flag]
            dimensions = [[80,300,50,80,80,80,80,80],[False,True,False,False,False,False,False]]
        
        # Items specific to custom bill
        elif self.billdata.bill_type == misc.BILL_CUSTOM:
            # Update model variables to include itemnos
            for itemno in itemnos:
                if itemno not in self.billdata.item_qty:
                    self.billdata.item_qty[itemno] = [0]
                    self.billdata.item_normal_amount[itemno] = 0
                    self.billdata.item_excess_amount[itemno] = 0
                
            for itemno in itemnos:
                item = self.schedule[itemno]
                description = item.extended_description_limited
                unit = item.unit
                rate = item.rate
                populated_items.append([itemno, description, unit, rate])
            captions = ['AgmtNo', 'Description', 'Unit', 'Rate', 'Qty', 'Amount', 'Ex.Amnt']
            columntypes = [misc.MEAS_CUST, misc.MEAS_CUST, misc.MEAS_CUST, misc.MEAS_CUST, misc.MEAS_L, misc.MEAS_L, misc.MEAS_L]
            cellrenderers = [callback_agmntno] + [callback_description] + [callback_unit] + [callback_rate] + [None] * 3
            dimensions = [[80,300,80,80,80,80,80],[False,True,False,False,False,False]]
            
        # Raise Dialog for entry of per item values
        dialog = scheduledialog.ScheduleDialog(toplevel, item_schedule, [], captions, columntypes, cellrenderers, dimensions)

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
            if self.billdata.bill_type == misc.BILL_NORMAL:
                x = self.billdata.item_excess_rates[item[0]]
                y = self.billdata.item_part_percentage[item[0]]
                z = self.billdata.item_excess_part_percentage[item[0]]
                records.append(['', '', '', '', str(x), str(y), str(z),''])
            elif self.billdata.bill_type == misc.BILL_CUSTOM:
                x = self.billdata.item_qty[item[0]][0]
                y = self.billdata.item_normal_amount[item[0]]
                z = self.billdata.item_excess_amount[item[0]]
                records.append(['', '', '', '', str(x), str(y), str(z)])
        old_val = [[], records, '', []]

        dialog.set_model(old_val)  # modify dialog with current values of data

        # Run Dialog and get modified values
        model = dialog.run()

        if model is not None:
            records = model[1]
            for count, item in enumerate(populated_items):
                if self.billdata.bill_type == misc.BILL_NORMAL:
                    self.billdata.item_excess_rates[item[0]] = float(eval(records[count][4]))
                    self.billdata.item_part_percentage[item[0]] = float(eval(records[count][5]))
                    self.billdata.item_excess_part_percentage[item[0]] = float(eval(records[count][6]))
                elif self.billdata.bill_type == misc.BILL_CUSTOM:
                    self.billdata.item_qty[item[0]][0] = float(eval(records[count][4]))
                    self.billdata.item_normal_amount[item[0]] = float(eval(records[count][5]))
                    self.billdata.item_excess_amount[item[0]] = float(eval(records[count][6]))

    def onToggleCellRendererToggle(self, toggle, path_str):
        """On toggle clicked"""
        path_obj = Gtk.TreePath.new_from_string(path_str)
        path = path_obj.get_indices()
        
        def select_item(item, path, state = None):
            """Change state of item at path if item is toggleable
            
                Arguments:
                    item: Item selected
                    path: Path to item selected
                    state: Final state of item. Toggles on None
            """
            if not isinstance(item, data.measurement.MeasurementItemHeading) and self.locked[path] != True:
                if state != None:
                    self.selected[path] = state
                else:
                    state = self.selected[path]
                    if state is None:
                        state = False
                    self.selected[path] = not(state)
            
        if len(path) == 1:
            # Measure all items under CMB
            for p2, meas in enumerate(self.measurements_view.cmbs[path[0]]):
                if not isinstance(meas, data.measurement.Completion):
                    for p3, meas_item in enumerate(meas):
                        select_item(meas_item, path + [p2, p3], True)
        elif len(path) == 2:
            # Measure all items under Measurement
            meas = self.measurements_view.cmbs[path[0]][path[1]]
            if not isinstance(meas, data.measurement.Completion):
                for p3, meas_item in enumerate(meas):
                    select_item(meas_item, path + [p3], True)
        elif len(path) == 3:
            meas_item = self.measurements_view.cmbs[path[0]][path[1]][path[2]]
            select_item(meas_item, path)
            
        self.update_store()

    # General class methods

    def get_model(self):
        self.update_store()
        return self.billdata.get_model()

    def clear(self):
        self.data = None

    def update_store(self):
         
        # Update store from lock states
        self.measurements_view.update_store(self.selected)
        for count_cmb, cmb in enumerate(self.data.cmbs):
            for count_meas, meas in enumerate(cmb):
                if isinstance(meas, data.measurement.Measurement):
                    for count_meas_item, meas_item in enumerate(meas):
                        path = [count_cmb, count_meas, count_meas_item]
                        if self.locked[path]:
                            # Apply color to locked items
                            self.measurements_view.set_colour(path, misc.MEAS_COLOR_LOCKED)
                        elif self.selected[path]:
                            # Apply color to selected items
                            self.measurements_view.set_colour(path, misc.MEAS_COLOR_SELECTED)
                        elif isinstance(meas_item, data.measurement.MeasurementItemHeading):
                            self.measurements_view.set_colour(path, misc.MEAS_COLOR_LOCKED)
                else:
                    path = [count_cmb, count_meas]
                    # Apply color to locked items
                    self.measurements_view.set_colour(path, misc.MEAS_COLOR_LOCKED)
        
        # Update Data elements
        if self.combobox_bill_last_bill.get_active() != 0:
            self.billdata.prev_bill = self.combobox_bill_last_bill.get_active() - 1
        else:
            self.billdata.prev_bill = None
        self.billdata.title = self.entry_bill_title.get_text()
        self.billdata.cmb_name = self.entry_bill_cmbname.get_text()
        self.billdata.bill_date = self.entry_bill_bill_date.get_text()
        self.billdata.starting_page = int(self.entry_bill_starting_page.get_text())
        # Update mitems
        self.billdata.mitems = self.selected.get_paths()

    def __init__(self, parent, data_object, bill_data, this_bill=None):
        # Setup variables
        self.parent = parent
        self.data = data_object
        self.billdata = data.bill.BillData()
        self.billdata.set_model(copy.deepcopy(bill_data))
        self.this_bill = this_bill  # if in edit mode, use this to filter out this bill entries
        # Derived data
        self.schedule = self.data.schedule
        self.bills = self.data.bills
        self.cmbs = self.data.cmbs
        self.selected = data.datamodel.LockState(self.billdata.mitems)
        self.locked = self.data.get_lock_states() - self.selected
                
        # Setup dialog window
        self.builder = Gtk.Builder()
        self.builder.add_from_file(misc.abs_path("interface","billdialog.glade"))
        self.window = self.builder.get_object("dialog")
        self.window.set_default_size(1000,500)
        self.window.set_transient_for(self.parent)
        self.builder.connect_signals(self)
        # Get required objects
        self.treeview_bill = self.builder.get_object("treeview_bill")
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
        self.measurements_view.renderer_toggle.connect("toggled", self.onToggleCellRendererToggle)
        
        # Load data
        if self.billdata.prev_bill is not None:
            self.combobox_bill_last_bill.set_active(self.billdata.prev_bill + 1)
        else:
            self.combobox_bill_last_bill.set_active(0)
        self.entry_bill_title.set_text(self.billdata.title)
        self.entry_bill_cmbname.set_text(self.billdata.cmb_name)
        self.entry_bill_bill_date.set_text(self.billdata.bill_date)
        self.entry_bill_starting_page.set_text(str(self.billdata.starting_page))
        # Special conditions for custom bill
        if self.billdata.bill_type == misc.BILL_CUSTOM:
            self.treeview_bill.set_sensitive(False)  # Deactivate measurement item entry
            self.combobox_bill_last_bill.set_sensitive(False)  # Deactivate last bill selection
        
        # Update GUI
        self.update_store()

    def run(self):
        self.window.show_all()
        response = self.window.run()

        if response == 1:
            data_model = self.get_model()
            self.window.destroy()
            return data_model
        else:
            self.window.destroy()
            return None

