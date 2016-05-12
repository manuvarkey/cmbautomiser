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

import pickle, copy, logging
from gi.repository import Gtk, Gdk, GLib
from undo import undoable

# local files import
from __main__ import data, misc
from . import billdialog

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
        toplevel = self.tree.get_toplevel()  # get toplevel window
        bill_dialog = BillDialog(toplevel, BillData(), self)
        data = bill_dialog.run()
        if data is not None:  # if cancel not pressed
            bill = Bill(self)
            bill.set_modal(data)
            self.insert_item_at_row(bill, None)

    def add_bill_custom(self):
        toplevel = self.tree.get_toplevel()  # get toplevel window
        bill_dialog = BillDialog(toplevel, BillData(BILL_CUSTOM), self)
        data = bill_dialog.run()
        if data is not None:  # if cancel not pressed
            bill = Bill(self)
            bill.set_modal(data)
            self.insert_item_at_row(bill, None)

    def insert_item_at_selection(self, item):
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0:  # if selection exists copy at selection
            [model, paths] = selection.get_selected_rows()
            path = paths[0].get_indices()
            row = path[0]
            self.insert_item_at_row(item, row)
        else:  # if no selection append a new item
            self.insert_item_at_row(item, None)

    @undoable
    def insert_item_at_row(self, item, row):  # note needs rows to be sorted
        if row is not None:
            new_row = row
            self.bills.insert(row, item)
        else:
            self.bills.append(item)
            new_row = len(self.bills) - 1
        self.update_store()

        yield "Insert data items to bill at row '{}'".format(new_row)
        # Undo action
        self.delete_row(new_row)
        self.update_store()

    def edit_selected_row(self):
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0:  # if selection exists copy at selection
            [model, paths] = selection.get_selected_rows()
            path = paths[0].get_indices()
            row = path[0]
            bill = self.bills[row]
            data = copy.deepcopy(bill.get_modal())  # make a copy of the modal

            toplevel = self.tree.get_toplevel()  # get toplevel window
            bill_dialog = BillDialog(toplevel, data, self, row)
            new_data = bill_dialog.run()
            if new_data is not None:  # if cancel not pressed
                self.edit_item_at_row(new_data, row)

    @undoable
    def edit_item_at_row(self, data, row):
        if row is not None:
            old_data = copy.deepcopy(self.bills[row].get_modal())
            self.bills[row].set_modal(data)
        self.update_store()

        yield "Edit bill item at row '{}'".format(row)
        # Undo action
        if row is not None:
            self.bills[row].set_modal(old_data)
        self.update_store()


    def delete_selected_row(self):
        # get rows to delete
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0:
            [model, path] = selection.get_selected_rows()
            row = path[0].get_indices()[0]
            # delete row
            self.delete_row(row)
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

    @undoable
    def delete_row(self, row):
        item = self.bills[row]
        del self.bills[row]
        self.update_store()

        yield "Delete data items from bill at row '{}'".format(row)
        # Undo action
        self.insert_item_at_row(item, row)
        self.update_store()

    def copy_selection(self):
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0:  # if selection exists
            [model, paths] = selection.get_selected_rows()
            path = paths[0]
            row = int(path.get_indices()[0])  # get item row
            item = self.bills[row]
            data = item.get_modal()
            text = pickle.dumps(data)  # dump item as text
            self.clipboard.set_text(text, -1)  # push to clipboard
        else:  # if no selection
            log.warning("No items selected to copy")

    def paste_at_selection(self):
        text = self.clipboard.wait_for_text()  # get text from clipboard
        if text is not None:
            try:
                data = pickle.loads(text)  # recover item from string
                if isinstance(data, BillData):
                    selection = self.tree.get_selection()
                    if selection.count_selected_rows() != 0:  # if selection exists copy at selection
                        [model, paths] = selection.get_selected_rows()
                        path = paths[0].get_indices()
                        row = path[0]
                        
                        # Handle different bill types
                        if self.bills[row].data.bill_type == BILL_CUSTOM:
                            if data.bill_type == BILL_NORMAL:
                                # create duplicate bill
                                bill = Bill(self)
                                bill.set_modal(data)
                                bill.update_values()
                                # Fill in values for custom bill from duplicate bill
                                data.item_normal_amount = bill.item_normal_amount
                                data.item_excess_amount = bill.item_excess_amount
                                data.item_qty = []
                                for qtys in bill.item_qty:
                                    data.item_qty.append([sum(qtys)])
                                # clear items for normal bill
                                data.mitems = []
                                data.item_part_percentage = []  # part rate for exess rate items
                                data.item_excess_part_percentage = []  # part rate for exess rate items
                                data.item_excess_rates = []  # list of excess rates above excess_percentage
                                # set bill type
                                data.bill_type = BILL_CUSTOM
                        else:
                            data.mitems = []  # clear measured items
                            # clear additional elements for custom bill
                            data.item_qty = []  # qtys of items b/f
                            data.item_normal_amount = []  # total item amount for qty at normal rate
                            data.item_excess_amount = []  # amounts for qty at excess rate
                            # set bill type
                            data.bill_type = BILL_NORMAL
                            
                        self.edit_item_at_row(data, row)
                    else:  # if selection do not exist paste at end
                        data.mitems = []  # clear measured items
                        bill = Bill(self)
                        bill.set_modal(data)
                        self.insert_item_at_row(bill, None)
            except:
                log.warning('No valid data in clipboard')
        else:
            log.warning("No text on the clipboard.")

    def clear(self):
        self.bills = []
        self.update_store()

    def update_store(self):

        # Update all bills
        for bill in self.bills:
            bill.update_values()

        # Update Bill Store
        self.store.clear()
        for count, bill in enumerate(self.bills):
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
            code = self.render(folder, replacement_dict, path)
            return code
        else:
            return (CMB_WARNING, 'Please select a Bill for rendering')

    def render(self, folder, replacement_dict, path, recursive=True):
        if self.bills[path[0]].data.bill_type == BILL_NORMAL:  # render only if normal bill
            bill = self.bills[path[0]]
            # fill in latex buffer
            bill.update_values()  # build all data structures
            if bill.prev_bill is not None:
                bill.prev_bill.update_values()
            latex_buffer = bill.get_latex_buffer([path[0]])
            latex_buffer_bill = bill.get_latex_buffer_bill()
            # make global variables replacements
            latex_buffer = replace_all(latex_buffer, replacement_dict)
            latex_buffer_bill = replace_all(latex_buffer_bill, replacement_dict)

            # include linked cmbs
            replacement_dict_cmbs = {}
            external_docs = ''
            for cmbpath in bill.cmb_ref:
                if cmbpath != -1:
                    external_docs += '\externaldocument{cmb_' + str(cmbpath + 1) + '}\n'
                elif bill.data.prev_bill is not None: # prev abstract
                    external_docs += '\externaldocument{abs_' + str(bill.data.prev_bill + 1) + '}\n'
            replacement_dict_cmbs['$cmbexternaldocs$'] = external_docs
            latex_buffer = replace_all_vanilla(latex_buffer, replacement_dict_cmbs)

            # write output
            filename = posix_path(folder, 'abs_' + str(path[0] + 1) + '.tex')
            file_latex = open(filename, 'w')
            file_latex.write(latex_buffer)
            file_latex.close()

            filename_bill = posix_path(folder, 'bill_' + str(path[0] + 1) + '.tex')
            file_latex_bill = open(filename_bill, 'w')
            file_latex_bill.write(latex_buffer_bill)
            file_latex_bill.close()

            filename_bill_ods = posix_path(folder, 'bill_' + str(path[0] + 1) + '.xlsx')

            # run latex on file and dependencies
            if recursive:  # if recursive call
                # Render all cmbs depending on the bill
                for cmb_ref in bill.cmb_ref:
                    if cmb_ref is not -1:  # if not prev bill
                        code = self.measurement_view.render(folder, replacement_dict, self.bills, [cmb_ref], False)
                        if code[0] == CMB_ERROR:
                            return code
                # Render prev bill
                if bill.prev_bill is not None and bill.prev_bill.data.bill_type == BILL_NORMAL:
                    code = self.render(folder, replacement_dict, [bill.data.prev_bill], False)
                    if code[0] == CMB_ERROR:
                        return code

            # Render this bill
            code = run_latex(posix_path(folder), filename)
            if code == CMB_ERROR:
                return (CMB_ERROR, 'Rendering of Bill: ' + self.bill.data.title + ' failed')
            code_bill = run_latex(posix_path(folder), filename_bill)
            if code_bill == CMB_ERROR:
                return (CMB_ERROR, 'Rendering of Bill Schedule: ' + self.bill.data.title + ' failed')

            # Render all cmbs again to rebuild indexes on recursive run
            if recursive:  # if recursive call
                for cmb_ref in bill.cmb_ref:
                    if cmb_ref is not -1:  # if not prev bill
                        code = self.measurement_view.render(folder, replacement_dict, self.bills, [cmb_ref], False)
                        if code[0] == CMB_ERROR:
                            return code
                bill.export_ods_bill(filename_bill_ods,replacement_dict)

            return (CMB_INFO, 'Bill: ' + self.bills[path[0]].data.title + ' rendered successfully')
        else:
            return (CMB_WARNING, 'Rendering of custom bill not supported')

    def __init__(self, parent, data, tree):
        self.data = data
        self.tree = tree
        self.schedule = data.schedule
        self.cmbs = data.cmbs
        self.bills = data.bills
        
        ## Setup treeview store
        # Item Description, Billed Flag, Tooltip, Colour
        self.store = Gtk.ListStore([str,str])
        # Treeview columns
        self.column_slno = Gtk.TreeViewColumn('Sl.No.')
        self.column_slno.props.expand = True
        self.column_desc = Gtk.TreeViewColumn('Bill Description')
        self.column_slno.props.fixed_width = 150
        self.column_slno.props.min_width = 150
        # Treeview renderers
        self.renderer_slno = Gtk.CellRendererText()
        self.renderer_desc = Gtk.CellRendererText()
        # Pack renderers
        self.column_slno.pack_start(self.renderer_slno, True)
        self.column_desc.pack_start(self.renderer_desc, True)
        # Add attributes
        self.column_slno.add_attribute(self.renderer_desc, "text", 0)
        self.column_desc.add_attribute(self.renderer_toggle, "text", 1)
        # Set model for store
        self.tree.set_model(self.store)

        # Initialise clipboard
        self.clipboard = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD)
        
        # Connect Callbacks
        self.tree.connect("key-press-event", self.onKeyPressTreeviewSchedule)
