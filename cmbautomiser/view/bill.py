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
        """Add bill at selection"""
        bill_dialog = BillDialog(self.parent, data.bill.BillData(), self.data)
        data = bill_dialog.run()
        if data is not None:  # if cancel not pressed
            self.insert_bill_at_selection(data)
        self.update_store()

    def add_bill_custom(self):
        """Add custom bill at selection"""
        bill_dialog = BillDialog(self.parent, data.bill.BillData(misc.BILL_CUSTOM), self.data)
        data = bill_dialog.run()
        if data is not None:  # if cancel not pressed
            self.insert_bill_at_selection(data)
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
            bill_dialog = BillDialog(self.parent, bill.get_modal(), self.data)
            new_data = bill_dialog.run()
            if new_data is not None:  # if cancel is not pressed
                self.data.edit_bill_at_row(new_data, row)
        self.update_store()

    def delete_selected_row(self):
        # get rows to delete
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

    def copy_selection(self):
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0:  # if selection exists
            test_string = "BillView"
            [model, paths] = selection.get_selected_rows()
            path = paths[0]
            row = int(path.get_indices()[0])  # get item row
            item = self.bills[row]
            data = item.get_modal()
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
                    data = itemlist[1]
                if isinstance(data, data.bill.BillData):
                    selection = self.tree.get_selection()
                    if selection.count_selected_rows() != 0:  # if selection exists copy at selection
                        [model, paths] = selection.get_selected_rows()
                        path = paths[0].get_indices()
                        row = path[0]
                        # Handle different bill types
                        if self.bills[row].data.bill_type == misc.BILL_CUSTOM:
                            if data.bill_type == misc.BILL_NORMAL:
                                # create duplicate bill
                                bill = data.bill.Bill(self)
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
                            
                        self.data.edit_bill_at_row(data, row)
                    else:  # if selection do not exist paste at end
                        data.mitems = []  # clear measured items
                        self.data.insert_bill_at_row(data, None)
            except:
                log.warning('BillView - paste_at_selection - No valid data in clipboard')
        else:
            log.warning('BillView - paste_at_selection - No text on the clipboard')

    def clear(self):
        self.store.clear()

    def update_store(self):
        """Update Bill Store"""
        self.store.clear()
        for count, bill in enumerate(data.bills):
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
            return (misc.CMB_WARNING, 'Please select a Bill for rendering')

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
