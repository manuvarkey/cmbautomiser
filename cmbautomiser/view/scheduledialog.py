#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# scheduledialog.py
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

import undo

# local files import
from . import schedule
from __main__ import misc, data

# Setup logger object
log = logging.getLogger(__name__)

class ScheduleDialog:
    """Class implements a dialog box for entry of measurement records"""

    # General Methods

    def get_item_schedule_combobox(self):
        """Return packed combo box"""

        # combo box
        combo = Gtk.ComboBox.new_with_model(self.item_schedule_store)

        # text entries to be packed
        combo_text_agmntno = Gtk.CellRendererText()
        combo_divider1 = Gtk.CellRendererText()
        combo_text_desc = Gtk.CellRendererText()
        combo_divider2 = Gtk.CellRendererText()
        combo_text_unit = Gtk.CellRendererText()

        # start packing
        combo.pack_start(combo_text_agmntno, False)
        combo.pack_start(combo_divider1, False)
        combo.pack_start(combo_text_desc, True)
        combo.pack_start(combo_divider2, False)
        combo.pack_start(combo_text_unit, False)

        # set attributes
        combo.add_attribute(combo_text_agmntno, 'text', 0)
        combo.add_attribute(combo_text_desc, 'text', 1)
        combo.add_attribute(combo_text_unit, 'text', 2)

        combo.props.id_column = 0

        combo_divider1.props.text = '|'
        combo_divider2.props.text = '|'
        combo_text_desc.props.max_width_chars = misc.CMB_DESCRIPTION_WIDTH
        combo_text_desc.props.wrap_width = misc.CMB_DESCRIPTION_WIDTH
        combo_text_desc.props.ellipsize = 2
        combo_text_desc.set_fixed_size(-1, 5)
        combo_text_desc.set_fixed_height_from_font(1)

        return combo
    
    def model_width(self):
        """Width of schedule model loaded"""
        return self.schedule_view.model_width()

    def get_model(self):
        """Get data model"""
        remark = self.remark_cell.get_text()
        item_remarks = [cell.get_text() for cell in self.item_remarks_cell]
        data = [self.itemnos, self.schedule_view.get_model(), remark, item_remarks]
        return data

    def set_model(self, data):
        """Set data model"""
        self.itemnos = data[0]
        # Set combo boxes
        for itemno, combo in zip(self.itemnos, self.item_combos):
            combo.set_active_id(itemno)
        # Set schedule
        self.schedule_view.clear()
        self.schedule_view.set_model(copy.deepcopy(data[1]))
        self.schedule_view.update_store()
        # Set remark cells
        self.remark_cell.set_text(data[2])
        for cell, text in zip(self.item_remarks_cell, data[3]):
            cell.set_text(text)

    # Callbacks for GUI elements

    def onComboChanged(self, combo, index):
        """Select item from schedule on selection using combo box"""
        tree_iter = combo.get_active_iter()
        if tree_iter is not None:
            model = combo.get_model()
            itemno = model[tree_iter][0]
            self.itemnos[index] = itemno
        else:
            self.itemnos[index] = None

    def onButtonScheduleAddPressed(self, button):
        """Add row to schedule"""
        items = []
        items.append(data.schedule.ScheduleItemGeneric([""] * self.model_width()))
        self.schedule_view.insert_item_at_selection(items)

    def onButtonScheduleAddMultPressed(self, button):
        """Add multiple rows to schedule"""
        user_input = misc.get_user_input_text(self.window, "Please enter the number \nof rows to be inserted",
                                             "Number of Rows")
        try:
            number_of_rows = int(user_input)
        except:
            log.warning("Invalid number of rows specified")
            return
        items = []
        for i in range(0, number_of_rows):
            items.append(data.schedule.ScheduleItemGeneric([""] * self.model_width()))
        self.schedule_view.insert_item_at_selection(items)

    def onButtonScheduleDeletePressed(self, button):
        """Delete item from schedule"""
        self.schedule_view.delete_selected_rows()

    def onUndoSchedule(self, button):
        """Undo changes in schedule"""
        undo.setstack(self.stack)  # select schedule undo stack
        log.info('ScheduleViewGeneric - ' + str(self.stack.undotext()))
        self.stack.undo()

    def onRedoSchedule(self, button):
        """Redo changes in schedule"""
        undo.setstack(self.stack)  # select schedule undo stack
        log.info('ScheduleViewGeneric - ' + str(self.stack.redotext()))
        self.stack.redo()

    def onCopySchedule(self, button):
        """Copy rows from schedule"""
        self.schedule_view.copy_selection()

    def onPasteSchedule(self, button):
        """Paste rows in schedule"""
        self.schedule_view.paste_at_selection()

    def onImportScheduleClicked(self, button):
        """Import xlsx file into schedule"""
        filename = self.builder.get_object("filechooserbutton_schedule").get_filename()
        spreadsheet = misc.Spreadsheet(filename, 'r')
        models = spreadsheet.read_rows(columntypes = self.columntypes)
        items = []
        for model in models:
            item = data.schedule.ScheduleItemGeneric(*model)
            items.append(item)
        self.schedule_view.insert_item_at_selection(items)

    def onClearButtonPressed(self, button, combo):
        """Clear combobox selecting schedule item"""
        combo.set_active(-1)
        self.onComboChanged(combo, -1)

    def __init__(self, parent, item_schedule, itemnos, captions, columntypes, render_funcs, dimensions=None):
        """Initialise ScheduleDialog class
        
            Arguments:
                parent: Parent widget (Main window)
                item_schedule: Agreement schedule
                itemnos: Itemsnos of items being meaured
                captions: Captions of columns
                columntypes: Data types of columns. 
                             Takes following values:
                                misc.MEAS_NO: Integer
                                misc.MEAS_L: Float
                                misc.MEAS_DESC: String
                                misc.MEAS_CUSTOM: Value from render funstion provided
                render_funcs: Generates values of CUSTOM columns
        """
        # Setup variables
        self.parent = parent
        self.itemnos = itemnos
        self.captions = captions
        self.columntypes = columntypes
        self.render_funcs = render_funcs
        self.item_schedule = item_schedule
        
        # Save undo stack of parent
        self.stack_old = undo.stack()
        # Initialise undo/redo stack
        self.stack = undo.Stack()
        undo.setstack(self.stack)

        # Setup dialog window
        self.builder = Gtk.Builder()
        self.builder.add_from_file(misc.abs_path("interface","scheduledialog.glade"))
        self.window = self.builder.get_object("dialog")
        self.window.set_transient_for(self.parent)
        self.window.set_default_size(1000,500)
        self.builder.connect_signals(self)

        # Get required objects
        self.listbox_itemnos = self.builder.get_object("listbox_itemnos")
        self.tree = self.builder.get_object("treeview_schedule")
        
        # Setup schdule view for items
        self.schedule_view = schedule.ScheduleViewGeneric(self.parent,
            self.tree, self.captions, self.columntypes, self.render_funcs)
        if dimensions is not None:
            self.schedule_view.setup_column_props(*dimensions)

        # Setup liststore model for combobox from item schedule
        self.item_schedule_store = Gtk.ListStore(str, str, str, float, str)
        for row in range(self.item_schedule.length()):
            item = self.item_schedule.get_item_by_index(row)
            self.item_schedule_store.append([item.itemno, item.description, item.unit, float(item.rate), item.reference])

        # Setup remarks row
        row = Gtk.ListBoxRow()
        hbox = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)
        label = Gtk.Label("Remarks:")
        entry = Gtk.Entry()

        # Pack row
        row.add(hbox)
        hbox.pack_start(label, False, True, 3)
        hbox.pack_start(entry, True, True, 3)

        # Set additional properties
        label.props.width_request = 50

        # Add to list box
        self.listbox_itemnos.add(row)

        self.remark_cell = entry
        self.item_combos = []
        self.item_remarks_cell = []

        for itemno, index in zip(self.itemnos, list(range(len(self.itemnos)))):
            # Get items in row
            row = Gtk.ListBoxRow()
            hbox = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)
            entry = Gtk.Entry()
            button = Gtk.Button(stock=Gtk.STOCK_CLEAR)
            combo = self.get_item_schedule_combobox()

            # Pack row
            row.add(hbox)
            hbox.pack_start(entry, False, True, 3)
            hbox.pack_start(combo, True, True, 3)
            hbox.pack_start(button, False, True, 0)

            # Set additional properties
            entry.props.width_request = 50
            button.connect("clicked", self.onClearButtonPressed, combo)
            combo.connect("changed", self.onComboChanged, index)

            # Add to list box
            self.listbox_itemnos.add(row)

            # Save variables
            self.item_combos.append(combo)
            self.item_remarks_cell.append(entry)

    def run(self):
        """Display dialog box and return data model
        
            Returns:
                Data Model on Ok
                None on Cancel
        """
        self.window.show_all()
        response = self.window.run()
        # Reset undo stack of parent
        undo.setstack(self.stack_old)

        if response == 1:
            data = self.get_model()
            self.window.destroy()
            return data
        else:
            self.window.destroy()
            return None
