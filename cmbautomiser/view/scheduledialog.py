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
    
    def select_schedule(self, selected=None):
        '''Shows a dialog to select a schedule item 
        
            Arguments:
                selected: Current selected item
            Returns:
                Returns selected item or 'None' if user does not select any item.
        '''
        title = 'Select an item to be measured'
        dialog_window = Gtk.Dialog(title, self.window, Gtk.DialogFlags.MODAL,
            (Gtk.STOCK_CANCEL, Gtk.ResponseType.CANCEL,
             Gtk.STOCK_OK, Gtk.ResponseType.OK))
        dialog_window.set_transient_for(self.window)
        dialog_window.set_default_response(Gtk.ResponseType.OK)
        dialog_window.set_default_size(900,500)
        dialog_window.set_resizable(True)

        dialogBox = dialog_window.get_content_area()
        scrolled = Gtk.ScrolledWindow()
        scrolled.set_border_width(6)
        tree = Gtk.TreeView(self.item_schedule_store)
        dialogBox.pack_end(scrolled, True, True, 0)
        scrolled.add(tree)
        
        # Setup tree view
        tree.set_grid_lines(3)
        tree.set_enable_search(True)
        column1 = Gtk.TreeViewColumn("Agmt.No")
        column2 = Gtk.TreeViewColumn("Item Description")
        column3 = Gtk.TreeViewColumn("Unit")
        column4 = Gtk.TreeViewColumn("Reference")
        tree.append_column(column1)
        tree.append_column(column2)
        tree.append_column(column3)
        tree.append_column(column4)
        
        cell1 = Gtk.CellRendererText()
        cell2 = Gtk.CellRendererText()
        cell3 = Gtk.CellRendererText()
        cell4 = Gtk.CellRendererText()
        
        column1.pack_start(cell1, False)
        column2.pack_start(cell2, True)
        column3.pack_start(cell3, False)
        column4.pack_start(cell4, False)

        column1.add_attribute(cell1, "text", 0)
        column2.add_attribute(cell2, "text", 1)
        column3.add_attribute(cell3, "text", 2)
        column4.add_attribute(cell4, "text", 3)
        column1.set_fixed_width(80)
        column2.set_fixed_width(500)
        column3.set_fixed_width(100)
        column4.set_fixed_width(100)
        column2.props.expand = True
        
        cell2.props.wrap_width = 500
        cell2.props.wrap_mode = 2
        
        # Interactive search function
        def equal_func(model, column, key, iter, cols):
            """Equal function for interactive search"""
            search_string = ''
            for col in cols:
                search_string += ' ' + model[iter][col].lower()
            for word in key.split():
                if word.lower() not in search_string:
                    return True
            return False
        tree.set_search_equal_func(equal_func, [0,1,2,3])
        
        # Set old value
        if selected != None:
            itemnos = self.item_schedule.get_itemnos()
            if selected in itemnos:
                index = itemnos.index(selected)
                path = Gtk.TreePath.new_from_indices([index])
                tree.set_cursor(path)
                tree.scroll_to_cell(path, None)
        
        # Show Dialog window
        dialog_window.show_all()
        response = dialog_window.run()
        
        # Evaluate response
        if response == Gtk.ResponseType.OK:
            selection = tree.get_selection()
            if selection.count_selected_rows() != 0: # if selection exists
                [model, paths] = selection.get_selected_rows()
                path = paths[0].get_indices()
                itemno = self.item_schedule_store[path][0]
                dialog_window.destroy()
            return [True, itemno]
        else:
            dialog_window.destroy()
            return [False]
    
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
        self.itemnos = copy.copy(data[0])
        # Set item buttons
        for itemno, button in zip(self.itemnos, self.item_buttons):
            button.set_label(str(itemno))
        # Set schedule
        self.schedule_view.clear()
        self.schedule_view.set_model(copy.deepcopy(data[1]))
        self.schedule_view.update_store()
        # Set remark cells
        self.remark_cell.set_text(data[2])
        for cell, text in zip(self.item_remarks_cell, data[3]):
            cell.set_text(text)

    # Callbacks for GUI elements

    def OnItemSelectClicked(self, button, index):
        """Select item from schedule on selection using combo box"""
        response = self.select_schedule(self.itemnos[index])
        if response[0]:
            self.itemnos[index] = response[1]
            button.set_label(str(response[1]))
        
    def onClearButtonPressed(self, button, button_item, index):
        """Clear combobox selecting schedule item"""
        button_item.set_label('None')
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
        spreadsheet = misc.Spreadsheet(filename)
        models = spreadsheet.read_rows(columntypes = self.columntypes)
        items = []
        for model in models:
            item = data.schedule.ScheduleItemGeneric(model)
            items.append(item)
        self.schedule_view.insert_item_at_selection(items)

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

        # Setup liststore model for schedule selection from item schedule
        itemnos = self.item_schedule.get_itemnos()
        self.item_schedule_store = Gtk.ListStore(str, str, str, str)
        for itemno in itemnos:
            item = self.item_schedule[itemno]
            self.item_schedule_store.append([item.itemno, item.extended_description_limited, item.unit, item.reference])

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
        self.item_buttons = []
        self.item_remarks_cell = []

        for itemno, index in zip(self.itemnos, list(range(len(self.itemnos)))):
            # Get items in row
            row = Gtk.ListBoxRow()
            hbox = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)
            entry = Gtk.Entry()
            button_clear = Gtk.Button(stock=Gtk.STOCK_CLEAR)
            
            button_item = Gtk.Button.new_with_label("None")
            
            # Pack row
            row.add(hbox)
            hbox.pack_start(entry, False, True, 3)
            hbox.pack_start(button_item, True, True, 3)
            hbox.pack_start(button_clear, False, True, 0)

            # Set additional properties
            entry.props.width_request = 50
            button_clear.connect("clicked", self.onClearButtonPressed, button_item, index)
            button_item.connect("clicked", self.OnItemSelectClicked, index)

            # Add to list box
            self.listbox_itemnos.add(row)

            # Save variables
            self.item_buttons.append(button_item)
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
