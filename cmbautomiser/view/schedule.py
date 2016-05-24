#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# schedule.py
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

import logging, pickle, codecs

from gi.repository import Gtk, Gdk, GLib

# local files import
from __main__ import misc, data, undo
from undo import undoable

# Setup logger object
log = logging.getLogger(__name__)


class ScheduleViewGeneric:
    """Implements a view for display and manipulation of ScheduleGeneric over a treeview"""

    # Call backs for treeview

    def onScheduleCellEditedText(self, widget, row, new_text, column):
        """Treeview cell renderer for editable text field
        
            User Data:
                column: column in ListStore being edited
        """
        self.cell_renderer_text(int(row), column, new_text)

    def onScheduleCellEditedNum(self, widget, row, new_text, column):
        """Treeview cell renderer for editable number field
        
            User Data:
                column: column in ListStore being edited
        """
        try:  # check whether item evaluates fine
            eval(new_text)
        except:
            log.warning("ScheduleViewGeneric - onScheduleCellEditedNum - evaluation of [" 
            + new_text + "] failed")
            return
        self.cell_renderer_text(int(row), column, new_text)

    def onEditStarted(self, widget, editable, path, column):
        """Fill in text from schedule when schedule view column get edited
        
            User Data:
                column: column in ListStore being edited
        """
        row = int(path)
        item = self.schedule.get_item_by_index(row).get_model()
        editable.props.text = str(item[column])

    # for browsing with tab key
    def onKeyPressTreeviewSchedule(self, treeview, event):
        """Handle key presses"""
        keyname = event.get_keyval()[1]
        state = event.get_state()
        shift_pressed = bool(state & Gdk.ModifierType.SHIFT_MASK)
        path, col = treeview.get_cursor()
        if path != None:
            ## only visible columns!!
            columns = [c for c in treeview.get_columns() if c.get_visible() and self.celldict[c].props.editable]
            rows = [r for r in treeview.get_model()]
            if col in columns:
                colnum = columns.index(col)
                rownum = path[0]
                if keyname in [Gdk.KEY_Tab, Gdk.KEY_ISO_Left_Tab]:
                    if shift_pressed == 1:
                        if colnum - 1 >= 0:
                            prev_column = columns[colnum - 1]
                        else:
                            tmodel = treeview.get_model()
                            titer = tmodel.iter_previous(tmodel.get_iter(path))
                            if titer is None:
                                titer = tmodel.get_iter_first()
                                path = tmodel.get_path(titer)
                                prev_column = columns[0]
                            else:
                                path = tmodel.get_path(titer)
                                prev_column = columns[-1]
                        GLib.timeout_add(50, treeview.set_cursor, path, prev_column, True)
                    else:
                        if colnum + 1 < len(columns):
                            next_column = columns[colnum + 1]
                        else:
                            tmodel = treeview.get_model()
                            titer = tmodel.iter_next(tmodel.get_iter(path))
                            if titer is None:
                                titer = tmodel.get_iter_first()
                            path = tmodel.get_path(titer)
                            next_column = columns[0]
                        GLib.timeout_add(50, treeview.set_cursor, path, next_column, True)
                elif keyname in [Gdk.KEY_Return, Gdk.KEY_KP_Enter]:
                    if shift_pressed == 1:
                        if rownum - 1 >= 0:
                            rownum -= 1
                        else:
                            rownum = 0
                        path = [rownum]
                        GLib.timeout_add(50, treeview.set_cursor, path, col, True)
                    else:
                        if rownum + 1 < len(rows):
                            rownum += 1
                        else:
                            rownum = len(rows) - 1
                        path = [rownum]
                        GLib.timeout_add(50, treeview.set_cursor, path, col, True)
                elif keyname in [Gdk.KEY_Control_L, Gdk.KEY_Control_R, Gdk.KEY_Escape]:  # unselect all
                    self.tree.get_selection().unselect_all()

    # Class methods
    
    def setup_column_props(self, widths, expandables):
        """Set column properties
            Arguments:
                widths: List of column widths type-> [int, ...]. None values are skiped.
                expandables: List of expand property type-> [bool, ...]. None values are skiped.
        """
        for column, width, expandable in zip(self.columns, widths, expandables):
            if width != None:
                column.set_min_width(width)
                column.set_fixed_width(width)
                self.celldict[column].props.wrap_width = width
            if expandable != None:
                column.set_expand(expandable)
                

    def insert_item_at_selection(self, itemlist):
        """Insert items at selected row"""
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0:  # if selection exists
            [model, paths] = selection.get_selected_rows()
            rows = []
            for i in range(0, len(itemlist)):
                rows.append(paths[0].get_indices()[0])
            self.insert_item_at_row(itemlist, rows)
        else:  # if no selection
            self.append_item(itemlist)

    def delete_selected_rows(self):
        """Delete selected rows"""
        selection = self.tree.get_selection()
        [model, paths] = selection.get_selected_rows()
        rows = []
        for path in paths:  # get rows
            rows.append(int(path.get_indices()[0]))

        # delete rows
        self.delete_row(rows)
        
    @undoable
    def append_item(self, itemlist):
        """Undoable function to append items to schedule"""
        newrows = []

        for item in itemlist:
            self.schedule.append_item(item)
            newrows.append(self.schedule.length() - 1)
        self.update_store()

        yield "Append data items to schedule at row '{}'".format(newrows)
        # Undo action
        self.delete_row(newrows)
    
    @undoable
    def insert_item_at_row(self, itemlist, rows):
        """Undoable function to insert items to schedule at given rows
        
            Remarks:
                Need rows to be sorted.
        """
        newrows = []
        for i in range(0, len(rows)):
            self.schedule.insert_item_at_index(rows[i], itemlist[i])
            newrows.append(rows[i] + i)
        self.update_store()

        yield "Insert data items to schedule at rows '{}'".format(rows)
        # Undo action
        self.delete_row(newrows)

    @undoable
    def delete_row(self, rows):
        """Undoable function to delete a set of rows"""
        newrows = []
        items = []

        rows.sort()
        for i in range(0, len(rows)):
            items.append(self.schedule[rows[i] - i])
            newrows.append(rows[i])
            self.schedule.remove_item_at_index(rows[i] - i)
        self.update_store()

        yield "Delete data items from schedule at rows '{}'".format(rows)
        # Undo action
        self.insert_item_at_row(items, newrows)

    @undoable
    def cell_renderer_text(self, row, column, newvalue):
        """Undoable function for modifying value of a treeview cell"""
        oldvalue = self.schedule[row][column]
        self.schedule[row][column] = newvalue
        self.update_store()

        yield "Change data item at row:'{}' and column:'{}'".format(row, column)
        # Undo action
        self.schedule[row][column] = oldvalue
        self.update_store()

    def copy_selection(self):
        """Copy selected rows to clipboard"""
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0:  # if selection exists
            test_string = "Schedule:" + str(self.columntypes)
            [model, paths] = selection.get_selected_rows()
            items = []
            for path in paths:
                row = int(path.get_indices()[0])  # get item row
                item = self.schedule[row]
                items.append(item)  # save items
            text = codecs.encode(pickle.dumps([test_string, items]), "base64").decode() # dump item as text
            self.clipboard.set_text(text, -1)  # push to clipboard
        else:  # if no selection
            log.warning("ScheduleViewGeneric - copy_selection - No items selected to copy")

    def paste_at_selection(self):
        """Paste copied item at selected row"""
        text = self.clipboard.wait_for_text()  # get text from clipboard
        if text is not None:
            test_string = "Schedule:" + str(self.columntypes)
            try:
                itemlist = pickle.loads(codecs.decode(text.encode(), "base64"))  # recover item from string
                if itemlist[0] == test_string:
                    selection = self.tree.get_selection()
                    if selection.count_selected_rows() != 0:  # if selection exists
                        [model, paths] = selection.get_selected_rows()
                        rows = []
                        for i in range(0, len(itemlist[1])):
                            rows.append(int(paths[0].get_indices()[0]))
                        self.insert_item_at_row(itemlist[1], rows)
                    else:  # if no selection
                        self.append_item(itemlist[1])
            except:
                log.warning('ScheduleViewGeneric - paste_at_selection - No valid data in clipboard')
        else:
            log.warning("ScheduleViewGeneric - paste_at_selection - No text in clipboard.")

    def model_width(self):
        """Returns the width of model"""
        return len(self.columntypes)

    def get_model(self):
        """Return data model"""
        return self.schedule.get_model()

    def set_model(self, schedule):
        """Set data model"""
        self.schedule.set_model(schedule)
        self.update_store()

    def clear(self):
        """Clear all schedule items"""
        self.schedule.clear()
        self.update_store()

    def update_store(self):
        """Update store to reflect modified schedule"""
        log.info('ScheduleViewGeneric - update_store')
        # Add or remove required rows
        rownum = 0
        for row in self.store:
            rownum += 1
        rownum = len(self.store)
        if rownum > self.schedule.length():
            for i in range(rownum - self.schedule.length()):
                del self.store[-1]
        else:
            for i in range(self.schedule.length() - rownum):
                self.store.append()

        # Find formated items and fill in values
        for row in range(0, self.schedule.length()):
            item = self.schedule.get_item_by_index(row).get_model()
            display_item = []
            for item_elem, columntype, render_func in zip(item, self.columntypes, self.render_funcs):
                try:
                    if item_elem != "" or columntype == misc.MEAS_CUST:
                        if columntype == misc.MEAS_CUST:
                            display_item.append(render_func(item, row))
                        if columntype == misc.MEAS_DESC:
                            display_item.append(item_elem)
                        elif columntype == misc.MEAS_NO:
                            value = str(int(eval(item_elem))) if item_elem not in ['0','0.0'] else ''
                            display_item.append(value)
                        elif columntype == misc.MEAS_L:
                            value = str(round(float(eval(item_elem)), 3)) if item_elem not in ['0','0.0'] else ''
                            display_item.append(value)
                    else:
                        display_item.append("")
                except TypeError:
                    display_item.append("")
                    log.warning('ScheduleViewGeneric - Wrong value loaded in store - '  + str(item_elem))
            self.store[row] = display_item

    def __init__(self, parent, tree, captions, columntypes, render_funcs):
        """Initialise ScheduleViewGeneric class
        
            Arguments:
                parent: Parent widget (dialog/window)
                tree: Treeview for implementing schedule
                captions: Captions to be displayed in columns
                columntypes: Column data type
                    (takes the values misc.MEAS_NO,
                                      misc.MEAS_L,
                                      misc.MEAS_DESC,
                                      misc.MEAS_CUST)
                render_funcs: Fucntions generating values of CUSTOM columns
        """
        log.info('ScheduleViewGeneric - Initialise')
        # Setup variables
        self.parent = parent
        self.tree = tree
        self.captions = captions
        self.columntypes = columntypes
        self.render_funcs = render_funcs
        self.schedule = data.schedule.ScheduleGeneric()

        # Setup treeview
        data_types = [str] * len(self.columntypes)
        self.store = Gtk.ListStore(*data_types)
        self.tree.set_model(self.store)
        self.celldict = dict()
        self.columns = []
        
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
            
        # Set interactive search function
        cols = [i for i,x in enumerate(self.columntypes) if x == misc.MEAS_DESC]
        self.tree.set_search_equal_func(equal_func, [0,1,5])

        for columntype, caption, render_func, i in zip(self.columntypes, self.captions, self.render_funcs,
                                                       range(len(self.columntypes))):
            cell = Gtk.CellRendererText()

            column = Gtk.TreeViewColumn(caption, cell, text=i)
            column.props.resizable = True
            
            self.columns.append(column)  # Add column to list of columns
            self.celldict[column] = cell  # Add cell to column map for future ref

            self.tree.append_column(column)
            self.tree.props.search_column = 0

            if columntype == misc.MEAS_NO:
                cell.set_property("editable", True)
                column.props.min_width = 75
                column.props.fixed_width = 75
                if render_func is None:
                    cell.connect("edited", self.onScheduleCellEditedNum, i)
                    cell.connect("editing_started", self.onEditStarted, i)
                else:
                    cell.connect("edited", render_func, i)
            elif columntype == misc.MEAS_L:
                cell.set_property("editable", True)
                column.props.min_width = 75
                column.props.fixed_width = 75
                if render_func is None:
                    cell.connect("edited", self.onScheduleCellEditedNum, i)
                    cell.connect("editing_started", self.onEditStarted, i)
                else:
                    cell.connect("edited", render_func, i)
            elif columntype == misc.MEAS_DESC:
                cell.set_property("editable", True)
                column.props.fixed_width = 150
                column.props.min_width = 150
                column.props.expand = True
                cell.props.wrap_width = 150
                cell.props.wrap_mode = 2
                if render_func is None:
                    cell.connect("edited", self.onScheduleCellEditedText, i)
                    cell.connect("editing_started", self.onEditStarted, i)
                else:
                    cell.connect("edited", render_func, i)
            elif columntype == misc.MEAS_CUST:
                cell.set_property("editable", False)
                column.props.fixed_width = 100
                column.props.min_width = 100
                cell.props.wrap_width = 100
                cell.props.wrap_mode = 2
                if render_func is None:
                    cell.connect("edited", self.onScheduleCellEditedText, i)
                    cell.connect("editing_started", self.onEditStarted, i)
                else:
                    cell.connect("edited", render_func, i)

        # Intialise clipboard
        self.clipboard = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD)  # initialise clipboard

        # Connect signals with custom userdata
        self.tree.connect("key-press-event", self.onKeyPressTreeviewSchedule)


class ScheduleView(ScheduleViewGeneric):
    """Implements a view for display and manipulation of schedule of rates over a treeview"""
    
    def __init__(self, parent, tree, schedule):
        """Initialises a ScheduleViewGeneric and sets up custom parameters
        
            Arguments:
                parent: Master window
                tree: TreeView for implementing ScheduleView
                schedule: Schedule Data model for storing values
        """
        log.info('ScheduleView - Initialise')
        captions = ['Agmt.No.','Item Description','Unit','Rate','Qty','Reference','Excess %']
        columntypes = [misc.MEAS_DESC, misc.MEAS_DESC, misc.MEAS_DESC,
                       misc.MEAS_L, misc.MEAS_L, misc.MEAS_DESC, misc.MEAS_L]
        render_funcs = [None,None,None,None,None,None,None]
        widths = [80,500,100,100,100,100,100]
        expandables = [False,True,False,False,False,False,False]
               
        # Initialise base class
        super(ScheduleView, self).__init__(parent, tree, captions, columntypes, render_funcs)
        self.setup_column_props(widths, expandables)
        self.schedule = schedule  # Override default schedule
        
