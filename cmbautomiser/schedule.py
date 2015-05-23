# !/usr/bin/env python
# -*- coding: utf-8 -*-
#
#  schedule.py
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

import pickle

from gi.repository import Gtk, Gdk, GLib

from undo import *

# local files import
from globalconstants import *
from misc import *

# class storing individual items in schedule of work
class ScheduleItem:
    def __init__(self, itemno="", decription="", unit="", rate=0, qty=0, reference="", excess_rate_percent=30):
        self.itemno = itemno
        self.description = decription
        self.unit = unit
        self.rate = rate
        self.qty = qty
        self.reference = reference
        self.excess_rate_percent = excess_rate_percent
        # extended descritption of item
        self.extended_description = ''
        self.extended_description_limited = ''

    def set(self, itemno="", decription="", unit="", rate=0.0, qty=0.0, reference="", excess_rate_percent=30):
        self.itemno = itemno
        self.description = decription
        self.unit = unit
        self.rate = rate
        self.qty = qty
        self.reference = reference
        self.excess_rate_percent = excess_rate_percent

        self.extended_description = ''
        self.extended_description_limited = ''

    def get(self):
        return [self.itemno, self.description, self.unit, self.rate, self.qty, self.reference, self.excess_rate_percent]

    def __setitem__(self, index, value):
        if index == 0:
            self.itemno = value
        elif index == 1:
            self.description = value
        elif index == 2:
            self.unit = value
        elif index == 3:
            self.rate = value
        elif index == 4:
            self.qty = value
        elif index == 5:
            self.reference = value
        elif index == 6:
            self.excess_rate_percent = value

    def __getitem__(self, index):
        if index == 0:
            return self.itemno
        elif index == 1:
            return self.description
        elif index == 2:
            return self.unit
        elif index == 3:
            return self.rate
        elif index == 4:
            return self.qty
        elif index == 5:
            return self.reference
        elif index == 6:
            return self.excess_rate_percent

    def print_item(self):
        print [self.itemno, self.description, self.unit, self.rate, self.qty, self.reference, self.excess_rate_percent]


# class storing schedule of rates for work
class Schedule:
    def __init__(self):
        self.items = []  # main data store of rows
        self.itemnos = []  # list of item numbers for easy search

    def get_item(self, itemno):
        try:
            index = self.itemnos.index(itemno)
        except:
            return None
        return self.items[index]

    def get_item_index(self, itemno):
        try:
            index = self.itemnos.index(itemno)
        except:
            return None
        return index

    def set_item(self, itemno, item):
        index = self.itemnos.index(itemno)
        self.items[index] = item
        self.itemnos[index] = item.itemno

    def append_item(self, item):
        self.items.append(item)
        self.itemnos.append(item.itemno)

    def remove_item(self, itemno):
        index = self.itemnos.index(itemno)
        del (self.items[index])
        del (self.itemnos[index])

    def get_item_by_index(self, index):
        return self.items[index]

    def set_item_at_index(self, index, item):
        self.items[index] = item
        self.itemnos[index] = item.itemno

    def insert_item_at_index(self, index, item):
        self.items.insert(index, item)
        self.itemnos.insert(index, item.itemno)

    def remove_item_at_index(self, index):
        del (self.items[index])
        del (self.itemnos[index])

    def __setitem__(self, index, value):
        self.items[index] = value
        self.itemnos[index] = value.itemno

    def __getitem__(self, index):
        return self.items[index]

    def length(self):
        return len(self.itemnos)

    def clear(self):
        del self.itemnos[:]
        del self.items[:]

    def update(self):
        # populate ScheduleItem.extended_description (Used for final billing)
        iter = 0
        extended_description = ''
        itemno = ''
        while iter < len(self.items):
            item = self.items[iter]
            # if main item
            if item.qty == 0 and item.unit == '' and item.rate == 0:
                if item.itemno != '':  # for main item start reset values
                    extended_description = ''
                    itemno = item.itemno
                    extended_description = extended_description + item.description
                else:
                    extended_description = extended_description + '\n' + item.description
                item.extended_description = item.description
            # if normal item
            elif item.itemno.startswith(itemno) and extended_description != '':  # if subitem
                item.extended_description = extended_description + '\n' + item.description
            else:
                item.extended_description = item.description
            if len(item.extended_description) > CMB_DESCRIPTION_MAX_LENGTH:
                item.extended_description_limited = item.extended_description[0:CMB_DESCRIPTION_MAX_LENGTH/2] + \
                    ' ... ' + item.extended_description[:CMB_DESCRIPTION_MAX_LENGTH/2]
            else:
                item.extended_description_limited = item.extended_description
            iter += 1

    def print_item(self):
        print "schedule start"
        for item in self.items:
            item.print_item()
        print "schedule end"


class ScheduleView:
    # Treeview event handler functions

    # Cell renderer treeview generic for editable text field
    def onScheduleCellEdited(self, widget, row, new_text, column):
        self.cell_renderer_text(int(row), column, new_text)

    # Cell renderer treeview generic for editable text field rates
    def onScheduleCellEditedRates(self, widget, row, new_text, column):
        try:
            validated_number = float(new_text)
        except:
            return
        self.cell_renderer_text(int(row), column, validated_number)

    # for browsing with tab key
    def onKeyPressTreeviewSchedule(self, treeview, event):
        keyname = event.get_keyval()[1]
        state = event.get_state()
        shift_pressed = bool(state & Gdk.ModifierType.SHIFT_MASK)
        path, col = treeview.get_cursor()
        if path <> None:
            ## only visible columns!!
            columns = [c for c in treeview.get_columns() if c.get_visible()]
            rows = [r for r in treeview.get_model()]
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
            elif keyname == Gdk.KEY_Escape:  # unselect all
                self.tree.get_selection().unselect_all()


    # tree view methods

    @undoable
    def append_item(self, itemlist):
        newrows = []

        for item in itemlist:
            self.schedule.append_item(item)
            newrows.append(self.schedule.length() - 1)
        self.update_store()

        yield "Append data items to schedule at row '{}'".format(newrows)
        # Undo action
        self.delete_row(newrows)
        self.update_store()

    def insert_item_at_selection(self, itemlist):
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0:  # if selection exists
            [model, paths] = selection.get_selected_rows()
            rows = []
            for i in range(0, len(itemlist)):
                rows.append(paths[0].get_indices()[0] + 1)
            self.insert_item_at_row(itemlist, rows)
        else:  # if no selection
            self.append_item(itemlist)

    @undoable
    def insert_item_at_row(self, itemlist, rows):  # note needs rows to be sorted
        newrows = []
        for i in range(0, len(rows)):
            self.schedule.insert_item_at_index(rows[i] + i - 1, itemlist[i])
            newrows.append(rows[i] + i - 1)
        self.update_store()

        yield "Insert data items to schedule at rows '{}'".format(rows)
        # Undo action
        self.delete_row(newrows)
        self.update_store()

    def delete_selection(self):
        # get rows to delete
        selection = self.tree.get_selection()
        [model, paths] = selection.get_selected_rows()
        rows = []
        for path in paths:  # get rows
            rows.append(int(path.get_indices()[0]))

        # delete rows
        self.delete_row(rows)

    @undoable
    def delete_row(self, rows):
        newrows = []
        items = []

        rows.sort()
        for i in range(0, len(rows)):
            items.append(self.schedule[rows[i] - i])
            newrows.append(rows[i] + i + 1)
            self.schedule.remove_item_at_index(rows[i] - i)
        self.update_store()

        yield "Delete data items from schedule at rows '{}'".format(rows)
        # Undo action
        self.insert_item_at_row(items, newrows)
        self.update_store()

    @undoable
    def cell_renderer_text(self, row, column, newvalue):
        oldvalue = self.schedule[row][column]
        self.schedule[row][column] = newvalue
        self.schedule.itemnos[row] = self.schedule[row][0]  # update agmntnos
        self.update_store()
        yield "Change data item at row:'{}' and column:'{}'".format(row, column)
        # Undo action
        self.schedule[row][column] = oldvalue
        self.schedule.itemnos[row] = self.schedule[row][0]  # update agmntnos
        self.update_store()

    def copy_selection(self):
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0:  # if selection exists
            [model, paths] = selection.get_selected_rows()
            items = []
            for path in paths:
                row = int(path.get_indices()[0])  # get item row
                item = self.schedule[row]
                items.append(item)  # save items
            text = pickle.dumps(items)  # dump item as text
            self.clipboard.set_text(text, -1)  # push to clipboard
        else:  # if no selection
            print("No items selected to copy")

    def paste_at_selection(self):
        text = self.clipboard.wait_for_text()  # get text from clipboard
        if text != None:
            try:
                itemlist = pickle.loads(text)  # recover item from string
                if isinstance(itemlist[0], ScheduleItem):
                    selection = self.tree.get_selection()
                    if selection.count_selected_rows() != 0:  # if selection exists
                        [model, paths] = selection.get_selected_rows()
                        index = paths[0].get_indices()[0]
                        rows = []
                        for i in range(0, len(itemlist)):
                            rows.append(int(index + 1))
                        self.insert_item_at_row(itemlist, rows)
                    else:  # if no selection
                        self.append_item(itemlist)
            except:
                print('No valid data in clipboard')
        else:
            print("No text on the clipboard.")

    def undo(self):
        setstack(self.stack)  # select schedule undo stack
        print self.stack.undotext()
        self.stack.undo()

    def redo(self):
        setstack(self.stack)  # select schedule undo stack
        print self.stack.redotext()
        self.stack.redo()

    def get_data_object(self):
        return self.schedule

    def set_data_object(self, data):
        self.schedule = data  # deprecated
        self.update_store()

    def clear(self):
        self.stack.clear()
        self.schedule.clear()
        self.update_store()

    def update_store(self):
        # add or remove required rows
        rownum = 0
        for row in self.store:
            rownum += 1
        if rownum > self.schedule.length():
            for i in range(rownum - self.schedule.length()):
                del self.store[-1]
        else:
            for i in range(self.schedule.length() - rownum):
                self.store.append()

        # Fill in Values
        for row in range(0, self.schedule.length()):
            item = self.schedule.get_item_by_index(row)
            self.store[row][0] = item.itemno
            self.store[row][1] = item.description
            self.store[row][2] = item.unit
            self.store[row][3] = str(round(item.rate, 2)) if item.rate != 0 else ''
            self.store[row][4] = str(round(item.qty, 2)) if item.qty != 0 else ''
            self.store[row][5] = item.reference
            self.store[row][6] = str(round(item.excess_rate_percent)) \
                                    if item.excess_rate_percent != 0 else ''
        setstack(self.stack)  # select undo stack

        # populate ScheduleItem.extended_description (Used for final billing)
        self.schedule.update()

    def __init__(self, schedule, ListStoreObject, TreeViewObject):
        self.schedule = schedule
        self.store = ListStoreObject
        self.tree = TreeViewObject

        self.stack = Stack()  # initialise undo/redo stack
        setstack(self.stack)  # select schedule undo stack

        self.clipboard = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD)  # initialise clipboard

        self.tree.connect("key-press-event", self.onKeyPressTreeviewSchedule)
      
