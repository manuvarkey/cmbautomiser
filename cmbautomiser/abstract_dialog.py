#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# abstract_dialog.py
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

from cmb import *
from misc import *



# Class storing Abstract dialog
class AbstractDialog:
    # General signal handler Methods

    def OnNumberTextChanged(self, entry):
        try:
            val = int(entry.get_text())
            entry.set_text(str(val))
        except ValueError:
            entry.set_text('')

    def onToggleCellRendererToggle(self, toggle, path_str):
        path_obj = Gtk.TreePath.new_from_string(path_str)
        path = path_obj.get_indices()
        if len(path) == 3:
            meas_item = self.measurements_view.cmbs[path[0]][path[1]][path[2]]
            if not isinstance(meas_item, MeasurementItemHeading) and not (path in self.locked):
                meas_item.set_billed_flag(not (meas_item.get_billed_flag()))
        self.update_store()

    # Class methods

    def get_model(self):
        self.update_store()
        return [self.mitems,self.int_m_item.get_model(),self.int_m_item.itemtype]

    def clear(self):
        self.data = None

    def update_store(self):
        # Update measurements store
        self.measurements_view.update_store()

        # update data.mitems from self.cmbs
        self.mitems = []  # clear mitems
        for count_cmb, cmb in enumerate(self.cmbs):
            for count_meas, meas in enumerate(cmb):
                if isinstance(meas, Measurement):
                    for count_meas_item, meas_item in enumerate(meas):
                        path = Gtk.TreePath.new_from_string(
                            str(count_cmb) + ':' + str(count_meas) + ':' + str(count_meas_item))
                        path_iter = self.measurements_view.store.get_iter(path)
                        if [count_cmb, count_meas, count_meas_item] in self.locked:
                            self.measurements_view.store.set_value(path_iter, 3,
                                                                   MEAS_COLOR_LOCKED)  # apply color to locked items
                        elif meas_item.get_billed_flag() is True:
                            self.mitems.append([count_cmb, count_meas, count_meas_item])
                            self.measurements_view.store.set_value(path_iter, 3,
                                                                   MEAS_COLOR_SELECTED)  # apply color to selected items
                        elif isinstance(meas_item, MeasurementItemHeading):
                            self.measurements_view.store.set_value(path_iter, 3, MEAS_COLOR_LOCKED)
                else:
                    path = Gtk.TreePath.new_from_string(str(count_cmb) + ':' + str(count_meas))
                    path_iter = self.measurements_view.store.get_iter(path)
                    self.measurements_view.store.set_value(path_iter, 3, MEAS_COLOR_LOCKED)  # apply color to locked

        # Update self.int_m_item
        if self.mitems:
            p = self.mitems[0]
            item = self.cmbs[p[0]][p[1]][p[2]]
            type = item.itemtype
            if self.int_m_item is None:
                self.int_m_item = MeasurementItemCustom(item.get_model(),type)
            elif item.itemtype != self.int_m_item.itemtype:
                self.int_m_item = MeasurementItemCustom(item.get_model(),type)
            # Populate values
            self.int_m_item.records = []
            for path in self.mitems:
                item = self.measurements_view.cmbs[path[0]][path[1]][path[2]]
                values = item.export_abstract(item.records, item.user_data)
                # Make dicionary magic
                cmbbf = 'ref:meas:'+ str(path) + ':1'
                label = 'ref:abs:'+ str(path) + ':1'
                values[0] = r'Qty B/F MB.No.\emph{\nameref{' + cmbbf + r'} Pg.No. \pageref{' + cmbbf \
                            + r'}}\phantomsection\label{' + label + '}'
                self.int_m_item.append_record(RecordCustom(values,
                    self.int_m_item.cust_funcs,self.int_m_item.total_func_item,
                    self.int_m_item.columntypes))

        # update remarks column
        self.int_m_item.set_remark(self.entry_abstract_remark.get_text())
        self.update_lock()  # lock item_type

    def update_lock(self):
        if self.int_m_item is not None:
            type = self.int_m_item.itemtype
            if type:
                for count_cmb, cmb in enumerate(self.cmbs):
                    for count_meas, meas in enumerate(cmb):
                        if isinstance(meas, Measurement):
                            for count_meas_item, meas_item in enumerate(meas):
                                path = [count_cmb,count_meas,count_meas_item]
                                if isinstance(meas_item,MeasurementItemCustom):
                                    if path not in self.locked:
                                        if meas_item.itemtype != type:
                                            self.locked.append(path)
                                    elif meas_item.itemtype == type:
                                            self.locked.remove(path)


    def __init__(self, parent, data, measurements_view, schedule):
        # Setup variables
        self.parent = parent
        self.measurements_view_master = measurements_view
        self.schedule = schedule
        self.cmbs = copy.deepcopy(self.measurements_view_master.cmbs)  # make copy of cmbs for tinkering
        self.locked = []  # paths to items locked from changing
        self.data = data
        self.mitems = []
        self.int_m_item = None

        # Setup dialog window
        self.builder = Gtk.Builder()
        self.builder.add_from_file(abs_path("interface","abstractdialog.glade"))
        self.window = self.builder.get_object("dialog")
        self.window.set_transient_for(self.parent)
        self.builder.connect_signals(self)

        # Get required objects
        self.treeview_abstract = self.builder.get_object("treeview_abstract")
        self.liststore_abstract = self.builder.get_object("liststore_abstract")
        self.tree_abstract_text_billed = self.builder.get_object("tree_abstract_text_billed")
        # Text Entries
        self.entry_abstract_remark = self.builder.get_object("entry_abstract_remark")
        # setup cmb tree view
        self.measurements_view = MeasurementsView(self.schedule, self.liststore_abstract, self.treeview_abstract)
        self.measurements_view.set_data_object(self.schedule, self.cmbs)

        # Connect signals
        self.tree_abstract_text_billed.connect("toggled", self.onToggleCellRendererToggle)

        # Restore UI elements from data

        if self.data is not None:
            self.mitems = data[0]
            self.int_m_item = MeasurementItemCustom(data[1],data[2])
            self.entry_abstract_remark.set_text(self.int_m_item.get_remark())

        # Measured Items
        # lock all measured and non similar items
        for count_cmb, cmb in enumerate(self.cmbs):
            for count_meas, meas in enumerate(cmb):
                if isinstance(meas, Measurement):
                    for count_meas_item, meas_item in enumerate(meas):
                        if meas_item.get_billed_flag() is True:
                            self.locked.append([count_cmb, count_meas, count_meas_item])
                        if not isinstance(meas_item,MeasurementItemCustom):
                            self.locked.append([count_cmb, count_meas, count_meas_item])
                        else:
                            if meas_item.export_abstract == None:
                                self.locked.append([count_cmb, count_meas, count_meas_item])
                            elif self.data[1] is not None:
                                if meas_item.itemtype != self.data[2]:
                                    self.locked.append([count_cmb, count_meas, count_meas_item])

        # remove any selected items
        for path in self.mitems:
            if path in self.locked:
                self.locked.remove(path)

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
