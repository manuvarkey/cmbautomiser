#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# abstractdialog.py
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

from __main__ import misc, data

# Setup logger object
log = logging.getLogger(__name__)

class AbstractDialog:
    """Class creates a dialog window for selecting items to be abstracted"""
    
    # Callbacks

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
            meas_item = self.data.cmbs[path[0]][path[1]][path[2]]
            if not isinstance(meas_item, data.measurement.MeasurementItemHeading) and self.locked[path] != True:
                state = self.selected[path[0]][path[1]][path[2]]
                if state is None:
                    state = False
                self.selected[path[0]][path[1]][path[2]] = not(state)
        self.update_store()

    # Class methods

    def get_model(self):
        self.update_store()
        return [self.mitems,self.int_m_item.get_model(),self.int_m_item.itemtype]

    def clear(self):
        self.data = None

    def update_store(self):
        
        # Update mitems
        self.mitems = self.selected.get_paths()

        # Update self.int_m_item
        if self.mitems:
            p = self.mitems[0]
            item = self.data.cmbs[p[0]][p[1]][p[2]]
            type_ = item.itemtype
            if self.int_m_item is None:
                self.int_m_item = data.measurement.MeasurementItemCustom(item.get_model(), type_)
            elif item.itemtype != self.int_m_item.itemtype:
                self.int_m_item = data.measurement.MeasurementItemCustom(item.get_model(), type_)
            # Populate values
            self.int_m_item.records = []
            for path in self.mitems:
                item = self.data.cmbs[path[0]][path[1]][path[2]]
                values = item.export_abstract(item.records, item.user_data)
                # Make dicionary magic
                cmbbf = 'ref:meas:'+ str(path) + ':1'
                label = 'ref:abs:'+ str(path) + ':1'
                values[0] = r'Qty B/F MB.No.\emph{\nameref{' + cmbbf + r'} Pg.No. \pageref{' + cmbbf \
                            + r'}}\phantomsection\label{' + label + '}'
                self.int_m_item.append_record(data.measurement.RecordCustom(values,
                    self.int_m_item.cust_funcs,self.int_m_item.total_func_item,
                    self.int_m_item.columntypes))
        
        # Lock all custom items apart from the current selected int_m_item
        self.locked = data.get_lock_states() - self.selected
        if self.int_m_item is not None:
            type_ = self.int_m_item.itemtype
            if type_:
                for count_cmb, cmb in enumerate(self.cmbs):
                    for count_meas, meas in enumerate(cmb):
                        if isinstance(meas, Measurement):
                            for count_meas_item, meas_item in enumerate(meas):
                                path = [count_cmb,count_meas,count_meas_item]
                                if isinstance(meas_item,MeasurementItemCustom):
                                    if self.locked[path] != True and meas_item.itemtype != type_:
                                        self.locked[path] = True
                                    elif meas_item.itemtype == type_:
                                            self.locked[path] = False
                                            
        # Update remarks column
        self.int_m_item.set_remark(self.entry_abstract_remark.get_text())
        
        # Update store from lock states
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
        
        # Update measurements view
        self.measurements_view.update_store()

    def __init__(self, parent, data):
        # Setup variables
        self.parent = parent
        self.data = data
        self.schedule = data.schedule
        self.locked = None
        self.selected = None
        
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
        # Text Entries
        self.entry_abstract_remark = self.builder.get_object("entry_abstract_remark")
        # Setup measurements view
        self.measurements_view = MeasurementsView(self.window, self.data, self.treeview_abstract)
        # Connect toggled signal of measurement view to callback
        self.measurements_view.renderer_toggle.connect("toggled", self.onToggleCellRendererToggle) #TODO

        # Load data
        if self.data is not None:
            self.mitems = data[0]
            self.int_m_item = MeasurementItemCustom(data[1],data[5])
            self.entry_abstract_remark.set_text(self.int_m_item.get_remark())
            self.selected = data.datamodel.LockState(self.mitems)
            self.locked = data.get_lock_states() - self.selected
        
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
