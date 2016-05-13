#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#  measuremets.py
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

import pickle, codecs, os.path, copy, logging

from gi.repository import Gtk, Gdk, GLib
from undo import undoable

# local files import
from __main__ import misc, data
from .scheduledialog import ScheduleDialog
from .abstractdialog import AbstractDialog

# Setup logger object
log = logging.getLogger(__name__)


class MeasurementsView:
    """Implements a view for display and manipulation of measurement items over a treeview"""
            
    # Callback functions
    
    def onKeyPressTreeview(self, treeview, event):
        """Handle keypress event"""
        if event.keyval == Gdk.KEY_Escape:  # unselect all
            self.tree.get_selection().unselect_all()
            
    # Public Methods
    
    def set_colour(self, path, color):
        """Sets the colour of item selected by path"""
        if len(path == 1):
            path_formated = Gtk.TreePath.new_from_string(str(path[0]))
        elif len(path == 2):
            path_formated = Gtk.TreePath.new_from_string(str(path[0]) + ':' + str(path[1]))
        elif len(path == 3):
            path_formated = Gtk.TreePath.new_from_string(str(path[0]) + ':' + str(path[1]) + ':' + str(path[2]))
        else
            return
        path_iter = self.store.get_iter(path_formated)
        self.measurements_view.store.set_value(path_iter, 3, color)

    def add_cmb(self):
        """Add a CMB to measurement view"""
        cmb_name = misc.get_user_input_text(self.parent, "Please input CMB Name", "Add new CMB")
        if cmb_name != None:
            cmb = ['CMB', [cmb_name, []]]
            # Get selection
            selection = self.tree.get_selection()
            if selection.count_selected_rows() != 0:  # If selection exists
                [model, paths] = selection.get_selected_rows()
                self.add_cmb_at_node(cmb, paths[0].get_indices()[0])
            else:  # If no selection append at end
                self.add_cmb_at_node(cmb, None)
            self.update_store()
        
    def add_measurement(self):
        """Add a Measurement to measurement view"""
        meas_name = misc.get_user_input_text(self.parent, "Please input Measurement Date", "Add new Measurement")
        if meas_name != None:
            meas = ['Measurement', [meas_name, []]]
            # Get selection
            selection = self.tree.get_selection()
            if selection.count_selected_rows() != 0:  # If selection exists
                [model, paths] = selection.get_selected_rows()
                path = paths[0].get_indices()
                self.add_measurement_at_node(meas,path)
            else: # If no selection append at end
                self.add_measurement_at_node(meas,None)
            self.update_store()

    def add_completion(self):
        """Add a Completion to measurement view"""
        compl_name = misc.get_user_input_text(self.parent, "Please input Completion Date", "Add Completion")
        if compl_name != None:
            compl = ['Completion', [compl_name]]
            # Get selection
            selection = self.tree.get_selection()
            if selection.count_selected_rows() != 0: # If selection exists
                [model, paths] = selection.get_selected_rows()
                path = paths[0].get_indices()
                self.add_measurement_at_node(compl,path)
            else: # If no selection append at end
                self.add_measurement_at_node(compl,None)
            self.update_store()
        
    def add_heading(self):
        """Add a Heading to measurement view"""
        heading_name = misc.get_user_input_text(self.parent, "Please input Heading", "Add new Item: Heading")
        
        if heading_name != None:
            # get selection
            selection = self.tree.get_selection()
            heading = ['MeasurementItemHeading', [heading_name]]
            if selection.count_selected_rows() != 0: # if selection exists
                [model, paths] = selection.get_selected_rows()
                path = paths[0].get_indices()
                self.add_measurement_item_at_node(heading,path)
            else: # if no selection append at end
                self.add_measurement_item_at_node(heading,None)
            self.update_store()
                    
    def add_custom(self, oldval=None, itemtype=None):
        """Add a Custom item to measurement view"""
        template = self.data.get_custom_item_template(itemtype)
        dialog = ScheduleDialog(self.parent, self.schedule, *template)

        if oldval is not None: # if edit mode add data
            # Obtain ScheduleDialog model from MeasurementItemCustom model
            schmod = data.get_schmod_from_custmod(oldval)
            dialog.set_model(schmod)
            data = dialog.run()
            if data is not None: # if edited
                # Obtain MeasurementItemCustom model from ScheduleDialog model
                custmod = data.get_custmod_from_schmod(data, oldval, itemtype)
                return custmod
            else: # if cancel pressed
                return None
        else: # if normal mode
            data = dialog.run()
            if data is not None:
                custmod = data.get_custmod_from_schmod(data, None, itemtype)
                # get selection
                selection = self.tree.get_selection()
                if selection.count_selected_rows() != 0: # if selection exists
                    [model, paths] = selection.get_selected_rows()
                    path = paths[0].get_indices()
                    self.add_measurement_item_at_node(custmod, path)
                else: # if no selection append at end
                    self.add_measurement_item_at_node(custmod, None)

    def add_abstract(self,oldval=None):
        """Add an Abstract item to measurement view"""
        if oldval is not None: # if edit mode add data
            dialog = AbstractDialog(self.parent, self.data, oldval)
            model = dialog.run()
            if model is not None: # if edited
                return model
            else: # if cancel pressed
                return None
        else: # if normal mode
            dialog = AbstractDialog(self.parent, self.data, None)
            model = dialog.run()
            if model is not None:
                # Get selection
                selection = self.tree.get_selection()
                if selection.count_selected_rows() != 0: # if selection exists
                    [model, paths] = selection.get_selected_rows()
                    path = paths[0].get_indices()
                    self.add_measurement_item_at_node(model, path)
                else: # if no selection append at end
                    self.add_measurement_item_at_node(model, None)
        
    def delete_selected_row(self):
        """Delete selected rows"""
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0: # if selection exists
            [model, paths] = selection.get_selected_rows()
            self.delete_row(paths[0].get_indices())

    def copy_selection(self):
        """Copy selected row to clipboard"""
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0: # if selection exists
            test_string = "MeasurementsView"
            [model, paths] = selection.get_selected_rows()
            path = paths[0].get_indices()
            if len(path) == 1:
                item = self.cmbs[path[0]]
            elif len(path) == 2:
                item = self.cmbs[path[0]][path[1]]
            elif len(path) == 3:
                item = self.cmbs[path[0]][path[1]][path[2]]
            text = codecs.encode(pickle.dumps([test_string, item]), "base64").decode() # dump item as text
            self.clipboard.set_text(text,-1) # push to clipboard
        else: # if no selection
            log.warning("MeasurementsView - copy_selection - No items selected to copy")

    def paste_at_selection(self):
        """Paste copied item at selected row"""
        text = self.clipboard.wait_for_text() # get text from clipboard
        if text != None:
            test_string = "MeasurementsView"
            try:
                itemlist = pickle.loads(codecs.decode(text.encode(), "base64"))  # recover item from string
                if itemlist[0] == test_string:
                    item = itemlist[1]
                    selection = self.tree.get_selection()
                    if selection.count_selected_rows() != 0: # if selection exists
                        [model, paths] = selection.get_selected_rows()
                        path = paths[0].get_indices()
                        if isinstance(item, data.measurement.Cmb):
                            self.add_cmb_at_node(item,path[0])
                        elif isinstance(item, data.measurement.Measurement) or isinstance(item, data.measurement.Completion):
                            self.add_measurement_at_node(item,path)
                        elif isinstance(item, data.measurement.MeasurementItem):
                            self.add_measurement_item_at_node(item,path)
                    else:
                        if isinstance(item, data.measurement.Cmb):
                            self.add_cmb_at_node(item,None)
                        elif isinstance(item, data.measurement.Measurement) or isinstance(item, data.measurement.Completion):
                            self.add_measurement_at_node(item,None)
                        elif isinstance(item, data.measurement.MeasurementItem):
                            self.add_measurement_item_at_node(item,None)
            except:
                log.warning('MeasurementsView - paste_at_selection - No valid data in clipboard')
        else:
            log.warning('MeasurementsView - paste_at_selection - No text on the clipboard')

    def update_store(self):
        """Update GUI of MeasurementsView from data model while trying to preserve selection"""
        # Get selection
        selection = self.tree.get_selection()
        old_path = []
        if selection.count_selected_rows() != 0: # if selection exists
            [model, paths] = selection.get_selected_rows()
            old_path = paths[0].get_indices()

        # Update StoreView
        self.store.clear()
        for cmb in self.cmbs:
            iter_cmb = self.store.append(None,[cmb.get_text(),False,cmb.get_tooltip(),misc.MEAS_COLOR_NORMAL])
            for meas in cmb.items:
                iter_meas = self.store.append(iter_cmb,[meas.get_text(),False,meas.get_tooltip(),misc.MEAS_COLOR_NORMAL])
                if isinstance(meas, data.measurement.Measurement):
                    for mitem in meas.items:
                        self.store.append(iter_meas,[mitem.get_text(),mitem.get_billed_flag(),mitem.get_tooltip(),misc.MEAS_COLOR_NORMAL])
                elif isinstance(meas, data.measurement.Completion):
                    pass
        self.tree.expand_all()

        # Set selection to the nearest item that was selected
        if old_path != []:
            if len(old_path) > 0 and len(self.cmbs) > 0:
                if len(self.cmbs) <= old_path[0]:
                    old_path[0] = len(self.cmbs)-1
                cmb = self.cmbs[old_path[0]]
                if len(old_path) > 1 and len(cmb.items) > 0:
                    if len(cmb.items) <= old_path[1]:
                        old_path[1] = len(cmb.items)-1
                    item = cmb.items[old_path[1]]
                    if isinstance(item, data.measurement.Measurement) and len(old_path) == 3 and len(item.items) > 0:
                        if len(item.items) <= old_path[2]:
                            old_path[2] = len(item.items)-1
                    else:
                        old_path = old_path[0:2]
                else:
                    old_path = [old_path[0]]
            else:
                old_path = []

            if old_path != []:
                path = Gtk.TreePath.new_from_indices(old_path)
                self.tree.set_cursor(path)

    def render_selection(self, folder, replacement_dict, bills):
        """Render selected CMB"""
        # get selection
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0 and folder != None: # if selection exists
            # get path of selection
            [model, paths] = selection.get_selected_rows()
            path = paths[0].get_indices()
            code = data.render_cmb(folder, replacement_dict, path)
            # Return status code for main application interface
            return code
        else:
            # Return status code for main application interface
            return (misc.CMB_WARNING,'Please select a CMB for rendering')
            
    def edit_selected_row(self):
        """Edit selected Row"""
        # Get selection
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0: # if selection exists
            [model, paths] = selection.get_selected_rows()
            path = paths[0].get_indices()
            if len(path) == 1:
                item = self.cmbs[path[0]]
                oldval = item.get_name()
                newval = misc.get_user_input_text(self.parent, "Please input CMB Name", "Edit CMB",oldval)
                data.edit_measurement_item(path,item,newval,oldval)
            elif len(path) == 2:
                item = self.cmbs[path[0]][path[1]]
                if isinstance(item, data.measurement.Measurement):
                    oldval = item.get_date()
                    newval = misc.get_user_input_text(self.parent, "Please input Measurement Date", "Edit Measurement",oldval)
                    self.edit_measurement_item(path,item,newval,oldval)
                elif isinstance(item, data.measurement.Completion):
                    oldval = item.get_date()
                    newval = misc.get_user_input_text(self.parent, "Please input Completion Date", "Edit Measurement",oldval)
                    self.edit_measurement_item(path,item,newval,oldval)
            elif len(path) == 3:
                item = self.cmbs[path[0]][path[1]][path[2]]
                if isinstance(item, data.measurement.MeasurementItemHeading):
                    oldval = item.get_remark()
                    newval = misc.get_user_input_text(self.parent, "Please input Heading", "Edit Heading",oldval)
                    self.edit_measurement_item(path,item,newval,oldval)
                elif isinstance(item, data.measurement.MeasurementItemCustom):
                    oldval = item.get_model()
                    newval = self.add_custom(oldval,item.itemtype)
                    self.edit_measurement_item(path,item,newval,oldval)
                elif isinstance(item, data.measurement.MeasurementItemAbstract):
                    oldval = item.get_model()
                    newval = self.add_abstract(oldval)
                    self.edit_measurement_item(path,item,newval,oldval)
            # Update GUI
            self.update_store()

    def edit_selected_properties(self):
        """Edit user data of selected item"""
        # Get Selection
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0: # if selection exists
            [model, paths] = selection.get_selected_rows()
            path = paths[0].get_indices()
            if len(path) == 3:
                item = self.cmbs[path[0]][path[1]][path[2]]
                if isinstance(item, MeasurementItemCustom):
                    oldval = item.get_model()
                    olddata = oldval[1][4]
                    # Setup user data dialog
                    newdata = olddata[:]
                    project_settings_dialog = misc.UserEntryDialog(self.parent, 
                                                  'Edit User Data',
                                                  newval,
                                                  item.captions_udata)
                    # Show user data dialog
                    code = project_settings_dialog.run()
                    # Edit data on success
                    if code:
                        newval = copy.deepcopy(oldval)
                        newval[1][4] = newdata
                        self.edit_item(path, item, newval, oldval)
                    return None
                if isinstance(item, MeasurementItemAbstract):
                    oldval = item.get_model()
                    olddata = oldval[1][1][4]
                    newdata = self.edit_user_data(olddata,oldval[2])
                    if newdata is not None:
                        newval = copy.deepcopy(oldval)
                        newval[1][1][4] = newdata
                        self.edit_item(path, item, newval, oldval)
                    return None
        return (CMB_WARNING,'User data not supported')
                    
    def __init__(self, parent, data, tree):
        """Initialise MeasurementsView class
        
            Arguments:
                parent: Parent widget (Main window)
                data: Main data model
                tree: Treeview for implementing MeasurementsView
        """
        self.parent = parent        
        self.tree = tree
        self.data = data
        self.schedule = data_model.schedule
        self.cmbs = data_model.cmbs
        
        ## Setup treeview store
        # Item Description, Billed Flag, Tooltip, Colour
        self.store = Gtk.ListStore([str,bool,str,str])
        # Treeview columns
        self.column_desc = Gtk.TreeViewColumn('Item Description')
        self.column_desc.props.expand = True
        self.column_toggle = Gtk.TreeViewColumn('Billed ?')
        self.column_toggle.props.fixed_width = 150
        self.column_toggle.props.min_width = 150
        # Treeview renderers
        self.renderer_desc = Gtk.CellRendererText()
        self.renderer_toggle = Gtk.CellRendererToggle()
        # Pack renderers
        self.column_desc.pack_start(self.renderer_desc, True)
        self.column_toggle.pack_start(self.renderer_toggle, True)
        # Add attributes
        self.column_desc.add_attribute(self.renderer_desc, "text", 0)
        self.column_toggle.add_attribute(self.renderer_toggle, "active", 1)
        # Set model for store
        self.tree.set_model(self.store)

        # Intialise clipboard
        self.clipboard = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD)

        # Connect Callbacks
        self.tree.connect("key-press-event", self.onKeyPressTreeview)
        
        # Update GUI elements according to data
        self.update_store()
