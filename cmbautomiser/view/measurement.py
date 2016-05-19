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

# Setup logger object
log = logging.getLogger(__name__)


class MeasurementsView:
    """Implements a view for display and manipulation of measurement items over a treeview"""
            
    # Callback functions
    
    def onKeyPressTreeview(self, treeview, event):
        """Handle keypress event"""
        if event.keyval == Gdk.KEY_Escape:  # unselect all
            self.tree.get_selection().unselect_all()
    
    def can_have_sub_item(self, level):
        """Return True if sub items can be added with current selection"""
        
        # CMB item can always be added
        if level == 1:
            return True
        
        # Get selection path if selection exists
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0:  # If selection exists
            [model, paths] = selection.get_selected_rows()
            path = paths[0].get_indices()
        # Else get tree node path
        else:
            if len(self.cmbs) > 0: # CMB exists
                p1 = len(self.cmbs)-1
                if len(self.cmbs[p1].items) > 0: # Measurement exists
                    p2 = len(self.cmbs[p1].items)-1
                    if isinstance(self.cmbs[p1][p2], data.measurement.Measurement) and len(self.cmbs[p1][p2].items) > 0: # Measurement item exists
                        p3 = len(self.cmbs[p1][p2].items)-1
                         # Path to measurement item
                        path = [p1,p2,p3]
                    else:
                         # Path to measurement
                        path = [p1,p2]
                else:
                    # Path to CMB
                    path = [p1]
            else:
                # No CMBs exist
                path = []

        # Measurement/Completion
        if level == 2:
            return True if len(path) > 0 else False
        # MEasurementItem
        elif level == 3:
            # If a measurement or measurement item selected
            if len(path) > 1:
                # Return True if measurement is a 'Measurement'
                return True if isinstance(self.cmbs[path[0]][path[1]], data.measurement.Measurement) else False
            elif len(path) == 1:
                # CMB Selected. Return True if last item of CMB is a 'Measurement'
                if len(self.cmbs[path[0]].items) > 0:
                    return True if isinstance(self.cmbs[path[0]].items[-1], data.measurement.Measurement) else False
                else:
                    return False
            else:
                # No CMB exists
                return False
            
    # Public Methods
    
    def set_colour(self, path, color):
        """Sets the colour of item selected by path"""
        if len(path) == 1:
            path_formated = Gtk.TreePath.new_from_string(str(path[0]))
        elif len(path) == 2:
            path_formated = Gtk.TreePath.new_from_string(str(path[0]) + ':' + str(path[1]))
        elif len(path) == 3:
            path_formated = Gtk.TreePath.new_from_string(str(path[0]) + ':' + str(path[1]) + ':' + str(path[2]))
        else:
            return
        path_iter = self.store.get_iter(path_formated)
        self.store.set_value(path_iter, 3, color)

    def add_cmb(self):
        """Add a CMB to measurement view"""
        cmb_name = misc.get_user_input_text(self.parent, "Please input CMB Name", "Add new CMB")
        if cmb_name != None:
            cmb = ['CMB', [cmb_name, []]]
            # Get selection
            selection = self.tree.get_selection()
            if selection.count_selected_rows() != 0:  # If selection exists
                [model, paths] = selection.get_selected_rows()
                self.data.add_cmb_at_node(cmb, paths[0].get_indices()[0])
            else:  # If no selection append at end
                self.data.add_cmb_at_node(cmb, None)
            self.update_store()
        
    def add_measurement(self):
        """Add a Measurement to measurement view"""
        if self.can_have_sub_item(2):
            meas_name = misc.get_user_input_text(self.parent,  "Please input Measurement Date", "Add Measurement Group")
            if meas_name != None:
                meas = ['Measurement', [meas_name, []]]
                # Get selection
                selection = self.tree.get_selection()
                if selection.count_selected_rows() != 0:  # If selection exists
                    [model, paths] = selection.get_selected_rows()
                    path = paths[0].get_indices()
                    self.data.add_measurement_at_node(meas,path)
                else: # If no selection append at end
                    self.data.add_measurement_at_node(meas,None)
                self.update_store()
        else:
            # Return status code for main application interface
            return (misc.CMB_WARNING,"Item not added - 'Measurement Group' can only be added under a 'CMB'")

    def add_completion(self):
        """Add a Completion to measurement view"""
        if self.can_have_sub_item(2):
            compl_name = misc.get_user_input_text(self.parent, "Please input Completion Date", "Add Completion Certificate")
            if compl_name != None:
                compl = ['Completion', [compl_name]]
                # Get selection
                selection = self.tree.get_selection()
                if selection.count_selected_rows() != 0: # If selection exists
                    [model, paths] = selection.get_selected_rows()
                    path = paths[0].get_indices()
                    self.data.add_measurement_at_node(compl,path)
                else: # If no selection append at end
                    self.data.add_measurement_at_node(compl,None)
                self.update_store()
        else:
            # Return status code for main application interface
            return (misc.CMB_WARNING,"Item not added - 'Completion Certificate' can only be added under a 'CMB'")
        
    def add_heading(self):
        """Add a Heading to measurement view"""
        if self.can_have_sub_item(3):
            heading_name = misc.get_user_input_text(self.parent, "Please input Heading", "Add new Item: Heading")
            if heading_name != None:
                # get selection
                selection = self.tree.get_selection()
                heading = ['MeasurementItemHeading', [heading_name]]
                if selection.count_selected_rows() != 0: # if selection exists
                    [model, paths] = selection.get_selected_rows()
                    path = paths[0].get_indices()
                    self.data.add_measurement_item_at_node(heading,path)
                else: # if no selection append at end
                    self.data.add_measurement_item_at_node(heading,None)
                self.update_store()
        else:
            # Return status code for main application interface
            return (misc.CMB_WARNING,"Item not added - 'Heading' can only be added under a 'Measurement Group'")
                    
    def add_custom(self, oldval=None, itemtype=None):
        """Add a Custom item to measurement view"""
        if self.can_have_sub_item(3):
            template = self.data.get_custom_item_template(itemtype)
            dialog = ScheduleDialog(self.parent, self.schedule, *template)
            if oldval is not None: # if edit mode add data
                # Obtain ScheduleDialog model from MeasurementItemCustom model
                schmod = self.data.get_schmod_from_custmod(oldval)
                dialog.set_model(schmod)
                data = dialog.run()
                if data is not None: # if edited
                    # Obtain MeasurementItemCustom model from ScheduleDialog model
                    custmod = self.data.get_custmod_from_schmod(data, oldval, itemtype)
                    return custmod
                else: # if cancel pressed
                    return None
            else: # if normal mode
                data = dialog.run()
                if data is not None:
                    # Obtain custom item model from returned data
                    custmod = self.data.get_custmod_from_schmod(data, None, itemtype)
                    # get selection
                    selection = self.tree.get_selection()
                    if selection.count_selected_rows() != 0: # if selection exists
                        [model, paths] = selection.get_selected_rows()
                        path = paths[0].get_indices()
                        self.data.add_measurement_item_at_node(custmod, path)
                    else: # if no selection append at end
                        self.data.add_measurement_item_at_node(custmod, None)
                    self.update_store()
        else:
            # Return status code for main application interface
            return (misc.CMB_WARNING,"Item not added - Item can only be added under a 'Measurement Group'")

    def add_abstract(self,oldval=None):
        """Add an Abstract item to measurement view"""
        if self.can_have_sub_item(3):
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
                        self.data.add_measurement_item_at_node(model, path)
                    else: # if no selection append at end
                        self.data.add_measurement_item_at_node(model, None)
                    self.update_store()
        else:
            # Return status code for main application interface
            return (misc.CMB_WARNING,"Item not added - 'Abstract of Item' can only be added under a 'Measurement Group'")
        
    def delete_selected_row(self):
        """Delete selected rows"""
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0: # if selection exists
            [model, paths] = selection.get_selected_rows()
            self.data.delete_row_meas(paths[0].get_indices())
            self.update_store()

    def copy_selection(self):
        """Copy selected row to clipboard"""
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0: # if selection exists
            test_string = "MeasurementsView"
            [model, paths] = selection.get_selected_rows()
            path = paths[0].get_indices()
            if len(path) == 1:
                item = self.cmbs[path[0]].get_model()
            elif len(path) == 2:
                item = self.cmbs[path[0]][path[1]].get_model()
            elif len(path) == 3:
                item = self.cmbs[path[0]][path[1]][path[2]].get_model()
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
                        if item[0] == 'CMB':
                            self.data.add_cmb_at_node(item,path[0])
                        elif item[0] in ['Measurement', 'Completion']:
                            self.data.add_measurement_at_node(item,path)
                        elif item[0] in ['MeasurementItemHeading', 'MeasurementItemCustom', 'MeasurementItemAbstract']:
                            self.data.add_measurement_item_at_node(item,path)
                    else:
                        if item[0] == 'CMB':
                            self.data.add_cmb_at_node(item,None)
                        elif item[0] in ['Measurement', 'Completion']:
                            self.data.add_measurement_at_node(item,None)
                        elif item[0] in ['MeasurementItemHeading', 'MeasurementItemCustom', 'MeasurementItemAbstract']:
                            self.data.add_measurement_item_at_node(item,None)
                    self.update_store()
            except:
                log.warning('MeasurementsView - paste_at_selection - No valid data in clipboard')
        else:
            log.warning('MeasurementsView - paste_at_selection - No text on the clipboard')

    def update_store(self, lock_state = None):
        """Update GUI of MeasurementsView from data model while trying to preserve selection
        
            Arguments:
                lock_state: (Optional) Display custom lock state
        """
        if lock_state is None:
            lock_state = self.data.get_lock_states()
                                
        # Get selection
        selection = self.tree.get_selection()
        old_path = []
        if selection.count_selected_rows() != 0: # if selection exists
            [model, paths] = selection.get_selected_rows()
            old_path = paths[0].get_indices()

        # Update StoreView
        self.store.clear()
        for p1, cmb in enumerate(self.cmbs):
            iter_cmb = self.store.append(None,[cmb.get_text(),False,cmb.get_tooltip(),misc.MEAS_COLOR_NORMAL])
            for p2, meas in enumerate(cmb.items):
                iter_meas = self.store.append(iter_cmb,[meas.get_text(),False,meas.get_tooltip(),misc.MEAS_COLOR_NORMAL])
                if isinstance(meas, data.measurement.Measurement):
                    for p3, mitem in enumerate(meas.items):
                        m_flag = lock_state[[p1, p2, p3]]
                        self.store.append(iter_meas,[mitem.get_text(), m_flag, mitem.get_tooltip(),misc.MEAS_COLOR_NORMAL])
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

    def render_selection(self, folder, replacement_dict):
        """Render selected CMB"""
        # get selection
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0 and folder != None: # if selection exists
            # get path of selection
            [model, paths] = selection.get_selected_rows()
            path = paths[0].get_indices()
            code = self.data.render_cmb(folder, replacement_dict, path)
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
                self.data.edit_measurement_item(path,item,newval,oldval)
            elif len(path) == 2:
                item = self.cmbs[path[0]][path[1]]
                if isinstance(item, data.measurement.Measurement):
                    oldval = item.get_date()
                    newval = misc.get_user_input_text(self.parent, "Please input Measurement Date", "Edit Measurement",oldval)
                    self.data.edit_measurement_item(path,item,newval,oldval)
                elif isinstance(item, data.measurement.Completion):
                    oldval = item.get_date()
                    newval = misc.get_user_input_text(self.parent, "Please input Completion Date", "Edit Measurement",oldval)
                    self.data.edit_measurement_item(path,item,newval,oldval)
            elif len(path) == 3:
                item = self.cmbs[path[0]][path[1]][path[2]]
                if isinstance(item, data.measurement.MeasurementItemHeading):
                    oldval = item.get_remark()
                    newval = misc.get_user_input_text(self.parent, "Please input Heading", "Edit Heading",oldval)
                    self.data.edit_measurement_item(path,item,newval,oldval)
                elif isinstance(item, data.measurement.MeasurementItemCustom):
                    oldval = item.get_model()
                    newval = self.add_custom(oldval,item.itemtype)
                    if newval is not None:
                        self.data.edit_measurement_item(path,item,newval,oldval)
                elif isinstance(item, data.measurement.MeasurementItemAbstract):
                    oldval = item.get_model()
                    newval = self.add_abstract(oldval)
                    if newval is not None:
                        self.data.edit_measurement_item(path,item,newval,oldval)
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
                if isinstance(item, data.measurement.MeasurementItemCustom):
                    if item.captions_udata:
                        oldval = item.get_model()
                        olddata = oldval[1][4]
                        # Setup user data dialog
                        newdata = olddata[:]
                        project_settings_dialog = misc.UserEntryDialog(self.parent, 
                                                    'Edit User Data',
                                                    newdata,
                                                    item.captions_udata)
                        # Show user data dialog
                        code = project_settings_dialog.run()
                        # Edit data on success
                        if code:
                            newval = copy.deepcopy(oldval)
                            newval[1][4] = newdata
                            self.data.edit_measurement_item(path, item, newval, oldval)
                        return None
                elif isinstance(item, data.measurement.MeasurementItemAbstract):
                    if item.int_mitem.captions_udata:
                        oldval = item.get_model()
                        olddata = oldval[1][1][1][4]
                        # Setup user data dialog
                        newdata = olddata[:]
                        project_settings_dialog = misc.UserEntryDialog(self.parent, 
                                                    'Edit User Data',
                                                    newdata,
                                                    item.int_mitem.captions_udata)
                        # Show user data dialog
                        code = project_settings_dialog.run()
                        # Edit data on success
                        if code:
                            newval = copy.deepcopy(oldval)
                            newval[1][1][1][4] = newdata
                            self.data.edit_measurement_item(path, item, newval, oldval)
                        return None
        return (misc.CMB_WARNING,'User data not supported')
                    
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
        
        # Derived data
        self.schedule = data.schedule
        self.cmbs = data.cmbs
        
        ## Setup treeview store
        # Item Description, Billed Flag, Tooltip, Colour
        self.store = Gtk.TreeStore(str,bool,str,str)
        # Treeview columns
        self.column_desc = Gtk.TreeViewColumn('Item Description')
        self.column_desc.props.expand = True
        self.column_toggle = Gtk.TreeViewColumn('Billed ?')
        self.column_toggle.props.fixed_width = 100
        self.column_toggle.props.min_width = 100
        # Pack Columns
        self.tree.append_column(self.column_desc)
        self.tree.append_column(self.column_toggle)
        # Treeview renderers
        self.renderer_desc = Gtk.CellRendererText()
        self.renderer_toggle = Gtk.CellRendererToggle()
        # Pack renderers
        self.column_desc.pack_start(self.renderer_desc, True)
        self.column_toggle.pack_start(self.renderer_toggle, True)
        # Add attributes
        self.column_desc.add_attribute(self.renderer_desc, "markup", 0)
        self.column_desc.add_attribute(self.renderer_desc, "background", 3)
        self.column_toggle.add_attribute(self.renderer_toggle, "active", 1)
        self.tree.set_tooltip_column(2)
        # Set model for store
        self.tree.set_model(self.store)

        # Intialise clipboard
        self.clipboard = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD)

        # Connect Callbacks
        self.tree.connect("key-press-event", self.onKeyPressTreeview)
        
        # Update GUI elements according to data
        self.update_store()
        
        
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
        """On toggle clicked"""
        path_obj = Gtk.TreePath.new_from_string(path_str)
        path = path_obj.get_indices()
        
        def select_item(item, path, state = None):
            """Change state of item at path if item is toggleable
            
                Arguments:
                    item: Item selected
                    path: Path to item selected
                    state: Final state of item. Toggles on None"""
            if not isinstance(item, data.measurement.MeasurementItemHeading) and not isinstance(meas_item, data.measurement.MeasurementItemAbstract) and self.locked[path] != True:
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

    # Class methods

    def get_model(self):
        self.update_store()
        if self.int_mitem != None:
            model = [self.mitems, self.int_mitem.get_model()] 
            return ['MeasurementItemAbstract', model]
        else:
            return None

    def clear(self):
        self.data = None

    def update_store(self):
        
        # Update mitems
        self.mitems = self.selected.get_paths()

        # Update self.int_mitem
        if self.mitems:
            p = self.mitems[0]
            item = self.data.cmbs[p[0]][p[1]][p[2]]
            type_ = item.itemtype
            if self.int_mitem is None:
                self.int_mitem = data.measurement.MeasurementItemCustom(item.get_model()[1], type_)
            # Populate values
            self.int_mitem.records = []
            for path in self.mitems:
                item = self.data.cmbs[path[0]][path[1]][path[2]]
                values = item.export_abstract(item.records, item.user_data)
                # Make dicionary magic
                cmbbf = 'ref:meas:'+ str(path) + ':1'
                label = 'ref:abs:'+ str(path) + ':1'
                values[0] = r'Qty B/F MB.No.\emph{\nameref{' + cmbbf + r'} Pg.No. \pageref{' + cmbbf \
                            + r'}}\phantomsection\label{' + label + '}'
                self.int_mitem.append_record(data.measurement.RecordCustom(values,
                    self.int_mitem.cust_funcs,self.int_mitem.total_func_item,
                    self.int_mitem.columntypes))
        else:
            self.int_mitem = None
            
        # Lock all custom items apart from the current selected int_mitem
        self.locked = self.data.get_lock_states() - self.initial_selected
        for count_cmb, cmb in enumerate(self.cmbs):
            for count_meas, meas in enumerate(cmb):
                if isinstance(meas, data.measurement.Measurement):
                    for count_meas_item, meas_item in enumerate(meas):
                        path = [count_cmb,count_meas,count_meas_item]
                        if isinstance(meas_item,data.measurement.MeasurementItemCustom):
                            # Lock if non abstratable item
                            if meas_item.export_abstract == None:
                                self.locked[path] = True
                            # If one item selected, lock all other itemtypes
                            elif self.int_mitem is not None:
                                type_ = self.int_mitem.itemtype                               
                                if meas_item.itemtype != type_:
                                    self.locked[path] = True
                                            
        if self.int_mitem is not None:
            # Update remarks column
            self.int_mitem.set_remark(self.entry_abstract_remark.get_text())
        
        # Update store from lock states
        self.measurements_view.update_store(self.selected)
        for count_cmb, cmb in enumerate(self.data.cmbs):
            flag_cmb = False # Check if there exists an abstractable item
            for count_meas, meas in enumerate(cmb):
                if isinstance(meas, data.measurement.Measurement):
                    flag_meas = False # Check if there exists an abstractable item
                    for count_meas_item, meas_item in enumerate(meas):
                        path = [count_cmb, count_meas, count_meas_item]
                        if self.locked[path]:
                            # Apply color to locked items
                            self.measurements_view.set_colour(path, misc.MEAS_COLOR_LOCKED)
                        elif isinstance(meas_item, data.measurement.MeasurementItemHeading) or isinstance(meas_item, data.measurement.MeasurementItemAbstract):
                            self.measurements_view.set_colour(path, misc.MEAS_COLOR_LOCKED)
                        elif self.selected[path]:
                            # Apply color to selected items
                            self.measurements_view.set_colour(path, misc.MEAS_COLOR_SELECTED)
                            flag_meas = True
                            flag_cmb = True
                        else:
                            # Non selected unlocked items
                            flag_meas = True
                            flag_cmb = True
                    # If no abstractable item exists, color locked
                    if not flag_meas:
                        path = [count_cmb, count_meas]
                        self.measurements_view.set_colour(path, misc.MEAS_COLOR_LOCKED)
                else:
                    path = [count_cmb, count_meas]
                    # Apply color to locked items
                    self.measurements_view.set_colour(path, misc.MEAS_COLOR_LOCKED)
            if not flag_meas:
                path = [count_cmb]
                self.measurements_view.set_colour(path, misc.MEAS_COLOR_LOCKED)

    def __init__(self, parent, datamodel, model):
        # Setup variables
        self.parent = parent
        self.data = datamodel
        self.selected = data.datamodel.LockState()
        # Derived data
        self.schedule = self.data.schedule
        self.cmbs = self.data.cmbs
        self.locked = self.data.get_lock_states()
        self.initial_selected = self.data.get_lock_states()
        # Private variables
        self.mitems = []
        self.int_mitem = None

        # Setup dialog window
        self.builder = Gtk.Builder()
        self.builder.add_from_file(misc.abs_path("interface","abstractdialog.glade"))
        self.window = self.builder.get_object("dialog")
        self.window.set_transient_for(self.parent)
        self.window.set_default_size(1000,500)
        self.builder.connect_signals(self)
        # Get required objects
        self.treeview_abstract = self.builder.get_object("treeview_abstract")
        # Text Entries
        self.entry_abstract_remark = self.builder.get_object("entry_abstract_remark")
        # Setup measurements view
        self.measurements_view = MeasurementsView(self.window, self.data, self.treeview_abstract)
        # Connect toggled signal of measurement view to callback
        self.measurements_view.renderer_toggle.connect("toggled", self.onToggleCellRendererToggle)

        # Load data
        if model is not None:
            if model[0] == 'MeasurementItemAbstract':
                self.mitems = model[1][0]
                self.int_mitem = data.measurement.MeasurementItemCustom(model[1][1][1], model[1][1][1][5])
                self.entry_abstract_remark.set_text(self.int_mitem.get_remark())
                self.selected = data.datamodel.LockState(self.mitems)
                self.initial_selected = data.datamodel.LockState(self.mitems)
                self.locked = self.data.get_lock_states() - self.selected
        
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

