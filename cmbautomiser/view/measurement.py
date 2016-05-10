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

from gi.repository import Gtk, Gdk, GLib
from undo import undoable

import pickle, os.path, copy, logging

# local files import
from __main__ import misc
from .scheduledialog import ScheduleDialog
from .abstractdialog import AbstractDialog

# Setup logger object
log = logging.getLogger(__name__)

# Measurements object
class MeasurementsView:
            
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
        

        item = MeasurementItemAbstract(None)
        
        if oldval is not None: # if edit mode add data
            dialog = AbstractDialog(self.parent,oldval, self, self.schedule)
            data = dialog.run()
            if data is not None: # if edited
                return data
            else: # if cancel pressed
                return None
        else: # if normal mode
            dialog = AbstractDialog(toplevel,[[],None,None], self, self.schedule)
            data = dialog.run()
            if data is not None:
                item.set_model(data)
                # get selection
                selection = self.tree.get_selection()
                if selection.count_selected_rows() != 0: # if selection exists
                    [model, paths] = selection.get_selected_rows()
                    path = paths[0].get_indices()
                    self.add_measurement_item_at_node(item,path)
                else: # if no selection append at end
                    self.add_measurement_item_at_node(item,None)
        
    def delete_selected_row(self):
        """Delete selected rows"""
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0: # if selection exists
            [model, paths] = selection.get_selected_rows()
            self.delete_row(paths[0].get_indices())

    def copy_selection(self):
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0: # if selection exists
            [model, paths] = selection.get_selected_rows()
            path = paths[0].get_indices()
            if len(path) == 1:
                item = self.cmbs[path[0]]
            elif len(path) == 2:
                item = self.cmbs[path[0]][path[1]]
            elif len(path) == 3:
                item = self.cmbs[path[0]][path[1]][path[2]]
            text = pickle.dumps(item) # dump item as text
            self.clipboard.set_text(text,-1) # push to clipboard
        else: # if no selection
            log.warning("No items selected to copy")

    def paste_at_selection(self):
        text = self.clipboard.wait_for_text() # get text from clipboard
        if text != None:
            try:
                item = pickle.loads(text) # recover item from string
                selection = self.tree.get_selection()
                if selection.count_selected_rows() != 0: # if selection exists
                    [model, paths] = selection.get_selected_rows()
                    path = paths[0].get_indices()
                    if isinstance(item,Cmb):
                        self.add_cmb_at_node(item,path[0])
                    elif isinstance(item,Measurement) or isinstance(item,Completion):
                        self.add_measurement_at_node(item,path)
                    elif isinstance(item,MeasurementItem):
                        self.add_measurement_item_at_node(item,path)
                else:
                    if isinstance(item,Cmb):
                        self.add_cmb_at_node(item,None)
                    elif isinstance(item,Measurement) or isinstance(item,Completion):
                        self.add_measurement_at_node(item,None)
                    elif isinstance(item,MeasurementItem):
                        self.add_measurement_item_at_node(item,None)
            except:
                log.warning('No valid data in clipboard')
        else:
            log.warning("No text on the clipboard.")

    def update_store(self):

        # Update all measurement flags
        ManageResourses().update_billed_flags()

        # Get selection
        selection = self.tree.get_selection()
        old_path = []
        if selection.count_selected_rows() != 0: # if selection exists
            [model, paths] = selection.get_selected_rows()
            old_path = paths[0].get_indices()

        # Update StoreView
        self.store.clear()
        for cmb in self.cmbs:
            iter_cmb = self.store.append(None,[cmb.get_text(),False,cmb.get_tooltip(),MEAS_COLOR_NORMAL])
            for meas in cmb.items:
                iter_meas = self.store.append(iter_cmb,[meas.get_text(),False,meas.get_tooltip(),MEAS_COLOR_NORMAL])
                if isinstance(meas,Measurement):
                    for mitem in meas.items:
                        self.store.append(iter_meas,[mitem.get_text(),mitem.get_billed_flag(),mitem.get_tooltip(),MEAS_COLOR_NORMAL])
                elif isinstance(meas,Completion):
                    pass
        self.tree.expand_all()

        # Set selection
        if old_path != []:
            if len(old_path) > 0 and len(self.cmbs) > 0:
                if len(self.cmbs) <= old_path[0]:
                    old_path[0] = len(self.cmbs)-1
                cmb = self.cmbs[old_path[0]]
                if len(old_path) > 1 and len(cmb.items) > 0:
                    if len(cmb.items) <= old_path[1]:
                        old_path[1] = len(cmb.items)-1
                    item = cmb.items[old_path[1]]
                    if isinstance(item,Measurement) and len(old_path) == 3 and len(item.items) > 0:
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

    def render_selection(self,folder,replacement_dict,bills):
        # get selection
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0 and folder != None: # if selection exists
            # get path of selection
            [model, paths] = selection.get_selected_rows()
            path = paths[0].get_indices()
            code = self.render(folder,replacement_dict,bills,path)
            return code
        else:
            return (CMB_WARNING,'Please select a CMB for rendering')

    def render(self,folder,replacement_dict,bills,path,recursive = True):
        # fill in latex buffer
        latex_buffer = self.cmbs[path[0]].get_latex_buffer([path[0]])

        # make global variables replacements
        latex_buffer = replace_all(latex_buffer,replacement_dict)

        # include linked bills
        replacement_dict_bills = {}
        external_docs = ''
        for count,bill in enumerate(bills):
            external_docs += '\externaldocument{abs_' + str(count+1) + '}\n'
        replacement_dict_bills['$cmbexternaldocs$'] = external_docs
        latex_buffer = replace_all_vanilla(latex_buffer,replacement_dict_bills)

        # write output
        filename = posix_path(folder,'cmb_' + str(path[0]+1) + '.tex')
        file_latex = open(filename,'w')
        file_latex.write(latex_buffer)
        file_latex.close()

        # run latex on file and dependencies

        # run on all bills refering cmb
        if recursive: # if recursive call
            for bill_count,bill in enumerate(bills):
                if path[0] in bill.cmb_ref:
                    code = bill.bill_view.render(folder,replacement_dict,[bill_count],False)
                    if code[0] == CMB_ERROR:
                        return code
        # run latex on file
        code = run_latex(posix_path(folder),filename)
        if code == CMB_ERROR:
            return (CMB_ERROR,'Rendering of CMB No.' + self.cmbs[path[0]].get_name() + ' failed')

        # run on all bills refering cmb again to rebuild indexes on recursive run
        if recursive: # if recursive call
            for bill_count,bill in enumerate(bills):
                if path[0] in bill.cmb_ref:
                    code = bill.bill_view.render(folder,replacement_dict,[bill_count],False)
                    if code[0] == CMB_ERROR:
                        return code

        return (CMB_INFO,'CMB No.' + self.cmbs[path[0]].get_name() + ' rendered successfully')
            
    def edit_selected_row(self):
        toplevel = self.tree.get_toplevel() # get current top level window
        # get selection
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0: # if selection exists
            [model, paths] = selection.get_selected_rows()
            path = paths[0].get_indices()
            if len(path) == 1:
                item = self.cmbs[path[0]]
                oldval = item.get_name()
                newval = misc.get_user_input_text(toplevel, "Please input CMB Name", "Edit CMB",oldval)
                self.edit_item(path,item,newval,oldval)
            elif len(path) == 2:
                item = self.cmbs[path[0]][path[1]]
                if isinstance(item,Measurement):
                    oldval = item.get_date()
                    newval = misc.get_user_input_text(toplevel, "Please input Measurement Date", "Edit Measurement",oldval)
                    self.edit_item(path,item,newval,oldval)
                elif isinstance(item,Completion):
                    oldval = item.get_date()
                    newval = misc.get_user_input_text(toplevel, "Please input Completion Date", "Edit Measurement",oldval)
                    self.edit_item(path,item,newval,oldval)
            elif len(path) == 3:
                item = self.cmbs[path[0]][path[1]][path[2]]
                if isinstance(item,MeasurementItemHeading):
                    oldval = item.get_remark()
                    newval = misc.get_user_input_text(toplevel, "Please input Heading", "Edit Heading",oldval)
                    self.edit_item(path,item,newval,oldval)
                elif isinstance(item,MeasurementItemNLBH):
                    oldval = item.get_model()
                    newval = self.add_nlbh(oldval)
                    self.edit_item(path,item,newval,oldval)
                elif isinstance(item,MeasurementItemLLLLL):
                    oldval = item.get_model()
                    newval = self.add_lllll(oldval)
                    self.edit_item(path,item,newval,oldval)
                elif isinstance(item,MeasurementItemNNNNNNNN):
                    oldval = item.get_model()
                    newval = self.add_nnnnnnnn(oldval)
                    self.edit_item(path,item,newval,oldval)
                elif isinstance(item,MeasurementItemnnnnnT):
                    oldval = item.get_model()
                    newval = self.add_nnnnnt(oldval)
                    self.edit_item(path,item,newval,oldval)
                elif isinstance(item,MeasurementItemCustom):
                    oldval = item.get_model()
                    newval = self.add_custom(oldval,item.itemtype)
                    self.edit_item(path,item,newval,oldval)
                elif isinstance(item,MeasurementItemAbstract):
                    oldval = item.get_model()
                    newval = self.add_abstract(oldval)
                    self.edit_item(path,item,newval,oldval)

    def edit_user_data(self,oldval,itemtype):

        item = MeasurementItemCustom(None,itemtype)
        captions_udata = item.captions_udata
        columntypes_udata = item.columntypes_udata
        toplevel = self.tree.get_toplevel() # get current top level window
        newdata = []
        entrys = []

        dialogWindow = Gtk.MessageDialog(toplevel,
                              Gtk.DialogFlags.MODAL | Gtk.DialogFlags.DESTROY_WITH_PARENT,
                              Gtk.MessageType.QUESTION,
                              Gtk.ButtonsType.OK_CANCEL,
                              'Edit User Data')

        dialogWindow.set_transient_for(toplevel)
        dialogWindow.set_title('Edit User Data')
        dialogWindow.set_default_response(Gtk.ResponseType.OK)

        # Pack Dialog
        dialogBox = dialogWindow.get_content_area()
        grid = Gtk.Grid()
        grid.set_column_spacing(5)
        grid.set_row_spacing(5)
        grid.set_border_width(5)
        dialogBox.add(grid)
        for caption, columntype in zip(captions_udata, columntypes_udata):
            userLabel = Gtk.Label(caption)
            userEntry = Gtk.Entry()
            userEntry.set_activates_default(True)
            grid.attach_next_to(userLabel, None, Gtk.PositionType.BOTTOM, 1, 1)
            grid.attach_next_to(userEntry, userLabel, Gtk.PositionType.RIGHT, 1, 1)
            entrys.append(userEntry)

        # Add data
        if oldval is not None:
            for data_str,userEntry in zip(oldval,entrys):
                userEntry.set_text(data_str)

        # Run dialog
        dialogWindow.show_all()
        response = dialogWindow.run()

        # Get formated text
        for userEntry,column_type in zip(entrys,columntypes_udata):
            cell = userEntry.get_text()
            try:  # try evaluating string
                if column_type == MEAS_DESC:
                    cell_formated = str(cell)
                elif column_type == MEAS_L:
                    cell_formated = str(float(cell))
                elif column_type == MEAS_NO:
                    cell_formated = str(int(cell))
                else:
                    cell_formated = ''
            except:
                cell_formated = ''
            newdata.append(cell_formated)
        dialogWindow.destroy()
        if response == Gtk.ResponseType.OK:
            return newdata
        else:
            return None

    def edit_selected_properties(self):
        # get selection
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0: # if selection exists
            [model, paths] = selection.get_selected_rows()
            path = paths[0].get_indices()
            if len(path) == 3:
                item = self.cmbs[path[0]][path[1]][path[2]]
                if isinstance(item,MeasurementItemCustom):
                    oldval = item.get_model()
                    olddata = oldval[5]
                    newdata = self.edit_user_data(olddata,item.itemtype)
                    if newdata is not None:
                        newval = copy.deepcopy(oldval)
                        newval[5] = newdata
                        self.edit_item(path,item,newval,oldval)
                    return None
                if isinstance(item,MeasurementItemAbstract):
                    oldval = item.get_model()
                    olddata = oldval[1][5]
                    newdata = self.edit_user_data(olddata,oldval[2])
                    if newdata is not None:
                        newval = copy.deepcopy(oldval)
                        newval[1][5] = newdata
                        self.edit_item(path,item,newval,oldval)
                    return None
        return (CMB_WARNING,'User data not supported')

    @undoable
    def edit_item(self,path,item,newval,oldval):
        if newval != None:
            if len(path) == 1:
                item.set_name(newval)
            elif len(path) == 2:
                if isinstance(item,Measurement):
                    item.set_date(newval)
                elif isinstance(item,Completion):
                    item.set_date(newval)
            elif len(path) == 3:
                if isinstance(item,MeasurementItemHeading):
                    item.set_remark(newval)
                else:
                    item.set_model(newval)
            self.update_store()
        
        yield "Edit measurement items at '{}'".format(path)
        # Undo action
        if oldval != None and newval != None:
            if len(path) == 1:
                item.set_name(oldval)
            elif len(path) == 2:
                if isinstance(item,Measurement):
                    item.set_date(oldval)
                elif isinstance(item,Completion):
                    item.set_date(oldval)
            elif len(path) == 3:
                if isinstance(item,MeasurementItemHeading):
                    item.set_remark(oldval)
                else:
                    item.set_model(oldval)
            self.update_store()
                    
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
        self.bills = data_model.cmbs
        
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
        self.clipboard = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD) # initialise clipboard

        # Connect Callbacks
        self.tree.connect("key-press-event", self.onKeyPressTreeview)
        
        # Update GUI elements according to data
        self.update_store()
