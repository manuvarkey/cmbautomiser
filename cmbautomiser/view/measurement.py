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
import misc
from data.schedule import *
from view.scheduledialog import ScheduleDialog

from misc import *

# Setup logger object
log = logging.getLogger(__name__)

# Measurements object
class MeasurementsView:
            
    # Callback functions
    
    def onKeyPressTreeview(self, treeview, event):
        if event.keyval == Gdk.KEY_Escape:  # unselect all
            self.tree.get_selection().unselect_all()

    def get_data_object(self):
        return self.cmbs

    def set_data_object(self,schedule,data):
        del self.cmbs
        self.cmbs = data
        for cmb in self.cmbs: # required since cutom object can be dynamic
            for meas in cmb:
                if not isinstance(meas,Completion):
                    for mitem in meas:
                        if isinstance(mitem,MeasurementItemCustom) or isinstance(mitem,MeasurementItemAbstract):
                            billedflag = mitem.get_billed_flag() # keep billed flag info. Not copied with set_modal
                            mitem.set_model(mitem.get_model()) # refresh changes on disk
                            mitem.set_billed_flag(billedflag)
        self.schedule = schedule
        self.update_store()

    def clear(self):
        self.cmbs = []
        self.update_store()

    def add_cmb(self):
        toplevel = self.tree.get_toplevel() # get current top level window
        cmb_name = misc.get_user_input_text(toplevel, "Please input CMB Name", "Add new CMB")
        
        if cmb_name != None:
            cmb = Cmb(cmb_name)
            
            # get selection
            selection = self.tree.get_selection()
            if selection.count_selected_rows() != 0: # if selection exists
                [model, paths] = selection.get_selected_rows()
                self.add_cmb_at_node(cmb,paths[0].get_indices()[0])
            else: # if no selection append at end
                self.add_cmb_at_node(cmb,None)
        
    @undoable
    def add_cmb_at_node(self,cmb,row):
        if row != None:
            self.cmbs.insert(row,cmb)
            row_delete = row
        else:
            self.cmbs.append(cmb)
            row_delete = len(self.cmbs) - 1
        self.update_store()

        yield "Add CMB at '{}'".format([row])
        # Undo action
        self.delete_row([row_delete])
        
    def add_measurement(self):
        toplevel = self.tree.get_toplevel() # get current top level window
        meas_name = misc.get_user_input_text(toplevel, "Please input Measurement Date", "Add new Measurement")
        
        if meas_name != None:
            meas = Measurement(meas_name)
            # get selection
            selection = self.tree.get_selection()
            if selection.count_selected_rows() != 0: # if selection exists
                [model, paths] = selection.get_selected_rows()
                path = paths[0].get_indices()
                self.add_measurement_at_node(meas,path)
            else: # if no selection append at end
                self.add_measurement_at_node(meas,None)
        self.update_store()

    def add_completion(self):
        toplevel = self.tree.get_toplevel() # get current top level window
        compl_name = misc.get_user_input_text(toplevel, "Please input Completion Date", "Add Completion")

        if compl_name != None:
            compl = Completion(compl_name)
            # get selection
            selection = self.tree.get_selection()
            if selection.count_selected_rows() != 0: # if selection exists
                [model, paths] = selection.get_selected_rows()
                path = paths[0].get_indices()
                self.add_measurement_at_node(compl,path)
            else: # if no selection append at end
                self.add_measurement_at_node(compl,None)
        self.update_store()

    @undoable
    def add_measurement_at_node(self,meas,path):
        selection = self.tree.get_selection()
        delete_path = None
        if path != None:
            if len(path) > 1: # if a measurement selected
                self.cmbs[path[0]].insert_item(path[1],meas)
                delete_path = [path[0],path[1]]
            else: # append to selected cmb
                self.cmbs[path[0]].append_item(meas)
                delete_path = [path[0],self.cmbs[path[0]].length()-1]
        else: # if no selection append at end
            if len(self.cmbs) != 0:
                self.cmbs[-1].append_item(meas)
                delete_path = [len(self.cmbs)-1,self.cmbs[-1].length()-1]
        self.update_store()

        yield "Add Measurement at '{}'".format(path)
        # Undo action
        self.delete_row(delete_path)
        
    def add_heading(self):
        toplevel = self.tree.get_toplevel() # get current top level window
        heading_name = misc.get_user_input_text(toplevel, "Please input Heading", "Add new Item: Heading")
        
        if heading_name != None:
            # get selection
            selection = self.tree.get_selection()
            heading = MeasurementItemHeading(heading_name)
            if selection.count_selected_rows() != 0: # if selection exists
                [model, paths] = selection.get_selected_rows()
                path = paths[0].get_indices()
                self.add_measurement_item_at_node(heading,path)
            else: # if no selection append at end
                self.add_measurement_item_at_node(heading,None)

    def add_nlbh(self,oldval=None):
        toplevel = self.tree.get_toplevel() # get current top level window
        itemnos = [None]
        captions = ['Description','Breakup','Nos','L','B','H','Total']
        columntypes = [MEAS_DESC,MEAS_CUST,MEAS_L,MEAS_L,MEAS_L,MEAS_L,MEAS_CUST]
        callback_total = lambda values,row:str(RecordNLBH(*values[1:6]).find_total())
        callback_breakup = lambda values,row:str(RecordNLBH(*values[1:6]).find_breakup())
        cellrenderers = [None] + [callback_breakup] + [None]*4 + [callback_total]
        item_schedule = self.schedule
        
        dialog = ScheduleDialog(toplevel,itemnos,captions,columntypes,cellrenderers,item_schedule)

        if oldval != None: # if edit mode add data
            dialog.set_model(oldval)
            data = dialog.run()
            if data != None: # if edited
                return data
            else: # if cancel pressed
                return None
        else: # if normal mode
            data = dialog.run()
            if data != None:
                item = MeasurementItemNLBH()
                item.set_model(data)
                # get selection
                selection = self.tree.get_selection()
                if selection.count_selected_rows() != 0: # if selection exists
                    [model, paths] = selection.get_selected_rows()
                    path = paths[0].get_indices()
                    self.add_measurement_item_at_node(item,path)
                else: # if no selection append at end
                    self.add_measurement_item_at_node(item,None)
                    
    def add_lllll(self,oldval=None):
        toplevel = self.tree.get_toplevel() # get current top level window
        itemnos = [None]*5
        captions = ['Description','Breakup','L1','L2','L3','L4','L5']
        columntypes = [MEAS_DESC,MEAS_CUST,MEAS_L,MEAS_L,MEAS_L,MEAS_L,MEAS_L]
        callback_breakup = lambda values,row:str(RecordLLLLL(*values[1:7]).find_breakup()) # check later
        cellrenderers = [None] + [callback_breakup] + [None]*5
        item_schedule = self.schedule
        
        dialog = ScheduleDialog(toplevel,itemnos,captions,columntypes,cellrenderers,item_schedule)

        if oldval != None: # if edit mode add data
            dialog.set_model(oldval)
            data = dialog.run()
            if data != None: # if edited
                return data
            else: # if cancel pressed
                return None
        else: # if normal mode
            data = dialog.run()
            if data != None:
                item = MeasurementItemLLLLL()
                item.set_model(data)
                # get selection
                selection = self.tree.get_selection()
                if selection.count_selected_rows() != 0: # if selection exists
                    [model, paths] = selection.get_selected_rows()
                    path = paths[0].get_indices()
                    self.add_measurement_item_at_node(item,path)
                else: # if no selection append at end
                    self.add_measurement_item_at_node(item,None)
                    
    def add_nnnnnnnn(self,oldval=None):
        toplevel = self.tree.get_toplevel() # get current top level window
        itemnos = [None]*8
        captions = ['Description','N1','N2','N3','N4','N5','N6','N7','N8']
        columntypes = [MEAS_DESC,MEAS_NO,MEAS_NO,MEAS_NO,MEAS_NO,MEAS_NO,MEAS_NO,MEAS_NO,MEAS_NO]
        cellrenderers = [None] + [None]*8
        item_schedule = self.schedule
        
        dialog = ScheduleDialog(toplevel,itemnos,captions,columntypes,cellrenderers,item_schedule)

        if oldval != None: # if edit mode add data
            dialog.set_model(oldval)
            data = dialog.run()
            if data != None: # if edited
                return data
            else: # if cancel pressed
                return None
        else: # if normal mode
            data = dialog.run()
            if data != None:
                item = MeasurementItemNNNNNNNN()
                item.set_model(data)
                # get selection
                selection = self.tree.get_selection()
                if selection.count_selected_rows() != 0: # if selection exists
                    [model, paths] = selection.get_selected_rows()
                    path = paths[0].get_indices()
                    self.add_measurement_item_at_node(item,path)
                else: # if no selection append at end
                    self.add_measurement_item_at_node(item,None)
                    
    def add_nnnnnt(self,oldval=None):
        toplevel = self.tree.get_toplevel() # get current top level window
        itemnos = [None]
        captions = ['Description','N1','N2','N3','N4','N5','Total']
        columntypes = [MEAS_DESC,MEAS_L,MEAS_L,MEAS_L,MEAS_L,MEAS_L,MEAS_CUST]
        callback_total = lambda values,row:str(RecordnnnnnT(*values[0:6]).find_total())
        cellrenderers = [None] + [None]*5 + [callback_total]
        item_schedule = self.schedule
        
        dialog = ScheduleDialog(toplevel,itemnos,captions,columntypes,cellrenderers,item_schedule)

        if oldval != None: # if edit mode add data
            dialog.set_model(oldval)
            data = dialog.run()
            if data != None: # if edited
                return data
            else: # if cancel pressed
                return None
        else: # if normal mode
            data = dialog.run()
            if data != None:
                item = MeasurementItemnnnnnT()
                item.set_model(data)
                # get selection
                selection = self.tree.get_selection()
                if selection.count_selected_rows() != 0: # if selection exists
                    [model, paths] = selection.get_selected_rows()
                    path = paths[0].get_indices()
                    self.add_measurement_item_at_node(item,path)
                else: # if no selection append at end
                    self.add_measurement_item_at_node(item,None)
                    
    def add_custom(self,oldval=None,itemtype=None):
        item = MeasurementItemCustom(None,itemtype)

        toplevel = self.tree.get_toplevel() # get current top level window
        itemnos_mask = item.itemnos_mask
        captions = item.captions
        columntypes = item.columntypes
        cellrenderers = []
        cust_iter = 0
        for columntype in columntypes:
            if columntype == MEAS_CUST:
                cellrenderers.append(item.cust_funcs[cust_iter])
                cust_iter += 1
            else:
                cellrenderers.append(None)
        item_schedule = self.schedule
        
        dialog = ScheduleDialog(toplevel, itemnos_mask, captions, columntypes, cellrenderers, item_schedule)

        if oldval is not None: # if edit mode add data
            dialog.set_model(oldval[:-2])
            data = dialog.run()
            if data is not None: # if edited
                return data + [oldval[5]] + [itemtype]
            else: # if cancel pressed
                return None
        else: # if normal mode
            data = dialog.run()
            if data is not None:
                item.set_model(data + [item.user_data] + [itemtype])
                # get selection
                selection = self.tree.get_selection()
                if selection.count_selected_rows() != 0: # if selection exists
                    [model, paths] = selection.get_selected_rows()
                    path = paths[0].get_indices()
                    self.add_measurement_item_at_node(item,path)
                else: # if no selection append at end
                    self.add_measurement_item_at_node(item,None)

    def add_abstract(self,oldval=None):
        from abstract_dialog import AbstractDialog

        item = MeasurementItemAbstract(None)

        toplevel = self.tree.get_toplevel() # get current top level window

        if oldval is not None: # if edit mode add data
            dialog = AbstractDialog(toplevel,oldval, self, self.schedule)
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
        
    @undoable
    def add_measurement_item_at_node(self,item,path):
        delete_path = None
        if path != None:
            if len(path) > 2: # if a measurement item selected
                self.cmbs[path[0]][path[1]].insert_item(path[2],item)
                delete_path = [path[0],path[1],path[2]]
            elif len(path) > 1: # if measurement group selected
                if isinstance(self.cmbs[path[0]][path[1]],Measurement): # check if meas item
                    self.cmbs[path[0]][path[1]].append_item(item)
                    delete_path = [path[0],path[1],self.cmbs[path[0]][path[1]].length()-1]
            elif len(path) == 1: # if cmb selected
                index_meas = self.cmbs[path[0]].length()-1
                if isinstance(self.cmbs[path[0]][index_meas],Measurement): # check if meas item
                    self.cmbs[path[0]][index_meas].append_item(item)
                    index_item = self.cmbs[path[0]][index_meas].length()-1
                    delete_path = [path[0],index_meas,index_item]
        else: # if path is None append at end
            if len(self.cmbs) != 0:
                if self.cmbs[-1].length() > 0:
                    if isinstance(self.cmbs[-1][-1],Measurement):
                        self.cmbs[-1][-1].append_item(item)
                        delete_path = [len(self.cmbs)-1,self.cmbs[-1].length()-1,self.cmbs[-1][-1].length()-1]
        self.update_store()
        
        yield "Add Measurement item at '{}'".format(path)
        # Undo action
        if delete_path != None:
            self.delete_row(delete_path)
        
    def delete_selected_row(self):
        # get selection
        selection = self.tree.get_selection()
        if selection.count_selected_rows() != 0: # if selection exists
            [model, paths] = selection.get_selected_rows()
            self.delete_row(paths[0].get_indices())
    
    @undoable
    def delete_row(self,path):
        item = None
        # get selection
        if len(path) == 1:
            item = self.cmbs[path[0]]
            del self.cmbs[path[0]]
        elif len(path) == 2:
            item = self.cmbs[path[0]][path[1]]
            self.cmbs[path[0]].remove_item(path[1])
        elif len(path) == 3:
            item = self.cmbs[path[0]][path[1]][path[2]]
            self.cmbs[path[0]][path[1]].remove_item(path[2])
        self.update_store()
        
        yield "Delete measurement items at '{}'".format(path)
        # Undo action
        if len(path) == 1:
            self.add_cmb_at_node(item,path[0])
        elif len(path) == 2:
            self.add_measurement_at_node(item,path)
        elif len(path) == 3:
            self.add_measurement_item_at_node(item,path)
        self.update_store()

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
                    
    def __init__(self,schedule,TreeStoreObject,TreeViewObject):
        self.schedule = schedule
        self.store = TreeStoreObject
        self.tree = TreeViewObject
            
        self.cmbs = [] # initialise item list
        
        self.clipboard = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD) # initialise clipboard
        
        # connect callbacks
        self.tree.connect("key-press-event", self.onKeyPressTreeview)
        
        self.update_store()
