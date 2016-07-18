#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#  datamodel.py
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
import os.path, copy, logging

# local files import
from __main__ import misc, undo
from undo import undoable
from . import schedule, measurement, bill

# Setup logger object
log = logging.getLogger(__name__)


class DataModel:
    """Undoable class for storing agregated data model for CMBAutomiser App"""
    def __init__(self, data=None):
        log.info('DataModel - Initialise')
        # Base data
        self.schedule = schedule.Schedule()
        self.cmbs = []
        self.bills = []
        # Derived data
        self.lock_state = LockState()  # Billed/Abstracted states of measurement items
        self.cmb_ref = []  # Array of sets corresponding to cmbs refered to by particular cmb
        
        if data is not None:
            self.schedule.set_model(data[0])
            for cmb_model in data[1]:
                cmb = measurement.Cmb()
                cmb.set_model(cmb_model)
                self.cmbs.append(cmb)
            for bill_model in data[2]:
                bill_item = bill.Bill()
                bill_item.set_model(bill_model)
                self.bills.append(bill_item)
        # Update values
        self.update()
    
    def get_model(self):
        """Return data model"""
        schedule_model = self.schedule.get_model()
        cmb_models = []
        bill_models = []
        for cmb in self.cmbs:
            cmb_models.append(cmb.get_model())
        for bill in self.bills:
            bill_models.append(bill.get_model())
        return ['DataModel', [schedule_model, cmb_models, bill_models]]
    
    def set_model(self, model):
        """Set data model"""
        if model[0] == 'DataModel':
            self.schedule.set_model(model[1][0])
            self.cmbs.clear()
            self.bills.clear()
            for cmb_model in model[1][1]:
                cmb = measurement.Cmb()
                cmb.set_model(cmb_model)
                self.cmbs.append(cmb)
            for bill_model in model[1][2]:
                bill_item = bill.Bill()
                bill_item.set_model(bill_model)
                self.bills.append(bill_item)
            # Update values
            self.update()
    
    def update(self):
        """Update derived data values"""
        
        log.info('DataModel - update called')

        # Calculate extended descriptions
        self.schedule.update_values()
        
        # Update all bills
        for bill in self.bills:
            bill.update(self.schedule, self.cmbs, self.bills)
        
        # Update locks
        self.lock_state = LockState()
        for bill in self.bills:
            self.lock_state += LockState(bill.get_billed_items())
        for cmb in self.cmbs:
            for meas in cmb:
                if isinstance(meas, measurement.Measurement):
                    for measitem in meas:
                        if isinstance(measitem, measurement.MeasurementItemAbstract):
                            self.lock_state += LockState(measitem.get_abstracted_items())
        
        # Update 1) dependency tree of cmbs. 2) measurement abstracts
        self.cmb_ref = [set()]*len(self.cmbs)
        for cmb_no, cmb in enumerate(self.cmbs):
            for meas_no, meas in enumerate(cmb.items):
                if isinstance(meas, measurement.Measurement):
                    for meas_item_nos, meas_item in enumerate(meas.items):
                        if isinstance(meas_item, measurement.MeasurementItemAbstract):
                            # Update MeasurementItemAbstract
                            meas_item.update(self.cmbs)
                            # Update Dependency
                            for mitem in meas_item.mitems:
                                if mitem[0] != cmb_no:
                                    # Update cmb with abstract
                                    self.cmb_ref[cmb_no] |= set([mitem[0]])
                                    # Update cmb with item abstracted
                                    self.cmb_ref[mitem[0]] |= set([cmb_no])
            
    def get_lock_states(self):
        """Return underlying LockState object for App"""
        self.update()
        return self.lock_state
        
    # Measurement Methods
    
    def get_custom_item_template(self, module):
        """Get custom item template for ScheduleDialog and others"""
        item = measurement.MeasurementItemCustom(None,module)
        return [item.itemnos_mask, item.captions, item.columntypes, item.cust_funcs, item.dimensions]
        
    def get_schmod_from_custmod(self, model):
        """Get custom item template for ScheduleDialog and others"""
        return model[1][0:4]
        
    def get_custmod_from_schmod(self, schmod, custmodel, itemtype):
        """Get custom item template for ScheduleDialog and others"""
        if custmodel is not None:
            return ['MeasurementItemCustom', schmod + custmodel[1][4:6]]
        else:
            item = measurement.MeasurementItemCustom(None, itemtype)
            nullmodel = item.get_model()
            return ['MeasurementItemCustom', schmod + nullmodel[1][4:6]]
    
    def update_static_paths(self, path, add_flag = True):
        """Updates static paths in bills and AbstractMeasurement on change in measurements model
        
            Arguments:
                path: Path to item being added/deleted
                add_flag: True for additions, False for deletions
        """
        log.info('DataModel - update_static_paths - ' + str([path, add_flag]))
        # Initialise replacement dictionary
        mod_dict = dict()
        # Paths to be deleted
        del_path = []
        # Saved values for reversal
        bill_mitems_old = []
        bill_paths_old = []
        abs_mitems_old = []
        abs_paths_old = []
        # Increment path by 1 for add and decrement by 1 for remove
        increment = 1 if add_flag else -1
        
        # CMB added/removed
        if len(path) == 1:
            for p1, cmb in enumerate(self.cmbs[path[0]:], path[0]):
                for p2, meas in enumerate(self.cmbs[p1].items):
                    if isinstance(self.cmbs[p1][p2], measurement.Measurement):
                        for p3, meas_item in enumerate(self.cmbs[p1][p2].items):
                            # Update items to be deleted
                            if [p1] == path and not add_flag:
                                del_path.append([p1,p2,p3])
                            # Update items to be changed
                            else:
                                mod_dict[(p1,p2,p3)] = (p1+increment,p2,p3)
                    mod_dict[(p1,p2)] = (p1+increment,p2)
                mod_dict[tuple([p1])] = tuple([p1+increment])
        # Measurement added/removed
        elif len(path) == 2:
            p1 = path[0]
            for p2, meas in enumerate(self.cmbs[p1].items[path[1]:], path[1]):
                if isinstance(self.cmbs[p1][p2], measurement.Measurement):
                    for p3, meas_item in enumerate(self.cmbs[p1][p2].items):
                        # Update items to be deleted
                        if [p1,p2] == path and not add_flag:
                            del_path.append([p1,p2,p3])
                        # Update items to be changed
                        else:
                            mod_dict[(p1,p2,p3)] = (p1,p2+increment,p3)
                mod_dict[(p1,p2)] = (p1,p2+increment)
        # Measurement item added/removed
        elif len(path) == 3:
            p1 = path[0]
            p2 = path[1]
            for p3, meas_item in enumerate(self.cmbs[p1][p2].items[path[2]:], path[2]):
                # Update items to be deleted
                if [p1,p2,p3] == path  and not add_flag:
                    del_path.append([p1,p2,p3])
                # Update items to be changed
                else:
                    mod_dict[(p1,p2,p3)] = (p1,p2,p3+increment)
        
        # Make replacements in bill
        for row, bill in enumerate(self.bills):
            changed = False
            mitem_copy = copy.deepcopy(bill.data.mitems)
            # If remove flag, remove path from bill 
            if not add_flag:
                for mitem in mitem_copy:
                    if mitem in del_path:
                        bill.data.mitems.remove(mitem)
                        changed = True
            for index, item in enumerate(bill.data.mitems):
                if tuple(item) in mod_dict:
                    bill.data.mitems[index] = list(mod_dict[tuple(item)])
                    changed = True

            # If bill changed
            if changed:
                bill_mitems_old.append(mitem_copy)
                bill_paths_old.append(row)
                
        # Make replacements in abstract measurements
        for p1, cmb in enumerate(self.cmbs):
            for p2, meas in enumerate(self.cmbs[p1].items):
                if isinstance(self.cmbs[p1][p2], measurement.Measurement):
                    for p3, meas_item in enumerate(self.cmbs[p1][p2].items):
                        if isinstance(meas_item, measurement.MeasurementItemAbstract):
                            changed = False
                            mitem_copy = copy.deepcopy(meas_item.mitems)
                            # If remove flag, remove path from bill 
                            if not add_flag: 
                                for mitem in mitem_copy:
                                    if mitem in del_path:
                                        meas_item.mitems.remove(mitem)
                                        changed = True
                            for index, item in enumerate(meas_item.mitems):
                                if tuple(item) in mod_dict:
                                    meas_item.mitems[index] = list(mod_dict[tuple(item)])
                                    changed = True
                            # If abstract changed
                            if changed:
                                abs_mitems_old.append(mitem_copy)
                                abs_paths_old.append([p1,p2,p3])
        return [bill_paths_old, bill_mitems_old, abs_paths_old, abs_mitems_old]
        
    def replace_static_paths(self, data=None):
        """Function reverses effect of update_static_paths"""
        log.info('DataModel - replace_static_paths')
        log.debug('DataModel - replace_static_paths - \n' + str(data))
        if data is not None:
            bill_paths_old = data[0]
            bill_mitems_old = data[1]
            abs_paths_old = data[2]
            abs_mitems_old = data[3]
            
            # Make replacements in bill
            for billno, mitems in zip(bill_paths_old, bill_mitems_old):
                self.bills[billno].data.mitems = mitems
            # Make replacements in abstract
            for abspath, mitems in zip(abs_paths_old, abs_mitems_old):
                self.cmbs[abspath[0]][abspath[1]][abspath[2]].mitems = mitems
    
    @undoable
    def add_cmb_at_node(self, cmb_model, row):
        """Undoable function for adding a Cmb to model"""
        log.info('DataModel - add_cmb_at_node - ' + str(row))
        # Obtain CMB item
        if cmb_model[0] == 'CMB':
            cmb = measurement.Cmb(cmb_model[1])
            if row != None:
                # Update static paths
                self.update_static_paths([row])
                
                self.cmbs.insert(row,cmb)
                row_delete = row
            else:
                self.cmbs.append(cmb)
                row_delete = len(self.cmbs) - 1
            self.update()
        else:
            log.warning('add_cmb_at_node - Wrong model loaded')

        yield "Add CMB at '{}'".format([row])
        # Undo action
        self.delete_row_meas([row_delete])
        
    @undoable
    def add_measurement_at_node(self, meas_model, path):
        """Undoable function for adding a Measurement/Completion to model"""
        log.info('DataModel - add_measurement_at_node - ' + str(path))
        # Obtain Measurement item
        if meas_model[0] == 'Measurement':
            meas = measurement.Measurement(meas_model[1])
        elif meas_model[0] == 'Completion':
            meas = measurement.Completion(meas_model[1])
        else:
            log.warning('add_measurement_at_node - Wrong model loaded')
            return
        
        delete_path = None
        if path != None:
            if len(path) > 1: # If a measurement selected
                # Update static paths
                self.update_static_paths(path)
                
                self.cmbs[path[0]].insert_item(path[1],meas)
                delete_path = [path[0],path[1]]
            else: # Append to selected Cmb
                self.cmbs[path[0]].append_item(meas)
                delete_path = [path[0],self.cmbs[path[0]].length()-1]
        else: # If no selection append at end
            if len(self.cmbs) != 0:
                self.cmbs[-1].append_item(meas)
                delete_path = [len(self.cmbs)-1,self.cmbs[-1].length()-1]
        self.update()

        yield "Add Measurement at '{}'".format(path)
        # Undo action
        self.delete_row_meas(delete_path)
        
    @undoable
    def add_measurement_item_at_node(self,item_model,path):        
        """Undoable function for adding a MeasurementItem to model"""
        log.info('DataModel - add_measurement_item_at_node - ' + str(path))
        delete_path = None
        item = None
        if item_model[0] == 'MeasurementItemHeading':
            item = measurement.MeasurementItemHeading(item_model[1])
        elif item_model[0] == 'MeasurementItemCustom':
            item = measurement.MeasurementItemCustom(item_model[1], item_model[1][5])
        elif item_model[0] == 'MeasurementItemAbstract':
            item = measurement.MeasurementItemAbstract(item_model[1])
        else:
            log.warning('add_measurement_item_at_node - Wrong model loaded')
            return
            
        if path != None:
            if len(path) > 2: # if a measurement item selected
                # Update static paths
                self.update_static_paths(path)
                
                self.cmbs[path[0]][path[1]].insert_item(path[2],item)
                delete_path = [path[0],path[1],path[2]]
            elif len(path) > 1: # if measurement group selected
                if isinstance(self.cmbs[path[0]][path[1]], measurement.Measurement): # check if meas item
                    self.cmbs[path[0]][path[1]].append_item(item)
                    delete_path = [path[0],path[1],self.cmbs[path[0]][path[1]].length()-1]
            elif len(path) == 1: # if cmb selected
                index_meas = self.cmbs[path[0]].length()-1
                if isinstance(self.cmbs[path[0]][index_meas], measurement.Measurement): # check if meas item
                    self.cmbs[path[0]][index_meas].append_item(item)
                    index_item = self.cmbs[path[0]][index_meas].length()-1
                    delete_path = [path[0],index_meas,index_item]
        else: # if path is None append at end
            if len(self.cmbs) != 0:
                if self.cmbs[-1].length() > 0:
                    if isinstance(self.cmbs[-1][-1], measurement.Measurement):
                        self.cmbs[-1][-1].append_item(item)
                        delete_path = [len(self.cmbs)-1,self.cmbs[-1].length()-1,self.cmbs[-1][-1].length()-1]
        self.update()
        
        yield "Add Measurement item at '{}'".format(path)
        # Undo action
        if delete_path != None:
            self.delete_row_meas(delete_path)
            
    @undoable
    def edit_measurement_item(self, path, item, newval, oldval):
        """Undoable function for editing a MeasurementItem in model"""
        log.info('DataModel - edit_measurement_item - ' + str(path))
        if newval != None:
            if len(path) == 1:
                item.set_name(newval)
            elif len(path) == 2:
                if isinstance(item, measurement.Measurement):
                    item.set_date(newval)
                elif isinstance(item, measurement.Completion):
                    item.set_date(newval)
            elif len(path) == 3:
                if isinstance(item, measurement.MeasurementItemHeading):
                    item.set_remark(newval)
                else:
                    item.set_model(newval)
            self.update()
        
        yield "Edit measurement items at '{}'".format(path)
        # Undo action
        if oldval != None and newval != None:
            if len(path) == 1:
                item.set_name(oldval)
            elif len(path) == 2:
                if isinstance(item, measurement.Measurement):
                    item.set_date(oldval)
                elif isinstance(item, measurement.Completion):
                    item.set_date(oldval)
            elif len(path) == 3:
                if isinstance(item, measurement.MeasurementItemHeading):
                    item.set_remark(oldval)
                else:
                    item.set_model(oldval)
            self.update()
        
    @undoable
    def delete_row_meas(self, path):
        """Undoable function for deleting a measurement item from model"""
        log.info('DataModel - delete_row_meas - ' + str(path))
        item = None
        # Update static paths
        static_paths_old = self.update_static_paths(path, False)

        if len(path) == 1:
            item = self.cmbs[path[0]].get_model()
            del self.cmbs[path[0]]
        elif len(path) == 2:
            item = self.cmbs[path[0]][path[1]].get_model()
            self.cmbs[path[0]].remove_item(path[1])
        elif len(path) == 3:
            item = self.cmbs[path[0]][path[1]][path[2]].get_model()
            self.cmbs[path[0]][path[1]].remove_item(path[2])
        self.update()
        
        yield "Delete measurement items at '{}'".format(path)
        # Undo action
        if len(path) == 1:
            self.add_cmb_at_node(item,path[0])
        elif len(path) == 2:
            self.add_measurement_at_node(item,path)
        elif len(path) == 3:
            self.add_measurement_item_at_node(item,path)
        # Replace static paths lost on delete
        self.replace_static_paths(static_paths_old)
        
        self.update()
        
    def render_cmb(self, folder, replacement_dict, path, recursive = True, progress=None):
        """Render CMB
            
            Arguments:
                folder: Output folder
                replacement_dict: Replacement dictionary for global values
                path: Path of CMB to be rendered
                recursive: Flag to select rendering of dependent items
        """
        log.info('DataModel - render_cmb - ' + str([path, recursive]))
        # Build all data structures
        self.update()
        
        # Fill in latex buffer
        latex_buffer = self.cmbs[path[0]].get_latex_buffer([path[0]], self.schedule)

        # Make global variables replacements
        latex_buffer.replace_and_clean(replacement_dict)
        
        # Include linked cmbs and bills
        replacement_dict_external_docs = {}
        external_docs = ''
        # Include cmbs
        for count,cmb in enumerate(self.cmbs):
            if path[0] in self.cmb_ref[count]:
                external_docs += '\externaldocument{cmb_' + str(count+1) + '}\n'
        # Include bills
        for count,bill in enumerate(self.bills):
            if path[0] in bill.cmb_ref:
                external_docs += '\externaldocument{abs_' + str(count+1) + '}\n'
        replacement_dict_external_docs['$cmbexternaldocs$'] = external_docs
        latex_buffer.replace(replacement_dict_external_docs)

        # Write output
        filename = misc.posix_path(folder,'cmb_' + str(path[0]+1) + '.tex')
        latex_buffer.write(filename)

        # Run latex on files and dependencies
        
        # Setup progress window
        if progress is not None:
            billno = 0
            cmbno = 0
            for bill in self.bills:
                if path[0] in bill.cmb_ref:
                    billno += 1
            for cmb_count, cmb in enumerate(self.cmbs):
                if path[0] in self.cmb_ref[cmb_count]:
                    cmbno += 1
            progress.set_pulse_step(1/((cmbno+billno)*2 + 2))
        
        # Run on all cmbs refered by cmb
        if recursive: # if recursive call
            for cmb_count, cmb in enumerate(self.cmbs):
                if path[0] in self.cmb_ref[cmb_count]:
                    if progress is not None:
                        progress.add_message('Rendering CMB No.' + cmb.name)
                    code = self.render_cmb(folder, replacement_dict, [cmb_count],False)
                    if progress is not None:
                        progress.pulse()
                    if code[0] == misc.CMB_ERROR:
                        return code
        # Run on all bills refering cmb
        if recursive: # if recursive call
            for bill_count,bill in enumerate(self.bills):
                if path[0] in bill.cmb_ref:
                    if progress is not None:
                        progress.add_message('Rendering Bill No.' + bill.data.cmb_name)
                    code = self.render_bill(folder, replacement_dict, [bill_count], False)
                    if progress is not None:
                        progress.pulse()
                    if code[0] == misc.CMB_ERROR:
                        return code
                        
        # Run latex on file
        if progress is not None:
            progress.add_message('Rendering CMB No.' + self.cmbs[path[0]].name)
        code = misc.run_latex(misc.posix_path(folder), filename)
        if progress is not None:
            progress.pulse()
        if code == misc.CMB_ERROR:
            return (misc.CMB_ERROR,'Rendering of CMB No.' + self.cmbs[path[0]].get_name() + ' failed')

        # Run on all cmbs and bills refering cmb again to rebuild indexes on recursive run
        # Run on all cmbs refered by cmb
        if recursive: # if recursive call
            for cmb_count, cmb in enumerate(self.cmbs):
                if path[0] in self.cmb_ref[cmb_count]:
                    if progress is not None:
                        progress.add_message('Rebuilding index for CMB No.' + cmb.name)
                    code = self.render_cmb(folder, replacement_dict, [cmb_count],False)
                    if progress is not None:
                        progress.pulse()
                    if code[0] == misc.CMB_ERROR:
                        return code
        # Run on all bills refering cmb
        if recursive: # if recursive call
            for bill_count,bill in enumerate(self.bills):
                if path[0] in bill.cmb_ref:
                    if progress is not None:
                        progress.add_message('Rebuilding index for Bill No.' + bill.data.cmb_name)
                    code = self.render_bill(folder, replacement_dict, [bill_count], False)
                    if progress is not None:
                        progress.pulse()
                    if code[0] == misc.CMB_ERROR:
                        return code
        
        # Get spreadsheet buffer
        if progress is not None:
            progress.add_message('Rendering CMB No.' + self.cmbs[path[0]].name + ' to .xlsx')
        spreadsheet = self.cmbs[path[0]].get_spreadsheet_buffer([path[0]], self.schedule)
        filename = misc.posix_path(folder,'cmb_' + str(path[0]+1) + '.xlsx')
        spreadsheet.save(filename)
        if progress is not None:
            progress.add_message('<b>Rendering Finished</b>')
            progress.pulse(end=True)
        # Return status code for main application interface
        return (misc.CMB_INFO,'CMB No.' + self.cmbs[path[0]].get_name() + ' rendered successfully')
    
    # Bill methods
    
    @undoable
    def insert_bill_at_row(self, data_model, row):  # note needs rows to be sorted
        """Undoable function for adding a bill to model"""
        log.info('DataModel - insert_bill_at_row - ' + str(row))
        item = bill.Bill(data_model)
        if row is not None:
            new_row = row
            self.bills.insert(row, item)
        else:
            self.bills.append(item)
            new_row = len(self.bills) - 1
        self.update()

        yield "Insert data items to bill at row '{}'".format(new_row)
        # Undo action
        self.delete_bill(new_row)
        self.update()
        
    @undoable
    def edit_bill_at_row(self, data_model, row):
        """Undoable function for editing a bill in model"""
        log.info('DataModel - edit_bill_at_row - ' + str(row))
        if row is not None:
            old_data = copy.deepcopy(self.bills[row].get_model())
            self.bills[row].set_model(data_model)
        self.update()

        yield "Edit bill item at row '{}'".format(row)
        # Undo action
        if row is not None:
            self.bills[row].set_model(old_data)
        self.update()
    
    @undoable
    def delete_bill(self, row):
        """Undoable function for deleting a bill from model"""
        log.info('DataModel - delete_bill - ' + str(row))
        data_model = self.bills[row].get_model()
        del self.bills[row]
        self.update()

        yield "Delete data items from bill at row '{}'".format(row)
        # Undo action
        self.insert_bill_at_row(data_model, row)
        self.update()
        
    def render_bill(self, folder, replacement_dict, path, recursive=True, progress=None):
        """Render bill to file
            
            Arguments:
                folder: Output folder
                replacement_dict: Replacement dictionary for global values
                path: Path of Bill to be rendered
                recursive: Flag to select rendering of dependent items
        """
        log.info('DataModel - render_bill - ' + str([path, recursive]))
                
        # Build all data structures
        self.update()
        
        # Render only if bill is a normal bill
        if self.bills[path[0]].data.bill_type == misc.BILL_NORMAL:
            bill = self.bills[path[0]]
            
            # Fill in latex buffer
            latex_buffer = bill.get_latex_buffer([path[0]], self.schedule)
            latex_buffer_bill = bill.get_latex_buffer_bill(self.schedule)
            # Make global variables replacements
            latex_buffer.replace_and_clean(replacement_dict)
            latex_buffer_bill.replace_and_clean(replacement_dict)
            
            # Include linked cmbs
            cmb_refs = bill.cmb_ref.copy()
            replacement_dict_cmbs = {}
            external_docs = ''
            # Add all cmbs depending on cmbs billed to include abstracted items
            for count in bill.cmb_ref:
                cmb_refs |= self.cmb_ref[count]
            for cmbpath in cmb_refs:
                if cmbpath != -1:
                    external_docs += '\externaldocument{cmb_' + str(cmbpath + 1) + '}\n'
                elif bill.data.prev_bill is not None: # prev abstract
                    external_docs += '\externaldocument{abs_' + str(bill.data.prev_bill + 1) + '}\n'
            replacement_dict_cmbs['$cmbexternaldocs$'] = external_docs
            latex_buffer.replace(replacement_dict_cmbs)
            
            # Setup progress window
            if progress is not None:
                billno = 0
                cmbno = len(cmb_refs)
                if -1 in cmb_refs:
                    cmbno -= 1
                if bill.prev_bill is not None and bill.prev_bill.data.bill_type == misc.BILL_NORMAL:
                    billno = 1
                progress.set_pulse_step(1/(cmbno*2+billno+3))

            # Write output
            filename = misc.posix_path(folder, 'abs_' + str(path[0] + 1) + '.tex')
            latex_buffer.write(filename)
            filename_bill = misc.posix_path(folder, 'bill_' + str(path[0] + 1) + '.tex')
            latex_buffer_bill.write(filename_bill)

            # run latex on file and dependencies
            if recursive:  # if recursive call
                # Render all cmbs depending on the bill
                for cmb_ref in cmb_refs:
                    if cmb_ref is not -1:  # if not prev bill
                        if progress is not None:
                            progress.add_message('Rendering CMB No.' + self.cmbs[cmb_ref].name)
                        code = self.render_cmb(folder, replacement_dict, [cmb_ref], False)
                        if progress is not None:
                            progress.pulse()
                        if code[0] == misc.CMB_ERROR:
                            return code
                # Render prev bill
                if bill.prev_bill is not None and bill.prev_bill.data.bill_type == misc.BILL_NORMAL:
                    if progress is not None:
                        progress.add_message('Rendering Bill No.' + bill.prev_bill.data.cmb_name)
                    code = self.render_bill(folder, replacement_dict, [bill.data.prev_bill], False)
                    if progress is not None:
                        progress.pulse()
                    if code[0] == misc.CMB_ERROR:
                        return code

            # Render this bill
            if progress is not None:
                progress.add_message('Rendering Bill No.' + bill.data.cmb_name + ' Abstract')
            code = misc.run_latex(misc.posix_path(folder), filename)
            if progress is not None:
                progress.pulse()
            if code == misc.CMB_ERROR:
                return (misc.CMB_ERROR, 'Rendering of Bill: ' + bill.data.title + ' failed')
            if progress is not None:
                progress.add_message('Rendering Bill No.' + bill.data.cmb_name + ' Schedule')
            code_bill = misc.run_latex(misc.posix_path(folder), filename_bill)
            if progress is not None:
                progress.pulse()
            if code_bill == misc.CMB_ERROR:
                return (misc.CMB_ERROR, 'Rendering of Bill Schedule: ' + bill.data.title + ' failed')

            # Render all cmbs again to rebuild indexes on recursive run
            if recursive:  # if recursive call
                for cmb_ref in cmb_refs:
                    if cmb_ref is not -1:  # if not prev bill
                        if progress is not None:
                            progress.add_message('Rebuilding index for CMB No.' + self.cmbs[cmb_ref].name)
                        code = self.render_cmb(folder, replacement_dict, [cmb_ref], False)
                        if progress is not None:
                            progress.pulse()
                        if code[0] == misc.CMB_ERROR:
                            return code
                # Write spreadsheet output
                filename_bill_spreadsheet = misc.posix_path(folder, 'bill_' + str(path[0] + 1) + '.xlsx')
                if progress is not None:
                    progress.add_message('Rendering Bill No.' + bill.data.cmb_name + ' to .xlsx')
                bill.export_spreadsheet_bill(filename_bill_spreadsheet, replacement_dict, self.schedule, self.cmbs)
                if progress is not None:
                    progress.pulse()
                if progress is not None:
                    progress.add_message('<b>Rendering Finished</b>')
                    progress.pulse(end=True)

            return (misc.CMB_INFO, 'Bill: ' + self.bills[path[0]].data.title + ' rendered successfully')
        else:
            return (misc.CMB_WARNING, 'Rendering of custom bill not supported')

        
class LockState:
    """Implements variable for storing N-D array of bools for tracking variable lock states"""
    
    def __init__(self, mitems = None):
        """Initialises class with list of path indices"""
        self.flags = []
        if mitems != None:
            for mitem in mitems:
                self.__setitem__(mitem, True)
                
    def resize(self, path):
        """Resize array to path"""
        flag_part = self.flags
        for index in path[:-1]:
            if len(flag_part) <= index:
                # Expand array
                for i in range(index - len(flag_part) + 1):
                    flag_part.append([])
            flag_part = flag_part[index]
        if len(flag_part) <= path[-1]:
            flag_part.extend([False]*(path[-1] - len(flag_part) + 1))
            
    def get_paths(self, paths=None, flags = None, level=None):
        """Returns a list of paths from model"""
        if level is None:
            level = []
        if paths == None:
            paths = []
            flags = self.flags
        for index, flag_part in enumerate(flags):
            if isinstance(flag_part, list):
                self.get_paths(paths, flag_part, level + [index])
            # If at terminal node, append path
            elif flag_part == True:
                paths.append(level + [index])
        return paths
                        
    def __setitem__(self, path, value):
        """Set path"""
        # Resize array to path
        self.resize(path)
        flag_part = self.flags
        if value in [True, False]:
            for index in path[:-1]:
                flag_part = flag_part[index]
            flag_part[path[-1]] = value
            
    def __getitem__(self, path):
        """Get path
            
            Returns:
                True/False: if path exist
                None: If path does not exist
        """
        # Resize array to path
        flag_part = self.flags
        for index in path[:-1]:
            if len(flag_part) > index:
                flag_part = flag_part[index]
            else:
                return None
        if len(flag_part) > path[-1]:
            return flag_part[path[-1]]
        else:
            return None
                
    def __add__(self, other):
        """Add items"""
        # Copy current array
        new_lock = copy.deepcopy(self)
        # Resize array to size of other and set item
        paths = other.get_paths()
        for path in paths:
            new_lock.resize(path)
            new_lock[path] = True
        return new_lock
                
    def __sub__(self, other):
        """Add items"""
        # Copy current array
        new_lock = copy.deepcopy(self)
        # Resize array to size of other and set item
        paths = other.get_paths()
        for path in paths:
            new_lock.resize(path)
            new_lock[path] = False
        return new_lock
        
