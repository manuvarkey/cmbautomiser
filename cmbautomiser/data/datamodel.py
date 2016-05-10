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
from undo import undoable

# local files import
from __main__ import misc
from . import schedule, measurement, bill

# Setup logger object
log = logging.getLogger(__name__)


class DataModel:
    """Undoable class for storing agregated data model for CMBAutomiser App"""
    def __init__(self, data=None):
        # Base data
        self.schedule = schedule.Schedule()
        self.cmbs = []
        self.bills = []
        # Derived data
        self.lock_state = LockState()
        
        if data is not None:
            self.schedule.set_model(data[0])
            for cmb_model in data[1]:
                cmb = measurement.Cmb()
                cmb.set_model(cmb_model)
                self.cmbs.append(cmb)
            for bill_model in data[2]:
                bill = bill.Bill()
                bill.set_model(bill_model)
                self.bills.append(bill)
        # Update values
        self.update()
    
    def update(self):
        """Update derived data values"""
        # Update locks
        self.lock_state = LockState()
        for bill in self.bills:
            self.lock_state += LockState(bill.get_billed_items())
        for cmb in self.cmbs:
            for meas in cmb:
                if isinstance(meas, measurement.Measurement):
                    for measitem in meas:
                        if isinstance(meas, measurement.MeasurementItemAbstract):
                            self.lock_state += LockState(measitem.get_abstracted_items())
    
    def get_model(self):
        schedule_model = self.schedule.get_model()
        cmb_models = []
        bill_models = []
        for cmb in self.cmbs:
            cmb_models.append(cmb.get_model())
        for bill in self.bills:
            bill_models.append(bill.get_model())
        return ['DataModel', [schedule_model, cmb_models, bill_models]]
    
    def set_model(self, model):
        if model[0] == 'DataModel':
            self.__init__(model[1])
            
    def get_lock_states(self):
        self.update()
        return self.lock_state
        
    # Measurement Methods
    
    def get_custom_item_template(self, module):
        '''Get custom item template for ScheduleDialog and others'''
        item = data.measurement.MeasurementItemCustom(None,module)
        return [item.itemnos_mask, item.captions, item.columntypes, item.cellrenderers]
        
    def get_schmod_from_custmod(self, model):
        '''Get custom item template for ScheduleDialog and others'''
        return model[1][0:4]
        
    def get_custmod_from_schmod(self, schmod, custmodel, itemtype):
        '''Get custom item template for ScheduleDialog and others'''
        if custmodel is not None:
            return ['MeasurementItemCustom', schmod + [custmodel[1][4:6]]]
        else:
            item = data.measurement.MeasurementItemCustom(None, itemtype)
            nullmodel = item.get_model()
            return ['MeasurementItemCustom', schmod + [nullmodel[1][4:6]]]
    
    @undoable
    def add_cmb_at_node(self, cmb_model, row):
        # Obtain CMB item
        cmb = measurement.Cmb(cmb_model)
        if row != None:
            self.cmbs.insert(row,cmb)
            row_delete = row
        else:
            self.cmbs.append(cmb)
            row_delete = len(self.cmbs) - 1
        self.update()

        yield "Add CMB at '{}'".format([row])
        # Undo action
        self.delete_row_meas([row_delete])
        
    @undoable
    def add_measurement_at_node(self, meas_model, path):
        # Obtain Measurement item
        if meas_model[0] == 'Measurement':
            meas = measurement.Measurement(meas_model)
        elif meas_model[0] == 'Completion':
            meas = measurement.Completion(meas_model)
        
        delete_path = None
        if path != None:
            if len(path) > 1: # If a measurement selected
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
        self.update()
        
        yield "Add Measurement item at '{}'".format(path)
        # Undo action
        if delete_path != None:
            self.delete_row_meas(delete_path)
        
    @undoable
    def delete_row_meas(self,path):
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
        self.update()
        
        yield "Delete measurement items at '{}'".format(path)
        # Undo action
        if len(path) == 1:
            self.add_cmb_at_node(item,path[0])
        elif len(path) == 2:
            self.add_measurement_at_node(item,path)
        elif len(path) == 3:
            self.add_measurement_item_at_node(item,path)
        self.update()
        
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
            
    def get_paths(self, paths=None, flags = None, level=[]):
        """Returns a list of paths from model"""
        if paths == None:
            paths = []
            flags = self.flags
        for index, flag_part in enumerate(flags):
            if isinstance(flag_part, list):
                self.get_paths(paths, flag_part, level + [index])
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
