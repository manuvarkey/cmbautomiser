#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# data.py
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

import logging

# Local module import
from .. import misc

# Setup logger object
log = logging.getLogger(__name__)


# class storing individual items in schedule of work
class ScheduleItemGeneric:
    """Class stores a row in the generic schedule"""
    
    def __init__(self, item=None):
        if item is not None:
            self.item = item
        else:
            self.item = []

    def set_model(self, item):
        """Set data model"""
        self.item = item

    def get_model(self):
        """Get data model"""
        return self.item

    def __setitem__(self, index, value):
        self.item[index] = value

    def __getitem__(self, index):
        return self.item[index]

    def print_item(self):
        print(self.item)
        

class ScheduleItem(ScheduleItemGeneric):
    """Class stores a row in the schedule of rates"""
    
    def __init__(self, itemno='', description='', unit='', rate='0', qty='0', reference='', excess_rate_percent='100', percentage='True'):
        self.itemno = itemno
        self.description = description
        self.unit = unit
        try:
            self.rate = round(eval(rate), 2)
        except:
            log.warning('ScheduleItem - Wrong value loaded in model - rate - ' + rate)
            rate = '0'
            self.rate = 0
        try:
            self.qty = eval(qty)
        except:
            log.warning('ScheduleItem - Wrong value loaded in model - qty - ' + qty)
            qty = '0'
            self.qty = 0
        self.reference = reference
        try:
            self.excess_rate_percent = eval(excess_rate_percent)
        except:
            log.warning('ScheduleItem - Wrong value loaded in model - excess_rate_percent - ' + excess_rate_percent)
            excess_rate_percent = '100'
            self.excess_rate_percent = 100
        try:
            self.percentage = eval(percentage)
        except:
            log.warning('ScheduleItem - Wrong value loaded in model - percentage - ' + percentage)
            percentage = 'True'
            self.percentage = True
        # extended descritption of item
        self.extended_description = ''
        self.extended_description_limited = ''
        
        # Initialise base class
        super(ScheduleItem, self).__init__([itemno, description, unit, rate, qty, reference, excess_rate_percent, percentage])

    def __setitem__(self, index, value):
        self.item[index] = value
        
        if index == 0:
            self.itemno = value
        elif index == 1:
            self.description = value
        elif index == 2:
            self.unit = value
        elif index == 3:
            try:
                self.rate = round(float(eval(value)), 2)
            except:
                log.warning('ScheduleItem - Wrong value loaded in model - rate - ' + value)
                self.rate = 0
                value = '0'
        elif index == 4:
            try:
                self.qty = float(eval(value))
            except:
                log.warning('ScheduleItem - Wrong value loaded in model - qty - ' + value)
                self.qty = 0
                value = '0'
        elif index == 5:
            self.reference = value
        elif index == 6:
            try:
                self.excess_rate_percent = float(eval(value))
            except:
                log.warning('ScheduleItem - Wrong value loaded in model - excess_rate_percent - ' + value)
                self.excess_rate_percent = 100
                value = '100'
        elif index == 7:
            try:
                self.percentage = bool(eval(value))
            except:
                log.warning('ScheduleItem - Wrong value loaded in model - percentage - ' + value)
                self.percentage = True
                value = 'True'
        self.item[index] = value

class ScheduleGeneric:
    """Class stores a generic schedule"""
    
    def __init__(self, items=None):
        if items is not None:
            self.items = items  # main data store of rows
        else:
            self.items = []

    def append_item(self, item):
        """Append item at end of schedule"""
        self.items.append(item)

    def get_item_by_index(self, index):
        return self.items[index]

    def set_item_at_index(self, index, item):
        self.items[index] = item

    def insert_item_at_index(self, index, item):
        self.items.insert(index, item)

    def remove_item_at_index(self, index):
        del (self.items[index])

    def __setitem__(self, index, value):
        self.items[index] = value

    def __getitem__(self, index):
        return self.items[index]
    
    def get_model(self):
        """Get data model"""
        items = []
        for item in self.items:
            items.append(item.get_model())
        return items
        
    def set_model(self,items):
        """Set data model"""
        for item in items:
            self.items.append(ScheduleItemGeneric(item))
            
    def length(self):
        """Return number of rows"""
        return len(self.items)

    def clear(self):
        del self.items[:]

    def print_item(self):
        print("schedule start")
        for item in self.items:
            item.print_item()
        print("schedule end")
        
        
class Schedule(ScheduleGeneric):
    """Class stores the schedule of rates for work"""
    
    def __init__(self, items=None):
        #Initialise base class
        super(Schedule,self).__init__(items)
        self.update_values()
        
    def update_values(self):
        """Populate ScheduleItem.extended_description (Used for final billing)"""
        iter = 0
        extended_description = ''
        itemno = ''
        while iter < len(self.items):
            item = self.items[iter]
            # Item without quantity
            if item.qty == 0 and item.unit == '' and item.rate == 0:
                # If main item, reset values
                if item.itemno != '':
                    itemno = item.itemno
                    extended_description = item.description
                # If continuing item append description
                else:
                    extended_description = extended_description + '\n' + item.description
                # Keep description of non quantity item
                item.extended_description = item.description
            # If sub item of main item i.e. with quantity and starts with itemno
            elif itemno !='' and item.itemno.startswith(itemno):
                item.extended_description = extended_description + '\n' + item.description
            # If one line item
            else:
                item.extended_description = item.description
                # Reset values to nil
                itemno = item.itemno
                extended_description = item.description
                
            if len(item.extended_description) > misc.CMB_DESCRIPTION_MAX_LENGTH:
                item.extended_description_limited = item.extended_description[0:int(misc.CMB_DESCRIPTION_MAX_LENGTH/2)] + \
                    ' ... ' + item.extended_description[-int(misc.CMB_DESCRIPTION_MAX_LENGTH/2):]
            else:
                item.extended_description_limited = item.extended_description
            iter += 1
            
    def __setitem__(self, index, value):
        if isinstance(index, int):
            self.items[index] = value
        elif isinstance(index, str):
            for item in self.items:
                if item.itemno == index:
                    self.items[index] = value
                    break
            else:
                log.warning("Schedule - Itemno not found while assigning value")

    def __getitem__(self, index):
        if isinstance(index, int):
            return self.items[index]
        elif isinstance(index, str):
            for item in self.items:
                if item.itemno == index:
                    return item
            return None
    
    def set_model(self,items):
        """Set data model"""
        self.items.clear()
        for item in items:
            self.items.append(ScheduleItem(*item))
        self.update_values()
            
    def get_itemnos(self):
        """Returns a list of itemnos with order as in schedule"""
        itemnos = []
        for item in self.items:
            if item.itemno != '' and item.qty != 0:
                itemnos.append(item.itemno)
        return itemnos

