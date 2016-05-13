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

# Setup logger object
log = logging.getLogger(__name__)


# class storing individual items in schedule of work
class ScheduleItemGeneric:
    """Class stores a generic schedule item"""
    
    def __init__(self, item=[]):
        self.item = item

    def set_model(self, item):
        self.item = item

    def get_model(self):
        return self.item

    def __setitem__(self, index, value):
        self.item[index] = value

    def __getitem__(self, index):
        return self.item[index]

    def print_item(self):
        print(self.item)
        

class ScheduleItem(ScheduleItemGeneric):
    """Class stores a row in the schedule of rates"""
    
    def __init__(self, itemno="", description="", unit="", rate="0", qty="0", reference="", excess_rate_percent="30"):
        super(ScheduleItem, self).__init__([itemno, description, unit, rate, qty, reference, excess_rate_percent])
        self.itemno = itemno
        self.description = description
        self.unit = unit
        self.rate = rate
        self.qty = qty
        self.reference = reference
        self.excess_rate_percent = excess_rate_percent
        # extended descritption of item
        self.extended_description = ''
        self.extended_description_limited = ''


class ScheduleGeneric:
    """Class stores a generic schedule"""
    
    def __init__(self, items=[]):
        self.items = items  # main data store of rows

    def append_item(self, item):
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
        items = []
        for item in self.items:
            items.append(item.get_model())
        return self.items
        
    def set_model(self,items):
        for item in items:
            self.append(ScheduleItemGeneric(item))
        return self.items

    def length(self):
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
    
    def __init__(self, items=[]):
        #Initialise base class
        super(Schedule,self).__init__(items)
        
    def update_values(self):
        # Populate ScheduleItem.extended_description (Used for final billing)
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
                item.extended_description_limited = item.extended_description[0:int(CMB_DESCRIPTION_MAX_LENGTH/2)] + \
                    ' ... ' + item.extended_description[-int(CMB_DESCRIPTION_MAX_LENGTH/2):]
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
            
    def get_itemnos(self):
        """Returns a list of itemnos with order as in schedule"""
        itemnos = []
        for item in self.items:
            if item.itemno != '' and item.unit != '' and item.qty != 0:
                itemnos.append(item.itemno)
        return itemnos

