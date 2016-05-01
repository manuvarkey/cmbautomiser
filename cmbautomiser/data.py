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


class ScheduleGeneric:
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
            items.append(item.get())
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
        
