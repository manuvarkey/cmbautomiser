#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#  cmb.py
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

import pickle
import os.path
import copy

from undo import *

# local files import
from globalconstants import *
from schedule import *
from schedule_dialog import ScheduleDialog

from misc import *

# class storing Cmb
class Cmb:
    def __init__(self,name_ = ""):
        self.name = name_
        self.items = []

    def append_item(self,item):
        item.set_cmb(self)
        self.items.append(item)
                
    def insert_item(self,index,item):
        item.set_cmb(self)
        self.items.insert(index,item)
        
    def remove_item(self,index):
        del(self.items[index])

    def get_latex_buffer(self,path):
        file_preamble = open(abs_path('latex','preamble.tex'),'r')
        latex_buffer =file_preamble.read()
        cmb_local_vars = {}
        cmb_local_vars['$cmbbookno$'] = self.name
        cmb_local_vars['$cmbheading$'] = ' '
        cmb_local_vars['$cmbtitle$'] = 'DETAILS OF MEASUREMENTS'
        cmb_local_vars['$cmbstartingpage$'] = str(1)
        latex_buffer = replace_all(latex_buffer,cmb_local_vars)
        for count,item in enumerate(self.items):
            newpath = list(path) + [count]
            latex_buffer += item.get_latex_buffer(newpath)
        file_end = open(abs_path('latex','end.tex'),'r')
        latex_buffer += file_end.read()
        file_preamble.close()
        file_end.close()
        return latex_buffer

    def __setitem__(self, index, value):
        self.items[index] = value
    
    def __getitem__(self, index):
        return self.items[index]
        
    def set_name(self,name):
        self.name = name
        
    def get_name(self):
        return self.name
        
    def get_schedule(self):
        return self.schedule

    def length(self):
        return len(self.items)
        
    def clear(self):
        self.items = []
        
    def get_text(self):
        return "<b>CMB No." + clean_markup(self.name) + "</b>"
    
    def get_tooltip(self):
        return None
    
    def print_item(self):
        print "CMB " + self.name
        for item in self.items:
            item.print_item()
        
# class storing Measurement groups
class Measurement:
    def __init__(self,date = ""):
        self.cmb = None
        self.date = date
        self.items = []

    def append_item(self,item):
        item.set_measurement(self)
        self.items.append(item)
                
    def insert_item(self,index,item):
        item.set_measurement(self)
        self.items.insert(index,item)
        
    def remove_item(self,index):
        del(self.items[index])
       
    def get_latex_buffer(self,path):
        file_measgroup = open(abs_path('latex','measgroup.tex'),'r')
        latex_buffer = file_measgroup.read()
        file_measgroup.close()
        
        # replace local variables
        measgroup_local_vars = {}
        measgroup_local_vars['$cmbmeasurementdate$'] = self.date
        latex_buffer = replace_all(latex_buffer,measgroup_local_vars)
        for count,item in enumerate(self.items):
            newpath = list(path) + [count]
            latex_buffer += item.get_latex_buffer(newpath)
        return latex_buffer

    def __setitem__(self, index, value):
        self.items[index] = value
    
    def __getitem__(self, index):
        return self.items[index]
        
    def set_date(self,date):
        self.date = date
        
    def get_date(self):
        return self.date

    def length(self):
        return len(self.items)
        
    def clear(self):
        self.items = []
    
    def get_text(self):
        return "<b>Measurement dated." + clean_markup(self.date) + "</b>"
    
    def get_tooltip(self):
        return None
        
    def get_cmb(self):
        return self.cmb
        
    def set_cmb(self,cmb_):
        self.cmb = cmb_
    
    def print_item(self):
        print "  " + "Measurement dated " + self.date
        for item in self.items:
            item.print_item()
        
# BaseClass for storing Measurement items
class MeasurementItem:
    def __init__(self,itemnos=[],items=[],records=[],remark="",item_remarks=[]):
        self.measurement = None
        self.itemnos = itemnos
        self.items = items
        self.records = records
        self.remark = remark
        self.item_remarks = item_remarks
        self.billed = False # whether item is billed

    def set_item(self,index,itemno,item):
        self.itemnos[index] = itemno
        self.items[index] = item
        
    def get_item(self,index):
        return self.items[index]
        
    def append_record(self,record):
        self.records.append(record)
                
    def insert_record(self,index,record):
        self.records.insert(index,record)
        
    def remove_record(self,index):
        del(self.records[index])
        
    def __setitem__(self, index, value):
        self.records[index] = value
    
    def __getitem__(self, index):
        return self.records[index]
        
    def set_remark(self,remark):
        self.remark = remark
        
    def get_remark(self):
        return self.remark

    def length(self):
        return len(self.records)
        
    def clear(self):
        self.measurement = None
        self.records = []
        self.remark = ""
        
    def get_cmb(self):
        return self.measurement.cmb
        
    def get_measurement(self):
        return self.measurement
        
    def set_measurement(self,measurement):
        self.measurement = measurement
    
    def set_billed_flag(self,flag=True):
        self.billed = flag

    def get_billed_flag(self):
        return self.billed
        
class MeasurementItemHeading(MeasurementItem):
    def __init__(self,remark):
        MeasurementItem.__init__(self,itemnos=[],items=[],records=[],remark=remark,item_remarks = None)

    def print_item(self):
        print "    " + self.remark
        
    def get_latex_buffer(self,path):
        file_measheading = open(abs_path('latex','measheading.tex'),'r')
        latex_buffer = file_measheading.read()
        file_measheading.close()
        
        # replace local variables
        measheading_local_vars = {}
        measheading_local_vars['$cmbmeasurementheading$'] = self.remark
        latex_buffer = replace_all(latex_buffer,measheading_local_vars)
        
        return latex_buffer
    
    def get_text(self):
        return "<b><i>" + clean_markup(self.remark) + "</i></b>"
    
    def get_tooltip(self):
        return None

class RecordNLBH:
    def __init__(self,desc,n="",l="",b="",h=""):
        self.desc = desc
        self.data_str = [n,l,b,h]
        self.data = []
        for x in self.data_str:
            try:
                num = eval(x)
                self.data.append(num)
            except:
                self.data.append(0)
        self.breakup = self.find_breakup()
        self.total = self.find_total()
        
    def get(self):
        return [self.desc,self.breakup,self.data,self.total]
        
    def get_model(self):
        data = []
        data.append(self.desc) # add description
        data.append("") # add breakup
        data += self.data_str
        data.append("") # for total column
        return data

    def find_total(self):
        nonzero = [x for x in self.data if x!= 0]
        total = 1
        for x in nonzero:
            total *= x
        if len(nonzero) == 0:
            return 0
        else:
            return total
    
    def find_breakup(self):
        breakup = "["
        for x in self.data_str:
            if x != "" and x!= '0':
                breakup = breakup + str(x) + ","
            else:
                breakup = breakup + ','      
        breakup = breakup[:-1] + "]"
        return breakup
    
    def print_item(self):
        print "      " + str([self.desc,self.breakup,self.data,self.total])
        
# record format [description,breakup,no,l,b,h,total]
class MeasurementItemNLBH(MeasurementItem):
    def __init__(self,data = None):
        if data != None:
            itemnos = data[0]
            items = data[1]
            records = data[2]
            remark = data[3]
            item_remarks = data[4]
            MeasurementItem.__init__(self,itemnos,items,records,remark,item_remarks)
        else:
            MeasurementItem.__init__(self)
            
    def get_model(self):
        item_schedule = []
        for item in self.records:
            item_schedule.append(item.get_model())
        data = [self.itemnos, self.items, item_schedule, self.remark, self.item_remarks]
        return data
        
    def set_model(self,data):
        self.clear()
        self.itemnos = data[0]
        self.items = data[1]
        for item in data[2]:
            self.append_record(RecordNLBH(item[0],item[2],item[3],item[4],item[5]))
        self.remark = data[3]
        self.item_remarks = data[4]
        
    def get_latex_buffer(self,path):
        file_measnlbh = open(abs_path('latex','measnlbh.tex'),'r')
        latex_buffer = file_measnlbh.read()
        file_measnlbh.close()
        # read part latex code for records
        latex_records = ''
        data_str = [None]*4
        for slno,record in enumerate(self.records):
            for i in range(4): # evaluate string of data entries, suppress zero.
                data_str[i] = str(record.data[i]) if record.data[i] != 0 else ''
            latex_records += '        ' + str(slno+1) + ' & ' + clean_latex(record.desc) + ' & ' + '\\AddBreakableChars{'+record.breakup+'}' + ' & ' + data_str[0] + ' & ' + \
                            data_str[1] + ' & ' + data_str[2] + ' & ' + data_str[3] + ' & ' + str(record.total) + '\\\\ \n'
            
        # replace local variables
        measnlbh_local_vars = {}
        try:
            measnlbh_local_vars['$cmbitemno$'] = str(self.itemnos[0])
            measnlbh_local_vars['$cmbitemdesc$'] = str(self.items[0].extended_description)
        except:
            measnlbh_local_vars['$cmbitemno$'] = 'ERROR'
            measnlbh_local_vars['$cmbitemdesc$'] = ''
        measnlbh_local_vars['$cmbtotal$'] = str(sum(self.get_total()))
        measnlbh_local_vars['$cmbcarriedover$'] = 'ref:abs:'+str(path) + ':1'
        measnlbh_local_vars['$cmblabel$'] = 'ref:meas:'+str(path) + ':1'
        measnlbh_local_vars['$cmbremark$'] = str(self.remark)
        latex_buffer = replace_all(latex_buffer,measnlbh_local_vars)
        
        # fill in records - vanilla function used since latex_records contains latex code
        measnlbh_local_vars_vannilla = {}
        measnlbh_local_vars_vannilla['$cmbrecords$'] = latex_records
        latex_buffer = replace_all_vanilla(latex_buffer,measnlbh_local_vars_vannilla)
        
        return latex_buffer
                
    def print_item(self):
        print "    Item No." + str(self.itemnos[0])
        for i in range(self.length()):
            self[i].print_item()
        print "    " + "Total: " + str(self.get_total())
    
    def get_total(self):
        total = 0
        for i in range(self.length()):
            total += self.records[i].total
        return [total]
    
    def get_text(self):
        total = ['{:.1f}'.format(x) for x in self.get_total()]
        return "Item No.<b>" + str(self.itemnos) + "    |NLBH|</b>    # of records: <b>" + str(self.length()) + "</b>, Total: <b>" + str(total) + "</b>"

    def get_tooltip(self):
        if self.remark != "":
            return "Remark: " + self.remark
        else:
            return None

class RecordLLLLL:
    def __init__(self,desc,l1='',l2='',l3='',l4='',l5=''):
        self.desc = desc
        self.data_str = [l1,l2,l3,l4,l5]
        self.data = []
        for x in self.data_str:
            try:
                num = eval(x)
                self.data.append(num)
            except:
                self.data.append(0)
        self.breakup = self.find_breakup()
        
    def get(self):
        return [self.desc,self.breakup,self.data]
        
    def get_model(self):
        data = []
        data.append(self.desc) # add description
        data.append("") # add breakup
        data += self.data_str
        return data
    
    def find_breakup(self):
        breakup = "["
        for x in self.data_str:
            if x != "" and x != 0:
                breakup = breakup + str(x) + ","
            else:
                breakup = breakup + ','            
        breakup = breakup[:-1] + "]"
        return breakup
    
    def print_item(self):
        print "      " + str([self.desc,self.breakup,self.data])
        
# record format [description,breakup,l1,l2,l3,l4,l5]
class MeasurementItemLLLLL(MeasurementItem):
    def __init__(self,data = None):
        if data != None:
            itemnos = data[0]
            items = data[1]
            records = data[2]
            remark = data[3]
            item_remarks = data[4]
            MeasurementItem.__init__(self,itemnos,items,records,remark,item_remarks)
        else:
            MeasurementItem.__init__(self)
            
    def get_model(self):
        item_schedule = []
        for item in self.records:
            item_schedule.append(item.get_model())
        data = [self.itemnos, self.items, item_schedule, self.remark, self.item_remarks]
        return data
        
    def set_model(self,data):
        self.clear()
        self.itemnos = data[0]
        self.items = data[1]
        for item in data[2]:
            self.append_record(RecordLLLLL(item[0],item[2],item[3],item[4],item[5],item[6]))
        self.remark = data[3]
        self.item_remarks = data[4]
        
    def get_latex_buffer(self,path):
        file_measlllll = open(abs_path('latex','measlllll.tex'),'r')
        latex_buffer = file_measlllll.read()
        file_measlllll.close()
        # read part latex code for records
        latex_records = ''
        data_str = [None]*5
        for slno,record in enumerate(self.records):
            for i in range(5): # evaluate string of data entries, suppress zero.
                data_str[i] = str(record.data[i]) if record.data[i] != 0 else ''
            latex_records += '        ' + str(slno+1) + ' & ' + clean_latex(record.desc) + ' & ' + '\\AddBreakableChars{'+record.breakup+'}' + ' & ' + data_str[0] + ' & ' + \
                            data_str[1] + ' & ' + data_str[2] + ' & ' + data_str[3] + ' & ' + data_str[4] + '\\\\ \n'
            
        # replace local variables
        measlllll_local_vars = {}
        measlllll_local_vars_vannilla = {}
        for i in range(0,5):
            try:
                measlllll_local_vars['$cmbitemdesc' + str(i+1) + '$'] = str(self.items[i].extended_description)
                measlllll_local_vars['$cmbitemno' + str(i+1) + '$'] = str(self.itemnos[i])
                measlllll_local_vars['$cmbtotal' + str(i+1) + '$'] = str(self.get_total()[i])
                measlllll_local_vars['$cmbitemremark' + str(i+1) + '$'] = str(self.item_remarks[i])
                measlllll_local_vars['$cmbcarriedover' + str(i+1) + '$'] = 'ref:abs:'+str(path)+':'+str(i+1)
                measlllll_local_vars['$cmblabel' + str(i+1) + '$'] = 'ref:meas:'+str(path)+':'+str(i+1)
                measlllll_local_vars_vannilla['$cmbitemexist' + str(i+1) + '$'] = '\\iftrue'
            except:
                measlllll_local_vars['$cmbitemdesc' + str(i+1) + '$'] = ''
                measlllll_local_vars['$cmbitemno' + str(i+1) + '$'] = 'ERROR'
                measlllll_local_vars['$cmbtotal' + str(i+1) + '$'] = ''
                measlllll_local_vars['$cmbitemremark' + str(i+1) + '$'] = ''
                measlllll_local_vars_vannilla['$cmbitemexist' + str(i+1) + '$'] = '\\iffalse'
        measlllll_local_vars['$cmbremark$'] = str(self.remark)
        latex_buffer = replace_all(latex_buffer,measlllll_local_vars)
        
        # fill in records - vanilla function used since latex_records contains latex code
        measlllll_local_vars_vannilla['$cmbrecords$'] = latex_records
        latex_buffer = replace_all_vanilla(latex_buffer,measlllll_local_vars_vannilla)
        
        return latex_buffer
                
    def print_item(self):
        print "    Item No." + str(self.itemnos[0])
        for i in range(self.length()):
            self[i].print_item()
        print "    " + "Total: " + str(self.get_total())
    
    def get_total(self):
        total = [0]*5
        for i in range(self.length()):
            for j in range(0,5):
                total[j] += self.records[i].data[j]
        return total
    
    def get_text(self):
        total = ['{:.1f}'.format(x) for x in self.get_total()]
        return "Item No.<b>" + str(self.itemnos) + "    |LLLLL|</b>    # of records: <b>" + str(self.length()) + "</b>, Total: <b>" + str(total) + "</b>"
    
    def get_tooltip(self):
        if self.remark != "":
            return "Remark: " + self.remark
        else:
            return None

class RecordNNNNNNNN:
    def __init__(self,desc,n1='',n2='',n3='',n4='',n5='',n6='',n7='',n8=''):
        self.desc = desc
        self.data_str = [n1,n2,n3,n4,n5,n6,n7,n8]
        self.data = []
        for x in self.data_str:
            try:
                num = eval(x)
                self.data.append(num)
            except:
                self.data.append(0)
        
    def get(self):
        return [self.desc,self.data]
        
    def get_model(self):
        data = []
        data.append(self.desc) # add description
        data += self.data_str
        return data
    
    def print_item(self):
        print "      " + str([self.desc,self.data])
        
# record format [description,n1,n2,n3,n4,n5,n6,n7,n8]
class MeasurementItemNNNNNNNN(MeasurementItem):
    def __init__(self,data = None):
        if data != None:
            itemnos = data[0]
            items = data[1]
            records = data[2]
            remark = data[3]
            item_remarks = data[4]
            MeasurementItem.__init__(self,itemnos,items,records,remark,item_remarks)
        else:
            MeasurementItem.__init__(self)
            
    def get_model(self):
        item_schedule = []
        for item in self.records:
            item_schedule.append(item.get_model())
        data = [self.itemnos, self.items, item_schedule, self.remark, self.item_remarks]
        return data
        
    def set_model(self,data):
        self.clear()
        self.itemnos = data[0]
        self.items = data[1]
        for item in data[2]:
            self.append_record(RecordNNNNNNNN(item[0],item[1],item[2],item[3],item[4],item[5],item[6],item[7],item[8]))
        self.remark = data[3]
        self.item_remarks = data[4]
        
    def get_latex_buffer(self,path):
        file_measnnnnnnnn = open(abs_path('latex','measnnnnnnnn.tex'),'r')
        latex_buffer = file_measnnnnnnnn.read()
        file_measnnnnnnnn.close()
        # read part latex code for records
        latex_records = ''
        data_str = [None]*8
        for slno,record in enumerate(self.records):
            for i in range(8): # evaluate string of data entries, suppress zero.
                data_str[i] = str(record.data[i]) if record.data[i] != 0 else ''
            latex_records += '        ' + str(slno+1) + ' & ' + clean_latex(record.desc) + ' & ' + data_str[0] + ' & ' + \
                            data_str[1] + ' & ' + data_str[2] + ' & ' + data_str[3] + ' & ' + data_str[4] + '&' + \
                            data_str[5] + ' & ' + data_str[6] + ' & ' + data_str[7] + '\\\\ \n'
            
        # replace local variables
        measnnnnnnnn_local_vars = {}
        measnnnnnnnn_local_vars_vannilla = {}
        for i in range(0,8):
            try:
                measnnnnnnnn_local_vars['$cmbitemdesc' + str(i+1) + '$'] = str(self.items[i].extended_description)
                measnnnnnnnn_local_vars['$cmbitemno' + str(i+1) + '$'] = str(self.itemnos[i])
                measnnnnnnnn_local_vars['$cmbtotal' + str(i+1) + '$'] = str(self.get_total()[i])
                measnnnnnnnn_local_vars['$cmbitemremark' + str(i+1) + '$'] = str(self.item_remarks[i])
                measnnnnnnnn_local_vars['$cmbcarriedover' + str(i+1) + '$'] = 'ref:abs:'+str(path)+':'+str(i+1)
                measnnnnnnnn_local_vars['$cmblabel' + str(i+1) + '$'] = 'ref:meas:'+str(path)+':'+str(i+1)
                measnnnnnnnn_local_vars_vannilla['$cmbitemexist' + str(i+1) + '$'] = '\\iftrue'
            except:
                measnnnnnnnn_local_vars['$cmbitemdesc' + str(i+1) + '$'] = ''
                measnnnnnnnn_local_vars['$cmbitemno' + str(i+1) + '$'] = 'ERROR'
                measnnnnnnnn_local_vars['$cmbtotal' + str(i+1) + '$'] = ''
                measnnnnnnnn_local_vars['$cmbitemremark' + str(i+1) + '$'] = ''
                measnnnnnnnn_local_vars_vannilla['$cmbitemexist' + str(i+1) + '$'] = '\\iffalse'
        measnnnnnnnn_local_vars['$cmbremark$'] = str(self.remark)
        latex_buffer = replace_all(latex_buffer,measnnnnnnnn_local_vars)
        
        # fill in records - vanilla function used since latex_records contains latex code
        measnnnnnnnn_local_vars_vannilla['$cmbrecords$'] = latex_records
        latex_buffer = replace_all_vanilla(latex_buffer,measnnnnnnnn_local_vars_vannilla)
        
        return latex_buffer
                
    def print_item(self):
        print "    Item No." + str(self.itemnos[0])
        for i in range(self.length()):
            self[i].print_item()
        print "    " + "Total: " + str(self.get_total())
    
    def get_total(self):
        total = [0]*8
        for i in range(self.length()):
            for j in range(0,8):
                total[j] += self.records[i].data[j]
        return total
    
    def get_text(self):
        total = ['{:.1f}'.format(x) for x in self.get_total()]
        return "Item No.<b>" + str(self.itemnos) + "    |NNNNNNNN|</b>    # of records: <b>" + str(self.length()) + "</b>, Total: <b>" + str(total) + "</b>"
    
    def get_tooltip(self):
        if self.remark != "":
            return "Remark: " + self.remark
        else:
            return None


class RecordnnnnnT:
    def __init__(self,desc,n1="",n2="",n3="",n4="",n5=""):
        self.desc = desc
        self.data_str = [n1,n2,n3,n4,n5]
        self.data = []
        for x in self.data_str:
            try:
                num = eval(x)
                self.data.append(num)
            except:
                self.data.append(0)
        self.total = self.find_total()
        
    def get(self):
        return [self.desc,self.data,self.total]
        
    def get_model(self):
        data = []
        data.append(self.desc) # add description
        data += self.data_str
        data.append("") # for total column
        return data

    def find_total(self):
        total = sum(self.data)
        return total
    
    def print_item(self):
        print "      " + str([self.desc,self.data,self.total])
        
# record format [description,n1,n2,n3,n4,n5,total]
class MeasurementItemnnnnnT(MeasurementItem):
    def __init__(self,data = None):
        if data != None:
            itemnos = data[0]
            items = data[1]
            records = data[2]
            remark = data[3]
            item_remarks = data[4]
            MeasurementItem.__init__(self,itemnos,items,records,remark,item_remarks)
        else:
            MeasurementItem.__init__(self)
            
    def get_model(self):
        item_schedule = []
        for item in self.records:
            item_schedule.append(item.get_model())
        data = [self.itemnos, self.items, item_schedule, self.remark, self.item_remarks]
        return data
        
    def set_model(self,data):
        self.clear()
        self.itemnos = data[0]
        self.items = data[1]
        for item in data[2]:
            self.append_record(RecordnnnnnT(item[0],item[1],item[2],item[3],item[4],item[5]))
        self.remark = data[3]
        self.item_remarks = data[4]
        
    def get_latex_buffer(self,path):
        file_measnnnnnt = open(abs_path('latex','measnnnnnt.tex'),'r')
        latex_buffer = file_measnnnnnt.read()
        file_measnnnnnt.close()
        # read part latex code for records
        latex_records = ''
        data_str = [None]*5
        for slno,record in enumerate(self.records):
            for i in range(5): # evaluate string of data entries, suppress zero.
                data_str[i] = str(record.data[i]) if record.data[i] != 0 else ''
            latex_records += '        ' + str(slno+1) + ' & ' + clean_latex(record.desc) + ' & ' + data_str[0] + ' & ' + \
                            data_str[1] + ' & ' + data_str[2] + ' & ' + data_str[3] + ' & ' + data_str[4] + ' & ' + str(record.total) + '\\\\ \n'
            
        # replace local variables
        measnnnnnt_local_vars = {}
        try:
            measnnnnnt_local_vars['$cmbitemno$'] = str(self.itemnos[0])
            measnnnnnt_local_vars['$cmbitemdesc$'] = str(self.items[0].extended_description)
        except:
            measnnnnnt_local_vars['$cmbitemno$'] = 'ERROR'
            measnnnnnt_local_vars['$cmbitemdesc$'] = ''
        measnnnnnt_local_vars['$cmbtotal$'] = str(sum(self.get_total()))
        measnnnnnt_local_vars['$cmbcarriedover$'] = 'ref:abs:'+str(path) + ':1'
        measnnnnnt_local_vars['$cmblabel$'] = 'ref:meas:'+str(path) + ':1'
        measnnnnnt_local_vars['$cmbremark$'] = str(self.remark)
        latex_buffer = replace_all(latex_buffer,measnnnnnt_local_vars)
        
        # fill in records - vanilla function used since latex_records contains latex code
        measnnnnnt_local_vars_vannilla = {}
        measnnnnnt_local_vars_vannilla['$cmbrecords$'] = latex_records
        latex_buffer = replace_all_vanilla(latex_buffer,measnnnnnt_local_vars_vannilla)
        
        return latex_buffer
                
    def print_item(self):
        print "    Item No." + str(self.itemnos[0])
        for i in range(self.length()):
            self[i].print_item()
        print "    " + "Total: " + str(self.get_total())
    
    def get_total(self):
        total = 0
        for i in range(self.length()):
            total += self.records[i].total
        return [total]
    
    def get_text(self):
        total = ['{:.1f}'.format(x) for x in self.get_total()]
        return "Item No.<b>" + str(self.itemnos) + "    |nnnnnT|</b>    # of records: <b>" + str(self.length()) + "</b>, Total: <b>" + str(total) + "</b>"
    
    def get_tooltip(self):
        if self.remark != "":
            return "Remark: " + self.remark
        else:
            return None
            
class RecordCustom:
    def __init__(self,items,cust_funcs,total_func,columntypes):
        self.data_str = items
        self.data = []
        # Populate Data
        for x,columntype in zip(self.data_str,columntypes):
            if columntype not in [MEAS_DESC,MEAS_CUST]:
                try:
                    num = eval(x)
                    self.data.append(num)
                except:
                    self.data.append(0)
            else:
                self.data.append(0)
        self.cust_funcs = cust_funcs
        self.total_func = total_func
        self.columntypes = columntypes
        self.total = self.find_total()

    def get(self):
        iter_d = 0
        skip = 0
        data = []
        for columntype in self.columntypes:
            if columntype == MEAS_CUST:
                data.append('')
                skip += 1
            else:
                data.append(self.data_str[iter_d-skip])
            iter_d += 1
        return data

    def get_model(self):
        return self.data_str

    def find_total(self):
        return self.total_func(self.data)

    def find_custom(self,index):
        return self.cust_funcs[index](self.data)

    def print_item(self):
        print "      " + str([self.data_str,self.total])


# record format [As per file read]
class MeasurementItemCustom(MeasurementItem):
    def __init__(self,data = None,plugin=None):
        self.name = ''
        self.itemtype = None
        self.itemnos_mask = []
        self.captions = []
        self.columntypes = []
        self.cust_funcs = []
        self.total_func_item = None
        self.total_func = None
        self.latex_item = ''
        self.latex_record = ''
        # For user data support
        self.captions_udata = []
        self.columntypes_udata = []
        self.user_data = None
        self.latex_postproc_func = None
        self.export_abstract = None

        # Read description from file
        if plugin is not None:
            try:
                package = __import__('templates.' + plugin)
                module = getattr(package, plugin)
                self.custom_object = module.CustomItem()
                self.name = self.custom_object.name
                self.itemtype = plugin
                self.itemnos_mask = self.custom_object.itemnos_mask
                self.captions = self.custom_object.captions
                self.columntypes = self.custom_object.columntypes
                self.cust_funcs = self.custom_object.cust_funcs
                self.total_func_item = self.custom_object.total_func_item
                self.total_func = self.custom_object.total_func
                self.latex_item = self.custom_object.latex_item
                self.latex_record = self.custom_object.latex_record
                # For user data support
                self.captions_udata = self.custom_object.captions_udata
                self.columntypes_udata = self.custom_object.columntypes_udata
                self.latex_postproc_func = self.custom_object.latex_postproc_func
                self.user_data = self.custom_object.user_data_default
                self.export_abstract = self.custom_object.export_abstract
            except ImportError:
                print('Error Loading plugin')

            if data != None:
                itemnos = data[0]
                items = data[1]
                records = data[2]
                remark = data[3]
                item_remarks = data[4]
                self.user_data = data[5]
                MeasurementItem.__init__(self,itemnos,items,records,remark,item_remarks)
            else:
                MeasurementItem.__init__(self,[None]*self.item_width(),[None]*self.item_width(),[],
                                         '',['']*self.item_width())
        else:
            MeasurementItem.__init__(self,[None]*self.item_width(),[None]*self.item_width(),[],
                                         '',['']*self.item_width())

    def model_width(self):
        return len(self.columntypes)

    def item_width(self):
        return len(self.itemnos_mask)

    def get_model(self):
        item_schedule = []
        for item in self.records:
            item_schedule.append(item.get_model())
        data = [self.itemnos, self.items, item_schedule, self.remark, self.item_remarks,
                 self.user_data, self.itemtype]
        return data

    def set_model(self,data):
        self.clear()
        self.itemnos = data[0]
        self.items = data[1]
        for item in data[2]:
            if item is not None:
                self.append_record(RecordCustom(item,
                               self.cust_funcs,self.total_func_item,self.columntypes))
        self.remark = data[3]
        self.item_remarks = data[4]
        self.user_data = data[5]
        self.itemtype = data[6]
        if self.itemtype is not None:
            try:
                package = __import__('templates.' + self.itemtype)
                module = getattr(package, self.itemtype)
                self.custom_object = module.CustomItem()
                self.name = self.custom_object.name
                self.itemnos_mask = self.custom_object.itemnos_mask
                self.captions = self.custom_object.captions
                self.columntypes = self.custom_object.columntypes
                self.cust_funcs = self.custom_object.cust_funcs
                self.total_func_item = self.custom_object.total_func_item
                self.total_func = self.custom_object.total_func
                self.latex_item = self.custom_object.latex_item
                self.latex_record = self.custom_object.latex_record
                # For user data support
                self.captions_udata = self.custom_object.captions_udata
                self.columntypes_udata = self.custom_object.columntypes_udata
                self.latex_postproc_func = self.custom_object.latex_postproc_func
                self.export_abstract = self.custom_object.export_abstract
            except ImportError:
                print('Error Loading plugin')

    def get_latex_buffer(self,path,isabstract=False):
        # read part latex code for records
        latex_records = ''
        data_str = [None]*self.model_width()
        for slno,record in enumerate(self.records):
            meascustom_rec_vars = {}
            meascustom_rec_vars_van = {}
            # Evaluate string to make replacement
            cust_iter = 0
            for i,columntype in enumerate(self.columntypes): # evaluate string of data entries, suppress zero.
                if columntype == MEAS_CUST:
                    try:
                        value =  str(record.cust_funcs[cust_iter](record.get(),slno))
                        data_str[i] = value if value not in ['0','0.0'] else ''
                    except:
                        data_str[i] = ''
                    cust_iter += 1
                elif columntype == MEAS_DESC:
                    try:
                        data_str[i] = str(record.data_str[i-cust_iter])
                    except:
                        data_str[i] = ''
                else:
                    try:
                        data_str[i] = str(record.data[i-cust_iter]) if record.data[i-cust_iter] != 0 else ''
                    except:
                        data_str[i] = ''
                # Check for carry over item possibly contains code
                if columntype == MEAS_DESC and data_str[i].find('Qty B/F') != -1:
                    meascustom_rec_vars_van['$data' + str(i+1) + '$'] = data_str[i]
                else:
                    meascustom_rec_vars['$data' + str(i+1) + '$'] = data_str[i]
            meascustom_rec_vars['$slno$'] = str(slno+1)
            latex_record = self.latex_record[:]
            latex_record = replace_all_vanilla(latex_record,meascustom_rec_vars_van)
            latex_record = replace_all(latex_record,meascustom_rec_vars)
            latex_records += latex_record
            
        # replace local variables
        meascustom_local_vars = {}
        meascustom_local_vars_vannilla = {}
        for i in range(0,self.item_width()):
            try:
                meascustom_local_vars['$cmbitemdesc' + str(i+1) + '$'] = str(self.items[i].extended_description)
                meascustom_local_vars['$cmbitemno' + str(i+1) + '$'] = str(self.itemnos[i])
                meascustom_local_vars['$cmbtotal' + str(i+1) + '$'] = str(self.get_total()[i])
                meascustom_local_vars['$cmbitemremark' + str(i+1) + '$'] = str(self.item_remarks[i])
                meascustom_local_vars['$cmbcarriedover' + str(i+1) + '$'] = 'ref:abs:'+str(path)+':'+str(i+1)
                meascustom_local_vars['$cmblabel' + str(i+1) + '$'] = 'ref:meas:'+str(path)+':'+str(i+1)
                meascustom_local_vars_vannilla['$cmbitemexist' + str(i+1) + '$'] = '\\iftrue'
            except:
                meascustom_local_vars['$cmbitemdesc' + str(i+1) + '$'] = ''
                meascustom_local_vars['$cmbitemno' + str(i+1) + '$'] = 'ERROR'
                meascustom_local_vars['$cmbtotal' + str(i+1) + '$'] = ''
                meascustom_local_vars['$cmbitemremark' + str(i+1) + '$'] = ''
                meascustom_local_vars_vannilla['$cmbitemexist' + str(i+1) + '$'] = '\\iffalse'
        meascustom_local_vars['$cmbremark$'] = str(self.remark)
        if isabstract:
            meascustom_local_vars_vannilla['$cmbasbstractitem$'] = '\\iftrue'
        else:
            meascustom_local_vars_vannilla['$cmbasbstractitem$'] = '\\iffalse'
        latex_buffer = self.latex_item[:]
        latex_buffer = replace_all(latex_buffer,meascustom_local_vars)

        # fill in records - vanilla function used since latex_records contains latex code
        meascustom_local_vars_vannilla['$cmbrecords$'] = latex_records
        latex_buffer = replace_all_vanilla(latex_buffer,meascustom_local_vars_vannilla)
        return self.latex_postproc_func(self.records,self.user_data,latex_buffer,isabstract)

    def print_item(self):
        print "    Item No." + str(self.itemnos)
        for i in range(self.length()):
            self[i].print_item()
        print "    " + "Total: " + str(self.get_total())

    def get_total(self):
        if self.total_func is not None:
            return self.total_func(self.records,self.user_data)
        else:
            return []

    def get_text(self):
        total = ['{:.1f}'.format(x) for x in self.get_total()]
        return "Item No.<b>" + str(self.itemnos) + "    |Custom: " + self.name + "|</b>    # of records: <b>" + \
               str(self.length()) + "</b>, Total: <b>" + str(total) + "</b>"

    def get_tooltip(self):
        if self.remark != "":
            return "Remark: " + self.remark
        else:
            return None

# Abstract of measurement
class MeasurementItemAbstract(MeasurementItem):
    def __init__(self,data = None):
        self.int_m_item = None # used to store all records and stuff
        self.m_items = None

        if data is not None:
            self.set_model(data)
        MeasurementItem.__init__(self,itemnos=[],items=[],records=[],remark='',item_remarks = [])

    def get_model(self):
        model = None
        itemtype = None
        if self.int_m_item is not None:
            model = self.int_m_item.get_model()
            itemtype = self.int_m_item.itemtype
        data = [self.m_items,model,itemtype]
        return data

    def set_model(self,data):
        self.int_m_item = None # used to store all records and stuff
        self.m_items = None

        if data is not None:
            self.m_items = data[0]
            if self.m_items is not None:
                self.int_m_item = MeasurementItemCustom(None,data[2])
                self.int_m_item.set_model(data[1])
                MeasurementItem.__init__(self,self.int_m_item.itemnos,self.int_m_item.items,self.int_m_item.records,
                              self.int_m_item.remark,self.int_m_item.item_remarks)

    def get_latex_buffer(self,path):
        if self.m_items is not None:
            return self.int_m_item.get_latex_buffer(path,True)

    def print_item(self):
        print '    Abstract Item'
        self.int_m_item.print_item()

    def get_total(self):
        if self.int_m_item is not None:
            return self.int_m_item.get_total()
        else:
            return []

    def get_text(self):
        if self.int_m_item is not None:
            return 'Abs: ' + self.int_m_item.get_text()
        else:
            return 'Abs: NOT DEFINED'

    def get_tooltip(self):
        if self.int_m_item is not None:
            if self.int_m_item.get_tooltip() is not None:
                return 'Abs: ' + self.int_m_item.get_tooltip()

# class storing Measurement groups
class Completion:
    def __init__(self,date = "",remark = ""):
        self.cmb = None
        self.date = date

    def get_latex_buffer(self,path):
        file_meascompletion = open(abs_path('latex','meascompletion.tex'),'r')
        latex_buffer = file_meascompletion.read()
        file_meascompletion.close()
        # replace local variables
        measgroup_local_vars = {}
        measgroup_local_vars['$cmbcompletiondate$'] = self.date
        latex_buffer = replace_all(latex_buffer,measgroup_local_vars)
        return latex_buffer

    def get_cmb(self):
        return self.cmb

    def set_cmb(self,cmb_):
        self.cmb = cmb_

    def set_date(self,date):
        self.date = date

    def get_date(self):
        return self.date

    def get_text(self):
        return "<b>Completion recorded on " + clean_markup(self.date) + "</b>"

    def get_tooltip(self):
        return None

    def print_item(self):
        print "  " + "Completion recorded on " + self.date

# Measurements object
class MeasurementsView:
    # for CMB, measurement and Heading
    def get_user_input_text(self, parent, message, title='',value=None):
        # Returns user input as a string or None
        # If user does not input text it returns None, NOT AN EMPTY STRING.
        dialogWindow = Gtk.MessageDialog(parent,
                              Gtk.DialogFlags.MODAL | Gtk.DialogFlags.DESTROY_WITH_PARENT,
                              Gtk.MessageType.QUESTION,
                              Gtk.ButtonsType.OK_CANCEL,
                              message)
        
        dialogWindow.set_transient_for(parent)
        dialogWindow.set_title(title)
        dialogWindow.set_default_response(Gtk.ResponseType.OK)

        dialogBox = dialogWindow.get_content_area()
        userEntry = Gtk.Entry()
        userEntry.set_activates_default(True)
        userEntry.set_size_request(100,0)
        dialogBox.pack_end(userEntry, False, False, 0)
        
        if value != None:
            userEntry.set_text(value)

        dialogWindow.show_all()
        response = dialogWindow.run()
        text = userEntry.get_text() 
        dialogWindow.destroy()
        if (response == Gtk.ResponseType.OK) and (text != ''):
            return text
        else:
            return None
            
    # Callback functions
    
    def onKeyPressTreeview(self, treeview, event):
        if event.keyval == Gdk.KEY_Escape:  # unselect all
            self.tree.get_selection().unselect_all()

    def undo(self):
        setstack(self.stack) # select schedule undo stack
        print self.stack.undotext()
        self.stack.undo()

    def redo(self):
        setstack(self.stack) # select schedule undo stack
        print self.stack.redotext()
        self.stack.redo()

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
        self.stack.clear()
        self.cmbs = []
        self.update_store()

    def add_cmb(self):
        toplevel = self.tree.get_toplevel() # get current top level window
        cmb_name = self.get_user_input_text(toplevel, "Please input CMB Name", "Add new CMB")
        
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
        meas_name = self.get_user_input_text(toplevel, "Please input Measurement Date", "Add new Measurement")
        
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
        compl_name = self.get_user_input_text(toplevel, "Please input Completion Date", "Add Completion")

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
        heading_name = self.get_user_input_text(toplevel, "Please input Heading", "Add new Item: Heading")
        
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
            print("No items selected to copy")

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
                print('No valid data in clipboard')
        else:
            print("No text on the clipboard.")

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
        setstack(self.stack) # select schedule undo stack

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
            return [CMB_WARNING,'Please select a CMB for rendering']

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
            return [CMB_ERROR,'Rendering of CMB No.' + self.cmbs[path[0]].get_name() + ' failed']

        # run on all bills refering cmb again to rebuild indexes on recursive run
        if recursive: # if recursive call
            for bill_count,bill in enumerate(bills):
                if path[0] in bill.cmb_ref:
                    code = bill.bill_view.render(folder,replacement_dict,[bill_count],False)
                    if code[0] == CMB_ERROR:
                        return code

        return [CMB_INFO,'CMB No.' + self.cmbs[path[0]].get_name() + ' rendered successfully']
            
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
                newval = self.get_user_input_text(toplevel, "Please input CMB Name", "Edit CMB",oldval)
                self.edit_item(path,item,newval,oldval)
            elif len(path) == 2:
                item = self.cmbs[path[0]][path[1]]
                if isinstance(item,Measurement):
                    oldval = item.get_date()
                    newval = self.get_user_input_text(toplevel, "Please input Measurement Date", "Edit Measurement",oldval)
                    self.edit_item(path,item,newval,oldval)
                elif isinstance(item,Completion):
                    oldval = item.get_date()
                    newval = self.get_user_input_text(toplevel, "Please input Completion Date", "Edit Measurement",oldval)
                    self.edit_item(path,item,newval,oldval)
            elif len(path) == 3:
                item = self.cmbs[path[0]][path[1]][path[2]]
                if isinstance(item,MeasurementItemHeading):
                    oldval = item.get_remark()
                    newval = self.get_user_input_text(toplevel, "Please input Heading", "Edit Heading",oldval)
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
        return [CMB_WARNING,'User data not supported']

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
        print oldval,newval
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
        
        self.stack = Stack() # initialise undo/redo stack
        setstack(self.stack) # select schedule undo stack
            
        self.cmbs = [] # initialise item list
        
        self.clipboard = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD) # initialise clipboard
        
        # connect callbacks
        self.tree.connect("key-press-event", self.onKeyPressTreeview)
        
        self.update_store()
