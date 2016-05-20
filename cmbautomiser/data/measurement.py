#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#  measurement.py
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
import copy, logging

# local files import
from __main__ import misc

# Setup logger object
log = logging.getLogger(__name__)

class Cmb:
    """Stores a CMB data instance"""
    def __init__(self, model=None):
        if model is not None:
            self.name = model[0]
            self.items = []
            for item_model in model[1]:
                if item_model[0] in ['Measurement','Completion']:
                    item_type = globals()[item_model[0]]
                    item = item_type()
                    item.set_model(item_model)
                    self.items.append(item)
        else:
            self.name = ''
            self.items = []

    def append_item(self,item):
        self.items.append(item)
                
    def insert_item(self,index,item):
        self.items.insert(index,item)
        
    def remove_item(self,index):
        del(self.items[index])

    def __setitem__(self, index, value):
        self.items[index] = value
    
    def __getitem__(self, index):
        return self.items[index]
        
    def set_name(self,name):
        self.name = name
        
    def get_name(self):
        return self.name
        
    def length(self):
        return len(self.items)
        
    def get_model(self):
        items_model = []
        for item in self.items:
            items_model.append(item.get_model())
        return ['CMB', [self.name, items_model]]
    
    def set_model(self, model):
        if model[0] == 'CMB':
            self.__init__(model[1])

    def clear(self):
        self.items = []
        
    def get_latex_buffer(self, path, schedule):
        latex_buffer = misc.LatexFile()
        cmb_local_vars = {}
        cmb_local_vars['$cmbbookno$'] = self.name
        cmb_local_vars['$cmbheading$'] = ' '
        cmb_local_vars['$cmbtitle$'] = 'DETAILS OF MEASUREMENTS'
        cmb_local_vars['$cmbstartingpage$'] = str(1)
        
        latex_buffer.add_preffix_from_file(misc.abs_path('latex','preamble.tex'))
        latex_buffer.replace_and_clean(cmb_local_vars)
        for count,item in enumerate(self.items):
            newpath = list(path) + [count]
            latex_buffer += item.get_latex_buffer(newpath, schedule)
        latex_buffer.add_suffix_from_file(misc.abs_path('latex','end.tex'))
        return latex_buffer
    
    def get_spreadsheet_buffer(self, path, schedule):
        spreadsheet = misc.Spreadsheet()
        # Set datas
        spreadsheet.add_merged_cell(value='DETAILS OF MEASUREMENT FOR ' + self.name + '  (ref:' + str(path) + ')', bold=True, width=6, horizontal='center')
        spreadsheet.append_data([[None]])
        # Set datas of children
        for slno, item in enumerate(self.items):
            spreadsheet.append(item.get_spreadsheet_buffer(path + [slno], schedule))
        # Set sheet properties
        spreadsheet.set_title('CMB')
        spreadsheet.set_column_widths([10, 50] + [10]*15)
        
        return spreadsheet
        
    def get_text(self):
        return "<b>CMB No." + misc.clean_markup(self.name) + "</b>"
    
    def get_tooltip(self):
        return None
    
    def print_item(self):
        print("CMB " + self.name)
        for item in self.items:
            item.print_item()
        
class Measurement:
    """Stores a Measurement groups"""
    def __init__(self, model = None):
        if model is not None:
            self.date = model[0]
            self.items = []
            class_list = ['MeasurementItemHeading',
                        'MeasurementItemCustom',
                        'MeasurementItemAbstract']
            for item_model in model[1]:
                if item_model[0] in class_list:
                    item_type = globals()[item_model[0]]
                    item = item_type()
                    item.set_model(item_model)
                    self.items.append(item)
        else:
            self.date = ''
            self.items = []

    def append_item(self,item):
        self.items.append(item)
                
    def insert_item(self,index,item):
        self.items.insert(index,item)
        
    def remove_item(self,index):
        del(self.items[index])

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
    
    def get_model(self):
        items_model = []
        for item in self.items:
            items_model.append(item.get_model())
        return ['Measurement', [self.date, items_model]]
    
    def set_model(self, model):
        if model[0] == 'Measurement':
            self.__init__(model[1])
        
    def get_latex_buffer(self, path, schedule):
        latex_buffer = misc.LatexFile()
        latex_buffer.add_preffix_from_file(misc.abs_path('latex', 'measgroup.tex'))
        # Replace local variables
        measgroup_local_vars = {}
        measgroup_local_vars['$cmbmeasurementdate$'] = self.date
        latex_buffer.replace_and_clean(measgroup_local_vars)
        
        for count,item in enumerate(self.items):
            newpath = list(path) + [count]
            latex_buffer += item.get_latex_buffer(newpath, schedule)
        return latex_buffer
    
    def get_spreadsheet_buffer(self, path, schedule):
        spreadsheet = misc.Spreadsheet()
        # Set datas
        rows = [[str(path), 'Date of measurement:', self.date], [None]]
        spreadsheet.append_data(rows, bold=True, wrap_text=False)
        # Set datas of children
        for slno, item in enumerate(self.items):
            spreadsheet.append(item.get_spreadsheet_buffer(path + [slno], schedule))
            
        return spreadsheet
        
    def clear(self):
        self.items = []
    
    def get_text(self):
        return "<b>Measurement dated." + misc.clean_markup(self.date) + "</b>"
    
    def get_tooltip(self):
        return None
    
    def print_item(self):
        print("  " + "Measurement dated " + self.date)
        for item in self.items:
            item.print_item()
        
class MeasurementItem:
    """Base class for storing Measurement items"""
    def __init__(self, itemnos=[], records=[], remark="", item_remarks=[]):
        self.itemnos = itemnos
        self.records = records
        self.remark = remark
        self.item_remarks = item_remarks

    def set_item(self,index,itemno):
        self.itemnos[index] = itemno
        
    def get_item(self,index):
        return self.itemnos[index]
        
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
        self.itemnos = []
        self.records = []
        self.remark = ''
        self.item_remarks = []
                
class MeasurementItemHeading(MeasurementItem):
    """Stores an item heading"""
    def __init__(self, model=None):
        if model is not None:
            MeasurementItem.__init__(self,remark=model[0])
        else:
            MeasurementItem.__init__(self)
    
    def get_model(self):
        model = ['MeasurementItemHeading', [self.remark]]
        return model
    
    def set_model(self, model):
        if model[0] == 'MeasurementItemHeading':
            self.__init__(model[1])
        
    def get_latex_buffer(self, path, schedule):
        latex_buffer = misc.LatexFile()
        latex_buffer.add_preffix_from_file(misc.abs_path('latex', 'measheading.tex'))
        # replace local variables
        measheading_local_vars = {}
        measheading_local_vars['$cmbmeasurementheading$'] = self.remark
        latex_buffer.replace_and_clean(measheading_local_vars)
        return latex_buffer
    
    def get_spreadsheet_buffer(self, path, schedule):
        spreadsheet = misc.Spreadsheet()
        spreadsheet.append_data([[str(path), self.remark], [None]], bold=True, wrap_text=False)
        return spreadsheet
        
    def get_text(self):
        return "<b><i>" + misc.clean_markup(self.remark) + "</i></b>"
    
    def get_tooltip(self):
        return None
        
    def print_item(self):
        print("    " + self.remark)

class RecordCustom:
    """An individual record of a MeasurementItemCustom"""
    def __init__(self, items, cust_funcs, total_func, columntypes):
        self.data_string = items
        self.data = []
        # Populate Data
        for x,columntype in zip(self.data_string,columntypes):
            if columntype not in [misc.MEAS_DESC, misc.MEAS_CUST]:
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

    def get_model(self):
        return self.data_string
        
    def get_model_rendered(self, row=None):
        item = self.get_model()
        rendered_item = []
        for item_elem, columntype, render_func in zip(item, self.columntypes, self.cust_funcs):
            try:
                if item_elem != "" or columntype == misc.MEAS_CUST:
                    if columntype == misc.MEAS_CUST:
                        try:
                            # Try for numerical values
                            value = float(render_func(item, row))
                        except:
                            # If evaluation fails gracefully fallback to string
                            value = render_func(item, row)
                        rendered_item.append(value)
                    if columntype == misc.MEAS_DESC:
                        rendered_item.append(item_elem)
                    elif columntype == misc.MEAS_NO:
                        value = int(eval(item_elem)) if item_elem not in ['0','0.0'] else 0
                        rendered_item.append(value)
                    elif columntype == misc.MEAS_L:
                        value = eval(item_elem) if item_elem not in ['0','0.0'] else 0
                        rendered_item.append(value)
                else:
                    rendered_item.append(None)
            except TypeError:
                rendered_item.append(None)
                log.warning('RecordCustom - Wrong value loaded in item - ' + str(item_elem))
        return rendered_item
        
    def set_model(self, items, cust_funcs, total_func, columntypes):
        self.__init__(items, cust_funcs, total_func, columntypes)

    def find_total(self):
        return self.total_func(self.data)

    def find_custom(self,index):
        return self.cust_funcs[index](self.data)

    def print_item(self):
        print("      " + str([self.data_string,self.total]))


class MeasurementItemCustom(MeasurementItem):
    """Stores a custom record set [As per plugin loaded]"""
    def __init__(self, data = None, plugin=None):
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
        self.dimensions = None

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
                self.dimensions = self.custom_object.dimensions
            except ImportError:
                log.error('Error Loading plugin - MeasurementItemCustom - ' + str(plugin))

            if data != None:
                itemnos = data[0]
                records = []
                for item_model in data[1]:
                    item = RecordCustom(item_model, self.cust_funcs,
                                        self.total_func_item, self.columntypes)
                    records.append(item)
                remark = data[2]
                item_remarks = data[3]
                self.user_data = data[4]
                MeasurementItem.__init__(self, itemnos, records, remark, item_remarks)
            else:
                MeasurementItem.__init__(self, [None]*self.item_width(), [],
                                        '', ['']*self.item_width())
        else:
            MeasurementItem.__init__(self)

    def model_width(self):
        return len(self.columntypes)

    def item_width(self):
        return len(self.itemnos_mask)

    def get_model(self):
        item_schedule = []
        for item in self.records:
            item_schedule.append(item.get_model())
        data = [self.itemnos, item_schedule, self.remark, self.item_remarks,
                self.user_data, self.itemtype]
        return ['MeasurementItemCustom', data]

    def set_model(self, model):
        if model[0] == 'MeasurementItemCustom':
            self.clear()
            self.__init__(model[1], model[1][5])

    def get_latex_buffer(self, path, schedule, isabstract=False):
        latex_records = misc.LatexFile()
        
        data_string = [None]*self.model_width()
        for slno,record in enumerate(self.records):
            meascustom_rec_vars = {}
            meascustom_rec_vars_van = {}
            # Evaluate string to make replacement
            for i,columntype in enumerate(self.columntypes): # evaluate string of data entries, suppress zero.
                if columntype == misc.MEAS_CUST:
                    try:
                        value =  str(record.cust_funcs[i](record.get_model(),slno))
                        data_string[i] = value if value not in ['0','0.0'] else ''
                    except:
                        data_string[i] = ''
                elif columntype == misc.MEAS_DESC:
                    try:
                        data_string[i] = str(record.data_string[i])
                    except:
                        data_string[i] = ''
                elif columntype == misc.MEAS_NO:
                    try:
                        data_string[i] = str(int(record.data[i])) if record.data[i] != 0 else ''
                    except:
                        data_string[i] = ''
                else:
                    try:
                        data_string[i] = str(record.data[i]) if record.data[i] != 0 else ''
                    except:
                        data_string[i] = ''
                # Check for carry over item possibly contains code
                if columntype == misc.MEAS_DESC and data_string[i].find('Qty B/F') != -1:
                    saved_path = data_string[i][9:]
                    cmbbf = 'ref:meas:'+ saved_path + ':1'
                    label = 'ref:abs:'+ saved_path + ':1'
                    record_code = r'Qty B/F MB.No.\emph{\nameref{' + cmbbf + r'} Pg.No. \pageref{' + cmbbf \
                            + r'}}\phantomsection\label{' + label + '}'
                    meascustom_rec_vars_van['$data' + str(i+1) + '$'] = record_code
                else:
                    meascustom_rec_vars['$data' + str(i+1) + '$'] = data_string[i]
            meascustom_rec_vars['$slno$'] = str(slno+1)
            
            latex_record = misc.LatexFile(self.latex_record)
            latex_record.replace(meascustom_rec_vars_van)
            latex_record.replace_and_clean(meascustom_rec_vars)
            latex_records += latex_record
            
        # replace local variables
        meascustom_local_vars = {}
        meascustom_local_vars_vannilla = {}
        for i in range(0,self.item_width()):
            try:
                meascustom_local_vars['$cmbitemdesc' + str(i+1) + '$'] = str(schedule[self.itemnos[i]].extended_description)
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
        # fill in records - vanilla used since latex_records contains latex code
        meascustom_local_vars_vannilla['$cmbrecords$'] = latex_records.get_buffer()
            
        latex_buffer = misc.LatexFile(self.latex_item)
        latex_buffer.replace_and_clean(meascustom_local_vars)
        latex_buffer.replace(meascustom_local_vars_vannilla)
        
        latex_post = self.latex_postproc_func(self.records, self.user_data, latex_buffer, isabstract)
        return latex_post
        
    def get_spreadsheet_buffer(self, path, schedule):
        spreadsheet = misc.Spreadsheet()
        # Item no and description
        for itemno in self.itemnos:
            spreadsheet.append_data([[str(path), 'Item No:' + itemno]], bold=True)
            spreadsheet.append_data([[None, schedule[itemno].extended_description]])
        # Data rows
        spreadsheet.append_data([[None], [None] + self.captions], bold=True)
        for slno, record in enumerate(self.records,1):
            values = record.get_model_rendered(slno)
            spreadsheet.append_data([[slno] + values])
        # User data
        if self.captions_udata:
            spreadsheet.append_data([[None], [None, 'User Data Captions'] + self.captions_udata], bold=True)
            spreadsheet.append_data([[None, 'User Datas'] + self.user_data])
        # Total values
        spreadsheet.append_data([[None], [None, 'TOTAL'] + self.get_total(), [None]], bold=True)
        return spreadsheet

    def print_item(self):
        print("    Item No." + str(self.itemnos))
        for i in range(self.length()):
            self[i].print_item()
        print("    " + "Total: " + str(self.get_total()))

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


class MeasurementItemAbstract(MeasurementItem):
    """Stores an abstract of measurements"""
    def __init__(self, data = None):
        self.int_mitem = None  # MeasurementItemCustom for storing abstract
        self.mitems = []  # Paths to items to be abstracted

        if data is not None:
            self.mitems = data[0]
            self.int_mitem = MeasurementItemCustom(data[1][1],data[1][1][5])
            MeasurementItem.__init__(self, itemnos=self.int_mitem.itemnos, 
                records=self.int_mitem.records, remark=self.int_mitem.remark, 
                item_remarks = self.int_mitem.item_remarks)
        else:
            MeasurementItem.__init__(self, itemnos=[], records=[], 
                remark='', item_remarks = [])

    def get_model(self):
        model = None
        if self.int_mitem is not None:
            model = self.int_mitem.get_model()
        data = [self.mitems, model]
        return ['MeasurementItemAbstract', data]

    def set_model(self, model):
        if model[0] == 'MeasurementItemAbstract':
            self.clear()
            self.__init__(model[1])
            
    def get_abstracted_items(self):
        return self.mitems

    def get_latex_buffer(self, path, schedule):
        if self.mitems is not None:
            return self.int_mitem.get_latex_buffer(path, schedule, True)
            
    def get_spreadsheet_buffer(self, path, schedule):
        if self.mitems is not None:
            return self.int_mitem.get_spreadsheet_buffer(path, schedule)

    def print_item(self):
        print('    Abstract Item')
        self.int_mitem.print_item()

    def get_total(self):
        if self.int_mitem is not None:
            return self.int_mitem.get_total()
        else:
            return []

    def get_text(self):
        if self.int_mitem is not None:
            return 'Abs: ' + self.int_mitem.get_text()
        else:
            return 'Abs: NOT DEFINED'

    def get_tooltip(self):
        if self.int_mitem is not None:
            if self.int_mitem.get_tooltip() is not None:
                return 'Abs: ' + self.int_mitem.get_tooltip()

class Completion:
    """Class storing Completion date"""
    def __init__(self, model=None):
        if model is not None:
            self.date = model[0]
        else:
            self.date = ''
        MeasurementItem.__init__(self)

    def set_date(self,date):
        self.date = date

    def get_date(self):
        return self.date
        
    def get_model(self):
        return ['Completion',[self.date]]
    
    def set_model(self, model):
        if model[0] == 'Completion':
            self.__init__(model[1])
    
    def get_latex_buffer(self, path, schedule):
        latex_buffer = misc.LatexFile()
        latex_buffer.add_preffix_from_file(misc.abs_path('latex', 'meascompletion.tex'))
        # replace local variables
        measgroup_local_vars = {}
        measgroup_local_vars['$cmbcompletiondate$'] = self.date
        latex_buffer.replace_and_clean(measgroup_local_vars)
        return latex_buffer
        
    def get_spreadsheet_buffer(self, path, schedule):
        spreadsheet = misc.Spreadsheet()
        spreadsheet.append_data([[None], [str(path), 'DATE OF COMPLETION', self.date], [None]], bold=True, wrap_text=False)
        return spreadsheet

    def get_text(self):
        return "<b>Completion recorded on " + misc.clean_markup(self.date) + "</b>"

    def get_tooltip(self):
        return None

    def print_item(self):
        print("  " + "Completion recorded on " + self.date)\
        
