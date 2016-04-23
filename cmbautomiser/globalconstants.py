#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# globalconstants.py
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

# Program name
PROGRAM_NAME = 'CMB Automiser'
# Item codes for schedule dialog
MEAS_NO = 1
MEAS_L = 2
MEAS_DESC = 3
MEAS_CUST = 4
# CMB error codes used for displaying info in main window
CMB_ERROR = -1
CMB_WARNING = -2
CMB_OK = 0
CMB_INFO = 0
# For tracking window management
CHILD_WINDOW = 1
PARENT_WINDOW = 0
main_hidden = 0
child_windows = 0
# Used in bill module for indicating type of bill
BILL_CUSTOM = 1
BILL_NORMAL = 2
# background colors for treeview
MEAS_COLOR_LOCKED = '#BABDB6'
MEAS_COLOR_NORMAL = '#FFFFFF'
MEAS_COLOR_SELECTED = '#729FCF'
# Timeout for killing Latex subprocess
LATEX_TIMEOUT = 300 # 5 minutes
# Item description wrap-width for screen purpose
CMB_DESCRIPTION_WIDTH = 60
CMB_DESCRIPTION_MAX_LENGTH = 1000
# Deviation statement
DEV_LIMIT_STATEMENT = 10
        
# List of units which will be considered as integer values
INT_ITEMS = ['point', 'points', 'pnt', 'pnts', 'number', 'numbers', 'no', 'nos', 'lot', 'lots',
             'lump', 'lumpsum', 'lump-sum', 'lump sum', 'ls', 'each','job','jobs','set','sets',
             'pair','pairs',
             'pnt.', 'no.', 'nos.', 'l.s.', 'l.s']
# String used for checking file version
PROJECT_FILE_VER = 'CMBAUTOMISER_FILE_REFERENCE_VER_3'

# Item codes for cmb
global_vars = ['$cmbnameofwork$', '$cmbagency$', '$cmbagmntno$', '$cmbsituation$', '$cmbdateofstart$',
               '$cmbdateofstartasperagmnt$',
               '$cmbissuedto$', '$cmbvarifyingauthority$', '$cmbvarifyingauthorityoffice$', '$cmbissuingauthority$',
               '$cmbissuingauthorityoffice$']
# Widget names equivalent to item codes
builder_vars = ["proj_nameofwork", "proj_agency", "proj_agmntno", "proj_situation", "proj_dateofstart",
                "proj_dateofstartasperagmnt",
                "proj_issuedto", "proj_varifyingauthority", "proj_varifyingauthorityoffice", "proj_issuingauthority",
                "proj_issuingauthorityoffice"]

