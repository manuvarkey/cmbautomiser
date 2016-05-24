#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# tableofpoints.py
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

# Item codes for schedule dialog * DONT CHANGE *
MEAS_NO = 1
MEAS_L = 2
MEAS_DESC = 3
MEAS_CUST = 4

# Define your latex strings here
latex_record = r"""        $slno$ & $data1$ & \AddBreakableChars{$data2$} & $data3$ & $data4$ & $data5$ & $data6$ & $data7$ \\
                """
latex_item = r"""
                % Item LLLLL
                \fakesubsubsection{$cmbitemno1$,$cmbitemno2$,$cmbitemno3$,$cmbitemno4$,$cmbitemno5$}
                $cmbitemexist1$
                \noindent\textbf{Item No.$cmbitemno1$}\\
                \noindent $cmbitemdesc1$\\
                \fi
                $cmbitemexist2$
                \noindent\textbf{Item No.$cmbitemno2$}\\
                \noindent $cmbitemdesc2$\\
                \fi
                $cmbitemexist3$
                \noindent\textbf{Item No.$cmbitemno3$}\\
                \noindent $cmbitemdesc3$\\
                \fi
                $cmbitemexist4$
                \noindent\textbf{Item No.$cmbitemno4$}\\
                \noindent $cmbitemdesc4$\\
                \fi
                $cmbitemexist5$
                \noindent\textbf{Item No.$cmbitemno5$}\\
                \noindent $cmbitemdesc5$\\
                \fi
                \noindent\emph{(Remarks:$cmbremark$)}\\
                \vspace*{-\baselineskip}
                \begin{longtabu} to \textwidth {|X[1,c]|X[10,l]|X[5,l]|X[3,r]|X[3,r]|X[3,r]|X[3,r]|X[3,r]|}
                   \hline
                    \textbf{Sl.\newline No.} & \textbf{Description} & \textbf{Breakup} & $cmbitemexist1$\textbf{Item No. $cmbitemno1$}\fi & $cmbitemexist2$\textbf{Item No. $cmbitemno2$}\fi & $cmbitemexist3$\textbf{Item No. $cmbitemno3$}\fi & $cmbitemexist4$\textbf{Item No. $cmbitemno4$}\fi & $cmbitemexist5$\textbf{Item.No. $cmbitemno5$}\fi \\
                   \hline
                     &  &  & \emph{$cmbitemremark1$} & \emph{$cmbitemremark2$} & \emph{$cmbitemremark3$} & \emph{$cmbitemremark4$} & \emph{$cmbitemremark5$} \\
                   \hline
                    \endhead
                $cmbrecords$
                    \hline
                      & \textbf{TOTAL} & & $cmbitemexist1$ \textbf{$cmbtotal1$} \fi & $cmbitemexist2$ \textbf{$cmbtotal2$} \fi & $cmbitemexist3$ \textbf{$cmbtotal3$} \fi & $cmbitemexist4$ \textbf{$cmbtotal4$} \fi & $cmbitemexist5$ \textbf{$cmbtotal5$} \fi \\
                   \hline
                      & \emph{Qty C/o MB.No. \nameref{$cmbcarriedover1$}} &  & $cmbitemexist1$\emph{Pg.No. \pageref{$cmbcarriedover1$}} \phantomsection\label{$cmblabel1$}\fi & $cmbitemexist2$\emph{Pg.No. \pageref{$cmbcarriedover2$}} \phantomsection\label{$cmblabel2$}\fi & $cmbitemexist3$\emph{Pg.No. \pageref{$cmbcarriedover3$}} \phantomsection\label{$cmblabel3$} \fi & $cmbitemexist4$\emph{Pg.No. \pageref{$cmbcarriedover4$}} \phantomsection\label{$cmblabel4$} \fi & $cmbitemexist5$\emph{Pg.No. \pageref{$cmbcarriedover5$}} \phantomsection\label{$cmblabel5$}\fi \\
                   \hline
                \end{longtabu}

             """

class CustomItem:
    def __init__(self):
    
        def callback_breakup(values,row=None):
            data_str = values[2:7]
            breakup = "["
            for x in data_str:
                if x != "" and x!= '0':
                    breakup = breakup + str(x) + ","
                else:
                    breakup = breakup + ','      
            breakup = breakup[:-1] + "]"
            return breakup

        def total_func(item_list,userdata=None):
            total = [0,0,0,0,0]
            for item in item_list:
                if item is not None:
                    itemtot = item.find_total()
                    for i in range(5):
                        total[i] += itemtot[i]
            for i in range(5):
                total[i] = round(total[i],3)
            return total
        
        def total_func_item(values):
            return values[2:7]
                
        def latex_postproc_func(item_list,userdata,latex_buffer,isabstract=False):
            # Do nothing return as it is
            return latex_buffer
        
        # Define your variables here
        self.name = 'Item LLLLL'
        self.itemnos_mask = [None,None,None,None,None]
        self.captions = ['Description','Breakup','L1','L2','L3','L4','L5']
        self.columntypes = [MEAS_DESC,MEAS_CUST,MEAS_L,MEAS_L,MEAS_L,MEAS_L,MEAS_L]
        self.latex_item = latex_item
        self.latex_record = latex_record
        self.captions_udata = []
        self.columntypes_udata = []
        self.user_data_default = []
        # Define functions here
        self.cust_funcs = [None, callback_breakup, None, None, None, None, None]
        self.total_func = total_func
        self.total_func_item = total_func_item
        self.latex_postproc_func = latex_postproc_func
        self.export_abstract = None
        self.dimensions = [[200,150,80,80,80,80,80], [True,False,False,False,False,False,False]]
          
        
        
