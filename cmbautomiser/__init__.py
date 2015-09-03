#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# main.py
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

import os
import dill as pickle
import urlparse
import json
import sys

from gi.repository import Gtk, Gdk, GLib

from undo import *
import appdirs
from openpyxl import Workbook, load_workbook

# local files import
from globalconstants import *
from schedule import *
from cmb import *
from bill import *
from misc import *
import globalvars

class MainWindow:
    # General Methods

    def display_error(self, error_text):
        infobar_main = self.builder.get_object("infobar_main")
        label_infobar_main = self.builder.get_object("label_infobar_main")
        infobar_main.set_message_type(Gtk.MessageType.ERROR)
        label_infobar_main.set_text(error_text)
        infobar_main.show()

    def display_info(self, error_text):
        infobar_main = self.builder.get_object("infobar_main")
        label_infobar_main = self.builder.get_object("label_infobar_main")
        infobar_main.set_message_type(Gtk.MessageType.INFO)
        label_infobar_main.set_text(error_text)
        infobar_main.show()

    def display_warning(self, error_text):
        infobar_main = self.builder.get_object("infobar_main")
        label_infobar_main = self.builder.get_object("label_infobar_main")
        infobar_main.set_message_type(Gtk.MessageType.WARNING)
        label_infobar_main.set_text(error_text)
        infobar_main.show()

    def display_status(self, status):
        status_code = status[0]
        message = status[1]
        if status_code == CMB_ERROR:
            self.display_error(message)
        elif status_code == CMB_WARNING:
            self.display_warning(message)
        elif status_code == CMB_INFO:
            self.display_info(message)


    def get_user_input(self, parent, message, title=''):
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
        userEntry.set_size_request(50, 0)
        dialogBox.pack_end(userEntry, False, False, 0)

        dialogWindow.show_all()
        response = dialogWindow.run()
        text = userEntry.get_text()
        dialogWindow.destroy()
        if (response == Gtk.ResponseType.OK) and (text != ''):
            return text
        else:
            return None

    # General signal handler Methods

    def onProjectSettingsClicked(self, button):
        self.project_settings_dialog.show()

        # write values to dialog
        for item_code, builder_var in zip(global_vars, builder_vars):
            self.builder.get_object(builder_var).set_text(self.project_settings_dict[item_code])

        # run dialog
        response = self.project_settings_dialog.run()

        if response == Gtk.ResponseType.APPLY:
            # read values to dictionary
            for item_code, builder_var in zip(global_vars, builder_vars):
                self.project_settings_dict[item_code] = self.builder.get_object(builder_var).get_text()
            self.project_settings_dialog.hide()
        elif response == Gtk.ResponseType.CANCEL:
            self.project_settings_dialog.hide()


    def onProjectSettingsClose(self, *args):
        self.project_settings_dialog.hide()
        return True

    # About Dialog

    def onAboutDialogClose(self, *args):
        self.about_dialog.hide()
        return True

    def onAboutClick(self, button):
        self.about_dialog.show()
        response = self.about_dialog.run()
        if response == Gtk.ResponseType.CANCEL :
            self.about_dialog.hide()

    # Main Window

    def onDeleteWindow(self, *args):
        global CHILD_WINDOW, PARENT_WINDOW, main_hidden, child_windows
        if self.windowtype == CHILD_WINDOW:
            child_windows = child_windows - 1
            if child_windows == 0 and main_hidden == 1:
                Gtk.main_quit()
            self.window.destroy()
        else:
            if child_windows == 0:
                Gtk.main_quit()
            else:
                self.window.hide()
                main_hidden = 1

    def onNewProjectClicked(self, button):
        global CHILD_WINDOW, PARENT_WINDOW, main_hidden, child_windows
        child_window = MainWindow(CHILD_WINDOW)
        child_windows = child_windows + 1
        child_window.run()
        child_window.display_info("New project window created")

    def onOpenProjectClicked(self, button):
        # create a filechooserdialog to open:
        # the arguments are: title of the window, parent_window, action,
        # (buttons, response)
        open_dialog = Gtk.FileChooserDialog("Open project File", self.window,
                                            Gtk.FileChooserAction.OPEN,
                                            (Gtk.STOCK_CANCEL, Gtk.ResponseType.CANCEL,
                                             Gtk.STOCK_OPEN, Gtk.ResponseType.ACCEPT))
        # not only local files can be selected in the file selector
        open_dialog.set_local_only(False)
        # dialog always on top of the textview window
        open_dialog.set_modal(True)
        # set filters
        open_dialog.set_filter(self.builder.get_object("filefilter_project"))
        # set window position
        open_dialog.set_gravity(Gdk.Gravity.CENTER)
        open_dialog.set_position(Gtk.WindowPosition.CENTER_ON_PARENT)


        response_id = open_dialog.run()
        # if response is "ACCEPT" (the button "Save" has been clicked)
        if response_id == Gtk.ResponseType.ACCEPT:
            # get filename and set project as active
            self.filename = open_dialog.get_filename()
            fileobj = open(self.filename, 'rb')
            if fileobj == None:
                print("Error opening file " + self.filename)
                self.display_error("Project could not be opened: Error opening file")
            else:
                data = pickle.load(fileobj)  # load data structure
                fileobj.close()
                if data[0] == PROJECT_FILE_VER:
                    # clear application states
                    self.schedule_view.clear()
                    self.measurements_view.clear()
                    self.bill_view.clear()
                    # load application data

                    self.schedule = data[1]
                    self.schedule_view.set_data_object(self.schedule)
                    self.measurements_view.set_data_object(self.schedule, data[2])
                    self.bill_view.set_data_object(self.schedule, self.measurements_view, data[3])
                    self.project_settings_dict = data[4]

                    # set project as active
                    self.PROJECT_ACTIVE = 1

                    self.display_info("Project successfully opened")
                    # Setup paths for folder chooser objects
                    self.builder.get_object("filechooserbutton_meas").set_current_folder(posix_path(
                        os.path.split(self.filename)[0]))
                    self.builder.get_object("filechooserbutton_bill").set_current_folder(posix_path(
                        os.path.split(self.filename)[0]))
                    # setup window name
                    self.window.set_title(self.filename + ' - ' + PROGRAM_NAME)

                else:
                    self.display_error("Project could not be opened: Wrong file type selected")
        # if response is "CANCEL" (the button "Cancel" has been clicked)
        elif response_id == Gtk.ResponseType.CANCEL:
            print("cancelled: FileChooserAction.OPEN")
        # destroy dialog
        open_dialog.destroy()

    def onSaveProjectClicked(self, button):
        if self.PROJECT_ACTIVE == 0:
            self.onSaveAsProjectClicked(button)
        else:
            # parse data into object
            data = []
            dataschedule = self.schedule_view.get_data_object()
            datameasurement = self.measurements_view.get_data_object()
            databill = self.bill_view.get_data_object()

            data.append(PROJECT_FILE_VER)
            data.append(dataschedule)
            data.append(datameasurement)
            data.append(databill)
            data.append(self.project_settings_dict)

            # try to open file
            fileobj = open(self.filename, 'wb')
            if fileobj == None:
                print "Error opening file " + self.filename
                self.display_error("Project file could not be opened for saving")
            pickle.dump(data, fileobj)
            fileobj.close()
            self.display_info("Project successfully saved")

    def onSaveAsProjectClicked(self, button):
        # create a filechooserdialog to open:
        # the arguments are: title of the window, parent_window, action,
        # (buttons, response)
        open_dialog = Gtk.FileChooserDialog("Save project as...", self.window,
                                            Gtk.FileChooserAction.SAVE,
                                            (Gtk.STOCK_CANCEL, Gtk.ResponseType.CANCEL,
                                             Gtk.STOCK_SAVE, Gtk.ResponseType.ACCEPT))
        # not only local files can be selected in the file selector
        open_dialog.set_local_only(False)
        # dialog always on top of the textview window
        open_dialog.set_modal(True)
        # set filters
        open_dialog.set_filter(self.builder.get_object("filefilter_project"))
        # set window position
        open_dialog.set_gravity(Gdk.Gravity.CENTER)
        open_dialog.set_position(Gtk.WindowPosition.CENTER_ON_PARENT)
        # set overwrite confirmation
        open_dialog.set_do_overwrite_confirmation(True)
        # set default name
        open_dialog.set_current_name("newproject.proj")
        response_id = open_dialog.run()
        # if response is "ACCEPT" (the button "Save" has been clicked)
        if response_id == Gtk.ResponseType.ACCEPT:
            # get filename and set project as active
            self.filename = open_dialog.get_filename()
            self.PROJECT_ACTIVE = 1
            # call save project
            self.onSaveProjectClicked(button)
            # Setup paths for folder chooser objects
            self.builder.get_object("filechooserbutton_meas").set_current_folder(posix_path(
                os.path.split(self.filename)[0]))
            self.builder.get_object("filechooserbutton_bill").set_current_folder(posix_path(
                os.path.split(self.filename)[0]))
        # if response is "CANCEL" (the button "Cancel" has been clicked)
        elif response_id == Gtk.ResponseType.CANCEL:
            print "cancelled: FileChooserAction.OPEN"
        # destroy dialog
        open_dialog.destroy()

    def onInfobarClose(self, widget, response=0):
        widget.hide()

    def onRedo(self, button):
        print stack().redotext()
        stack().redo()

    def onUndo(self, button):
        print stack().undotext()
        stack().undo()

    # Schedule signal handler methods

    def onButtonScheduleAddPressed(self, button):
        items = []
        items.append(ScheduleItem("", "", "", 0, 0, "", 30))
        self.schedule_view.insert_item_at_selection(items)

    def onButtonScheduleAddMultPressed(self, button):
        userInput = self.get_user_input(self.window, "Please enter the number \nof rows to be inserted",
                                        "Number of Rows")
        try:
            number_of_rows = int(userInput)
        except:
            self.display_warning("Invalid number of rows specified")
            return
        items = []
        for i in range(0, number_of_rows):
            items.append(ScheduleItem("", "", "", 0, 0, "", 30))
        self.schedule_view.insert_item_at_selection(items)

    def onButtonScheduleDeletePressed(self, button):
        self.schedule_view.delete_selection()

    def onUndoSchedule(self, button):
        self.schedule_view.undo()

    def onRedoSchedule(self, button):
        self.schedule_view.redo()

    def onCopySchedule(self, button):
        self.schedule_view.copy_selection()

    def onPasteSchedule(self, button):
        self.schedule_view.paste_at_selection()

    def onImportScheduleClicked(self, button):
        filename = self.builder.get_object("filechooserbutton_schedule").get_filename()
        spreadsheet = load_workbook(filename)
        sheet = spreadsheet.active
        # get count of rows
        rowcount = len(sheet.columns)
        items = []
        for row in range(0, rowcount):
            item = ScheduleItem()
            for col in range(0, 7):
                cell = sheet.cell(row = row+1, column = col+1).value
                # Get formated string input
                if cell is None:
                    cell_formated = ""
                else:
                    try:  # try evaluating string
                        cell_formated = str(cell)
                    except:
                        cell_formated = ""
                # Try adding to item
                if col in [0,1,2,5]:
                    item[col] = cell_formated
                elif col in [3,4,6]:
                    try:
                        item[col] = float(cell_formated)
                    except:
                        item[col] = ScheduleItem()[col]
            # Hack for proper formating of AgmntNos
            try:
                agmntno_float = float(item[0])
                agmntno_int = int(agmntno_float)
                if float(agmntno_int) == agmntno_float:
                    item[0] = str(agmntno_int)
            except ValueError:
                pass
            items.append(item)
        self.schedule_view.insert_item_at_selection(items)



    # Measuremets signal handlers

    def onAddCmbClicked(self, button):
        self.measurements_view.add_cmb()

    def onAddMeasClicked(self, button):
        self.measurements_view.add_measurement()

    def onAddComplClicked(self, button):
        self.measurements_view.add_completion()

    def onAddHeadingClicked(self, button):
        self.measurements_view.add_heading()

    def onAddNLBHClicked(self, button):
        self.measurements_view.add_nlbh(None)

    def onAddLLLLLClicked(self, button):
        self.measurements_view.add_lllll(None)

    def onAddNNNNNNNNClicked(self, button):
        self.measurements_view.add_nnnnnnnn(None)

    def onAddnnnnnTClicked(self, button):
        self.measurements_view.add_nnnnnt(None)

    def onAddAbstractClicked(self, button):
        self.measurements_view.add_abstract(None)

    def OnMeasDeleteClicked(self, button):
        self.measurements_view.delete_selected_row()

    def OnMeasRenderClicked(self, button):
        filechooserbutton_meas = self.builder.get_object("filechooserbutton_meas")
        folder_uri = filechooserbutton_meas.get_uri()
        if folder_uri <> None:
            folder = urlparse.urlparse(folder_uri).path
            code = self.measurements_view.render_selection(folder, self.project_settings_dict, self.bill_view.bills)
            self.display_status(code)
            # remove temporary files
            onlytempfiles = [f for f in os.listdir(posix_path(folder)) if (f.find('.aux')!=-1 or f.find('.log')!=-1
                                or f.find('.out')!=-1) or f.find('.tex')!=-1 or f.find('.bak')!=-1]
            for f in onlytempfiles:
                os.remove(posix_path(folder,f))

    def OnMeasUndoClicked(self, button):
        self.measurements_view.undo()

    def OnMeasRedoClicked(self, button):
        self.measurements_view.redo()

    def OnMeasEditClicked(self, button):
        self.measurements_view.edit_selected_row()

    def OnMeasPropertiesClicked(self, button):
        code = self.measurements_view.edit_selected_properties()
        if code is not None:
            self.display_status(code)

    def OnMeasCopyClicked(self, button):
        self.measurements_view.copy_selection()

    def OnMeasPasteClicked(self, button):
        self.measurements_view.paste_at_selection()
        
    def onMeasCustomMenuClicked(self,button,module):
        self.measurements_view.add_custom(None,module)

    # Bills signal handlers

    def onAddBillClicked(self, button):
        self.bill_view.add_bill()

    def onAddBillCustomClicked(self, button):
        self.bill_view.add_bill_custom()

    def OnBillDeleteClicked(self, button):
        self.bill_view.delete_selected_row()

    def OnBillRenderClicked(self, button):
        filechooserbutton_bill = self.builder.get_object("filechooserbutton_bill")
        folder_uri = filechooserbutton_bill.get_uri()
        if folder_uri <> None:
            folder = urlparse.urlparse(folder_uri).path
            code = self.bill_view.render_selected(folder, self.project_settings_dict)
            self.display_status(code)
            # remove temporary files
            onlytempfiles = [f for f in os.listdir(posix_path(folder)) if (f.find('.aux')!=-1 or f.find('.log')!=-1
                                or f.find('.out')!=-1) or f.find('.tex')!=-1 or f.find('.bak')!=-1]
            for f in onlytempfiles:
                os.remove(posix_path(folder,f))

    def OnBillUndoClicked(self, button):
        self.bill_view.undo()

    def OnBillRedoClicked(self, button):
        self.bill_view.redo()

    def OnBillEditClicked(self, button):
        self.bill_view.edit_selected_row()

    def OnBillCopyClicked(self, button):
        self.bill_view.copy_selection()

    def OnBillPasteClicked(self, button):
        self.bill_view.paste_at_selection()

    # Tab methods

    def onSwitchTab(self, widget, page, pagenum):
        self.schedule_view.update_store()
        self.measurements_view.update_store()
        self.bill_view.update_store()

    def __init__(self, _windowtype=PARENT_WINDOW):
        
        # Variable to check multiple window instances
        if _windowtype == PARENT_WINDOW:
            self.windowtype = PARENT_WINDOW
        else:
            self.windowtype = CHILD_WINDOW

        # check for project active state
        self.PROJECT_ACTIVE = 0

        # other variables

        self.filename = None

        # Setup main window
        self.builder = Gtk.Builder()
        self.builder.add_from_file(abs_path("interface","mainwindow.glade"))
        self.window = self.builder.get_object("window_main")
        self.builder.connect_signals(self)

        # Setup project settings dialog
        self.project_settings_dict = {}
        for item_code in global_vars:  # initialise settings dictionary
            self.project_settings_dict[item_code] = ''
        self.project_settings_dialog = self.builder.get_object("projectsettingsdialog")
        
        # Load global Variables
        globalvars.set_global_platform_vars()

        # dirs = appdirs.AppDirs("cmb_automiser", "cpwd", version="1.0")
        # settings_filename = os.path.join(dirs.user_data_dir,'settings.conf')
        # try:
            # settings_file = open(settings_filename,'r')
            # globalvars.global_settings_dict = json.loads(settings_file.read())
            # settings_file.close()
        # except (ValueError,IOError):
            # try:
                # os.makedirs(dirs.user_data_dir)
            # except OSError:
                # pass
            # try:
                # settings_file = open(settings_filename,'w')
                # settings_file.write(json.dumps(globalvars.global_settings_dict))
                # settings_file.close()
            # except (IOError,ValueError):
                # pass
        
        # Setup about dialog
        self.about_dialog = self.builder.get_object("aboutdialog")

        # Setup a schedule of items
        self.schedule = Schedule()  # initialise a schedule

        # Setup measurement View
        self.treeview_meas = self.builder.get_object("treeview_meas")
        self.liststore_meas = self.builder.get_object("liststore_meas")
        self.measurements_view = MeasurementsView(self.schedule, self.liststore_meas, self.treeview_meas)
        
        # Setup custom measurement items
        file_names = [f for f in os.listdir(abs_path('templates'))]
        module_names = []
        for f in file_names:
            if f[-3:] == '.py' and f != '__init__.py':
                module_names.append(f[:-3])
        self.custom_menus = []
        popupmenu = self.builder.get_object("popupmenu_meas")
        module_names.sort()
        for module_name in module_names:
            try:
                package = __import__('templates.' + module_name)
                module = getattr(package, module_name)
                custom_object = module.CustomItem()
                name = custom_object.name
                menuitem = Gtk.MenuItem(label=name)
                popupmenu.append(menuitem)
                menuitem.set_visible(True)
                menuitem.connect("activate",self.onMeasCustomMenuClicked,module_name)
                self.custom_menus.append(menuitem)
            except ImportError:
                print('Error Loading plugin')

        # Setup bill View
        self.treeview_bill = self.builder.get_object("treeview_bill")
        self.liststore_bill = self.builder.get_object("liststore_bill")
        self.bill_view = BillView(self.schedule, self.measurements_view, self.liststore_bill, self.treeview_bill)

        # Setup schedule View
        self.treeview_schedule = self.builder.get_object("treeview_schedule")
        self.liststore_schedule = self.builder.get_object("liststore_schedule")
        self.schedule_view = ScheduleView(self.schedule, self.liststore_schedule, self.treeview_schedule)
        # Connect signals with custom userdata
        self.builder.get_object("tree_schedule_agmntno").connect("edited", self.schedule_view.onScheduleCellEdited, 0)
        self.builder.get_object("tree_schedule_description").connect("edited", self.schedule_view.onScheduleCellEdited,
                                                                     1)
        self.builder.get_object("tree_schedule_unit").connect("edited", self.schedule_view.onScheduleCellEdited, 2)
        self.builder.get_object("tree_schedule_rate").connect("edited", self.schedule_view.onScheduleCellEditedRates, 3)
        self.builder.get_object("tree_schedule_qty").connect("edited", self.schedule_view.onScheduleCellEditedRates, 4)
        self.builder.get_object("tree_schedule_reference").connect("edited", self.schedule_view.onScheduleCellEdited, 5)
        self.builder.get_object("tree_schedule_percent").connect("edited", self.schedule_view.onScheduleCellEditedRates,
                                                                 6)
        # setup infobar
        self.builder.get_object("infobar_main").hide()

        # setup class for intra process communication
        self.manage_resources = ManageResourses(self.schedule_view,self.measurements_view,self.bill_view)

    def run(self, *args):
        self.window.show_all()


def main():
    MainWindow().run()  # initialise main window
    Gtk.main()  # loop
    return 0


if __name__ == '__main__':
    main()

