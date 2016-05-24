#!/usr/bin/env python3
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

import subprocess, os, platform, sys, tempfile, logging, json
from gi.repository import Gtk, Gdk, GLib

import undo

# local files import
import misc, data.datamodel, data.schedule, view.measurement, view.bill

# Get logger object
log = logging.getLogger()

class MainWindow:

    # General Methods

    def display_status(self, status_code, message):
        """Displays a formated message in Infobar
            
            Arguments:
                status_code: Specifies the formatting of message.
                             (Takes the values misc.CMB_ERROR,
                              misc.CMB_WARNING, misc.CMB_INFO]
                message: The message to be displayed
        """
        infobar_main = self.builder.get_object("infobar_main")
        label_infobar_main = self.builder.get_object("label_infobar_main")
        
        if status_code == misc.CMB_ERROR:
            infobar_main.set_message_type(Gtk.MessageType.ERROR)
            label_infobar_main.set_text(message)
            infobar_main.show()
        elif status_code == misc.CMB_WARNING:
            infobar_main.set_message_type(Gtk.MessageType.WARNING)
            label_infobar_main.set_text(message)
            infobar_main.show()
        elif status_code == misc.CMB_INFO:
            infobar_main.set_message_type(Gtk.MessageType.INFO)
            label_infobar_main.set_text(message)
            infobar_main.show()
    
    def update(self):
        """Refreshes all displays"""
        log.info('MainWindow update called')
        self.data.update()
        self.schedule_view.update_store()
        self.measurements_view.update_store()
        self.bill_view.update_store()
            
    # About Dialog

    def onAboutDialogClose(self, *args):
        """Hide about dialog"""
        self.about_dialog.hide()
        return True

    def onAboutClick(self, button):
        """Show about dialog"""
        log.info('onAboutClick - Show About window')
        self.about_dialog.show()
        response = self.about_dialog.run()
        if response == Gtk.ResponseType.CANCEL:
            self.about_dialog.hide()        

    def onHelpClick(self, button):
        """Launch help file"""
        log.info('onHelpClick - Launch Help file')
        if platform.system() == 'Linux':
            subprocess.call(('xdg-open', misc.abs_path('documentation', 'cmbautomisermanual.pdf')))
        elif platform.system() == 'Windows':
            os.startfile(abs_path('documentation','cmbautomisermanual.pdf'))

    # Main Window

    def onDeleteWindow(self, *args):
        """Callback called on pressing the close button of main window"""
        
        log.info('onDeleteWindow called')
        
        def wait_for_exit(child_windows):
            """Wait for child windows to exit before calling Gtk.main_quit()"""
            for slno, window in enumerate(child_windows):
                log.info('onDeleteWindow - wait_for_exit - Waiting for - ' + str(slno))
                window.wait()
            log.info('onDeleteWindow - wait_for_exit - All windows closed. Exiting')
            Gtk.main_quit()
            
        # Ask confirmation from user
        if self.stack.haschanged():
            message = 'You have unsaved changes which will be lost if you continue.\n Are you sure you want to exit ?'
            title = 'Confirm Exit'
            dialogWindow = Gtk.MessageDialog(self.window,
                                     Gtk.DialogFlags.MODAL | Gtk.DialogFlags.DESTROY_WITH_PARENT,
                                     Gtk.MessageType.QUESTION,
                                     Gtk.ButtonsType.YES_NO,
                                     message)
            dialogWindow.set_transient_for(self.window)
            dialogWindow.set_title(title)
            dialogWindow.set_default_response(Gtk.ResponseType.NO)
            dialogWindow.show_all()
            response = dialogWindow.run()
            dialogWindow.destroy()
            if response == Gtk.ResponseType.NO:
                # Do not propogate signal
                log.info('onDeleteWindow - Cancelled by user')
                return True
        
        # Check for status of child windows
        if self.child_windows:
            self.window.hide()
            # Wait for all child windows to exist after returning from method
            GLib.timeout_add(50, wait_for_exit, self.child_windows)
            # Propogate delete event to destroy window
            log.info('onDeleteWindow - Waiting for child to exit')
            return False
        else:
            log.info('onDeleteWindow - No child windows. Exiting')
            Gtk.main_quit()

    def onNewProjectClicked(self, button):
        """Create a new window"""
        proc = subprocess.Popen([__file__], stdin=None, stdout=None, stderr=None)
        self.child_windows.append(proc)
        log.info('onNewProjectClicked - New window raised')

    def onOpenProjectClicked(self, button):
        """Open project selected by  the user"""
        
        # Create a filechooserdialog to open:
        # The arguments are: title of the window, parent_window, action,
        # (buttons, response)
        open_dialog = Gtk.FileChooserDialog("Open project File", self.window,
                                            Gtk.FileChooserAction.OPEN,
                                            (Gtk.STOCK_CANCEL, Gtk.ResponseType.CANCEL,
                                             Gtk.STOCK_OPEN, Gtk.ResponseType.ACCEPT))
        # Remote files can be selected in the file selector
        open_dialog.set_local_only(True)
        # Dialog always on top of the textview window
        open_dialog.set_modal(True)
        # Set filters
        open_dialog.set_filter(self.builder.get_object("filefilter_project"))
        # Set window position
        open_dialog.set_gravity(Gdk.Gravity.CENTER)
        open_dialog.set_position(Gtk.WindowPosition.CENTER_ON_PARENT)

        response_id = open_dialog.run()
        # If response is "ACCEPT" (the button "Save" has been clicked)
        if response_id == Gtk.ResponseType.ACCEPT:
            # get filename and set project as active
            self.filename = open_dialog.get_filename()
            fileobj = open(self.filename, 'r')
            if fileobj == None:
                log.error("onOpenProjectClicked - Error opening file - " + self.filename)
                self.display_status(misc.CMB_ERROR, "Project could not be opened: Error opening file")
            else:
                try:
                    data = json.load(fileobj)  # load data structure
                    fileobj.close()
                    if data[0] == misc.PROJECT_FILE_VER:
                        self.data.set_model(data[1])
                        self.project_settings_dict = data[2]

                        self.display_status(misc.CMB_INFO, "Project successfully opened")
                        log.info('onOpenProjectClicked - Project successfully opened - ' +self.filename)
                        # Setup paths for folder chooser objects
                        self.builder.get_object("filechooserbutton_meas").set_current_folder(misc.posix_path(
                            os.path.split(self.filename)[0]))
                        self.builder.get_object("filechooserbutton_bill").set_current_folder(misc.posix_path(
                            os.path.split(self.filename)[0]))
                        # Setup window name
                        self.window.set_title(self.filename + ' - ' + misc.PROGRAM_NAME)
                        # Clear undo/redo stack
                        self.stack.clear()
                        # Set flags
                        self.project_active = True
                        # Save point in stack for checking change state
                        self.stack.savepoint()
                        # Refresh all displays
                        self.update()
                    else:
                        self.display_status(misc.CMB_ERROR, "Project could not be opened: Wrong file type selected")
                        log.warning('onOpenProjectClicked - Project could not be opened: Wrong file type selected - ' +self.filename)
                except:
                    log.exception("Error parsing project file - " + self.filename)
                    self.display_status(misc.CMB_ERROR, "Project could not be opened: Error opening file")
        # If response is "CANCEL" (the button "Cancel" has been clicked)
        elif response_id == Gtk.ResponseType.CANCEL:
            log.info("cancelled: FileChooserAction.OPEN")
        # Destroy dialog
        open_dialog.destroy()

    def onSaveProjectClicked(self, button):
        """Save project to file already opened"""
        if self.project_active is False:
            self.onSaveAsProjectClicked(button)
        else:
            # Parse data into object
            data = []
            data.append(misc.PROJECT_FILE_VER)
            data.append(self.data.get_model())
            data.append(self.project_settings_dict)
            # Try to open file
            fileobj = open(self.filename, 'w')
            if fileobj == None:
                log.error("onSaveProjectClicked - Error opening file - " + self.filename)
                self.display_status(misc.CMB_ERROR, "Project file could not be opened for saving")
            json.dump(data, fileobj)
            fileobj.close()
            self.display_status(misc.CMB_INFO, "Project successfully saved")
            log.info('onSaveProjectClicked -  Project successfully saved')
            self.window.set_title(self.filename + ' - ' + misc.PROGRAM_NAME)

    def onSaveAsProjectClicked(self, button):
        """Save project to file selected by the user"""
        # Create a filechooserdialog to open:
        # The arguments are: title of the window, parent_window, action,
        # (buttons, response)
        open_dialog = Gtk.FileChooserDialog("Save project as...", self.window,
                                            Gtk.FileChooserAction.SAVE,
                                            (Gtk.STOCK_CANCEL, Gtk.ResponseType.CANCEL,
                                             Gtk.STOCK_SAVE, Gtk.ResponseType.ACCEPT))
        # Remote files can be selected in the file selector
        open_dialog.set_local_only(False)
        # Dialog always on top of the textview window
        open_dialog.set_modal(True)
        # Set filters
        open_dialog.set_filter(self.builder.get_object("filefilter_project"))
        # Set window position
        open_dialog.set_gravity(Gdk.Gravity.CENTER)
        open_dialog.set_position(Gtk.WindowPosition.CENTER_ON_PARENT)
        # Set overwrite confirmation
        open_dialog.set_do_overwrite_confirmation(True)
        # Set default name
        open_dialog.set_current_name("newproject.proj")
        response_id = open_dialog.run()
        # If response is "ACCEPT" (the button "Save" has been clicked)
        if response_id == Gtk.ResponseType.ACCEPT:
            # Get filename and set project as active
            self.filename = open_dialog.get_filename()
            self.project_active = True
            # Call save project
            self.onSaveProjectClicked(button)
            # Setup paths for folder chooser objects
            self.builder.get_object("filechooserbutton_meas").set_current_folder(misc.posix_path(
                os.path.split(self.filename)[0]))
            self.builder.get_object("filechooserbutton_bill").set_current_folder(misc.posix_path(
                os.path.split(self.filename)[0]))
            # Setup window name
            self.window.set_title(self.filename + ' - ' + misc.PROGRAM_NAME)
            
            log.info('onSaveAsProjectClicked -  Project successfully saved - ' + self.filename)
        # If response is "CANCEL" (the button "Cancel" has been clicked)
        elif response_id == Gtk.ResponseType.CANCEL:
            log.info("cancelled: FileChooserAction.OPEN")
        # Destroy dialog
        open_dialog.destroy()
        
    def onProjectSettingsClicked(self, button):
        """Display dialog to input project settings"""
        log.info('onProjectSettingsClicked - Launch project settings')
        item_values = [self.project_settings_dict[key] for key in misc.global_vars]
        # Setup project settings dialog
        project_settings_dialog = misc.UserEntryDialog(self.window, 
                                      'Project Settings',
                                      item_values,
                                      misc.global_vars_captions)
        # Show settings dialog
        project_settings_dialog.run()
        for key,item in zip(misc.global_vars,item_values):
            self.project_settings_dict[key] = item

    def onInfobarClose(self, widget, response=0):
        """Hides the infobar"""
        widget.hide()

    def onRedoClicked(self, button):
        """Redo action from stack"""
        log.info('Redo:' + str(self.stack.redotext()))
        self.stack.redo()
        self.update()

    def onUndoClicked(self, button):
        """Undo action from stack"""
        log.info('Undo:' + str(self.stack.undotext()))
        self.stack.undo()
        self.update()

    # Schedule signal handler methods

    def onButtonScheduleAddPressed(self, button):
        """Add empty row to schedule view"""
        items = []
        items.append(data.schedule.ScheduleItem())
        self.schedule_view.insert_item_at_selection(items)

    def onButtonScheduleAddMultPressed(self, button):
        """Add multiple empty rows to schedule view"""
        user_input = misc.get_user_input_text(self.window, "Please enter the number \nof rows to be inserted",
                                        "Number of Rows")
        try:
            number_of_rows = int(user_input)
        except:
            self.display_status(misc.CMB_WARNING, "Invalid number of rows specified")
            return
        items = []
        for i in range(0, number_of_rows):
            items.append(data.schedule.ScheduleItem())
        self.schedule_view.insert_item_at_selection(items)

    def onButtonScheduleDeletePressed(self, button):
        """Delete selected rows from schedule view"""
        self.schedule_view.delete_selected_rows()

    def onCopySchedule(self, button):
        """Copy selected rows from schedule view to clipboard"""
        self.schedule_view.copy_selection()

    def onPasteSchedule(self, button):
        """Paste rows from clipboard into schedule view"""
        self.schedule_view.paste_at_selection()

    def onImportScheduleClicked(self, button):
        """Imports rows from spreadsheet selected by 'filechooserbutton_schedule' into schedule view"""
        filename = self.builder.get_object("filechooserbutton_schedule").get_filename()
        spreadsheet = misc.Spreadsheet(filename)
        columntypes = [misc.MEAS_DESC, misc.MEAS_DESC, misc.MEAS_DESC,
                       misc.MEAS_L, misc.MEAS_L, misc.MEAS_DESC, misc.MEAS_L]
        models = spreadsheet.read_rows(columntypes = columntypes)
        items = []
        for model in models:
            item = data.schedule.ScheduleItem(*model)
            items.append(item)
        self.schedule_view.insert_item_at_selection(items)

    # Measuremets signal handlers

    def onAddCmbClicked(self, button):
        """Add a CMB object to measurement view"""
        self.measurements_view.add_cmb()

    def onAddMeasClicked(self, button):
        """Add a Measurement object to measurement view"""
        code = self.measurements_view.add_measurement()
        if code is not None:
            self.display_status(*code)

    def onAddComplClicked(self, button):
        """Add a Completion object to measurement view"""
        code = self.measurements_view.add_completion()
        if code is not None:
            self.display_status(*code)

    def onAddHeadingClicked(self, button):
        """Add a Heading object to measurement view"""
        code = self.measurements_view.add_heading()
        if code is not None:
            self.display_status(*code)
            
    def onMeasCustomMenuClicked(self,button,module):
        """Callback function for click event of custom measurement item"""
        code = self.measurements_view.add_custom(None,module)
        if code is not None:
            self.display_status(*code)

    def onAddAbstractClicked(self, button):
        """Add a measurement abstract object to measurement view"""
        code = self.measurements_view.add_abstract(None)
        if code is not None:
            self.display_status(*code)

    def OnMeasDeleteClicked(self, button):
        """Delete selected item from measurement view"""
        self.measurements_view.delete_selected_row()

    def OnMeasRenderClicked(self, button):
        """Renders selected CMB to directory selected by 'filechooserbutton_meas'"""
        filechooserbutton_meas = self.builder.get_object("filechooserbutton_meas")
        if filechooserbutton_meas.get_file() != None:
            folder = filechooserbutton_meas.get_file().get_path()
            code = self.measurements_view.render_selection(folder, self.project_settings_dict)
            self.display_status(*code)
            # remove temporary files
            onlytempfiles = [f for f in os.listdir(misc.posix_path(folder)) if (f.find('.aux')!=-1 or f.find('.log')!=-1
                                or f.find('.out')!=-1) or f.find('.tex')!=-1 or f.find('.bak')!=-1]
            for f in onlytempfiles:
                os.remove(misc.posix_path(folder,f))
        else:
            self.display_status(misc.CMB_ERROR, 'Please select an output directory for rendering')
        
    def OnMeasClickEvent(self, button, event):
        """Edit measurement view object on double click"""
        if event.type == Gdk.EventType._2BUTTON_PRESS:
            self.measurements_view.edit_selected_row()

    def OnMeasEditClicked(self, button):
        """Edit selected item in measurement view"""
        self.measurements_view.edit_selected_row()

    def OnMeasPropertiesClicked(self, button):
        """Edit properties of supported items in measurement view"""
        code = self.measurements_view.edit_selected_properties()
        if code is not None:
            self.display_status(*code)

    def OnMeasCopyClicked(self, button):
        """Copy selected item in measurement view to clipboard"""
        self.measurements_view.copy_selection()

    def OnMeasPasteClicked(self, button):
        """Paste item from clipboard to measurement view"""
        self.measurements_view.paste_at_selection()

    # Bills signal handlers

    def onAddBillClicked(self, button):
        """Add a bill to bill view"""
        self.bill_view.add_bill()

    def onAddBillCustomClicked(self, button):
        """Add a custom bill to bill view"""
        self.bill_view.add_bill_custom()

    def OnBillDeleteClicked(self, button):
        """Delete a bill from bill view"""
        self.bill_view.delete_selected_row()

    def OnBillRenderClicked(self, button):
        """Renders selected Bill to directory selected by 'filechooserbutton_bill'"""
        filechooserbutton_bill = self.builder.get_object("filechooserbutton_bill")
        if filechooserbutton_bill.get_file() != None:
            folder = filechooserbutton_bill.get_file().get_path()
            code = self.bill_view.render_selected(folder, self.project_settings_dict)
            self.display_status(*code)
            # remove temporary files
            onlytempfiles = [f for f in os.listdir(misc.posix_path(folder)) if (f.find('.aux')!=-1 or f.find('.log')!=-1
                                or f.find('.out')!=-1) or f.find('.tex')!=-1 or f.find('.bak')!=-1]
            for f in onlytempfiles:
                os.remove(misc.posix_path(folder,f))
        else:
            self.display_status(misc.CMB_ERROR, 'Please select an output directory for rendering')
        
    def OnBillClickEvent(self, button, event):
        """Edit bill object on double click"""
        if event.type == Gdk.EventType._2BUTTON_PRESS:
            self.bill_view.edit_selected_row()

    def OnBillEditClicked(self, button):
        """Edit selected bill"""
        self.bill_view.edit_selected_row()

    def OnBillCopyClicked(self, button):
        """Copy bill to clipboard"""
        self.bill_view.copy_selection()

    def OnBillPasteClicked(self, button):
        """Paste bill in clipboard to bill view"""
        self.bill_view.paste_at_selection()

    # Tab methods

    def onSwitchTab(self, widget, page, pagenum):
        """Refresh display on switching between views"""
        log.info('onSwitchTab called - ' + str(pagenum))
        self.update()

    def __init__(self):
        log.info('MainWindow - initialise')
        # Variable used to store handles for child window processes
        self.child_windows = []

        # Check for project active status
        self.project_active = False
        
        # Setup main data model
        self.data = data.datamodel.DataModel()
        
        # Initialise undo/redo stack
        self.stack = undo.Stack()
        undo.setstack(self.stack)
        # Save point in stack for checking change state
        self.stack.savepoint()

        # Other variables
        self.filename = None

        # Setup main window
        self.builder = Gtk.Builder()
        self.builder.add_from_file(misc.abs_path("interface", "mainwindow.glade"))
        self.window = self.builder.get_object("window_main")
        self.builder.connect_signals(self)
        # Setup infobar
        self.builder.get_object("infobar_main").hide()

        # Setup project settings dictionary
        self.project_settings_dict = dict()
        for item_code in misc.global_vars:
            self.project_settings_dict[item_code] = ''
        
        # Load global Variables
        misc.set_global_platform_vars()
        
        # Setup about dialog
        self.about_dialog = self.builder.get_object("aboutdialog")

        # Setup schedule View
        self.treeview_schedule = self.builder.get_object("treeview_schedule")
        self.schedule_view = view.schedule.ScheduleView(self.window, self.treeview_schedule, self.data.schedule)
        
        # Setup bill View
        self.treeview_bill = self.builder.get_object("treeview_bill")
        self.bill_view = view.bill.BillView(self.window, self.data, self.treeview_bill)
        
        # Setup measurement View
        self.treeview_meas = self.builder.get_object("treeview_meas")
        self.measurements_view = view.measurement.MeasurementsView(self.window, self.data, self.treeview_meas)
        
        # Setup custom measurement items
        file_names = [f for f in os.listdir(misc.abs_path('templates'))]
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
                log.info('Plugin loaded - ' + module_name)
            except ImportError:
                log.error('Error Loading plugin - ' + module_name)

    def run(self, *args):
        self.window.show_all()


def main():

    # Setup Logging to temporary file
    log_file = tempfile.NamedTemporaryFile(mode='w', prefix='cmbautomiser_', 
                                               suffix='.log', delete=False)
    logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                        stream=log_file,level=logging.INFO)
    logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                        stream=sys.stdout,level=logging.INFO)
    # Logging to stdout
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    ch.setFormatter(formatter)
    log.addHandler(ch)
    # Log all uncaught exceptions
    def handle_exception(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        log.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))
    sys.excepthook = handle_exception

    # Initialise main window
    log.info('Start Program Execution')
    MainWindow().run()
    log.info('Main window initialised - Running Gtk.main()')
    Gtk.main()
    return 0


if __name__ == '__main__':
    main()

