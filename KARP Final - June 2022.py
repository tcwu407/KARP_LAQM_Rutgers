# -*- coding: utf-8 -*-
"""
Created on Thu Jun 10 16:56:39 2021

@author: pbuttles
"""

""" An application for connecting to and controlling a RP100 power supply, """
""" a Keysight E4980AL capcacitance bridge, plotting, and saving data from both. """
""" Programmed by Paul Buttles under the supervision of Tsung-Chi Wu """
""" and Dr. Jak Chakhalian of the LAQM at Rutgers University. """
""" Modification of Jack Barraclough's (jack@razorbillinstruments.com) app. """


""""NOTE: Safety checks have been turned off for voltage between -210V and 210V"""

import serial.tools.list_ports as list_ports
import serial
from tkinter import *
import tkinter.simpledialog
from enum import Enum
import time
import datetime
import os
import numpy as np
import matplotlib.pyplot as plt
import csv
from itertools import count
import matplotlib.animation as animation
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import ttk
import pyvisa

rm = pyvisa.ResourceManager()


"""Changes working directory to the folder of the script.
Fixes "image 'pyimageX' doesn't exist" error."""
os.chdir(os.path.dirname(os.path.abspath(__file__)))

class ScpiProperty:
    """
    Each SCPI property of the instrument has a ScpiProperty in the GUI.  This is a base class which should be subclassed 
    to create a class for each type of SCPI property (float, int, bool...). 
    """
    def __init__(self, parent, row, ser, command, description, can_set=True, can_get=True):
        self.command = command
        self.ser = ser
        self.value = None
        self.heldvalue = None
        self.can_get = can_get
        self._interactable_widgets = []
        """create text labels to the left of the widgets"""
        if description != "":
            if type(description) == list:
                for i in range(len(description)):
                    Label(parent, text=description[i]).grid(row=row+i, column=1)
            else:
                Label(parent, text=description).grid(row=row, column=1)

    def scpi2human(self, scpi_bytes):
        """ Must be overridden by subclass"""
        raise NotImplementedError

    def human2scpi(self, value):
        """ Must be overriden by subclass"""
        raise NotImplementedError

    def scpi_get(self):
        """ Get a property's value from the instrument, and update the GUI to reflect it"""
        """ if statement is for RP100, elif is for Keysight, else is for anything else"""
        if type(self.command) == bytes:
            self.ser.write(self.command + b"?\n")
            resp = self.ser.read()
            self.value.set(self.scpi2human(resp))
        elif self.command == ":FETCh:IMPedance:FORMatted?":
            perfTime2=time.time()
            self.ser.write(self.command)
            resp = self.ser.read().strip("\n").strip().split(",")
            for i in range(len(resp)):
                self.value[i].set(float(resp[i]))
            print(time.time()-perfTime2)
                
        else:
            self.ser.write(self.command)
            resp = self.ser.read()
            self.value.set(resp)

    def scpi_set(self):
        """ Take a property's value from the GUI, and send it to the instrument """
        scpi_value = self.human2scpi(self.value.get())
        self.ser.write(self.command + b" " + scpi_value + b"\n")
        #win.focus_set()
        try:
            self.heldvalue.set(self.value.get())
        except: pass

    def disable(self):
        """ Disable all the widgets which make up this ScpiProperty, e.g. if there is no instrument to work with"""
        for widget in self._interactable_widgets:
            widget.config(state="disabled")

    def enable(self):
        """ Enable all the widgets which make up this ScpiProperty, e.g. when the instrument is connected"""
        for widget in self._interactable_widgets:
            if str(widget) == ".!notebook.!frame.!frame5.!combobox" or str(widget) == ".!notebook.!frame.!frame5.!combobox2" or str(widget) == ".!notebook.!frame.!frame5.!combobox3":
                widget.config(state="readonly")
            else:
                widget.config(state="normal")
        if self.can_get:
            self.scpi_get()
            
    def lock(self):
        """sets widgets to readonly"""
        for widget in self._interactable_widgets:
            widget.config(state="disabled")
    def unlock(self):
        for widget in self._interactable_widgets:
            widget.config(state="normal")
            
    def snapback(self,event):
        """ Entry boxes return to their actual held value if clicked off of without hitting Enter"""
        self.value.set(self.heldvalue.get())
            
    def setwrapper(self,event):
        """ A wrapper of warning dialogues for scpi_set """
        """ Target Voltage warnings for RP100, at room temperature (-20 to 120), or absolute (-200 to 200)"""
        if str(self.command) == "b'SOUR1:VOLT'" or str(self.command) == "b'SOUR2:VOLT'":
            if float(self.value.get()) >= 210.0 or float(self.value.get()) <= -210.0:
                if float(self.value.get()) >= 210.0 or float(self.value.get()) <= -210.0:
                    MsgBox = messagebox.askquestion("High Voltage Detected!", "Warning: The RP100 Power Supply is rated for up to Â±200V. The power supply also allows some over-range capability, of no less than Â±210 volts, and typically around Â±225 V depending on the load characteristics and small variations from supply to supply. When using this over-range, the noise performance is reduced, and at the ends of the range, the accuracy and linearity will be poor. Using this over-range capability with a Razorbill Instruments cell will probably reduce its service life considerably, and should be done with caution. \n\nAre you sure you want to continue?",icon="error")
                    if MsgBox == 'yes': self.scpi_set()
                    else: return
                else:
                    MsgBox = messagebox.askquestion("High Voltage Detected!", "Warning: At room temperature, ensure that the RP100 Power Supply voltages do not go below -20V or above +120V. \n\nThe more extreme voltages the RP100 can supply should only be applied at cryogenic temperature in accordance with the strain cell datasheet. Exceeding these limits will damage the piezoelectric stacks. \n\nWould you like to continue?",icon="warning")
                    if MsgBox == 'yes': self.scpi_set()
                    else: return
            else: self.scpi_set()
        """ Slew Rate warnings for RP100, for above 100 V/s, or for low resolution piece-wise stepping below 0.0005 V/s"""
        if str(self.command) == "b'SOUR1:VOLT:SLEW'" or str(self.command) == "b'SOUR2:VOLT:SLEW'":
            if float(self.value.get()) >= 100.0:
                MsgBox = messagebox.askquestion("High Slew Rate Detected!", "It is generally advisable to keep slew rates below 100V/s for piezoelectric devices which are not designed for high frequency operation, and if the device is in a cryostat, slower rates will reduce unwanted heating. \n\nAre you sure you want to continue?", icon="question")
                if MsgBox == 'yes': self.scpi_set()
                else: return
            if float(self.value.get()) <= 0.0005:
                MsgBox = messagebox.askquestion("Low Slew Rate Detected!", "A smooth ramp is possible for slew rates above 0.5mV/s. For rates below that, the output can take on a staircase shape, as the output changes by one least significant bit at a time. \n\nWould you like to continue?",icon="question")
                if MsgBox == 'yes': self.scpi_set()
            else: self.scpi_set()
        else: self.scpi_set()

class ScpiPropertyFloat(ScpiProperty):
    """ A ScpiProperty representing a SCPI property on the instrument which can take float values"""
    def __init__(self, parent, row, ser, command, description, can_set=True, can_get=True, width=20):
        super().__init__(parent, row, ser, command, description, can_set, can_get)
        self.value = StringVar()
        self.heldvalue = StringVar()
        self.row = row
        """ generates setttable widgets (target voltages, slew rates)"""
        if can_set:
            textbox = Entry(parent, textvariable=self.value)
            textbox.grid(row=row, column=2, padx=(10,2),pady=5)
            textbox.configure(width=width)
            textbox.bind("<Return>",self.setwrapper)
            textbox.bind("<FocusOut>",self.snapback)
            self.description = description
            self._interactable_widgets.append(textbox)
        else:
            """ Parse's Keysight's single command into two outputs and generates their widgets"""
            if self.command == ":FETCh:IMPedance:FORMatted?":
                self.description = [description[0],description[1]]
                self.value = [StringVar(),StringVar(),StringVar()]
                infolabel1 = Label(parent, textvariable=self.value[0])#self.value.get())#.split(",")[0])
                infolabel1.grid(row=row, column=2)
                infolabel2 = Label(parent, textvariable=self.value[1])#self.value.get().split(",")[1])
                infolabel2.grid(row=row + 1, column=2)
            else:
                """ generates most widgets """
                self.description = description
                infolabel = Label(parent, textvariable=self.value)
                infolabel.grid(row=row, column=2)

    def human2scpi(self, value):
        return bytes(value, 'utf8')

    def scpi2human(self, scpi_bytes):
        if scpi_bytes is None:
            return ""
        else:
            try:
                resp = scpi_bytes.decode()
            except:
                resp = scpi_bytes
            try:
                num = float(resp)
            except ValueError:
                return ""
            return str(num)


class ScpiPropertyBool(ScpiProperty):
    """ A ScpiProperty representing a SCPI property on the instrument which can take boolean values"""
    def __init__(self, parent, row, ser, command, description, can_set=True, can_get=True):
        super().__init__(parent, row, ser, command, description, can_set, can_get)
        self.value = IntVar()
        self.value.set(0)
        self.heldvalue = IntVar()
        self.description = description
        """ generates two radio buttons inside a frame, only used for output relay"""
        frame = Frame(parent)
        frame.grid(row=row, column=2)
        radio = Radiobutton(frame, text="Enable", variable=self.value, value=1)
        radio.grid(row=1, column=1)
        radio.bind('<Button-1>',self.setwrapper)
        self._interactable_widgets.append(radio)
        radio = Radiobutton(frame, text="Disable", variable=self.value, value=0)
        radio.grid(row=1, column=2)
        radio.bind("<ButtonRelease-1>",self.setwrapper)
        self._interactable_widgets.append(radio)
        

        #generates lock/unlock radiobuttons WIP

        

    def scpi2human(self, scpi_bytes):
        if (scpi_bytes is None) or (scpi_bytes.decode().rstrip() == ""):
            return -1
        else:
            return int(scpi_bytes.decode().rstrip())

    def human2scpi(self, value):
        return bytes(str(value), 'utf8')

class ScpiPropertyCombobox(ScpiProperty):
    """ A ScpiProperty representing the values that can be chosen for independent/dependent variables when graphing"""
    def __init__(self, parent, row, column, ser, command, values_list, description, can_set=True, can_get=True):
        super().__init__(parent, row, ser, command, description, can_set, can_get)
        self.value = StringVar()
        #print("Hello")
        #print(values_list[0])
        #self.value.set(values_list[0])
        combo = ttk.Combobox(parent)
        combo.grid(row=row,column=column)
        combo.configure(values=values_list)
        combo.configure(state='readonly')
        combo.configure(textvariable=values_list[0])
        combo.configure(width=3)
        combo.configure(takefocus="")
        combo.configure(text=values_list[0])
        #print(combo.value[0])
        #print("hiya")
        self.value.set(values[0].get())
        self._interactable_widgets.append(combo)

    def human2scpi(self, value):
        return bytes(value, 'utf8')

    def scpi2human(self, scpi_bytes):
        if scpi_bytes is None:
            return ""
        else:
            resp = scpi_bytes.decode()
            try:
                num = str(resp)
            except ValueError:
                return ""
            return str(num)

class ScpiErrorReporter(ScpiProperty):
    """ A class for reporting the error state of the instrument """
    def __init__(self, parent, row, ser):
        super().__init__(parent, row, ser, b'SYST:ERR', "Last Error:", can_set=False)
        self.value = StringVar()
        textbox = Label(parent, textvariable=self.value, relief=SUNKEN)
        textbox.grid(row=row, column=1, sticky="NSWE")

    def scpi2human(self, scpi_bytes):
        if scpi_bytes is None:
            return ""
        else:
            return scpi_bytes.decode().strip()


class SerialStates(Enum):
    UNCONFIGURED = 1
    CONNECTED = 2
    DROPPED = 3
    
class USBStates(Enum):
    UNCONFIGURED = 1
    CONNECTED = 2
    DROPPED = 3


class MonitoredSerial:
    """ 
    A class for serial connections, with some extra wrappers to release the port if the device is unplugged
    or dropped, and grab it again when it reappears. Call update() about once every millisecond. 
    """
    def __init__(self, printer=None, print_io=False, print_conn=False):
        self._pid = None
        self._vid = None
        self._serial_number = None
        self._port = None
        self._printer = printer
        self._print_io = print_io
        self._print_conn = print_conn
        self.state = SerialStates.UNCONFIGURED
        self.needs_reset = False

    """ Used in choose_serial_port, takes the result of PortChooser as port_info, and attempts to open a serial connection"""
    def connect(self, port_info):
        try:
            self._port = serial.Serial(port_info.device, timeout=0.1)
        except Exception as e:
            if self._printer is not None:
                self._printer("Failed to open serial port: " + str(e))
        else:
            if not self._port.is_open:
                self._port.open()
            if self._print_conn:
                self._printer("Opened serial port: " + port_info.description)
            self._serial_number = port_info.serial_number
            self._pid = port_info.pid
            self._vid = port_info.vid
            self.state = SerialStates.CONNECTED

    """ Checks if the RP100 is still there"""
    def update(self):
        ports = list_ports.comports()
        if self._port is not None:
            if (self._port.name in (i.device for i in ports)) and not self.needs_reset:
                # All is OK, the port is still there
                self.state = SerialStates.CONNECTED
                return False
            else:
                # Our port has vanished, or needs_reset has been set by something else
                self._port.close()
                self._port = None
                if self._print_conn:
                    self._printer("Lost connection to serial port")
                self.state = SerialStates.DROPPED
                time.sleep(0.05)
                return True
        else:
            if self._serial_number is not None:
                # Our port is missing. Look for it
                for port in ports:
                    if port.serial_number == self._serial_number and port.vid == self._vid and port.pid == self._pid:
                        try:
                            self._port = serial.Serial(port.device, timeout=0.1)
                            if not self._port.is_open:
                                self._port.open()
                            if self._print_conn:
                                self._printer("Reopened serial port " + port.device)
                            self.state = SerialStates.CONNECTED
                            return True
                        except:
                            if self._printer is not None:
                                self._printer("Failed to reopen serial port")
                                time.sleep(0.1)
            else:
                # Not configured, so no change.
                return False

    """ Back-end command for disconnecting the RP100"""
    def disconnect(self):
        if self._print_conn:
            self._printer("Manually disconnected RP100 from serial port")
        self._port.close()
        self._pid = None
        self._vid = None
        self._serial_number = None
        self._port = None
        self.state = SerialStates.UNCONFIGURED

    """ Reads from the RP100 using readline(), otherwise prints errors to the printer"""
    def read(self):
        if self.state != SerialStates.CONNECTED:
            return None
        else:
            try:
                resp = self._port.readline()
                if self._print_io:
                    if resp.decode().strip() == "":
                        self._printer("Timeout or empty line on serial read")
            except Exception as e:
                resp = b""
                if self._print_io:
                    self._printer("IO Error on Serial Read: " + str(e))
                self.needs_reset = True
            return resp

    """ Writes to the RP100 using write(), otherwise prints errors to the printer"""
    def write(self, message):
        if self.state != SerialStates.CONNECTED:
            pass
        else:
            try:
                self._port.write(message)
            except Exception as e:
                if self._print_io:
                    self._printer("IO Error on Serial Write: " + str(e))
                self.needs_reset = True

class MonitoredUSB:
    """ 
    A class for USB connections, with some extra wrappers to release the port if the device is unplugged
    or dropped, and grab it again when it reappears. Call update() about once every millisecond. 
    """
    def __init__(self, printer=None, print_io=False, print_conn=False):
        self._alias = None
        self._name = None
        self._serial_number = None
        self._port = None
        self._printer = printer
        self._print_io = print_io
        self._print_conn = print_conn
        self.state = USBStates.UNCONFIGURED
        self.needs_reset = False
        self.usb_ports = []

    """ Used in choose_usb_port, takes the result of PortChooser as port_info, and attempts to open a USB connection"""
    def connect(self, port_info):
        try:
            self._port = rm.open_resource(port_info)
            try: self._port = rm.open_resource(self._port.resource_info.alias)
            except: pass
        except Exception as e:
            if self._printer is not None:
                self._printer("Failed to open serial port: " + str(e))
        else:
            if self._print_conn:
                self._printer("Opened port: " + self._port.resource_info.alias)
            self._alias = self._port.resource_info.alias
            self._name = self._port.resource_info.resource_name
            self._serial_number = self._name.split("::")[3]
            self.state = USBStates.CONNECTED
            

    """ Checks if the Keysight is still there """
    def update(self):
        self.usb_ports = []
        self.resources = rm.list_resources()
        for i in range(len(self.resources)):
            if str(self.resources[i])[0:3] == "USB":
                self.usb_ports.append(self.resources[i])
        if self._port is not None:
            if (self._serial_number in (rm.open_resource(i).resource_info.resource_name.split("::")[3] for i in self.usb_ports)) and not self.needs_reset:
                # All is OK, the port is still there
                self.state = USBStates.CONNECTED
                return False
            else:
                # Our port has vanished, or needs_reset has been set by something else
                self._port = None
                if self._print_conn:
                    self._printer("Lost connection to USB port")
                self.state = USBStates.DROPPED
                self.MainGui.status_box2.configure(background='red')
                time.sleep(0.05)
                return True
        else:
            if self._serial_number is not None:
                # Our port is missing. Look for it
                for port in self.usb_ports:
                    try:
                        portcheck = rm.open_resource(port)
                    except: pass
                    else:
                        namesplit=portcheck.resource_info.resource_name.split("::")
                        if namesplit[3] == self._serial_number:
                            self._printer("Reopened port " + str(namesplit[0]))
                            self.state = USBStates.CONNECTED
                            return True
                if self._printer is not None:
                    self._printer("Failed to find USB port")
                    time.sleep(0.1)
            else:
                # Not configured, so no change.
                return False

    """ Back-end command for disconnecting the Keysight"""
    def disconnect(self):
        if self._print_conn:
            self._printer("Manually disconnected Keysight from USB port")
        self._serial_number = None
        self._port = None
        self._alias = None
        self._name = None
        self.state = USBStates.UNCONFIGURED
        
    
    """ Reads from the Keysight using read(), otherwise prints errors to the printer"""
    def read(self):
        if self.state != USBStates.CONNECTED:
            return None
        else:
            try:
                resp = self._port.read()
                if self._print_io:
                    if resp == "":
                        self._printer("Timeout or empty line on read")
            except Exception as e:
                resp = ""
                if self._print_io:
                    self._printer("IO Error on USB Read: " + str(e))
                self.needs_reset = True
            return resp
    
    """ Writes to the Keysight using write(), otherwise prints errors to the printer"""
    def write(self, message):
        if self.state != USBStates.CONNECTED:
            pass
        else:
            try:
                
                self._port.write(message)
            except Exception as e:
                if self._print_io:
                    self._printer("IO Error on USB Write: " + str(e))
                self.needs_reset = True
    
class MainGui:
    """Main class"""
    def __init__(self):
        self.serial_port = MonitoredSerial(printer=self.printer, print_conn=True, print_io=True)
        self.usb_port = MonitoredUSB(printer=self.printer, print_conn=True, print_io=True)
        self.win = None
        self.log_text = None
        self.recording = False
        self.isPlotOn = False
        self.indvar = None
        self.depvar = None
        self.counter = 0
        self._scpi_properties = []
        self.data = []

        self.build_main_window()
        self.start()

    """ Prints a message to the printer"""
    def printer(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(END, message + '\n')
        self.log_text.config(state="disabled")

    """ Starts the program and begins the main_task loop"""
    def start(self):
        self.win.after(1, self.main_task)
        self.win.mainloop()
    
    """ Main loop for the software, repeats until the program is closed"""
    def main_task(self):
        
        #perfTime = time.time()
        
        """ Live-plot animation function """
        def animate(self):
            
            self.xval.append(timestepvalues[self.indcombo.current()])        
            self.yval.append(timestepvalues[self.depcombo.current()])
            self.scattery.set_offsets(np.c_[self.xval,self.yval])
            
            if self.indcombo.current() == 14:
                self.ax.set_xlim([time.time()-init_time-10,time.time()-init_time+5]) #updates x axis as time passes
            self.canvas.draw_idle()
        
        """ Check if connections have changed and act accordingly"""
        serial_has_changed = self.serial_port.update()
        usb_has_changed = self.usb_port.update()
        if serial_has_changed:
            self.status_box1.config(text=str(self.serial_port.state.name))
            if self.serial_port.state == SerialStates.CONNECTED:
                for i in range(12):
                    self._scpi_properties[i].enable()
            else:
                for i in range(12):
                    self._scpi_properties[i].disable()
        if usb_has_changed:
            self.status_box2.config(text=str(self.usb_port._port))
            if self.usb_port.state == USBStates.CONNECTED:
                for i in range(len(self._scpi_properties)-12):
                    self._scpi_properties[i+12].enable()
            else:
                for i in range(len(self._scpi_properties)-12):
                    self._scpi_properties[i+12].disable()
        
        """ Live-update GUI with new values from instruments"""
        if self.serial_port.state == SerialStates.CONNECTED:
            for i in range(3):
                self._scpi_properties[i+3].scpi_get()
                self._scpi_properties[i+9].scpi_get()
        if self.usb_port.state.name == "CONNECTED":
            self._scpi_properties[12].scpi_get()
            
        """ Live-plot and record data """
        if self.recording == True:
            #init_time = time.time()
            if self.serial_port.state == SerialStates.CONNECTED or self.usb_port.state == USBStates.CONNECTED:
                """ Record Data """
                timestepvalues = np.zeros(15, dtype = float)
                if self.serial_port.state == SerialStates.CONNECTED:
                    #for i in range(3):
                        #timestepvalues[i] = float(self._scpi_properties[i].heldvalue.get())
                        #timestepvalues[i+6] = float(self._scpi_properties[i+6].heldvalue.get())
                        #print(float(self._scpi_properties[1].heldvalue.get()))
                    for i in range(3):
                        timestepvalues[i+3] = float(self._scpi_properties[i+3].value.get())
                        timestepvalues[i+9] = float(self._scpi_properties[i+9].value.get())
                if self.usb_port.state == USBStates.CONNECTED:
                    for i in range(3):
                        timestepvalues[12+i] = float(self._scpi_properties[12].value[i].get())
                #timestepvalues[12] = timestepvalues[12]*(10**12) #from F to pF
                #timestepvalues[13] = timestepvalues[13]/(10**3) #changes from Ohm to kOhm
                #timestepvalues[15] = time.strftime("%H:%M:%S", time.localtime())
                timestepvalues[14] = (time.time() - init_time)
                self.data.append(timestepvalues)
                """ Plot Data """
                animate(self)
                    
        self.counter = 0
        self.counter += 1
        self.win.after(1, self.main_task)
        
        #print(time.time() - perfTime)

    def build_main_window(self):
        self.win = Tk()
        #self.win.geometry('1025x700')
        self.win.state('zoomed')
        self.win.wm_title("KARP")
        #self.win.geometry('%dx%d+%d+%d' % (1200, self.win.winfo_screenheight(), self.win.winfo_screenwidth()/2 - 1200/2, -10))
        
        ############################################
        ### INITIALIZING FUNCTIONS FOR LATER USE ###
        ############################################

        
        """ Sequence to start recording data, bound to Start Recording button"""
        def startrecord(event):
            global init_time
            init_time = time.time()
            self.recording = True
            self.indcombo.configure(state="disabled")
            self.depcombo.configure(state="disabled")
            self.ax.set_xlabel(self.indcombo.get())
            self.ax.set_ylabel(self.depcombo.get())
            #list of axes ranges, corresponding in order to the associated _scpi_property, being assigned based on selection
            lims = [[-1,2],[-20,120],[0,100],[-20,120],[-20,120],[-20,100],[-1,2],[-20,120],[0,100],[-20,120],[-20,120],[-20,100],[-20/(10**12),10/(10**12)],[-200*(10**3),100*(10**3)],[0,10]]
            self.ax.set_xlim(lims[self.indcombo.current()])
            self.ax.set_ylim(lims[self.depcombo.current()])
            #assigning which _scpi_property has been chosen for the independent/dependent variables for graphing
            #self.indvar = self._scpi_properties[self.indcombo.current()]
            #print(type(self.indvar))
            #self.depvar = self._scpi_properties[self.depcombo.current()]
            self.canvas.draw()
            recordbutton.configure(state="disabled",background="white")
            stoprecbutton.configure(state="normal",background="light grey")
            datalabels=["Output Relay 1","Target Voltage 1 (V)","Slew Rate 1 (V/s)","Output Voltage 1 (V)","Measured Voltage 1 (V)","Measured Current 1 (A)","Output Relay 2","Target Voltage 2 (V)","Slew Rate 2 (V/s)","Output Voltage 2 (V)","Measured Voltage 2 (V)","Measured Current 2 (A)","Primary Keysight Measurement","Secondary Keysight Measurement", "Time (s)"]
            with open('data_in_progress.csv','w',newline='') as file:
                    writer = csv.writer(file)
                    writer.writerow(datalabels)
            
        """ Sequence to stop recording data, bound to Stop Recording button"""
        def stoprecord(event):
            self.recording = False
            self.indcombo.configure(state="readonly")
            self.depcombo.configure(state="readonly")
            recordbutton.configure(state="normal",background="firebrick1")
            stoprecbutton.configure(state="disabled",background="white")
            savetime=time.strftime("%Y %m %d - %H_%M_%S")
            os.rename(r'data_in_progress.csv',savetime+'.csv')
            with open(savetime+'.csv','a',newline='') as file:
                writer = csv.writer(file)
                writer.writerows(self.data)
            self.fig.savefig(savetime+'.png')

        """ Facilitates serial port selection, links front-end (PortChooser) with back-end (connect())"""
        def choose_port_serial():
            p = PortChooser(self.win)
            if p.result is not None:
                self.serial_port.connect(p.result)
                if self.serial_port.state == SerialStates.CONNECTED:
                    connect_button1.config(state="disabled")
                    disconnect_button1.config(state="normal")
                    self.status_box1.config(text=self.serial_port.state.name)
                    self.status_box1.config(background='lime')
                    time.sleep(0.05)
                    self.serial_port.write(b'*IDN?\n')
                    resp = self.serial_port.read()
                    if resp is not None:
                        self.idn_box1.config(text=resp.strip())
                    for i in range(12):
                        self._scpi_properties[i].enable()
                        try:
                            self._scpi_properties[i].heldvalue.set(self._scpi_properties[i].value.get())
                        except: pass
        
        """ Facilitates USB selection, links front-end (PortChooser) with back-end (connect())"""
        def choose_port_usb():
            p = PortChooser(self.win)
            if p.result is not None:
                self.usb_port.connect(p.result)
                if self.usb_port.state == USBStates.CONNECTED:
                    connect_button2.config(state="disabled")
                    disconnect_button2.config(state="normal")
                    time.sleep(0.05)
                    self.status_box2.config(text=self.usb_port.state.name)
                    self.status_box2.config(background='lime')
                    self.idn_box2.config(text=rm.open_resource(p.result).query('*IDN?').strip())
                    for i in range(len(self._scpi_properties)-12):
                        self._scpi_properties[i+12].enable()

        """ Front-end command for disconnecting the RP100 plus a safe disconnect sequences"""
        def disconnect_serial():
            if self.recording:
                MsgBox = messagebox.askquestion("Stop Recording?","You are currently recording. Quitting now will stop recording and save data collected up to this point. Are you sure you'd like to continue?",icon="warning")
                if MsgBox == 'yes':
                    stoprecord()
                else: return
            waittime1 = int(np.ceil(float(self._scpi_properties[1].value.get())/float(self._scpi_properties[2].value.get())*1000))
            waittime2 = int(np.ceil(float(self._scpi_properties[1+6].value.get())/float(self._scpi_properties[2+6].value.get())*1000))
            waittime = max(waittime1,waittime2)
            MsgBox = messagebox.askquestion("Turn Off Power?","Would you like to set the output voltages to 0V before disconnecting? It will take about " + str(int(np.floor(waittime/1000))) + " seconds.",icon="question")
                # safe disconnect sequence, sets target voltage to 0, waits while ramping down, then sets output relay to disabled
            if MsgBox == 'yes':
                self._scpi_properties[1].value.set(0.00)
                self._scpi_properties[1+6].value.set(0.00)
                self._scpi_properties[1].scpi_set()
                self._scpi_properties[1+6].scpi_set()
                self.win.after(waittime)
                self._scpi_properties[0].value.set("0")
                self._scpi_properties[0].scpi_set()
                self._scpi_properties[0+6].value.set("0")
                self._scpi_properties[0+6].scpi_set()
            self.serial_port.disconnect()
            self.status_box1.configure(background='lightcoral')
            connect_button1.config(state="normal")
            disconnect_button1.config(state="disabled")
            self.status_box1.config(text=self.serial_port.state.name)
            self.idn_box1.config(text="None")
            for i in range(12):
                self._scpi_properties[i].disable()
                
        """ Front-end command for disconnecting the Keysight"""
        def disconnect_usb():
            self.usb_port.disconnect()
            connect_button2.config(state="normal")
            disconnect_button2.config(state="disabled")
            self.status_box2.config(text=self.usb_port.state.name)
            self.status_box2.configure(background='lightcoral')
            self.idn_box2.config(text="None")
            for i in range(len(self._scpi_properties)-12):
                self._scpi_properties[i+12].disable()
        
        """front end unlock/lock"""
        def unlocker():
            if self.serial_port.state == SerialStates.CONNECTED:
                for i in [0,1,2,6,7,8]:
                    self._scpi_properties[i].unlock()
                unlockbutton.config(state="disabled")
                lockbutton.config(state="normal")
        def locker():
            if self.serial_port.state == SerialStates.CONNECTED:
                for i in [0,1,2,6,7,8]:
                    self._scpi_properties[i].lock()
                unlockbutton.config(state="normal")
                lockbutton.config(state="disabled")
        
        def plotOn(event):
            self.isPlotOn = True
            plotButtonOff.config(state='normal')
            plotButtonOn.config(state='disabled')

        def plotOff(event):
            self.isPlotOn = False
            plotButtonOn.config(state='normal', background='white')
            plotButtonOff.config(state='disabled')
        
        """ Command to quit the program and prompt to stop recording if recording is active"""
        def quitexe():
            MsgBox = messagebox.askquestion("Quit Application","You are about to quit the application. Would you like to proceeed?",icon="warning")
            if MsgBox == 'yes':
                if self.serial_port.state == SerialStates.CONNECTED or self.usb_port.state == USBStates.CONNECTED:
                    if self.recording:
                        MsgBox = messagebox.askquestion("Stop Recording?","You are currently recording. Quitting now will stop recording and save data collected up to this point. Are you sure you'd like to continue?",icon="warning")
                        if MsgBox == 'yes':
                            stoprecord()
                            disconnect_serial()
                        else: return
                self.win.destroy()
            return
        
        
        #############################################################
        ######################## GENERATE GUI #######################
        #############################################################
        
        
        """ Partitions the program into two tabs """
        tabControl = ttk.Notebook(self.win)
        tab1 = ttk.Frame(tabControl)
        tab2 = ttk.Frame(tabControl)
        tabControl.add(tab1, text ='Instrumentation Measurement & Control')
        tabControl.add(tab2, text ='User Guide + Error Reporting')
        tabControl.pack(expand = 1, fill ="both")
        tab1.columnconfigure(index=1,weight=1)
        tab1.columnconfigure(index=2,weight=1)
        self.win.iconbitmap('LAQM.ico')


        #############################################################
        ### GENERATE TAB 1: INSTRUMENTATION MEASUREMENT & CONTROL ###
        #############################################################

        """ Generates the title """
        frame = Frame(tab1, border=2, relief=GROOVE)
        frame.grid(row=0, column=0, columnspan=3, sticky='WE', padx=10, pady=5)
        frame.icon1 = PhotoImage(file='LAQM.png')
        #frame.icon2 = PhotoImage(file='Rutgers.png')
        Label(frame, image=frame.icon1).grid(row=1, rowspan=2, column=0)
        #Label(frame, image=frame.icon2).grid(row=1, rowspan=2, column=1)
        label = Label(frame, text="  KARP: Keysight and Razorbill Product Control Tool", justify='right', font='Helvetica 26 bold')
        label.grid(row=2, column=2, columnspan=10, pady=10, sticky="NSEW")
        
        """ Generates the RP100 Serial Connection Box """
        frame = Frame(tab1, width=100, height = 125, border=2, relief=GROOVE)
        frame.grid_propagate(False)
        frame.grid(row=1, column=1, sticky="WE", padx=10, pady=5)
        frame.grid_columnconfigure(3, weight=2)
        Label(frame, text="Razorbill RP100 Serial Connection").grid(row=0, column=1, columnspan=3)
        quit_button = Button(frame,text='Quit',command=quitexe)
        quit_button.grid(row=0,column=1)
        connect_button1 = Button(frame, text="Connect", command=choose_port_serial)
        connect_button1.grid(row=1, column=1, sticky="NSEW")
        disconnect_button1 = Button(frame, text="Disconnect", command=disconnect_serial, state="disabled")
        disconnect_button1.grid(row=2, column=1, sticky="NSEW")
        unlockbutton= Button(frame, text="Unlock", command=unlocker, state="disabled")
        unlockbutton.grid(row=3, column=1, sticky="NEWS", pady=15)
        lockbutton = Button(frame, text="Lock", command=locker, state="normal")
        lockbutton.grid(row=3, column=2, sticky="NWS", pady=15, ipadx=15)
        Label(frame, text="Connection Status:").grid(row=1, column=2, sticky="E", padx=(30, 1))
        Label(frame, text="Connected To (*IDN?):").grid(row=2, column=2, sticky="E", padx=(30, 1))
        self.status_box1 = Label(frame, text="UNCONFIGURED", relief=SUNKEN, background='lightcoral')
        self.status_box1.grid(row=1, column=3, sticky="WE")
        self.idn_box1 = Label(frame, text="None", relief=SUNKEN)
        self.idn_box1.grid(row=2, column=3, sticky="WE")
        
        """ Generates the Keysight E4980AL USB Connection Box """
        frame = Frame(tab1, width = 150, height = 90,  border=2, relief=GROOVE)
        frame.grid_propagate(False)
        frame.grid(row=1, column=2, sticky="NWE", padx=10, pady=5)
        frame.grid_columnconfigure(3, weight=2)
        Label(frame, text="Keysight E4980AL USB Connection").grid(row=0, column=1, columnspan=3)
        quit_button = Button(frame,text='Quit',command=quitexe)
        quit_button.grid(row=0,column=1)
        connect_button2 = Button(frame, text="Connect", command=choose_port_usb)
        connect_button2.grid(row=1, column=1, sticky="NSEW")
        disconnect_button2 = Button(frame, text="Disconnect", command=disconnect_usb, state="disabled")
        disconnect_button2.grid(row=2, column=1, sticky="NSEW")
        
        Label(frame, text="Connection Status:").grid(row=1, column=2, sticky="E", padx=(30, 1), ipadx=0, ipady=0)
        Label(frame, text="Connected To (*IDN?):").grid(row=2, column=2, sticky="E", padx=(30, 1), ipadx=0, ipady=0)
        self.status_box2 = Label(frame, text="UNCONFIGURED", relief=SUNKEN, background='lightcoral')
        self.status_box2.grid(row=1, column=3, sticky="WE")
        self.idn_box2 = Label(frame, text="None", relief=SUNKEN)
        self.idn_box2.grid(row=2, column=3, sticky="WE")

        """ Generates the RP100 control widgets for each channel """
        
        for channel in (1, 2):
            frame = Frame(tab1, width=150, height = 205, border=2, relief=GROOVE)
            frame.grid_propagate(False)
            frame.grid(row=2+channel, padx=10, pady=5, column=1, sticky='NEW')
            frame.columnconfigure(index=1,weight=1)
            frame.columnconfigure(index=2,weight=2)
            Label(frame, text="RP100 Channel " + str(channel)).grid(row=1, column=0, columnspan=5)
            ch_bytes = bytes(str(channel), 'utf8')
            prop = ScpiPropertyBool(frame, 2, self.serial_port, b"OUTP" + ch_bytes, "Output relay " + str(channel))
            self._scpi_properties.append(prop)
            prop = ScpiPropertyFloat(frame, 3, self.serial_port, b"SOUR" + ch_bytes + b":VOLT", "Target Voltage " + str(channel) + " (V)")
            self._scpi_properties.append(prop)
            prop = ScpiPropertyFloat(frame, 4, self.serial_port, b"SOUR" + ch_bytes + b":VOLT:SLEW", "Slew Rate " + str(channel) + " (V/s)")
            self._scpi_properties.append(prop)
            prop = ScpiPropertyFloat(frame, 5, self.serial_port, b"SOUR" + ch_bytes + b":VOLT:NOW", "Output Voltage " + str(channel) + " (V)", can_set=False)
            self._scpi_properties.append(prop)
            prop = ScpiPropertyFloat(frame, 6, self.serial_port, b"MEAS" + ch_bytes + b":VOLT", "Measured Voltage " + str(channel) + " (V)", can_set=False)
            self._scpi_properties.append(prop)
            prop = ScpiPropertyFloat(frame, 7, self.serial_port, b"MEAS" + ch_bytes + b":CURR", "Measured Current " + str(channel) + " (A)", can_set=False)
            self._scpi_properties.append(prop)
        
        """ Generates the Keysight control widgets """
        frame = Frame(tab1, width=120, height=100, border=2, relief=GROOVE)
        frame.grid_propagate(False)
        frame.grid(row=5, padx=10, pady=5, column=1, sticky=NSEW)
        Label(frame, text="E4980AL Channel").grid(row=1, column=1, columnspan=5)
        frame.columnconfigure(index=1,weight=1)
        frame.columnconfigure(index=2,weight=1)
        #frame.grid_forget()
        prop = ScpiPropertyFloat(frame, 2, self.usb_port, ":FETCh:IMPedance:FORMatted?", ["Capacitance (F)","Resistance (Ω)"], can_set=False)
        self._scpi_properties.append(prop)
        
        """ Generate the Live Plotting graph """
        plotframe = Frame(tab1, border=2, relief=GROOVE)
        plotframe.grid(row=3,column=2,columnspan=1,rowspan=2, padx=50, pady=5, sticky="NSEW")
        self.fig = plt.Figure(tight_layout=True)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_xlim([-20,120])
        self.ax.set_ylim([-1,1])
        self.xval=[0]
        self.yval=[0]
        self.scattery = self.ax.scatter(self.xval,self.yval,color='red')
        self.ax.grid()
        self.ax.axvline(x=0,color='black')
        self.ax.axhline(y=0,color='black')
        self.canvas = FigureCanvasTkAgg(self.fig, master=plotframe)
        self.canvas.get_tk_widget().pack(fill=tkinter.BOTH, expand=1)
        
        """ Generates the Plotting/Recording Control panel (bottom right) """
        frame = Frame(tab1, border=2, relief=GROOVE)
        frame.grid(row=5,column=2)
        values_list = []
        for prop in self._scpi_properties:
            if type(prop.description) == list: #keysight returns measurements in [x,y] list
                for i in range(len(prop.description)):
                    values_list.append(prop.description[i])
            else:
                values_list.append(prop.description)
        values_list.append("Time")
        
        """
        plotButtonOn = Button(frame, text="On", background='lime')
        plotButtonOn.grid(row=1, column=2)
        plotButtonOn.bind('<ButtonRelease-1>', plotOn)
        plotButtonOff = Button(frame, text="Off", background='red')
        plotButtonOff.grid(row=2, column=2)
        plotButtonOff.bind('<ButtonRelease-1>', plotOff) """
        
        
        
        
        self.indcombo = ttk.Combobox(frame)
        self.indcombo.grid(row=1,column=0)
        value = StringVar()
        self.indcombo.configure(state="readonly",textvariable = value, values=values_list)
        
        label = Label(frame, text="Independent variable: ")
        label.grid(row=0,column=0)
        label = Label(frame, text="Dependent variable: ")
        label.grid(row=2,column=0)
        self.depcombo = ttk.Combobox(frame)
        self.depcombo.grid(row=3,column=0)
        value = StringVar()
        self.depcombo.configure(state="readonly",textvariable = value, values=values_list)
        recordbutton = Button(frame, text="Start Record", background="firebrick1")
        recordbutton.grid(row=0,column=1,rowspan=2,padx=10)
        recordbutton.bind("<ButtonRelease-1>",startrecord)
        recordbutton.configure(disabledforeground="#a3a3a3")
        stoprecbutton = Button(frame, text="Stop & Save", background="white")
        stoprecbutton.grid(row=2,column=1,rowspan=2,padx=10)
        stoprecbutton.bind("<ButtonRelease-1>",stoprecord)
        stoprecbutton.configure(state="disabled")
        
        
        #############################################################
        ####### GENERATE TAB 2: USER GUIDE + ERROR REPORTING ########
        #############################################################
        
        """ Generate User Guide """
        s = ("Get started by clicking "
             "'connect' and choosing the serial port corresponding to the proper instrument. Below are the instruent "
             "control options. To send a command to the instrument, toggle the radiobuttons by double clicking, or type "
             "in the box and hit Enter. The values being read will update below in realtime. The graph to the right "
             "displays live data according to the independent/dependent variables chosen below. Click the red 'Start Record' "
             "button to begin recording data, and the grey 'Stop Record' button to save to a CSV file. To quit or "
             "disconnect, click the respective buttons in the serial connection box(es). ")
        frame = Frame(tab2, border=2, relief=GROOVE)
        frame.grid(row=1, column=1, columnspan=2, padx=10, pady=5, sticky="WE")
        frame.grid_columnconfigure(2, weight=2)
        label = Label(frame, text="How To Use:", justify='center', font='Helvetica 16')
        label.grid(row=0,column=0)
        label = Label(frame, text=s, wraplength=1000, justify="left", anchor="w")
        label.grid(row=1, column=0, sticky="NWE")
        
        """ Generate the Error Reporter Printer """
        frame = Frame(tab2, border=2, relief=GROOVE)
        frame.grid(row=10, column=1, columnspan=2, padx=10, pady=5, sticky="WE")
        frame.grid_columnconfigure(2, weight=2)
        prop = ScpiErrorReporter(frame, 1, self.serial_port)
        self._scpi_properties.append(prop)
        frame = Frame(tab2)
        frame.grid(row=20, column=1, columnspan=2)
        self.log_text = Text(frame, borderwidth=3, relief="sunken")
        self.log_text.config(font=("consolas", 10), undo=True, wrap='word')
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        scroll = Scrollbar(frame, command=self.log_text.yview)
        scroll.grid(row=0, column=1, sticky='nsew')
        self.log_text['yscrollcommand'] = scroll.set
        self.log_text.config(state="disabled")
        
        """ Initialize all widgets as disabled """
        for prop in self._scpi_properties:
            prop.disable()
        
        
class PortChooser(tkinter.simpledialog.Dialog):
    """ A popup dialog for selecting a serial port to connect to. """
    def body(self, master):
        self.usb_ports = []
        for i in range(len(rm.list_resources())):
            if str(rm.list_resources()[i])[0:3] == "USB":
                self.usb_ports.append(rm.list_resources()[i])
        self.iconbitmap('LAQM.ico')
        self.ports = list_ports.comports()
        self.choice = StringVar(master)
        self.choice.set("None")
        Label(master, text="Please choose a serial port to connect to.").grid(row=1, column=1, columnspan=2)
        n = 10
        if len(self.ports) > 0 or len(self.usb_ports) > 0:
            for port in self.ports:
                Radiobutton(master, text=port.description, variable=self.choice, value=port.device)\
                    .grid(row=n, column=1, columnspan=2, sticky=W)
                n += 1
            for port in self.usb_ports:
                openedport = rm.open_resource(port)
                try:
                    Radiobutton(master, text=openedport.resource_info.alias, variable=self.choice, value=port).grid(row=n, column=1, columnspan=2, sticky=W)
                    n += 1
                except:
                    Radiobutton(master, text=openedport.resource_info.resource_name, variable=self.choice, value=port).grid(row=n, column=1, columnspan=2, sticky=W)
                    n += 1
        else:
            Label(master, text="Error: Could not find a suitable serial or USB port.\n"
                               "Make sure your devices are plugged in and turned on,\n"
                               "and the E4980AL is properly configured (see guide). ").grid(row=n, column=1, columnspan=2)

    """ Find the device associated with the selection and return it as the selected port"""
    def apply(self):
        choice = self.choice.get()
        for port in self.ports:
            if choice == port.device:
                self.result = port
                return
        for port in self.usb_ports:
            if choice == port:
                self.result = port
                return
        self.result = None


if __name__ == "__main__":
    MainGui()