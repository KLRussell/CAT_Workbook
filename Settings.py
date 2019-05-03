from tkinter import *
from tkinter import messagebox
from Global import grabobjs
import os

curr_dir = os.path.dirname(os.path.abspath(__file__))
main_dir = os.path.dirname(curr_dir)
Global_Objs = grabobjs(main_dir)


class SettingsGUI:
    def __init__(self, header=None):
        if header:
            self.header_text = header
        else:
            self.header_text = 'Welcome to Vacuum Settings!\nSettings can be changed below.\nPress save when finished'

        self.asql = Global_Objs['SQL']
        self.asql.connect('alch')
        self.main = Tk()

        # GUI Variables
        self.server = StringVar()
        self.database = StringVar()
        self.w1s = StringVar()
        self.w2s = StringVar()
        self.w3s = StringVar()
        self.w4s = StringVar()
        self.we = StringVar()
        self.wne = StringVar()

    def build_gui(self):
        # Set GUI Geometry and GUI Title
        self.main.geometry('480x400+500+150')
        self.main.title('Vacuum Settings')
        self.main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.main)
        network_frame = LabelFrame(self.main, text='Network Settings', width=480, height=70)
        wrkbook_main_frame = LabelFrame(self.main, text='SQL Worksheet Staging Tables', width=480, height=54)
        wrkbook_left_frame = Frame(wrkbook_main_frame, width=238, height=54)
        wrkbook_right_frame = Frame(wrkbook_main_frame, width=238, height=54)
        error_frame = LabelFrame(self.main, text='SQL Workbook Error Tables', width=508, height=60)
        buttons_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        network_frame.pack(fill="both")
        wrkbook_main_frame.pack(fill="both")
        wrkbook_left_frame.grid(row=0, column=0, ipady=5)
        wrkbook_right_frame.grid(row=0, column=1, ipady=5)
        error_frame.pack(fill="both")
        buttons_frame.pack(fill='both')

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.main, text=self.header_text, width=375, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply Network Labels & Input boxes to the Network_Frame
        #     SQL Server Input Box
        server_label = Label(self.main, text='SQL Server:', padx=15, pady=7)
        server_txtbox = Entry(self.main, textvariable=self.server)
        server_label.pack(in_=network_frame, side=LEFT)
        server_txtbox.pack(in_=network_frame, side=LEFT)

        #     Server Database Input Box
        database_label = Label(self.main, text='Server Database:')
        database_txtbox = Entry(self.main, textvariable=self.database)
        database_txtbox.pack(in_=network_frame, side=RIGHT, pady=7, padx=15)
        database_label.pack(in_=network_frame, side=RIGHT)

        # Apply Labels & Input boxes to the Wrkbook_Left_Frame
        #     Worksheet One Table Input Box
        w1s_label = Label(wrkbook_left_frame, text='W1S TBL:')
        w1s_txtbox = Entry(wrkbook_left_frame, textvariable=self.w1s)
        w1s_label.grid(row=0, column=0, padx=8, pady=5, sticky='w')
        w1s_txtbox.grid(row=0, column=1, padx=13, pady=5, sticky='e')

        #     Worksheet Two Table Input Box
        w2s_label = Label(wrkbook_left_frame, text='W2S TBL:')
        w2s_txtbox = Entry(wrkbook_left_frame, textvariable=self.w2s)
        w2s_label.grid(row=1, column=0, padx=8, pady=5, sticky='w')
        w2s_txtbox.grid(row=1, column=1, padx=13, pady=5, sticky='e')

        # Apply Labels & Input boxes to the Wrkbook_Right_Frame
        #     Worksheet Three Table Input Box
        w3s_label = Label(wrkbook_right_frame, text='W3S TBL:')
        w3s_txtbox = Entry(wrkbook_right_frame, textvariable=self.w3s)
        w3s_label.grid(row=0, column=0, padx=8, pady=5, sticky='w')
        w3s_txtbox.grid(row=0, column=1, padx=13, pady=5, sticky='e')

        #     Worksheet Four Table Input Box
        w4s_label = Label(wrkbook_right_frame, text='W4S TBL:')
        w4s_txtbox = Entry(wrkbook_right_frame, textvariable=self.w4s)
        w4s_label.grid(row=1, column=0, padx=8, pady=5, sticky='w')
        w4s_txtbox.grid(row=1, column=1, padx=13, pady=5, sticky='e')

        # Apply Labels & Input boxes to the Wrkbook_Right_Frame
        #     Workbook Errors Table Input Box
        we_label = Label(error_frame, text='WE TBL:')
        we_txtbox = Entry(error_frame, textvariable=self.we)
        we_label.grid(row=0, column=0, padx=8, pady=5, sticky='w')
        we_txtbox.grid(row=0, column=1, padx=13, pady=5, sticky='e')

        #     Workbook Norm Errors Table Input Box
        wne_label = Label(error_frame, text='WNE TBL:')
        wne_txtbox = Entry(error_frame, textvariable=self.wne)
        wne_label.grid(row=0, column=2, padx=8, pady=5, sticky='w')
        wne_txtbox.grid(row=0, column=3, padx=13, pady=5, sticky='e')

        # Apply buttons to the Buttons_Frame
        #       Save Button
        save_settings_button = Button(buttons_frame, text='Save Settings', width=20, command=self.save_settings)
        save_settings_button.grid(row=0, column=0, pady=6, padx=15)

        #       Cancel Button
        cancel_button = Button(buttons_frame, text='Cancel', width=20, command=self.cancel)
        cancel_button.grid(row=0, column=1, pady=6, padx=165)

        # Show GUI Dialog
        self.main.mainloop()

    def save_settings(self):
        print('saving settings')

    def cancel(self):
        self.main.destroy()


if __name__ == '__main__':
    header_text = 'Welcome to Vacuum Settings!\nSettings haven''t been established.\nPlease fill out all empty fields below:'
    obj = SettingsGUI()
    obj.build_gui()