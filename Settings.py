from tkinter import *
from tkinter import messagebox
from Global import grabobjs
from Global import CryptHandle
import os

curr_dir = os.path.dirname(os.path.abspath(__file__))
main_dir = os.path.dirname(curr_dir)
global_objs = grabobjs(main_dir)


class SettingsGUI:
    def __init__(self, header=None):
        if header:
            self.header_text = header
        else:
            self.header_text = 'Welcome to Vacuum Settings!\nSettings can be changed below.\nPress save when finished'

        self.asql = global_objs['SQL']
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

    @staticmethod
    def fill_textbox(setting_list, val, key):
        assert(key and val and setting_list)
        item = global_objs[setting_list].grab_item(key)

        if isinstance(item, CryptHandle):
            val.set(item.decrypt_text())

    @staticmethod
    def add_setting(setting_list, val, key):
        assert(key and val and setting_list)

        global_objs[setting_list].del_item(key)
        global_objs[setting_list].add_item(key=key, val=val, encrypt=True)

    def check_table(self, table):
        table2 = table.split('.')

        if len(table2) == 2:
            myresults = self.asql.query('''
                SELECT
                    1
                FROM information_schema.tables
                WHERE
                    Table_Schema = '{0}'
                        AND
                    Table_Name = '{1}'
            '''.format(table2[0], table2[1]))

            if myresults.empty:
                return False
            else:
                return True
        else:
            return False

    def build_gui(self):
        # Set GUI Geometry and GUI Title
        self.main.geometry('444x285+500+150')
        self.main.title('Vacuum Settings')
        self.main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.main)
        network_frame = LabelFrame(self.main, text='Network Settings', width=444, height=70)
        wrkbook_main_frame = LabelFrame(self.main, text='SQL Worksheet Staging Tables', width=444, height=54)
        wrkbook_left_frame = Frame(wrkbook_main_frame, width=219, height=54)
        wrkbook_right_frame = Frame(wrkbook_main_frame, width=219, height=54)
        error_frame = LabelFrame(self.main, text='SQL Workbook Error Tables', width=444, height=60)
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
        server_label = Label(self.main, text='Server:', padx=15, pady=7)
        server_txtbox = Entry(self.main, textvariable=self.server)
        server_label.pack(in_=network_frame, side=LEFT)
        server_txtbox.pack(in_=network_frame, side=LEFT)

        #     Server Database Input Box
        database_label = Label(self.main, text='Database:')
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
        save_settings_button.grid(row=0, column=0, pady=6, padx=10)

        #       Cancel Button
        cancel_button = Button(buttons_frame, text='Cancel', width=20, command=self.cancel)
        cancel_button.grid(row=0, column=1, pady=6, padx=115)

        # Fill Textboxes with settings
        self.fill_gui()

        # Show GUI Dialog
        self.main.mainloop()

    def fill_gui(self):
        self.fill_textbox('Settings', self.server, 'Server')
        self.fill_textbox('Settings', self.database, 'Database')
        self.fill_textbox('Local_Settings', self.w1s, 'W1S_TBL')
        self.fill_textbox('Local_Settings', self.w2s, 'W2S_TBL')
        self.fill_textbox('Local_Settings', self.w3s, 'W3S_TBL')
        self.fill_textbox('Local_Settings', self.w4s, 'W4S_TBL')
        self.fill_textbox('Local_Settings', self.we, 'WE_TBL')
        self.fill_textbox('Local_Settings', self.wne, 'WNE_TBL')

    def save_settings(self):
        if not self.w1s.get():
            messagebox.showerror('W1S Empty Error!', 'No value has been inputed for W1S TBL (Worksheet One Staging)',
                                 parent=self.main)
        elif not self.w2s.get():
            messagebox.showerror('W2S Empty Error!', 'No value has been inputed for W2S TBL (Worksheet Two Staging)',
                                 parent=self.main)
        elif not self.w3s.get():
            messagebox.showerror('W3S Empty Error!', 'No value has been inputed for W3S TBL (Worksheet Three Staging)',
                                 parent=self.main)
        elif not self.w4s.get():
            messagebox.showerror('W4S Empty Error!', 'No value has been inputed for W4S TBL (Worksheet Four Staging)',
                                 parent=self.main)
        elif not self.server.get():
            messagebox.showerror('Server Empty Error!', 'No value has been inputed for Server',
                                 parent=self.main)
        elif not self.database.get():
            messagebox.showerror('Database Empty Error!', 'No value has been inputed for Database',
                                 parent=self.main)
        elif not self.we.get():
            messagebox.showerror('WE Empty Error!', 'No value has been inputed for WE TBL (Workbook Errors)',
                                 parent=self.main)
        elif not self.wne.get():
            messagebox.showerror('WNE Empty Error!', 'No value has been inputed for WNE TBL (Workbook Norm Errors)',
                                 parent=self.main)
        else:
            self.add_setting('Settings', self.server.get(), 'Server')
            self.add_setting('Settings', self.database.get(), 'Database')
            self.asql.change_config(server=self.server.get(), database=self.database.get())

            if self.asql.test_conn('alch'):
                self.asql.connect('alch')

                if not self.check_table(self.w1s.get()):
                    messagebox.showerror('Invalid W1S Table!',
                                         'W1S, Worksheet One Staging, table does not exist in sql server',
                                         parent=self.main)
                elif not self.check_table(self.w2s.get()):
                    messagebox.showerror('Invalid W2S Table!',
                                         'W2S, Worksheet Two Staging, table does not exist in sql server',
                                         parent=self.main)
                elif not self.check_table(self.w3s.get()):
                    messagebox.showerror('Invalid W3S Table!',
                                         'W3S, Worksheet Three Staging, table does not exist in sql server',
                                         parent=self.main)
                elif not self.check_table(self.w4s.get()):
                    messagebox.showerror('Invalid W4S Table!',
                                         'W4S, Worksheet Four Staging, table does not exist in sql server',
                                         parent=self.main)
                elif not self.check_table(self.we.get()):
                    messagebox.showerror('Invalid WE Table!',
                                         'WE, Workbook Errors, table does not exist in sql server',
                                         parent=self.main)
                elif not self.check_table(self.wne.get()):
                    messagebox.showerror('Invalid WNE Table!',
                                         'WNE, Workbook Norm Errors, table does not exist in sql server',
                                         parent=self.main)
                else:
                    self.add_setting('Local_Settings', self.w1s.get(), 'W1S_TBL')
                    self.add_setting('Local_Settings', self.w2s.get(), 'W2S_TBL')
                    self.add_setting('Local_Settings', self.w3s.get(), 'W3S_TBL')
                    self.add_setting('Local_Settings', self.w4s.get(), 'W4S_TBL')
                    self.add_setting('Local_Settings', self.we.get(), 'WE_TBL')
                    self.add_setting('Local_Settings', self.wne.get(), 'WNE_TBL')

                    self.main.destroy()
            else:
                messagebox.showerror('Network Test Error!', 'Unable to connect to {0} server and {1} database',
                                     parent=self.main)

    def cancel(self):
        self.main.destroy()


if __name__ == '__main__':
    header_text = 'Welcome to Vacuum Settings!\nSettings haven''t been established.\nPlease fill out all empty fields below:'
    obj = SettingsGUI()
    obj.build_gui()
