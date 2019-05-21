# Global Module import
from Global import grabobjs
from Global import XMLParseClass
from Global import XMLAppendClass
from Settings import SettingsGUI
from time import sleep

import os
import gc
import pathlib as pl
import random
import numpy as np
import pandas as pd
import traceback

# Global Variable declaration
curr_dir = os.path.dirname(os.path.abspath(__file__))
main_dir = os.path.dirname(curr_dir)
process_dir = os.path.join(main_dir, '02_To_Process')
uploaded_dir = os.path.join(main_dir, '03_Processed', '01_Uploaded')
failed_dir = os.path.join(main_dir, '03_Processed', '02_Failed')
errors_dir = os.path.join(main_dir, '04_Errors')
global_objs = grabobjs(main_dir, 'Vacuum')


# CATWorkbook Class Function
class CATWorkbook:
    # Function that is executed upon creation of CATWorkbook class
    def __init__(self, df, file):
        df['Source_File'] = file
        df['Source_File_ID'] = np.arange(len(df))
        self.sheet_name = df['Sheet_Name'][0]
        del df['Sheet_Name']

        # Opens SQL Connection for class object
        self.asql = global_objs['SQL']
        try:
            self.asql.connect('alch')
            self.df = df
            self.file = file
        finally:
            self.asql.close()

    # Closes SQL connection for class object
    def close_conn(self):
        self.asql.close()

    # Cleans Dataframe of double spaces and leading/preceeding spaces
    def clean_df(self):
        for col in self.df.columns.tolist():
            self.df[col] = self.df[col].astype(str).str.strip().str.replace('  ', ' ')

    # Finds CSR in CSR directory for items sent to LV
    def lv_operations(self):
        if self.sheet_name == 'Sheet1':
            myitems = self.df[self.df['Action'] == 'Send to LV']

            if not myitems.empty:
                self.df['CSR_File_Name'] = None

                for index, row in self.df.head().iterrows():
                    f = list(pl.Path(global_objs['Local_Settings'].grab_item('CSR_Dir').decrypt_text())
                             .glob('{0}_{1}*'.format(row['Source_TBL'], row['Source_ID'])))
                    if f:
                        self.df.loc[self.df.index == index, 'CSR_File_Name'] = str(
                            os.path.basename(max(f, key=os.path.getctime)))

    # Function to upload data into SQL server tables according to CAT Workbook Sheet name
    def upload(self):
        try:
            if self.sheet_name == 'Sheet1':
                table = global_objs['Local_Settings'].grab_item('W1S_TBL').decrypt_text()
            elif self.sheet_name == 'Sheet2':
                table = global_objs['Local_Settings'].grab_item('W2S_TBL').decrypt_text()
            elif self.sheet_name == 'Sheet3':
                table = global_objs['Local_Settings'].grab_item('W3S_TBL').decrypt_text()
            elif self.sheet_name == 'Sheet4':
                table = global_objs['Local_Settings'].grab_item('W4S_TBL').decrypt_text()
            else:
                global_objs['Event_Log'].write_log(
                    'Unable to upload {0}. Spreadsheet has invalid Sheet_Name ''{1}'''.format(self.file,
                                                                                              self.sheet_name))
                return False

            global_objs['Event_Log'].write_log('Uploading {0} items to {1} SQL table'.format(len(self.df), table))
            return self.asql.upload(self.df, table, index=False)
        except:
            return False

    # Migrate files from 02_To_Process to 03_Processed
    #   (Files may be either 01_Uploaded or 02_Failed according to the success/failure of the upload process)
    def migrate_file(self, processed=True):
        if processed:
            global_objs['Event_Log'].write_log('Upload successful. Migrating file to 03_Processed\\01_Uploaded folder')
            os.rename(os.path.join(process_dir, self.file), os.path.join(uploaded_dir, self.file))
        else:
            global_objs['Event_Log'].write_log(
                'Failed Upload! Migrating file to 03_Processed\\02_Failed folder')
            os.rename(os.path.join(process_dir, self.file), os.path.join(failed_dir, self.file))


# Class object for ErrorProcessing
class ErrorProcessing:
    # Private variable declaring for ErrorProcessing class
    df = pd.DataFrame()

    # Function that is executed upon creation of ErrorProcessing class
    def __init__(self):
        self.asql = global_objs['SQL']
        try:
            self.asql.connect('alch')
        finally:
            self.asql.close()

    # Closes SQL connection for class object
    def close_conn(self):
        self.asql.close()

    # Checks for errors in the error tables within SQL Server
    def check_errors(self):
        self.df = self.asql.query('''
            SELECT
                WE.Source_File,
                WE.Source_File_ID,
                WE.Error_Col,
                WE.Error_Msg
                
            FROM {0} As WE
        '''.format(global_objs['Local_Settings'].grab_item('WE_TBL').decrypt_text()))

        if self.df.empty:
            return False
        else:
            global_objs['Event_Log'].write_log('Vacuum clogged with Errors. Cleaning vacuum...')
            return True

    # Export errors into the 04_Errors folder if errors exists in the Error tables in SQL Server
    def process_errors(self):
        try:
            for source_file in self.df['Source_File'].unique().tolist():
                xmlobj = XMLParseClass(os.path.join(uploaded_dir, source_file))
                df = xmlobj.parsexml('./{urn:schemas-microsoft-com:rowset}data/')
                df['Source_File_ID'] = np.arange(len(df))
                df = pd.merge(df, self.df, on='Source_File_ID')
                del df['Source_File'], df['Source_File_ID']

                for comp_serial in df['Comp_Serial'].tolist():
                    global_objs['Event_Log'].write_log('{0} error(s) appended to {1}\'s workbook'.format(
                        len(df[df['Comp_Serial'] == comp_serial].index), comp_serial), 'warning')

                    if not os.path.exists(os.path.join(errors_dir, comp_serial)):
                        os.makedirs(os.path.join(errors_dir, comp_serial))

                    obj = XMLAppendClass(os.path.join(errors_dir, comp_serial, source_file))
                    obj.write_xml(df.loc[df['Comp_Serial'] == comp_serial])
                    del obj

                del xmlobj, df
            return True
        except:
            return False

    # Truncate error table in SQL Server after errors are processed
    def truncate_tbl(self):
        self.asql.execute('TRUNCATE TABLE %s' % global_objs['Local_Settings'].grab_item('WE_TBL').decrypt_text())


# Generate random talk whenever Vacuum is idled (Text will not show in event logs)
def gen_talk():
    f = open(os.path.join(curr_dir, 'Vacuum_Talk.txt'), 'r')
    lines = f.readlines()
    talkid = random.randint(0, len(lines) - 1)

    print('The Vacuum: {0}'.format(lines[talkid]))
    f.close()


# Function to read text from a .sql file and execute that code in SQL Server
def newuser(file):
    f = open(file, "r")

    try:
        data = f.read()

        if len(data) > 0:
            asql = global_objs['SQL']

            try:
                asql.connect('alch')
                asql.execute(data)
            finally:
                asql.close()
    finally:
        f.close()


# Function to check whether updates exists in the 02_To_Process
def find_updates():
    f = list(pl.Path(process_dir).glob('*.xml')) + list(pl.Path(process_dir).glob('*.sql'))

    if f:
        global_objs['Event_Log'].write_log('Vacuum found crumbs... Switching to max suction')
        return f


# Function to process updates in the 02_To_Process folder
def proc_updates(files):
    write_blank = False

    for file in files:
        folder_name = os.path.basename(os.path.dirname(file))
        file_name = os.path.basename(file)

        if write_blank:
            global_objs['Event_Log'].write_log('')

        global_objs['Event_Log'].write_log('Reading file ({0}/{1})'.format(folder_name, file_name))

        if os.path.splitext(file_name)[1].lower() == '.xml':
            xmlobj = XMLParseClass(file)

            if xmlobj:
                df = xmlobj.parsexml('./{urn:schemas-microsoft-com:rowset}data/')

                if not df.empty:
                    obj = CATWorkbook(df, file_name)

                    try:
                        obj.clean_df()
                        obj.lv_operations()

                        if obj.upload():
                            obj.migrate_file()
                        else:
                            obj.migrate_file(False)
                    finally:
                        obj.close_conn()
                        del obj
                del df
            del xmlobj
        else:
            global_objs['Event_Log'].write_log('Adding new user to CAT Employee table')
            newuser(file)
            os.remove(file)

        write_blank = True
        del folder_name, file_name


# Function to process Errors if errors exists in the SQL Server Error table
def proc_errors():
    obj = ErrorProcessing()
    try:
        if obj.check_errors() and obj.process_errors():
            obj.truncate_tbl()
    finally:
        obj.close_conn()


# Checks network and sql server table settings to see if it exists and whether those settings are valid
#   A GUI will pop-up if settings need to be inputed or whether settings are invalid
def check_settings():
    my_return = False
    obj = SettingsGUI()

    if not os.path.exists(errors_dir):
        os.makedirs(errors_dir)

    if not os.path.exists(process_dir):
        os.makedirs(process_dir)

    if not os.path.exists(uploaded_dir):
        os.makedirs(uploaded_dir)

    if not os.path.exists(failed_dir):
        os.makedirs(failed_dir)

    if not global_objs['Settings'].grab_item('Server')\
            or not global_objs['Settings'].grab_item('Database') \
            or not global_objs['Local_Settings'].grab_item('CSR_Dir') \
            or not global_objs['Local_Settings'].grab_item('W1S_TBL')\
            or not global_objs['Local_Settings'].grab_item('W2S_TBL')\
            or not global_objs['Local_Settings'].grab_item('W3S_TBL')\
            or not global_objs['Local_Settings'].grab_item('WE_TBL'):
        header_text = 'Welcome to Vacuum Settings!\nSettings haven''t been established.\nPlease fill out all empty fields below:'
        obj.build_gui(header_text)
    else:
        mylist = []

        try:
            if not obj.sql_connect():
                mylist.append('network')
            if not os.path.exists(global_objs['Local_Settings'].grab_item('CSR_Dir').decrypt_text()):
                mylist.append('CSR Dir')
            if not obj.check_table(global_objs['Local_Settings'].grab_item('W1S_TBL').decrypt_text()):
                mylist.append('W1S')
            if not obj.check_table(global_objs['Local_Settings'].grab_item('W2S_TBL').decrypt_text()):
                mylist.append('W2S')
            if not obj.check_table(global_objs['Local_Settings'].grab_item('W3S_TBL').decrypt_text()):
                mylist.append('W3S')
            if not obj.check_table(global_objs['Local_Settings'].grab_item('W4S_TBL').decrypt_text()):
                mylist.append('W4S')
            if not obj.check_table(global_objs['Local_Settings'].grab_item('WE_TBL').decrypt_text()):
                mylist.append('WE')

            if len(mylist) > 0:
                header_text = 'Welcome to Vacuum Settings!\n{0} settings are invalid.\nPlease fix the network settings below:'\
                    .format(', '.join(mylist))
                obj.build_gui(header_text)
            else:
                my_return = True
        finally:
            obj.sql_close()
        del mylist

    obj.cancel()
    del obj

    return my_return


# Main loop for script that continuously searches the 02_To_Process and processes updates/errors whenever any exists
if __name__ == '__main__':
    if check_settings():
        global_objs['Event_Log'].write_log('')
        global_objs['Event_Log'].write_log('Starting Vacuum...')

        to_continue = False

        try:
            while 1 != 0:
                has_updates = None
                global_objs['Event_Log'].write_log('Vacuum sniffing floor for crumbs...')
                proc_errors()

                while has_updates is None:
                    has_updates = find_updates()
                    rand = random.randint(1, 10000000000)
                    sleep(1)

                    if to_continue:
                        global_objs['Event_Log'].write_log('')
                        to_continue = False
                    elif rand % 777 == 0:
                        gen_talk()

                proc_updates(has_updates)
                to_continue = True
        except:
            global_objs['Event_Log'].write_log(traceback.format_exc(), 'critical')

        finally:
            os.system('pause')
    else:
        global_objs['Event_Log'].write_log('Settings Mode was established. Need to re-run script', 'warning')

gc.collect()
