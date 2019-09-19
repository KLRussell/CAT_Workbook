from Global import grabobjs
from Settings import SettingsGUI
from Global import XMLParseClass
from Global import XMLAppendClass

import os
import pandas as pd
import numpy as np
import traceback
import gc

# Global Variable declaration
curr_dir = os.path.dirname(os.path.abspath(__file__))
main_dir = os.path.dirname(curr_dir)
uploaded_dir = os.path.join(main_dir, '03_Processed', '01_Uploaded')
errors_dir = os.path.join(main_dir, '04_Errors')
global_objs = grabobjs(main_dir, 'Vacuum')


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
                WE.Error_Msg,
                ROW_NUMBER()
                    OVER(PARTITION BY Source_File, Source_File_ID ORDER BY Edit_DT) Filter

            FROM {0} As WE
        '''.format(global_objs['Local_Settings'].grab_item('WE_TBL').decrypt_text()))

        if self.df.empty:
            return False
        else:
            self.df = self.df.loc[self.df['Filter'] == 1]
            self.df.drop(['Filter'], axis=1)
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


# Function to process Errors if errors exists in the SQL Server Error table
def proc_errors():
    obj = ErrorProcessing()
    try:
        if obj.check_errors() and obj.process_errors():
            obj.truncate_tbl()
        else:
            global_objs['Event_Log'].write_log('No Errors Found')
    finally:
        obj.close_conn()


# Checks network and sql server table settings to see if it exists and whether those settings are valid
#   A GUI will pop-up if settings need to be inputed or whether settings are invalid
def check_settings():
    my_return = False
    obj = SettingsGUI()

    if not os.path.exists(errors_dir):
        os.makedirs(errors_dir)

    if not global_objs['Settings'].grab_item('Server')\
            or not global_objs['Settings'].grab_item('Database') \
            or not global_objs['Local_Settings'].grab_item('WE_TBL'):
        header_text = 'Welcome to Vacuum Settings!\nSettings haven''t been established.\nPlease fill out all empty fields below:'
        obj.build_gui(header_text)
    else:
        mylist = []

        try:
            if not obj.sql_connect():
                mylist.append('network')
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
        global_objs['Event_Log'].write_log('Vacuum searching for errors')

        try:
            proc_errors()
        except:
            global_objs['Event_Log'].write_log(traceback.format_exc(), 'critical')
    else:
        global_objs['Event_Log'].write_log('Settings Mode was established. Need to re-run script', 'warning')

gc.collect()
