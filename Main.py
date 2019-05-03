from Global import grabobjs
from Global import XMLParseClass
from Global import XMLAppendClass
from time import sleep

import os
import gc
import pathlib as pl
import random
import numpy as np
import pandas as pd
import traceback

curr_dir = os.path.dirname(os.path.abspath(__file__))
main_dir = os.path.dirname(curr_dir)
process_dir = os.path.join(main_dir, '02_To_Process')
processed_dir = os.path.join(main_dir, '03_Processed')
errors_dir = os.path.join(main_dir, '04_Errors')
global_objs = grabobjs(main_dir)


class CATWorkbook:
    def __init__(self, df, file):
        df['Source_File'] = file
        df['Source_File_ID'] = np.arange(len(df))
        self.sheet_name = df['Sheet_Name'][0]
        del df['Sheet_Name']

        self.asql = global_objs['SQL']
        try:
            self.asql.connect('alch')
            self.df = df
            self.file = file
        finally:
            self.asql.close()

    def close_conn(self):
        self.asql.close()

    def upload(self):
        try:
            if self.sheet_name == 'Sheet1':
                table = global_objs['Local_Settings'].grab_item('W1S_TBL')
                global_objs['Event_Log'].write_log('Uploading {0} items to {1} SQL table'.format(len(self.df), table))
                self.asql.upload(self.df, table, index=False)
            elif self.sheet_name == 'Sheet2':
                table = global_objs['Local_Settings'].grab_item('W2S_TBL')
                global_objs['Event_Log'].write_log('Uploading {0} items to {1} SQL table'.format(len(self.df), table))
                self.asql.upload(self.df, table, index=False)
            elif self.sheet_name == 'Sheet3':
                table = global_objs['Local_Settings'].grab_item('W3S_TBL')
                global_objs['Event_Log'].write_log('Uploading {0} items to {1} SQL table'.format(len(self.df), table))
                self.asql.upload(self.df, table, index=False)
            elif self.sheet_name == 'Sheet4':
                table = global_objs['Local_Settings'].grab_item('W4S_TBL')
                global_objs['Event_Log'].write_log('Uploading {0} items to {1} SQL table'.format(len(self.df), table))
                self.asql.upload(self.df, table, index=False)
            else:
                global_objs['Event_Log'].write_log(
                    'Unable to upload {0}. Spreadsheet has invalid Sheet_Name ''{1}'''.format(self.file,
                                                                                              self.sheet_name))
                return False
            return True
        except:
            return False

    def migrate_file(self, processed=True):
        if processed:
            global_objs['Event_Log'].write_log('Upload successful. Migrating file to Processed folder')
            os.rename(os.path.join(process_dir, self.file), os.path.join(processed_dir, self.file))
        else:
            global_objs['Event_Log'].write_log(
                'Migrating file to Processed folder as Upload_Error_{0}'.format(self.file))
            os.rename(os.path.join(process_dir, self.file), os.path.join(processed_dir,
                                                                         'Upload_Error_{0}'.format(self.file)))


class ErrorProcessing:
    df = pd.DataFrame()

    def __init__(self):
        self.asql = global_objs['SQL']
        try:
            self.asql.connect('alch')
        finally:
            self.asql.close()

    def close_conn(self):
        self.asql.close()

    def check_errors(self):
        self.df = self.asql.query('''
            SELECT
                WE.Source_File,
                WE.Source_File_ID,
                WNE.Norm_Error,
                WE.Error_Msg
                
            FROM {0} As WE
            INNER JOIN {1} As WNE
            ON
                WE.WNE_ID = WNE.WNE_ID
        '''.format(global_objs['Local_Settings'].grab_item('WE_TBL'),
                   global_objs['Local_Settings'].grab_item('WNE_TBL')))

        if self.df.empty:
            return False
        else:
            global_objs['Event_Log'].write_log('Vacuum clogged with Errors. Cleaning vacuum...')
            return True

    def process_errors(self):
        try:
            for source_file in self.df['Source_File'].unique().tolist():
                xmlobj = XMLParseClass(os.path.join(processed_dir, source_file))
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

    def truncate_tbl(self):
        self.asql.execute('TRUNCATE TABLE %s' % global_objs['Local_Settings'].grab_item('WE_TBL'))


def gen_talk():
    f = open(os.path.join(curr_dir, 'random_talk.txt'), 'r')
    lines = f.readlines()
    talkid = random.randint(0, len(lines) - 1)

    print('The Vacuum: {0}'.format(lines[talkid]))
    f.close()


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


def find_updates():
    f = list(pl.Path(process_dir).glob('*.xml')) + list(pl.Path(process_dir).glob('*.sql'))

    if f:
        global_objs['Event_Log'].write_log('Vacuum found crumbs... Switching to max suction')
        return f


def proc_updates(files):
    write_blank = False

    for file in files:
        folder_name = os.path.basename(os.path.dirname(file))
        file_name = os.path.basename(file)

        if write_blank:
            global_objs['Event_Log'].write_log('')

        global_objs['Event_Log'].write_log('Reading file ({0}/{1})'.format(folder_name, file_name))
        global_objs['Event_Log'].write_log('')

        if os.path.splitext(file_name)[1].lower() == '.xml':
            xmlobj = XMLParseClass(file)

            if xmlobj:
                df = xmlobj.parsexml('./{urn:schemas-microsoft-com:rowset}data/')

                if not df.empty:
                    obj = CATWorkbook(df, file_name)

                    try:
                        if obj.upload():
                            obj.migrate_file()
                        else:
                            obj.migrate_file(False)
                    finally:
                        obj.close_conn()
                        del obj

            del xmlobj
        else:
            global_objs['Event_Log'].write_log('Adding new user to CAT Employee table')
            newuser(file)
            os.remove(file)


def proc_errors():
    obj = ErrorProcessing()
    try:
        if obj.check_errors() and obj.process_errors():
            obj.truncate_tbl()
    finally:
        obj.close_conn()


def load_settings():
    if not os.path.exists(errors_dir):
        os.makedirs(errors_dir)

    if not os.path.exists(process_dir):
        os.makedirs(process_dir)

    if not os.path.exists(processed_dir):
        os.makedirs(processed_dir)




if __name__ == '__main__':
    load_settings()

    try:
        global_objs['SQL'].connect('alch')
    finally:
        global_objs['SQL'].close()

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

gc.collect()
