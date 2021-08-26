#-------------------------------------------------------------------------
# 
# (C) 2019, jdmorise@yahoo.com
#  
# 	 This program is free software: you can redistribute it and/or modify
# 	 it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
#
#-------------------------------------------------------------------------


version = '0.7.0'

import os
import sys
import argparse
from training_d import *
from trd_oview_xlswriter import * 
from trd_session_xlswriter import * 

def main(): 
    print_header()
    
    usage = """"Examples: 
    PyPersonalTrainer -m scan -of Overview_2019.xlsx -sf Trainingsheets_2019 -tf Training_2019 
    PyPersonalTrainer -m add -t Training_2019/19010601.gpx -of Overview_2019.xlsx -sf Trainingsheets_2019
    """

    parser = argparse.ArgumentParser()
    parser.add_argument('-m','--mode', default='scan', choices=['scan', 'SCAN', 'add', 'ADD'],  help='Operation Mode: SCAN checks and processes new files in training folder, ADD will add an individual file (-t argument required)')
    parser.add_argument('-l','--license',help='show license terms', action='store_true')
    parser.add_argument('-f','--force',help='force processing of files', action='store_true')
    parser.add_argument('-o','--overview_workbook_filename', default='Overview.xls', help='Filename of excel worksheet where monthly and yearly training overview will be stored') 
    parser.add_argument('-sf','--session_workbook_foldername', default='Trainingsheets', help='Foldername of excel worksheet where individual training sessions will be stored')
    parser.add_argument('-tf','--training_foldername', default='Training', help='Foldername where gpx and hrm training data is stored')
    parser.add_argument('-t','--training_filename', help='GPX Filename of individual training session')
    
    args = parser.parse_args()

    if len(sys.argv) == 1:
        parser.print_help()
        return(1)
    elif (args.license): 
        print_license()
        return(0)
        
    else : 
      
        # Default Settings
        # Folder of HRM and GPX Files
        tr_foldername = args.training_foldername
        # Folder with results
        wb_foldername = args.session_workbook_foldername
        mode = args.mode
        
        dot = '.'
        filename_dict = []
        if ((mode == 'scan') or  (mode == 'SCAN')): 

            # Check if other filenames are given on the commandline           
            tr_filepath = os.getcwd() + '\\' + tr_foldername
            tr_filenames = os.listdir(tr_filepath)
        elif((mode == 'add') or  (mode == 'ADD')):
            tr_filenames = []
            tr_filenames.append(args.training_filename)
        
        else: 
            parser.print_help()
            return 1

        # Iterate over files in directory
        for filename_s in tr_filenames: 
            if '.gpx' in filename_s:
                filename_trunk_s = filename_s.split('.')[0]
                if (os.path.isfile(tr_foldername + '/' + filename_trunk_s + '.hrm')):
                    filename_dict_s = {'hrm_file' : tr_foldername + '/' + filename_trunk_s + '.hrm'}
                    filename_dict_s['gpx_file'] = tr_foldername + '/' + filename_trunk_s + '.gpx'
                    filename_dict_s['xlsx_file'] = wb_foldername + '/' + filename_trunk_s + '.xlsx'
                    filename_dict_s['overview_xlsx_file'] = args.overview_workbook_filename
                    
                    filename_dict.append(filename_dict_s)

        # Iterate over list of filenames
        for filename_dict_s in filename_dict: 

            if(args.force or not(os.path.exists(filename_dict_s['xlsx_file']))):
                analyze_training(filename_dict_s)
                print('Session Sheet: ', filename_dict_s['xlsx_file'], ' written')
            else:
                print('Session Sheet: ', filename_dict_s['xlsx_file'], ' exists')
        
        print('-------------------------------------------------------------------------') 
        print('Execution Finished') 
        
def print_header(): 
    print('-------------------------------------------------------------------------')
    print('PyPersonalTrainer V', version)
    print('Training analyzer for Polar HRM and GPX files')
    print('(c) 2019 JDMORISE at YAHOO.COM')
    print('-------------------------------------------------------------------------')
    

def print_license(): 
    # print('-------------------------------------------------------------------------')
    # print(' ') 
    # print(' (C) 2019, jdmorise@yahoo.com')
    print(' ')
    print(' Traducción realizada al español por Luis Zurita Herrera')
    print(' This program is free software: you can redistribute it and/or modify')
    print(' it under the terms of the GNU General Public License as published by')
    print(' the Free Software Foundation, either version 3 of the License, or')
    print(' (at your option) any later version.')
    print(' ')
    print(' This program is distributed in the hope that it will be useful,')
    print(' but WITHOUT ANY WARRANTY; without even the implied warranty of')
    print(' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the')
    print(' GNU General Public License for more details.')
    print(' ')
    print(' You should have received a copy of the GNU General Public License')
    print(' along with this program.  If not, see <https://www.gnu.org/licenses/>')
    print(' ')
    print('-------------------------------------------------------------------------')

# Test box for training class
def analyze_training(filename_dict_s): 
    
    my_training = training_d(filename_dict_s)
    my_training.analyze_trdata()
    # Create Session and Write Excel Sheet
    my_trd_session_xlsx = trd_session_xlswriter(filename_dict_s,my_training)
    my_trd_session_xlsx.create_all()
    my_trd_session_xlsx.save()
    
    # Create Overview and write Excel Sheet
    my_trd_oview_xlsx = trd_oview_xlswriter(filename_dict_s,my_training)
    my_trd_oview_xlsx.update_workbook()
    my_trd_oview_xlsx.save()
    
    # Garbage Collection after Finishing
    del my_training
    del my_trd_oview_xlsx 
    del my_trd_session_xlsx
    
if (__name__ == '__main__'):
    main()