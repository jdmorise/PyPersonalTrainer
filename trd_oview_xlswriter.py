# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# Class for overview xlswriter  
# Writes and formats the respective data to the excel file in the 
# filename
# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

#-------------------------------------------------------------------------
# 
# (C) 2019, jdmorise at yahoo.com
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



from trd_xlswriter import * 
import os
from math import sqrt, floor
from statistics import mean
import datetime
import time
import numpy as np

class trd_oview_xlswriter(trd_xlswriter): 
    def __init__(self, filename_dict_s, my_training): 
        self.xlsx_filename = filename_dict_s['overview_xlsx_file']
        super(trd_oview_xlswriter,self).__init__(self.xlsx_filename)

        self.my_tr = my_training
        self.months = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
        self.weekdays = ['Lu', 'Ma', 'Mi', 'Ju', 'Vi', 'Sa', 'Do' ]

        if(os.path.exists(self.xlsx_filename)):
            
            #print('File Exists')
            self.wb = load_workbook(self.xlsx_filename)
            self.ws = self.wb.active
        else: 
            #print('Create New File')
            self.create_oview_workbook()
            
    def create_oview_workbook(self): 
        self.wb = Workbook()
        # Write header files to newly created file
        self.create_oview_sheet()
        self.create_interval_sheet()
        self.create_months_sheet()
    
    def create_oview_sheet(self): 
        # Überblickssheet
        
        self.ws = self.wb.active
        self.ws.title = "General"
        self.ws['A1'] = 'Mess' 
        self.ws['B1'] = 'Trayecto' 
        self.ws['C1'] = 'Duración' 
        self.ws['D1'] = 'Sesión de entrenamiento'
        
        for i in range(12): 
            month_s = self.months[i]
            self.ws.cell(row=i+2, column=1).value = month_s
            self.ws.cell(row=i+2, column=2).value = '=Sum(' + month_s + '!G2:' + month_s + '!G32)'
            self.ws.cell(row=i+2, column=3).value = '=Sum(' + month_s + '!D2:' + month_s + '!D32)'
            self.ws.cell(row=i+2, column=3).number_format = 'hh:mm:ss'
            
        self.format_header(1, 1, 4, 13)
        
        self.ws.column_dimensions["A"].width = 13.0
        self.ws.column_dimensions["B"].width = 10.0
        self.ws.column_dimensions["C"].width = 10.0
        self.ws.column_dimensions["D"].width = 10.0
        



    # Intervalle
    def create_interval_sheet(self): 
        self.ws = self.wb.create_sheet('Intervalle')
        self.ws['A1'] = 'Datos'
        self.ws['B1'] = 'Tiempo' 
        self.ws['C1'] = 'Etiqueta'
        self.ws['D1'] = 'Intervalos'
        self.ws['E1'] = 'Distancia media'
        self.ws['F1'] = 'Duración media'
        self.ws['G1'] = 'Velocidad media'

        for i in range(10): 
            self.ws.cell(row=1, column=8+3*i).value = ('Distancia #' + str(i))
            self.ws.cell(row=1, column=8+3*i+1).value = ('Duración #' + str(i))
            self.ws.cell(row=1, column=8+3*i+2).value = ('Velocidad #' + str(i))

        # Header of Laps
        self.ws.column_dimensions["A"].width = 13.0
        self.ws.column_dimensions["B"].width = 10.0
        self.ws.column_dimensions["D"].width = 10.0
        self.ws.column_dimensions["E"].width = 10.0
        self.ws.column_dimensions["F"].width = 10.0
        self.ws.column_dimensions["G"].width = 10.0
        self.ws.column_dimensions["H"].width = 8.0
        self.ws.column_dimensions["I"].width = 8.0
        self.ws.column_dimensions["J"].width = 8.0
        
        self.format_header(1, 1, 37, 1)
        
    def create_months_sheet(self): 
        # Sheet für jeden Monat
        for i in range(12): 
            self.ws = self.wb.create_sheet(self.months[i])

            self.ws['A1'] = 'Datos'
            self.ws['B1'] = 'Tiempo' 
            self.ws['C1'] = 'Etiqueta'
            self.ws['D1'] = 'Duración'
            self.ws['E1'] = 'Pulsaciones Medias'
            self.ws['F1'] = 'Pulsaciones Máximas'
            self.ws['G1'] = 'Distancia'
            self.ws['H1'] = 'Velocidad'
            self.ws['I1'] = 'Índice de carrera'
            self.ws['J1'] = 'Entrenamiento'
            
            self.format_header(1, 1, 10, 1)

            # Header of Laps
            self.ws.column_dimensions["A"].width = 13.0
            self.ws.column_dimensions["B"].width = 10.0
            self.ws.column_dimensions["D"].width = 10.0
            self.ws.column_dimensions["E"].width = 10.0
            self.ws.column_dimensions["F"].width = 10.0
            self.ws.column_dimensions["G"].width = 10.0
            self.ws.column_dimensions["H"].width = 10.0
            self.ws.column_dimensions["I"].width = 15.0
            self.ws.column_dimensions["J"].width = 15.0
            
    
    def update_oview_sheet(self): 
        st = datetime.datetime(1970,1,1,0,0,0,0)
        dur_str0 = '00:00:00'
        dt0 = st.strptime(dur_str0,"%H:%M:%S")
        ts1 = 82800

        # Überblicks Sheet aktualisieren
        dt_t = self.my_tr.hrm_d.Start_Date

        self.ws = self.wb[self.months[dt_t.month-1]]

        sum_len = 0
        training_no = self.ws.max_row-1
        sum_dur = 0
        dur_str0 = '00:00:00'

        for row in self.ws.iter_rows(min_col=7,  min_row=2, max_col=7, max_row=self.ws.max_row): 
            for cell in row:
                sum_len += cell.value
        for row in self.ws.iter_rows(min_col=4,  min_row=2, max_col=4, max_row=self.ws.max_row): 
            for cell in row:
                dt1 = st.strptime(cell.value,"%H:%M:%S")
                t = dt1 - dt0 
                sum_dur += t.total_seconds()

        self.ws = self.wb['General'] 
        self.ws.cell(column=2, row = dt_t.month+1).value = sum_len
        self.ws.cell(column=4, row = dt_t.month+1).value = training_no
        sum_dur_dt = datetime.datetime.fromtimestamp(sum_dur+ts1)
        self.ws.cell(column=3, row = dt_t.month+1).value = sum_dur_dt.strftime('%H:%M:%S')
        
    def update_months_row(self): 
        # Add content to excel
        st = datetime.datetime(1970,1,1,0,0,0,0)

        data_row = []
        dt_t = self.my_tr.hrm_d.Start_Date
        date_s = dt_t.strftime('%d.%m.%Y')
        time_s = dt_t.strftime("%H:%M:%S")
        
        weekday_s = self.weekdays[dt_t.weekday()]

        data_row.append(date_s)
        data_row.append(time_s)
        data_row.append(weekday_s)
        
        #dur_ts = sum(self.my_tr.gpx_d.time_dif)
        dur_h = self.my_tr.hrm_d.Length_dt.tm_hour
        dur_m = self.my_tr.hrm_d.Length_dt.tm_min
        dur_s = self.my_tr.hrm_d.Length_dt.tm_sec
               
        data_row.append('{0:02d}'.format(dur_h) + ':' + '{0:02d}'.format(dur_m) + ':' + '{0:02d}'.format(dur_s))
        data_row.append(int(round(mean(self.my_tr.hrm_d.hr),0)))
        data_row.append(max(self.my_tr.hrm_d.hr))
        data_row.append(round(int(self.my_tr.gpx_d.dist_geod_no_alt[-1])/1000,3))
        data_row.append('{0:02d}'.format(floor(self.my_tr.gpx_d.pace)) + ':' + '{0:02d}'.format(floor((self.my_tr.gpx_d.pace%1*60))))
        data_row.append(self.my_tr.hrm_d.RunningIndex)
        data_row.append(self.my_tr.trtype)

        self.ws = self.wb[self.months[dt_t.month-1]]

        if(self.ws.max_row == 1): 
            style_row = self.append_row(data_row)
            self.format_data(1,style_row, len(data_row), style_row)
            
            
        else: 
            append=True

            for row in self.ws.iter_rows(min_row=2, max_row=self.ws.max_row, min_col=1, max_col=1):
                for cell in row:

                    date_cell = st.strptime(cell.value,"%d.%m.%Y")
                    time_cell = st.strptime(self.ws.cell(row=cell.row, column=2).value, "%H:%M:%S")
                    dt_cell = date_cell.replace(hour=time_cell.hour, minute=time_cell.minute, second=time_cell.second)    

                    if(append & (dt_cell > dt_t)):
                        append=False
                        style_row = cell.row
                        self.insert_row(data_row, style_row)
                        self.format_data(1,style_row, len(data_row), style_row)

                    elif (dt_cell == dt_t):
                        append=False
                        style_row = cell.row
                        
            if(append): 
                style_row = self.append_row(data_row)
                self.format_data(1,style_row, len(data_row), style_row)
                
    #+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    # Update Interval Sheet
    def update_intervals_sheet(self): 
        st = datetime.datetime(1970,1,1,0,0,0,0)

        data_row = []
        dt_t = self.my_tr.hrm_d.Start_Date
        date_s = dt_t.strftime('%d.%m.%Y')
        time_s = dt_t.strftime("%H:%M:%S")
        
        weekday_s = self.weekdays[dt_t.weekday()]
        
        int_len_avg = int(np.mean(self.my_tr.intlaps.length))
        int_dur_avg = int(np.mean(self.my_tr.intlaps.dur))
        int_pace_avg = np.mean(self.my_tr.intlaps.pace)
        
        data_row.append(date_s)
        data_row.append(time_s)
        data_row.append(weekday_s)
        data_row.append(len(self.my_tr.intlaps.idx))
        data_row.append(int_len_avg)
        data_row.append('{0:02d}'.format(floor(int_dur_avg/60)) + ':' + '{0:02d}'.format(int_dur_avg%60))
        data_row.append('{0:02d}'.format(floor(int_pace_avg)) + ':' + '{0:02d}'.format(floor(int_pace_avg%1*60)))
                        
        for int_idx_s in range(len(self.my_tr.intlaps.idx)): 
            int_pace_s = self.my_tr.intlaps.pace[int_idx_s]
            int_dur_s = int(self.my_tr.intlaps.dur[int_idx_s])
            data_row.append(int(self.my_tr.intlaps.length[int_idx_s]))
            data_row.append('{0:02d}'.format(floor(int_dur_s/60)) + ':' + '{0:02d}'.format(int_dur_s%60))
            data_row.append('{0:02d}'.format(floor(int_pace_s)) + ':' + '{0:02d}'.format(floor(int_pace_s%1*60)))

        self.ws = self.wb['Intervalos']
        
        if(self.ws.max_row == 1): 
            style_row = self.append_row(data_row)
            self.format_data(1,style_row, len(data_row), style_row)
            
        else: 
            append=True

            for row in self.ws.iter_rows(min_row=2, max_row=self.ws.max_row, min_col=1, max_col=1):
                for cell in row:

                    date_cell = st.strptime(cell.value,"%d.%m.%Y")
                    time_cell = st.strptime(self.ws.cell(row=cell.row, column=2).value, "%H:%M:%S")
                    dt_cell = date_cell.replace(hour=time_cell.hour, minute=time_cell.minute, second=time_cell.second)    

                    if(append & (dt_cell > dt_t)):
                        append=False
                        style_row = cell.row
                        self.insert_row(data_row, style_row)
                        self.format_data(1,style_row, len(data_row), style_row)

                    elif (dt_cell == dt_t):
                        append=False
                        style_row = cell.row
                        
            if(append): 
                style_row = self.append_row(data_row)
                self.format_data(1,style_row, len(data_row), style_row)
            
    #+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    # Append and Insert functions
    def insert_row(self,data_row, row_ins):
        idx=0
        self.ws.insert_rows(row_ins)
        for row in self.ws.iter_rows(min_col=1,  min_row=row_ins, max_col=len(data_row), max_row=row_ins):
            for cell in row: 
                cell.value = data_row[idx]
                idx+=1
                
    def append_row(self, data_row): 
        self.ws.append(data_row)
        style_row = self.ws.max_row
        return style_row
    
    # high level methods of class
    
    def update_workbook(self): 
        self.update_months_row()
        self.update_oview_sheet() 
        #print(self.my_tr.trtype)
        if(self.my_tr.trtype == 'Intervalos'): 
            self.update_intervals_sheet()
    
        
        
        
    