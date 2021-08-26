# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# Class for individual xlswriter  
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
from math import sqrt, floor
from statistics import mean
import datetime
import time
import numpy as np



class trd_session_xlswriter(trd_xlswriter): 

    def __init__(self, training_filename_dict, my_training):
        self.xlsx_filename = training_filename_dict['xlsx_file']
        super(trd_session_xlswriter, self).__init__(self.xlsx_filename)
        
        self.mytr = my_training
        
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = "General"
        
        # Location of HRZones
        self.hrzones_col = 14
        self.hrzones_row = 2
        
        # Location of Autolaps
        self.lap_col = 14
        self.lap_row = 16
        self.no_autolaps = len(self.mytr.autolaps_idx)
        self.no_userlaps = len(self.mytr.userlaps_idx)

        
        
    ## Write pace and HRM Data to second Sheet  
    def write_hrmdata(self): 
        sec_unit = 1/(24*60*60) 

        my_row = 2
        ts1 = 82800
        self.ws1 = self.wb.create_sheet("Datos")
        self.ws1['A1'] = 'Segundos'
        #self.ws1['B1'] = 'Time'
        self.ws1['C1'] = 'Pulsaciones'
        self.ws1['D1'] = 'Velocidad'
        self.ws1['E1'] = 'Paso'
        self.ws1['F1'] = 'Alt'
        hrm_length = len(self.mytr.hrm_d.pace_u)
        for i in range(hrm_length):
            self.ws1.cell(row=my_row+i, column=1, value = i)
            #ts = datetime.datetime.fromtimestamp(i+ts1)
            #ts_str = ts.strftime('%H:%M:%S')
            self.ws1.cell(row=my_row+i, column=2, value = i*sec_unit)
            self.ws1.cell(row=my_row+i, column=3, value = self.mytr.hrm_d.hr[i])
            self.ws1.cell(row=my_row+i, column=4, value = self.mytr.hrm_d.speed[i])
            self.ws1.cell(row=my_row+i, column=5, value = self.mytr.hrm_d.pace_u[i])
            self.ws1.cell(row=my_row+i, column=6, value = self.mytr.hrm_d.alt[i])
            
    #+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
    ## Insert Overview 
    
    def write_overview(self):
        self.ws['A1'] = 'Datos'
        self.ws['A2'] = self.mytr.hrm_d.Start_Date.strftime('%d.%m.%Y')

        self.ws['C1'] = 'Duración'
        
        dur_h = self.mytr.hrm_d.Length_dt.tm_hour
        dur_m = self.mytr.hrm_d.Length_dt.tm_min
        dur_s = self.mytr.hrm_d.Length_dt.tm_sec
        
        self.ws['C2'] = '{0:02d}'.format(dur_h) + ':' + '{0:02d}'.format(dur_m) + ':' + '{0:02d}'.format(dur_s)

        self.ws['B1'] = 'Hora' 
        self.ws['B2'] = self.mytr.hrm_d.Start_Date.strftime("%H:%M:%S")

        self.ws['F1'] = 'Pulsaciones medias'
        self.ws['F2'] = int(round(mean(self.mytr.hrm_d.hr),0))

        self.ws['G1'] = 'Pulsaciones máximas'
        self.ws['G2'] = max(self.mytr.hrm_d.hr)

        self.ws['D1'] = 'Distancia'
        self.ws['D2'] = int(self.mytr.gpx_d.dist_geod_no_alt[-1])/1000

        self.ws['E1'] = 'Paso'
        self.ws['E2'] = '{0:02d}'.format(int(floor(self.mytr.gpx_d.pace))) + ':' + '{0:02d}'.format(int((self.mytr.gpx_d.pace*60)%60)) 

        self.ws['H1'] = 'Índice de carrera'
        self.ws['H2'] = self.mytr.hrm_d.RunningIndex

        self.ws['I1'] = 'Entrenamiento'
        self.ws['I2'] = self.mytr.trtype

    #+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
    ## Insert Autolaps 
    
    def write_autolaps(self): 
        
        my_col = self.lap_col 
        my_row = self.lap_row + 2 # Workaround as indices below are offset by 2
        #self.no_autolaps = len(self.mytr.autolaps.idx)
        self.ws.merge_cells(start_row=my_row-2, start_column=my_col, end_row=my_row-2, end_column=my_col+5)  
        self.ws.cell(row=my_row-2, column = my_col, value='Autolaps')
        self.ws.cell(row=my_row-1, column = my_col, value='Carrera')
        self.ws.cell(row=my_row-1, column = my_col+1, value='Distancia')
        self.ws.cell(row=my_row-1, column = my_col+2, value='Duración')
        self.ws.cell(row=my_row-1, column = my_col+3, value='Paso')
        self.ws.cell(row=my_row-1, column = my_col+4, value='Pulsaciones medias')
        self.ws.cell(row=my_row-1, column = my_col+5, value='Pulsaciones máximas')
        
        for i in range(self.no_autolaps):
            self.ws.cell(row=my_row+i, column=my_col, value = i+1)
            self.ws.cell(row=my_row+i, column=my_col+2, value = '{0:02d}'.format(floor(self.mytr.autolaps.dur[i]/60)) + ':' + '{0:02d}'.format(int(self.mytr.autolaps.dur[i]%60)))
            self.ws.cell(row=my_row+i, column=my_col+1, value = int(round(self.mytr.autolaps.length[i],0)))
            lap_pace_s = '{0:02d}'.format(int(floor(self.mytr.autolaps.pace[i]))) + ':' + '{0:02d}'.format(int((self.mytr.autolaps.pace[i]*60)%60))        
            self.ws.cell(row=my_row+i, column=my_col+3, value = lap_pace_s)
            self.ws.cell(row=my_row+i, column=my_col+4, value = self.mytr.autolaps.hr_avg[i])
            self.ws.cell(row=my_row+i, column=my_col+5, value = self.mytr.autolaps.hr_max[i])
        
            
    #+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
    ## Insert Userlaps
    
    def write_userlaps(self): 
        my_col = self.lap_col 
        my_row = self.lap_row + self.no_autolaps + 7
        
        #self.no_userlaps = len(self.mytr.userlaps.idx)
        
        #Runden Einfügen
        self.ws.merge_cells(start_row=my_row-2, start_column=my_col, end_row=my_row-2, end_column=my_col+5)
        self.ws.cell(row=my_row-2, column = my_col, value='Vueltas')
        self.ws.cell(row=my_row-1, column = my_col, value='Ronda')
        self.ws.cell(row=my_row-1, column = my_col+1, value='Distancia')
        self.ws.cell(row=my_row-1, column = my_col+2, value='Duración')
        self.ws.cell(row=my_row-1, column = my_col+3, value='Paso')
        self.ws.cell(row=my_row-1, column = my_col+4, value='Pulsaciones medias')
        self.ws.cell(row=my_row-1, column = my_col+5, value='Pulsaciones máximas')
        for i in range(self.no_userlaps):
            self.ws.cell(row=my_row+i, column=my_col, value = i+1)
            self.ws.cell(row=my_row+i, column=my_col+2, value = '{0:02d}'.format(floor(self.mytr.userlaps.dur[i]/60)) + ':' + '{0:02d}'.format(int(self.mytr.userlaps.dur[i]%60)))
            self.ws.cell(row=my_row+i, column=my_col+1, value = int(round(self.mytr.userlaps.length[i],0)))
            lap_pace_s = '{0:02d}'.format(int(floor(self.mytr.userlaps.pace[i]))) + ':' + '{0:02d}'.format(int((self.mytr.userlaps.pace[i]*60)%60))              
            self.ws.cell(row=my_row+i, column=my_col+3, value = lap_pace_s)
            self.ws.cell(row=my_row+i, column=my_col+4, value = self.mytr.userlaps.hr_avg[i])
            self.ws.cell(row=my_row+i, column=my_col+5, value = self.mytr.userlaps.hr_max[i])
            
    #+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++     
    ## Write Heart Zone 
    
    def write_hrzones(self):
        hr = self.mytr.hrm_d.hr
        MaxHR = self.mytr.hrm_d.MaxHR
        
        my_col = self.hrzones_col
        my_row = self.hrzones_row
        
        ts1 = 82800
        np_hr = np.array(hr)
        hist_edges = [floor(MaxHR*0.5), floor(MaxHR*0.6), floor(MaxHR*0.7), floor(MaxHR*0.8), floor(MaxHR*0.9), MaxHR]
        hr_hist = np.histogram(np_hr, hist_edges)
        np_hr_hist = np.array(hr_hist[0])/len(hr)
        #np_hr_hist/len(hr)
        total_s = len(hr)*self.mytr.hrm_d.Interval
        ts_z = []
        ts_z_str = []
        
        # Write Header
        # self.ws.merge_cells(start_row=my_row-2, start_column=my_col, end_row=my_row-2, end_column=my_col+1)
        self.ws.merge_cells(start_row=my_row-1, start_column=my_col, end_row=my_row-1, end_column=my_col+3)
        self.ws.cell(column=my_col, row=my_row-1).value = 'Zonas de esfuero'
        self.ws.cell(column=my_col, row=my_row).value = 'Zona'
        self.ws.cell(column=my_col+2, row=my_row).value = '%'
        self.ws.cell(column=my_col+3, row=my_row).value = 'Duración'
        
        # Write HRzones Data
        for i in range(6): 
            self.ws.merge_cells(start_row=my_row+2*i, start_column=my_col+1, end_row=my_row+(2*i+1), end_column=my_col+1)
            self.ws.cell(column=my_col+1, row=my_row+2*i).value = hist_edges[5-i]
            
        for i in range(5): 
            self.ws.merge_cells(start_row=my_row+2*i+1, start_column=my_col, end_row=my_row+(2*i)+2, end_column=my_col)
            self.ws.merge_cells(start_row=my_row+2*i+1, start_column=my_col+2, end_row=my_row+(2*i)+2, end_column=my_col+2)
            self.ws.merge_cells(start_row=my_row+2*i+1, start_column=my_col+3, end_row=my_row+(2*i)+2, end_column=my_col+3)
            
            self.ws.cell(column=my_col, row=my_row+2*i+1).value='Z' + str(5-i)
            self.ws.cell(column=my_col+2, row=my_row+2*i+1).value=float(round(np_hr_hist[4-i],4))
            self.ws.cell(column=my_col+2, row=my_row+2*i+1).number_format = '0.00%'
            ts_z.append(datetime.datetime.fromtimestamp(float(round(np_hr_hist[4-i],4))*total_s+ts1))
            ts_z_str.append(ts_z[i].strftime('%H:%M:%S'))
            self.ws.cell(column=my_col+3, row=my_row+2*i+1).value=ts_z_str[i]      
               
    #+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
    ## Format Cells 
    
                
    def format_overview(self): 
        my_col_d = 8
        my_row_d = 1
        
        self.format_header(1, 1, 1+my_col_d, 1+my_row_d)
                
    def format_hrzones(self): 
        my_row = self.hrzones_row
        my_row_d = 10
        my_col = self.hrzones_col
        my_col_d = 3
        
        self.format_header(my_col, my_row-1, my_col+my_col_d, my_row-1)

        self.format_header(my_col, my_row, my_col, my_row)
        self.format_header(my_col+2, my_row, my_col+my_col_d, my_row)

        self.format_data(my_col, my_row+1, my_col+my_col_d, my_row+my_row_d)
        self.format_data(my_col+1, my_row, my_col+1, my_row)
        self.format_data(my_col+1, my_row+my_row_d+1, my_col+1, my_row+my_row_d+1)

                
    def format_autolaps(self): 
        my_col = self.lap_col 
        my_row = self.lap_row
        my_col_d = 5
        my_row_d = self.no_autolaps
        
        self.format_header(my_col,my_row, my_col+my_col_d, my_row+1)
        if(my_row_d > 0): 
            self.format_data(my_col, my_row+2, my_col+my_col_d, my_row+my_row_d+1)
                
    def format_userlaps(self):
        my_col = self.lap_col 
        my_row = self.lap_row + self.no_autolaps + 5
        my_col_d = 5
        my_row_d = self.no_userlaps 
        
        self.format_header(my_col,my_row, my_col+my_col_d, my_row+2)
        self.format_data(my_col, my_row+2, my_col+my_col_d, my_row+my_row_d+1)
        
        
    # Change Column width    
    def change_columnwidth(self):        
        self.ws.column_dimensions["A"].width = 13.0
        self.ws.column_dimensions["B"].width = 10.0
        self.ws.column_dimensions["D"].width = 10.0
        self.ws.column_dimensions["E"].width = 10.0
        self.ws.column_dimensions["F"].width = 10.0
        self.ws.column_dimensions["G"].width = 10.0
        self.ws.column_dimensions["H"].width = 15.0
        self.ws.column_dimensions["I"].width = 15.0        
    
    
    #+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ## Create Chart of Heart Rate and Speed
    
    def create_hrmgraph(self): 
        
        data_len = len(self.mytr.hrm_d.pace_u)

        # First Chart (HeartRate)
        c1 = ScatterChart()
        #c1.title = ""
        c1.style = 13
        c1.y_axis.title = 'Pulsaciones (bpm)'
        c1.x_axis.title = 'Duración'
        c1.y_axis.title.tx.rich.p[0].r[0].rPr = CharacterProperties(sz=1200)
        c1.x_axis.title.tx.rich.p[0].r[0].rPr = CharacterProperties(sz=1200)


        hr_values = Reference(self.ws1, min_col=3, min_row=1, max_col=3, max_row=data_len+1)
        xvalues = Reference(self.ws1, min_col=2, min_row=1, max_col=2, max_row=data_len+1)
        hr_series = Series(hr_values, xvalues, title_from_data=True)
        hr_series.graphicalProperties.line.width = 2000
        hr_series.graphicalProperties.line.solidFill = "FF0000"

        c1.series.append(hr_series)

        # Second Chart: Speed
        c2 = ScatterChart()

        c2.y_axis.title = 'Velocidad (km/h)'
        c2.y_axis.title.tx.rich.p[0].r[0].rPr = CharacterProperties(sz=1200)


        spd_values = Reference(self.ws1, min_col=4, min_row=1, max_col=4, max_row=data_len+1)
        #xvalues = Reference(ws1, min_col=2, min_row=1, max_col=2, max_row=len(pace_u)+1)
        spd_series = Series(spd_values, xvalues, title_from_data=True)
        spd_series.graphicalProperties.line.width = 2000
        spd_series.graphicalProperties.line.solidFill ="00AAAA"

        c2.series.append(spd_series)

        c2.y_axis.axId = 200

        c2.y_axis.crosses = "max"
        c1 += c2
     
        # Plot the chart
        sec_unit = 1/(24*60*60) 
        
        c1.x_axis.number_format = '[h]:mm:ss'

        c1.x_axis.scaling.max = data_len * sec_unit
        c1.x_axis.scaling.min = 0
        
        if(data_len < 60*40): 
            unit_minutes = 5
        elif (data_len < 60*75):
            unit_minutes = 10  
        elif (data_len < 60*125):
            unit_minutes = 15
        elif (data_len < 60*60*4):
            unit_minutes = 30
        else:
            unit_minutes = 60

        c1.x_axis.majorUnit = sec_unit * 60 * unit_minutes

        c1.width = 20
        c1.height = 10

        c1.legend = None

        self.ws.add_chart(c1, "B5")
    
    # high level functions to write all, format all, save. 
    
    def write_all(self): 
        self.write_overview()
        self.write_hrzones()
        self.write_hrmdata()
        if(self.no_autolaps > 0): 
            self.write_autolaps()
        if(self.no_userlaps > 0): 
            self.write_userlaps()
    
    def format_all(self): 
        self.format_overview()
        self.format_hrzones()
        if(self.no_autolaps > 0): 
            self.format_autolaps()
        if(self.no_userlaps > 0): 
            self.format_userlaps()
            
        self.change_columnwidth()

    
    def create_all(self): 
        self.write_all()
        self.create_hrmgraph()
        self.format_all()
        