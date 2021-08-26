#----------------------------------------------------------------------------------
# Main Training Class, contains methods to import, calculate, and write training 
# data to excel
# 
#----------------------------------------------------------------------------------

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



from hrm_d import * 
from gpx_d import * 
from laps import * 
import numpy as np
import datetime
from statistics import mean


class training_d: 
    # Constructor
    def __init__(self, training_filename_dict): 
        self.gpx_filename = training_filename_dict['gpx_file']
        self.hrm_filename = training_filename_dict['hrm_file']
        self.xlsx_filename = training_filename_dict['xlsx_file']
        self.ov_xlsx_filename = training_filename_dict['overview_xlsx_file']
    
    def import_gpx_hrm(self): 
        self.hrmdata_fromfile()
        self.gpxdata_fromfile()
        self.timeoffset_nogps = int((self.gpx_d.data[0].time-self.hrm_d.Start_Date).total_seconds())
        
    def analyze_trdata(self): 
        self.import_gpx_hrm()
        self.gpx_d.calc_distance()
        self.get_type()
        self.analyze_laps() 
        if(len(self.userlaps_idx) !=0): 
            self.get_intervals()
           
          
    # Constructor for HRM Data - Import HRM
    def hrmdata_fromfile(self): 
        self.hrm_d = hrm_d(self.hrm_filename)
        
    # Constructor for GPX - Import GPX Data
    def gpxdata_fromfile(self): 
        self.gpx_d = gpx_d(self.gpx_filename, self.hrm_d.Interval)
    
     
    
    def get_type(self): 
        # Bestimmung Trainingstyp
        self.trtype = 'GA1'
        AvgHR = int(round(mean(self.hrm_d.hr),0))
        MaxHR = self.hrm_d.MaxHR
        #np_hr = np.array(self.hrm_d.hr)
        #hist_edges = [MaxHR*0.5, MaxHR*0.6, MaxHR*0.7, MaxHR*0.8, MaxHR*0.9, MaxHR]
        #hr_hist = np.histogram(np_hr, hist_edges)
        #np_hr_hist = np.array(hr_hist[0])/len(self.hrm_d.hr)

        #if((np_hr_hist[4]>0.1) & (0.68*MaxHR < AvgHR < 0.90*MaxHR)):
        #    self.trtype = 'Intervalle'
        if (AvgHR > (0.90*MaxHR)): 
            self.trtype = 'Competición'
        elif (AvgHR > (0.84*MaxHR)): 
            self.trtype = 'Entrenamiento'
        elif (AvgHR > (0.75*MaxHR)): 
            self.trtype = 'Recuperación'
        elif (AvgHR > (0.50*MaxHR)): 
            self.trtype = 'Ruta del colesterol'
        # elif (AvgHR > (0.85*MaxHR)): 
        #     self.trtype = 'GA1'
        
    def analyze_laps(self): 
        if(len(self.hrm_d.lap_totals) > self.gpx_d.dist_geod_no_alt[-1]/1000):
            self.get_lap_idx()
            self.userlaps = laps(self.hrm_d, self.gpx_d, self.userlaps_idx)
            self.userlaps.calc_laps()
            
        else: 
            self.autolaps_idx = list(range(len(self.hrm_d.lap_totals)))
            self.userlaps_idx = []
            
        self.autolaps = laps(self.hrm_d, self.gpx_d, self.autolaps_idx)
        self.autolaps.calc_laps()
    
    def get_lap_idx(self): 
        self.lap_dist = []
        self.autolaps_idx = []
        self.userlaps_idx = []
        self.delta_laplen = 20
        
        no_autolaps = floor((self.gpx_d.dist_geod_no_alt[-1]/1000))
        
        for lapidx in self.hrm_d.lap_totals: 
            lapidx_c = lapidx - self.timeoffset_nogps
            
            
            lap_dist_s = self.gpx_d.dist_geod_no_alt[lapidx_c-1]
            self.lap_dist.append(lap_dist_s)
                
        for autolap_opt_dist_s in range(no_autolaps): 
            autolap_act_dist_s = min(self.lap_dist, key=lambda x:abs(x-1000*(autolap_opt_dist_s+1)))
            #print([autolap_opt_dist_s, autolap_act_dist_s])
            self.autolaps_idx.append(self.lap_dist.index(autolap_act_dist_s))
        
        for lapidx in range(len(self.hrm_d.lap_totals)):
            if (lapidx not in self.autolaps_idx): 
                self.userlaps_idx.append(lapidx)

        #    m = round(lap_dist_s/1000,0)
        #    delta = abs(lap_dist_s - (m * 1000))
        #    #print(str(lap_dist_s) + ' ' + str(delta))
        #    i = self.hrm_d.lap_totals.index(lapidx)
        #    #print(lap_dist[-1] - ((len(autolaps_idx)+1) * 1000) )
        #    #if( (-self.delta_laplen) < (self.lap_dist[-1] - ((len(self.autolaps_idx)+1) * 1000)) <  + self.delta_laplen ):   
        #    if(delta < self.delta_laplen): 
        #        self.autolaps_idx.append(i)
        #    else:
        #        self.userlaps_idx.append(i)
        # 
        
    def get_intervals(self):
        
        AvgHR = int(round(mean(self.hrm_d.hr),0))
        MaxHR = self.hrm_d.MaxHR
        
        sum_hrmax = sum(hr_avg > 0.85*MaxHR for hr_avg in self.userlaps.hr_avg)   
        sum_hrmin = sum(hr_avg < 0.79*MaxHR for hr_avg in self.userlaps.hr_avg)  
        
        if ((sum_hrmax > 2) & (sum_hrmin >= sum_hrmax)):  
            self.trtype = 'Intervalos'
            intlaps_idx = []
            for userlap_idx_s in self.userlaps_idx:
                idx_s = self.userlaps_idx.index(userlap_idx_s)
                if((self.userlaps.hr_avg[idx_s] > 0.825*self.hrm_d.MaxHR) & (self.userlaps.pace[idx_s] < 4.75)):
                    intlaps_idx.append(userlap_idx_s)


                self.intlaps = laps(self.hrm_d, self.gpx_d, intlaps_idx)
                for intlap_idx_s in intlaps_idx: 
                    idx_s = self.userlaps_idx.index(intlap_idx_s)
                    self.intlaps.dist.append(self.userlaps.dist[idx_s])
                    self.intlaps.dur.append(self.userlaps.dur[idx_s])
                    self.intlaps.length.append(self.userlaps.length[idx_s])
                    self.intlaps.pace.append(self.userlaps.pace[idx_s])
                    self.intlaps.hr_max.append(self.userlaps.hr_max[idx_s])
                    self.intlaps.hr_avg.append(self.userlaps.hr_avg[idx_s])

            # Overview 
            self.int_no = len(intlaps_idx)

        
