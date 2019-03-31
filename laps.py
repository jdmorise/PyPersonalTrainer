#----------------------------------------------------------------------------------
# Laps Parent Class, contains import functions and datastructure of hrm data
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


class laps: 
    def __init__(self, hrm_d, gpx_d, lap_idx):
        self.idx = lap_idx
        
        self.gpx_d = gpx_d
        self.hrm_d = hrm_d
        self.dist = []
        self.length = []
        self.dur = [] 
        self.pace = []
        self.hr_max = []
        self.hr_avg = []
    
    def calc_laps(self):
        
        laptime_offset = int((self.gpx_d.data[0].time-self.hrm_d.Start_Date).total_seconds())
        
        for idx_s in self.idx: 

            idx_c = self.hrm_d.lap_totals[idx_s] - laptime_offset

            if(self.length == []):
                lap_len_s = (self.gpx_d.dist_geod_no_alt[idx_c-1])
                lap_dur_s = self.hrm_d.lap_totals[idx_s] - laptime_offset
            else:
                idx_m1 = self.idx[self.idx.index(idx_s)-1]
                # print(
                lap_len_s = (self.gpx_d.dist_geod_no_alt[idx_c-1]-self.dist[-1])
                lap_dur_s = self.hrm_d.lap_totals[idx_s] - self.hrm_d.lap_totals[idx_m1]
            self.dist.append(self.gpx_d.dist_geod_no_alt[idx_c-1])
            self.dur.append(lap_dur_s)
            self.length.append(lap_len_s)
            self.pace.append(lap_dur_s * 1000/(60 * lap_len_s))
            self.hr_max.append(self.hrm_d.lap_hr_max[idx_s])
            self.hr_avg.append(self.hrm_d.lap_hr_avg[idx_s])
