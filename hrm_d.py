#----------------------------------------------------------------------------------
# HRM Files Class, contains import functions and datastructure of hrm data
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

import datetime
import time

class hrm_d: 
    def __init__(self,hrm_filename): 

        results = []
        self.hr = []
        self.speed = []
        self.alt = []

        self.lap_hr_avg = []
        self.lap_hr_mom = []
        self.lap_hr_max = []
        self.lap_starttime = []
        lap_h = []
        lap_m = []
        lap_s = []
        self.lap_totals = []
        self.pace_u = []

        import_hrm = 0
        import_laps = 0
        self.hrmVersion = 0
        st = datetime.datetime(1970,1,1,0,0,0,0)


        with open(hrm_filename) as inputfile:
            for line in inputfile:
                if(line[0] != '\n'):
                    if '[IntNotes]' in line: 
                        import_laps = 0
                    if ((import_hrm == 1) & (self.hrmVersion == 106)): 
                        #print(line)
                        [hr_s, speed_s, alt_s] = line.strip().split('\t')
                        self.hr.append(int(hr_s))
                        self.speed.append(float(speed_s)/10)
                        if(float(speed_s) != 0): 
                            self.pace_u.append(600/float(speed_s))
                        else:
                            self.pace_u.append(0)
                        self.alt.append(int(alt_s))
                    if ((import_laps == 1) & (self.hrmVersion == 106)): 
                        lap_s = line.strip().split('\t')
                        if (lap_s[0] != '0'): 
                            lap_starttime_s = lap_s[0].strip().split('.')[0]
                            self.lap_starttime.append(time.strptime(lap_starttime_s,"%H:%M:%S"))
                            [h_s, m_s, s_s] = lap_starttime_s.strip().split(':')
                            lap_h.append(int(h_s))
                            lap_m.append(int(m_s))
                            lap_s.append(int(s_s))
                            
                            self.lap_totals.append(int(h_s)*3600 + int(m_s)*60 + int(s_s))
                            self.lap_hr_mom.append(int(lap_s[1]))
                            self.lap_hr_avg.append(int(lap_s[3]))
                            self.lap_hr_max.append(int(lap_s[4]))
                    if 'Version' in line: 
                        self.hrmVersion = int(line.strip().split('=')[1])
                    if 'Date' in line: 
                        Date = line.strip().split('=')[1]
                        self.Start_Date = st.strptime(Date,"%Y%m%d")
                    if 'StartTime' in line: 
                        StartTime = line.strip().split('=')[1]
                        StartTime_S = StartTime.strip().split('.')[0] 

                        self.Start_Time = self.Start_Date.strptime(StartTime_S,"%H:%M:%S")
                    if 'Length' in line: 
                        Length = line.strip().split('=')[1]
                        Length_S = Length.strip().split('.')[0]
                        
                        self.Length_dt = time.strptime(Length_S,"%H:%M:%S")  

                    if 'Interval' in line: 
                        self.Interval = int(line.strip().split('=')[1])
                    if 'MaxHR' in line: 
                        self.MaxHR = int(line.strip().split('=')[1])
                    if 'RestHR' in line: 
                        self.RestHR = int(line.strip().split('=')[1])
                    if 'VO2max' in line: 
                        self.VO2max = int(line.strip().split('=')[1])
                    if 'Weight' in line: 
                        self.Weight = int(line.strip().split('=')[1])
                    if 'RunningIndex' in line: 
                        self.RunningIndex = int(line.strip().split('=')[1])    
                    if '[HRData]' in line: 
                        import_hrm = 1
                    if '[IntTimes]' in line: 
                        import_laps = 1

        self.Start_Date = self.Start_Date.replace(hour=self.Start_Time.hour, minute=self.Start_Time.minute, second=self.Start_Time.second)          