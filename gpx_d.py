

#----------------------------------------------------------------------------------
# GPX Files Class, contains import functions and datastructure of hrm data
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

import gpxpy
from geopy import distance
from math import sqrt, floor

class gpx_d: 
    def __init__(self,gpx_filename, gpx_interval):
        self.gpx_interval = gpx_interval
        
        gpx_file = open(gpx_filename,'r')
        gpx = gpxpy.parse(gpx_file)
        self.data = gpx.tracks[0].segments[0].points
        for i in range(len(self.data)):
            self.data[i].time = self.data[i].time.replace(tzinfo=None)

    def calc_distance(self): 
        self.alt_dif = [0]
        self.time_dif = [0]
        self.dist_geod_no_alt = [0]
        self.dist_dif_geod_2d = [0]
        #dist_dif_vin_2d

        for index in range(len(self.data)):
            if index == 0:
                pass
            else:
                start = self.data[index-1]

                stop = self.data[index]


                #dist_dif_vin_2d.append(distance_vin_2d)

                distance_geod_2d = distance.geodesic((start.latitude, start.longitude), (stop.latitude, stop.longitude)).m

                self.dist_dif_geod_2d.append(distance_geod_2d)


                if(distance_geod_2d < 20): 
                    self.dist_geod_no_alt.append(self.dist_geod_no_alt[-1] + distance_geod_2d)
                else: 
                    if(self.dist_geod_no_alt[-1] == 0): 
                        pace_tmp = 5.5
                    else: 
                        pace_tmp = sum(self.time_dif)*1000/(self.dist_geod_no_alt[-1]*60)
                    self.dist_geod_no_alt.append(self.dist_geod_no_alt[-1] + self.gpx_interval/(pace_tmp*60))

                time_delta = (stop.time - start.time).total_seconds()

                self.time_dif.append(time_delta)


        self.pace = sum(self.time_dif)*1000/(self.dist_geod_no_alt[-1]*60)
        
        #print('Geodesic 2D : ', self.dist_geod_no_alt[-1])
        #print('Total GPS Time : ', floor(sum(self.time_dif)/60),' min ', int(sum(self.time_dif)%60),' sec ')
        #print('Pace: ', floor(self.pace),' min ', int((self.pace*60)%60),' sec ')


        