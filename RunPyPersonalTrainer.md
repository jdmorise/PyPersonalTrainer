

```python
#-------------------------------------------------------------------------
# 
# (C) 2019, jdmorise@yahoo.com
#  
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
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
```


```python
%run PyPersonalTrainer.py -h
```

    -------------------------------------------------------------------------
    PyPersonalTrainer V 0.7.0
    Training analyzer for Polar HRM and GPX files
    (c) 2019 JDMORISE at YAHOO.COM
    -------------------------------------------------------------------------
    usage: PyPersonalTrainer.py [-h] [-m {scan,SCAN,add,ADD}] [-l] [-f]
                                [-o OVERVIEW_WORKBOOK_FILENAME]
                                [-sf SESSION_WORKBOOK_FOLDERNAME]
                                [-tf TRAINING_FOLDERNAME] [-t TRAINING_FILENAME]
    
    optional arguments:
      -h, --help            show this help message and exit
      -m {scan,SCAN,add,ADD}, --mode {scan,SCAN,add,ADD}
                            Operation Mode: SCAN checks and processes new files in
                            training folder, ADD will add an individual file (-t
                            argument required)
      -l, --license         show license terms
      -f, --force           force processing of files
      -o OVERVIEW_WORKBOOK_FILENAME, --overview_workbook_filename OVERVIEW_WORKBOOK_FILENAME
                            Filename of excel worksheet where monthly and yearly
                            training overview will be stored
      -sf SESSION_WORKBOOK_FOLDERNAME, --session_workbook_foldername SESSION_WORKBOOK_FOLDERNAME
                            Foldername of excel worksheet where individual
                            training sessions will be stored
      -tf TRAINING_FOLDERNAME, --training_foldername TRAINING_FOLDERNAME
                            Foldername where gpx and hrm training data is stored
      -t TRAINING_FILENAME, --training_filename TRAINING_FILENAME
                            GPX Filename of individual training session
    


```python
#Example with scan of folder, new training files were added
```


```python
%run PyPersonalTrainer.py -m scan -o ../Ueberblick_2019.xlsx -tf ../Trainingsdaten_2019 -sf ../Trainingsheets_2019
```

    -------------------------------------------------------------------------
    PyPersonalTrainer V 0.7.0
    Training analyzer for Polar HRM and GPX files
    (c) 2019 JDMORISE at YAHOO.COM
    -------------------------------------------------------------------------
    Session Sheet:  ../Trainingsheets_2019/19010601.xlsx  exists
    Session Sheet:  ../Trainingsheets_2019/19011001.xlsx  exists
    Session Sheet:  ../Trainingsheets_2019/19011301.xlsx  exists
    Session Sheet:  ../Trainingsheets_2019/19011701.xlsx  exists
    Session Sheet:  ../Trainingsheets_2019/19012701.xlsx  exists
    Session Sheet:  ../Trainingsheets_2019/19013101.xlsx  exists
    Session Sheet:  ../Trainingsheets_2019/19020301.xlsx  exists
    Session Sheet:  ../Trainingsheets_2019/19022101.xlsx  exists
    Session Sheet:  ../Trainingsheets_2019/19022701.xlsx  exists
    Session Sheet:  ../Trainingsheets_2019/19030901.xlsx  exists
    Session Sheet:  ../Trainingsheets_2019/19031201.xlsx  exists
    Session Sheet:  ../Trainingsheets_2019/19031401.xlsx  exists
    Session Sheet:  ../Trainingsheets_2019/19031701.xlsx  exists
    Session Sheet:  ../Trainingsheets_2019/19032401.xlsx  written
    -------------------------------------------------------------------------
    Execution Finished
    


```python
# Example with use of force flag 
```


```python
%run PyPersonalTrainer.py -f -m scan -o ../Ueberblick_2019.xlsx -tf ../Trainingsdaten_2019 -sf ../Trainingsheets_2019
```

    -------------------------------------------------------------------------
    PyPersonalTrainer V 0.7.0
    Training analyzer for Polar HRM and GPX files
    (c) 2019 JDMORISE at YAHOO.COM
    -------------------------------------------------------------------------
    Session Sheet:  ../Trainingsheets_2019/19010601.xlsx  written
    Session Sheet:  ../Trainingsheets_2019/19011001.xlsx  written
    Session Sheet:  ../Trainingsheets_2019/19011301.xlsx  written
    Session Sheet:  ../Trainingsheets_2019/19011701.xlsx  written
    Session Sheet:  ../Trainingsheets_2019/19012701.xlsx  written
    Session Sheet:  ../Trainingsheets_2019/19013101.xlsx  written
    Session Sheet:  ../Trainingsheets_2019/19020301.xlsx  written
    Session Sheet:  ../Trainingsheets_2019/19022101.xlsx  written
    Session Sheet:  ../Trainingsheets_2019/19022701.xlsx  written
    Session Sheet:  ../Trainingsheets_2019/19030901.xlsx  written
    Session Sheet:  ../Trainingsheets_2019/19031201.xlsx  written
    Session Sheet:  ../Trainingsheets_2019/19031401.xlsx  written
    Session Sheet:  ../Trainingsheets_2019/19031701.xlsx  written
    Session Sheet:  ../Trainingsheets_2019/19032401.xlsx  written
    -------------------------------------------------------------------------
    Execution Finished
    


```python
# Example with mode == add, adds an individual file to the training overview
```


```python
%run PyPersonalTrainer.py -m add -o ../Ueberblick_2019.xlsx -tf ../Trainingsdaten_2019  -t 19010601.gpx -sf ../Trainingsheets_2019
```

    -------------------------------------------------------------------------
    PyPersonalTrainer V 0.7.0
    Training analyzer for Polar HRM and GPX files
    (c) 2019 JDMORISE at YAHOO.COM
    -------------------------------------------------------------------------
    Session Sheet:  ../Trainingsheets_2019/19010601.xlsx  written
    -------------------------------------------------------------------------
    Execution Finished
    


```python
%run PyPersonalTrainer.py -l
```

    -------------------------------------------------------------------------
    PyPersonalTrainer V 0.7.0
    Training analyzer for Polar HRM and GPX files
    (c) 2019 JDMORISE at YAHOO.COM
    -------------------------------------------------------------------------
     
     This program is free software: you can redistribute it and/or modify
     it under the terms of the GNU General Public License as published by
     the Free Software Foundation, either version 3 of the License, or
     (at your option) any later version.
     
     This program is distributed in the hope that it will be useful,
     but WITHOUT ANY WARRANTY; without even the implied warranty of
     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
     GNU General Public License for more details.
     
     You should have received a copy of the GNU General Public License
     along with this program.  If not, see <https://www.gnu.org/licenses/>
     
    -------------------------------------------------------------------------
    


```python

```
