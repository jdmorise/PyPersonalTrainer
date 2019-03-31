# PyPersonalTrainer
A personal training analyzer based on Python and Excel. 

As the PolarPersonalTrainer Webpage will be shut down by 31.12.2019, older Polar Fitness and GPS watches will be deprecated. This software enables to further use these older devices and creates relevant running training information in form of Excel WorkSheets. 
The software imports the training data in form HRM and GPX Files, creates training session worksheet for each training, and adds the training to an overview sheet. 

## Introduction 
This software aims to provide a similar (of course restricted) functionality as the website PolarPersonalTrainer by creating excel sheets with a similar look as the website. For each training session, an individual excel is created which contains the following information:    
1. A detailed overview table with
    1. start date
    2. start time
    3. duration
    4. distance 
    5. pace
    6. average heart rate
    7. maximum heart rate
    8. Polar running index
    9. Type of Training
2. A figure with the heart rate and speed over time
3. A table with time and percentage in zones of the puls rates
4. A table with the autolaps of the watch, taken each 1 km 
5. A table with userlaps, if recorded

The filename of the individual excel sheet is derived from the filenames of the hrm and gpx files. An example can be found under "19010601.xlsx". The data heartrate, pace, speed, and altitude versus time is stored in the same workbook in worksheet 'Data' to be used for further analysis. 

The software then adds the information on the overview table to a separate excel which should give an overview over the yearly training. 
It contains of a first sheet which gives overview over total duration, total distance, and number of training sessions for each month. 
The workbook contains for each month a table with the detailed overview table.    
The software als tries to detect if a training session was of type interval by analyzing variation in heart rate and speed and then adds this information to the worksheet 'Intervals'. An example of the Yearly overview excel workbook can be found under "Overview_2019.xlsx".   

## Requirements
The software was tested with training data from a Polar RC3 GPS with HRM version 1.06.   

Python 3.5.1 or higher   
argparse   
openpyxl   
gpxpy   
geopy   
datetime   

## Usage 
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
    

```
 ## Examples 
 
General example to scan a folder for updated training files, analyze the GPX and HRM data, createa dedicated training session workbook, and update the overview workbook: 

```python
python PyPersonalTrainer.py -m scan -o ../Ueberblick_2019.xlsx -tf ../Trainingsdaten_2019 -sf ../Trainingsheets_2019
```

A jupyter notebook with more detailed examples is included in the repo.    


