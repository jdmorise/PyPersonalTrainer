# PyPersonalTrainer
A personal training analyzer based on Python and Excel. 

As the PolarPersonalTrainer Webpage will be shut down by 31.12.2019, older Polar Fitness and GPS watches will be deprecated. This software enables to further use these older devices and creates relevant running training information in form of Excel WorkSheets. 
The software imports the training data in form HRM and GPX Files, creates training session worksheet for each training, and adds the training to an overview sheet. 

## Introduction 


## Requirements
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


