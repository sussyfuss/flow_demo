# -*- coding: utf-8 -*-
"""
Created on Thu Apr 20 10:03:50 2020

@author: ckittrel
@purpose:  This refreshes the bsr_perstats.xlsx data tables that connect to SharePoint.  The data is loaded
           into the tableau server with a later step.

@note: This process will open an excel application where you will need to enter your CAC email certificate.  
       Once logged in, the data tables in the excel
       document will refresh and save.   The next step is to publish the data to the tableau server.
"""

#####################
##### Libraries #####
#####################
import sys

import win32com.client # for Excel
import time

import os
from pathlib import Path # path of the root parent
#import inspect # getting the file path
import pandas as pd  # data munging
pd.options.mode.chained_assignment = None  # default='warn'
from datetime import datetime # Current date time in local system
import subprocess # calling the Prep Flow

#################
##### Setup #####
#################


# get the current file
#src_file_path = inspect.getfile(lambda: None)

# current path to the file
curr_path = os.path.dirname(os.path.abspath(__file__))
# project source path
path_parent = str(Path(curr_path).parent.absolute())



#####################
##### Functions #####
#####################

#def get_project_root() -> Path:
#     """Returns project root folder."""
#     return Path(src_file_path).parent.parent
     

def refresh(path):
    xlapp = win32com.client.DispatchEx("Excel.Application")
    wb = xlapp.Workbooks.Open(path)
    
    # Keep it visible so you can sign into carepoint
    xlapp.Visible = True
    wb.RefreshAll()
    
    time.sleep(5)
    xlapp.DisplayAlerts = False
    wb.Save()
    xlapp.Quit()
    
def impExcel(path, sheet):
    
    df = pd.read_excel(open(path, 'rb'),
              sheet_name=sheet,
              dtype=str) 
    return(df)

def join(iterator, separator):
    it = map(str, iterator)
    separator = str(separator)
    string = next(it, '')
    for s in it:
        string += separator + s
    return string    
  
def publishHyper(path_parent):
    
    # https://python.readthedocs.io/en/latest/library/subprocess.html
    #subprocess.run(r'C:\Windows\System32\robocopy.exe "<source folder>" <destination folder> /s /xo', shell=True, check=True)
    
    path_cli = r'C:\"Program Files"\Tableau\"Tableau Prep Builder 2019.4"\scripts\tableau-prep-cli.bat '
    path_cred = r'-c C:\Users\ckittrel\Documents\projects\etl_extract_covid_data\tableau\credentials.json'
    path_tfl = r'-t C:\Users\ckittrel\Documents\projects\etl_extract_covid_data\tableau\etl_extract_and_publish_covid19_bsr_perstat_timeseries5.tfl'
        
    cmd = join([path_cli, path_cred, path_tfl], ' ')

    ###########################
    ##### Jazz hands here #####
    ###########################
    subprocess.run(cmd, shell=True, check=False)    
    
################    
##### Main #####    
################
    
    
def main():

    # create a global to return it to the main env
    global bsr, perstat # base status report, personnel status report
    global ref_facility, ref_site_service, ref_dmis # reference 
    global perstat2 # modified perstat daily load file
    global master_bsr  # master bsr that is appended daily
    global master_perstat # master perstat that is appended daily
    global master_bsr_filtered # master bsr filtered for today
    global master_perstat_filtered # master perstat filtered for today
    global master_bsr_new, master_perstat_new # newly appended master dataframes
      
    ######################################
    ##### Setup and Folder Structure #####
    ######################################

    
    # add the project path to the script
    sys.path.insert(0, path_parent)
    # where the xlsx is located
    path_source = path_parent +  "\\ext_bsr_perstat\\bsr_perstat.xlsx"
    
    # master data files for Prep
    f_master_bsr = path_parent + "\\data4prep\\bsr_master.csv"
    f_master_perstat = path_parent + "\\data4prep\\perstat_master.csv"
    
    
    # get today's date
    today = datetime.date(datetime.now()).strftime("%m/%d/%Y")
    
    # archive path of the bsr and perstat backups
    bsr_archive_name = path_parent + "\\data4prep\\backup\\bsr_" + str(datetime.date(datetime.now()).strftime("%m-%d-%Y")) + ".csv"
    perstat_archive_name = path_parent + "\\data4prep\\backup\\perstat_" + str(datetime.date(datetime.now()).strftime("%m-%d-%Y")) + ".csv"
    bsr_master_archive_name = path_parent + "\\data4prep\\backup\\bsr_master_" + str(datetime.date(datetime.now()).strftime("%m-%d-%Y")) + ".csv"
    perstat_master_archive_name = path_parent + "\\data4prep\\backup\\perstat_master_" + str(datetime.date(datetime.now()).strftime("%m-%d-%Y")) + ".csv"
    
    # secondary archiving
    path_sg9_archive = "<archive site>"
 
    sg9_bsr_archive_name = path_sg9_archive + "\\bsr_" + str(datetime.date(datetime.now()).strftime("%m-%d-%Y")) + ".csv"
    sg9_perstat_archive_name = path_sg9_archive + "\\perstat_" + str(datetime.date(datetime.now()).strftime("%m-%d-%Y")) + ".csv"
    sg9_bsr_master_archive_name = path_sg9_archive + "\\bsr_master_" + str(datetime.date(datetime.now()).strftime("%m-%d-%Y")) + ".csv"
    sg9_perstat_master_archive_name = path_sg9_archive + "\\perstat_master_" + str(datetime.date(datetime.now()).strftime("%m-%d-%Y")) + ".csv"
    
    sg9_bsr_master_archive_master_name = path_sg9_archive + "\\bsr_master.csv"
    sg9_perstat_master_archive_master_name = path_sg9_archive + "\\perstat_master.csv"
    
    
    ########################################################
    ##### refresh the workbook data using the win32com #####
    ########################################################
    
    # you'll be prompted for you CAC in Excel - WATCH FOR THIS
    refresh(path_source)

    ##################################################
    ##### import the excel data into data frames #####
    ##################################################
    
    print ("Preparing data...")
    
    bsr = impExcel(path_source, 'bsr_owssvr')
    perstat = impExcel(path_source, 'perstat_owssvr')
    ref_facility = impExcel(path_source, 'ref_facility_list')
    ref_site_service = impExcel(path_source, 'ref_site_service')
    ref_dmis = impExcel(path_source, 'ref_dmis')

    # add the column dmis_id
    perstat['dmis_id'] = perstat['DMIS/MTF'].str[:4]

    ##################################
    ##### Add Reference Metadata #####
    ##################################
    
    # ref_facility 
    perstat2 = pd.merge(perstat,
                        ref_facility[['dmis_id', 
                                      'facility_name',
                                      'Inpatient or Outpatient'
                                      ]],
               left_on='dmis_id', right_on='dmis_id', how='left')

    # ref_dmis
    perstat2 = pd.merge(perstat2,
                        ref_dmis[['dmis_id',
                                  'Market',
                                  'Region Name',
                                  'FEMA Region Number',
                                  'COCOM',
                                  'Country Code',
                                  'MAJCOM',
                                  'Site Service']],
               left_on='dmis_id', right_on='dmis_id', how='left')

    # ref_site_service
    perstat2 = pd.merge(perstat2,
                    ref_site_service[['Site Service',
                              'Site Service Name']],
           left_on='Site Service', right_on='Site Service', how='left')


    # add a time stamp
    bsr['refreshDt'] = pd.to_datetime('today').strftime("%m/%d/%Y")
    perstat2['refreshDt'] = pd.to_datetime('today').strftime("%m/%d/%Y")

    ###############################
    ##### Mutate Master Files #####
    ###############################
    
    master_bsr = pd.read_csv(f_master_bsr, dtype=str)
    master_perstat = pd.read_csv(f_master_perstat, dtype=str)


    # filter for today's data in case it's already there
    master_bsr_filtered = master_bsr[master_bsr['refreshDt'] != today]
    master_perstat_filtered = master_perstat[master_perstat['refreshDt'] != today]

    # append the new data to the master tables
    master_bsr_new = master_bsr_filtered.append(bsr, ignore_index=True, sort=False)
    master_perstat_new = master_perstat_filtered.append(perstat2, ignore_index=True, sort=False)

    # fix leading zero issues
    master_bsr_filtered['Tmt DMIS ID'] = master_bsr_filtered['Tmt DMIS ID'].map(lambda x: f'{x:0>4}')
    master_perstat_filtered['dmis_id'] = master_perstat_filtered['dmis_id'].map(lambda x: f'{x:0>4}')
    master_bsr_new['Tmt DMIS ID'] = master_bsr_new['Tmt DMIS ID'].map(lambda x: f'{x:0>4}')
    master_perstat_new['dmis_id'] = master_perstat_new['dmis_id'].map(lambda x: f'{x:0>4}')

    ###########################################
    ##### Export Master Files back to CSV #####
    ###########################################
    
    print("Exporting to CSV...")
    
    # back up today's files
    bsr.to_csv(bsr_archive_name, index=False)
    perstat2.to_csv(perstat_archive_name, index=False)
    
    # export date stamped files to project
    master_bsr_new.to_csv(bsr_master_archive_name, index=False)
    master_perstat_new.to_csv(perstat_master_archive_name, index=False)   

    # the files without the date are used for Tableau Prep
    master_bsr_new.to_csv(f_master_bsr, index=False)
    master_perstat_new.to_csv(f_master_perstat, index=False)
 
    # SG9 ARCHIVE - back up today's files
    bsr.to_csv(sg9_bsr_archive_name, index=False)
    perstat2.to_csv(sg9_perstat_archive_name, index=False)
    
    # SG9 ARCHIVE - export date stamped files to project
    master_bsr_new.to_csv(sg9_bsr_master_archive_name, index=False)
    master_perstat_new.to_csv(sg9_perstat_master_archive_name, index=False)   
       
    # the files without the date are used for Tableau Prep (loaded to the network drive for tfl v3)
    master_bsr_new.to_csv(sg9_bsr_master_archive_master_name, index=False)
    master_perstat_new.to_csv(sg9_perstat_master_archive_master_name, index=False)
    

    #####################################
    ##### Execute Tableau Prep Flow #####
    #####################################
    
    print("Calling Flow...")
    publishHyper(path_source)
   

    # notify script user of the status of the process
    print("Process update is complete...")
    

# Run main() function
main()
