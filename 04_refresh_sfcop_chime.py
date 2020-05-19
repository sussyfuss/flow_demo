# -*- coding: utf-8 -*-
"""
Created on Thu Apr 16 10:03:50 2020

@author: ckittrel
@backup datasource:  Check the TFL for the data source.   

@note: This process will update the CHIME data source covid19_chime on the tableau server via the tableau flow.
"""

#####################
##### Libraries #####
#####################

import sys
import os
from pathlib import Path # path of the root parent


#####################
##### Functions #####
#####################

def get_project_root() -> Path:
    """Returns project root folder."""
    return Path(__file__).parent.parent
    
def refreshTFL(credentials, flow):
    
    # location of the CLI for Tableau Prep
    path_exec = 'C:\\"Program Files"\\Tableau\\"Tableau Prep Builder 2019.4"\\scripts\\tableau-prep-cli.bat'
    
    # Location of the credentials json 
    path_credentials = " -c " + credentials
    path_flow = " -t " + flow
    command = path_exec + path_credentials + path_flow  
    
    # execute
    p = os.popen(command)
    
    # print out results to the console
    print(p.read())
    
################    
##### Main #####    
################
    
def main():

    # change to current directory path
    os.chdir(os.path.dirname(os.path.abspath(__file__)))  

    # get the parent path
    path_parent = str(get_project_root())
    
    # add the project path to the script
    sys.path.insert(0, path_parent)
    

    # where the credentials.json is located
    path_credentials =  path_parent + "\\tableau\\credentials.json"
    # where the prep tfl is located
    path_flow =  path_parent + "\\tableau\\etl_extract_and_publish_chime_data_to_tableau_v5.tfl"

    
    # publish the data over to the Tableau server
    refreshTFL(path_credentials, path_flow)
    
     # notify script user of the status of the process
    print("Data load to tableau complete!")   
    

if __name__ == '__main__':
    main()
