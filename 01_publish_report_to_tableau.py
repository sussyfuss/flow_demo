# -*- coding: utf-8 -*-
"""
Created on Fri Apr 12 13:47:44 2019

@author: ckittrel

Publishes the COVID report onto the Tableau Server

"""

import os
import sys
from pathlib import Path # path of the root parent

import tableauserverclient as TSC # Tableau Server Client get from https://github.com/tableau/server-client-python
import tableau_config #server configuration

#load configuration
username = tableau_config.tableau_username
password = tableau_config.tableau_password
site_id = tableau_config.tableau_site_id
serverURL = tableau_config.tableau_server

tableau_auth = TSC.TableauAuth(username, password)
server = TSC.Server(serverURL)


def get_project_root() -> Path:
    """Returns project root folder."""
    return Path(__file__).parent.parent


def main():
    
    
    # change working directory to the file location
    os.chdir(os.path.dirname(os.path.abspath(__file__)))  
    #print(os.path.dirname(os.path.abspath(__file__)))
    
    # get the parent path
    path_parent = str(get_project_root())
    
    # add the project path to the script
    sys.path.insert(0, path_parent)
    
    
    with server.auth.sign_in(tableau_auth):
        # create a workbook item
        wb_item = TSC.WorkbookItem(name='<tableau report name>', project_id='xxxxx', show_tabs=True)
        
       
        # call the publish method with the workbook item
        wb_item = server.workbooks.publish(wb_item, path_parent + '/tableau/<report name>.twbx', 'Overwrite')
        #wb_item = server.workbooks.publish(wb_item, filepath, 'Overwrite')
# 
 
    # call the sign-out method with the auth object
    server.auth.sign_out()
    
    print("All done.")
        
if __name__ == '__main__':
    main()



    
    
    
