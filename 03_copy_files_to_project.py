# -*- coding: utf-8 -*-
"""

@author: ckittrel
@script: Copies the remote file sets over to this project for publishing into Tableau Server
"""

import subprocess
import sys
import os 
import inspect # getting the file path
#import distutils.dir_util # for copying files
from pathlib import Path # path of the root parent

# get the current file
src_file_path = inspect.getfile(lambda: None)



def get_project_root() -> Path:
    """Returns project root folder."""
    return Path(src_file_path).parent.parent

def join(iterator, separator):
    it = map(str, iterator)
    separator = str(separator)
    string = next(it, '')
    for s in it:
        string += separator + s
    return string

def robocopyDHA(path_parent):
    
    # https://python.readthedocs.io/en/latest/library/subprocess.html
    #subprocess.run(r'C:\Windows\System32\robocopy.exe "<source folder>" <destination folder> /s /xo', shell=True, check=True)
    
    path_exe = r'C:\Windows\System32\robocopy.exe'
    path_source = r'<source folder>'
    path_dest = path_parent + r'\ext_data_dha'
    arg = r'/s /xo'
    
    cmd = join([path_exe, path_source, path_dest, arg], ' ')

    ###########################
    ##### Jazz hands here #####
    ###########################
    subprocess.run(cmd, shell=True, check=False)


def main():
    
    # change to current directory path
    os.chdir(os.path.dirname(os.path.abspath(src_file_path)))  
    
    # get the parent path
    path_parent = str(get_project_root())
    os.system("cd " + path_parent)
    print(path_parent)
    
    # add the project path to the script
    sys.path.insert(0, path_parent)
    
    # copy the files to local folder
    robocopyDHA(path_parent)
    

        
if __name__ == '__main__':
    main()

    


