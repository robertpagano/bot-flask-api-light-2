# %%
from docx import Document
from docx.shared import Inches
from docxcompose.composer import Composer

import requests
import unicodedata

import glob

## this work will need to be used if we need to use Python to grab files
## off sharepoint. For now, the plan is to do all of this through Flow, so no need to
## continue working on this.

# %%
## below is some scratch work on dynamically grabbing folders and shit

filepaths_test = []
for file in glob.glob('doc_builder_files/August/*'):
    filepaths_test.append(file)

filepaths_test

## so above grabs the content subfolders.
## I think what i can do is grab everything after '\\' and save that as a variable, 
## append it to a new list that has all section names.
## then i could change '\\' to '/', add a '\\*.docx' after content name,


# %%
def get_months(glob_path = 'doc_builder_files/*/'):
    # returns a list of months that exist in the subdirectory
    valid_months = []
    
    for subfolder in glob.glob(glob_path):
        valid_months.append(subfolder.split("\\", 1)[1][:-1])

    return valid_months

def get_articles_by_month(month):
    
    glob_path = 'doc_builder_files/REPLACE/*/'
    glob_path = glob_path.replace('REPLACE', month)
    
    return glob_path

def get_section_names(month):
    
    main_filepaths = []

    top_level = get_articles_by_month(month)
    #print(top_level)
    
    for subfolder in glob.glob(top_level):
        main_filepaths.append(subfolder)
    
    section_names = []

    for subfolder in main_filepaths:
        section_names.append(subfolder.split("\\", 1)[1][:-1])

    return section_names

def get_sections_all_months(glob_path = 'doc_builder_files/*/'):

    all_sections_dict = dict()

    for month in get_months(glob_path):
        all_sections_dict.update({month: get_section_names(month)})

    return all_sections_dict

# %%
get_section_names('August')

# %%
get_sections_all_months()

