'''
Python boilerplate filler takes a spread sheet of information and a word document for each row in the spreadsheet using
the data provided in that row to fill the template word document
uses the pandas library for working with excel and minidom for editing the docx xml
'''

from pandas import read_excel
import os
import sys
import inspect
import shutil
import zipfile
import datetime
from xml.dom import minidom

# temp excel and word file names
docx_file = 'test document.docx'
excel_file = 'test spreadsheet.xlsx'


# returns the a list of dicts for each finish doc to be created
def get_excel_data(excel_file):
    # reads in the excel spreadsheet and gets the column names which will be used to find what to replace
    df = read_excel(excel_file)
    column_names = df.columns

    # format the excel data into a list of dictionaries
    variable_list = []
    for i in df.index:
        row = {}
        for j in range(len(column_names)):
            row[column_names[j]] = str(df[column_names[j]][i])
        variable_list.append(row)

    return variable_list


# copies and unzips a docx file to edit xml
def copy_unzip_docx(docx_file):
    # want to unzip to temp folder of this program so it can be cleaned later
    folderpath = os.path.dirname(os.path.abspath(
        inspect.getfile(inspect.currentframe()))) + '/tmp/'

    # create the tmp folder if it doesn't exist and then copy
    if not(os.path.exists(folderpath)):
        os.mkdir(folderpath)

    # unzip the docx
    with zipfile.ZipFile(docx_file, 'r') as zip_ref:
        zip_ref.extractall(folderpath + '/template/')


def zip_and_save_docx(filename):
    folder_path = os.path.dirname(os.path.abspath(
        inspect.getfile(inspect.currentframe())))
    x = datetime.datetime.now()
    input_path = folder_path + '/tmp/'
    output_path = folder_path + '/output/'  # + x.strftime("%c") + '/'

    # create the output folder if it doesn't exist and then copy
    if not(os.path.exists(output_path)):
        print(output_path)
        os.mkdir(output_path)

    shutil.make_archive(output_path + filename, 'zip',
                        input_path + '/template/')

    # rename the file to .docx
    os.rename(output_path + filename + '.zip',
              output_path + filename + '.docx')


# clean the tmp directory after work is finished
def clean_tmp():
    shutil.rmtree(os.path.dirname(os.path.abspath(
        inspect.getfile(inspect.currentframe()))) + '/tmp/')


# searches the open template for text to replace and replaces it based on the replace dict
def fill_boilerplate(replace_dict):
    # verify tmp/template/word/document.xml exists
    filepath = os.path.dirname(os.path.abspath(
        inspect.getfile(inspect.currentframe()))) + '/tmp/template/word/document.xml'
    if not(os.path.exists(filepath)):
        raise Exception(
            'document.xml does not exists in the template folder!!!')

    # open the xml file
    template = minidom.parse(filepath)
    # get all the text elements
    template_text = template.getElementsByTagName('w:t')

    # check each element for a replacement form
    for text in template_text:
        # word creates empty text elements for styling so just ignore all the empty ones
        if text.firstChild != None:
            new_text = text.firstChild.data
            for key in replace_dict:
                new_text = new_text.replace(
                    '/' + key + '/', replace_dict[key])
            text.firstChild.data = new_text

    # save the document afterwards with changes
    with open(filepath, 'w') as f:
        f.write(template.toxml())


def main(docx_file, excel_file):
    replace_dicts = get_excel_data(excel_file)
    i = 1
    for replace_dict in replace_dicts:
        copy_unzip_docx(docx_file)
        fill_boilerplate(replace_dict)
        zip_and_save_docx('test_' + str(i))
        i += 1

    clean_tmp()


if __name__ == '__main__':
    if len(sys.argv) == 3:
        main(sys.argv[1], sys.argv[2])
    else:
        main(docx_file, excel_file)
