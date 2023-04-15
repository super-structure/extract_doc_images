"""This script will extract images as PNGs from a Word document. 

This script requires that Wordconv.exe be installed with MS Office.
The filepath to this program may vary.

https://stackoverflow.com/questions/2405417/automation-how-to-automate-transforming-doc-to-docx/2405508#2405508
 
Copyright Ⓒ Bentley Systems, Incorporated. All rights reserved.
"""
# standard libraries
import os
from pathlib import Path
import shutil
import sys
import subprocess
import uuid
import zipfile
 
# third-party libraries
from halo import Halo
from lxml import etree

VERSION = '1'
LANG = 'en'
PREFIX = 'c-re'

def main(arg_values):
    word_on_path()
    count = 0
    for count, value in enumerate(arg_values[1:], start=1):
        word_file_path = Path(value)
        print(f'Input file: {word_file_path.name} in {word_file_path.parent}')
        temp_file_created = False
        if word_file_path.suffix == '.doc':
            word_file_path = convert_to_docx(word_file_path)
            temp_file_created = True
        zip_file_name = Path(word_file_path.parent, f'{word_file_path.stem}.zip')
        export_directory = Path(word_file_path.parent, word_file_path.stem)
        os.rename(word_file_path, zip_file_name)
        docx_archive = zipfile.ZipFile(zip_file_name)
        for file in docx_archive.namelist():
            if file.startswith('word/media/'):
                docx_archive.extract(file, export_directory)
        flatten(export_directory)

        # rename the image files?
        for image_num, file in enumerate(export_directory.glob('**/*')): #os.walk(str(export_directory), topdown=False)
            id = ish_guid()
            ext = Path(file).suffix
            filename = f'{PREFIX} {word_file_path.stem}{image_num+1}={id}={VERSION}={LANG}=low{ext}'
            filepath = Path(export_directory, filename)
            os.rename(file, filepath)
            print(f'\t{filename}')
        docx_archive.close()
        if temp_file_created == True:
            zip_file_name.unlink()
        else:
            os.rename(zip_file_name, word_file_path)
    print(f'='*64)
    print(f'{count} Word documents processed.')
    os.system("pause")

def word_on_path():
    office16_path = Path(f"{os.environ['ProgramFiles']}",'Microsoft Office','root','Office16')
    os.environ["PATH"] += str(office16_path) + os.pathsep
    print(shutil.which('Wordconv.exe'))
    print(f'='*64)
    pass


def convert_to_docx(word_doc_file:Path):
    docx_file = Path(word_doc_file.parent, f'{word_doc_file.stem}.docx')
    if docx_file.exists() is True:
        docx_file.unlink()
    with Halo(text=f'\tConverting {word_doc_file.name} to .docx', spinner='pipe') as spinner:
        subprocess.run(['Wordconv', '-oice', '-nme', word_doc_file, docx_file])
        spinner.succeed(f'\tConverting: {word_doc_file.name} to .docx\t...Complete!')
    return docx_file


def get_author(word_xml_file:Path):
    w_root = etree.parse(word_xml_file).getroot()
    authors = w_root.xpath('.//o:Author/text()',
                               namespaces = {'o':'urn:schemas-microsoft-com:office:office'})
    print(f'\tAuthors: "{authors[0]}"')
    return authors[0]


def get_topic_title(word_xml_file:Path):
    w_root = etree.parse(word_xml_file).getroot()
    topic_title = w_root.xpath('.//w:p[child::w:pPr/w:pStyle[@w:val="Heading1"]]/w:r/w:t/text()',
                               namespaces = {'w':'http://schemas.microsoft.com/office/word/2003/wordml'})
    try:
        topic_title_str = ' '.join(topic_title)
    except:
        topic_title_str = ''
    print(f'\tTopic title: "{topic_title_str}"')
    return topic_title_str


def ish_guid():
    return f"GUID-{str(uuid.uuid4()).upper()}"


def flatten(directory):
    """Recursively flattens all the files in a dirctory tree
    
    source: https://amitd.co/code/python/flatten-a-directory

    Args:
      directory: the directory to be flattened

    """
    directory = str(directory)  # in case it's a pathlib Path
    for dirpath, _, filenames in os.walk(directory, topdown=False):
        for filename in filenames:
            i = 0
            source = os.path.join(dirpath, filename)
            target = os.path.join(directory, filename)

            while os.path.exists(target):
                i += 1
                file_parts = os.path.splitext(os.path.basename(filename))

                target = os.path.join(
                    directory,
                    file_parts[0] + "_" + str(i) + file_parts[1],
                )

            shutil.move(source, target)

        if dirpath != directory:
            os.rmdir(dirpath)


# ============================================================================================================================
if __name__ == '__main__':
    main(sys.argv)
