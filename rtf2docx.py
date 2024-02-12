# import libraries
import streamlit as st
import tkinter as tk
from tkinter import filedialog
import glob
import re
import os
import win32com.client as win32
from win32com.client import constants
from io import BytesIO
import pythoncom
pythoncom.CoInitialize()
import pandas as pd

# Set up tkinter
root = tk.Tk()
root.withdraw()

# Make folder picker dialog appear on top of other windows
root.wm_attributes('-topmost', 1)

#Folder picker button
st.title('Конвертер RTF в DOC')
st.write('Пожалуйста, выберите папку с вашими файлами формата RTF:')
clicked_path_to_folder = st.button('Выбрать папку:')

if 'folder_selected' not in st.session_state:
    st.session_state['folder_selected'] = False

if 'files_list' not in st.session_state:
    st.session_state['files_list'] = None

if clicked_path_to_folder:
    dirname = st.text_input('Выбрать папку:', filedialog.askdirectory(master=root))
    
    path = dirname
    pattern = '*.rtf'
    files = glob.glob(f'{path}/{pattern}')
    
    st.session_state['files_list'] = files

    if len(files):
        st.info(f'Найдено файлов RTF: {len(files)}', icon="ℹ")
        st.session_state['folder_selected']  = True
    else:
        st.warning(f'RTF файлы не найдены', icon="⚠️")




if st.session_state['folder_selected']:
    clicked_begin_convert = st.button('Начать конвертацию')

    if clicked_begin_convert:
        files = st.session_state['files_list'] 
        progress_text = "Пожалуйста, подождите"
        my_bar = st.progress(0, text=progress_text)
        i = 0
        part = 1 / len(files)
        for file_path in files:
            word = win32.gencache.EnsureDispatch('Word.Application')
            doc = word.Documents.Open(file_path)
            doc.Activate()
            # Rename path with .doc
            new_file_abs = os.path.abspath(file_path)
            new_file_abs = re.sub(r'\.\w+$', '.doc', new_file_abs)
            # Save and Close
            word.ActiveDocument.SaveAs(
                new_file_abs, FileFormat=constants.wdFormatDocument
            )
            doc.Close(False)
            i+=1 #Просто счётчик
            my_bar.progress(i*part, text=progress_text)
        st.write('Конвертация окончена')