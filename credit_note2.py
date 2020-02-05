import re

import pandas as pd
import os.path

import datetime

import pdfplumber
from PyPDF2 import PdfFileReader, PdfFileWriter
from win32com.client import Dispatch
from datetime import datetime
import fitz
from os import walk
from os.path import join
import comtypes.client


xl = Dispatch("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
which_quarter = input('西元+季:')
as_at = input('e.g.Sep 30,2019')
soa_period = input('e.g.201909N1')
AC_period = input('e.g. Apr.01, 2019 to Jun.30, 2019')
save_path = 'M:\\季帳\\{}\\starr\\CR'.format(which_quarter)
wbs_path = 'D:\\creditnote.xlsx'

