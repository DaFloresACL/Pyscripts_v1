import pandas as pd
import csv
import glob
import os
from os import path
import datetime
import pyodbc
import importlib.util
now = datetime.datetime.now()
timestamp = now.strftime('%Y%m%d_%H%M%S_')
ROOT_DIR = os.path.dirname(
    r'\\ac-hq-fs01\accounting\Finance\Underwriting\DataX\DataX Docs and Analytics\DataX Resources\Python\\')
dir = path.join(ROOT_DIR, 'DataX\\')
os.chdir(r'\\ac-hq-fs01\accounting\Finance\Underwriting\DataX\DataX Docs and Analytics\DataX Resources\Python\\')
dir
import cred

os.getcwd()

