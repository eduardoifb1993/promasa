import pyodbc
import pandas as pd
import datetime as dt
import numpy as np
import xlsxwriter
from datetime import datetime, timedelta

a = pd.read_csv('e_202019.csv', sep = ';')

a.columns = map(str.lower,a.columns)
a = a.groupby(['w prod', 'ens', 'especie', 'modelo', 'panel', 'espesor',
       'ancho ensamble', 'largo', 'pro_description', 'bvl', 'ppc',
       'pid ensamble'])['total'].sum()
a = pd.DataFrame(a)       
a = a.unstack(level = 0)
a.sort_values(['especie'], ascending=True)

