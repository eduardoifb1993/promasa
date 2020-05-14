import pyodbc
import pandas as pd
import datetime as dt
import numpy as np
import xlsxwriter
from datetime import datetime, timedelta
import matplotlib.pyplot as plt

from matplotlib.ticker import FuncFormatter
import seaborn as sns

cierre = pd.read_excel('C:/Users/efernandez/Desktop/CIERRE ÓRDENES SEM 19-2020.xlsx', sheet_name='PROGRAMA',header = 1,
                                  usecols=['SEM RE','PO','ID','CUSTOMER','QTY','X ENS','ENS','STOCK','X EMBALAR','ENSAMBLE','linea','X PINTAR'])


def xsemana(cierre, semana):
    cierre = cierre[cierre['ENSAMBLE'].isin(['A','M'])]
    cierre = cierre[cierre['SEM RE'].isin([semana])]
    cierre = cierre.replace(np.nan, 0 , regex=True)
    cierre.PO = cierre.PO.astype(str)
    cierre = cierre.groupby(['PO','CUSTOMER','linea'])['QTY','X ENS','ENS','STOCK','X PINTAR','X EMBALAR'].sum()
    return cierre

def resumen(cierre, semana):
    cierre = cierre[cierre['ENSAMBLE'].isin(['A','M'])]
    cierre = cierre[cierre['SEM RE'].isin([semana])]
    cierre = cierre.replace(np.nan, 0 , regex=True)
    cierre.PO = cierre.PO.astype(str)
    cierre = cierre.groupby(['PO','CUSTOMER','linea'])['QTY','X ENS','ENS','STOCK','X PINTAR','X EMBALAR'].sum()
    cierre = cierre.sum(level=['PO','CUSTOMER'])
    resumen = cierre.drop(columns = ['QTY','ENS','STOCK'])
    return resumen

writer = pd.ExcelWriter('Report_cierre_ordenes.xlsx', engine='xlsxwriter')
resumen(cierre, '19').to_excel(writer, sheet_name='Report_19')
resumen(cierre, '19').sum().to_excel(writer, sheet_name='Report_19S_')
resumen(cierre, '20').to_excel(writer, sheet_name='Report_20')
resumen(cierre, '20').sum().to_excel(writer, sheet_name='Report_20S_')
writer.save()




xsemana(cierre,'19').drop(columns = ['QTY','ENS','STOCK']).sum(axis = 1)


#fig, ax = plt.subplots()
#ax.barh(group_names, group_data)
#plt.leyend('asndlsad')

plt.plot()

plt.barh(resumen(cierre, '19').reset_index()['X ENS'])
#############
sns.set(style="whitegrid")

# Load the example Titanic dataset
titanic = sns.load_dataset(resumen(cierre, '19'))

# Draw a nested barplot to show survival for class and sex
g = sns.catplot(x="PO", y=["X ENS",'X PINTAR'], hue="CUSTOMER", data = resumen(cierre, '19').reset_index(),
                height=6, kind="bar", palette="muted")
g.despine(left=True)
g.set_ylabels("survival probability")

titanic = sns.load_dataset("titanic")

sns.catplot(y="CUSTOMER", hue="CUSTOMER", kind="count",
            palette="pastel", edgecolor=".6",
            data=carga)

