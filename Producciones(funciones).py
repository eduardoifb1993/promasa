import pyodbc
import pandas as pd
import datetime as dt
import numpy as np
import xlsxwriter
from datetime import datetime, timedelta

cnxn = pyodbc.connect('DSN=cpp;UID=mespinoza;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID=NTBK-SAP6;DATABASE=CPP;LANGUAGE=Español')
cursor = cnxn.cursor()
r = pd.read_sql("( SELECT PRODUCCIONES.TIPO_TARJETA, PRODUCCIONES.BODEGA, PRODUCCIONES.SECTOR,\
                                PRODUCCIONES.N_MAQUINA, PRODUCCIONES.VERSION, PRODUCCIONES.MAQUINA, PRODUCCIONES.TURNO,\
                                PRODUCCIONES.TARJETA, PRODUCCIONES.PIEZAS, PRODUCCIONES.M3_TARJETA,\
                                PRODUCCIONES.FECHA, PRODUCCIONES.PROCESO, PRODUCCIONES.TAG_ACTV,\
                                PRODUCCIONES.OBSERVACION_TARJETA, PRODUCCIONES.ESTADO,\
                                PRODUCCIONES.ID, PRODUCCIONES.CLASIFICACION, PRODUCCIONES.INV_CLAS,\
                                PRODUCCIONES.FAMILIA, PRODUCCIONES.NIVEL, PRODUCCIONES.TIPO,\
                                PRODUCCIONES.NOMBRE, PRODUCCIONES.ESPECIE, PRODUCCIONES.ESPESOR, PRODUCCIONES.ANCHO,\
                                PRODUCCIONES.LARGO, PRODUCCIONES.PERFIL, PRODUCCIONES.PANEL, PRODUCCIONES.STILE,\
                                PRODUCCIONES.DESCRIPCION, PRODUCCIONES.OBSERVACIONES, PRODUCCIONES.DETALLE,\
                                PRODUCCIONES.NEW_DETALLE, PRODUCCIONES.IMG, PRODUCCIONES.PO,\
                                PRODUCCIONES.PARTNUMBER, PRODUCCIONES.OPERADOR, PRODUCCIONES.SUPERVISOR,\
                                PRODUCCIONES.ETIQUETADOR, PRODUCCIONES.BODEGA_UBICACION, PRODUCCIONES.SECTOR_UBICACION,\
                                PRODUCCIONES.CALLE_UBICACION, PRODUCCIONES.OBSOLETO,\
                                PRODUCCIONES.BVL, PRODUCCIONES.mes \
                                FROM CPP.dbo.PRODUCCIONES PRODUCCIONES\
                                WHERE (PRODUCCIONES.FECHA>{ts '2020-04-01 00:00:00'}))", cnxn)
r.columns = map(str.lower,r.columns)
r = r[r.version != 0]
r = pd.DataFrame(r)
r['fecha_d_m'] = r.fecha.dt.strftime('%d/%m')


def produccionesxdia(r,dias): #
    df = pd.DataFrame()
    df1 = pd.DataFrame()
    
    #filtro por proceso
    r = r[r['proceso'].isin(['TERMINADO'])]
    
    #ciclo para los dias hacia atras de produccion
    for i in range(dias):
        PV = dt.datetime.today() - dt.timedelta(days=i)
        PV = PV.strftime('%d/%m')
        
        df = r[(r['fecha_d_m'] == PV )]\
                        .groupby(['familia','id','descripcion','espesor', 'ancho', 'largo','bvl'])['piezas'].sum()
        df = pd.DataFrame(df)
        df = df.rename(columns={"piezas": PV})
        df1 = pd.concat([df1, df], axis = 1).sort_index(axis=1)
        df1 = df1.replace(np.nan, 0 , regex=True)
    return df1

def prodxfamiliapuertas(r,dias): 
    df = pd.DataFrame()
    df1 = pd.DataFrame()
       
    #filtro por proceso y familia
    r = r[r['proceso'].isin(['TERMINADO'])]
    r = r[r['familia'].isin(['4 Y 6 PAN','FLAT PANEL','45 MM',\
                             'TDL','PLK','PUERTAS VARIABLE','LOUVER','FLAT PANEL','FULL LOUVER'])]
    
    #ciclo para los dias hacia atras de produccion
    for i in range(dias):
        PV = dt.datetime.today() - dt.timedelta(days=i)
        PV = PV.strftime('%d/%m')
        
        df = r[(r['fecha_d_m'] == PV )]\
                        .groupby(['familia','id','descripcion','espesor', 'ancho', 'largo','bvl'])['piezas'].sum()
        df = pd.DataFrame(df)
        df = df.rename(columns={"piezas": PV})
        df1 = pd.concat([df1, df], axis = 1).sort_index(axis=1)
        df1 = df1.replace(np.nan, 0 , regex=True)
    return df1


def prodxfamiliaexcel(r,dias): 
    df = pd.DataFrame()
    df1 = pd.DataFrame()
    writer = pd.ExcelWriter('prod_fam_term.xlsx', engine='xlsxwriter')
       
    #filtro por proceso y familia
    r = r[r['proceso'].isin(['TERMINADO'])]
    r = r[r['familia'].isin(['4 Y 6 PAN','FLAT PANEL','45 MM',\
                             'TDL','PLK','PUERTAS VARIABLE','LOUVER','FLAT PANEL','FULL LOUVER'])]
    
    #ciclo para los dias hacia atras de produccion
    for i in range(dias):
        PV = dt.datetime.today() - dt.timedelta(days=i)
        PV = PV.strftime('%d/%m')
        
        df = r[(r['fecha_d_m'] == PV )]\
                        .groupby(['familia','id','descripcion','espesor', 'ancho', 'largo','bvl'])['piezas'].sum()
        df = pd.DataFrame(df)
        df = df.rename(columns={"piezas": PV})
        df1 = pd.concat([df1, df], axis = 1).sort_index(axis=1)
        df1 = df1.replace(np.nan, 0 , regex=True)
    df2 = df1.reset_index()
    fam = df2.familia.unique()
    for j in range(len(fam)):
        df1.loc[fam[j]].to_excel(writer, sheet_name='{0}'.format(fam[j]))
    return writer.save()


def produccionesxdia_ASSEMBLY(dias):
    cnxn = pyodbc.connect('DSN=cpp;UID=mespinoza;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID=NTBK-SAP6;DATABASE=CPP;LANGUAGE=Español')
    cursor = cnxn.cursor()
    r = pd.read_sql("( SELECT PRODUCCIONES.TIPO_TARJETA, PRODUCCIONES.BODEGA, PRODUCCIONES.SECTOR,\
                                    PRODUCCIONES.N_MAQUINA, PRODUCCIONES.VERSION, PRODUCCIONES.MAQUINA, PRODUCCIONES.TURNO,\
                                    PRODUCCIONES.TARJETA, PRODUCCIONES.PIEZAS, PRODUCCIONES.M3_TARJETA,\
                                    PRODUCCIONES.FECHA, PRODUCCIONES.PROCESO, PRODUCCIONES.TAG_ACTV,\
                                    PRODUCCIONES.OBSERVACION_TARJETA, PRODUCCIONES.ESTADO,\
                                    PRODUCCIONES.ID, PRODUCCIONES.CLASIFICACION, PRODUCCIONES.INV_CLAS,\
                                    PRODUCCIONES.FAMILIA, PRODUCCIONES.NIVEL, PRODUCCIONES.TIPO,\
                                    PRODUCCIONES.NOMBRE, PRODUCCIONES.ESPECIE, PRODUCCIONES.ESPESOR, PRODUCCIONES.ANCHO,\
                                    PRODUCCIONES.LARGO, PRODUCCIONES.PERFIL, PRODUCCIONES.PANEL, PRODUCCIONES.STILE,\
                                    PRODUCCIONES.DESCRIPCION, PRODUCCIONES.OBSERVACIONES, PRODUCCIONES.DETALLE,\
                                    PRODUCCIONES.NEW_DETALLE, PRODUCCIONES.IMG, PRODUCCIONES.PO,\
                                    PRODUCCIONES.PARTNUMBER, PRODUCCIONES.OPERADOR, PRODUCCIONES.SUPERVISOR,\
                                    PRODUCCIONES.ETIQUETADOR, PRODUCCIONES.BODEGA_UBICACION, PRODUCCIONES.SECTOR_UBICACION,\
                                    PRODUCCIONES.CALLE_UBICACION, PRODUCCIONES.OBSOLETO,\
                                    PRODUCCIONES.BVL, PRODUCCIONES.mes \
                                    FROM CPP.dbo.PRODUCCIONES PRODUCCIONES\
                                    WHERE (PRODUCCIONES.FECHA>{ts '2020-04-01 00:00:00'}))", cnxn)
    r.columns = map(str.lower,r.columns)
    r = r[r.version != 0]
    r = pd.DataFrame(r)
    r['fecha_d_m'] = r.fecha.dt.strftime('%d/%m')

    df = pd.DataFrame()
    df1 = pd.DataFrame()
    
    #filtro por proceso
    #r = r[r['proceso'].isin(['LIJADO'])]
    r = r[r['turno'].isin(['1','2'])]
    #r = r[r['familia'].isin(['LOUVER'])]
    r = r[r['clasificacion'].isin(['PUERTAS'])]
    #r = r[r['estado'].isin(['A'])]
    r = r[r['maquina'].isin(['AUTO PRESS','MANUAL PRESS 2', 'POST. APLICACION',\
                             'MANUAL PRESS ' , 'MANUAL PRESS 3','MANUAL PRESS 4',\
                             'HINGE ASMB', 'MANUAL PRESS 5 (SUF)','EMBALAJE','LIJADORA','LIJADORA 2','SHRINKWRAP'])]
    
    #r = r[r['clasificacion'].isin(['PUERTAS'])]
    
    #ciclo para los dias hacia atras de produccion
    for i in range(dias):
        PV = dt.datetime.today() - dt.timedelta(days=i)
        PV = PV.strftime('%d/%m')
        
        df = r[(r['fecha_d_m'] == PV )]\
                        .groupby(['maquina','id','descripcion','espesor', 'ancho', 'largo','bvl'])['piezas'].sum()
        df = pd.DataFrame(df)
        df = df.rename(columns={"piezas": PV})
        df1 = pd.concat([df1, df], axis = 1).sort_index(axis=1)
        df1 = df1.replace(np.nan, 0 , regex=True)
    return df1

writer = pd.ExcelWriter('C:/Users/efernandez/Desktop/Assembly.xlsx', engine='xlsxwriter')
produccionesxdia_ASSEMBLY(1).to_excel(writer, sheet_name='Assemmbly')
writer.save()

