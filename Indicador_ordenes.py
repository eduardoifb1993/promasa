import pyodbc
import pandas as pd
import datetime as dt
import numpy as np
import xlsxwriter

from datetime import datetime, timedelta

cnxn = pyodbc.connect('DSN=cpp;UID=mespinoza;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID=NTBK-SAP6;DATABASE=CPP;LANGUAGE=Español')
cursor = cnxn.cursor()

s = pd.read_sql("SELECT STOCK.TIPO_TARJETA, STOCK.BODEGA, STOCK.SECTOR, STOCK.N_MAQUINA,\
                STOCK.MAQUINA, STOCK.TURNO, STOCK.TARJETA, STOCK.VERSION,\
                STOCK.PIEZAS, STOCK.M3_TARJETA, STOCK.FECHA, STOCK.PROCESO,\
                STOCK.TAG_ACTV, STOCK.OBSERVACION_TARJETA, STOCK.ESTADO, STOCK.ID,\
                STOCK.CLASIFICACION, STOCK.INV_CLAS, STOCK.FAMILIA, STOCK.NIVEL,\
                STOCK.TIPO, STOCK.NOMBRE, STOCK.ESPECIE, STOCK.ESPESOR, STOCK.ANCHO,\
                STOCK.LARGO, STOCK.PERFIL, STOCK.PANEL, STOCK.VIDRIO, STOCK.STILE,\
                STOCK.DESCRIPCION, STOCK.OBSERVACIONES, STOCK.PO, STOCK.PARTNUMBER,\
                STOCK.ETIQUETADOR, STOCK.BODEGA_UBICACION, STOCK.SECTOR_UBICACION,\
                STOCK.CALLE_UBICACION, STOCK.OBSOLETO, STOCK.BVL, STOCK.MN, STOCK.DIF, STOCK.KG\
                FROM CPP.dbo.STOCK STOCK", cnxn) #STOCK DE TODO! A LA FECHA

rc = pd.read_sql(" SELECT RECETAS.PID, RECETAS.NIVEL, RECETAS.CLASIFICACION,\
                        RECETAS.FAMILIA, RECETAS.TIPO, RECETAS.NOMBRE, RECETAS.DESCRIPCION, RECETAS.ESPESOR,\
                        RECETAS.ANCHO, RECETAS.LARGO, RECETAS.ESPECIE, RECETAS.PERFIL, RECETAS.PANEL, RECETAS.VIDRIO, RECETAS.STILE, RECETAS.ITEM, RECETAS.DETALLE, RECETAS.PID_H, RECETAS.CLASIFICACION_H, RECETAS.FAMILIA_H, RECETAS.TIPO_H, RECETAS.NOMBRE_H, RECETAS.DESCRIPCION_H, RECETAS.ESPESOR_H, RECETAS.ANCHO_H, RECETAS.LARGO_H, RECETAS.ESPECIE_H, RECETAS.PERFIL_H, RECETAS.ITEM_H, RECETAS.DETALLE_H, RECETAS.QTY, RECETAS.PROCESO, RECETAS.M3, RECETAS.REG, RECETAS.REC, RECETAS.FORMATO, RECETAS.UOM, RECETAS.COSTO, RECETAS.CREACION, RECETAS.MODIFICACION\
                        FROM CPP.dbo.RECETAS RECETAS ", cnxn) #recetas completas

rc.columns = map(str.lower,rc.columns)
rc = rc[rc['familia_h'].isin(['BARRAS', 'VIDRIOS', 'STILE', 'STILE LOUVER', 'PANELES','SLATS',
                                    'RAILS', 'RAILS LOUVER', 'APLICACIONES',
                                    'MOLDURAS SOLIDAS', 'MARCOS', 'OTRAS', 'MOLDURAS FINGER'])]


cierre = pd.read_excel('C:/Users/efernandez/Desktop/CIERRE ÓRDENES SEM 19-2020.xlsx', sheet_name='PROGRAMA',header = 1,
                                  usecols=['SEM RE','PO','ID','CUSTOMER','QTY','X ENS','ENS',
                                            'STOCK','X EMBALAR','ENSAMBLE','linea','X PINTAR'])

def xsemana(cierre, semana):
    cierre = cierre[cierre['ENSAMBLE'].isin(['A','M'])]
    cierre = cierre[cierre['SEM RE'].isin([semana])]
    cierre = cierre.replace(np.nan, 0 , regex=True)
    cierre.PO = cierre.PO.astype(str)
    cierre = cierre.groupby(['PO','CUSTOMER','linea'])['QTY','X ENS','ENS','STOCK','X PINTAR','X EMBALAR'].sum()
    return cierre

def valida_general(rc,s):
   #consultas SQL
   s = pd.read_sql("SELECT STOCK.TIPO_TARJETA, STOCK.BODEGA, STOCK.SECTOR, STOCK.N_MAQUINA,\
                STOCK.MAQUINA, STOCK.TURNO, STOCK.TARJETA, STOCK.VERSION,\
                STOCK.PIEZAS, STOCK.M3_TARJETA, STOCK.FECHA, STOCK.PROCESO,\
                STOCK.TAG_ACTV, STOCK.OBSERVACION_TARJETA, STOCK.ESTADO, STOCK.ID,\
                STOCK.CLASIFICACION, STOCK.INV_CLAS, STOCK.FAMILIA, STOCK.NIVEL,\
                STOCK.TIPO, STOCK.NOMBRE, STOCK.ESPECIE, STOCK.ESPESOR, STOCK.ANCHO,\
                STOCK.LARGO, STOCK.PERFIL, STOCK.PANEL, STOCK.VIDRIO, STOCK.STILE,\
                STOCK.DESCRIPCION, STOCK.OBSERVACIONES, STOCK.PO, STOCK.PARTNUMBER,\
                STOCK.ETIQUETADOR, STOCK.BODEGA_UBICACION, STOCK.SECTOR_UBICACION,\
                STOCK.CALLE_UBICACION, STOCK.OBSOLETO, STOCK.BVL, STOCK.MN, STOCK.DIF, STOCK.KG\
                FROM CPP.dbo.STOCK STOCK", cnxn) #STOCK DE TODO! A LA FECHA

   rc = pd.read_sql("SELECT recetas_completas.PARTNUMBER, recetas_completas.DESCRIPCION_PARTNUMBER, \
                  recetas_completas.PID, recetas_completas.NIVEL,\
                  recetas_completas.CLASIFICACION, recetas_completas.FAMILIA, recetas_completas.TIPO,\
                  recetas_completas.NOMBRE,\
                  recetas_completas.DESCRIPCION,\
                  recetas_completas.ESPESOR, recetas_completas.ANCHO,\
                  recetas_completas.LARGO, recetas_completas.ESPECIE, \
                  recetas_completas.PERFIL, recetas_completas.PANEL, \
                  recetas_completas.VIDRIO,\
                  recetas_completas.STILE, recetas_completas.ITEM, \
                  recetas_completas.DETALLE, recetas_completas.PID_H,\
                  recetas_completas.CLASIFICACION_H, recetas_completas.FAMILIA_H,\
                  recetas_completas.TIPO_H, recetas_completas.NOMBRE_H,\
                  recetas_completas.DESCRIPCION_H, recetas_completas.ESPESOR_H,\
                  recetas_completas.ANCHO_H, recetas_completas.LARGO_H,\
                  recetas_completas.ESPECIE_H, recetas_completas.PERFIL_H,\
                  recetas_completas.ITEM_H, recetas_completas.DETALLE_H,\
                  recetas_completas.QTY, recetas_completas.PROCESO,\
                  recetas_completas.M3, recetas_completas.REG, \
                  recetas_completas.REC, recetas_completas.FORMATO,\
                  recetas_completas.UOM, recetas_completas.COSTO, recetas_completas.BVL\
                  FROM CPP.dbo.recetas_completas recetas_completas where clasificacion='puertas' ", cnxn) #recetas completas


   s = pd.read_csv('s.csv', sep = ',')
   rc = pd.read_csv('rc.csv', sep = ',') 
   #pretratamiento DataFrame´s
   rc = pd.DataFrame(rc)
   s = pd.DataFrame(s)

   s = pd.read_csv('s.csv', sep = ',')
   rc.columns = map(str.lower,rc.columns)

   s.columns = map(str.lower,s.columns)
   s = s[s.version != 0]
   s = s.rename(columns={"id": "pid_h"})

   rc = rc.set_index('pid_h')
   s = s.set_index('pid_h')

   s = s[['piezas','tarjeta','m3_tarjeta','estado']]
   s = s[s['estado'].isin(['A'])]

   #El proceso JOIN

   rcs = rc.join(s)
   rcs = rcs.replace(np.nan, 0 , regex=True)
   rcs['validadas_h'] =  rcs.piezas/rcs.qty
   rcs = rcs.reset_index()

   rcs = rcs[rcs['familia_h'].isin(['BARRAS', 'VIDRIOS', 'STILE', 'STILE LOUVER', 'PANELES','SLATS',
                                    'RAILS', 'RAILS LOUVER', 'SPEC L', 'APLICACIONES',
                                    'MOLDURAS SOLIDAS', 'BLANK', 'MARCOS', 'OTRAS', 'MOLDURAS FINGER'])]

   rcs = rcs.groupby(['pid','descripcion','espesor',
                     'ancho', 'largo', 'especie', 'perfil','pid_h','familia_h',
                     'descripcion_h'])['validadas_h'].sum()

   rcs = pd.DataFrame(rcs)
   rcs = rcs.reset_index()

   if __name__ == "__main__":
         df = rcs
      #Para sacar el minimo p_v por pid
         table = df.groupby(['pid']).validadas_h.min()
      #Si lo quieres agregar al dataframe original
         rcs['puertas_validadas'] = table.loc[df.pid].values
         
   rcs = rcs.groupby(['pid','puertas_validadas','descripcion','espesor',
                     'ancho', 'largo', 'especie', 'perfil','pid_h','familia_h',
                     'descripcion_h'])['validadas_h'].sum()
   rcs = pd.DataFrame(rcs)
   return rcs

def resumen(cierre, semana):
    cierre = cierre[cierre['ENSAMBLE'].isin(['A','M'])]
    cierre = cierre[cierre['SEM RE'].isin([semana])]
    cierre = cierre.replace(np.nan, 0 , regex=True)
    cierre.PO = cierre.PO.astype(str)
    cierre = cierre.groupby(['PO','CUSTOMER','linea'])['QTY','X ENS','ENS','STOCK','X PINTAR','X EMBALAR'].sum()
    cierre = cierre.sum(level=['PO','CUSTOMER'])
    resumen = cierre.drop(columns = ['QTY','ENS','STOCK'])
    return resumen  

def valida_general_(rc,s):
   #pretratamiento DataFrame´s
   rc = pd.DataFrame(rc)
   s = pd.DataFrame(s)
   rc.columns = map(str.lower,rc.columns)
   rc = rc.set_index('pid_h')

   s.columns = map(str.lower,s.columns)
   s = s[s.version != 0]
   s = s.rename(columns={"id": "pid_h"})
   s = s.set_index('pid_h')
   s = s[['piezas','tarjeta','m3_tarjeta','estado']]
   s = s[s['estado'].isin(['A'])]

   #El proceso JOIN

   rcs = rc.join(s)
   rcs = rcs.replace(np.nan, 0 , regex=True)
   rcs['validadas_h'] =  rcs.piezas/rcs.qty
   rcs = rcs.reset_index()

   rcs = rcs[rcs['familia_h'].isin(['BARRAS', 'VIDRIOS', 'STILE', 'STILE LOUVER', 'PANELES','SLATS',
                                    'RAILS', 'RAILS LOUVER', 'APLICACIONES',
                                    'MOLDURAS SOLIDAS', 'MARCOS', 'OTRAS', 'MOLDURAS FINGER'])]

   rcs = rcs.groupby(['pid','descripcion','espesor',
                     'ancho', 'largo', 'especie', 'perfil','pid_h','familia_h',
                     'descripcion_h'])['validadas_h'].sum()

   rcs = pd.DataFrame(rcs)
   rcs = rcs.reset_index()

   if __name__ == "__main__":
         df = rcs
      #Para sacar el minimo p_v por pid
         table = df.groupby(['pid']).validadas_h.min()
      #Si lo quieres agregar al dataframe original
         rcs['puertas_validadas'] = table.loc[df.pid].values
         
   rcs = rcs.groupby(['pid','puertas_validadas','descripcion','espesor',
                     'ancho', 'largo', 'especie', 'perfil','pid_h','familia_h',
                     'descripcion_h'])['validadas_h'].sum()
   rcs = pd.DataFrame(rcs)
   return rcs

#valida las puertas de una lista


def ID_PO(cierre,semana):
    cierre = cierre[cierre['ENSAMBLE'].isin(['A','M'])]
    cierre = cierre[cierre['SEM RE'].isin([semana])]
    cierre = cierre.replace(np.nan, 0 , regex=True)
    cierre.PO = cierre.PO.astype(str)
    cierre = cierre.set_index(['SEM RE', 'PO', 'ID', 'CUSTOMER'])
    ID_PO  = cierre.sum(level = [1,2])
    return ID_PO


def stock_managment_validacion(rc,s,lista,orden):
    #pretratamiento DataFrame´s
    rc = pd.DataFrame(rc)
    s = pd.DataFrame(s)    
    
    rc.columns = map(str.lower,rc.columns)
    s.columns = map(str.lower,s.columns)
    
    s = s[s.version != 0]
    s = s.rename(columns={"id": "pid_h"})
    s = s[s['estado'].isin(['A'])]
    
    s = s.groupby(['pid_h'])['piezas'].sum()
    rc = rc.set_index(['pid','pid_h'])
    
    #############################
    lista = lista.loc[orden]
    ###########################

    #seteo como indice el pid y pid_h
    lista['puertas_validadas'] = 0

    for i in lista: #para cada puerta pid
        result_array = []    
        for j in range(len(rc.loc[i])): #cada pid_h
            a = rc.loc[i].loc[j]['qty']
            b = s.loc[j]
            result = b/a #validacion de componentes
            result_array.append(result)
        pv = np.amin(result_array)
        if lista.loc[i]['X ENS'] < pv:
            pv = lista.loc[i]['X ENS']
        for k in range(len  ( rc.loc[i] ) ): #componente k
            a = rc.loc[i].loc[k]['qty']
            b = s.loc[k]
            s.loc[k] = b-a*pv
        lista = lista.loc[i]['puertas_validadas'] = pv #por revisar si funca
    return lista

#pv = valida_general_(rc,s).sum(level = [0,1]).reset_index('puertas_validadas').drop(columns = ['validadas_h'])

def X_PO(cierre,semana,orden,rc,s):
    for j in ID_PO(cierre,semana).loc[orden]:
        lista = ID_PO(cierre,semana).loc[orden] #j para los pid
        pv = stock_managment_validacion(rc,s,lista).loc[j]['puertas_validadas'] #entregar lista con las puertas validadas

        rcc = rc.set_index(['pid','pid_h'])
        c = ID_PO(cierre,semana).loc[orden].loc[j]['X ENS']-pv #puertas reales a validar
        rcc.loc[j]['qty'] = rcc.loc[j]['qty']*(max(c.values,0)) #componentes multiplicados por lo faltante por validar para pid j

    #recorrer para los j de la lista de la PO

    for k in ID_PO(cierre,semana).loc[orden][0:-1]:
        a = rcc.loc[k]['qty']
        b = pd.concat([a , rcc.loc[k+1]['qty']], axis =0)

    return b    #contenacion de las DF (qty y pid_h(indexados))

for i in ID_PO(cierre,'19'):
    # sumar promedio blablabla X_PO(cierre,semana,orden,rc,s)

    return ID_PO(cierre,'19')['porcentaje rail']
