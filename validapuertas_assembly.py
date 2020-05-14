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