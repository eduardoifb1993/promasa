import pyodbc
import pandas as pd
import datetime as dt
import numpy as np
import datetime
from datetime import datetime, timedelta
from datetime import date
from operator import itemgetter

today = date.today()
last_monday = today - timedelta(days=today.weekday())
wk = date.today() + timedelta(days=7)
wk = wk.strftime('%Y%W')
wk = int(wk)

K = 16000

#Lee programa
programa = pd.read_csv('programa_ensamble_202018.csv', sep = ',')
arrive = pd.read_csv('arrive.csv',sep = ',') 

## cambio de nombre columnas

programa = programa.rename(columns={'W PROD': 'wk_prod'})
programa = programa.rename(columns={'W CIERRE': 'wk_cierre'})
programa = programa.rename(columns={'QTY REQUERIDO': 'qty_requerido'})
programa = programa.rename(columns={'QTY ENSAMBLAR': 'qty_ensamblar'})
programa = programa.rename(columns={'X ENSAMBLE': 'x_ensamble'})
programa = programa.rename(columns={'X PROD': 'x_producir'})
programa = programa.rename(columns={'PID': 'pid'})
programa = programa.rename(columns={'PID ENSAMBLE': 'pid_ensamble'})
programa = programa.rename(columns={'DESCRIPCION': 'descripcion'})
programa['wk_cierre'].replace({'STOCK': '202023','COMP.':'202023','reserva':'202023'},inplace = True)

programa.columns = map(str.lower,programa.columns)
programa = programa.groupby(['pid','wk_prod', 'wk_cierre','descripcion','perfil'])['qty_requerido'].sum()
programa = pd.DataFrame(programa)
programa_1 = pd.DataFrame(programa)

pid_programa = np.array(programa.index.get_level_values('pid').unique())

#REORDENA BY SORT METHOD

arrive = arrive.sort_values(['wk_cierre','familia'], ascending=True) #XREALIZAR INPUT O ITERAR POR MEJOR ORDEN PARA OPTIMO
arrive = arrive.reset_index()

perfil = programa.reset_index()
perfil = perfil.groupby(['perfil','wk_prod'])['qty_requerido'].sum()
perfil = pd.DataFrame(perfil)

for i in range(len(arrive)):
    if (arrive.wk_cierre.loc[i] > wk): 
        if (arrive.pid.loc[i] in pid_programa):

            a = np.array(programa.loc[arrive.pid.loc[i]].reset_index().wk_prod.unique())
            a = a[a <= arrive.wk_cierre.loc[i]] # y que cumpla con la semana de cierre
            
            if len(a) == 0 : #No puede cumplir con la semana en donde esta programado #X MEJORAR
                programa = programa.reset_index()
                programa = programa.append({'pid': arrive.pid.loc[i],'wk_prod':arrive.wk_cierre.loc[i],\
                            'wk_cierre':arrive.wk_cierre.loc[i], 'descripcion': '-',\
                            'perfil':arrive.perfil.loc[i],'qty_requerido':arrive.qty.loc[i]},ignore_index = True)
                programa = programa.groupby(['pid','wk_prod', 'wk_cierre','descripcion','perfil'])['qty_requerido'].sum()
                print('ESTA EN EL PROGRAMA PERO NO CUMPLE SEMANA DE CIERRE/PROD igual se programa')
                programa = pd.DataFrame(programa)
                
            else:
                result = []
                result_array = []
                for k in range(len(a)):
                    result = [a[k], int(programa.sum(level= 1).loc[a[k]]) + arrive.qty.loc[i]]
                     #todas las producciones de semana a[k](donde esta programado) + el arrivo de puertas
                    result_array.append(result)
         
                b = sorted(result_array, key=itemgetter(1))
            
                if b[0][1] < K:
                    programa = programa.reset_index()
                    programa = programa.append({'pid': arrive.pid.loc[i],'wk_prod':b[0][0],\
                                'wk_cierre':arrive.wk_cierre.loc[i], 'descripcion': '-',\
                                 'perfil':arrive.perfil.loc[i],'qty_requerido':arrive.qty.loc[i]},ignore_index = True)
                    programa = programa.groupby(['pid','wk_prod', 'wk_cierre','descripcion','perfil'])['qty_requerido'].sum()
                    programa = pd.DataFrame(programa)
                    print('aqui programo, debo sumar al programa esta puerta')      
                else:
                    p = np.array(perfil.loc[arrive.perfil.loc[i]].\
                                            reset_index()[perfil.loc[arrive.perfil.loc[i]].\
                                            reset_index().wk_prod <= arrive.wk_cierre.loc[i]])
                    if len(p) == 0 : #No puede cumplir con la semana en donde esta programado XMEJORAR
                        programa = programa.reset_index()
                        programa = programa.append({'pid': arrive.pid.loc[i],'wk_prod':arrive.wk_cierre.loc[i],\
                                    'wk_cierre':arrive.wk_cierre.loc[i], 'descripcion': '-',\
                                    'perfil':arrive.perfil.loc[i],'qty_requerido':arrive.qty.loc[i]},ignore_index = True)
                        programa = programa.groupby(['pid','wk_prod', 'wk_cierre','descripcion','perfil'])['qty_requerido'].sum()
                        print('ESTA EN EL PROGRAMA PERO NO CUMPLE SEMANA DE CIERRE/PROD programo en semana de cierre')
                        programa = pd.DataFrame(programa)

                    else:
                        p = sorted(p, key=itemgetter(1),reverse=True)
                        result = []
                        result_array = []
                        for j in range(len(p)):
                            result = [p[j][0], int(programa.sum(level= 1).loc[p[j][0]]) + arrive.qty.loc[i]]
                            #calculo la suma de puertas por semana en donde esta programado el perfil
                            result_array.append(result)
                        c = sorted(result_array, key=itemgetter(1)) 

                        if np.amin(c, axis=0)[1] > K:
                            programa = programa.reset_index()
                            programa = programa.append({'pid': arrive.pid.loc[i],'wk_prod':c[0][0],\
                                                        'wk_cierre':arrive.wk_cierre.loc[i], 'descripcion': '-',\
                                                        'perfil':arrive.perfil.loc[i],'qty_requerido':arrive.qty.loc[i]},ignore_index = True)
                            programa = programa.groupby(['pid','wk_prod', 'wk_cierre','descripcion','perfil'])['qty_requerido'].sum()
                            print('Supera QTY de la semana (por perfil) SE PROGRAMA',np.amin(c, axis=0)[1])

                        elif np.amax(c, axis=0)[1] <= K: #Programado donde más esta el perfil al menor
                                programa = programa.reset_index()
                                programa = programa.append({'pid': arrive.pid.loc[i],'wk_prod':c[0][0],\
                                            'wk_cierre':arrive.wk_cierre.loc[i], 'descripcion': '-',\
                                             'perfil':arrive.perfil.loc[i],'qty_requerido':arrive.qty.loc[i]},ignore_index = True)
                                programa = programa.groupby(['pid','wk_prod', 'wk_cierre','descripcion','perfil'])['qty_requerido'].sum()
                                programa = pd.DataFrame(programa)
                                print('sumada al perfil')   
        else:
            o = np.array(perfil.loc[arrive.perfil.loc[i]].\
                                    reset_index()[perfil.loc[arrive.perfil.loc[i]].\
                                    reset_index().wk_prod <= arrive.wk_cierre.loc[i]])
            if len(o) == 0 : #No puede cumplir con la semana en donde esta programado XMEJORAR
                programa = programa.reset_index()
                programa = programa.append({'pid': arrive.pid.loc[i],'wk_prod':arrive.wk_cierre.loc[i],\
                            'wk_cierre':arrive.wk_cierre.loc[i], 'descripcion': '-',\
                            'perfil':arrive.perfil.loc[i],'qty_requerido':arrive.qty.loc[i]},ignore_index = True)
                programa = programa.groupby(['pid','wk_prod', 'wk_cierre','descripcion','perfil'])['qty_requerido'].sum()
                print('ESTA EN EL PROGRAMA PERO NO CUMPLE SEMANA DE CIERRE/PROD programo en semana de cierre')
                programa = pd.DataFrame(programa)

            else:
                o = sorted(o, key=itemgetter(1),reverse=True)
                result = []
                result_array = []
                for j in range(len(o)):
                    result = [o[j][0], int(programa.sum(level= 1).loc[o[j][0]]) + arrive.qty.loc[i]]
                    #calculo la suma de puertas por semana en donde esta programado el perfil
                    result_array.append(result)
                c = sorted(result_array, key=itemgetter(1)) 

                if np.amin(c, axis=0)[1] > K:
                    programa = programa.reset_index()
                    programa = programa.append({'pid': arrive.pid.loc[i],'wk_prod':c[0][0],\
                                                'wk_cierre':arrive.wk_cierre.loc[i], 'descripcion': '-',\
                                                'perfil':arrive.perfil.loc[i],'qty_requerido':arrive.qty.loc[i]},ignore_index = True)
                    programa = programa.groupby(['pid','wk_prod', 'wk_cierre','descripcion','perfil'])['qty_requerido'].sum()
                    print('Supera QTY de la semana (por perfil) SE PROGRAMA',np.amin(c, axis=0)[1])

                elif np.amax(c, axis=0)[1] <= K: #Programado donde más esta el perfil al menor
                        programa = programa.reset_index()
                        programa = programa.append({'pid': arrive.pid.loc[i],'wk_prod':c[0][0],\
                                    'wk_cierre':arrive.wk_cierre.loc[i], 'descripcion': '-',\
                                        'perfil':arrive.perfil.loc[i],'qty_requerido':arrive.qty.loc[i]},ignore_index = True)
                        programa = programa.groupby(['pid','wk_prod', 'wk_cierre','descripcion','perfil'])['qty_requerido'].sum()
                        programa = pd.DataFrame(programa)
                        print('sumada al perfil') 
            print('esta no estaba en el programa')  
    else:
         print('no cumple semana de cierre')

programa = pd.DataFrame(programa)
programa.to_csv('programa_nuevo_david.csv')