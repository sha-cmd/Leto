# -*- coding: utf-8 -*-
"""
Created on Tue Feb 11 10:44:43 2020
Ne jamais lancer le programme vers une date, lorsque la table demande ne 
contient pas le mois en entier.
@author: rboyrie
"""
# for 3D histogram
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D

import pandas as pd
import numpy as np
from sklearn import linear_model
from item import data_mining as dm
from item import regression as rg
from machine import tag_cloud as tgc
from machine import stats
from machine import kpi
from machine import request as rq
from machine import qualite as qlt
from machine import do
from machine import indicateur
from machine import groupe

import datetime
# Pour mesurer la mémoire du processus fonctionnel
# $ python -m memory_profiler example.py
#@profile
def create_kpi(month_start, month_end, year_start, year_end):
    """
    Lance le programme de data science afin de créer les kpi à partir des 
    tables de production "demande, prise en charge et suivi intervention,
    suivi ressource". Les résultats sont écrit dans de nouvelles tables 
    (environ 20) et contiennent toutes sortes d’indicateur. La première des
    tables est suffixé par le mot data. Le nom de la base de données utilisée
    et les properties pour se connecter à la base sont à paramétrer dans le 
    module do dans le répertoire machine. Il est de plus nécessaire de modifier
    le code si l’on veut utiliser un serveur distant, en modifiant le paramètre
    nommé local à des fonctions qui accède à la base de données pour qu’elle 
    écrive sur le serveur de DSI.
    """
    
    print('Bienvenu(e) dans le programme kpi')
    db="helpdesk_data"
    ci = True
    with open('bullet_time.txt', 'a') as f:
        create_date_table()
        #Créé une base de tous les objets et leurs caractéristiques
        date1 = datetime.datetime.now()
        kpi.extickets(db=db, calcul_integral=ci)
        date2 = datetime.datetime.now()
        print('durée de la phase 1 :', date2-date1)
        print('durée de la phase 1 :', date2-date1, file=f)
    
        #########################################
        ## Génération des catégories          ###
        #########################################
        stats.scan_cat(month_start, month_end, year_start, year_end,\
                       file_name='helpdesk_kpi_tables_quantite.xlsx',\
                       calcul_integral=ci)
        date3 = datetime.datetime.now()
        print('durée de la phase 2 :',date3-date2)
        print('durée de la phase 2 :',date3-date2, file=f)
        #########################################
        ## Calculer t1 et le mettre en base #####
        #########################################
        stats.scan_cat_t1(month_start, month_end, year_start, year_end,\
                          file_name='data_t1.xlsx',\
                          calcul_integral=ci)        
        date4 = datetime.datetime.now()
        print('durée de la phase 3 :',date4-date3)
        print('durée de la phase 3 :',date4-date3, file = f)
        stats.scan_cat_t2(month_start, month_end, year_start, year_end,\
                          file_name='data_t2.xlsx',\
                          calcul_integral=ci)        
        date5 = datetime.datetime.now()
        print('durée de la phase 4 :',date5-date4)
        print('durée de la phase 4 :',date5-date4, file = f)
        stats.scan_cat_t3(month_start, month_end, year_start, year_end,\
                          file_name='data_t3.xlsx',\
                          calcul_integral=ci)        
        date6 = datetime.datetime.now()
        print('durée de la phase 5 :',date6-date5)
        print('durée de la phase 5 :',date6-date5, file = f)
        stats.scan_cat_tr(month_start, month_end, year_start, year_end,\
                          file_name='data_tr.xlsx',\
                          calcul_integral=ci)
        date7 = datetime.datetime.now()
        print('durée de la phase 6 :',date7-date6)
        print('durée de la phase 6 :',date7-date6, file = f)
        stats.scan_cat_tc(month_start, month_end, year_start, year_end,\
                          file_name='data_tc.xlsx',\
                          calcul_integral=ci)
        date8 = datetime.datetime.now()
        print('durée de la phase 7 :',date8-date7)
        print('durée de la phase 7 :',date8-date7, file = f)
        stats.scan_cat_trc(month_start, month_end, year_start, year_end,\
                           file_name='data_trc.xlsx',\
                          calcul_integral=ci)
        date9 = datetime.datetime.now()
        print('durée de la phase 8 :',date9-date8)
        print('durée de la phase 8 :',date9-date8, file = f)
        stats.scan_cat_Tt(month_start, month_end, year_start, year_end,\
                          file_name='data_Tt.xlsx',\
                          calcul_integral=ci)
        date10 = datetime.datetime.now()
        print('durée de la phase 9 :',date10-date9)
        print('durée de la phase 9 :',date10-date9, file = f)
        stats.moy_tt_by_month(month_start,month_end,year_start,year_end, calcul_integral=ci)
        stats.moy_tt_by_day_in_a_month(month_start,month_end,year_start,year_end, calcul_integral=ci)
        ####################
        ## Faire les KPI  ##
        ####################
        
        kpi.make(month_start,month_end,year_start,year_end, calcul_integral=ci)
        date11 = datetime.datetime.now()
        print('durée de la phase 10 :',date11-date10)
        print('durée de la phase 10 :',date11-date10, file = f)
        qlt.make()
        date12 = datetime.datetime.now()
        print('durée de la phase 11 :',date12-date11)
        print('durée de la phase 11 :',date12-date11, file = f)
        indicateur.make()
        date13 = datetime.datetime.now()
        print('durée de la phase 13 :',date13-date12)
        print('durée de la phase 13 :',date13-date12, file = f)
        dm.scores(month_start,month_end,year_start,year_end)
        date14 = datetime.datetime.now()
        print('durée de la phase 14 :',date14-date13)
        print('durée de la phase 14 :',date14-date13, file = f)
        dm.synopsis()
        date15 = datetime.datetime.now()
        print('durée de la phase 15 :',date15-date14)
        print('durée de la phase 15 :',date15-date14, file = f)
        date16 = datetime.datetime.now()
        data =  do.r_to_b(query='SELECT * FROM synopsis')
        groupe.getgroupprop(month_start,month_end,year_start,year_end, data, calcul_integral=ci)
        groupe.getgroupmeans(month_start,month_end,year_start,year_end, data, calcul_integral=ci)
        #groupe.getgroupsums(month_start,month_end,year_start,year_end, data)
        print('durée de la phase 12 :',date13-date12)
        print('durée de la phase 12 :',date13-date12, file = f)
        # Au lieu de la régression, faire un ridge regressor à plusieurs fonctionnalités
        # Cette table n’est pas utilisée dans PBI.
        rg.do_regression_times(month_start,month_end,year_start,year_end)
        date17 = datetime.datetime.now()
        print('durée de la phase 16 :',date17-date16)
        print('durée de la phase 16 :',date17-date16, file = f)
        tgc.make(2,280,700)
        date18 = datetime.datetime.now()
        print('durée de la phase 17 :',date18-date17)
        print('durée de la phase 17 :',date18-date17, file = f)
        dm.outliers(month_start,month_end,year_start,year_end,2,280,700)#Limit laissant N1, Résolution et Traitement entre 0 et 30 tickets par mois
        date19 = datetime.datetime.now()
        print('durée de la phase 18 :',date19-date18)
        print('durée de la phase 18 :',date19-date18, file = f)
        # Calcul du marché par client et par technicien
        rg.do_supply_and_demand(month_start,month_end,year_start,year_end, calcul_integral=ci, limit_r_basse=28, limit_r_moyenne=12, limit_r_haute=1.5, limit_r_tres_haute=0.5, limit_t_basse=40, limit_t_moyenne=16, limit_t_haute=2, limit_t_tres_haute=1)# Rajoute les discriminants des urgences en heure
        date20 = datetime.datetime.now()
        print('durée de la phase 20 :',date20-date19)
        print('durée de la phase 20 :',date20-date19, file = f)
        qlt.school()
        date21 = datetime.datetime.now()
        print('durée de la phase 21 :',date21-date20)
        print('durée de la phase 21 :',date21-date20, file = f)
        stats.productivity(month_start,month_end,year_start,year_end)
        date22 = datetime.datetime.now()
        print('durée de la phase 22 :',date22-date21)
        print('durée de la phase 22 :',date22-date21, file = f)
        stats.productivity_synopsis(10,8,2019,2020)
        date23 = datetime.datetime.now()
        print('durée de la phase 23 :',date23-date22)
        print('durée de la phase 23 :',date23-date22, file = f)
        stats.make(month_start,month_end,year_start,year_end,40,16,2,1)#Poids en heure des tickets urgence basse, moyenne, haute, très haute
        date24 = datetime.datetime.now()
        print('durée totale :',date24-date1)
        print('durée totale :',date24-date1, file = f)
####################################################
## Création d’une table de calendrier :
def create_date_table(start='2019-01-01', end='2020-12-31'):
    df = pd.DataFrame({"Date": pd.date_range(start, end)})
    df["Week"] = df.Date.dt.weekofyear
    df["str_Month"] = df.Date.dt.month_name()
    df["Month"] = df.Date.dt.month
    df["str_Day"] = df.Date.dt.weekday_name
    df["Day"] = df.Date.dt.day
    df["Quarter"] = df.Date.dt.quarter
    df["Year"] = df.Date.dt.year
    #df['Month_Year'] = pd.concat(str(df.Date.dt.month), str(df.Date.dt.year))
    df["Year_half"] = (df.Quarter + 1) // 2
    df["str_Month"]=df["str_Month"].map({'January': 'Janvier','February': 'Février','March': 'Mars','April': 'Avril',
      'May': 'Mai','June': 'Juin','July': 'Juillet','August': 'Août','September': 'Septembre',
      'October': 'Octobre','November': 'Novembre','December': 'Décembre'})
    df["str_Day"]=df["str_Day"].map({'Monday': 'Lundi','Tuesday': 'Mardi','Wednesday': 'Mercredi','Thursday': 'Jeudi',
      'Friday': 'Vendredi','Saturday': 'Samedi','Sunday': 'Dimanche'})
    do.to_sql(df,db='times')
    #print(df)
    return df
####################################################

def barchart3D():
    # setup the figure and axes
    fig = plt.figure(figsize=(8, 3))
    ax1 = fig.add_subplot(121, projection='3d')
    #ax2 = fig.add_subplot(122, projection='3d')
    
    # fake data
    _x = np.arange(4)
    _y = np.arange(5)
    _xx, _yy = np.meshgrid(_x, _y)
    x, y = _xx.ravel(), _yy.ravel()
    
    top = x + y
    bottom = np.zeros_like(top)
    width = depth = 1
    
    ax1.bar3d(x, y, bottom, width, depth, top, shade=True)
    ax1.set_title('Shaded')

    plt.show()

#create_date_table(start='2019-01-01', end='2020-12-31')

## Création de l’indice Mensuel    
#create_kpi(1,8,2019,2020)
    # Fonction extrême en complexité, à polariser.
#rg.do_supply_and_demand( 1, 6, 2019, 2020, calcul_integral=True, limit_r_basse=28, limit_r_moyenne=12, limit_r_haute=1.5, limit_r_tres_haute=0.5, limit_t_basse=40, limit_t_moyenne=16, limit_t_haute=2, limit_t_tres_haute=1)


#kpi.make(1,5,2019,2020)

#data = pd.read_excel('data.xlsx')   
#data['Times'] = data['Tr']
#data['Times'] = data['Times'].replace(0,data['Trc'])
#data['Tech'] = data['Tech_res']
#data['Tech'] = data['Tech'].replace(0,data['Tech_clos'])
#print(data.loc[data['demande']==1764]['Tech'])

#create_date_table2()
#data =  do.r_to_b(query='SELECT * FROM synopsis')

# Productivité 
#stats.make(1,6,2019,2020,40,16,2,1)

#stats.ticket_cloture_jour_ferie(1,6,2019,2020       )
#groupe.getgroupprop(1,6,2019,2020, data)
#groupe.getgroupmeans(1,6,2019,2020, data)

#dm.scores(1,4,2019,2020)
# Relancer à la fin 1 juin 2020     
#dm.scores(1,6,2019,2020)
#dm.synopsis()
#qlt.make()

#qlt.school()
#tgc.make(2,280,700)
#dm.outliers(1,5,2019,2020,2,280,700)
#rg.do_supply_and_demand(10,5,2019,2020)
#print(create_date_table2())
#rg.do_regression_times(1,6,2019,2020)
#rg.make_graph()
rg.do_regression_times(1,8,2019,2020)
#stats.productivity_synopsis(10,8,2019,2020)
