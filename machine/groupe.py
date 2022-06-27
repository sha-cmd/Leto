# -*- coding: utf-8 -*-
"""
Created on Mon May 18 10:43:04 2020

Dans ces trois fonctions nous trions les données aberrantes en réduisant la 
recherche dans la base aux traitement inférieur à 2 semaines de travail, soit
400 heures. Les fonction get group peuvent être réglées à notre guise.
Comme le travail est fait à partir du fichier excel de data.xlsx, lui-même 
obtenu par la fonction kpi.extickets(), donc elle doivent-être consécutive.

@author: rboyrie
"""

import pandas as pd
from machine import do

def getgroupprop(mstart,mend,ystart, yend, data):
    """ Cette fonction doit-être changé si le nombre de groupe augmente"""
    dfend = pd.DataFrame(columns=['date','quantite','groupe'])
    #data = data.mask(data['date_cloture'].eq('None')).dropna()
    #print(data)
    for year in range (ystart, yend+1):
        for month in range( 1, 12+1 ):
            if ( year == ystart ) & (month < mstart):
                continue
            if ( year >= yend ) & ( month >= mend+1 ):
                continue
            if month != 12:
                #df = data.loc[(pd.to_datetime(data['date_de_cloture'].dropna(),format='%Y-%m-%d')>=pd.datetime(year,month,1))&(data['Tt']<400*3600000000000)&(pd.to_datetime(data['date_de_cloture'].dropna(),format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]
                df = data.loc[(pd.to_datetime(data['date_cloture'].dropna(),format='%Y-%m-%d')>=pd.datetime(year,month,1))&(pd.to_datetime(data['date_cloture'].dropna(),format='%Y-%m-%d')<pd.datetime(year,month+1,1))&(data['id_statut_demande']==6)]
            else:
                #df = data.loc[(pd.to_datetime(data['date_de_cloture'].dropna(),format='%Y-%m-%d')>=pd.datetime(year,month,1))&(data['Tt']<400*3600000000000)&(pd.to_datetime(data['date_de_cloture'].dropna(),format='%Y-%m-%d')<pd.datetime(year+1,1,1))]
                df = data.loc[(pd.to_datetime(data['date_cloture'].dropna(),format='%Y-%m-%d')>=pd.datetime(year,month,1))&(pd.to_datetime(data['date_cloture'].dropna(),format='%Y-%m-%d')<pd.datetime(year + 1,1,1))&(data['id_statut_demande']==6)]
            g1,g2,g3,g4,g5 = 0,0,0,0,0
            g1 = df.loc[df['id_groupe']==1]
            g2 = df.loc[df['id_groupe']==2]
            g3 = df.loc[df['id_groupe']==3]
            g4 = df.loc[df['id_groupe']==4]
            g5 = df.loc[df['id_groupe']==5]
            g6 = df.loc[df['id_groupe']==6]
            #lan = {'date':pd.datetime(year,month,1),'groupe 1':g1.shape[0],'groupe 2':g2.shape[0],'groupe 3':g3.shape[0]\
             #      ,'groupe 4':g4.shape[0],'groupe 5':g5.shape[0]}
            #dfend = dfend.append(lan,ignore_index=True)
            liste_de_groupe = [g1,g2,g3,g4,g5,g6]
            #print(liste_de_groupe)
            for i, val in enumerate(liste_de_groupe):# 5 groupe
                #p =val['id_groupe'].unique()[0]
                #print(p)
                try:
                    #p =val['id_groupe'].unique()[0]
                    #print(p)
               #     print(val['id_groupe'].values[0][:])
                    lan = {'date':pd.datetime(year,month,1),'quantite':val.shape[0],'groupe':str(val['id_groupe'].unique()[0])}
                    dfend = dfend.append(lan,ignore_index=True)
                except:
                    continue
    dfend['groupe']=dfend['groupe'].map({'1':'Front Office','2':'Infrastructure','3':'Support Applicatif','4':'Support Matériel','5':'Développement','6':'maitrise d\'ouvrage'})
    dfFinal=pd.DataFrame.from_dict(dfend)
    dfFinal.index=dfFinal['date']
    dfFinal = dfFinal.drop('date',axis=1)
    #print(dfend)
    dfFinal.to_excel("grp_quantas.xlsx")
    do.to_sql(dfFinal,db='group_quantities')
    return dfFinal

def getgroupmeans(mstart,mend, ystart, yend,data):
    dfend = pd.DataFrame(columns=['date','Tt mean','groupe'])
    #data = data.mask(data['date_cloture'].eq('None')).dropna()
    for year in range (ystart, yend+1):
        for month in range( 1, 12+1 ):
            if ( year == ystart ) & (month < mstart):
                continue
            if ( year >= yend ) & ( month >= mend+1 ):
                continue
            if month != 12:
                #df = data.loc[(pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')>=pd.datetime(year,month,1))&(data['Tt']<400*3600000000000)&(pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]
                df = data.loc[(pd.to_datetime(data['date_cloture'].dropna(),format='%Y-%m-%d')>=pd.datetime(year,month,1))&(pd.to_datetime(data['date_cloture'].dropna(),format='%Y-%m-%d')<pd.datetime(year,month+1,1))&(data['id_statut_demande']==6)]

            else:
                #df = data.loc[(pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')>=pd.datetime(year,month,1))&(data['Tt']<400*3600000000000)&(pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year+1,1,1))]
                df = data.loc[(pd.to_datetime(data['date_cloture'].dropna(),format='%Y-%m-%d')>=pd.datetime(year,month,1))&(pd.to_datetime(data['date_cloture'].dropna(),format='%Y-%m-%d')<pd.datetime(year + 1,1,1))&(data['id_statut_demande']==6)]
            g1,g2,g3,g4,g5 = 0,0,0,0,0
            g1 = df.loc[df['id_groupe']==1]
            g2 = df.loc[df['id_groupe']==2]
            g3 = df.loc[df['id_groupe']==3]
            g4 = df.loc[df['id_groupe']==4]
            g5 = df.loc[df['id_groupe']==5]
            g6 = df.loc[df['id_groupe']==6]
            liste_de_groupe = [g1,g2,g3,g4,g5,g6]
            #print(liste_de_groupe)
            for i, val in enumerate(liste_de_groupe):# 5 groupe
             #   p =val['str_groupe'].values[0][:]
            #print(p)
                try:
                    lan = {'date':pd.datetime(year,month,1),'Tt mean':val['treatment_time'].mean(),'groupe':str(val['id_groupe'].unique()[0])}
                    dfend = dfend.append(lan,ignore_index=True)
                except:
                    continue
            #lan = {'date':pd.datetime(year,month,1),'groupe 1 Tt mean':g1,'groupe 2 Tt mean':g2,\
             #      'groupe 3 Tt mean':g3\
              #     ,'groupe 4 Tt mean':g4,'groupe 5 Tt mean':g5}
            #dfend = dfend.append(lan,ignore_index=True)
    dfend['groupe']=dfend['groupe'].map({'1':'Front Office','2':'Infrastructure','3':'Support Applicatif','4':'Support Matériel','5':'Développement','6':'maitrise d\'ouvrage'})
    dfFinal=pd.DataFrame.from_dict(dfend)
    dfFinal.index=dfFinal['date']
    dfFinal = dfFinal.drop('date',axis=1)
    dfFinal.to_excel("grp_means.xlsx")
    do.to_sql(dfFinal, db='group_means')
    return dfFinal

def getgroupsums(mstart,mend, ystart, yend, data):
    dfend = pd.DataFrame(columns=['date','groupe 1 Tt sum','groupe 2 Tt sum','groupe 3 Tt sum','groupe 4 Tt sum','groupe 5 Tt sum'])
    data = data.mask(data['date_de_cloture'].eq('None')).dropna()
    for year in range (ystart, yend+1):
        for month in range( 1, 12+1 ):
            if ( year == ystart ) & (month < mstart):
                continue
            if ( year >= yend ) & ( month >= mend+1 ):
                continue
            if month != 12:
                df = data.loc[(pd.to_datetime(data['date_cloture'],format='%Y-%m-%d')>=pd.datetime(year,month,1))&(data['Tt']<400*3600000000000)&(pd.to_datetime(data['date_cloture'],format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]
            else:
                df = data.loc[(pd.to_datetime(data['date_cloture'],format='%Y-%m-%d')>=pd.datetime(year,month,1))&(data['Tt']<400*3600000000000)&(pd.to_datetime(data['date_cloture'],format='%Y-%m-%d')<pd.datetime(year+1,1,1))]
            g1,g2,g3,g4,g5 = 0,0,0,0,0
            g1 = df.loc[df['groupe']==1]['Tt'].sum()/3600000000000
            g2 = df.loc[df['groupe']==2]['Tt'].sum()/3600000000000
            g3 = df.loc[df['groupe']==3]['Tt'].sum()/3600000000000
            g4 = df.loc[df['groupe']==4]['Tt'].sum()/3600000000000
            g5 = df.loc[df['groupe']==5]['Tt'].sum()/3600000000000
            lan = {'date':pd.datetime(year,month,1),'groupe 1 Tt sum':g1,'groupe 2 Tt sum':g2,\
                   'groupe 3 Tt sum':g3\
                   ,'groupe 4 Tt sum':g4,'groupe 5 Tt sum':g5}
            dfend = dfend.append(lan,ignore_index=True)
    
    dfFinal=pd.DataFrame.from_dict(dfend)
    dfFinal.index=dfFinal['date']
    dfFinal = dfFinal.drop('date',axis=1)
    dfFinal.to_excel("grp_sums.xlsx")
    do.to_sql(dfFinal,db='group_sums')
    return dfFinal