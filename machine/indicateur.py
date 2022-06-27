# -*- coding: utf-8 -*-
"""
Created on Mon Apr  6 08:22:08 2020
Ce module crée la note indicatrice de la performance du mois. 
Il y a un calcul des performances temps, puis elles sont pondérées par leur
quantité par rapport au total. Ces indicateurs sont mensuels, sont mensuel,
donc il font un appel à la base sur les tickets regroupés par mois et non 
par semaines.
@author: rboyrie
"""
from machine import do as do
import numpy as np
import pandas as pd
import datetime

#Constante
date = datetime.datetime.now()
nb = int(date.month) + 2 # Mois de Mars vaut 6, ajouter 1 pour chaque mois supplémentaire
#print(nb)
mode = 'replace'

def make():
    """
    Cette fonction est un appel simplifiant la création de la note de notre
    indicateur unique de performance.
    """
    quantdelta()
    quantproportion()
    attenteniveau1delta()
    resolutiondelta()
    avantcloturedelta()
    traitementdelta()
    note() 

def quantdelta():
    """
    Cette fonction calcul et écrit dans la table delta quantité, la variation
    des quantités de tickets pour un mois données.
    """
    # Calcul des deltas des quantités 
    df = pd.DataFrame()
    for process in range( 0, 5):
        data = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_data_kpi_process_" + str(process))
        dfd = data["Nombre de demande de ressource"][-nb:].div(data["Nombre de demande de ressource"][-nb-1:-1].values).sub(1).astype(np.float64)
        dfd.index = data['date'][-nb:]
        dfi = data["Nombre de demande de support"][-nb:].div(data["Nombre de demande de support"][-nb-1:-1].values).sub(1).astype(np.float64)
        dfi.index = data['date'][-nb:]
        dfd.fillna(float(0), inplace = True)
        dfi.fillna(float(0), inplace = True)
     
        df["Ressource_"+str(process)] = dfd
        df["Support_"+str(process)] = dfi
        df["Ressource_"+str(process)] = df["Ressource_"+str(process)].replace([np.inf, -np.inf],np.float64(0),regex=True)
        df["Support_"+str(process)] = df["Support_"+str(process)].replace([np.inf, -np.inf],np.float64(0),regex=True)
    do.to_sql(df,db="helpdesk_kpi_tables_delta_quantite", mode=mode)
    return df   
#df = quantdelta()

def quantproportion():
    """
    Cette fonction calcul et écrit en base la proportion de chaque processus 
    de chaque type par rapport au nombre total de ticket.
    """
    # Calcul des deltas des quantités 
    df = pd.DataFrame()
    temps = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_tt")
    data = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_quantite")
    quantite = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_data_kpi_process_0")
    for process in range( 1, 5):
        dfd = (data['nbdc'+str(process)].add(data['nbdrc'+str(process)])).div(quantite['Traitement Quantité'])
        dfd.index = temps['mois']
        dfi = (data['nbic'+str(process)].add(data['nbirc'+str(process)])).div(quantite['Traitement Quantité'])
        dfi.index = temps['mois']
        df["Proportion_Ressource_"+str(process)] = dfd
        df["Proportion_Support_"+str(process)] = dfi
        df.index.name = "date"
    do.to_sql(df,db="helpdesk_kpi_tables_poids_quantite_proportion", mode=mode)
    return df   

def attenteniveau1delta():
    """
    Cette fonction calcul et écrit en base la performance mensuelle d’attente
    de niveau 1 pour chaque processus, ainsi que pour chaque type.
    """
    # Calcul des deltas des résolutions
    df = pd.DataFrame()
    temps = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_tt")
    nombre = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_quantite")
    quantite = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_data_kpi_process_0")
    # pp est la proportion pour le processus 1, 2, 3, 4
    pp1 = (nombre['nbdc1'].add(nombre['nbdrc1'])).div(quantite['Traitement Quantité'])
    pp2 = (nombre['nbdc2'].add(nombre['nbdrc2'])).div(quantite['Traitement Quantité'])
    pp3 = (nombre['nbdc3'].add(nombre['nbdrc3'])).div(quantite['Traitement Quantité'])
    pp4 = (nombre['nbdc4'].add(nombre['nbdrc4'])).div(quantite['Traitement Quantité'])
    pp1.index = temps['mois']
    pp2.index = temps['mois']
    pp3.index = temps['mois']
    pp4.index = temps['mois']
    for process in range( 0, 5):
        data = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_data_kpi_process_" + str(process))
        data["Prise en charge (Ressource) inside SLA"][-nb-1:-1].fillna(float(0), inplace = True)
        data["Prise en charge (Ressource) inside SLA"]=data["Prise en charge (Ressource) inside SLA"].astype(np.float64)
        dfd1 = data['Prise en charge (Ressource) inside SLA'][-nb:]
        dfd2 = data['Prise en charge (Ressource) inside SLA'][-nb-1:-1]
        dfd1.index = data['date'][-nb:]
        dfd2.index = data['date'][-nb:]
        
        dfd = dfd1.div(dfd2).sub(1)
        dfd.fillna(float(0), inplace = True)
        
        data["Prise en charge (Support) inside SLA"][-nb-1:-1].fillna(float(0), inplace = True)
        data["Prise en charge (Support) inside SLA"]=data["Prise en charge (Support) inside SLA"].astype(np.float64)
        dfi1 = data['Prise en charge (Support) inside SLA'][-nb:]
        dfi2 = data['Prise en charge (Support) inside SLA'][-nb-1:-1]
        dfi1.index = data['date'][-nb:]
        dfi2.index = data['date'][-nb:]
         
        dfi = dfi1.div(dfi2).sub(1)
        dfi.fillna(float(0), inplace = True)
        
        df["Attente_N1_Ressource_"+str(process)] = dfd
        df["Attente_N1_Support_"+str(process)] = dfi
        df["Attente_N1_Ressource_"+str(process)] = df["Attente_N1_Ressource_"+str(process)].replace([np.inf, -np.inf],np.float64(0),regex=True)
        df["Attente_N1_Support_"+str(process)] = df["Attente_N1_Support_"+str(process)].replace([np.inf, -np.inf],np.float64(0),regex=True)
    
    df["Attente_N1_Ressource_0"] = (df["Attente_N1_Ressource_1"].mul(pp1))\
                                    .add((df["Attente_N1_Ressource_2"].mul(pp2)))\
                                    .add((df["Attente_N1_Ressource_3"].mul(pp3)))\
                                    .add((df["Attente_N1_Ressource_4"].mul(pp4)))
    df["Attente_N1_Support_0"] = (df["Attente_N1_Support_1"].mul(pp1))\
                                    .add((df["Attente_N1_Support_2"].mul(pp2)))\
                                    .add((df["Attente_N1_Support_3"].mul(pp3)))\
                                    .add((df["Attente_N1_Support_4"].mul(pp4)))
    
    do.to_sql(df,db="helpdesk_kpi_tables_delta_attente_niveau_1",mode=mode)
    
    return df

def resolutiondelta():
    """
    Cette fonction calcul et écrit en base la performance mensuelle de résolution
    pour chaque processus, ainsi que pour chaque type.
    """
    # Calcul des deltas des résolutions
    df = pd.DataFrame()
    temps = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_tt")
    nombre = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_quantite")
    quantite = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_data_kpi_process_0")
    pp1 = (nombre['nbdc1'].add(nombre['nbdrc1'])).div(quantite['Traitement Quantité'])
    pp2 = (nombre['nbdc2'].add(nombre['nbdrc2'])).div(quantite['Traitement Quantité'])
    pp3 = (nombre['nbdc3'].add(nombre['nbdrc3'])).div(quantite['Traitement Quantité'])
    pp4 = (nombre['nbdc4'].add(nombre['nbdrc4'])).div(quantite['Traitement Quantité'])
    pp1.index = temps['mois']
    pp2.index = temps['mois']
    pp3.index = temps['mois']
    pp4.index = temps['mois']
    for process in range( 0, 5):
        data = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_data_kpi_process_" + str(process))
        data["Trd"][-nb-1:-1].fillna(float(0), inplace = True)
        data["Trd"]=data["Trd"].astype(np.float64)
        dfd1 = data['Trd'][-nb:]
        dfd2 = data['Trd'][-nb-1:-1]
        dfd1.index = data['date'][-nb:]
        dfd2.index = data['date'][-nb:]
        
        dfd = dfd1.div(dfd2).sub(1)
        dfd.fillna(float(0), inplace = True)
        
           
        data["Tri"][-nb-1:-1].fillna(float(0), inplace = True)
        data["Tri"]=data["Tri"].astype(np.float64)
        dfi1 = data['Tri'][-nb:]
        dfi2 = data['Tri'][-nb-1:-1]
        dfi1.index = data['date'][-nb:]
        dfi2.index = data['date'][-nb:]
        
        dfi = dfi1.div(dfi2).sub(1)
        dfi.fillna(float(0), inplace = True)
           
        df["Resolution_Ressource_"+str(process)] = dfd
        df["Resolution_Support_"+str(process)] = dfi
        df["Resolution_Ressource_"+str(process)] = df["Resolution_Ressource_"+str(process)].replace([np.inf, -np.inf],np.float64(0),regex=True)
        df["Resolution_Support_"+str(process)] = df["Resolution_Support_"+str(process)].replace([np.inf, -np.inf],np.float64(0),regex=True)
    df["Resolution_Ressource_0"] = (df["Resolution_Ressource_1"].mul(pp1))\
                                    .add((df["Resolution_Ressource_2"].mul(pp2)))\
                                    .add((df["Resolution_Ressource_3"].mul(pp3)))\
                                    .add((df["Resolution_Ressource_4"].mul(pp4)))
    df["Resolution_Support_0"] = (df["Resolution_Support_1"].mul(pp1))\
                                    .add((df["Resolution_Support_2"].mul(pp2)))\
                                    .add((df["Resolution_Support_3"].mul(pp3)))\
                                    .add((df["Resolution_Support_4"].mul(pp4)))
    do.to_sql(df,db="helpdesk_kpi_tables_delta_resolution",mode=mode)

    return df

def avantcloturedelta():
    """
    Cette fonction calcul et écrit en base la performance mensuelle d’avant
    clôture pour chaque processus, ainsi que pour chaque type.
    """
    # Calcul des deltas des avant clôtures
    df = pd.DataFrame()
    temps = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_tt")
    nombre = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_quantite")
    quantite = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_data_kpi_process_0")
    pp1 = (nombre['nbdc1'].add(nombre['nbdrc1'])).div(quantite['Traitement Quantité'])
    pp2 = (nombre['nbdc2'].add(nombre['nbdrc2'])).div(quantite['Traitement Quantité'])
    pp3 = (nombre['nbdc3'].add(nombre['nbdrc3'])).div(quantite['Traitement Quantité'])
    pp4 = (nombre['nbdc4'].add(nombre['nbdrc4'])).div(quantite['Traitement Quantité'])
    pp1.index = temps['mois']
    pp2.index = temps['mois']
    pp3.index = temps['mois']
    pp4.index = temps['mois']
    
    for process in range( 0, 5):
        data = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_data_kpi_process_" + str(process))
        data["Durée moyenne avant clôture (demande de ressource)"][-nb-1:-1].fillna(float(0), inplace = True)
        data["Durée moyenne avant clôture (demande de ressource)"]=data["Durée moyenne avant clôture (demande de ressource)"].astype(np.float64)
        dfd1 = data['Durée moyenne avant clôture (demande de ressource)'][-nb:]
        dfd2 = data['Durée moyenne avant clôture (demande de ressource)'][-nb-1:-1]
        dfd1.index = data['date'][-nb:]
        dfd2.index = data['date'][-nb:]
        dfd = dfd1.div(dfd2).sub(1)
        dfd.fillna(float(0), inplace = True)
        
           
        data["Durée moyenne avant clôture (incident)"][-nb-1:-1].fillna(float(0), inplace = True)
        data["Durée moyenne avant clôture (incident)"]=data["Durée moyenne avant clôture (incident)"].astype(np.float64)
        dfi1 = data['Durée moyenne avant clôture (incident)'][-nb:]
        dfi2 = data['Durée moyenne avant clôture (incident)'][-nb-1:-1]
        dfi1.index = data['date'][-nb:]
        dfi2.index = data['date'][-nb:]
          
        dfi = dfi1.div(dfi2).sub(1)
        dfi.fillna(float(0), inplace = True)
          
        df["Avant_Cloture_Ressource_"+str(process)] = dfd
        df["Avant_Cloture_Support_"+str(process)] = dfi
        df["Avant_Cloture_Ressource_"+str(process)] = df["Avant_Cloture_Ressource_"+str(process)].replace([np.inf, -np.inf],np.float64(0),regex=True)
        df["Avant_Cloture_Support_"+str(process)] = df["Avant_Cloture_Support_"+str(process)].replace([np.inf, -np.inf],np.float64(0),regex=True)
    
    
    df["Avant_Cloture_Ressource_0"] = (df["Avant_Cloture_Ressource_1"].mul(pp1))\
                                    .add((df["Avant_Cloture_Ressource_2"].mul(pp2)))\
                                    .add((df["Avant_Cloture_Ressource_3"].mul(pp3)))\
                                    .add((df["Avant_Cloture_Ressource_4"].mul(pp4)))
    df["Avant_Cloture_Support_0"] = (df["Avant_Cloture_Support_1"].mul(pp1))\
                                    .add((df["Avant_Cloture_Support_2"].mul(pp2)))\
                                    .add((df["Avant_Cloture_Support_3"].mul(pp3)))\
                                    .add((df["Avant_Cloture_Support_4"].mul(pp4)))            
    
    do.to_sql(df,db="helpdesk_kpi_tables_delta_avant_cloture",mode=mode)
       
    return df
            

def traitementdelta():
    """
    Cette fonction calcul et écrit en base la performance mensuelle de
    traitement pour chaque processus, ainsi que pour chaque type.
    """
    # Calcul des deltas des avant clôtures
    df = pd.DataFrame()
    temps = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_tt")
    nombre = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_quantite")
    quantite = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_data_kpi_process_0")
    pp1 = (nombre['nbdc1'].add(nombre['nbdrc1'])).div(quantite['Traitement Quantité'])
    pp2 = (nombre['nbdc2'].add(nombre['nbdrc2'])).div(quantite['Traitement Quantité'])
    pp3 = (nombre['nbdc3'].add(nombre['nbdrc3'])).div(quantite['Traitement Quantité'])
    pp4 = (nombre['nbdc4'].add(nombre['nbdrc4'])).div(quantite['Traitement Quantité'])
    pp1.index = temps['mois']
    pp2.index = temps['mois']
    pp3.index = temps['mois']
    pp4.index = temps['mois']
    
    for process in range( 1, 5):
        data = ((temps['Ttnbdc'+str(process)].mul(nombre['nbdc'+str(process)])).add((temps['Ttnbdrc'+str(process)]).mul(nombre['nbdrc'+str(process)]))).div(nombre['nbdc'+str(process)].add(nombre['nbdrc'+str(process)]))
        dfd1 = data[-nb:]
        dfd2 = data[-nb-1:-1]
        dfd1.index = temps['mois'][-nb:]
        dfd2.index = temps['mois'][-nb:]
        
        dfd = dfd1.div(dfd2).sub(1)
        dfd.fillna(float(0), inplace = True)
        data = ((temps['Ttnbic'+str(process)].mul(nombre['nbic'+str(process)])).add((temps['Ttnbirc'+str(process)]).mul(nombre['nbirc'+str(process)]))).div(nombre['nbic'+str(process)].add(nombre['nbirc'+str(process)]))
        dfi1 = data[-nb:]
        dfi2 = data[-nb-1:-1]
        dfi1.index = temps['mois'][-nb:]
        dfi2.index = temps['mois'][-nb:]
        
        dfi = dfi1.div(dfi2).sub(1)
        dfi.fillna(float(0), inplace = True)
          
        df["Traitement_Ressource_"+str(process)] = dfd
        df["Traitement_Support_"+str(process)] = dfi
        df.index.name = 'date'

    

    df["Traitement_Ressource_0"] = (df["Traitement_Ressource_1"].mul(pp1))\
                                    .add((df["Traitement_Ressource_2"].mul(pp2)))\
                                    .add((df["Traitement_Ressource_3"].mul(pp3)))\
                                    .add((df["Traitement_Ressource_4"].mul(pp4)))
    df["Traitement_Support_0"] = (df["Traitement_Support_1"].mul(pp1))\
                                    .add((df["Traitement_Support_2"].mul(pp2)))\
                                    .add((df["Traitement_Support_3"].mul(pp3)))\
                                    .add((df["Traitement_Support_4"].mul(pp4)))            
    do.to_sql(df,db="helpdesk_kpi_tables_delta_traitement",mode=mode)
    return df


def note():
    """
    Cette fonction calcul et écrit en base la performance mensuelle.
    """
    traitement = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_delta_traitement")
    #print(traitement0)
    resolution = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_delta_resolution")
    #print(resolution0)
    avant_cloture = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_delta_avant_cloture")
    attente_N1 = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_delta_attente_niveau_1")
    
    traitementR= traitement['Traitement_Ressource_0']
    traitementS= traitement['Traitement_Support_0']
    resolutionR = resolution['Resolution_Ressource_0']
    resolutionS = resolution['Resolution_Support_0']
    avant_clotureR = avant_cloture['Avant_Cloture_Ressource_0']
    avant_clotureS = avant_cloture['Avant_Cloture_Support_0']
    attente_N1R = attente_N1['Attente_N1_Ressource_0']
    attente_N1S = attente_N1['Attente_N1_Support_0']

    noteS = (traitementS.add(resolutionS).add(avant_clotureS).add(attente_N1S)).div(4)
    noteR = (traitementR.add(resolutionR).add(avant_clotureR).add(attente_N1R)).div(4)
    note = (noteS+noteR)/2
    df = pd.DataFrame()
    
    df['note']=note
    df['note'] = df['note'].astype(np.float64)
    df['noteR']=noteR
    df['noteR'] = df['noteR'].astype(np.float64)
    df['noteS']=noteS
    df['noteS'] = df['noteS'].astype(np.float64)

    df.index = traitement['date']
    #print('Note ressource',noteR)
    #print('Note support',noteS)
    #print('note', df)
    do.to_sql(df,db="helpdesk_kpi_tables_note",mode=mode)
