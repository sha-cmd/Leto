# -*- coding: utf-8 -*-
"""
Created on Fri Jun 19 09:16:32 2020

@author: rboyrie
"""
from machine import do
import pandas as pd
import numpy as np
#import nltk
#import re
#import collections  # 1.5
#import multiprocessing as mp  # 1.2

def text_concat(row):
    if (row['id_statut_demande']==6):
        if (row['str_type'] == 'Incident'):
            if row['description_client_action_impossible_incident'] == None:
                row['description_client_action_impossible_incident']  = ''
            if row['description_client_incident'] == None:
                row['description_client_incident'] = ''
            return row['description_client_action_impossible_incident'] + ' ' + row['description_client_incident']
        if (row['str_type'] == 'Ressource'):
            if row['commentaire_du_client_pour_ressource'] == None:
                row['commentaire_du_client_pour_ressource'] = ''
            return row['commentaire_du_client_pour_ressource'] 
    else:
        return None

def stop_word(w, stop_words):
    try:
        string = [x for x in w if x not in stop_words]
        return " ".join(string)
    except:
        return ""

def make(limit_n1, limit_resolution, limit_treatment):
    df = do.r_to_b(query="SELECT * FROM synopsis")
    df = df.set_index(df['index'])
    df = df.drop('index',axis=1)
    #print(df)
    df['year_int'] = pd.to_datetime(df['date_cloture']).dt.year
    df['month_int'] = pd.to_datetime(df['date_cloture']).dt.month
    df['Text_client'] = df.apply(lambda row: text_concat(row), axis=1)
    ## Chargement de la liste de mot vide
    mot_vide = pd.read_csv('data\mot_vide.txt', sep="\r\n", header=None)
    liste_mot_vide = list(mot_vide[0]) # L’accesseur 0 désigne le nom de la colonne de la DATAFRAME
    ## Création des colonnes toutes durées et durée spécifique s m l (short medium long)
    df['Mot_Propre'] = df.loc[df['id_statut_demande']==6]\
        ['Text_client'].str.lower().str.replace(r'\s|\d|!?:,;', ' ', case=False)\
        .str.split(r"\W[\W']*\W|\W").dropna()
    df['Mot_Propre'] = df['Mot_Propre'].apply(lambda w: stop_word(w, liste_mot_vide))
    df['Mot_s'] = df.loc[df['treatment_time']<=2.5]['Mot_Propre']
    df['Mot_m'] = df.loc[(df['treatment_time']>2.5)&(df['treatment_time']<=16)]['Mot_Propre']
    df['Mot_l'] = df.loc[df['treatment_time']>16]['Mot_Propre']
    df['words_treatment_outliers'] = (df.loc[df['treatment_time'] > limit_treatment]['Mot_Propre'])
    df['words_resolution_outliers'] = (df.loc[df['Time_of_resolution'] > limit_resolution]['Mot_Propre'])
    df['words_n1_outliers'] = (df.loc[df['attente_n1_time'] > limit_n1]['Mot_Propre'])

    ## Calcul de la fréquence des mots sur l’ensemble des demandes clients
    process_corpus = []
    for index, row in df.loc[(df['id_statut_demande']==6)].iterrows():
        try:
            process_corpus.extend(tuple(row['Mot_Propre']))
        except:
            continue 
    #word_counts = collections.Counter(process_corpus)
    # Création de la liste des mots peu courant.
    #uncommon_words = word_counts.most_common()[:5800:-1]
    # Élimination des mots vides de la colonne pour toutes durées
    #df['Mot_Propre'] = df['Mot_Propre'].apply(lambda w: stop_word(w, uncommon_words)) 
    #df['Mot_s'] = df['Mot_s'].apply(lambda w: stop_word(w, uncommon_words)) 
    #df['Mot_m'] = df['Mot_m'].apply(lambda w: stop_word(w, uncommon_words)) 
    #df['Mot_l'] = df['Mot_l'].apply(lambda w: stop_word(w, uncommon_words)) 
    # Création d’une colonne de date compatible avec Power BI
    # Enlever le jour 1 et le laisser tel qu’il est dans la date de cloture
    df['date_cloture_coherente']=df['date_cloture'].dropna().apply(lambda x: pd.to_datetime(x).replace(hour=0,minute =0, second= 0))
    df_box_and_whiskers = df.loc[df['id_statut_demande']==6][['id_demande','process_label','str_type','attente_n1_time','treatment_time','Time_of_resolution','date_cloture_coherente']]
    df_box_and_whiskers.to_excel('box_and_whiskers.xlsx')
    do.to_sql(df_box_and_whiskers, db='box_and_whiskers')
    df = df.loc[df['id_statut_demande']==6][['id_demande','process_label','str_type',
                                'words_n1_outliers',
                                'words_resolution_outliers',
                                'words_treatment_outliers',
                                'Mot_s',
                                'Mot_m',
                                'Mot_l',
                                'date_cloture_coherente'
                                ]]

    df.to_excel('tag_cloud.xlsx')
    do.to_sql(df,db='tag_cloud')
