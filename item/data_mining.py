# -*- coding: utf-8 -*-
"""
Created on Mon May 11 08:07:49 2020

Ce fichier pourrait servir à la réalisation de table jde réalisation par
technicien.
En index le nom des techniciens et en colonnes les statuts et les actions,
et le nom du process et le type de ticket. Ce serait le tableau de score par 
technicien.

La fonction process_label modifie un peu les résultats d’identification du
process écrient dans le fichier data.xlsx, pour les besoins de la représenta-
-tion graphique dans Power BI. Plus de détail dans la fonction elle-même.

@author: rboyrie
"""
import pandas as pd
import matplotlib.pyplot as plt
from machine import do
import datetime

def do_list():
    dbpec = do.r_to_b(query = "SELECT id_demande, date_prise_en_charge,"+
                      "date_fin_prise_en_charge, id_technicien FROM prise_en"+
                      "_charge")
    dbpec = dbpec.dropna()
    
    
    nametech = {
    '1':'A01',
    '2':'A02',
    '3':'A03',
    '4':'A04',
    '5':'A05',
    '6':'A06',
    '7':'A07',
    '8':'A08',
    '9':'A09',
    '10':'A10',
    '11':'A11',
    '12':'A12',
    '13':'A13',
    '14':'A14',
    '15':'A15',
    '16':'A16',
    '17':'A17',
    '18':'A18',
    '19':'A19',
    '20':'A20',
    '21':'A21',
    '22':'A22',
    '23':'A23',
    '24':'A24',
    '25':'A25',
    '26':'A26',
    '27':'A27',
    '28':'A28',
    '29':'A29',
    '30':'A30',
    '31':'A31',
    '32':'A32',
    '33':'A33',
    '34':'A34',
    '35':'A35',
    '36':'A36',
    '37':'A37',
    '38':'A38',
    '39':'A39'
    }
    str_conceptnameaction = {
            '1':'Résoudre',
            '2':'Escalader',
            '3':'Transférer',
            '4':'En cours'
            }
    str_conceptname = {
            '1':'Att N1',
            '2':'Att N2',
            '3':'En cours N1',
            '4':'En cours N2',
            '5':'Résolu',
            '6':'Clôturé'
            }
    dbpec['id_technicien']=dbpec['id_technicien'].astype(str)
    dbpec['id_technicien'] = dbpec['id_technicien'].map(nametech)
    dbpec.index = dbpec['id_demande']
    dbpec = dbpec.drop('id_demande',1)
    print(dbpec)
    dbpec = dbpec.loc[dbpec['date_prise_en_charge'] >= pd.datetime(2020,4,1)]
    print(dbpec)
    dbpec.to_csv('PEC_process.csv')
    #print(dbpec)
    
    dbsuivis = do.r_to_b(query = "SELECT id_demande, date_suivi,"+
                      " id_statut_demande, id_nature_action, id_technicien_"+
                      "creation FROM suivi_intervention")
    dbsuivis = dbsuivis.dropna()
    dbsuivis['id_technicien_creation']=dbsuivis['id_technicien_creation'].astype(str)
    dbsuivis['id_technicien_creation'] = dbsuivis['id_technicien_creation'].map(nametech)
    dbsuivis['id_statut_demande'] = dbsuivis['id_statut_demande'].astype(str)
    dbsuivis['id_statut_demande'] = dbsuivis['id_statut_demande'].map(str_conceptname)
    dbsuivis['id_nature_action'] = dbsuivis['id_nature_action'].astype(str)
    dbsuivis['id_nature_action'] = dbsuivis['id_nature_action'].map(str_conceptnameaction)
    dbsuivis.index = dbsuivis['id_demande']
    dbsuivis = dbsuivis.drop('id_demande',1)
    dbsuivis = dbsuivis.loc[dbsuivis['date_suivi']>=pd.datetime(2020,4,1)]
    dbsuivis.to_csv('Suivis_process.csv')
    
    dbsuivir = do.r_to_b(query = "SELECT id_demande, date_suivi,"+
                      " id_statut_demande, id_nature_action, id_technicien"+
                      " FROM suivi_demande_ressource")
    dbsuivir = dbsuivir.dropna()
    dbsuivir['id_technicien']=dbsuivir['id_technicien'].astype(str)
    dbsuivir['id_technicien'] = dbsuivir['id_technicien'].map(nametech)
    dbsuivir['id_statut_demande'] = dbsuivir['id_statut_demande'].astype(str)
    dbsuivir['id_statut_demande'] = dbsuivir['id_statut_demande'].map(str_conceptname)
    dbsuivir['id_nature_action'] = dbsuivir['id_nature_action'].astype(str)
    dbsuivir['id_nature_action'] = dbsuivir['id_nature_action'].map(str_conceptnameaction)
    dbsuivir.index = dbsuivir['id_demande']
    dbsuivir = dbsuivir.drop('id_demande',1)
    dbsuivir= dbsuivir.loc[dbsuivir['date_suivi']>=pd.datetime(2020,4,1)]
    dbsuivir.to_csv('Suivir_process.csv')
    #print(dbsuivir)
    #print(dbsuivis['id_statut_demande'])
    dbressourcetoadd = do.r_to_b(query="SELECT demande.id_demande, date_creation_demande FROM demande WHERE demande.id_demande IN (SELECT id_demande FROM suivi_demande_ressource WHERE id_statut_demande = 6)")
    dbressourcetoadd['concept:name']="Création"
    dbressourcetoadd.columns = ['case:concept:name', 'time:timestamp','concept:name']
    dbressourcetoadd = dbressourcetoadd.loc[dbressourcetoadd['time:timestamp']>=pd.datetime(2020,4,1)]
    
    dbinterventiontoadd = do.r_to_b(query="SELECT demande.id_demande, date_creation_demande FROM demande WHERE demande.id_demande IN (SELECT id_demande FROM suivi_intervention WHERE id_statut_demande = 6)")
    dbinterventiontoadd['concept:name']="Création"
    dbinterventiontoadd.columns = ['case:concept:name', 'time:timestamp','concept:name']
    dbinterventiontoadd = dbinterventiontoadd.loc[dbinterventiontoadd['time:timestamp']>=pd.datetime(2020,4,1)]
    
    #print(dbinterventiontoadd)
    dbressource = do.r_to_b(query="SELECT suivi_demande_ressource.id_demande caseconceptname, id_statut_demande conceptname, date_prise_en_charge timetimestamp FROM (SELECT * FROM prise_en_charge WHERE prise_en_charge.id_demande IN (SELECT id_demande FROM suivi_demande_ressource WHERE id_statut_demande = 6)) as table1 INNER JOIN suivi_demande_ressource ON table1.id_demande = suivi_demande_ressource.id_demande ")
    
    dbressource['conceptname'] = dbressource['conceptname'].astype(str)
    dbressource['conceptname'] = dbressource['conceptname'].map(str_conceptname)
    dbressource.columns = ['case:concept:name', 'concept:name','time:timestamp']
    dbressource = dbressource.loc[dbressource['time:timestamp']>=pd.datetime(2020,4,1)]
    
    
    dbintervention = do.r_to_b(query="SELECT suivi_intervention.id_demande caseconceptname, id_statut_demande conceptname, date_prise_en_charge timetimestamp FROM (SELECT * FROM prise_en_charge WHERE prise_en_charge.id_demande IN (SELECT id_demande FROM suivi_intervention WHERE id_statut_demande = 6)) as table1 INNER JOIN suivi_intervention ON table1.id_demande = suivi_intervention.id_demande ")
    
    dbintervention['conceptname'] = dbintervention['conceptname'].astype(str)
    dbintervention['conceptname'] = dbintervention['conceptname'].map(str_conceptname)
    dbintervention.columns = ['case:concept:name', 'concept:name','time:timestamp']
    dbintervention = dbintervention.loc[dbintervention['time:timestamp']>=pd.datetime(2020,4,1)]
    
    # Concatenation des colonnes
    db = pd.DataFrame(dbintervention.append(dbressource.append(dbressourcetoadd.append(dbinterventiontoadd))))
    db.index=db['case:concept:name']
    db.columns = ['case', 'concept:name','time:timestamp']
    db = db.drop("case", 1)
    db.to_csv('data_for_mining.csv')
    print(db)

def label_type(row):
    data = pd.read_excel('data.xlsx')
    typ = data.loc[data['demande'] == row['id_demande']]['type']
    return typ.iloc[-1]

def label_str_type(row):
    data = pd.read_excel('data.xlsx')
    str_type = data.loc[data['demande'] == row['id_demande']]['str_type']
    return  str_type.iloc[-1]

def label_process(row, data):
    """
    Dans cette fonction nous analysons le nom du process. S’il est "pas de
    process clair" c’est qu’il fini par un r. Nous le remplaçons alors pas un
    process 3, palliant le défaut de la précédente fonction ayant créée le 
    fichier data.xlsx.
    """
    process = data.loc[data['demande'] == row['id_demande']]['Process']
    try:
        if process.iloc[-1][-1:] == 'r':
            return int(3)
        else:
            return int(process.iloc[-1][-1:])
    except:
        return None

def label_pec_deb(row, data):
    pec_deb = data.loc[data['id_prise_en_charge'] == row['id_prise_en_charge']]['date_prise_en_charge']
    try:
        return pec_deb.iloc[-1]
    except:
        return None

def label_pec_end(row, data):
    pec_end = data.loc[data['id_prise_en_charge'] == row['id_prise_en_charge']]['date_fin_prise_en_charge']
    try:
        return pec_end.iloc[-1]
    except:
        return None

def label_tech(row):
    nametech = {
    '1':'A01',
    '2':'A02',
    '3':'A03',
    '4':'A04',
    '5':'A05',
    '6':'A06',
    '7':'A07',
    '8':'A08',
    '9':'A09',
    '10':'A10',
    '11':'A11',
    '12':'A12',
    '13':'A13',
    '14':'A14',
    '15':'A15',
    '16':'A16',
    '17':'A17',
    '18':'A18',
    '19':'A19',
    '20':'A20',
    '21':'A21',
    '22':'A22',
    '23':'A23',
    '24':'A24',
    '25':'A25',
    '26':'A26',
    '27':'A27',
    '28':'A28',
    '29':'A29',
    '30':'A30',
    '31':'A31',
    '32':'A32',
    '33':'A33',
    '34':'A34',
    '35':'A35',
    '36':'A36',
    '37':'A37',
    '38':'A38',
    '39':'A39'
    }
    tech = ""
    try:
        tech = nametech[str(row['id_technicien'])]
    except:
        tech = nametech[str(int(row['Tech']))]
    return tech

def treatment_time_selecting(row, data):
    time = data.loc[data['demande'] == row['id_demande']]['Tt']
    #print(time)
    try:
        if ((row['id_statut_demande']==6)):
            return time.iloc[-1]/3600000000000
        else :
            return None
    except:
        return None

def attente_treatment_time(row, data):
    time = data.loc[data['demande'] == row['id_demande']]['T1']
    #print(time)
    # Code enlevé de la condition if pour mettre la ligne sur le suivi au statut valant 6
    # &(row['a_ete_d_abord_resolu']==False))|((row['id_statut_demande']==5)&(row['a_ete_d_abord_resolu']==True)
    try:
        if ((row['id_statut_demande']==6)):
            return time.iloc[-1]/3600000000000
        else :
            return None
    except:
        return None

def attente_T2_treatment_time(row, data):
    time = data.loc[data['demande'] == row['id_demande']]['T2']
    #print(time)
    try:
        if ((row['id_statut_demande']==6)):
            return time.iloc[-1]/3600000000000
        else :
            return None
    except:
        return None

def attente_T3_treatment_time(row, data):
    time = data.loc[data['demande'] == row['id_demande']]['T3']
    #print(time)
    try:
        if ((row['id_statut_demande']==6)):
            return time.iloc[-1]/3600000000000
        else :
            return None
    except:
        return None

def treatment_time(row, data):
    time = data.loc[data['demande'] == row['id_demande']]['Tt']
    #print(time)
    try:
            return time.iloc[-1]
    except:
        return None

def label_resolu(row, data):
    a_une_resolution = data.loc[data['demande'] == row['id_demande']]['resolu']
    #print(time)
    try:
        return a_une_resolution.iloc[-1]
    except:
        return None

def label_tech_work_done(row, data):
    
    a_un_tech_de_resolution = data.loc[data['demande'] == row['id_demande']]['Tech']
    try:
        return a_un_tech_de_resolution.iloc[-1]
    except:
        return None
    
def label_time_of_resolution(row, data):
   
    time_of_resolution = data.loc[data['demande'] == row['id_demande']]['Times']/3600000000000
    try:
        if ((row['id_statut_demande']==6)):
            return time_of_resolution.iloc[-1]
        else :
            return None
    except:
        return None

def label_descript_incid(row, data):
    
    descript_incid = data.loc[data['id_demande'] == row['id_demande']]['description_incident']
    try:
        if ((row['id_statut_demande']==6)):
            return descript_incid.iloc[-1]
        else :
            return None
    except:
        return None
    
def label_descript_action(row, data):
    
    descript_action = data.loc[data['id_demande'] == row['id_demande']]['description_action_impossible']
    try:
        if ((row['id_statut_demande']==6)):
            return descript_action.iloc[-1]
        else :
            return None
    except:
        return None

def label_commentaire(row, data):
    
    commentaire = data.loc[data['id_demande'] == row['id_demande']]['commentaire']
    try:
        if ((row['id_statut_demande']==6)):
            return commentaire.iloc[-1]
        else :
            return None
    except:
        return None

def label_com_suivi_ressource(row, data):
    
    commentaire = data.loc[data['id_suivi_demande_ressource'] == row['id_suivi']]['commentaire']
    try:
        return commentaire.iloc[-1]
    except:
        return None
    
def label_descript_diagnost(row, data):
    
    descript_diagnost = data.loc[data['id_suivi_intervention'] == row['id_suivi']]['description_diagnostic']
    try:
        return descript_diagnost.iloc[-1]
    except:
        return None
    
def label_action_a_mener(row, data):
    
    action_a_mener = data.loc[data['id_suivi_intervention'] == row['id_suivi']]['action_a_mener']
    try:
        return action_a_mener.iloc[-1]
    except:
        return None
    
def label_email_demandeur(row,data):
    
    email_demandeur = data.loc[data['id_demande'] == row['id_demande']]['email_demandeur']
    try:
        if ((row['id_statut_demande']==6)):
            return email_demandeur.iloc[-1]
        else :
            return None
    except:
        return None
    
def label_cloture(row, data):
    
    date_de_cloture = data.loc[data['demande'] == row['id_demande']]['date_de_cloture']
    try:
        return date_de_cloture.iloc[-1]
    except:
        return None

def label_creation(row, data):
    
    date_de_cloture = data.loc[data['demande'] == row['id_demande']]['creation']
    try:
        return date_de_cloture.iloc[-1]
    except:
        return None    

def label_urgence(row, data):
    
    urgence = data.loc[data['demande'] == row['id_demande']]['urgence']
    try:
        return urgence.iloc[-1]
    except:
        return None    

def synopsis():
    """
    Création du document Synopsis Continuum Tessera Tabular ou vue panoramique
    des tickets. Ce document forme un socle intéressant pour le management, et
    aussi une perspective très pertinente pour le machine learning non 
    supervisé. Il contient aussi l’indication de résolution, il est magique.
    """
    intervention = do.r_to_b(query="SELECT * FROM intervention")
    demande_ressource = do.r_to_b(query="SELECT * FROM demande_ressource")
    demande = do.r_to_b(query="SELECT * FROM demande")
    suivi_intervention = do.r_to_b(query="SELECT * FROM suivi_intervention")
    suivi_demande_ressource = do.r_to_b(query="SELECT * FROM suivi_demande_ressource")
    
    
    si = pd.read_excel('suivi_intervention.xlsx')   
    sd  = pd.read_excel('suivi_ressource.xlsx')  
    

    data = pd.read_excel('data.xlsx')   
    data['Times'] = data['Tr']
    data['Times'] = data['Times'].replace(0,data['Trc'])
    data['Tech'] = data['Tech_res']
    data['Tech'] = data['Tech'].replace(0,data['Tech_clos'])
    to_drop = ['None']
    data = data[~data['date_de_cloture'].isin(to_drop)]
    ### Suivi Intervention
    si['description_tech_diagnostic'] = si.apply(lambda row: label_descript_diagnost(row, suivi_intervention), axis=1)
    si['action_du_tech_a_mener'] = si.apply(lambda row: label_action_a_mener(row, suivi_intervention), axis=1)
    sd['commentaire_du_tech_pour_ressource'] = sd.apply(lambda row: label_com_suivi_ressource(row, suivi_demande_ressource), axis=1)
    si['a_ete_d_abord_resolu'] = si.apply (lambda row: label_resolu(row, data), axis=1)
    sd['a_ete_d_abord_resolu'] = sd.apply (lambda row: label_resolu(row, data), axis=1)
    si['email_demandeur'] = si.apply(lambda row: label_email_demandeur(row, demande), axis=1)
    sd['email_demandeur'] = sd.apply(lambda row: label_email_demandeur(row, demande), axis=1)
    ### Suivi demande ressource
    si['description_client_incident'] = si.apply (lambda row: label_descript_incid(row, intervention), axis=1)
    si['description_client_action_impossible_incident'] = si.apply (lambda row: label_descript_action(row, intervention), axis=1)
    sd['commentaire_du_client_pour_ressource'] = sd.apply (lambda row: label_commentaire(row, demande_ressource), axis=1)
    si['Tech_de_resolution'] = si.apply (lambda row: label_tech_work_done(row, data), axis=1)
    si['Tech_de_resolution'] = si.apply (lambda row: label_tech(row), axis=1)
    sd['Tech_de_resolution'] = sd.apply (lambda row: label_tech_work_done(row, data), axis=1)
    sd['Tech_de_resolution'] = sd.apply (lambda row: label_tech(row), axis=1)
    si['Time_of_resolution'] = si.apply (lambda row: label_time_of_resolution(row, data), axis=1)
    sd['Time_of_resolution'] = sd.apply (lambda row: label_time_of_resolution(row, data), axis=1)
    si['treatment_time'] = si.apply (lambda row: treatment_time_selecting(row, data), axis=1)
    sd['treatment_time'] = sd.apply (lambda row: treatment_time_selecting(row, data), axis=1)
    si['attente_n1_time'] = si.apply (lambda row: attente_treatment_time(row, data), axis=1)
    sd['attente_n1_time'] = sd.apply (lambda row: attente_treatment_time(row, data), axis=1)
    si['attente_t2_time'] = si.apply (lambda row: attente_T2_treatment_time(row, data), axis=1)
    sd['attente_t2_time'] = sd.apply (lambda row: attente_T2_treatment_time(row, data), axis=1)
    si['attente_t3_time'] = si.apply (lambda row: attente_T3_treatment_time(row, data), axis=1)
    sd['attente_t3_time'] = sd.apply (lambda row: attente_T3_treatment_time(row, data), axis=1)
    si['date_cloture'] = si.apply(lambda row: label_cloture(row, data), axis=1)
    sd['date_cloture'] = sd.apply(lambda row: label_cloture(row, data), axis=1)
    si = si.dropna(axis=0, how="any", subset=['date_cloture'], inplace=False)
    sd = sd.dropna(axis=0, how="any", subset=['date_cloture'], inplace=False)
    si['date_cloture_coherente']=si['date_cloture'].apply(lambda x: pd.to_datetime(x).replace(hour=0,minute =0, second= 0))
    sd['date_cloture_coherente']=sd['date_cloture'].apply(lambda x: pd.to_datetime(x).replace(hour=0,minute =0, second= 0))
    si['date_creation'] = si.apply(lambda row: label_creation(row, data), axis=1)
    sd['date_creation'] = sd.apply(lambda row: label_creation(row, data), axis=1)
    si['date_creation_coherente']=si['date_creation'].apply(lambda x: pd.to_datetime(x).replace(hour=0,minute =0, second= 0))
    sd['date_creation_coherente']=sd['date_creation'].apply(lambda x: pd.to_datetime(x).replace(hour=0,minute =0, second= 0))
    si['urgence'] = si.apply(lambda row: label_urgence(row, data), axis=1)
    sd['urgence'] = sd.apply(lambda row: label_urgence(row, data), axis=1)
    #si = si.loc[((si['id_statut_demande']==6)&(si['a_ete_d_abord_resolu']==False))|\
     #           ((si['id_statut_demande']==5)&(si['a_ete_d_abord_resolu']==True))]
    #sd = sd.loc[((sd['id_statut_demande']==6)&(sd['a_ete_d_abord_resolu']==False))|\
     #           ((sd['id_statut_demande']==5)&(sd['a_ete_d_abord_resolu']==True))]
    
    #si.index = si['id_demande']
    #si = si.drop(['id_demande'],axis=1)
    si = si.drop(['tech_label'],axis=1)
    #sd.index = sd['id_demande']
    #sd = sd.drop(['id_demande'],axis=1)  
    sd = sd.drop(['tech_label'],axis=1)
    df = si.append(sd)
    df['an']=pd.to_datetime(df['date_cloture']).dt.year
    df['mois']=pd.to_datetime(df['date_cloture']).dt.month
    #df = df.sort_index(axis=0)
    #df['id_demande'] = df.index
    df = df.reset_index()
    df = df.sort_values(by=['an','mois','str_type','process_label','id_demande','id_suivi'])
    
    df = df.reset_index()
    #df = df.drop(['number_demande'],axis=1)
    #si['pec_deb'] = si.apply (lambda row: label_pec_deb(row, data), axis=1)
    df = df[['id_demande','Tech_de_resolution','id_technicien','email_demandeur','a_ete_d_abord_resolu',
             'process_label','id_groupe','urgence','str_type','type',
             'commentaire_du_client_pour_ressource','description_client_action_impossible_incident',
             'description_client_incident','description_tech_diagnostic','commentaire_du_tech_pour_ressource',
             'action_du_tech_a_mener','attente_n1_time','attente_t2_time','attente_t3_time','Time_of_resolution','treatment_time','id_suivi','id_statut_demande',
             'id_prise_en_charge','id_nature_action','date_suivi','pec_deb','pec_end','date_creation','date_cloture','date_cloture_coherente','date_creation_coherente']]
    df.to_excel('Synopsis Continuum Tessera Tabular - vue panoramique des tickets.xlsx')
    df.to_csv('Synopsis Continuum Tessera Tabular - vue panoramique des tickets.csv')
    print('base')
    do.to_sql(df,db='synopsis')


def outliers(mstart,mend,ystart, yend, limit_n1, limit_resolution, limit_treatment):
    """
    À terminer
    Constate les outliers, notes le id du ticket et décompte le nombre dans
    chaque catégorie.
    """
    synopsis = do.r_to_b(query="SELECT * FROM synopsis")
    dfFinal = pd.DataFrame()
    for year in range (ystart, yend+1):
        for month in range( mstart, 12+1):
            #data = pd.read_excel('suivi_ressource.xlsx')
            if (year==yend) & (month == mend+1):
                break
            if month != 12:
                data = synopsis.loc[(pd.to_datetime(synopsis['date_cloture'])>=pd.datetime(year,month,1))&
                    (synopsis['id_statut_demande']==6)
                    &(pd.to_datetime(synopsis['date_cloture'],format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]
                
            else:
                data = synopsis.loc[(pd.to_datetime(synopsis['date_cloture'])>=pd.datetime(year,month,1))&
                                (synopsis['id_statut_demande']==6)
                            &(pd.to_datetime(synopsis['date_cloture'],format='%Y-%m-%d')<pd.datetime(year+1,1,1))]
                

            # utilisation des limites appliqués à la variable mensuelle de data
            treatment_outliers = (data.loc[data['treatment_time'] > limit_treatment])
            nb_treatment_outliers = treatment_outliers['treatment_time'].count()
            resolution_outliers = (data.loc[data['Time_of_resolution'] > limit_resolution])
            nb_resolution_outliers = resolution_outliers['Time_of_resolution'].count()
            n1_outliers = (data.loc[data['attente_n1_time'] > limit_n1])
            nb_n1_outliers = n1_outliers['attente_n1_time'].count()
            # transformation des id en liste pour chaque catégorie
            list_treatment_outliers = treatment_outliers['id_demande'].to_list()
            list_resolution_outliers = resolution_outliers['id_demande'].to_list()
            list_n1_outliers = n1_outliers['id_demande'].to_list()
            # conversion en string pour écriture possible en base par l’ORM Alchemy
            # Traitement
            string_treatment = [str(x) for x in list_treatment_outliers]
            integrable_list_treatment = ', '.join(string_treatment)
            # Résolution
            string_resolution = [str(x) for x in list_resolution_outliers]
            integrable_list_resolution = ', '.join(string_resolution)
            # Attente niveau 1
            string_n1 = [str(x) for x in list_n1_outliers]
            integrable_list_n1 = ', '.join(string_n1)

            ###########
            date = pd.datetime(year,month,1)
            ##########
            # Ajout des données calculées dans le cadre de données
            dfFinal = dfFinal.append({'date':date,
                            'nb_n1_outliers': nb_n1_outliers,
                            'nb_resolution_outliers': nb_resolution_outliers,
                            'nb_treatment_outliers': nb_treatment_outliers,
                            'liste_n1_id':integrable_list_n1,
                            'liste_resolution_id':integrable_list_resolution,
                            'liste_treatment_id':integrable_list_treatment
                            }, ignore_index=True)
    dfFinal = dfFinal.reset_index()
    dfFinal = dfFinal.drop(['index'],axis=1)
    dfFinal.to_excel('outliers.xlsx')
    do.to_sql(dfFinal,db='outliers')
    dfFinal.index = dfFinal['date']
    df= dfFinal.iloc[9:][['nb_n1_outliers', 'nb_resolution_outliers', 'nb_treatment_outliers']]
    fig, ax = plt.subplots(figsize=(10, 5))
    df.plot(ax=ax)
    ax.legend(["Niveaux 1", "Résolutions","Traitements"]);
    ax.set_ylabel('Nombre de tickets exclus')
    ax.set_xlabel('Mois')
    plt.savefig('exclusion_par_cretes.png')
    plt.show()
    print('Moyenne d’exclusion de N1: ',dfFinal.iloc[9:]['nb_n1_outliers'].mean())
    print('Moyenne d’exclusion de résolution: ', dfFinal.iloc[9:]['nb_resolution_outliers'].mean())
    print('Moyenne d’exclusion de traitement: ', dfFinal.iloc[9:]['nb_treatment_outliers'].mean())
    
    
    #return pd.Series(data_purified).div(3600000000000).mean()


def scores(mstart,mend,ystart, yend):
    """
    Cette fonction construit une table joignant les suivis et les process de
    tickets. En plus, elle joint les prises en charges et les temps de 
    traitement de tickets. Attention toutefois, il y a plusieurs fois le 
    même id_demande, et plusieurs le même temps de traitement. Les tickets
    peuvent-être triés par le statut_demande = 6
    """
    suivi_intervention = do.r_to_b(query="SELECT * FROM suivi_intervention")
    suivi_demande_ressource = do.r_to_b(query="SELECT * FROM suivi_demande_ressource")
    si = suivi_intervention.rename(columns={'id_technicien_creation':'id_technicien',\
                                'id_suivi_intervention':'id_suivi'})
    sd  = suivi_demande_ressource.rename(columns={'id_technicien_creation':'id_technicien',\
                                'id_suivi_demande_ressource':'id_suivi',\
                                })
    dataSI = pd.concat([\
                       si['id_suivi'], si['id_demande'],
                       si['id_prise_en_charge'],
                       si['id_statut_demande'],
                       si['id_nature_action'],si['id_groupe'],
                       si['id_technicien'],si['date_suivi']],
                        axis=1)
    dataSD = pd.concat([\
                       sd['id_suivi'], sd['id_demande'],
                       sd['id_prise_en_charge'],
                       sd['id_statut_demande'],
                       sd['id_nature_action'],sd['id_groupe'],
                       sd['id_technicien'],sd['date_suivi']],
                        axis=1)
    date1 = datetime.datetime.now()
    data = pd.read_excel('data.xlsx')
    dataSI['process_label'] = dataSI.apply (lambda row: label_process(row, data), axis=1)
    dataSI['date_de_cloture'] = dataSI.apply (lambda row: label_cloture(row, data), axis=1)
    date2 = datetime.datetime.now()
    print('durée de la phase lambda_process :',date2-date1)
    date1 = datetime.datetime.now()
    dataSI['treatment_time'] = dataSI.apply (lambda row: treatment_time(row,data), axis=1)
    date2 = datetime.datetime.now()
    print('durée de la phase lambda_treatment_time :',date2-date1)
    date1 = datetime.datetime.now()
    data = do.r_to_b(query="SELECT * FROM prise_en_charge")
    dataSI['pec_deb'] = dataSI.apply (lambda row: label_pec_deb(row, data), axis=1)
    dataSI['pec_end'] = dataSI.apply (lambda row: label_pec_end(row, data), axis=1)
    date2 = datetime.datetime.now()
    print('durée de la phase lambda_pec :',date2-date1)
    date1 = datetime.datetime.now()
    dataSI['tech_label'] = dataSI.apply (lambda row: label_tech(row), axis=1)
    date2 = datetime.datetime.now()
    print('durée de la phase lambda label_tech :',date2-date1)
    
    dataSI['type']= 1
    dataSI['str_type']= 'Incident'
    dataSI.index = dataSI['id_suivi']
    dataSI = dataSI.drop(['id_suivi'],axis=1)
    dataSI.to_excel('suivi_intervention.xlsx')
    dataSI.to_csv('suivi_intervention.csv')

    
    date1 = datetime.datetime.now()
    data = pd.read_excel('data.xlsx')
    dataSD['process_label'] = dataSD.apply (lambda row: label_process(row, data), axis=1)
    dataSD['date_de_cloture'] = dataSD.apply (lambda row: label_cloture(row, data), axis=1)
    #dataSD['process_label'] = dataSD.astype(int)
    date2 = datetime.datetime.now()
    print('durée de la phase label_process :',date2-date1)
    date1 = datetime.datetime.now()
    dataSD['treatment_time'] = dataSD.apply (lambda row: treatment_time(row,data), axis=1)
    date2 = datetime.datetime.now()
    print('durée de la phase treatment_time :',date2-date1)
    date1 = datetime.datetime.now()
    data = do.r_to_b(query="SELECT * FROM prise_en_charge")
    dataSD['pec_deb'] = dataSD.apply (lambda row: label_pec_deb(row, data), axis=1)
    dataSD['pec_end'] = dataSD.apply (lambda row: label_pec_end(row, data), axis=1)
    date2 = datetime.datetime.now()
    print('durée de la phase label_pec :',date2-date1)
    date1 = datetime.datetime.now()
    dataSD['tech_label'] = dataSD.apply (lambda row: label_tech(row), axis=1)
    date2 = datetime.datetime.now()
    print('durée de la phase label_tech :',date2-date1)
    dataSD['type']= 2
    dataSD['str_type']= 'Ressource'
    dataSD.index = dataSD['id_suivi']
    dataSD = dataSD.drop(['id_suivi'],axis=1)
    dataSD.to_excel('suivi_ressource.xlsx')
    dataSD.to_csv('suivi_ressource.csv')
    
    ## Création des scores par regroupement des colonnes suivant les tech
    ## et les process
    ## Scores pour les interventions.
    ## Ces scores sont mieux réaliser par la fonction regression
    date1 = datetime.datetime.now()
    dfFinal = pd.DataFrame()
    for year in range (ystart, yend+1):
        for month in range( mstart, 12+1):
            data = pd.read_excel('suivi_intervention.xlsx')            
            if (year==yend) & (month == mend+1):
                break
            if month != 12:
                data = data.loc[(pd.to_datetime(data['date_suivi'])>=pd.datetime(year,month,1))&
                                (data['id_statut_demande']==6)
                                &(pd.to_datetime(data['date_suivi'],format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]
            else:
                data = data.loc[(pd.to_datetime(data['date_suivi'])>=pd.datetime(year,month,1))&
                                (data['id_statut_demande']==6)
                                &(pd.to_datetime(data['date_suivi'],format='%Y-%m-%d')<pd.datetime(year+1,1,1))]
            df = data.groupby(['tech_label','process_label','id_demande']).count()
            df=df.rename(columns={'id_suivi':'Quantité de clôture','id_prise_en_charge':'Quantité de prise en charge'})
            df['date']=pd.datetime(year,month,1)
            df = df.reset_index()
            dfFinal = dfFinal.append(df)
    dfFinal[['tech_label','process_label','treatment_time','Quantité de clôture','date']].to_excel('scores_intervention.xlsx')
    do.to_sql(dfFinal[['tech_label','process_label','Quantité de clôture','date']],db='scores_intervention')
    ## Scores pour les ressources
    date1 = datetime.datetime.now()
    dfFinal = pd.DataFrame()
    for year in range (ystart, yend+1):
        for month in range( mstart, 12+1):
            data = pd.read_excel('suivi_ressource.xlsx')
            if (year==yend) & (month == mend+1):
                break
            if month != 12:
                data = data.loc[(pd.to_datetime(data['date_suivi'])>=pd.datetime(year,month,1))&
                    (data['id_statut_demande']==6)
                    &(pd.to_datetime(data['date_suivi'],format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]
            else:
                data = data.loc[(pd.to_datetime(data['date_suivi'])>=pd.datetime(year,month,1))&
                                (data['id_statut_demande']==6)
                            &(pd.to_datetime(data['date_suivi'],format='%Y-%m-%d')<pd.datetime(year+1,1,1))]

            df = data.groupby(['tech_label','process_label']).count()
            df=df.rename(columns={'id_suivi':'Quantité de clôture','id_prise_en_charge':'Quantité de prise en charge'})
            df['date']=pd.datetime(year,month,1)
            df = df.reset_index()
            dfFinal = dfFinal.append(df)
    dfFinal[['tech_label','process_label','treatment_time','Quantité de clôture','date']].to_excel('scores_ressource.xlsx')
    do.to_sql(dfFinal[['tech_label','process_label','treatment_time','Quantité de clôture','date']],db='scores_ressource')
    
    ## MOYENNE Par technicien, par mois pour les Interventions
    date1 = datetime.datetime.now()
    dfFinal = pd.DataFrame()
    for year in range (ystart, yend+1):
        for month in range( mstart, 12+1):
            data = pd.read_excel('suivi_intervention.xlsx')            
            if (year==yend) & (month == mend+1):
                break
            if month != 12:
                data = data.loc[(pd.to_datetime(data['date_suivi'])>=pd.datetime(year,month,1))&
                                (data['id_statut_demande']==6)
                                &(pd.to_datetime(data['date_suivi'],format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]
            else:
                data = data.loc[(pd.to_datetime(data['date_suivi'])>=pd.datetime(year,month,1))&
                                (data['id_statut_demande']==6)
                                &(pd.to_datetime(data['date_suivi'],format='%Y-%m-%d')<pd.datetime(year+1,1,1))]
            list_tech = data['tech_label'].unique()
            for tech in list_tech:
                if month != 12:
                    new_row = {'Tech':tech,'means':data.loc[(data['tech_label']==tech)&(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['type']==1)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]['treatment_time'].mean()/3600000000000,
                           'date':pd.datetime(year,month,1)}
                    dfFinal= dfFinal.append(new_row, ignore_index=True)
                else:
                    new_row = {'Tech':tech,'means':data.loc[(data['tech_label']==tech)&(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['type']==1)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year+1,1,1))]['treatment_time'].mean()/3600000000000,
                           'date':pd.datetime(year,month,1)}
                    dfFinal= dfFinal.append(new_row, ignore_index=True)
            
    dfFinal.to_excel('scores_trait_int_means.xlsx')
    do.to_sql(dfFinal,db='scores_trait_int_means')
    date2 = datetime.datetime.now()
    print('durée de la phase moyenne traitement pour intervention :',date2-date1)
    ## MOYENNE Par technicien, par mois pour les Ressources
    date1 = datetime.datetime.now()
    dfFinal = pd.DataFrame()
    for year in range (ystart, yend+1):
        for month in range( mstart, 12+1):
            data = pd.read_excel('suivi_ressource.xlsx')
            if (year==yend) & (month == mend+1):
                break
            if month != 12:
                data = data.loc[(pd.to_datetime(data['date_suivi'])>=pd.datetime(year,month,1))&
                    (data['id_statut_demande']==6)
                    &(pd.to_datetime(data['date_suivi'],format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]
            else:
                data = data.loc[(pd.to_datetime(data['date_suivi'])>=pd.datetime(year,month,1))&
                                (data['id_statut_demande']==6)
                            &(pd.to_datetime(data['date_suivi'],format='%Y-%m-%d')<pd.datetime(year+1,1,1))]

            list_tech = data['tech_label'].unique()
            for tech in list_tech:
                if month != 12:
                    new_row = {'Tech':tech,'means':data.loc[(data['tech_label']==tech)&(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['type']==2)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]['treatment_time'].mean()/3600000000000,
                           'date':pd.datetime(year,month,1)}
                    dfFinal= dfFinal.append(new_row, ignore_index=True)
                else:
                    new_row = {'Tech':tech,'means':data.loc[(data['tech_label']==tech)&(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['type']==2)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year+1,1,1))]['treatment_time'].mean()/3600000000000,
                           'date':pd.datetime(year,month,1)}
                    dfFinal= dfFinal.append(new_row, ignore_index=True)


    dfFinal.to_excel('scores_trait_res_means.xlsx')
    do.to_sql(dfFinal,db='scores_trait_res_means')
    date2 = datetime.datetime.now()
    print('durée de la phase moyenne traitement pour ressource :',date2-date1)
    
    ## MOYENNE Des Résolutions Par technicien, par mois pour les Types
    # Boucle pour faire par types les moyennes
    liste = [1,2]#Liste des types
    noms = ['scores_res_int_means','scores_res_res_means']
    for it, val in enumerate(liste):
        date1 = datetime.datetime.now()
        dfFinal = pd.DataFrame()            
        for year in range (ystart, yend+1):
            for month in range( mstart, 12+1):
                data = pd.read_excel('data.xlsx')   
                data['Times'] = data['Tr']
                data['Times'] = data['Times'].replace(0,data['Trc'])
                data['Tech'] = data['Tech_res']
                data['Tech'] = data['Tech'].replace(0,data['Tech_clos'])
                to_drop = ['None']
                data = data[~data['date_de_cloture'].isin(to_drop)]
                if (year==yend) & (month == mend+1):
                    break
                if month != 12:
                    data = data.loc[(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['type']==val)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]
                else:
                    data = data.loc[(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['type']==val)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year+1,1,1))]
                data = data
                list_tech = data['Tech'].unique()
                for tech in list_tech:
                    if month != 12:
                        new_row = {'Tech':tech,'means':data.loc[(data['Tech']==tech)&(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['Process'].str.contains(str(val), regex=True)==True)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]['Times'].mean()/3600000000000,
                               'date':pd.datetime(year,month,1)}
                        dfFinal= dfFinal.append(new_row, ignore_index=True)
                    else:
                        new_row = {'Tech':tech,'means':data.loc[(data['Tech']==tech)&(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['Process'].str.contains(str(val), regex=True)==True)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year+1, 1,1))]['Times'].mean()/3600000000000,
                               'date':pd.datetime(year,month,1)}
                        dfFinal= dfFinal.append(new_row, ignore_index=True)
        dfFinal['tech_label'] = dfFinal.apply (lambda row: label_tech(row), axis=1)
        dfFinal.to_excel(noms[it]+'.xlsx')
        do.to_sql(dfFinal,db=noms[it])
        date2 = datetime.datetime.now()
        print('durée de la phase moyenne résolution pour', noms[it], ' :',date2-date1)
    
    ## MOYENNE Des Résolutions Par technicien, par mois pour les résolutions par Process
    # Boucle pour faire par types les moyennes
    liste = [1,2,3,4]#Liste des types
    noms = ['scores_res_p1_means','scores_res_p2_means','scores_res_p3_means','scores_res_p4_means']
    for it, val in enumerate(liste):
        date1 = datetime.datetime.now()
        dfFinal = pd.DataFrame()            
        for year in range (ystart, yend+1):
            for month in range( mstart, 12+1):
                data = pd.read_excel('data.xlsx')   
                data['Times'] = data['Tr']
                data['Times'] = data['Times'].replace(0,data['Trc'])
                data['Tech'] = data['Tech_res']
                data['Tech'] = data['Tech'].replace(0,data['Tech_clos'])
                to_drop = ['None']
                data = data[~data['date_de_cloture'].isin(to_drop)]
                if (year==yend) & (month == mend+1):
                    break
                if month != 12:
                    data = data.loc[(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['Process'].str.contains(str(val), regex=True)==True)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]
                else:
                    data = data.loc[(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['Process'].str.contains(str(val), regex=True)==True)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year+1,1,1))]
                data = data
                list_tech = data['Tech'].unique()
                for tech in list_tech:
                    if month != 12:
                        new_row = {'Tech':tech,'means':data.loc[(data['Tech']==tech)&(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['Process'].str.contains(str(val), regex=True)==True)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]['Times'].mean()/3600000000000,
                               'date':pd.datetime(year,month,1)}
                        dfFinal= dfFinal.append(new_row, ignore_index=True)
                    else:
                        new_row = {'Tech':tech,'means':data.loc[(data['Tech']==tech)&(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['Process'].str.contains(str(val), regex=True)==True)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year+1, 1,1))]['Times'].mean()/3600000000000,
                               'date':pd.datetime(year,month,1)}
                        dfFinal= dfFinal.append(new_row, ignore_index=True)
        #if (dfFinal.shape[0] != 0):
        dfFinal['tech_label'] = dfFinal.apply (lambda row: label_tech(row), axis=1)
        dfFinal.to_excel(noms[it]+'.xlsx')
        do.to_sql(dfFinal,db=noms[it])
        date2 = datetime.datetime.now()
        print('durée de la phase moyenne pour résolution P', val, ' pour tickets :',date2-date1)
        
         ## MOYENNE Des Traitements Par technicien, par mois pour les Traitement par process
    # Boucle pour faire par types les moyennes
    liste = [1,2,3,4]#Liste des types
    noms = ['scores_trait_p1_means','scores_trait_p2_means','scores_trait_p3_means','scores_trait_p4_means']
    for it, val in enumerate(liste):
        date1 = datetime.datetime.now()
        dfFinal = pd.DataFrame()            
        for year in range (ystart, yend+1):
            for month in range( mstart, 12+1):
                data = pd.read_excel('data.xlsx')   
                #data['Times'] = data['Tr']
                #data['Times'] = data['Times'].replace(0,data['Trc'])
                data['Tech'] = data['Tech_res']
                data['Tech'] = data['Tech'].replace(0,data['Tech_clos'])
                to_drop = ['None']
                data = data[~data['date_de_cloture'].isin(to_drop)]
                if (year==yend) & (month == mend+1):
                    break
                if month != 12:
                    data = data.loc[(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['Process'].str.contains(str(val), regex=True)==True)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]
                else:
                    data = data.loc[(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['Process'].str.contains(str(val), regex=True)==True)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year+1,1,1))]
                data = data
                list_tech = data['Tech'].unique()
                for tech in list_tech:
                    if month != 12:
                        new_row = {'Tech':tech,'means':data.loc[(data['Tech']==tech)&(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['Process'].str.contains(str(val), regex=True)==True)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]['Tt'].mean()/3600000000000,
                               'date':pd.datetime(year,month,1)}
                        dfFinal= dfFinal.append(new_row, ignore_index=True)
                    else:
                        new_row = {'Tech':tech,'means':data.loc[(data['Tech']==tech)&(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                        (data['Process'].str.contains(str(val), regex=True)==True)&
                                        (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year+1,1,1))]['Tt'].mean()/3600000000000,
                                   'date':pd.datetime(year,month,1)}
                        dfFinal= dfFinal.append(new_row, ignore_index=True)
                        
        #if (dfFinal.shape[0] != 0):
        dfFinal['tech_label'] = dfFinal.apply (lambda row: label_tech(row), axis=1)
        dfFinal.to_excel(noms[it]+'.xlsx')
        do.to_sql(dfFinal,db=noms[it])
        date2 = datetime.datetime.now()
        print('durée de la phase moyenne pour traitement P', val, ' pour tickets :',date2-date1)
        
          ## Ratio Des Tickets à Résolutions Par technicien, par mois pour les Traitement par process
    
    liste = [1,2,3,4]#Liste des types
    noms = ['scores_ratio_p1_means','scores_ratio_p2_means','scores_ratio_p3_means','scores_ratio_p4_means']
    for it, val in enumerate(liste):
        date1 = datetime.datetime.now()
        dfFinal = pd.DataFrame()            
        for year in range (ystart, yend+1):
            for month in range( mstart, 12+1):
                data = pd.read_excel('data.xlsx')   
                #data['Times'] = data['Tr']
                #data['Times'] = data['Times'].replace(0,data['Trc'])
                data['Tech'] = data['Tech_res']
                data['Tech'] = data['Tech'].replace(0,data['Tech_clos'])
                to_drop = ['None']
                data = data[~data['date_de_cloture'].isin(to_drop)]
                if (year==yend) & (month == mend+1):
                    #print('le_mois vaut : ',month, mend)
                    break
                if month != 12:
                    data = data.loc[(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['Process'].str.contains(str(val), regex=True)==True)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]
                else:
                    data = data.loc[(pd.to_datetime(data['date_de_cloture'])>=pd.datetime(year,month,1))&
                                    (data['Process'].str.contains(str(val), regex=True)==True)&
                                    (pd.to_datetime(data['date_de_cloture'],format='%Y-%m-%d')<pd.datetime(year+1,1,1))]
                data = data
                list_tech = data['Tech'].unique()
                for tech in list_tech:
                    new_row = {'Tech':tech,
                               'resolu':data.loc[(data['resolu']==True) & (data['Tech']==tech)]['resolu'].count(),
                               'clos':data.loc[(data['resolu']==False) & (data['Tech']==tech)]['resolu'].count(),
                               'ratio_res':data.loc[(data['resolu']==True) & (data['Tech']==tech)]['resolu'].count()/data.loc[(data['Tech']==tech)]['resolu'].count() if (data.loc[(data['Tech']==tech)]['resolu'].count()!=0) else 0,
                               'ratio_clos':data.loc[(data['resolu']==False) & (data['Tech']==tech)]['resolu'].count()/data.loc[(data['Tech']==tech)]['resolu'].count() if (data.loc[(data['Tech']==tech)]['resolu'].count()!=0) else 0,
                               'date':pd.datetime(year,month,1)}
                    dfFinal= dfFinal.append(new_row, ignore_index=True)
        #if (dfFinal.shape[0] != 0):
        dfFinal['tech_label'] = dfFinal.apply (lambda row: label_tech(row), axis=1)
        dfFinal.to_excel(noms[it]+'.xlsx')
        do.to_sql(dfFinal,db=noms[it])
        date2 = datetime.datetime.now()
        print('durée de la phase moyenne pour ratio P', val, ' pour tickets :',date2-date1)
