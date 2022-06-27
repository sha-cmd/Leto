# -*- coding: utf-8 -*-
"""
Created on Tue Feb 11 10:45:39 2020
ATTENTION À CHANGER LES FONCTIONS l_o_d et r_to_b quand la base de production
est actualisé dans un nouveau schema.
@author: rboyrie
"""

import mysql.connector
from sqlalchemy import create_engine
import pandas as pd
import numpy as np
from item import ticket
import pymysql

dbase='helpdesk_data'

def to_sql(df,debug=0, db='helpdesk_data', mode="replace", local=1):
    """
    Écrit dans une table de la base la Data Frame passé en argument.
    Si le mode est append, il écrit à la suite.
    """
    # Conserver ! Historique de la connexion par mysql connector à la base 
    # locale.
    #cnx = mysql.connector.connect(user='romain', password='oXzhZestDhg3',
    #host='127.0.0.1',
    #database='helpdesk_data')
    # Utilisation du framework sqlalchemy
    
    engine = create_engine('mysql+pymysql://romain:oXzhZestDhg3@localhost/'+\
                        dbase +'?charset=utf8', encoding="utf8")
    conn = engine.connect()
    conn.detach()
    df.to_sql(db, conn, if_exists = mode)
    
    if local==0:
        conweb = create_engine(\
            'mysql+pymysql://romain:LwRL345T6OPi*2s!WgSEK@192.168.2.141/'+\
            'helpdesk_data' +'?charset=utf8', encoding="utf8")
        df.to_sql(db, conweb, if_exists = mode)
        conweb.close()
    # TO DO : fermeture de la connexion à la base de données.
    #cursor = connection.cursor()
    conn.close()

def r_to_b(db='helpdesk_data', query="SELECT * FROM `prise_en_charge`",\
           debug=0, local=1):
    """
    Exécute une requête à la base de données. La requête précise la table 
    a utilisée. Si local vaut 0, la requête s’effectuera sur le serveur de 
    DSI, Sinon sur le serveur local.
    """
    # Connection par serveur local
    
    #cnx = mysql.connector.connect(user='romain', password='oXzhZestDhg3',
    #host='127.0.0.1', #J’avais écrit db à la place de la string de database
    #database=dbase)
    #cnx = engine.connect()
    #cnx.detach()
    
    
    engine = create_engine('mysql+pymysql://romain:oXzhZestDhg3@localhost/'+\
                        dbase +'?charset=utf8', encoding="utf8")
    cnx = engine.connect()
    cnx.detach()
    
    # Transformation pour lecture depuis le serveur ETL
    if local==0:
        cnx = mysql.connector.connect(user='romain',\
                                      password='LwRL345T6OPi*2s!WgSEK',
        host='192.168.2.141', database=dbase)
        
    result = pd.read_sql(query, cnx)
    cnx.close()
#datapc.to_excel('pec.xlsx')
    data = pd.DataFrame(result)
    if debug == 1:
        print(data)
    return(data)

##TO DO THAT FUNCTION
#def g_type_t(debug=0, ide):
 #   r_to_b(query="SELECT `id_type_demande` FROM `demande` WHERE `id_demande` = " + str(ide))

def g_row_nb(debug=0, year=0, month=0, day=0):
    """
    Renvoi le numero du dernier ticket du mois d'une année, ou directement de 
    toute la base.
    """
    if (year!=0)&(month!=0)&(day==0):
        result = r_to_b(query=\
            "SELECT * FROM `demande` WHERE YEAR(`date_creation_demande`) = "\
            + year + " AND MONTH(`date_creation_demande`) = " + month \
            + " ORDER BY `id_demande` DESC LIMIT 1 ")
        if debug == 1:
            print('Le dernier ticket est', result['id_demande'].iloc[0])
        return result['id_demande'].iloc[0]
    elif (year==0)&(month==0)&(day==0):
        result = r_to_b(query="SELECT * FROM `demande` ORDER BY `id_demande` DESC LIMIT 1 ")
        if debug == 1:
            print('Le dernier ticket est', result['id_demande'].iloc[0])
        return result['id_demande'].iloc[0]
    else:
        print('Ne fonctionne pas encore selon les jours')
    

def d_of_t(datedeb, datefin):
    """
    Version simplifié de la fonction delta dans le module machine. Servait à
    calculer le temps écoulé entre deux dates en journée travaillée.
    """
    deb = pd.to_datetime(datedeb, format='%y%m%d %h:%m:%s')
    fin = pd.to_datetime(datefin, format='%y%m%d %h:%m:%s')
    if (type(deb) == type(fin)):
    #Est-ce que c'est plus d'un jour ?
        A = deb.date()
        B = fin.date()
        days = 0
        if A <= B:
            days = np.busday_count(A,B)
        if A > B:
            days = np.busday_count(B,A)
    #Est-ce en dehors des heures de bureau COB
        if int(deb.strftime('%H')) > 17:
            deb = deb.replace(hour=17, minute=0, second=0)
        if int(deb.strftime('%H')) < 8:
            deb = deb.replace(hour=8, minute=0, second=0)
        if int(fin.strftime('%H')) > 17:
            fin = fin.replace(hour=17, minute=0, second=0)
        if int(fin.strftime('%H')) < 8:
            fin = fin.replace(hour=8, minute=0, second=0)
        #print(deb)
    #Pour la même journée
        if days == 0:
            return fin - deb
    #Pour un jour de décalage
        if days == 1:
            temp1 = pd.Timedelta('17 hours') 
            temp2 = temp1 - pd.Timedelta(deb.strftime('%H:%M:%S'))
            temp3 = pd.Timedelta('8 hours')
            temp4 = pd.Timedelta(fin.strftime('%H:%M:%S')) - temp3
            return  temp2 + temp4
    #Pour plusieurs jours de décalage
        if days > 1:
            temp1 = pd.Timedelta('17 hours')
            temp2 = temp1 - pd.Timedelta(deb.strftime('%H:%M:%S'))
            temp3 = pd.Timedelta('8 hours')
            temp4 = pd.Timedelta(fin.strftime('%H:%M:%S')) - temp3
            temp5 = pd.Timedelta( '8 hours') * (days - 1)
            return temp2 + temp4 + temp5
    else:
        return pd.Timedelta('0 hours')

def p_b_s():    
    """
    Renvoie les listes des tickets suivant leur nombre de prise en charge et
    suivant leur nombre de suivi de manière croisée. 'suivi' est la variable
    des suivis alors que 'pec' est la variable des prises en charge.
    """
    for suivi in range(2,3):
        for pec in range(3,4):
            nb = 0
            list_of_tickets = np.array([0])
            list_of_actions = np.full([suivi], 0, dtype=int)
            list_of_statuts = np.full([suivi], 0, dtype=int)
            print( 'boucle{}{}'.format(suivi,pec))
            for i in range (0,int(g_row_nb())):
                tkt = ticket.Ticket(2037)
               ##Échantillon de test pour sélectionner les suivis
                if (tkt.nb_pec == pec ) & (tkt.nb_suivi == suivi)\
                & (tkt.leticket_existe == True) :
                    
                    nb += 1
                    list_of_tickets = np.vstack((list_of_tickets,tkt.demande))
                    list_of_actions = np.vstack((list_of_actions,\
                                                 tkt.liste_action))
                    list_of_statuts = np.vstack((list_of_statuts,\
                                                 tkt.liste_statut))
            dft = pd.DataFrame(list_of_tickets)
            dft.to_excel('dft{}{}.xlsx'.format(suivi,pec))
            dfa = pd.DataFrame(np.unique(list_of_actions,axis=0))
            dfa.to_excel('dfa{}{}.xlsx'.format(suivi,pec))
            dfs = pd.DataFrame(np.unique(list_of_statuts,axis=0))
            dfs.to_excel('dfs{}{}.xlsx'.format(suivi,pec))

def l_o_d(month, year, db=dbase):
    """
    Renvoie la liste des id des demandes d'un mois en particulier. Concerne 
    donc à la fois les demandes de ressources, et les demandes de support
    (les interventions)
    """
    result = r_to_b(query=\
                   'SELECT `id_demande` FROM `suivi_intervention` WHERE '\
                   'YEAR(`date_suivi`) =' +\
                   str(year) +' AND MONTH(`date_suivi`)='\
                   + str(month) +' AND `id_statut_demande`= 6  ',db=db)
    liste = []
    for ide in result['id_demande']:
        liste.append(ide)            
    #print(len(liste))
    result = r_to_b(query=\
                'SELECT `id_demande` FROM `suivi_demande_ressource` WHERE '\
                 'YEAR(`date_suivi`) =' +\
                 str(year) +' AND MONTH(`date_suivi`)='\
                   + str(month) +' AND `id_statut_demande`= 6  ',db=db)
    for ide in result['id_demande']:
        liste.append(ide)
    return liste


def instantiate_a_ticket(ide, debug=0):
    """
    La fonction instancie un ticket et le retourne s'il existe en base.
    """
    item = ticket.Ticket(ide)
    if debug == 1:
        print('=========================')
        print('Ticket {}'.format(item.demande))
        item.__str__()
    if item.leticket_existe:
        return item
    else:
        if debug == 1:
            print('Le ticket n\'existe pas')
        return None
    
def m_p_b_s(debug=0):
    """
    Écrit la matrice des suivis selon les Pec
    """
    liste = np.zeros((5,4))
    for pec in range(1,5):
        for suivi in range(0,5):
            nb = 0
            #print('boucle {}{}'.format(k,l))
            for ide in range ( 0, g_row_nb() ):
                tkt = instantiate_a_ticket(ide)
                try:
                    if (tkt.nb_pec == pec )&(tkt.nb_suivi == suivi)\
                    & (tkt.leticket_existe == True) :
                        nb += 1
                except:
                    pass
            liste[suivi,pec-1]=nb            
           # print(liste)            
    #print(liste)
    df = pd.DataFrame(liste, columns=[1,2,3,4])
    df.to_excel('matrice_des_suivis.xlsx')
