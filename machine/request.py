# -*- coding: utf-8 -*-
"""
Created on Thu Feb 27 11:56:13 2020

@author: rboyrie
"""

import machine.do as do
import pymysql
#CONNECTION À LA BASE DE DONNÉES
database = 'helpdesk_data'
def data_bymonth(month="1", year='2019', db='helpdesk_kpi_tables_data'):
    """
    Utilise la table de data extraite sur la base de production pour renvoyer
    une dataframe de toutes les valeurs close durant un certain mois d’une
    année.
    """

    result = do.r_to_b(query="SELECT * FROM " + db + " WHERE MONTH(`date_de_cloture`)=" + str(month) +\
                    " AND YEAR(`date_de_cloture`) = " + str(year))
    try :
        return result
    except :
        print('Erreur dans le retour de première pec pour le ticket X')
        return None
    
def data_bymonth_byprocess(month="1", year='2019', process=0, db='helpdesk_kpi_tables_data'):
    """
    Utilise la table de data extraite sur la base de production pour renvoyer
    une dataframe de toutes les valeurs close durant un certain mois d’une
    année. Et si le process n’est pas zéro, trie selon le process donné en 
    paramètre.
    """


    if process != 0:
        result = do.r_to_b(query=f"SELECT * FROM {db} WHERE MONTH(`date_de_cloture`)={month} AND YEAR(`date_de_cloture`) = {year} AND `Process` REGEXP \'.*{process}.*\'")
    if process == 0:
        
        result = do.r_to_b(query="SELECT * FROM " + db + " WHERE MONTH(`date_de_cloture`)=" + str(month) +\
                    " AND YEAR(`date_de_cloture`) = " + str(year))

    try :
        return result
    except :
        print('Erreur dans le retour de première pec pour le ticket X')
        return None

def id_bymonth(month="1", year='2019', db='helpdesk_kpi_tables_data'):
    """
    renvoie la liste des id clos dans un mois d’une année données.
    """

    result = do.r_to_b(query="SELECT `demande` FROM " + db + " WHERE MONTH(`date_de_cloture`)=" + str(month) +\
                    " AND YEAR(`date_de_cloture`) = " + str(year))
    
    try :
        liste = result.values
        lis = [i for i in liste.flatten()]
        return lis
    except :
        print('Erreur dans le retour de première pec pour le ticket X')
        return None

def opened_ticket_number_in_a_month(month="1",process="1", nature="i", year="2019", table='helpdesk_kpi_tables_data'):
    """
    renvoie le nombre de tickets ouverts dans un mois d’une année donnée.
    """
    result = do.r_to_b(query=f"SELECT COUNT(*) FROM {table} WHERE MONTH(`creation`)={month} AND YEAR(`creation`) = {year} AND `Process` REGEXP \'.*{process}.*\' AND `Process` REGEXP \'.*{nature}.*\'")

    return int(result.values.flatten()[0])
