# -*- coding: utf-8 -*-
"""
Created on Tue Feb 18 10:58:29 2020

@author: rboyrie
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import timedelta
from datetime import datetime
from machine import do
from scipy.stats import genextreme
import scipy.stats as st
import math
import sys
#import matplotlib
#matplotlib.use('Agg')

def delta(datedeb, datefin):
    """
    Calcul le temps entre deux en ne comptant que les jours ouvrés et tenant 
    compte des jours fériés français.
    """
    #Définition des jours fériés par le fichier téléchargé sur le site du gouvernement
    #et dont le code de production est dans le fichier 
    holidayfile ='jours-feries-seuls.csv'   
    df_holiday = pd.read_csv(holidayfile)
    df_holiday['date'] = pd.to_datetime(df_holiday['date'], format='%Y-%m-%d')
    holidays = np.array(df_holiday['date'], dtype='datetime64[D]')

    deb = pd.to_datetime(datedeb, format='%y%m%d %h:%m:%s')
    fin = pd.to_datetime(datefin, format='%y%m%d %h:%m:%s')
    deb_copy = deb
    fin_copy = fin
    deb_out = False
    fin_out = False
    #deb = np.busday_offset(deb.date(), 0, roll='forward',holidays=holidays)
    #Est-ce un jour chômé:
    #
    try:
        if not (np.is_busday(deb.date(),holidays=holidays)):
            deb = pd.to_datetime(np.busday_offset(deb.date(), 0, roll='forward',holidays=holidays))
        #print(fin.date())
        if not (np.is_busday(fin.date(),holidays=holidays)):
            fin = pd.to_datetime(np.busday_offset(fin.date(), 0, roll='forward',holidays=holidays))
    #print(deb.year)
     #   print(deb)
        if (type(deb) == type(fin)):
        #Est-ce que c'est plus d'un jour ?
            A = deb.date()
            B = fin.date()
            days = 0
            days = np.busday_count(A,B,holidays = holidays)
            if deb > fin:
                # La fonction contient deux nombre de 
                return pd.Timedelta('0 hours')
                #days = np.busday_count(B,A,holidays = holidays)
        #Est-ce en dehors des heures de bureau COB
            if int(deb.strftime('%H')) >= 17:
                deb = deb.replace(hour=17, minute=0, second=0)
                deb_out = True
            if int(deb.strftime('%H')) <= 8:
                deb = deb.replace(hour=8, minute=0, second=0)
                deb_out = True
            if int(fin.strftime('%H')) >= 17:
                fin = fin.replace(hour=17, minute=0, second=0)
                fin_out = True
            if int(fin.strftime('%H')) <= 8:
                fin = fin.replace(hour=8, minute=0, second=0)
                fin_out = True
    
        #Pour la même journée
            if days == 0:
                if (int(deb.strftime('%H')) < 13) and (int(fin.strftime('%H')) > 13) and (pd.Timedelta(fin-deb).seconds//3600 > 1):
                    return fin - deb - pd.Timedelta('1 hours')
                if (deb_out == True) or (fin_out == True):
                    return fin_copy - deb_copy
                return fin - deb
        #Pour un jour de décalage
            if days == 1:
                temp1 = pd.Timedelta('17 hours') 
                temp2 = temp1 - pd.Timedelta(deb.strftime('%H:%M:%S'))
                temp3 = pd.Timedelta('8 hours')
                temp4 = pd.Timedelta(fin.strftime('%H:%M:%S')) - temp3
                if (int(deb.strftime('%H')) < 13) and (int(fin.strftime('%H')) >= 13):
                    return temp2 + temp4 - pd.Timedelta('2 hours')
                elif (int(deb.strftime('%H')) < 13) and (int(fin.strftime('%H')) < 13):
                    return temp2 + temp4 - pd.Timedelta('1 hours')
                elif (int(deb.strftime('%H')) >= 13) and (int(fin.strftime('%H')) < 13):
                    return temp2 + temp4
                elif (int(deb.strftime('%H')) >= 13) and (int(fin.strftime('%H')) >= 13):
                    return temp2 + temp4 - pd.Timedelta('1 hours')
                else:
                    return  temp2 + temp4
        #Pour plusieurs jours de décalage
            if days > 1:
                temp1 = pd.Timedelta('17 hours')
                temp2 = temp1 - pd.Timedelta(deb.strftime('%H:%M:%S'))
                temp3 = pd.Timedelta('8 hours')
                temp4 = pd.Timedelta(fin.strftime('%H:%M:%S')) - temp3
                temp5 = pd.Timedelta( '8 hours') * (days - 1)
                if (int(deb.strftime('%H')) < 13) and (int(fin.strftime('%H')) >= 13) and (pd.Timedelta(fin-deb).seconds//3600 > 1):
                    return temp2 + temp4 + temp5  - pd.Timedelta('2 hours')
                elif (int(deb.strftime('%H')) < 13) and (int(fin.strftime('%H')) < 13):
                    return temp2 + temp4 + temp5  - pd.Timedelta('1 hours')
                elif (int(deb.strftime('%H')) >= 13) and (int(fin.strftime('%H')) < 13):
                    return temp2 + temp4 + temp5 
                elif (int(deb.strftime('%H')) >= 13) and (int(fin.strftime('%H')) >= 13):
                    return temp2 + temp4 + temp5 - pd.Timedelta('1 hours')
                else:
                    return  temp2 + temp4 + temp5
        else:
            print('Fonction delta : Les dates ne sont pas de même type, retourne 0')
            return pd.Timedelta('0 hours')
    except:
        print('Fonction delta : Une date est nulle fonction delta, retourne 0')
        return pd.Timedelta('0 hours')

def is_busday(date):
    """
    Exprime si une date correspond à un jour de travail
    """
    #Définition des jours fériés par le fichier téléchargé sur le site du gouvernement
    #et dont le code de production est dans le fichier 
    holidayfile ='jours-feries-seuls.csv'   
    df_holiday = pd.read_csv(holidayfile)
    df_holiday['date'] = pd.to_datetime(df_holiday['date'], format='%Y-%m-%d')
    holidays = np.array(df_holiday['date'], dtype='datetime64[D]')

    date_to_check = pd.to_datetime(date, format='%y%m%d')
    if (np.is_busday(date_to_check.date(),holidays=holidays)):
        return True
    else:
        return False

def fit_to_all_distributions(data):
    """
    Calcul la loi statistique d’un échantillon en parallèle avec la fonction 
    get_best_distribution_using_chisquared_test().
    """
    dist_names = ['fatiguelife', 'invgauss', 'johnsonsu', 'johnsonsb', 'lognorm', 'norminvgauss', 'powerlognorm', 'exponweib','genextreme', 'pareto']

    params = {}
    for dist_name in dist_names:
        try:
            dist = getattr(st, dist_name)
            param = dist.fit(data)

            params[dist_name] = param
        except Exception:
            print("Error occurred in fitting")
            params[dist_name] = "Error"

    return params 


def get_best_distribution_using_chisquared_test(data, params):
    """
    Calcul la meilleure loi statistique d’un échantillon pour fonctionner, utilise la fonction
    fit_to_all_distribution()
    """
    histo, bin_edges = np.histogram(data, bins='auto', normed=False)
    observed_values = histo

    dist_names = ['fatiguelife', 'invgauss', 'johnsonsu', 'johnsonsb', 'lognorm', 'norminvgauss', 'powerlognorm', 'exponweib','genextreme', 'pareto']

    dist_results = []

    for dist_name in dist_names:

        param = params[dist_name]
        if (param != "Error"):
            # Applying the SSE test
            arg = param[:-2]
            loc = param[-2]
            scale = param[-1]
            cdf = getattr(st, dist_name).cdf(bin_edges, loc=loc, scale=scale, *arg)
            expected_values = len(data) * np.diff(cdf)
            c , p = st.chisquare(observed_values, expected_values, ddof=len(param))
            dist_results.append([dist_name, c, p, expected_values])


    # select the best fitted distribution
    best_dist, best_c, best_p, best_expected_values, best_observed_values = None, sys.maxsize, 0, [],[]

    for item in dist_results:
        name = item[0]
        c = item[1]
        p = item[2]
        if (not math.isnan(c)):
            if (c < best_c):
                best_c = c
                best_dist = name
                best_p = p
                best_expected_values = expected_values
                best_observed_values = observed_values

    # print the name of the best fit and its p value

    #print("Best fitting distribution: " + str(best_dist))
    #print("Best c value: " + str(best_c))
    #print("Best p value: " + str(best_p))
    #print("Parameters for the best fit: " + str(params[best_dist]))

    return best_dist, best_c, best_p, best_expected_values, best_observed_values, params[best_dist], dist_results

def scan_cat(start_month, month_end, year_start=2019, year_end=2020, file_name='quantite_2019.xlsx', mode="replace"):
    """
    Scan la base de données et en sort les quantités de tickets par mois selon leur
    processus. Écrit le résultat dans un fichier
    """
    df = pd.DataFrame({
        'nbdc1' : [],# Nombre de demande close directement de Processus 1
        'nbdc2' : [],
        'nbdc3' : [],
        'nbdc4' : [],
        'nbdrc1' : [],# Nombre de demande close avec résolution de Process 1
        'nbdrc2' : [],
        'nbdrc3' : [],
        'nbdrc4' : [],
        'nbic1' : [],# Nombre d’intervention close directement de Processus 1
        'nbic2' : [],
        'nbic3' : [],
        'nbic4' : [],
        'nbirc1' : [],
        'nbirc2' : [],
        'nbirc3' : [],
        'nbirc4' : [],
        'nbbec' : [],# Décompte des éléments n’entrant dans aucun Processus
        'mois' : [],
        'an' : []
                })
    for year in range (year_start, year_end+1):
        if (year < year_end):
            end_month = 12
        else:
            end_month = month_end
        for i in range(start_month, end_month+1):
            nbdc1 = 0
            nbdc2 = 0
            nbdc3 = 0
            nbdc4 = 0
            nbdrc1 = 0
            nbdrc2 = 0
            nbdrc3 = 0
            nbdrc4 = 0
            nbic1 = 0
            nbic2 = 0
            nbic3 = 0
            nbic4 = 0
            nbirc1 = 0
            nbirc2 = 0
            nbirc3 = 0
            nbirc4 = 0
            nbbec = 0
            date = 0
            liste_demandes = do.l_o_d(month = i, year = year) 
            for j in liste_demandes:
                b = do.instantiate_a_ticket(j)
                    
                if (b is not None) and (b.est_clos_en_suivi): 
                    if (b.date_de_cloture.year == year) and (b.date_de_cloture.month == i):
                        date = b.date_de_cloture
                        
                        if (b.str_process == 'Process d c 1'):
                            nbdc1+=1
                        elif (b.str_process == 'Process d c 2'):
                            nbdc2+=1
                        elif (b.str_process == 'Process d c 3'):
                            nbdc3+=1
                        elif (b.str_process == 'Process d c 4'):
                            nbdc4+=1
                        elif (b.str_process == 'Process d rc 1'):
                            nbdrc1+=1
                        elif (b.str_process == 'Process d rc 2'):
                            nbdrc2+=1
                        elif (b.str_process == 'Process d rc 3'):
                            nbdrc3+=1
                        elif (b.str_process == 'Process d rc 4'):
                            nbdrc4+=1
                        elif (b.str_process == 'Process i c 1'):
                            nbic1+=1
                        elif (b.str_process == 'Process i c 2'):
                            nbic2+=1
                        elif (b.str_process == 'Process i c 3'):
                            nbic3+=1
                        elif (b.str_process == 'Process i c 4'):
                            nbic4+=1
                        elif (b.str_process == 'Process i rc 1'):
                            nbirc1+=1
                        elif (b.str_process == 'Process i rc 2'):
                            nbirc2+=1
                        elif (b.str_process == 'Process i rc 3'):
                            nbirc3+=1
                        elif (b.str_process == 'Process i rc 4'):
                            nbirc4+=1
                        else:
                            nbbec +=1
            df = df.append({\
            'nbdc1' : nbdc1,
            'nbdc2' : nbdc2,
            'nbdc3' : nbdc3,
            'nbdc4' : nbdc4,
            'nbdrc1' : nbdrc1,
            'nbdrc2' : nbdrc2,
            'nbdrc3' : nbdrc3,
            'nbdrc4' : nbdrc4,
            'nbic1' : nbic1,
            'nbic2' : nbic2,
            'nbic3' : nbic3,
            'nbic4' : nbic4,
            'nbirc1' : nbirc1,
            'nbirc2' : nbirc2,
            'nbirc3' : nbirc3,
            'nbirc4' : nbirc4,
            'nbbec' : nbbec,
            'mois' :  pd.to_datetime('{}/{}'.format(date.month,date.year),format\
                                     ='%m/%Y').date(),
            'an' : date.year
                                    }, ignore_index=True)
    df.index = df['mois']
    df = df.drop('mois',axis=1)
    df.to_excel(file_name)
    do.to_sql(df, db='helpdesk_kpi_tables_quantite', mode=mode)
    print('Scan_cat s’est bien déroulé')
    
def scan_cat_Tt(start_month, month_end, year_start=2019, year_end=2020, file_name='Tt_2019.xlsx', mode="replace"):
    """
    Scan la base et en sort les temps de traitement de tickets par mois selon leur
    processus. Écrit le résultat dans un fichier
    """
    df = pd.DataFrame({
        'Ttnbdc1' : [],
        'Ttnbdc2' : [],
        'Ttnbdc3' : [],
        'Ttnbdc4' : [],
        'Ttnbdrc1' : [],
        'Ttnbdrc2' : [],
        'Ttnbdrc3' : [],
        'Ttnbdrc4' : [],
        'Ttnbic1' : [],
        'Ttnbic2' : [],
        'Ttnbic3' : [],
        'Ttnbic4' : [],
        'Ttnbirc1' : [],
        'Ttnbirc2' : [],
        'Ttnbirc3' : [],
        'Ttnbirc4' : [],
        'Ttnbbec' : [],
        'mois' : [],
        'an' : []
                })
    for year in range (year_start, year_end+1):
        if (year < year_end):
            end_month = 12
        else:
            end_month = month_end
        for i in range(start_month, end_month+1):
            Ttnbdc1 = pd.Timedelta('0 hours')
            Ttnbdc2 = pd.Timedelta('0 hours')
            Ttnbdc3 = pd.Timedelta('0 hours')
            Ttnbdc4 = pd.Timedelta('0 hours')
            Ttnbdrc1 = pd.Timedelta('0 hours')
            Ttnbdrc2 = pd.Timedelta('0 hours')
            Ttnbdrc3 = pd.Timedelta('0 hours')
            Ttnbdrc4 = pd.Timedelta('0 hours')
            Ttnbic1 = pd.Timedelta('0 hours')
            Ttnbic2 = pd.Timedelta('0 hours')
            Ttnbic3 = pd.Timedelta('0 hours')
            Ttnbic4 = pd.Timedelta('0 hours')
            Ttnbirc1 = pd.Timedelta('0 hours')
            Ttnbirc2 = pd.Timedelta('0 hours')
            Ttnbirc3 = pd.Timedelta('0 hours')
            Ttnbirc4 = pd.Timedelta('0 hours')
            Ttnbbec = pd.Timedelta('0 hours')
            nbdc1 = 0
            nbdc2 = 0
            nbdc3 = 0
            nbdc4 = 0
            nbdrc1 = 0
            nbdrc2 = 0
            nbdrc3 = 0
            nbdrc4 = 0
            nbic1 = 0
            nbic2 = 0
            nbic3 = 0
            nbic4 = 0
            nbirc1 = 0
            nbirc2 = 0
            nbirc3 = 0
            nbirc4 = 0
            nbbec = 0
            date = pd.Timedelta('0 hours')
        
            liste_demandes = do.l_o_d(month = i, year = year) 
            for j in liste_demandes:
                b = do.instantiate_a_ticket(j)
                    
                if (b is not None) and (b.est_clos_en_suivi): 
                    if (b.date_de_cloture.year == year) and (b.date_de_cloture.month == i):
                        date = b.date_de_cloture
                        
                        if (b.str_process == 'Process d c 1'):
                            Ttnbdc1+=b.Tt
                            nbdc1+=1
                        elif (b.str_process == 'Process d c 2'):
                            Ttnbdc2+=b.Tt
                            nbdc2+=1
                        elif (b.str_process == 'Process d c 3'):
                            Ttnbdc3+=b.Tt
                            nbdc3+=1
                        elif (b.str_process == 'Process d c 4'):
                            Ttnbdc4+=b.Tt
                            nbdc4+=1
                        elif (b.str_process == 'Process d rc 1'):
                            Ttnbdrc1+=b.Tt
                            nbdrc1+=1
                        elif (b.str_process == 'Process d rc 2'):
                            Ttnbdrc2+=b.Tt
                            nbdrc2+=1
                        elif (b.str_process == 'Process d rc 3'):
                            Ttnbdrc3+=b.Tt
                            nbdrc3+=1
                        elif (b.str_process == 'Process d rc 4'):
                            Ttnbdrc4+=b.Tt
                            nbdrc4+=1
                        elif (b.str_process == 'Process i c 1'):
                            Ttnbic1+=b.Tt
                            nbic1+=1
                        elif (b.str_process == 'Process i c 2'):
                            Ttnbic2+=b.Tt
                            nbic2+=1
                        elif (b.str_process == 'Process i c 3'):
                            Ttnbic3+=b.Tt
                            nbic3+=1
                        elif (b.str_process == 'Process i c 4'):
                            Ttnbic4+=b.Tt
                            nbic4+=1
                        elif (b.str_process == 'Process i rc 1'):
                            Ttnbirc1+=b.Tt
                            nbirc1+=1
                        elif (b.str_process == 'Process i rc 2'):
                            Ttnbirc2+=b.Tt
                            nbirc2+=1
                        elif (b.str_process == 'Process i rc 3'):
                            Ttnbirc3+=b.Tt
                            nbirc3+=1
                        elif (b.str_process == 'Process i rc 4'):
                            Ttnbirc4+=b.Tt
                            nbirc4+=1
                        else:
                            Ttnbbec +=b.Tt
            df = df.append({\
            'Ttnbdc1' : (Ttnbdc1.value/nbdc1)/3600000000000 if (nbdc1!=0) else 0,
            'Ttnbdc2' : (Ttnbdc2.value/nbdc2)/3600000000000 if (nbdc2!=0) else 0,
            'Ttnbdc3' : (Ttnbdc3.value/nbdc3)/3600000000000 if (nbdc3!=0) else 0,
            'Ttnbdc4' : (Ttnbdc4.value/nbdc4)/3600000000000 if (nbdc4!=0) else 0,
            'Ttnbdrc1' : (Ttnbdrc1.value/nbdrc1)/3600000000000 if (nbdrc1!=0) else 0,
            'Ttnbdrc2' : (Ttnbdrc2.value/nbdrc2)/3600000000000 if (nbdrc2!=0) else 0,
            'Ttnbdrc3' : (Ttnbdrc3.value/nbdrc3)/3600000000000 if (nbdrc3!=0) else 0,
            'Ttnbdrc4' : (Ttnbdrc4.value/nbdrc4)/3600000000000 if (nbdrc4!=0) else 0,
            'Ttnbic1' : (Ttnbic1.value/nbic1)/3600000000000 if (nbic1!=0) else 0,
            'Ttnbic2' : (Ttnbic2.value/nbic2)/3600000000000 if (nbic2!=0) else 0,
            'Ttnbic3' : (Ttnbic3.value/nbic3)/3600000000000 if (nbic3!=0) else 0,
            'Ttnbic4' : (Ttnbic4.value/nbic4)/3600000000000 if (nbic4!=0) else 0,
            'Ttnbirc1' : (Ttnbirc1.value/nbirc1)/3600000000000 if (nbirc1!=0) else 0,
            'Ttnbirc2' : (Ttnbirc2.value/nbirc2)/3600000000000 if (nbirc2!=0) else 0,
            'Ttnbirc3' : (Ttnbirc3.value/nbirc3)/3600000000000 if (nbirc3!=0) else 0,
            'Ttnbirc4' : (Ttnbirc4.value/nbirc4)/3600000000000 if (nbirc4!=0) else 0,
            'Ttnbbec' : (Ttnbbec.value/nbbec)/3600000000000 if (nbbec!=0) else 0,
            'mois' :  pd.to_datetime('{}/{}'.format(date.month,date.year),format\
                                     ='%m/%Y').date(),
            'an' : date.year
                                    }, ignore_index=True)
        # Typage des colonnes de la dataframe
    df['Ttnbdc1'] = df['Ttnbdc1'].astype(float)
    df['Ttnbdc2'] = df['Ttnbdc2'].astype(float)
    df['Ttnbdc3'] = df['Ttnbdc3'].astype(float)
    df['Ttnbdc4'] = df['Ttnbdc4'].astype(float)
    df['Ttnbdrc1'] = df['Ttnbdrc1'].astype(float)
    df['Ttnbdrc2'] = df['Ttnbdrc2'].astype(float)
    df['Ttnbdrc3'] = df['Ttnbdrc3'].astype(float)
    df['Ttnbdrc4'] = df['Ttnbdrc4'].astype(float)
    df['Ttnbic1'] = df['Ttnbic1'].astype(float)
    df['Ttnbic2'] = df['Ttnbic2'].astype(float)
    df['Ttnbic3'] = df['Ttnbic3'].astype(float)
    df['Ttnbic4'] = df['Ttnbic4'].astype(float)
    df['Ttnbirc1'] = df['Ttnbirc1'].astype(float)
    df['Ttnbirc2'] = df['Ttnbirc2'].astype(float)
    df['Ttnbirc3'] = df['Ttnbirc3'].astype(float)
    df['Ttnbirc4'] = df['Ttnbirc4'].astype(float)
    df['Ttnbbec'] = df['Ttnbbec'].astype(float)
    df.index = df['mois']
    df = df.drop('mois',axis=1)
    df.to_excel(file_name)
    do.to_sql(df, db='helpdesk_kpi_tables_tt', mode=mode)
    print('Tt s’est bien déroulé')

def scan_cat_t1(start_month, month_end, year_start=2019,year_end=2020, file_name='helpdesk_kpi_tables_t1.xlsx', mode="replace"):
    """
    Scan la base et en sort l'indicateur moyen t1 pour chaque type de tickets 
    par mois selon leur processus. Écrit le résultat dans un fichier
    """
    df = pd.DataFrame({
            't1nbdc1' : [],
        't1nbdc2' : [],
        't1nbdc3' : [],
        't1nbdc4' : [],
        't1nbdrc1' : [],
        't1nbdrc2' : [],
        't1nbdrc3' : [],
        't1nbdrc4' : [],
        't1nbic1' : [],
        't1nbic2' : [],
        't1nbic3' : [],
        't1nbic4' : [],
        't1nbirc1' : [],
        't1nbirc2' : [],
        't1nbirc3' : [],
        't1nbirc4' : [],
        't1nbbec' : [],
        'mois' : [],
        'an' : []
                })
    for year in range (year_start, year_end+1):
        if (year < year_end):
            end_month = 12
        else:
            end_month = month_end
        for i in range(start_month, end_month+1):
            nbdc1 = 0
            nbdc2 = 0
            nbdc3 = 0
            nbdc4 = 0
            nbdrc1 = 0
            nbdrc2 = 0
            nbdrc3 = 0
            nbdrc4 = 0
            nbic1 = 0
            nbic2 = 0
            nbic3 = 0
            nbic4 = 0
            nbirc1 = 0
            nbirc2 = 0
            nbirc3 = 0
            nbirc4 = 0
            nbbec = 0
            t1nbdc1 = pd.Timedelta('0 hours')
            t1nbdc2 = pd.Timedelta('0 hours')
            t1nbdc3 = pd.Timedelta('0 hours')
            t1nbdc4 = pd.Timedelta('0 hours')
            t1nbdrc1 = pd.Timedelta('0 hours')
            t1nbdrc2 = pd.Timedelta('0 hours')
            t1nbdrc3 = pd.Timedelta('0 hours')
            t1nbdrc4 = pd.Timedelta('0 hours')
            t1nbic1 = pd.Timedelta('0 hours')
            t1nbic2 = pd.Timedelta('0 hours')
            t1nbic3 = pd.Timedelta('0 hours')
            t1nbic4 = pd.Timedelta('0 hours')
            t1nbirc1 = pd.Timedelta('0 hours')
            t1nbirc2 = pd.Timedelta('0 hours')
            t1nbirc3 = pd.Timedelta('0 hours')
            t1nbirc4 = pd.Timedelta('0 hours')
            t1nbbec = pd.Timedelta('0 hours')
            date = 0
            liste_demandes = do.l_o_d(month = i, year = year) 
            for j in liste_demandes:
                b = do.instantiate_a_ticket(j)
                    
                if (b is not None) and (b.est_clos_en_suivi):
                    if (b.date_de_cloture.year == year) and (b.date_de_cloture.month == i):
                        date = b.date_de_cloture
                        
                        if (b.str_process == 'Process d c 1'):
                            nbdc1+=1
                            t1nbdc1 += b.t1
                        elif (b.str_process == 'Process d c 2'):
                            nbdc2+=1
                            t1nbdc2+=b.t1
                        elif (b.str_process == 'Process d c 3'):
                            nbdc3+=1
                            t1nbdc3+=b.t1
                        elif (b.str_process == 'Process d c 4'):
                            nbdc4+=1
                            t1nbdc4+=b.t1
                        elif (b.str_process == 'Process d rc 1'):
                            nbdrc1+=1
                            t1nbdrc1+=b.t1
                        elif (b.str_process == 'Process d rc 2'):
                            nbdrc2+=1
                            t1nbdrc2+=b.t1
                        elif (b.str_process == 'Process d rc 3'):
                            nbdrc3+=1
                            t1nbdrc3+=b.t1
                        elif (b.str_process == 'Process d rc 4'):
                            nbdrc4+=1
                            t1nbdrc4+=b.t1
                        elif (b.str_process == 'Process i c 1'):
                            nbic1+=1
                            t1nbic1+=b.t1
                        elif (b.str_process == 'Process i c 2'):
                            nbic2+=1
                            t1nbic2+=b.t1
                        elif (b.str_process == 'Process i c 3'):
                            nbic3+=1
                            t1nbic3+=b.t1
                        elif (b.str_process == 'Process i c 4'):
                            nbic4+=1
                            t1nbic4+=b.t1
                        elif (b.str_process == 'Process i rc 1'):
                            nbirc1+=1
                            t1nbirc1+=b.t1
                        elif (b.str_process == 'Process i rc 2'):
                            nbirc2+=1
                            t1nbirc2+=b.t1
                        elif (b.str_process == 'Process i rc 3'):
                            nbirc3+=1
                            t1nbirc3+=b.t1
                        elif (b.str_process == 'Process i rc 4'):
                            nbirc4+=1
                            t1nbirc4+=b.t1
                        else:
                            nbbec +=1
                            t1nbbec+=b.t1
            df = df.append({
            't1nbdc1' : ( (t1nbdc1.value/(1000000000*3600))/nbdc1) if (nbdc1!=0)else 0,
            't1nbdc2' : ((t1nbdc2.value/(1000000000*3600))/nbdc2) if (nbdc2!=0)else 0,
            't1nbdc3' : ((t1nbdc3.value/(1000000000*3600))/nbdc3) if (nbdc3!=0)else 0,
            't1nbdc4' : ((t1nbdc4.value/(1000000000*3600))/nbdc4) if (nbdc4!=0)else 0,
            't1nbdrc1' :((t1nbdrc1.value/(1000000000*3600))/nbdrc1) if (nbdrc1!=0)else 0,
            't1nbdrc2' :((t1nbdrc2.value/(1000000000*3600))/nbdrc2) if (nbdrc2!=0)else 0,
            't1nbdrc3' :((t1nbdrc3.value/(1000000000*3600))/nbdrc3) if (nbdrc3!=0)else 0,
            't1nbdrc4' :((t1nbdrc4.value/(1000000000*3600))/nbdrc4) if (nbdrc4!=0)else 0,
            't1nbic1' : ((t1nbic1.value/(1000000000*3600))/nbic1) if (nbic1!=0)else 0,
            't1nbic2' : ((t1nbic2.value/(1000000000*3600))/nbic2) if (nbic2!=0)else 0,
            't1nbic3' : ((t1nbic3.value/(1000000000*3600))/nbic3) if (nbic3!=0)else 0,
            't1nbic4' : ((t1nbic4.value/(1000000000*3600))/nbic4) if (nbic4!=0)else 0,
            't1nbirc1' :((t1nbirc1.value/(1000000000*3600))/nbirc1 )if (nbirc1!=0) else 0,
            't1nbirc2' :((t1nbirc2.value/(1000000000*3600))/nbirc2 )if (nbirc2!=0) else 0,
            't1nbirc3' :((t1nbirc3.value/(1000000000*3600))/nbirc3 )if (nbirc3!=0) else 0,
            't1nbirc4' :((t1nbirc4.value/(1000000000*3600))/nbirc4) if (nbirc4!=0) else 0,
            't1nbbec' : ((t1nbbec.value/(1000000000*3600))/nbbec) if (nbbec!=0) else 0,
            'mois' :  pd.to_datetime('{}/{}'.format(date.month,date.year),format\
                                     ='%m/%Y').date(),
            'an' : date.year
                                    }, ignore_index=True)
    df.index = df['mois']
    df = df.drop('mois', axis=1)
    df.to_excel(file_name)
    do.to_sql(df,db="helpdesk_kpi_tables_t1",mode=mode)
    print('T1 s’est bien déroulé')
    
def scan_cat_t2(start_month, month_end, year_start=2019, year_end=2020, file_name='helpdesk_kpi_tables_t2.xlsx', mode="replace"):
    """
    Scan la base et en sort l'indicateur moyen t1 pour chaque type de tickets 
    par mois selon leur processus. Écrit le résultat dans un fichier
    """
    df = pd.DataFrame({
        't2nbdc2' : [],
        't2nbdrc2' : [],
        't2nbic2' : [],
        't2nbirc2' : [],
        't2nbbec' : [],
        't2nbdc3' : [],
        't2nbdrc3' : [],
        't2nbic3' : [],
        't2nbirc3' : [],
        'mois' : [],
        'an' : []
                })
    for year in range (year_start, year_end+1):
        if (year < year_end):
            end_month = 12
        else:
            end_month = month_end
        for i in range(start_month, end_month+1):
            nbdc2 = 0
            nbdrc2 = 0
            nbic2 = 0
            nbirc2 = 0
            nbbec = 0
            t2nbdc2 = pd.Timedelta('0 hours')
            t2nbdrc2 = pd.Timedelta('0 hours')
            t2nbic2 = pd.Timedelta('0 hours')
            t2nbirc2 = pd.Timedelta('0 hours')
            t2nbbec = pd.Timedelta('0 hours')
            nbdc3 = 0
            nbdrc3 = 0
            nbic3 = 0
            nbirc3 = 0
            nbbec = 0
            t2nbdc3 = pd.Timedelta('0 hours')
            t2nbdrc3 = pd.Timedelta('0 hours')
            t2nbic3 = pd.Timedelta('0 hours')
            t2nbirc3 = pd.Timedelta('0 hours')
            date = 0
            liste_demandes = do.l_o_d(month = i, year = year) 
            for j in liste_demandes:
                b = do.instantiate_a_ticket(j)
                if (b is not None) and (b.est_clos_en_suivi): 
                    if (b.date_de_cloture.year == year) and (b.date_de_cloture.month == i) \
                    and ((b.est_de_Process_ic2()) or (b.est_de_Process_dc2())\
                        or (b.est_de_Process_irc2()) or(b.est_de_Process_drc2())):
                        date = pd.to_datetime(b.date_de_cloture)
                        if (b.str_process == 'Process d c 2'):
                            nbdc2+=1
                            t2nbdc2+=b.t2
                        elif (b.str_process == 'Process d rc 2'):
                            nbdrc2+=1
                            t2nbdrc2+=b.t2
                        elif (b.str_process == 'Process i c 2'):
                            nbic2+=1
                            t2nbic2+=b.t2
                        elif (b.str_process == 'Process i rc 2'):
                            nbirc2+=1
                            t2nbirc2+=b.t2
                        else:
                            nbbec +=1
                            t2nbbec+=b.t2
                    if (b.date_de_cloture.year == year) and (b.date_de_cloture.month == i) \
                    and ((b.est_de_Process_ic3()) or (b.est_de_Process_dc3())\
                        or (b.est_de_Process_irc3()) or(b.est_de_Process_drc3())):
                        date = pd.to_datetime(b.date_de_cloture)
                        if (b.str_process == 'Process d c 3'):
                            nbdc3+=1
                            t2nbdc3+=b.t2
                        elif (b.str_process == 'Process d rc 3'):
                            nbdrc3+=1
                            t2nbdrc3+=b.t2
                        elif (b.str_process == 'Process i c 3'):
                            nbic3+=1
                            t2nbic3+=b.t2
                        elif (b.str_process == 'Process i rc 3'):
                            nbirc3+=1
                            t2nbirc3+=b.t2        
            df = df.append({
            't2nbdc2' : (t2nbdc2.value/(1000000000*3600))/nbdc2 if (nbdc2!=0)else 0,
            't2nbdrc2' : (t2nbdrc2.value/(1000000000*3600))/nbdrc2 if (nbdrc2!=0)else 0,
            't2nbic2' : (t2nbic2.value/(1000000000*3600))/nbic2 if (nbic2!=0)else 0,
            't2nbirc2' : (t2nbirc2.value/(1000000000*3600))/nbirc2 if (nbirc2!=0) else 0,
            't2nbbec' : (t2nbbec.value/(1000000000*3600))/nbbec if (nbbec!=0) else 0,
            't2nbdc3' : (t2nbdc3.value/(1000000000*3600))/nbdc3 if (nbdc3!=0)else 0,
            't2nbdrc3' : (t2nbdrc3.value/(1000000000*3600))/nbdrc3 if (nbdrc3!=0)else 0,
            't2nbic3' : (t2nbic3.value/(1000000000*3600))/nbic3 if (nbic3!=0)else 0,
            't2nbirc3' : (t2nbirc3.value/(1000000000*3600))/nbirc3 if (nbirc3!=0) else 0,
            'mois' :  pd.to_datetime('{}/{}'.format(date.month,date.year),format\
                                     ='%m/%Y').date() if (date !=0) else 0,
            'an' : date.year if (date != 0) else 0
                                    }, ignore_index=True)
    df.index = df['mois']
    df = df.drop('mois',axis=1)
    df.to_excel(file_name)
    do.to_sql(df,db="helpdesk_kpi_tables_t2",mode=mode)
    print('T2 s’est bien déroulé')
    
def scan_cat_t3(start_month, month_end, year_start=2019, year_end=2020, file_name='helpdesk_kpi_tables_t3".xlsx', mode="replace"):
    """
    Scan la base et en sort l'indicateur moyen t1 pour chaque type de tickets 
    par mois selon leur processus. Écrit le résultat dans un fichier
    """
    df = pd.DataFrame({
        't3nbdc3' : [],
        't3nbdrc3' : [],
        't3nbic3' : [],
        't3nbirc3' : [],
        't3nbbec' : [],
        'mois' : [],
        'an' : []
                })
    for year in range (year_start, year_end+1):
        if (year < year_end):
            end_month = 12
        else:
            end_month = month_end
        for i in range(start_month, end_month+1):
            nbdc3 = 0
            nbdrc3 = 0
            nbic3 = 0
            nbirc3 = 0
            nbbec = 0
            t3nbdc3 = pd.Timedelta('0 hours')
            t3nbdrc3 = pd.Timedelta('0 hours')
            t3nbic3 = pd.Timedelta('0 hours')
            t3nbirc3 = pd.Timedelta('0 hours')
            t3nbbec = pd.Timedelta('0 hours')
            date = 0
            liste_demandes = do.l_o_d(month = i, year = year) 
            for j in liste_demandes:
                b = do.instantiate_a_ticket(j)
                if (b is not None) and (b.est_clos_en_suivi): 
                    if (b.date_de_cloture.year == year) and (b.date_de_cloture.month == i) \
                    and ((b.est_de_Process_ic3()) or (b.est_de_Process_dc3())\
                        or (b.est_de_Process_irc3()) or(b.est_de_Process_drc3())):
                        date = pd.to_datetime(b.date_de_cloture)
                        if (b.str_process == 'Process d c 3'):
                            nbdc3+=1
                            t3nbdc3+=b.t3
                        elif (b.str_process == 'Process d rc 3'):
                            nbdrc3+=1
                            t3nbdrc3+=b.t3
                        elif (b.str_process == 'Process i c 3'):
                            nbic3+=1
                            t3nbic3+=b.t3
                        elif (b.str_process == 'Process i rc 3'):
                            nbirc3+=1
                            t3nbirc3+=b.t3
                        else:
                            nbbec +=1
                            t3nbbec+=b.t3
            df = df.append({
            't3nbdc3' : (t3nbdc3.value/(1000000000*3600))/nbdc3 if (nbdc3!=0)else 0,
            't3nbdrc3' : (t3nbdrc3.value/(1000000000*3600))/nbdrc3 if (nbdrc3!=0)else 0,
            't3nbic3' : (t3nbic3.value/(1000000000*3600))/nbic3 if (nbic3!=0)else 0,
            't3nbirc3' : (t3nbirc3.value/(1000000000*3600))/nbirc3 if (nbirc3!=0) else 0,
            't3nbbec' : (t3nbbec.value/(1000000000*3600))/nbbec if (nbbec!=0) else 0,
            'mois' :  pd.to_datetime('{}/{}'.format(date.month,date.year),format\
                                     ='%m/%Y').date() if (date !=0) else 0,
            'an' : date.year if (date != 0) else 0
                                    }, ignore_index=True)
    df.index = df['mois']
    df = df.drop('mois',axis=1)
    df.to_excel(file_name)
    do.to_sql(df,db="helpdesk_kpi_tables_t3",mode=mode)
    print('T3 s’est bien déroulé')
    
def scan_cat_tr(start_month, month_end, year_start=2019, year_end=2020, file_name='helpdesk_kpi_tables_tr".xlsx',mode="replace"):
    """
    Scan la base et en sort l'indicateur moyen t1 pour chaque type de tickets 
    par mois selon leur processus. Écrit le résultat dans un fichier
    """
    df = pd.DataFrame({
        'trnbdrc1' : [],
        'trnbdrc2' : [],
        'trnbdrc3' : [],
        'trnbdrc4' : [],
        'trnbirc1' : [],
        'trnbirc2' : [],
        'trnbirc3' : [],
        'trnbirc4' : [],
        'trnbbec' : [],
        'mois' : [],
        'an' : []
                })

    for year in range (year_start, year_end+1):
        if (year < year_end):
            end_month = 12
        else:
            end_month = month_end
        for i in range(start_month, end_month+1):
            nbdrc1 = 0
            nbdrc2 = 0
            nbdrc3 = 0
            nbdrc4 = 0
            nbirc1 = 0
            nbirc2 = 0
            nbirc3 = 0
            nbirc4 = 0
            nbbec = 0
            trnbdrc1 = pd.Timedelta('0 hours')
            trnbdrc2 = pd.Timedelta('0 hours')
            trnbdrc3 = pd.Timedelta('0 hours')
            trnbdrc4 = pd.Timedelta('0 hours')
            trnbirc1 = pd.Timedelta('0 hours')
            trnbirc2 = pd.Timedelta('0 hours')
            trnbirc3 = pd.Timedelta('0 hours')
            trnbirc4 = pd.Timedelta('0 hours')
            date = 0
            
            liste_demandes = do.l_o_d(month = i, year = year) 
            for j in liste_demandes:
                b = do.instantiate_a_ticket(j)
                    
                if (b is not None) and (b.est_clos_en_suivi): 
                    if (b.date_de_cloture.year == year) and (b.date_de_cloture.month == i):
                        date = b.date_de_cloture
                        
                        if (b.str_process == 'Process d rc 1'):
                            nbdrc1+=1
                            trnbdrc1+=b.tr
                        elif (b.str_process == 'Process d rc 2'):
                            nbdrc2+=1
                            trnbdrc2+=b.tr
                        elif (b.str_process == 'Process d rc 3'):
                            nbdrc3+=1
                            trnbdrc3+=b.tr
                        elif (b.str_process == 'Process d rc 4'):
                            nbdrc4+=1
                            trnbdrc4+=b.tr
                        elif (b.str_process == 'Process i rc 1'):
                            nbirc1+=1
                            trnbirc1+=b.tr
                        elif (b.str_process == 'Process i rc 2'):
                            nbirc2+=1
                            trnbirc2+=b.tr
                        elif (b.str_process == 'Process i rc 3'):
                            nbirc3+=1
                            trnbirc3+=b.tr
                        elif (b.str_process == 'Process i rc 4'):
                            nbirc4+=1
                            trnbirc4+=b.tr
                        else:
                            nbbec +=1
            df = df.append({
            'trnbdrc1' :((trnbdrc1.value/(1000000000*3600))/nbdrc1) if (nbdrc1!=0)else 0,
            'trnbdrc2' :((trnbdrc2.value/(1000000000*3600))/nbdrc2) if (nbdrc2!=0)else 0,
            'trnbdrc3' :((trnbdrc3.value/(1000000000*3600))/nbdrc3) if (nbdrc3!=0)else 0,
            'trnbdrc4' :((trnbdrc4.value/(1000000000*3600))/nbdrc4) if (nbdrc4!=0)else 0,
            'trnbirc1' :((trnbirc1.value/(1000000000*3600))/nbirc1 )if (nbirc1!=0) else 0,
            'trnbirc2' :((trnbirc2.value/(1000000000*3600))/nbirc2 )if (nbirc2!=0) else 0,
            'trnbirc3' :((trnbirc3.value/(1000000000*3600))/nbirc3 )if (nbirc3!=0) else 0,
            'trnbirc4' :((trnbirc4.value/(1000000000*3600))/nbirc4) if (nbirc4!=0) else 0,
            'mois' :  pd.to_datetime('{}/{}'.format(date.month,date.year),format\
                                     ='%m/%Y').date(),
            'an' : date.year
                                    }, ignore_index=True)
    df.index = df['mois']
    df = df.drop('mois', axis=1)
    df.to_excel(file_name)
    do.to_sql(df, db="helpdesk_kpi_tables_tr",mode=mode)
    print('Tr s’est bien déroulé')

def scan_cat_tc(start_month, month_end, year_start=2019, year_end=2020, file_name='helpdesk_kpi_tables_tc".xlsx',mode="replace"):
    """
    Scan la base et en sort l'indicateur moyen t1 pour chaque type de tickets 
    par mois selon leur processus. Écrit le résultat dans un fichier
    """
    df = pd.DataFrame({
        'tcnbdrc1' : [],
        'tcnbdrc2' : [],
        'tcnbdrc3' : [],
        'tcnbdrc4' : [],
        'tcnbirc1' : [],
        'tcnbirc2' : [],
        'tcnbirc3' : [],
        'tcnbirc4' : [],
        'tcnbbec' : [],
        'mois' : [],
        'an' : []
                })

    for year in range (year_start, year_end+1):
        if (year < year_end):
            end_month = 12
        else:
            end_month = month_end
        for i in range(start_month, end_month+1):
            nbdrc1 = 0
            nbdrc2 = 0
            nbdrc3 = 0
            nbdrc4 = 0
            nbirc1 = 0
            nbirc2 = 0
            nbirc3 = 0
            nbirc4 = 0
            nbbec = 0
            tcnbdrc1 = pd.Timedelta('0 hours')
            tcnbdrc2 = pd.Timedelta('0 hours')
            tcnbdrc3 = pd.Timedelta('0 hours')
            tcnbdrc4 = pd.Timedelta('0 hours')
            tcnbirc1 = pd.Timedelta('0 hours')
            tcnbirc2 = pd.Timedelta('0 hours')
            tcnbirc3 = pd.Timedelta('0 hours')
            tcnbirc4 = pd.Timedelta('0 hours')
            date = 0
            
            liste_demandes = do.l_o_d(month = i, year = year) 
            for j in liste_demandes:
                b = do.instantiate_a_ticket(j)
                    
                if (b is not None) and (b.est_clos_en_suivi): 
                    if (b.date_de_cloture.year == year) and (b.date_de_cloture.month == i):
                        date = b.date_de_cloture
                        
                        if (b.str_process == 'Process d rc 1'):
                            nbdrc1+=1
                            tcnbdrc1+=b.tc
                        elif (b.str_process == 'Process d rc 2'):
                            nbdrc2+=1
                            tcnbdrc2+=b.tc
                        elif (b.str_process == 'Process d rc 3'):
                            nbdrc3+=1
                            tcnbdrc3+=b.tc
                        elif (b.str_process == 'Process d rc 4'):
                            nbdrc4+=1
                            tcnbdrc4+=b.tc
                        elif (b.str_process == 'Process i rc 1'):
                            nbirc1+=1
                            tcnbirc1+=b.tc
                        elif (b.str_process == 'Process i rc 2'):
                            nbirc2+=1
                            tcnbirc2+=b.tc
                        elif (b.str_process == 'Process i rc 3'):
                            nbirc3+=1
                            tcnbirc3+=b.tc
                        elif (b.str_process == 'Process i rc 4'):
                            nbirc4+=1
                            tcnbirc4+=b.tc
                        else:
                            nbbec +=1
    
            df = df.append({
            'tcnbdrc1' :((tcnbdrc1.value/(1000000000*3600))/nbdrc1) if (nbdrc1!=0)else 0,
            'tcnbdrc2' :((tcnbdrc2.value/(1000000000*3600))/nbdrc2) if (nbdrc2!=0)else 0,
            'tcnbdrc3' :((tcnbdrc3.value/(1000000000*3600))/nbdrc3) if (nbdrc3!=0)else 0,
            'tcnbdrc4' :((tcnbdrc4.value/(1000000000*3600))/nbdrc4) if (nbdrc4!=0)else 0,
            'tcnbirc1' :((tcnbirc1.value/(1000000000*3600))/nbirc1 )if (nbirc1!=0) else 0,
            'tcnbirc2' :((tcnbirc2.value/(1000000000*3600))/nbirc2 )if (nbirc2!=0) else 0,
            'tcnbirc3' :((tcnbirc3.value/(1000000000*3600))/nbirc3 )if (nbirc3!=0) else 0,
            'tcnbirc4' :((tcnbirc4.value/(1000000000*3600))/nbirc4) if (nbirc4!=0) else 0,
            'mois' :  pd.to_datetime('{}/{}'.format(date.month,date.year),format\
                                     ='%m/%Y').date(),
            'an' : date.year
                                    }, ignore_index=True)
    df.index = df['mois']
    df = df.drop('mois', axis=1)
    df.to_excel(file_name)
    do.to_sql(df, db="helpdesk_kpi_tables_tc", mode=mode)
    print('Tc s’est bien déroulé')


def scan_cat_trc(start_month, month_end, year_start=2019, year_end=2020, file_name='helpdesk_kpi_tables_trc.xlsx', mode="replace"):
    """
    Scan la base et en sort l'indicateur moyen trc des processus n'étant que
    clôturés directement et PAS Résolue en booléen. Il d'adresse au processus
    noté d'un SEUL C . Car ceux-là ont une simple clôture et l'on mesure le 
    temps comme une résolution clôture, en même temps. Nous pourrions
    modifier le code pour réajuster ces valeurs
    """
    df = pd.DataFrame({
            'trcnbdc1' : [],#Nous considérons les process seulememt clos
        'trcnbdc2' : [],# La résolution n'entre pas en compte
        'trcnbdc3' : [],
        'trcnbdc4' : [],
        'trcnbic1' : [],
        'trcnbic2' : [],
        'trcnbic3' : [],
        'trcnbic4' : [],
        'trcnbbec' : [],
        'mois' : [],
        'an' : []
                })

    for year in range (year_start, year_end+1):
        if (year < year_end):
            end_month = 12
        else:
            end_month = month_end
        for i in range(start_month, end_month+1):
            nbdc1 = 0
            nbdc2 = 0
            nbdc3 = 0
            nbdc4 = 0
            nbic1 = 0
            nbic2 = 0
            nbic3 = 0
            nbic4 = 0
            nbbec = 0
            trcnbdc1 = pd.Timedelta('0 hours')
            trcnbdc2 = pd.Timedelta('0 hours')
            trcnbdc3 = pd.Timedelta('0 hours')
            trcnbdc4 = pd.Timedelta('0 hours')
            trcnbic1 = pd.Timedelta('0 hours')
            trcnbic2 = pd.Timedelta('0 hours')
            trcnbic3 = pd.Timedelta('0 hours')
            trcnbic4 = pd.Timedelta('0 hours')
            date = 0
            ##Chiffre 2200 à remettre à zéro lors de la production
            liste_demandes = do.l_o_d(month = i, year = year) 
            for j in liste_demandes:
                b = do.instantiate_a_ticket(j)
                    
                if (b is not None) and (b.est_clos_en_suivi): #and (b.str_process == 'Process i rc 1'):
                    if (b.date_de_cloture.year == year) and (b.date_de_cloture.month == i):
                        date = b.date_de_cloture
                        
                        if (b.str_process == 'Process d c 1'):
                            nbdc1+=1
                            trcnbdc1 += b.trc
                        elif (b.str_process == 'Process d c 2'):
                            nbdc2+=1
                            trcnbdc2+=b.trc
                        elif (b.str_process == 'Process d c 3'):
                            nbdc3+=1
                            trcnbdc3+=b.trc
                        elif (b.str_process == 'Process d c 4'):
                            nbdc4+=1
                            trcnbdc4+=b.trc
                        elif (b.str_process == 'Process i c 1'):
                            nbic1+=1
                            trcnbic1+=b.trc
                        elif (b.str_process == 'Process i c 2'):
                            nbic2+=1
                            trcnbic2+=b.trc
                        elif (b.str_process == 'Process i c 3'):
                            nbic3+=1
                            trcnbic3+=b.trc
                        elif (b.str_process == 'Process i c 4'):
                            nbic4+=1
                            trcnbic4+=b.trc
                        else:
                            nbbec +=1
    
          
            df = df.append({
            'trcnbdc1' : ( (trcnbdc1.value/(1000000000*3600))/nbdc1) if (nbdc1!=0)else 0,
            'trcnbdc2' : ((trcnbdc2.value/(1000000000*3600))/nbdc2) if (nbdc2!=0)else 0,
            'trcnbdc3' : ((trcnbdc3.value/(1000000000*3600))/nbdc3) if (nbdc3!=0)else 0,
            'trcnbdc4' : ((trcnbdc4.value/(1000000000*3600))/nbdc4) if (nbdc4!=0)else 0,
            'trcnbic1' : ((trcnbic1.value/(1000000000*3600))/nbic1) if (nbic1!=0)else 0,
            'trcnbic2' : ((trcnbic2.value/(1000000000*3600))/nbic2) if (nbic2!=0)else 0,
            'trcnbic3' : ((trcnbic3.value/(1000000000*3600))/nbic3) if (nbic3!=0)else 0,
            'trcnbic4' : ((trcnbic4.value/(1000000000*3600))/nbic4) if (nbic4!=0)else 0,
            'mois' :  pd.to_datetime('{}/{}'.format(date.month,date.year),format\
                                     ='%m/%Y').date(),
            'an' : date.year
                                    }, ignore_index=True)
    df.index = df['mois']
    df = df.drop('mois', axis=1)
    df.to_excel(file_name)
    do.to_sql(df, db="helpdesk_kpi_tables_trc", mode=mode)
    print('Trc s’est bien déroulé')

def moy_tt_by_month(start_month, month_end, year_start=2019, year_end=2020, file_name='helpdesk_kpi_tables_ratio_tt.xlsx', mode="replace"):
    """
    Moyenne ticket technicien par jour dans un mois donnée
    Requête la base pour les tickets d'un mois donné. Renvoie la moyenne du
    nombre de tickets par le nombre de technicien en moyenne par jour
    """
    df = pd.DataFrame({
        'tickets' : [],#Nous considérons les process seulememt clos
        'techs' : [],# La résolution n'entre pas en compte
        'moyenne' : [],
        'mois' : [],
        
                })
    for year in range (year_start, year_end+1):
        if (year < year_end):
            end_month = 12
        else:
            end_month = month_end
        for i in range(start_month, end_month+1):# i est le mois "month"
            nbtick = 0.
            nbtech = 0.
            moyenne = 0.
            date = pd.to_datetime('{}-{}'.format(year, i,format="%Y-%m"))
            nbtick,nbtech,moyenne = ratio_mois(i,year)
            df = df.append({
            'tickets' : nbtick,
            'techs' : nbtech,
            'moyenne':moyenne,
            'mois' :  pd.to_datetime('{}/{}'.format(date.month,date.year),\
                                     format='%m/%Y').date(),
            
                                    }, ignore_index=True)
    df.index = df['mois']
    df = df.drop('mois', axis=1)
    df.to_excel(file_name)
    do.to_sql(df, db="helpdesk_kpi_tables_ratio_tt", mode=mode)
    print('Ratio Tt s’est bien déroulé')

def moy_tt_by_day_in_a_month(start_month, month_end, year_start=2019, year_end=2020, file_name='helpdesk_kpi_tables_ratio_tt_by_day.xlsx', mode="replace"):
    """
    Moyenne ticket technicien par jour dans un mois donnée
    Requête la base pour les tickets d'un mois donné. Renvoie la moyenne du
    nombre de tickets par le nombre de technicien en moyenne par jour
    """
    df = pd.DataFrame({
        'tickets' : [],#Nous considérons les process seulememt clos
        'techs' : [],# La résolution n'entre pas en compte
        'moyenne' : [],
        'jour' : []
                })
    for year in range (year_start, year_end+1):
        if (year < year_end):
            end_month = 12
        else:
            end_month = month_end
        for month in range(start_month, end_month+1):# i est le mois "month"
            nbtick = 0.
            nbtech = 0.
            moyenne = 0.
            
            
            for day in range (1,32):
                try:
                    date = pd.to_datetime('{}-{}-{}'.format(year, month,day),format="%Y-%m-%d")
                    if is_busday(pd.to_datetime("{}-{}-{}".format(year,month,day),format='%Y-%m-%d')):
                        nbtick,nbtech,moyenne = ratio_jour(day,month,year)

                        df = df.append({
                        'tickets' : nbtick,
                        'techs' : nbtech,
                        'moyenne':moyenne,
                        'jour' :  pd.to_datetime('{}/{}/{}'.format(date.day,date.month,date.year),\
                                                 format='%d/%m/%Y').date()                        
                                                }, ignore_index=True)
                except:
                    continue
                   # print('Le jour {}/{}/{} n\'existe pas '.format(day, month, year))
        
    df.index = df['jour']
    df = df.drop('jour', axis=1)
    df.to_excel(file_name)
    do.to_sql(df, db="helpdesk_kpi_tables_ratio_tt_by_day", mode=mode)
    print('Ratio Tt par jour s’est bien déroulé')


def ratio_jour(day, month, year):
    """
    Calcul le nombre de tickets quotidiens, puis le nombre de techniciens qui 
    sont intervenus sur au moins un ticket, quotidiennement. La fonction
    renvoie aussi le ratio nombre de ticket par nombre de technicien.
    """
    liste_interv = do.r_to_b(query="SELECT `id_demande` FROM suivi_intervention WHERE MONTH(date_suivi) = " + str(month)+ " AND YEAR(date_suivi)=" + str(year) + " AND DAY(date_suivi)=" + str(day) +" AND id_statut_demande = 6")
    liste_dem = do.r_to_b(query="SELECT `id_demande` FROM `suivi_demande_ressource` WHERE MONTH(date_suivi) = " + str(month)+ " AND YEAR(date_suivi)=" + str(year) + " AND DAY(date_suivi)=" + str(day) +" AND id_statut_demande = 6")
    liste_tech_interv = do.r_to_b(query="SELECT `id_technicien_creation` FROM suivi_intervention WHERE MONTH(date_suivi) = " + str(month) + " AND YEAR(date_suivi)=" + str(year) + " AND DAY(date_suivi)=" + str(day) )
    liste_tech_dem_ress = do.r_to_b(query="SELECT `id_technicien` FROM `suivi_demande_ressource` WHERE MONTH(date_suivi) = " + str(month) + " AND YEAR(date_suivi)=" + str(year) + " AND DAY(date_suivi)=" + str(day) )
    liste_tech_pec = do.r_to_b(query="SELECT `id_technicien` FROM `prise_en_charge` WHERE MONTH(date_prise_en_charge) =" + str(month) + " AND YEAR(date_prise_en_charge)=" + str(year) + " AND DAY(date_prise_en_charge)=" + str(day) )
    nbtick = float(liste_interv.count() + liste_dem.count())
#    nbtech = float(len(pd.Series(liste_tech_interv['id_technicien_creation']).unique()))  
 #   nbtech += float(len(pd.Series(liste_tech_dem_ress['id_technicien']).unique()))
    liste = liste_tech_interv['id_technicien_creation']
    liste = liste_tech_interv['id_technicien_creation'].append(liste_tech_dem_ress['id_technicien']\
                             .append(liste_tech_pec['id_technicien'], ignore_index=True), ignore_index=True)
    nbtech = float(len(pd.Series(liste).unique()))
    #moyenneheure = 
    #print('Nombre de tech',nbtech)
    """ À détruire après une période de quarantaine
    for i in liste_interv['id_demande'].values:
        b= do.instantiate_a_ticket(i)
        if b is not None and b.est_clos_en_suivi:
            nbtick+=1
    for i in liste_dem['id_demande'].values:
        b= do.instantiate_a_ticket(i)
        if b is not None and b.est_clos_en_suivi:
            nbtick+=1
    """        
    return (nbtick,nbtech,(nbtick/nbtech)) if (nbtech != 0.) else (nbtick,nbtick,0.)

def ratio_mois(month,year):
    nbtick=0.
    nbtech = 0.
    nbday = 0.
    ratio = 0.
    for day in range (1,32):
        try:
            if is_busday(pd.to_datetime("{}-{}-{}".format(year,month,day),format='%Y-%m-%d')):
                nbticktemp,nbtechtemp,ratiotemp = ratio_jour(day,month,year)
                nbtick += nbticktemp
                nbtech += nbtechtemp
                ratio += ratiotemp
                nbday = 1 + nbday
        except:
            continue
            #print('Le jour {}/{}/{} n\'existe pas '.format(i, month, year))
    #print('jour',nbday)
    #print('ratio',ratio)
    return (nbtick/nbday, nbtech/nbday, (ratio/nbday)) if (nbday != 0.) else (nbtick,nbtick,0.)

def productivity(mstart,mend,ystart, yend):
    """
    Demande de beaucoup de réflexion en économie et ne colle pas à une information 
    pertinente dans le métier du ticketing : un homme peut faire autant de ticket que 12 en 
    moins de temps. Donc cet indicateur n’est pas 'relevant'
    Calcul la productivité marginale, le nombre moyen de ticket par technicien en
    intégrant l’échantillonnage de chaque mois pour fixer des limites supérieures et 
    inférieur comme autant d’écart type (2).
    """
    data = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_ratio_tt_by_day")
    dfend = pd.DataFrame({
        'tickets' : [],
        'techs' : [],
        'date' : []
                })

#    print(data['techs'].unique())
    a = [int(x) for x in data['techs'].unique()]
 #   print(a)
    for year in range (ystart, yend+1):
        for month in range( 1, 12+1 ):
            if ( year == ystart ) & (month < mstart):
                continue
            if ( year >= yend ) & ( month >= mend+1 ):
                continue
            if month != 12:
                #df = data.loc[(pd.to_datetime(data['date_de_cloture'].dropna(),format='%Y-%m-%d')>=pd.datetime(year,month,1))&(data['Tt']<400*3600000000000)&(pd.to_datetime(data['date_de_cloture'].dropna(),format='%Y-%m-%d')<pd.datetime(year,month + 1,1))]
                df = data.loc[(pd.to_datetime(data['jour'],format='%Y-%m-%d')>=pd.datetime(year,month,1))&(pd.to_datetime(data['jour'],format='%Y-%m-%d')<pd.datetime(year,month+1,1))]
            else:
                df = data.loc[(pd.to_datetime(data['jour'],format='%Y-%m-%d')>=pd.datetime(year,month,1))&(pd.to_datetime(data['jour'].dropna(),format='%Y-%m-%d')<pd.datetime(year + 1,1,1))]
            for i in a:#i est le nombre de technicien
                if not df.loc[df['techs']==i].empty:
                    nbtick = round(df.loc[df['techs']==i]['tickets'].mean(),2)
                    moyheure = round(df.loc[df['techs']==i]['moyenne'].mean(),2)
                    nbtech = i
                    if not df.loc[df['techs']==i].empty:
                        high = round(df.loc[df['techs']==i]['tickets'].mean()+df.loc[df['techs']==i]['tickets'].std(),2)
                        low = round(df.loc[df['techs']==i]['tickets'].mean()-df.loc[df['techs']==i]['tickets'].std(),2)
                    else:
                        high = 0
                        low = 0
#                    print(i, 'technicien, le ',year, month, 'ont fait ',round(df.loc[df['techs']==i]['tickets'].mean(),2))
 #                   print(i, round(df.loc[df['techs']==i]['tickets'].mean() + df.loc[df['techs']==i]['tickets'].std(),2),round(df.loc[df['techs']==i]['tickets'].mean() - (df.loc[df['techs']==i]['tickets'].std()),2))
                    date = pd.to_datetime('{}-{}-{}'.format(year, month,1),format="%Y-%m-%d")
                    dfend = dfend.append({
                    'tickets' : nbtick,
                    'techs' : nbtech,
                    'moyenne_heure':moyheure,
                    'high':high,
                    'low':low,
                    'date' :  pd.to_datetime('{}/{}/{}'.format(date.day,date.month,date.year),\
                                             format='%d/%m/%Y').date()                        
                                            }, ignore_index=True)
            dfend = dfend.append({
            'tickets' : 0,
            'techs' : 0,
            'moyenne_heure': 0,
            'high':0,
            'low':0,
            'date' :  pd.to_datetime('{}/{}/{}'.format(date.day,date.month,date.year),\
                                     format='%d/%m/%Y').date()                        
                                    }, ignore_index=True)
    dfend.index = dfend['date']
    dfend = dfend.drop('date', axis=1)
    dfend = dfend.sort_values(by=['date','techs'])

    dfend['marginal_productivity'] = (dfend['tickets']-dfend['tickets'].shift(1))/(dfend['techs']-dfend['techs'].shift(1))
    dfend['marginal_productivity'].loc[dfend['techs']==0] = 0
    dfend.to_excel('productivity.xlsx')
    do.to_sql(dfend, db="helpdesk_kpi_tables_productivity")
    
def productivity_synopsis(mstart,mend,ystart, yend, ub = 40, um = 16 ,uh= 2,uth= 1):
    """
    Calcul de la productivité en heure et selon le nombre de techniciens participant à un ticket donné.
    De plus, d’après l’urgence du ticket considéré, une valeur de besoin client en heure est 
    ajouté sur la ligne du ticket. Cela permet de comparer la demande en heure et l’offre de service en heure
    
    La table crée permet dans les fonctions :
        productivity_synopsis_day
        productivity_synopsis_month
        day_ticket
        day_prob
    de créer des information sur la productivité quotidienne et mensuelle, et même de calculer la probabilité
    de ticket par jour.
    """
    data = do.r_to_b(query="SELECT * FROM synopsis")
    dfend = pd.DataFrame({
        'id_demande':[],
        'tickets' : [],
        'techs' : [],
        'nb_techs':[],
        'heure':[],
        'date' : [],
        'urgence':[]
                })
    for year in range (ystart, yend+1):
        for month in range( 1, 12+1 ):
            if ( year == ystart ) & (month < mstart):
                continue
            if ( year >= yend ) & ( month >= mend+1 ):
                continue
            if month != 12:
                df = data.loc[(pd.to_datetime(data['date_cloture'],format='%Y-%m-%d')>=pd.datetime(year,month,1))&(pd.to_datetime(data['date_cloture'],format='%Y-%m-%d')<pd.datetime(year,month+1,1))]
            else:
                df = data.loc[(pd.to_datetime(data['date_cloture'],format='%Y-%m-%d')>=pd.datetime(year,month,1))&(pd.to_datetime(data['date_cloture'].dropna(),format='%Y-%m-%d')<pd.datetime(year + 1,1,1))]
            list_of_dates = pd.to_datetime(df['date_cloture']).dt.strftime('%Y/%m/%d').unique().tolist()
            for date in list_of_dates:#i est le nombre de technicien
                tickets = df.loc[(pd.to_datetime(df['date_cloture'])>=pd.to_datetime(date))&(pd.to_datetime(df['date_cloture'])<(pd.to_datetime(date)+timedelta(days = 1)))]
                list_of_ticket = tickets['id_demande'].unique().tolist()
                for id_demande in list_of_ticket:
                    tech = tickets.loc[tickets['id_demande']==id_demande]['Tech_de_resolution'].unique().tolist()
                    tech = ", ".join(tech)
                    nbtech = tickets.loc[tickets['id_demande']==id_demande]['Tech_de_resolution'].nunique()
                    heures = tickets.loc[(tickets['id_demande']==id_demande)&(tickets['id_statut_demande']==6)][['attente_n1_time','attente_t2_time','attente_t3_time','Time_of_resolution']]
                    heuretemp = heures['attente_n1_time']+heures['attente_t2_time']+heures['attente_t3_time']+heures['Time_of_resolution']
                    heure = round(heuretemp.values[0],2)
                    urgence = tickets.loc[(tickets['id_demande']==id_demande)&(tickets['id_statut_demande']==6)]['urgence']
                    #print(urgence.values[0])
                    if urgence.values[0] == 1:
                        urgence_heure = ub
                    if urgence.values[0] == 2:
                        urgence_heure = um
                    if urgence.values[0] == 3:
                        urgence_heure = uh
                    if urgence.values[0] == 4:
                        urgence_heure = uth
                    #date = pd.to_datetime('{}-{}-{}'.format(year, month,1),format="%Y-%m-%d")
                    dfend = dfend.append({
                    'id_demande': id_demande,
                    'tickets' : 1,
                    'techs' : tech,
                    'nb_techs': nbtech,
                    'heure':heure,
                    'urgence': urgence.values[0],
                    'heure_client':urgence_heure,
                    
    #                'date' :  pd.to_datetime(tickets.loc[(tickets['id_demande']==id_demande)&(tickets['id_statut_demande']==6)]['date_cloture'].iloc[0])}, ignore_index=True)
                    'date' :  pd.to_datetime(date)}, ignore_index=True)
            
    dfend.index = dfend['date']
    dfend = dfend.drop('date', axis=1)
    dfend = dfend.sort_values(by=['date'])
    #dfend['marginal_productivity'] = (dfend['tickets']-dfend['tickets'].shift(1))/(dfend['techs']-dfend['techs'].shift(1))
    #dfend['marginal_productivity'].loc[dfend['techs']==0] = 0
    dfend.to_excel('productivity_synopsis.xlsx')
    do.to_sql(dfend, db="helpdesk_kpi_tables_productivity_synopsis")
    
def productivity_synopsis_creation(mstart,mend,ystart, yend, ub = 40, um = 16 ,uh= 2,uth= 1):
    """
    Calcul la productivité marginale, le nombre moyen de ticket par technicien en
    intégrant l’échantillonnage de chaque mois pour fixer des limites supérieures et 
    inférieur comme autant d’écart type (2).
    """
    data = do.r_to_b(query="SELECT * FROM synopsis")
    dfend = pd.DataFrame({
        'id_demande':[],
        'tickets' : [],
        'techs' : [],
        'nb_techs':[],
        'heure':[],
        'date' : [],
        'urgence':[]
                })
    for year in range (ystart, yend+1):
        for month in range( 1, 12+1 ):
            if ( year == ystart ) & (month < mstart):
                continue
            if ( year >= yend ) & ( month >= mend+1 ):
                continue
            if month != 12:
                df = data.loc[(pd.to_datetime(data['date_creation'],format='%Y-%m-%d')>=pd.datetime(year,month,1))&(pd.to_datetime(data['date_creation'],format='%Y-%m-%d')<pd.datetime(year,month+1,1))]
            else:
                df = data.loc[(pd.to_datetime(data['date_creation'],format='%Y-%m-%d')>=pd.datetime(year,month,1))&(pd.to_datetime(data['date_creation'].dropna(),format='%Y-%m-%d')<pd.datetime(year + 1,1,1))]
            list_of_dates = pd.to_datetime(df['date_creation']).dt.strftime('%Y/%m/%d').unique().tolist()
            for date in list_of_dates:#i est le nombre de technicien
                tickets = df.loc[(pd.to_datetime(df['date_creation'])>=pd.to_datetime(date))&(pd.to_datetime(df['date_creation'])<(pd.to_datetime(date)+timedelta(days = 1)))]
                list_of_ticket = tickets['id_demande'].unique().tolist()
                for id_demande in list_of_ticket:
                    tech = tickets.loc[tickets['id_demande']==id_demande]['Tech_de_resolution'].unique().tolist()
                    tech = ", ".join(tech)
                    nbtech = tickets.loc[tickets['id_demande']==id_demande]['Tech_de_resolution'].nunique()
                    heures = tickets.loc[(tickets['id_demande']==id_demande)&(tickets['id_statut_demande']==6)][['attente_n1_time','attente_t2_time','attente_t3_time','Time_of_resolution']]
                    heuretemp = heures['attente_n1_time']+heures['attente_t2_time']+heures['attente_t3_time']+heures['Time_of_resolution']
                    heure = round(heuretemp.values[0],2)
                    urgence = tickets.loc[(tickets['id_demande']==id_demande)&(tickets['id_statut_demande']==6)]['urgence']
                    #print(urgence.values[0])
                    if urgence.values[0] == 1:
                        urgence_heure = ub
                    if urgence.values[0] == 2:
                        urgence_heure = um
                    if urgence.values[0] == 3:
                        urgence_heure = uh
                    if urgence.values[0] == 4:
                        urgence_heure = uth
                    #date = pd.to_datetime('{}-{}-{}'.format(year, month,1),format="%Y-%m-%d")
                    dfend = dfend.append({
                    'id_demande': id_demande,
                    'tickets' : 1,
                    'techs' : tech,
                    'nb_techs': nbtech,
                    'heure':heure,
                    'urgence': urgence.values[0],
                    'heure_client':urgence_heure,
                    
    #                'date' :  pd.to_datetime(tickets.loc[(tickets['id_demande']==id_demande)&(tickets['id_statut_demande']==6)]['date_cloture'].iloc[0])}, ignore_index=True)
                    'date' :  pd.to_datetime(date)}, ignore_index=True)
            
    dfend.index = dfend['date']
    dfend = dfend.drop('date', axis=1)
    dfend = dfend.sort_values(by=['date'])
    #dfend['marginal_productivity'] = (dfend['tickets']-dfend['tickets'].shift(1))/(dfend['techs']-dfend['techs'].shift(1))
    #dfend['marginal_productivity'].loc[dfend['techs']==0] = 0
    dfend.to_excel('productivity_synopsis_creation.xlsx')
    do.to_sql(dfend, db="helpdesk_kpi_tables_productivity_synopsis_creation")
    
def productivity_synopsis_day(mstart,mend,ystart, yend):
    """
    Calcul la productivité marginale, le nombre moyen de ticket par technicien en
    intégrant l’échantillonnage de chaque mois pour fixer des limites supérieures et 
    inférieur comme autant d’écart type (2).
    """
    data = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_productivity_synopsis")
    dfend = data[['date','nb_techs','tickets','heure','heure_client']]
    dfend = dfend.groupby(['date','nb_techs']).sum()
    dfend['heure_moyenne_groupe'] = dfend['heure']/dfend['tickets']
    dfend = dfend.reset_index()
    dfend['heure_moyenne_par_tech'] = dfend['heure_moyenne_groupe']/dfend['nb_techs']
    #dfend.index = dfend['date']    
    #dfend = dfend.drop('heure')
    #print(dfend)

    dfend.to_excel('productivity_synopsis_productivity_day.xlsx')
    do.to_sql(dfend, db="helpdesk_kpi_tables_productivity_synopsis_day")
    
def productivity_synopsis_day_creation(mstart,mend,ystart, yend):
    """
    Calcul la productivité marginale, le nombre moyen de ticket par technicien en
    intégrant l’échantillonnage de chaque mois pour fixer des limites supérieures et 
    inférieur comme autant d’écart type (2).
    """
    data = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_productivity_synopsis_creation")
    dfend = data[['date','nb_techs','tickets','heure','heure_client']]
    dfend = dfend.groupby(['date','nb_techs']).sum()
    dfend['heure_moyenne_groupe'] = dfend['heure']/dfend['tickets']
    dfend = dfend.reset_index()
    dfend['heure_moyenne_par_tech'] = dfend['heure_moyenne_groupe']/dfend['nb_techs']
    #dfend.index = dfend['date']    
    #dfend = dfend.drop('heure')
    #print(dfend)

    dfend.to_excel('productivity_synopsis_productivity_day_creation.xlsx')
    do.to_sql(dfend, db="helpdesk_kpi_tables_productivity_synopsis_day_creation")
    

def productivity_synopsis_month(mstart,mend,ystart, yend, cout_horaire = 10):
    """
    Calcul la productivité par mois selon le nombre de technicien qui interviennent.
    """
    data = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_productivity_synopsis")
    data['date'] = pd.to_datetime(data['date'].dt.strftime('%m/1/%Y'))
    #print(data['date'],data['date_month'])#print(pd.DatetimeIndex(data['date']).year)
    dfend = data[['date','nb_techs','tickets','heure','heure_client']]
    dfend = dfend.groupby(['date','nb_techs']).sum()
    dfend['heure_moyenne_groupe'] = dfend['heure']/dfend['tickets']
    dfend = dfend.reset_index()
    dfend['heure_moyenne_par_tech'] = dfend['heure_moyenne_groupe']/dfend['nb_techs']
    dfend['cout_horaire'] = 10
    dfend['prix_par_ticket'] = round(dfend['heure_moyenne_groupe'] * cout_horaire,2)
    #print(dfend)

    dfend.to_excel('productivity_synopsis_productivity_month.xlsx')
    do.to_sql(dfend, db="helpdesk_kpi_tables_productivity_synopsis_month")

def productivity_synopsis_month_creation(mstart,mend,ystart, yend, cout_horaire = 10):
    """
    Calcul la productivité par mois selon le nombre de technicien qui interviennent.
    """
    data = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_productivity_synopsis_creation")
    data['date'] = pd.to_datetime(data['date'].dt.strftime('%m/1/%Y'))
    #print(data['date'],data['date_month'])#print(pd.DatetimeIndex(data['date']).year)
    dfend = data[['date','nb_techs','tickets','heure','heure_client']]
    dfend = dfend.groupby(['date','nb_techs']).sum()
    dfend['heure_moyenne_groupe'] = dfend['heure']/dfend['tickets']
    dfend = dfend.reset_index()
    dfend['heure_moyenne_par_tech'] = dfend['heure_moyenne_groupe']/dfend['nb_techs']
    dfend['cout_horaire'] = 10
    dfend['prix_par_ticket'] = round(dfend['heure_moyenne_groupe'] * cout_horaire,2)
    #print(dfend)

    dfend.to_excel('productivity_synopsis_productivity_month_creation.xlsx')
    do.to_sql(dfend, db="helpdesk_kpi_tables_productivity_synopsis_month_creation")
    
def day_ticket(mstart,mend,ystart, yend):
    """
    Calcul la quantité de ticket chaque jour avec la date en index.
    """
    data = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_productivity_synopsis")
    dfend = data[['date','tickets']]#,'heure','heure_client']]
    #print(data)
    dfend = dfend.groupby(['date']).sum()
  #  print(dfend)

    dfend.to_excel('day_tickets.xlsx')
    do.to_sql(dfend, db="day_tickets")

def day_ticket_volatility(mstart,mend,ystart, yend):
    """
    Calcul la quantité de ticket chaque jour avec la date en index.
    """
    data = do.r_to_b(query="SELECT * FROM day_tickets")
    #data = pd.DataFrame(data[['date','tickets']], index=data['date'])
    data.index=data['date']
    data.rename(columns={'tickets':'volat_closed'}, inplace=True)
    df_filled = data['volat_closed'].asfreq('D', method='ffill') 
    df_returns = df_filled.pct_change()   
    df_std = df_returns.rolling(window=30, min_periods=30).std()
    df_std = df_std.fillna(0)

    #df_std.plot();
    
    df_std.to_excel('day_tickets_volatility.xlsx')
    do.to_sql(df_std, db="day_tickets_volatility")
    

def day_creation_ticket(mstart,mend,ystart, yend):
    """
    Calcul la quantité de ticket chaque jour avec la date en index.
    """
    data = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_productivity_synopsis_creation")
    dfend = data[['date','tickets']]#,'heure','heure_client']]
    dfend = dfend.groupby(['date']).sum()
 #   print(dfend)
    dfend.to_excel('day_creation_tickets.xlsx')
    do.to_sql(dfend, db="day_creation_tickets")


def day_creation_ticket_volatility(mstart,mend,ystart, yend):
    """
    Calcul la quantité de ticket chaque jour avec la date en index.
    """
    data = do.r_to_b(query="SELECT * FROM day_creation_tickets")
    #data = pd.DataFrame(data[['date','tickets']], index=data['date'])
    data.index=data['date']
    data.rename(columns={'tickets':'volat_creation'}, inplace=True)
    df_filled = data['volat_creation'].asfreq('D', method='ffill') 
    df_returns = df_filled.pct_change()   
    df_std = df_returns.rolling(window=30, min_periods=30).std()
    #df_std.plot();
    df_std = df_std.fillna(0)
    df_std.to_excel('day_creation_tickets_volatility.xlsx')
    do.to_sql(df_std, db="day_creation_tickets_volatility")
    

def day_prob(mstart,mend,ystart, yend):
    """
    Calcul sur une période la table de fréquence des tickets, en éliminant les jours fériés.
    """
    # Chargement des jours fériés
    holidayfile ='jours-feries-seuls.csv'   
    df_holiday = pd.read_csv(holidayfile)
    df_holiday['date'] = pd.to_datetime(df_holiday['date'], format='%Y-%m-%d')
    holidays = np.array(df_holiday['date'], dtype='datetime64[D]')
    # Chargement de la période considérés en calendrier
    dates_load = do.r_to_b(query="SELECT * FROM times WHERE (Month >= "+str(mstart)+" AND Year = "+str(ystart)+
                                                        ") OR (Month <= " + str(mend) + " AND Year = " + str(yend)+");")
    dates = dates_load
    dates = dates.loc[np.is_busday(pd.to_datetime(dates['Date']).values.astype('datetime64[D]'),holidays=holidays)]['Date']
    dates.index = dates
    data = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_productivity_synopsis WHERE (MONTH(date) >= "+str(mstart)+" AND YEAR(date) = "+str(ystart)+
                                                        ") OR (MONTH(date) <= " + str(mend) + " AND YEAR(date) = " + str(yend)+");")
    dfend = data[['date','tickets']]#,'heure','heure_client']]
    dfend = dfend.groupby(['date']).sum()
    result = pd.concat([dates, dfend], axis=1, sort=False)
    result.columns =['Date','nb_tickets']
    result = result.fillna(0)
    result['nb_tickets']= result['nb_tickets'].astype(int)
    result.index.name = 'Date'
    result = result.drop('Date',axis=1)
    #print(result)
    #print(result.mean())
    #print(result.nunique(), result.count())
    #print(result['nb_tickets'].value_counts(normalize=True).iloc[0:].values)
    #a = result['nb_tickets'].value_counts(normalize=True)
    #a=pd.DataFrame(a)
    #a=a.sort_index()
    #x = result['nb_tickets'].sort_values().value_counts(normalize=True).index.values
    #y = result['nb_tickets'].sort_values().value_counts(normalize=True).iloc[0:].values
    #print(x,y)
    #plt.stem(x, y, use_line_collection=True)
    #plt.show()
    #result.to_excel('day_tickets.xlsx')
    #do.to_sql(result, db="day_tickets")
    count_value_norm = pd.DataFrame(result['nb_tickets'].value_counts(normalize=True)).sort_index()
    count_value_norm.to_excel('day_tickets_proba.xlsx')
    do.to_sql(count_value_norm, db="day_tickets_proba")
    #print(count_value_norm)
    ## GET THE BEST DISTRIBUTION
    # Get the best distribution
    values = result['nb_tickets']
    data = pd.Series(values)
    params = fit_to_all_distributions(data)
    best_dist_chi, best_chi, best_p, best_expected_values,best_observed_values, params_chi, dist_results_chi = get_best_distribution_using_chisquared_test(values, params)
    dataframe_chi_goodness_of_fit = pd.DataFrame({'Best_dist_chi': best_dist_chi,'Best_chi': best_chi, 'Best_p': best_p}, index=[0])
    do.to_sql(dataframe_chi_goodness_of_fit, db="chisquare_gof_closed_tickets")
    #Plot the result
    dataframe_plot = pd.DataFrame({'Obs':best_observed_values,'Exp':best_expected_values})
    do.to_sql(dataframe_plot, db="day_tickets_proba_plot")

def day_creation_prob(mstart,mend,ystart, yend):
    """
    Calcul sur une période la table de fréquence des tickets, à partir du jour de création, en éliminant les jours fériés.
    """
    # Chargement des jours fériés
    holidayfile ='jours-feries-seuls.csv'   
    df_holiday = pd.read_csv(holidayfile)
    df_holiday['date'] = pd.to_datetime(df_holiday['date'], format='%Y-%m-%d')
    holidays = np.array(df_holiday['date'], dtype='datetime64[D]')
    # Chargement de la période considérés en calendrier
    dates_load = do.r_to_b(query="SELECT * FROM times WHERE (Month >= "+str(mstart)+" AND Year = "+str(ystart)+
                                                        ") OR (Month <= " + str(mend) + " AND Year = " + str(yend)+");")
    dates = dates_load
    dates = dates.loc[np.is_busday(pd.to_datetime(dates['Date']).values.astype('datetime64[D]'),holidays=holidays)]['Date']
    dates.index = dates
    data = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_productivity_synopsis_creation WHERE (MONTH(date) >= "+str(mstart)+" AND YEAR(date) = "+str(ystart)+
                                                        ") OR (MONTH(date) <= " + str(mend) + " AND YEAR(date) = " + str(yend)+");")
    dfend = data[['date','tickets']]#,'heure','heure_client']]
    dfend = dfend.groupby(['date']).sum()
    result = pd.concat([dates, dfend], axis=1, sort=False)
    result.columns =['Date','nb_tickets']
    result = result.fillna(0)
    result['nb_tickets']= result['nb_tickets'].astype(int)
    result.index.name = 'Date'
    result = result.drop('Date',axis=1)
    #print(result)
    #print(result.mean())
    #print(result.nunique(), result.count())
    #print(result['nb_tickets'].value_counts(normalize=True).iloc[0:].values)
    #a = result['nb_tickets'].value_counts(normalize=True)
    #a=pd.DataFrame(a)
    #a=a.sort_index()
    #x = result['nb_tickets'].sort_values().value_counts(normalize=True).index.values
    #y = result['nb_tickets'].sort_values().value_counts(normalize=True).iloc[0:].values
    #print(x,y)
    #plt.stem(x, y, use_line_collection=True)
    #plt.show()
    #result.to_excel('day_tickets.xlsx')
    #do.to_sql(result, db="day_tickets")
    count_value_norm = pd.DataFrame(result['nb_tickets'].value_counts(normalize=True)).sort_index()
    count_value_norm.to_excel('day_tickets_creation_proba.xlsx')
    do.to_sql(count_value_norm, db="day_tickets_creation_proba")
    #print(count_value_norm)
    ## GET THE BEST DISTRIBUTION
    # Get the best distribution
    #a, m = 3., 2.
    values = result['nb_tickets']
    data = pd.Series(values)
    params = fit_to_all_distributions(data)
    best_dist_chi, best_chi, best_p, best_expected_values,best_observed_values, params_chi, dist_results_chi = get_best_distribution_using_chisquared_test(values, params)
    dataframe_chi_goodness_of_fit = pd.DataFrame({'Best_dist_chi': best_dist_chi,'Best_chi': best_chi, 'Best_p': best_p}, index=[0])
    do.to_sql(dataframe_chi_goodness_of_fit, db="chisquare_gof_opened_tickets")
    #Plot the result
    dataframe_plot = pd.DataFrame({'Obs':best_observed_values,'Exp':best_expected_values})
    do.to_sql(dataframe_plot, db="day_tickets_creation_proba_plot")

def ticket_cloture_jour_ferie(mstart,mend,ystart, yend):
    """
    Calcul sur une période la table de fréquence des tickets, en ne comptant que les cloture s’étant déroulé un jour férié.
    """
    holidayfile ='jours-feries-seuls.csv'   
    df_holiday = pd.read_csv(holidayfile)
    df_holiday['date'] = pd.to_datetime(df_holiday['date'], format='%Y-%m-%d')
    holidays = pd.DataFrame(pd.to_datetime(df_holiday['date']))
    dates_load = do.r_to_b(query="SELECT * FROM times WHERE (Month >= "+str(mstart)+" AND Year = "+str(ystart)+
                                        ") OR (Month <= " + str(mend) + " AND Year = " + str(yend)+");")
    dates = dates_load
    dates.index = dates['Date']
    dates = dates['Date']
    data = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_productivity_synopsis")
    dfend = data[['date','tickets']]
    result = dfend
    result.columns = ['Date','nb_tickets']
    result.loc[:]['nb_tickets'] = result.loc[:]['nb_tickets'].astype(int)
    result_filtered = result.loc[pd.to_datetime(result['Date']).isin( holidays['date'])]
    result_filtered = result_filtered.reset_index()
    #result_filtered.index = result_filtered['Date']    
    #print(result_filtered)
    result_filtered.to_excel('day_out_of_work.xlsx')
    do.to_sql(result_filtered, db="day_out_of_work")
    
def make(month_start,month_end,year_start,year_end,urg_basse = 40, urg_moy = 16, urg_haute = 2, urg_th = 1):
    """
    Crée des statistiques à partir de la table universelle.
    """
    
    productivity_synopsis(month_start,month_end,year_start,year_end,urg_basse,urg_moy,urg_haute,urg_th)#les 4 derniers sont les durées SLA des urgences
    productivity_synopsis_creation(month_start,month_end,year_start,year_end,urg_basse,urg_moy,urg_haute,urg_th)#les 4 derniers sont les urgences
    productivity_synopsis_day(month_start,month_end,year_start,year_end)
    productivity_synopsis_day_creation(month_start,month_end,year_start,year_end)
    productivity_synopsis_month(month_start,month_end,year_start,year_end)
    productivity_synopsis_month_creation(month_start,month_end,year_start,year_end)
    day_ticket(month_start,month_end,year_start,year_end)
    day_creation_ticket(month_start,month_end,year_start,year_end)
    day_ticket_volatility(month_start,month_end,year_start,year_end)
    day_creation_ticket_volatility(month_start,month_end,year_start,year_end)
    day_prob(month_start,month_end,year_start,year_end)
    day_creation_prob(month_start,month_end,year_start,year_end)
