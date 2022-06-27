# -*- coding: utf-8 -*-
"""
Created on Tue Jun  9 15:48:09 2020

@author: rboyrie
"""
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
from scipy.optimize import leastsq
import pandas as pd
from machine import do

def do_regression_times(month_start,month_end,year_start,year_end):
    #data = pd.read_excel('Synopsis Continuum Tessera Tabular - vue panoramique des tickets.xlsx')
    data =  do.r_to_b(query='SELECT * FROM synopsis')
    data = data.loc[data['id_statut_demande']==6]
    temps_de_resolution = data.loc[data['id_statut_demande']==6]["Time_of_resolution"]
    temps_de_treatment = data["treatment_time"]
    date = data['date_cloture'].dropna()
    y = temps_de_treatment.dropna()
    x = temps_de_resolution.dropna()
    process = data['process_label']
    process_list = process.dropna().unique().astype(int)
    str_type = data['str_type']
    id_demande = data['id_demande']
    df = pd.DataFrame({'x':x,'y':y,'date':date,'str_type':str_type,'id_demande':id_demande,'process':process})
    df = df.loc[(pd.isnull(df['x'])==False)][['x','y','date','str_type','id_demande','process']]
    month = 0
    monthend = 0
    year = 0
    yearend = 0
    regression = pd.DataFrame()
    for year in range( year_start, year_end + 1 ):
        #print(year)
        for month in range( 1, 13 ):
            #print(month)
            if month == 12:
                yearend = year + 1
                monthend = 1
            else:
                yearend = year
                monthend = month + 1
            if (year==year_start)&(month<month_start):
                continue
            if (year==year_end)&(month>month_end):
                continue
#    for year in range( 2019, 2021 ):
#        for month in range( 1, 13 ):
#            if month == 12:
#                yearend = year + 1
#                monthend = 1
#            else:
#                yearend = year
#                monthend = month + 1
#            if (month >= 6) & (year==2020):
#                continue
            for i, p in enumerate(process_list):
                for j, model in enumerate(['Incident','Ressource']):
            #Select x and y from a x < 2 weeks of work.
                    dtemp = df.loc[(pd.isnull(df['x'])==False)
                    &(df['process'] == p)
                    &(df['str_type'].str.contains(model))
                    &(pd.to_datetime(df['date'])>=pd.datetime(year,month,1))
                    &(pd.to_datetime(df['date'])<pd.datetime(yearend,monthend,1))][['x','y','process','date','str_type','id_demande']]
                    #x = df.loc[ (pd.to_datetime(df['date'])>=pd.datetime(year,month,1))&(pd.to_datetime(df['date'])<pd.datetime(yearend,monthend,1))]['x']
                    #y = df.loc[ (pd.to_datetime(df['date'])>=pd.datetime(year,month,1))&(pd.to_datetime(df['date'])<pd.datetime(yearend,monthend,1))]['y']
                    if (dtemp['id_demande'].count()>=1):
                        try:
                            slope, intercept, r_value, p_value, std_err = stats.linregress(dtemp['x'], dtemp['y'])
                            Rsquared = r_value**2
                            temp=pd.DataFrame({'qut':dtemp['id_demande'].count(),'str_type':model,'process':p,'mean_resolution':x.mean(),
                                               'mean_traitement':y.mean(),'slope':slope,'intercept':intercept,
                                               'r_value':r_value,'p_value':p_value,'R-squared':Rsquared,'date':
                                               pd.datetime(year,month,1)}, index=[0])
                    
                            regression = regression.append(temp)        
                        except:
                            continue
   # regression.index = regression['date']
    regression = regression.drop('date',axis=1)
    regression.to_excel('regression_globale.xlsx')
    do.to_sql(regression,db='regression_globale')
    
    ## Étude des regressions par date par technicien et par process
    #data = pd.read_excel('Synopsis Continuum Tessera Tabular - vue panoramique des tickets.xlsx')
    data =  do.r_to_b(query='SELECT * FROM synopsis')
    data = data.loc[data['id_statut_demande']==6]
    temps_de_resolution = data["Time_of_resolution"]
    temps_de_treatment = data["treatment_time"]
    date = data['date_cloture'].dropna()
    process = data['process_label']
    str_type = data['str_type']
    id_demande = data['id_demande']
    tech = data['Tech_de_resolution']
    data['date_cloture']=data['date_cloture'].apply(pd.to_datetime)
    y = temps_de_treatment.dropna()
    x = temps_de_resolution.dropna()
    df = pd.DataFrame({'x':x,'y':y,'process':process,'tech':tech,'date':date,'str_type':str_type,'id_demande':id_demande})
    tech_list = tech.unique()
    #date_list = date.unique()
    process_list = process.dropna().unique().astype(int)
    regression = pd.DataFrame()
    for year in range( year_start, year_end + 1 ):
        #print(year)
        for month in range( 1, 13 ):
            #print(month)
            if month == 12:
                yearend = year + 1
                monthend = 1
            else:
                yearend = year
                monthend = month + 1
            if (year==year_start)&(month<month_start):
                continue
            if (year==year_end)&(month>month_end):
                continue
#    for year in range( 2019, 2021 ):
#        for month in range( 1, 13 ):
#            if month == 12:
#                yearend = year + 1
#                monthend = 1
#            else:
#                yearend = year
#                monthend = month + 1
#            if (month >= 6) & (year==2020):
#                continue
            for i, t in enumerate(tech_list):
                #print(i)
                for j, p in enumerate(process_list):
                    for k, model in enumerate(['Incident','Ressource']):
                        #print('process', str(p))
                        #print(process_list[:-1])
                        #for d in pd.datetime(2020,5,1):
                        dtemp = df.loc[(pd.isnull(df['x'])==False)&(df['tech'].str.contains(t))
                        &(df['process'] == p)
                        &(df['str_type'].str.contains(model))
                        &(pd.to_datetime(df['date'])>=pd.datetime(year,month,1))
                        &(pd.to_datetime(df['date'])<pd.datetime(yearend,monthend,1))][['x','y','tech','process','date','str_type','id_demande']]
                        #&(df['y'] <= 700)
                        #print(dtemp)
                        if (dtemp['id_demande'].count()>=1):
                            try:
                                slope, intercept, r_value, p_value, std_err = stats.linregress(dtemp['x'], dtemp['y'])
                                Rsquared = r_value**2
                                temp=pd.DataFrame({'qut':dtemp['id_demande'].count(),'mean_resolution':dtemp['x'].mean(),'mean_traitement':dtemp['y'].mean(),'slope':slope,'intercept':intercept,'r_value':r_value,'p_value':p_value,'R-squared':Rsquared,'date':\
                                                   pd.datetime(year,month,1),'process':p, 'str_type': model, 'tech':t}, index=[0])
                        
                                regression = regression.append(temp)   
                            except:
                                continue
    #                temp=pd.DataFrame({'qut':dtemp['id_demande'].count(),'mean_resolution':dtemp['x'].mean(),'mean_traitement':dtemp['y'].mean(),'slope':0,'intercept':0,'r_value':0,'p_value':0,'R-squared':0,'date':\
     #                                  pd.datetime(2020,5,1),'process':p, 'tech':t}, index=[0])
      #              regression = regression.append(temp)   
                        
                
                #print(dtemp)
    regression.index = regression['date']
    regression = regression.drop('date',axis=1)
    regression.to_excel('regression_par_tech.xlsx')
    do.to_sql(regression,db='regression_par_tech')      
        #fig = sns.lmplot(x="x", y="y", data=dtemp)

def make_graph():
    #data = do.r_to_b(query="SELECT * FROM regression")
    #sns.set(color_codes=True)
    #data = pd.read_excel('Synopsis Continuum Tessera Tabular - vue panoramique des tickets.xlsx')
    #temps_de_resolution = data["Time_of_resolution"]
    #temps_de_treatment = data["treatment_time"]
    #date = data['date_cloture'].dropna()
    #process = data['process_label']
    #tech = data['Tech_de_resolution']
    #data['date_cloture']=data['date_cloture'].apply(pd.to_datetime)
    #y = temps_de_treatment.dropna()
    #x = temps_de_resolution.dropna()
    #df = pd.DataFrame({'x':x,'y':y,'process':process,'tech':tech,'date':date})
    #tech_list = tech.unique()
   # print(tech_list)
    #for i, t in enumerate(tech_list):
        #print(t)
        #dtemp = df.loc[(pd.isnull(df['x'])==False)&(df['tech'].str.contains(t))][['x','y','tech','process','date']]
#        fig = sns.lmplot(x="x", y="y", data=dtemp)
    #fig = sns.lmplot(x="x", y="y", hue="process",
     #      col="tech", row="date", data=df);
    
    #fig.set_axis_labels('temps de résolution', 'temps de clôture')
    #fig.fig.suptitle(t);
        
        #sns.lmplot(x="date", y="y", hue="process", data=df);
    #    fig = sns_plot.get_figure()
    #plt.savefig("output.png")
        #tips = sns.load_dataset("tips")
        #print(tips)
    data = do.r_to_b(query="SELECT * FROM regression_par_tech")
    date_list = data['date'].unique()
    tech_list = data['tech'].unique()
 #   liste_de_coefficient = np.zeros((data.shape[0],data.shape[3]))
    for j, d in enumerate(date_list):
        for i, tech in enumerate(tech_list):
            for t in ['Incident','Ressource']:
                for p in [1,2,3,4]:
                    try:
                        m = data.loc[(pd.to_datetime(data['date'])==pd.to_datetime(d))
                        &(data['process']==p)
                        &(data['tech']==tech)
                        &(data['str_type']==t)]['slope'].iloc[-1]
                        c = data.loc[(pd.to_datetime(data['date'])==pd.to_datetime(d))
                        &(data['process']==p)
                        &(data['tech']==tech)
                        &(data['str_type']==t)]['intercept'].iloc[-1]
        #                liste_de_coefficient
                        x = np.array(range(0,1200))
          #              print(formula)
                        y = eval(str(c) + ' + ' + 'x*'+str(m))
                        plt.plot(x, y)
                    except:
                        continue
            plt.title(label='Process '+str(p)+' Type ' + t + ' Tech ' + tech + ' ' + str(d))
            plt.show()
#    graph(str(c) + ' + ' + 'x*'+str(m), range(0, 1200))
def graph(formula, x_range):  
    x = np.array(x_range)
    print(formula)
    y = eval(formula)
    plt.plot(x, y)  
    plt.show()

def do_supply_and_demand(month_start,month_end,year_start,year_end, calcul_integral=True, limit_r_basse=28, limit_r_moyenne=12, limit_r_haute=1.5, limit_r_tres_haute=0.5, limit_t_basse=40, limit_t_moyenne=16, limit_t_haute=2, limit_t_tres_haute=1):
## Étude des regressions par date par technicien et par process
    #data = pd.read_excel('Synopsis Continuum Tessera Tabular - vue panoramique des tickets.xlsx')
    data = do.r_to_b(query="SELECT * FROM synopsis")
    #kpi = do.r_to_b(query="SELECT * FROM helpdesk_kpi_tables_data_kpi_process_0")
    temps_de_resolution = data["Time_of_resolution"]
    temps_de_treatment = data["treatment_time"]
    date = data['date_cloture'].dropna()
    process = data['process_label']
    str_type = data['str_type']
    id_demande = data['id_demande']
    tech = data['Tech_de_resolution']
    client = data['email_demandeur']
    urgence = data['urgence']
    data['date_cloture']=data['date_cloture'].apply(pd.to_datetime)
    y = temps_de_treatment.dropna()
    x = temps_de_resolution.dropna()
    df = pd.DataFrame({'x':x,'y':y,'process':process,'tech':tech,'client':client,'date':date,'urgence':urgence,'str_type':str_type,'id_demande':id_demande})
    tech_list = tech.unique()
    client_list = client.unique()

    #date_list = date.unique()
    process_list = process.dropna().unique().astype(int)
    supply_demand_tech = pd.DataFrame()
    supply_demand_client = pd.DataFrame()
    sdtp = pd.DataFrame()
    sdcp = pd.DataFrame()

    for year in range( year_start, year_end + 1 ):
        #print(year)
        for month in range( 1, 13 ):
            #print(month)
            if month == 12:
                yearend = year + 1
                monthend = 1
            else:
                yearend = year
                monthend = month + 1
            if (year==year_start)&(month<month_start):
                continue
            if (year==year_end)&(month>month_end):
                continue
            #print( yearend, monthend)
            for i, t in enumerate(tech_list):
                #print(t)
                for j, p in enumerate(process_list):
                    #print(p)
                    for k, model in enumerate(['Incident','Ressource']):
                       # print(model)
                        #print('process', str(p))
                        #print(process_list[:-1])
                        #for d in pd.datetime(2020,5,1):
                        dtemp = df.loc[(pd.isnull(df['x'])==False)&(df['tech'].str.contains(t))
                        &(df['process'] == p)
                        &(df['str_type'].str.contains(model))
                        &(pd.to_datetime(df['date'])>=pd.datetime(year,month,1))
                        &(pd.to_datetime(df['date'])<pd.datetime(yearend,monthend,1))][['x','y','tech','urgence','process','date','str_type','id_demande']]
                        #print(dtemp['id_demande'].count()>=1)
                        ## Calcul des quantités dans chaques urgences et du taux de réussite
                        # qb : quantité basse , qtbi : quantité urgence basse in (temps de traitement réussi)
                        if (dtemp['id_demande'].count()>=1):
                            ##try:
                            qb = 0
                            qm = 0
                            qh = 0
                            qt = 0
                            qtbi = 0
                            qtmi = 0
                            qthi = 0
                            qtti = 0
                            qrbi = 0
                            qrmi = 0
                            qrhi = 0
                            qrti = 0
                            #for l, urgence in enumerate({'1':'Basse','2':'Moyenne','3':'Haute','4':'Très Haute'}):
                            qb = dtemp.loc[dtemp['urgence']==1]['id_demande'].count()
                            qm = dtemp.loc[dtemp['urgence']==2]['id_demande'].count()
                            qh = dtemp.loc[dtemp['urgence']==3]['id_demande'].count()
                            qt = dtemp.loc[dtemp['urgence']==4]['id_demande'].count()
                            qtbi = dtemp.loc[(dtemp['urgence']==1)&(dtemp['y']<limit_t_basse)]['id_demande'].count()
                            qtmi = dtemp.loc[(dtemp['urgence']==2)&(dtemp['y']<limit_t_moyenne)]['id_demande'].count()
                            qthi = dtemp.loc[(dtemp['urgence']==3)&(dtemp['y']<limit_t_haute)]['id_demande'].count()
                            qtti = dtemp.loc[(dtemp['urgence']==4)&(dtemp['y']<limit_t_tres_haute)]['id_demande'].count()
                            qrbi = dtemp.loc[(dtemp['urgence']==1)&(dtemp['x']<limit_r_basse)]['id_demande'].count()
                            qrmi = dtemp.loc[(dtemp['urgence']==2)&(dtemp['x']<limit_r_moyenne)]['id_demande'].count()
                            qrhi = dtemp.loc[(dtemp['urgence']==3)&(dtemp['x']<limit_r_haute)]['id_demande'].count()
                            qrti = dtemp.loc[(dtemp['urgence']==4)&(dtemp['x']<limit_r_tres_haute)]['id_demande'].count()
                            temp=pd.DataFrame({'qut':dtemp['id_demande'].count(),
                                               'qb':qb,'qm':qm,'qh':qh,'qt':qt,
                                               'qtbi':qtbi,'qtmi':qtmi,'qthi':qthi,'qtti':qtti,
                                               'qrbi':qrbi,'qrmi':qrmi,'qrhi':qrhi,'qrti':qrti,
                                               'mean_resolution':dtemp['x'].mean(),'mean_traitement':dtemp['y'].mean(),'date':\
                                               pd.datetime(year,month,1),'process':str(p), 'str_type': model, 'tech':t}, index=[0])
                            supply_demand_tech = supply_demand_tech.append(temp) 
                           # print(temp)
                           # except:
                           #     continue
    for year in range( year_start, year_end + 1 ):
        for month in range( 1, 13 ):
            if month == 12:
                yearend = year + 1
                monthend = 1
            else:
                yearend = year
                monthend = month + 1
            if (year==year_start)&(month<month_start):
                continue
            if (year==year_end)&(month>month_end):
                continue
            #print( yearend, monthend)
            for i, t in enumerate(client_list):
                #print(i)
                if t == '':
                    t = 'inconnu'
                for j, p in enumerate(process_list):
                    for k, model in enumerate(['Incident','Ressource']):
                        #print('process', str(p))
                        #print(process_list[:-1])
                        #for d in pd.datetime(2020,5,1):
                        dtemp = df.loc[(pd.isnull(df['x'])==False)&(df['client'].str.contains(str(t)))
                        &(df['process'] == p)
                        &(df['str_type'].str.contains(model))
                        &(pd.to_datetime(df['date'])>=pd.datetime(year,month,1))
                        &(pd.to_datetime(df['date'])<pd.datetime(yearend,monthend,1))][['x','y','client','urgence','process','date','str_type','id_demande']]
                        #&(df['y'] <= 700)
                        #print(dtemp)
                        if (dtemp['id_demande'].count()>=1):
                            #try:
                            qb = 0
                            qm = 0
                            qh = 0
                            qt = 0
                            qtbi = 0
                            qtmi = 0
                            qthi = 0
                            qtti = 0
                            qrbi = 0
                            qrmi = 0
                            qrhi = 0
                            qrti = 0
                            #for l, urgence in enumerate({'1':'Basse','2':'Moyenne','3':'Haute','4':'Très Haute'}):
                            qb = dtemp.loc[dtemp['urgence']==1]['id_demande'].count()
                            qm = dtemp.loc[dtemp['urgence']==2]['id_demande'].count()
                            qh = dtemp.loc[dtemp['urgence']==3]['id_demande'].count()
                            qt = dtemp.loc[dtemp['urgence']==4]['id_demande'].count()
                            qtbi = dtemp.loc[(dtemp['urgence']==1)&(dtemp['y']<limit_t_basse)]['id_demande'].count()
                            qtmi = dtemp.loc[(dtemp['urgence']==2)&(dtemp['y']<limit_t_moyenne)]['id_demande'].count()
                            qthi = dtemp.loc[(dtemp['urgence']==3)&(dtemp['y']<limit_t_haute)]['id_demande'].count()
                            qtti = dtemp.loc[(dtemp['urgence']==4)&(dtemp['y']<limit_t_tres_haute)]['id_demande'].count()
                            qrbi = dtemp.loc[(dtemp['urgence']==1)&(dtemp['x']<limit_r_basse)]['id_demande'].count()
                            qrmi = dtemp.loc[(dtemp['urgence']==2)&(dtemp['x']<limit_r_moyenne)]['id_demande'].count()
                            qrhi = dtemp.loc[(dtemp['urgence']==3)&(dtemp['x']<limit_r_haute)]['id_demande'].count()
                            qrti = dtemp.loc[(dtemp['urgence']==4)&(dtemp['x']<limit_r_tres_haute)]['id_demande'].count()
                            temp=pd.DataFrame({'qut':dtemp['id_demande'].count(),
                                               'qb':qb,'qm':qm,'qh':qh,'qt':qt,
                                               'qtbi':qtbi,'qtmi':qtmi,'qthi':qthi,'qtti':qtti,
                                               'qrbi':qrbi,'qrmi':qrmi,'qrhi':qrhi,'qrti':qrti,
                                               'mean_resolution':dtemp['x'].mean(),'mean_traitement':dtemp['y'].mean(),'date':\
                                               pd.datetime(year,month,1),'process':str(p), 'str_type': model, 'client':t}, index=[0])
                        
                            supply_demand_client = supply_demand_client.append(temp)   
                            #except:
                             #   continue
    supply_demand_tech
    sdtp = supply_demand_tech
    #print(sdtp)
    sdtp.index = sdtp['date']
    sdtp = sdtp.drop('date',axis=1)
    sdtp.to_excel('supply_demand_par_tech_proportion.xlsx')
    do.to_sql(sdtp,db='supply_demand_par_tech_proportion') 
    supply_demand_tech.set_index(['date', 'qut'], inplace=True)
    supply_demand_tech = supply_demand_tech.groupby(level=['date','qut']).mean()
    supply_demand_tech = supply_demand_tech.reset_index()
    supply_demand_tech.index = supply_demand_tech['date']
    supply_demand_tech.drop(['date','qb','qm','qh','qt','qtbi','qtmi','qthi','qtti','qrbi','qrmi','qrhi','qrti'],axis=1,inplace=True)
    supply_demand_tech.to_excel('supply_demand_par_tech.xlsx')
    do.to_sql(supply_demand_tech,db='supply_demand_par_tech') 
    
    sdcp = supply_demand_client
    sdcp.index = sdcp['date']
    sdcp= supply_demand_client.drop('date',axis=1)
    sdcp.to_excel('supply_demand_par_client_proportion.xlsx')
    do.to_sql(sdcp,db='supply_demand_par_client_proportion')     
    supply_demand_client.set_index(['date', 'qut'], inplace=True)
    supply_demand_client=supply_demand_client.groupby(level=['date','qut']).mean()
    supply_demand_client = supply_demand_client.reset_index()
    supply_demand_client.index = supply_demand_client['date']
    supply_demand_client.drop(['date','qb','qm','qh','qt','qtbi','qtmi','qthi','qtti','qrbi','qrmi','qrhi','qrti'],axis=1,inplace=True)
    supply_demand_client.to_excel('supply_demand_par_client.xlsx')
    do.to_sql(supply_demand_client,db='supply_demand_par_client')     
        #fig = sns.lmplot(x="x", y="y", data=dtemp)
