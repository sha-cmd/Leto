# -*- coding: utf-8 -*-
"""
Created on Wed Jul  8 10:01:52 2020

@author: rboyrie
"""
from datetime import timedelta
import pandas as pd
def datepaques(an):
    """Calcule la date de Pâques d'une année donnée an (=nombre entier)"""
    a=an//100
    b=an%100
    c=(3*(a+25))//4
    d=(3*(a+25))%4
    e=(8*(a+11))//25
    f=(5*a+b)%19
    g=(19*f+c-e)%30
    h=(f+11*g)//319
    j=(60*(5-d)+b)//4
    k=(60*(5-d)+b)%4
    m=(2*j-k-g+h)%7
    n=(g-h+m+114)//31
    p=(g-h+m+114)%31
    jour=p+1
    mois=n
    return [jour, mois, an]
 
# Exemple d'utilisation:
print (datepaques(2020))  #  affiche [23, 3, 2008]
paqq = datepaques(2020)
print(paqq[0])
paques = pd.datetime(datepaques(2020)[2],datepaques(2020)[1],datepaques(2020)[0])
pentecote = paques + timedelta(days=50)
print(pentecote)