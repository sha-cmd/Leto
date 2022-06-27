# -*- coding: utf-8 -*-
"""
Created on Tue Feb 18 13:12:28 2020
https://github.com/AntoineAugusti/jours-feries-france-datagouv
@author: Antoine Augusti
"""

"""
Mettre à jour la Pentecôte en l’excluant des jours fériés.
"""
from datetime import date
from datetime import timedelta

import pandas as pd
from ics import Calendar, Event
from jours_feries_france.compute import JoursFeries


def bank_holiday_name(data, date):
    return [k for k, v in data[date.year].items() if v == date][0]


def to_csv(df, filename):
    df.to_csv(filename, index=False, encoding="utf-8")


def add_event(calendar, name, date):
    event = Event()
    event.name = name
    event.begin = date.strftime("%Y-%m-%d")
    event.make_all_day()
    calendar.events.add(event)

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
    paques = pd.datetime(an,mois,jour)
    pentecote = paques + timedelta(days=50)# Jour de solidarité DSI
    ascencion = paques + timedelta(days=39)# Si besoin dans le futur
    return pentecote

START, END = date(1950, 1, 1), date(2050, 12, 31)

modes = [({"include_alsace": False}, ""), ({"include_alsace": True}, "_alsace_moselle")]

for mode in modes:
    args, suffix = mode
    bank_holidays = {}
    calendar = Calendar()

    for year in range(START.year, END.year + 1):
        bank_holidays[year] = JoursFeries.for_year(year, **args)

    data = []
    dates = map(lambda d: d.to_pydatetime().date(), pd.date_range(START, END))
    print(dates)
    for the_date in dates:
        est_jour_ferie = the_date in bank_holidays[the_date.year].values()
        nom_jour_ferie = None
        current_year = date.today().year
        is_recent = the_date.year in range(current_year - 5, current_year + 11)

        if est_jour_ferie:
            nom_jour_ferie = bank_holiday_name(bank_holidays, the_date)

        # Generate ICS calendar only from 5 years ago to 10 years in the future
        if est_jour_ferie and is_recent:
            add_event(calendar, nom_jour_ferie, the_date)

        data.append(
            {
                "date": the_date.strftime("%Y-%m-%d"),
                "est_jour_ferie": est_jour_ferie,
                "nom_jour_ferie": nom_jour_ferie,
            }
        )

    df = pd.DataFrame(data)
    print(df)
    #to_csv(df, "jours_feries" + suffix + ".csv")
    #to_csv(df[df.est_jour_ferie], "jours_feries_seuls" + suffix + ".csv")

    #with open("jours_feries" + suffix + ".ics", "w") as my_file:
     #   my_file.writelines(calendar)