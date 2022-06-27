# -*- coding: utf-8 -*-
"""
Created on Wed Feb 12 17:03:53 2020

@author: rboyrie
"""

def improd(id):
    if id == 1:
        return "Ralentissement modéré, l'activité est toujours possible"
    elif id == 2:
        return "Fort ralentissement, certaines actions sont impossibles"
    elif id == 3:
        return "Arrêt de la production"
    else:
        return "N'a pas d'impact précis"    
    
def oripanne(id):
    if id == 1:
        return "Logiciel"
    elif id == 2:
        return "Matériel"
    elif id == 3:
        return "Réseau"
    elif id == 4:
        return "Autre"
    else:
        return "N'a pas d'origine"
    
def nomticket(id):
    if id == 1:
        return 'incident'
    elif id == 2:
        return  'demande'
    else:
        return "Le numéro de ticket n'a pas de référence en base"
    
def nomgrp(id):
    if id == 1:
        return 'Front Office'
    if id == 2:
        return "Développement"
    if id == 3:
        return "Infrastructure"
    if id == 4:
        return "Support Applicatif"
    if id == 5:
        return "Support Matériel"
    
def statut(id):
    if id == 1:
        return 'En attente niveau 1'
    if id == 2:
        return "En atttente niveau 2"
    if id == 3:
        return "En cours de traitement niveau 1"
    if id == 4:
        return "En cours de traitement niveau 2"
    if id == 5:
        return "Résolu"
    if id == 6:
        return "Clôturé"
    
def nomact(id):
    if id == 1:
        return 'Résoudre directement'
    if id == 2:
        return "Escalader"
    if id == 3:
        return "Transférer"
    if id == 4:
        return "Indiquer En Cours de Traitement"

