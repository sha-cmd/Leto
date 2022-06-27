# -*- coding: utf-8 -*-
"""
Created on Tue Feb 11 10:49:48 2020

@author: rboyrie
"""

from machine import do
from machine import stats
import pandas as pd
class Ticket():

        def __init__(self, id):
            ##Initialisation des valeurs
            ###########################
            self.demande = id
            self.tbD = do.r_to_b(query=\
                         'SELECT * FROM `demande`  WHERE `id_demande` = ' \
                         + str(self.demande))
#            self.tbDR = do.r_to_b(query=\
#                      'SELECT * FROM `demande_ressource` WHERE `id_demande` = '\
#                      + str(self.demande))
            self.tbPEC = do.r_to_b(query=\
                       'SELECT * FROM `prise_en_charge` WHERE `id_demande` = '\
                       + str(self.demande) + '  ORDER BY `date_prise_en_charge`')
            self.tbSD = do.r_to_b(query=\
                  'SELECT * FROM `suivi_demande_ressource` WHERE `id_demande` = '\
                  + str(self.demande) + ' ORDER BY `date_suivi`')
            self.tbSI = do.r_to_b(query=\
                      'SELECT * FROM `suivi_intervention` WHERE `id_demande` = ' \
                      + str(self.demande) + ' ORDER BY `date_suivi`')
            ##Ne Fonctionne pas
            self.type = 0 if (do.r_to_b(query=\
                "SELECT `id_type_demande` FROM `demande` WHERE `id_demande` = "\
                + str(self.demande)).size == 0) else (int(do.r_to_b(query=\
                        "SELECT `id_type_demande` FROM `demande` WHERE `id_demande` = " + str(self.demande)).iloc[0]))
            ##################################################
            ##### CONSTRUCTION DES PARAMÈTRES DE L'OBJET #####
            self.bD = False
            self.bDR = False
            self.a_une_pec = False
            self.urgence = 0
            self.a_une_cloture = False
            self.est_resolu_en_suivi = False
            ####################################################
            self.t1 = pd.Timedelta('0 hours')
            self.t2 = pd.Timedelta('0 hours')
            self.t3 = pd.Timedelta('0 hours')
            self.tc = pd.Timedelta('0 hours')
            self.tr = pd.Timedelta('0 hours')
            self.trc = pd.Timedelta('0 hours')
            self.Tt = pd.Timedelta('0 hours')
            #################################################
            self.liste_date_pec = None
            self.liste_date_fin_pec = None
            
            ################################################
            self.liste_action = 0
            self.liste_statut = 0
            self.liste_id_tech_pec = 0
            self.liste_id_tech_suivi = 0
            self.liste_pec_sur_suivi = 0
            self.tech_de_cloture = 0
            self.tech_de_res = 0
            self.groupe = 0
            self.est_clos_en_suivi = False
            self.nb_demande_de_ressource = 0
            #self.est_clos_en_demande_de_ressource = False
            self.creation = None
            self.ouverture = None
            #self.a_demande_ressource = False
            ################################################
            #Calcul du nombre de PEC et de SUIVI
            ################################################
            self.nb_pec = self.tbPEC.\
                      loc[self.tbPEC['id_demande'] ==\
                          self.demande]['id_demande'].size
            self.nb_suivi_resolu = self.tbSI.loc[(self.tbSI['id_demande'] ==\
                          self.demande) & (self.tbSI['id_statut_demande'] ==\
                          5)]['id_demande'].size
            if self.nb_suivi_resolu == 0:
                self.nb_suivi_resolu = self.tbSD.loc[(self.tbSD['id_demande'] ==\
                          self.demande) & (self.tbSD['id_statut_demande'] ==\
                          5)]['id_demande'].size
            self.nb_suivi = self.tbSI.loc[self.tbSI['id_demande'] ==\
                          self.demande]['id_demande'].size 
            if self.nb_suivi == 0:
                self.nb_suivi = self.tbSD.loc[self.tbSD['id_demande'] ==\
                          self.demande]['id_demande'].size
            #################################################                                  
            #Retourne en annulant l'instanciation si la table Demande ne 
            #contient pas de ligne pour ce ticket
            #################################################
            if self.tbD.loc[self.tbD['id_demande']==self.demande]['id_demande'].size==0:
                self.leticket_existe = False
                return None
            #################################################
            #Si le ticket existe, toutes les commandes qui suivent seront
            #lu
            #################################################
            else : #Sinon le ticket existe
                self.bD = True
                self.leticket_existe = True
                #############################################
                #Calcul divers sur la table demande
                #############################################
                self.est_hors_perimetre = self.tbD.loc[self.tbD['id_demande']\
                                        == self.demande]['hors_perimetre'].iloc[0] != 0
                # La Date de création (vu qu'il existe)
                self.creation = pd.to_datetime(self.tbD.loc[self.tbD['id_demande']\
                                    ==self.demande]['date_creation_demande'].iloc[0],\
                                    format='%Y-%m-%d %H:%M:%S')
                self.a_une_creation = True if self.creation != None else False
                ####### ACTION SUR LES SUIVIS ###############
                # TESTS SUR LA DEMANDE TYPE 1 INTERVENTION###
                #############################################
                if (self.tbD.loc[self.tbD['id_demande']==self.demande]\
                ['id_type_demande'].iloc[0]==1)&(self.nb_suivi > 0):
                   #################################
                   ## Note type dans table demande##
                   #################################
                    ##########################################
                    ## Calcul le type en string et en entier##
                    ##########################################
                    self.type = 1 if self\
                        .tbD.loc[self.tbD['id_demande']==self.demande]\
                                    ['id_type_demande'].iloc[0]==1 else 2
                    self.str_type = 'intervention' if self.tbD.loc[self.tbD\
                                                    ['id_demande']==self.demande]\
                           ['id_type_demande'].iloc[0]==1 else 'demandes de res.'
                    ##########################################
                    ##########################################
                    self.liste_action = self.tbSI.loc[self.tbSI['id_demande']==\
                                  self.demande]['id_nature_action'].tolist()
                    self.liste_statut = self.tbSI.loc[self.tbSI['id_demande']==\
                                  self.demande]['id_statut_demande'].tolist()
                    self.liste_date_suivi = self.tbSI.loc[self.tbSI['id_demande']==\
                                  self.demande]['date_suivi'].tolist()
                    ############################                    
                    ##Dire s'il est résolu######
                    ############################
                    self.resolu = pd.Series(5).isin(self.liste_statut) 
                    self.est_resolu_en_suivi = True if (pd.Series(5)\
                        .isin(self.tbSI.loc[self.tbSI['id_demande']\
                                ==self.demande]['id_statut_demande']).any())else \
                                    False                    
                    self.est_escalade = True if (pd.Series(2)\
                        .isin(self.tbSI.loc[self.tbSI['id_demande']\
                                ==self.demande]['id_nature_action']).any()) else \
                                    False
                    self.est_transfere = True if (pd.Series(3)\
                        .isin(self.tbSI.loc[self.tbSI['id_demande']\
                                ==self.demande]['id_nature_action']).any())else \
                                    False
                    #####################################################                
                    # Test si le ticket a été clôturé une fois au moins##
                    #####################################################
                    self.est_clos_en_suivi = True if pd.Series(6).isin(self.tbSI.\
                                        loc[self.tbSI['id_demande']\
                                ==self.demande]['id_statut_demande']).any()else \
                                    False
                    #################################################
                    #################################################
                    #Test la vérité sur la clôture d'un ticket#######
                    #################################################
                    try:
                        self.date_de_cloture = pd.to_datetime(self.tbSI.loc[(\
                                                        self.tbSI['id_demande']\
                                                         == self.demande)\
                                         & (self.tbSI['id_statut_demande'] == 6)]\
                                ['date_suivi'].iloc[0],format='%Y-%m-%d %H:%M:%S')
                        self.tech_de_cloture = self.tbSI.loc[(\
                                                        self.tbSI['id_demande']\
                                                         == self.demande)\
                                         & (self.tbSI['id_statut_demande'] == 6)]\
                                ['id_technicien_creation'].iloc[0]
                        try:
                            self.tech_de_res = self.tbSI.loc[(\
                                                        self.tbSI['id_demande']\
                                                         == self.demande)\
                                         & (self.tbSI['id_statut_demande'] == 5)]\
                                ['id_technicien_creation'].iloc[0]
                        except:
                            #print(self.tech_de_res)
                            self.tech_de_res = self.tbSI.loc[(\
                                                        self.tbSI['id_demande']\
                                                         == self.demande)\
                                         & (self.tbSI['id_statut_demande'] == 4)]\
                                ['id_technicien_creation'].iloc[0]
                        self.a_une_cloture = self.est_clos_en_suivi
                    except:
                        self.a_une_cloture = self.est_clos_en_suivi
                    ###################################################
                    ###Renvoi la liste des techniciens de création, un#
                    # unique() pourrait donner un seul résultat       #
                    ###################################################
                    self.liste_id_tech_suivi = self.tbSI.loc[self.tbSI['id_demande']==\
                                  self.demande]['id_technicien_creation'].tolist()
                    ######################################
                    ##GROUPE##############################
                    ######################################
                    self.groupe = self.tbSI.loc[self.tbSI['id_demande']==\
                                       self.demande]['id_groupe'].iloc[0] 
                    ################################################
                    #NOMBRE DE SUIVI CALCULÉ EN PREMIER LIEU AVANT##
                    ################################################    
                    self.liste_pec_sur_suivi = self.tbSI.loc[self.tbSI['id_demande']==\
                                  self.demande]['id_prise_en_charge'].tolist()
                    ################################################
                    ##PREMIER SUIVI#################################
                    ################################################
                    self.ouverture_suivi = pd.to_datetime(self.tbSI.loc[self.tbSI['id_demande']\
                                    ==self.demande]['date_suivi'].iloc[0],\
                                        format='%Y-%m-%d %H:%M:%S')
                ##########################################    
                ## TEST SUR LA DEMANDE TYPE 2 RESSOURCE ##
                ##########################################
                elif (self.tbD.loc[self.tbD['id_demande']==self.demande]\
                ['id_type_demande'].iloc[0]==2) & (self.nb_suivi > 0):
                    self.str_type = 'demande de res.' if self.tbD.loc[self.tbD\
                                                    ['id_demande']==self.demande]\
                           ['id_type_demande'].iloc[0]==1 else 'intervention'
                    self.type = 2 if self\
                        .tbD.loc[self.tbD['id_demande']==self.demande]\
                                    ['id_type_demande'].iloc[0]==2 else 1
                    self.liste_action = self.tbSD.loc[self.tbSD['id_demande']==\
                                  self.demande]['id_nature_action'].tolist()
                    self.liste_statut = self.tbSD.loc[self.tbSD['id_demande']==\
                                  self.demande]['id_statut_demande'].tolist()
                    self.liste_date_suivi = self.tbSD.loc[self.tbSD['id_demande']==\
                                  self.demande]['date_suivi'].tolist()
                       ############################                                         
                       ##Inscrire si il est résolu#
                       ############################
                    self.est_resolu_en_suivi = True if pd.Series(5)\
                        .isin(self.tbSD.loc[self.tbSD['id_demande']\
                                ==self.demande]['id_statut_demande']).any()else \
                                    False
                        #####################################################
                        # Test si le ticket a été clôturé une fois au moins #
                        #####################################################
                    self.est_escalade = True if (pd.Series(2)\
                        .isin(self.tbSD.loc[self.tbSD['id_demande']\
                                ==self.demande]['id_nature_action']).any())else \
                                    False
                    self.est_transfere = True if (pd.Series(3)\
                        .isin(self.tbSD.loc[self.tbSD['id_demande']\
                                ==self.demande]['id_nature_action']).any())else \
                                    False
                    self.est_clos_en_suivi = True if (pd.Series(6)\
                        .isin(self.tbSD.loc[self.tbSD['id_demande']\
                                ==self.demande]['id_statut_demande']).any())else \
                                    False
                    try:
                        self.date_de_cloture = pd.to_datetime(self.tbSD.loc[(\
                                                        self.tbSD['id_demande']\
                                                         == self.demande)\
                                         & (self.tbSD['id_statut_demande'] == 6)]\
                                ['date_suivi'].iloc[0],format='%Y-%m-%d %H:%M:%S')
                        self.tech_de_cloture = self.tbSD.loc[(\
                                                        self.tbSD['id_demande']\
                                                         == self.demande)\
                                         & (self.tbSD['id_statut_demande'] == 6)]\
                                ['id_technicien'].iloc[0]
                        self.tech_de_res = self.tbSD.loc[(\
                                                        self.tbSD['id_demande']\
                                                         == self.demande)\
                                         & (self.tbSD['id_statut_demande'] == 5)]\
                                ['id_technicien'].iloc[0]
                        self.a_une_cloture = self.est_clos_en_suivi
                    except:
                        self.a_une_cloture = self.est_clos_en_suivi
                        

                    self.liste_id_tech_suivi = self.tbSD.loc[self.tbSD['id_demande']==\
                                  self.demande]['id_technicien'].tolist()
                    self.groupe = self.tbSD.loc[self.tbSD['id_demande']==\
                                       self.demande]['id_groupe'].iloc[0]
                    self.liste_pec_sur_suivi = self.tbSD.loc[self.tbSD['id_demande']==\
                                     self.demande]['id_prise_en_charge'].tolist()
                    self.ouverture_suivi = pd.to_datetime(self.tbSD.loc[self.tbSD['id_demande']\
                                    ==self.demande]['date_suivi'].iloc[0],\
                                        format='%Y-%m-%d %H:%M:%S')
                else:
                    #self.nb_suivi = 0
                    self.groupe = 0
                    self.liste_action = []
                    self.liste_id_tech_suivi = []
                    #print('Ce n\'est ni une intervention, ni une demande')
                ############################################
                #Test l'existence de la demande ressource en table
#            self.a_demande_ressource = True if self.tbDR.loc[self.\
#                                         tbDR['id_demande'] == self.demande]\
#                                             ['id_demande'].size != 0 else False
                #Test si la table demande de ressource affiche un statut actuel clos
#            if self.a_demande_ressource:
#                if self.tbDR.loc[self.tbDR['id_demande']==self.demande]\
#                            ['id_statut_actuel'].iloc[0] == 6:
#                    self.est_clos_en_demande_de_ressource = True
#                self.nb_demande_de_ressource = int(self.tbDR.\
#                      loc[self.tbDR['id_demande'] ==\
#                          self.demande]['id_demande'].size)
#                self.statut_actuel = self.tbDR.loc[self.tbDR['id_demande']==\
#                                   self.demande]['id_statut_actuel'].iloc[0]
#                        ##########################################
#            else:
#                self.statut_actuel = 0
#                self.est_clos_en_demande_de_ressource = False
#                self.nb_demande_de_ressource = 0
                    ## DÉTAIL DE LA TABLE Prise En Charge
            if self.tbPEC.loc[self.tbPEC['id_demande'] == self.demande]\
            ['id_demande'].size != 0:
                #print('Le ticket n\'a pas de prise en charge')
                self.a_une_pec = True
                self.liste_id_tech_pec = self.tbPEC.loc[self.tbPEC['id_demande']==\
                                  self.demande]['id_technicien'].tolist()
                self.urgence = int(self.tbPEC.loc[self.tbPEC['id_demande']==\
                                       self.demande]['id_urgence'].iloc[0])
                self.nb_pec = self.tbPEC.\
                      loc[self.tbPEC['id_demande'] ==\
                          self.demande]['id_demande'].size
                self.liste_date_pec = self.tbPEC.loc[self.tbPEC['id_demande']==\
                                  self.demande]['date_prise_en_charge'].tolist()
                self.liste_date_fin_pec = self.tbPEC.loc[self.tbPEC['id_demande']==\
                                  self.demande]['date_fin_prise_en_charge'].tolist()
            
            ##Définir l'indicateur t1
            if self.a_une_pec:
                self.ouverture = pd.to_datetime(self.tbPEC.loc[self.tbPEC['id_demande']\
                                    ==self.demande]['date_prise_en_charge'].iloc[0],\
                                        format='%Y-%m-%d %H:%M:%S')
                self.t1 = stats.delta(self.creation, self.ouverture)
                if self.t1 == -1:
                    print('Le ticket {} a un t1 négatif, vérifier en base'.\
                          format(self.demande))                
            self.str_process = self.de_Process()
            #self.t2_t3()
            self.kpi()
            ############################################
            ############################################
            #Définir les process
        def est_de_Process_dc1(self):
            if self is not None:
                #Test de Process 1
                if (self.nb_suivi > 0 ) and (self.type == 2) and (self.a_une_cloture) and (self.a_une_creation) and \
                            (not self.est_hors_perimetre) and (self.a_une_pec)\
                        and (not self.est_resolu_en_suivi) and (not self.est_escalade) \
                        and (not self.est_transfere):
                    return True
                else:
                    return False
        def est_de_Process_dc2(self):
            if self is not None:
                #Test de Process 1
                if (self.nb_suivi > 0 ) and (self.type == 2) and (self.a_une_cloture) and (self.a_une_creation) and \
                            (not self.est_hors_perimetre) and (self.a_une_pec)\
                        and (not self.est_resolu_en_suivi) and (self.est_escalade) \
                        and (not self.est_transfere):
                    return True
                else:
                    return False
        def est_de_Process_dc3(self):
            if self is not None:
                #Test de Process 1
                if (self.nb_suivi > 0 ) and (self.type == 2) and (self.a_une_cloture) and (self.a_une_creation) and \
                            (not self.est_hors_perimetre) and (self.a_une_pec)\
                        and (not self.est_resolu_en_suivi) and (self.est_escalade) \
                        and (self.est_transfere):
                    return True
                else:
                    return False
        def est_de_Process_dc4(self):
            if self is not None:
                #Test de Process 1
                if (self.nb_suivi > 0 ) and (self.type == 2) and (self.a_une_cloture) and (self.a_une_creation) and \
                            (self.est_hors_perimetre) and (self.a_une_pec)\
                        and (not self.est_resolu_en_suivi):
                    return True
                else:
                    return False
        def est_de_Process_drc1(self):
            if self is not None:
                #Test de Process 1
                if (self.nb_suivi > 0 ) and (self.type == 2) and (self.a_une_cloture) and (self.a_une_creation) and \
                            (not self.est_hors_perimetre) and (self.a_une_pec)\
                        and (self.est_resolu_en_suivi) and (not self.est_escalade) \
                        and (not self.est_transfere):
                    return True
                else:
                    return False
        def est_de_Process_drc2(self):
            if self is not None:
                #Test de Process 1
                if (self.nb_suivi > 0 ) and (self.type == 2) and (self.a_une_cloture) and (self.a_une_creation) and \
                            (not self.est_hors_perimetre) and (self.a_une_pec)\
                        and (self.est_resolu_en_suivi) and (self.est_escalade) \
                        and (not self.est_transfere):
                    return True
                else:
                    return False
        def est_de_Process_drc3(self):
            if self is not None:
                #Test de Process 1
                if (self.nb_suivi > 0 ) and (self.type == 2) and (self.a_une_cloture) and (self.a_une_creation) and \
                            (not self.est_hors_perimetre) and (self.a_une_pec)\
                        and (self.est_resolu_en_suivi) and (self.est_escalade) \
                        and (self.est_transfere):
                    return True
                else:
                    return False
        def est_de_Process_drc4(self):
            if self is not None:
                #Test de Process 1
                if (self.nb_suivi > 0 ) and (self.type == 2) and (self.a_une_cloture) and (self.a_une_creation) and \
                            (self.est_hors_perimetre) and (self.a_une_pec)\
                        and (self.est_resolu_en_suivi):
                    return True
                else:
                    return False
        def est_de_Process_ic1(self):
            if self is not None:
                #Test de Process 1
                if (self.nb_suivi > 0 ) and (self.type == 1) and (self.a_une_cloture) and (self.a_une_creation) and \
                            (not self.est_hors_perimetre) and (self.a_une_pec)\
                        and (not self.est_resolu_en_suivi) and (not self.est_escalade) \
                        and (not self.est_transfere):
                    return True
                else:
                    return False
        def est_de_Process_ic2(self):
            if self is not None:
                #Test de Process 1
                if (self.nb_suivi > 0 ) and (self.type == 1) and (self.a_une_cloture) and (self.a_une_creation) and \
                            (not self.est_hors_perimetre) and (self.a_une_pec)\
                        and (not self.est_resolu_en_suivi) and (self.est_escalade) \
                        and (not self.est_transfere):
                    return True
                else:
                    return False
        def est_de_Process_ic3(self):
            if self is not None:
                #Test de Process 1
                if (self.nb_suivi > 0 ) and (self.type == 1) and (self.a_une_cloture) and (self.a_une_creation) and \
                            (not self.est_hors_perimetre) and (self.a_une_pec)\
                        and (not self.est_resolu_en_suivi) and (self.est_escalade) \
                        and (self.est_transfere):
                    return True
                else:
                    return False
        def est_de_Process_ic4(self):
            if self is not None:
                #Test de Process 1
                if (self.nb_suivi > 0 ) and (self.type == 1) and (self.a_une_cloture) and (self.a_une_creation) and \
                            (self.est_hors_perimetre) and (self.a_une_pec)\
                        and (not self.est_resolu_en_suivi):
                    return True
                else:
                    return False
        def est_de_Process_irc1(self):
            if self is not None:
                #Test de Process 1
                if (self.nb_suivi > 0 ) and (self.type == 1) and (self.a_une_cloture) and (self.a_une_creation) and \
                            (not self.est_hors_perimetre) and (self.a_une_pec)\
                        and (self.est_resolu_en_suivi) and (not self.est_escalade) \
                        and (not self.est_transfere):
                    return True
                else:
                    return False
        def est_de_Process_irc2(self):
            if self is not None:
                #Test de Process 1
                if (self.nb_suivi > 0 ) and (self.type == 1) and (self.a_une_cloture) and (self.a_une_creation) and \
                            (not self.est_hors_perimetre) and (self.a_une_pec)\
                        and (self.est_resolu_en_suivi) and (self.est_escalade) \
                        and (not self.est_transfere):
                    return True
                else:
                    return False
        def est_de_Process_irc3(self):
            if self is not None:
                #Test de Process 1
                if (self.nb_suivi > 0 ) and (self.type == 1) and (self.a_une_cloture) and (self.a_une_creation) and \
                            (not self.est_hors_perimetre) and (self.a_une_pec)\
                        and (self.est_resolu_en_suivi) and (self.est_escalade) \
                        and (self.est_transfere):
                    return True
                else:
                    return False
        def est_de_Process_irc4(self):
            if self is not None:
                #Test de Process 1
                if (self.nb_suivi > 0 ) and (self.type == 1) and (self.a_une_cloture) and (self.a_une_creation) and \
                            (self.est_hors_perimetre) and (self.a_une_pec)\
                        and (self.est_resolu_en_suivi):
                    return True
                else:
                    return False
                
                
        def de_Process(self):
            """
            Retourne 1 si le ticket ne fait pas parti d'un process
            Définie le processus auquel appartient le ticket
            """
            if self.est_de_Process_dc1() == True:
                return 'Process d c 1'
            elif self.est_de_Process_dc2() == True:
                self.t2_t3()
                return 'Process d c 2'
            elif self.est_de_Process_dc3() == True:
                self.t2_t3()
                return 'Process d c 3'
            elif self.est_de_Process_dc4() == True:
                return 'Process d c 4'
            elif self.est_de_Process_drc1() == True:
                return 'Process d rc 1'
            elif self.est_de_Process_drc2() == True:
                self.t2_t3()
                return 'Process d rc 2'
            elif self.est_de_Process_drc3() == True:
                self.t2_t3()
                return 'Process d rc 3'
            elif self.est_de_Process_drc4() == True:
                return 'Process d rc 4'
            elif self.est_de_Process_ic1() == True:
                return 'Process i c 1'
            elif self.est_de_Process_ic2() == True:
                self.t2_t3()
                return 'Process i c 2'
            elif self.est_de_Process_ic3() == True:
                self.t2_t3()
                return 'Process i c 3'
            elif self.est_de_Process_ic4() == True:
                return 'Process i c 4'
            elif self.est_de_Process_irc1() == True:
                return 'Process i rc 1'
            elif self.est_de_Process_irc2() == True:
                self.t2_t3()
                return 'Process i rc 2'
            elif self.est_de_Process_irc3() == True:
                self.t2_t3()
                return 'Process i rc 3'
            elif self.est_de_Process_irc4() == True:
                return 'Process i rc 4'
            else:
                return 'Pas de process clair'
        
        def kpi(self):
            if ( self.a_une_cloture ):
                self.Tt = stats.delta(self.creation,self.date_de_cloture)
            if (self.est_de_Process_dc1() == True)|\
            (self.est_de_Process_dc2() == True)|\
            (self.est_de_Process_dc3() == True)|\
            self.est_de_Process_dc4() == True:
                ###############################################################
                ### LE TICKET 1184 M'EMPÊCHE DE FOURNIR LE LAST OF THE LIST ###
                ###############################################################
                self.trc = stats.delta(self.tbPEC.loc[\
                        self.tbPEC['id_prise_en_charge'] \
                        == self.tbSD.loc[self.tbSD['id_statut_demande']==6]\
                        ['id_prise_en_charge'].iloc[0]]['date_prise_en_charge'].iloc[0],
                        self.tbPEC.loc[\
                        self.tbPEC['id_prise_en_charge'] \
                        == self.tbSD.loc[self.tbSD['id_statut_demande']==6]\
                        ['id_prise_en_charge'].iloc[0]]['date_fin_prise_en_charge'].iloc[0])
                if self.trc == -1:
                    print('Le ticket {} a un temps négatif'.format(self.demande))
                ########################
                #### UNION DES TABLES ##
                ########################
            elif (self.est_de_Process_drc1() == True)\
            |(self.est_de_Process_drc2() == True)\
            |(self.est_de_Process_drc3() == True)\
            |(self.est_de_Process_drc4() == True):
                try:
                    self.tr = stats.delta(self.tbPEC.loc[\
                            self.tbPEC['id_prise_en_charge'] \
                            == self.tbSD.loc[self.tbSD['id_statut_demande']==5]\
                            ['id_prise_en_charge'].iloc[-1]]['date_prise_en_charge'].iloc[0],
                            self.tbPEC.loc[\
                            self.tbPEC['id_prise_en_charge'] \
                            == self.tbSD.loc[self.tbSD['id_statut_demande']==5]\
                            ['id_prise_en_charge'].iloc[-1]]['date_fin_prise_en_charge'].iloc[0])
                except:
                    self.tr = pd.Timedelta('0 hours')
                self.tc = stats.delta(self.tbSD.loc[self.tbSD['id_statut_demande']==5]\
                        ['date_suivi'].iloc[-1],
                        self.tbSD.loc[self.tbSD['id_statut_demande']==6]\
                        ['date_suivi'].iloc[-1])
                #if (temp == -1) or (temp == -1):
                 #   print('Le ticket {} a un temps négatif'.format(self.demande))
            #################################
            ### TRAVAIL DES INTERVENTIONS ###
            #################################
            elif (self.est_de_Process_ic1() == True)\
            |(self.est_de_Process_ic2() == True)\
            |(self.est_de_Process_ic3() == True)\
            |(self.est_de_Process_ic4() == True):
                self.trc = stats.delta(self.tbPEC.loc[\
                        self.tbPEC['id_prise_en_charge'] \
                        == self.tbSI.loc[self.tbSI['id_statut_demande']==6]\
                        ['id_prise_en_charge'].iloc[-1]]['date_prise_en_charge'].iloc[0],
                        self.tbPEC.loc[\
                        self.tbPEC['id_prise_en_charge'] \
                        == self.tbSI.loc[self.tbSI['id_statut_demande']==6]\
                        ['id_prise_en_charge'].iloc[-1]]['date_fin_prise_en_charge'].iloc[0])
                #############################################
                ## CALCUL DES RESOLUTIONS ET DES CLOTURES ###
                #############################################
            elif (self.est_de_Process_irc1() == True)\
            |(self.est_de_Process_irc2() == True)\
            |(self.est_de_Process_irc3() == True)\
            |(self.est_de_Process_irc4() == True):
                self.tr = stats.delta(self.tbPEC.loc[\
                        self.tbPEC['id_prise_en_charge'] \
                        == self.tbSI.loc[self.tbSI['id_statut_demande']==5]\
                        ['id_prise_en_charge'].iloc[-1]]['date_prise_en_charge'].iloc[0],
                        self.tbPEC.loc[\
                        self.tbPEC['id_prise_en_charge'] \
                        == self.tbSI.loc[self.tbSI['id_statut_demande']==5]\
                        ['id_prise_en_charge'].iloc[-1]]['date_fin_prise_en_charge'].iloc[0])
                self.tc = stats.delta(self.tbSI.loc[self.tbSI['id_statut_demande']==5]\
                        ['date_suivi'].iloc[-1],
                        self.tbSI.loc[self.tbSI['id_statut_demande']==6]['date_suivi'].iloc[-1])
                #if (temp == -1) or (temp == -1):
                 #   print('Le ticket {} a un temps négatif'.format(self.demande))
            else:
                return 'Pas de process clair'
        
        def t2_t3(self):
            ########################################################
            # Réparation d'un bug à cause d'une sale donnée 2048  #
            ########################################################
            a=pd.notnull(self.tbSD.loc[self.tbSD['id_nature_action']==2]\
                        ['id_prise_en_charge'])
           
            if (self.est_de_Process_dc2() == True)\
            |(self.est_de_Process_drc2() == True)|(self.est_de_Process_dc3() == True)\
            |(self.est_de_Process_drc3() == True):
                self.t2 = stats.delta(self.tbPEC.loc[\
                        self.tbPEC['id_prise_en_charge'] \
                        == self.tbSD.loc[self.tbSD['id_nature_action']==2]\
                        ['id_prise_en_charge'][a].iloc[-1]]['date_prise_en_charge'].iloc[0],
                        self.tbPEC.loc[\
                        self.tbPEC['id_prise_en_charge'] \
                        == self.tbSD.loc[self.tbSD['id_nature_action']==2]\
                        ['id_prise_en_charge'][a].iloc[-1]]['date_fin_prise_en_charge'].iloc[0])
            if (self.est_de_Process_ic3() == True)\
            |(self.est_de_Process_irc3() == True):
                self.t3 = stats.delta(self.tbPEC.loc[\
                        self.tbPEC['id_prise_en_charge'] \
                        == self.tbSI.loc[self.tbSI['id_nature_action']==3]\
                        ['id_prise_en_charge'].iloc[-1]]['date_prise_en_charge'].iloc[0],
                        self.tbPEC.loc[\
                        self.tbPEC['id_prise_en_charge'] \
                        == self.tbSI.loc[self.tbSI['id_nature_action']==3]\
                        ['id_prise_en_charge'].iloc[-1]]['date_fin_prise_en_charge'].iloc[0])
            if (self.est_de_Process_ic2() == True)\
            |(self.est_de_Process_irc2() == True)|(self.est_de_Process_ic3() == True)\
            |(self.est_de_Process_irc3() == True):
                self.t2 = stats.delta(self.tbPEC.loc[\
                        self.tbPEC['id_prise_en_charge'] \
                        == self.tbSI.loc[self.tbSI['id_nature_action']==2]\
                        ['id_prise_en_charge'].iloc[-1]]['date_prise_en_charge'].iloc[0],
                        self.tbPEC.loc[\
                        self.tbPEC['id_prise_en_charge'] \
                        == self.tbSI.loc[self.tbSI['id_nature_action']==2]\
                        ['id_prise_en_charge'].iloc[-1]]['date_fin_prise_en_charge'].iloc[0])
            if (self.est_de_Process_dc3() == True)\
            |(self.est_de_Process_drc3() == True):
                self.t3 = stats.delta(self.tbPEC.loc[\
                        self.tbPEC['id_prise_en_charge'] \
                        == self.tbSD.loc[self.tbSD['id_nature_action']==3]\
                        ['id_prise_en_charge'].iloc[-1]]['date_prise_en_charge'].iloc[0],
                        self.tbPEC.loc[\
                        self.tbPEC['id_prise_en_charge'] \
                        == self.tbSD.loc[self.tbSD['id_nature_action']==3]\
                        ['id_prise_en_charge'].iloc[-1]]['date_fin_prise_en_charge'].iloc[0])
            
            
        def __str__(self):
            print('liste des statuts {}'.format(self.liste_statut))
            print('liste des actions {}'.format(self.liste_action))
            print('liste des pec sur suivi{}'.format(self.liste_pec_sur_suivi))
            print('liste des déb pec {}'.format(self.liste_date_pec))
            print('liste des fin pec {}'.format(self.liste_date_fin_pec))
            print('liste des tech de pec{}'.format(self.liste_id_tech_pec))      
            print('liste des tech de suivi {}'.format(self.liste_id_tech_suivi))
            print('Annonce du groupe {}'.format(self.groupe))
            print('Nombre de pec {}'.format(self.nb_pec))
            print('Nombre de suivi {}'.format(self.nb_suivi))
            print('Tech de résolution {}'.format(self.tech_de_res))
            print('Tech de clôture {}'.format(self.tech_de_cloture))
           # print('A une demande ressource {}'.format(self.a_demande_ressource))
           # print('Nombre de demande {}'.format(self.nb_demande_de_ressource))
            print('Date de création {}'.format(self.creation))
           # print('Il est résolu en suivi {}'.format(self.est_résolu_en_suivi))
            print('Est hors périmètre {}'.format(self.est_hors_perimetre))
            print('{}'.format(self.ouverture))