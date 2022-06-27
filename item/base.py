# -*- coding: utf-8 -*-
"""
Created on Wed Feb 12 16:30:35 2020

@author: rboyrie
"""
import numpy as np
import machine.do as do
import item.ticket as ticket
import machine.transcription as trn
class Base:
    def __init__(self):
        last_ticket_in_base = do.g_row_nb()
        db='helpdesk_200210'):
    #pec.helloworld()
        df = pd.DataFrame({
                'demande': [], 
                'urgence': [],
                'creation': [],
                '1ere_pec':[],
                'type':[],
                'str_type':[],
                })
        
    
        for num in range(0,last_ticket_in_base):
            tkt = ticket.Ticket(num)
         #   lst.append(ticket)
            
            if ticket.isExist:
                df = df.append({\
                                'demande': tkt.demande \
                                'urgence': tkt.urgence \
                                'groupe': tkt.groupe,\
                                'str_groupe': str(trn.nomgrp(tkt.groupe)),
                                'clos en suivi': tkt.est_clos_en_suivi,\
                                'clos en demande ressource': tkt.est_clos_en_demande_de_ressource,\
                                'resolu':tkt.isResolu,\
                                'creation':tkt.creation,\
                                'ouverture': tkt.ouverture,\
                                'derniere_resolution': ticket.lsu.dateLastResolution,\
                                'str_type': str(trn.nomticket(ticket.idtype)),
                                'type':ticket.idtype,
                                'traitement': ticket.traitement.value,
                                'attente_niv_1': ticket.delaispec.value,
                                'attente_de_cloture': ticket.lsu.delta_rc.value,
                                'impact': ticket.idimpact,
                                'str_impact': str(trn.improd(ticket.idimpact)),
                                'origine': ticket.idorigine,
                                'str_origine': str(trn.oripanne(ticket.idorigine)),
                                'nombre_de_resolution' : ticket.lsu.nombre_de_resolutions,
                                'de_niveau_2': ticket.lsu.isNiveau2,
                                'sum_pec': ticket.lpec.sum.value,
                                'traitement_s': ticket.traitement.value/1000000000,
                                'traitement_m': ticket.traitement.value/(1000000000*60),
                                'traitement_h': ticket.traitement.value/(1000000000*3600),
                                'traitement_d': ticket.traitement.value/(1000000000*60*60*24),
                                'sum_pec_s' : ticket.lpec.sum.value/1000000000,
                                'sum_pec_m' : ticket.lpec.sum.value/(1000000000*60),
                                'sum_pec_h' : ticket.lpec.sum.value/(1000000000*60*60),
                                'sum_pec_d' : ticket.lpec.sum.value/(1000000000*60*60*24),
                                }, ignore_index=True)
        
        df['demande'] = df['demande'].astype(int)
        df['urgence'] = df['urgence'].astype(int)
        df['groupe'] = df['groupe'].astype(int)
        df['str_groupe'] = df['str_groupe'].astype(str)
        df['clos'] = df['clos'].astype(bool)
        df['resolu'] = df['resolu'].astype(bool)
        df['creation'] = df['creation'].astype(str)
        df['1ere_pec'] = df['1ere_pec'].astype(str)
        df['derniere_resolution'] = df['derniere_resolution'].astype(str)
        df['str_type'] = df['str_type'].astype(str)
        df['type'] = df['type'].astype(int)
        df['traitement'] = df['traitement'].astype(float)
        df['sum_pec'] = df['sum_pec'].astype(float)
        df['sum_pec_s'] = df['sum_pec_s'].astype(float)
        df['sum_pec_m'] = df['sum_pec_m'].astype(float)
        df['sum_pec_h'] = df['sum_pec_h'].astype(float)
        df['sum_pec_d'] = df['sum_pec_d'].astype(float)
        df['traitement_s'] = df['traitement_s'].astype(float)
        df['traitement_m'] = df['traitement_m'].astype(float)
        df['traitement_h'] = df['traitement_h'].astype(float)
        df['traitement_d'] = df['traitement_d'].astype(float)
        df['attente_niv_1'] = df['attente_niv_1'].astype(float)
        df['attente_de_cloture'] = df['attente_de_cloture'].astype(float)
        df['impact'] = df['impact'].astype(int)
        df['str_impact'] = df['str_impact'].astype(str)
        df['origine'] = df['origine'].astype(int)
        df['str_origine'] = df['str_origine'].astype(str)
        df['nombre_de_resolution'] = df['nombre_de_resolution'].astype(int)
        df['de_niveau_2'] = df['de_niveau_2'].astype(bool)
        
        df.to_excel('data.xlsx')
        
        return df
        