     # -*- coding: utf-8 -*-
"""Outils d'extraction de dossiers aléatoirement
Usage:
======
    python Extract_Aleatoire.py
"""

__authors__ = ("Pierre-Yves Léglise")
__contact__ = ("pierre-yves.leglise@cliniquesaintpaul.fr")
__copyright__ = "PYL"
__date__ = "2022-04-26"
__version__= "1.0.0"

from genericpath import exists
import pymssql
import xlsxwriter
import os
import configparser
from Cryptodome.Cipher import AES
import base64
import hashlib
from Cryptodome.Random import get_random_bytes
from beautifultable import BeautifulTable

def decrypt(enc):
    unpad = lambda s: s[:-ord(s[-1:])]
    enc = base64.b64decode(enc)
    iv = enc[:AES.block_size]
    cipher = AES.new(__key__, AES.MODE_CFB, iv)
    return unpad(base64.b64decode(cipher.decrypt(enc[AES.block_size:])).decode('utf8'))

def print_table(result):
     table=BeautifulTable()
     #table.column_headers["NIP"]
     for row in result:
          table.append_row(row)
     print(table)

def not_in_use(filename):
        try:
            os.rename(filename,filename)
            return True
        except:    
            return False

def fetchService(service):
    reqSQLServ = """select top(15) s.IdAdministratif_Sejour as NIP, format(S.DateDebut,'dd-MM-yyyy') as "Date entree", s.Nom as Patient , pi2.Nom as "Responsable", s2.Nom as Service 
    from Sejours s , Services s2 , PraticiensInternes pi2, V_SejourPratResp vspr
    where  s2.id=s.RefService 
    and vspr.RefSejour = s.id 
    and pi2.id = vspr.RefPraticienInt 
    and (S2.Nom like '%"""+service+"""%')
    and (not pi2.nom like '%POEY%' and not pi2.nom like '%GUY%' and not pi2.nom like '%ANQUI%')
    and (s.DateDebut > '"""+datedebut+"""' and s.DateDebut < '"""+datefin+"""' and format(S.DateDebut,'dd-MM-yyyy') <> format(S.DateFin,'dd-MM-yyyy'))
    and S.DateFin <> ''
    and s.IdAdministratif_Sejour <> ''
    ORDER BY NEWID(), s.DateDebut ASC"""
    cursorCSP = connectionCSP.cursor() # to access field as dictionary use cursor(as_dict=True)
    cursorCSP.execute(reqSQLServ)
    tableCSP=cursorCSP.fetchall()
    # print(reqSQLServ)
    worksheet = workbook.add_worksheet(tableCSP[0][4])
    worksheet.write_row(0, 0, ('NIP','Date entrée','Nom','Médecin','Service'))
    for ligne, data in enumerate(tableCSP):
        # print(ligne +' : ' + data)
        worksheet.write_row(ligne+1, 0, data)

def fetchAmbu(service):
    reqSQLServ = """select top(15) s.IdAdministratif_Sejour as NIP, format(S.DateDebut,'dd-MM-yyyy') as "Date entree", s.Nom as Patient , pi2.Nom as "Responsable", s2.Nom as Service 
    from Sejours s , Services s2 , PraticiensInternes pi2, V_SejourPratResp vspr
    where  s2.id=s.RefService 
    and vspr.RefSejour = s.id 
    and pi2.id = vspr.RefPraticienInt 
    and (S2.Nom like '%"""+service+"""%')
    and (not pi2.nom like '%POEY%' and not pi2.nom like '%GUY%' and not pi2.nom like '%ANQUI%')
    and (s.DateDebut > '"""+datedebut+"""' and s.DateDebut < '"""+datefin+"""' and format(S.DateDebut,'dd-MM-yyyy') = format(S.DateFin,'dd-MM-yyyy'))
    and S.DateFin <> ''
    and s.IdAdministratif_Sejour <> ''
    ORDER BY NEWID(), s.DateDebut ASC"""
    cursorCSP = connectionCSP.cursor() # to access field as dictionary use cursor(as_dict=True)
    cursorCSP.execute(reqSQLServ)
    tableCSP=cursorCSP.fetchall()
    # print(reqSQLServ)
    worksheet = workbook.add_worksheet(tableCSP[0][4])
    worksheet.write_row(0, 0, ('NIP','Date entrée','Nom','Médecin','Service'))
    for ligne, data in enumerate(tableCSP):
        # print(ligne +' : ' + data)
        worksheet.write_row(ligne+1, 0, data)

def fetchServiceMEDINTER():
    reqSQLMedInter = """select top(15) s.IdAdministratif_Sejour as NIP, format(S.DateDebut,'dd-MM-yyyy') as "Date entree", s.Nom as Patient , pi2.Nom as "Responsable", s2.Nom as Service 
    from Sejours s , Services s2 , PraticiensInternes pi2, V_SejourPratResp vspr
    where  s2.id=s.RefService 
    and vspr.RefSejour = s.id 
    and pi2.id = vspr.RefPraticienInt 
    and (S2.Nom like '%NI2%' OR S2.Nom like '%NI3%')
    and (pi2.nom like '%POEY%' or pi2.nom like '%GUY%' or pi2.nom like '%ANQUI%')
    and (s.DateDebut > '"""+datedebut+"""' and s.DateDebut < '"""+datefin+"""')
    and S.DateFin <> ''
    and s.IdAdministratif_Sejour <> ''
    ORDER BY NEWID(), s.DateDebut ASC"""
    cursorCSP = connectionCSP.cursor() # to access field as dictionary use cursor(as_dict=True)
    cursorCSP.execute(reqSQLMedInter)
    # print(reqSQL)
    worksheet = workbook.add_worksheet('MEDECINE INTERVENTIONNELLE')
    worksheet.write_row(0, 0, ('NIP','Date entrée','Nom','Médecin','Service'))
    for ligne, data in enumerate(cursorCSP):
        # print(ligne +' : ' + data)
        datalist=list(data)
        datalist[4]='MEDECINE INTERVENTIONNELLE'
        worksheet.write_row(ligne+1, 0, datalist)

def fetchUSC():
    reqSQLUSC = """select  top(15)  s.IdAdministratif_Sejour as NIP,  Format(ml.DateUtilisateur , 'dd/MM/yyyy') as Entree,  s.Nom as Patient , pi2.Nom as 'Responsable', s2.Nom as Service 
    from MouvementsLits ml ,Sejours s, Services s2,PraticiensInternes pi2, V_SejourPratResp vspr
    where ml.RefService = s2.id 
    and ml.RefSejour = s.id 
    and ml.RefService  ='F772FE6F-D0DC-4C3A-B1FB-5C17B30DB420'
    and vspr.RefSejour = s.id 
    and pi2.id = vspr.RefPraticienInt 
    and ml.DateUtilisateur >= '"""+datedebut+"""' and ml.DateUtilisateur < '"""+datefin+"""'
    and S.DateFin <> ''
    and s.IdAdministratif_Sejour <>''
    order by newid()"""
    cursorCSP = connectionCSP.cursor() # to access field as dictionary use cursor(as_dict=True)
    cursorCSP.execute(reqSQLUSC)
    #cursorCSP.insert(['NIP','Date entrée','Nom','Médecin','Service'])
    #print_table(cursorCSP)
    #exit
    worksheet = workbook.add_worksheet('USC')
    worksheet.write_row(0, 0, ('NIP','Date entrée','Nom','Médecin','Service'))
    
    for ligne, data in enumerate(cursorCSP):
        # print(ligne +' : ' + data)
        worksheet.write_row(ligne+1, 0, data)

def fetchACO(service):
    reqSQLACO = """select top(15) s.IdAdministratif_Sejour as NIP, format(S.DateDebut,'dd-MM-yyyy') as "Date entree", s.Nom as Patient , pi2.Nom as "Responsable", s2.Nom as Service 
    from Sejours s , Services s2 , PraticiensInternes pi2, V_SejourPratResp vspr
    where  s2.id=s.RefService 
    and vspr.RefSejour = s.id 
    and pi2.id = vspr.RefPraticienInt 
    and (S2.Nom like '%"""+service+"""%')
    and (s.DateFin > '"""+datedebut+"""' and s.DateFin < '"""+datefin+"""')
    and s.IdAdministratif_Sejour <> ''
    ORDER BY NEWID(), s.DateDebut ASC"""
    cursorACO = connectionACO.cursor() # to access field as dictionary use cursor(as_dict=True)
    cursorACO.execute(reqSQLACO)
    tableACO=cursorACO.fetchall()
    # print(reqSQLACO)
    worksheet = workbook.add_worksheet("ACO "+tableACO[0][4])
    worksheet.write_row(0, 0, ('NIP','Date sortie','Nom','Médecin','Service'))
    for ligne, data in enumerate(tableACO):
        # print(ligne +' : ' + data)
        worksheet.write_row(ligne+1, 0, data)

datedebut = '20220501'
datefin = '20220601'
nomfichierXl='\Extraction_Aléatoire_mai.xlsx'

__key__ = hashlib.sha256(b'pjqFX32pfaZaOkkC').digest()

scriptName=os.path.abspath(__file__)[::-1][3:][::-1]
iniName=scriptName+".ini"
# print(iniName)
if exists(iniName):
    config = configparser.ConfigParser()
    config.read(iniName)
    if "Connection" in config:
        server = decrypt(config.get("Connection","Serveur"))
        user = decrypt(config.get("Connection", "Utilisateur"))
        passwd = decrypt(config.get("Connection", "MdP"))
        dbCSP = decrypt(config.get("Connection", "BaseCSP"))
        dbACO = decrypt(config.get("Connection", "BaseACO"))
    else:
        print("Section 'Connection' introuvable dans le fichier ini")
        exit
else:
    print("Fichier ini introuvable")
    exit

# print(iniName)
# exit

fichierXL = os.path.dirname(__file__)+nomfichierXl

connectionCSP = pymssql.connect(server=server, user=user, password=passwd, database=dbCSP)
connectionACO = pymssql.connect(server=server, user=user, password=passwd, database=dbACO)

if (os.path.exists(fichierXL) and not_in_use(fichierXL)) or (not os.path.exists(fichierXL)): 
    workbook = xlsxwriter.Workbook(fichierXL,{'strings_to_numbers': True})
    for serv in ('AMBU 1','AMBU 2','AMBU 3'):
        fetchAmbu(serv)
    for serv in ('NI1','NI2','NI3','NI4','SSR CARDIO','JOUR CARDIO', 'SSR PNEUMO'):
        fetchService(serv)
    fetchUSC()
    fetchServiceMEDINTER()
    for serv in ('PSY','HOSP DE JOUR'):
        fetchACO(serv)
    workbook.close()
else: 
    print("Fichier Excel ouvert ! Fermeture du programme")
    exit

