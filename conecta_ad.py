#----------------------Import da biblioteca que mexe com o AD-----------------------#
#Para a instalação da biblioteca, deve-se ter o Python instalado na maquina#
#e utilizar o comando: python -m pip install pyad pypiwin32#
from pyad import *
#----------------------Import da biblioteca que mexe com o CSV-----------------------#
import csv

#Linha que conecta com o AD, é necessário colocar o nome completo da máquina que está o AD#
#Colocar um usuário válido que tenha permissões no AD#
con = pyad.set_defaults(ldap_server="UERGSRSAD01.UERGS.RS", username="alleff", password="uergs@2019")

#ad_object = pyad.adobject.ADObject.from_dn("CN=Alleff Dymytry Pereira de Deus, OU=Guaiba, OU=Unidades, OU=Alunos, OU=Vinculados, OU=Pessoas, DC=UERGS, DC=RS")
#pyad.adobject.ADObject.update_attribute(ad_object,"telephoneNumber","3721")

#Linha que abre o CSV específico com a codificação necessária#
with open('teste.csv', encoding="utf8") as func:
#Linha que lê as informações do arquivo CSV#
    csv_read = csv.reader(func, delimiter=',')
    line = 0
    for row in csv_read:
        if line == 0:
            print(row[0] +","+ row[1] +","+ row[2] +","+ row[3] +","+ row[4])
            line += 1
        else:
            #print(row[0])
#Linha que procura no AD a pessoa específica, pelo CN atrelado a ele, sendo esse o seu nome completo#
            ad_object = pyad.adobject.ADObject.from_cn(row[0], search_base="OU=Funcionarios,OU=Vinculados,OU=Pessoas,DC=UERGS,DC=RS")
            print(ad_object)
#Linha que modifica o objeto e adiciona as informações que estão contidas no CSV#
            #pyad.adobject.ADObject.update_attributes(ad_object,
            #                                         {"telephoneNumber": row[4],
            #                                          "title": row[1],
            #                                          "company": row[2],
            #                                          "department": row[3]})
#Linha que modifica o objeto e adiciona as informações que estão contidas no CSV#
            pyad.adobject.ADObject.update_attributes(ad_object,
                                                     {"telephoneNumber": "",
                                                      "title": "",
                                                      "company": "",
                                                      "department": ""})
            line += 1
