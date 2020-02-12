# ----------------------Import da biblioteca que mexe com o AD-----------------------#
# Para a instalação da biblioteca, deve-se ter o Python instalado na maquina#
# e utilizar o comando: python -m pip install pyad pypiwin32#
from pyad import *
# ----------------------Import da biblioteca que mexe com o CSV-----------------------
import csv

import msvcrt

# Função que abre um arquivo e faz a configuração da conexão com o AD
def conecta_ad():
    flag_open = 0
    while flag_open == 0:
        try:
            arquivo = input("Digite somente o nome do arquivo de configuração:")
            arquivo = arquivo + ".txt"
            arquivo = open(arquivo, 'r')
            config = []
            for linha in arquivo:
                linha = linha.strip()
                config.append(linha)
            flag_open = 1
        except IOError:
            print("O arquivo "+str(arquivo)+" não existe.")
            flag_open = 0
    # Linha que conecta com o AD, é necessário colocar o nome completo da máquina que está o AD
    # Colocar um usuário válido que tenha permissões no AD#
    con = pyad.set_defaults(ldap_server=config[0], username=config[1], password=config[2])


# Função que abre o arquivo CSV e procura os usuário no AD
def abre_csv():
    # Linha que abre o CSV específico com a codificação necessária
    flag_open = 0
    while flag_open == 0:
        arquivo = input("Digite somente o nome do arquivo CSV:")
        arquivo = arquivo + ".csv"
        try:
            with open(arquivo, encoding="utf8") as arqcsv:
                csv_read = csv.reader(arqcsv, delimiter=",")
                line = 0
                for row in csv_read:
                    if line == 0:
                        line += 1
                    else:
                        # Linha que procura no AD a pessoa específica, pelo CN atrelado a ele, sendo esse o seu nome completo
                        ad_object = pyad.adobject.ADObject.from_cn(row[0],
                                                                search_base="OU=Funcionarios,"
                                                                            "OU=Vinculados,"
                                                                            "OU=Pessoas,"
                                                                            "DC=UERGS,"
                                                                            "DC=RS")
                        print(ad_object)
                        pyad.adobject.ADObject.update_attributes(ad_object,
                                                                {"telephoneNumber": row[4],
                                                                "title": row[1],
                                                                "company": row[2],
                                                                "department": row[3]})
                flag_open = 1

        except IOError:
            print("O arquivo "+str(arquivo)+" não existe.")
            flag_open = 0


def main():
    conecta_ad()
    abre_csv()
    print("Processo concluído, aperte 'ENTER' para fechar.")
    ch = msvcrt.getch()


main()
