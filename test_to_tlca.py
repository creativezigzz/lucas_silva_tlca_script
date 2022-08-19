import csv
from csv import DictReader
from datetime import datetime
import argparse
import pandas as pd

COLS = ["Select", "Date", "Comment"]
COMPETENCE = []


class MissingParameter(Exception):
    def __init__(self, message):
        Exception.__init__(self, message)


def write_in_xlsx(src, dest, data, col_names):
    """
    Crée un fichier excel contenant les données des test qu'on passée les élèves.
    Il note et coche les cases de compétences si la note du test est supérieur à 75%


    PRE:Les chemins de fichier doivent être correct, cad que les fichiers src et les données provenant de moodle
    doivent être dans le bon format.
    Il doit y avoir exactement le même nombre d'élèves sur le fichier tlca et les données provenant du csv.µ
    POST: Crée un fichier avec le nom spécifié dans la variable dest. Si le fichier existe déjà, il l'update.

    :param src: Le fichier que nous fourni tlca avec le nom des élèves et leurs id respectif
    :param dest: Le nom du fichier dans lequel l'utilisateur veut écrire.
    :param data: Les données traduites dans un format de liste de dictionnaire provenant du fichier moodle.csv
    :param col_names: Une liste contenant les différentes colonnes où 'on va écrire,
    ainsi que les différentes compétences que l'on souhaite validé.
    :return: /
    """
    df = pd.read_excel(src)
    for col in col_names: #On écrit dans la colonne selcetionnée les datas correspondantes
        list_to_add = selected_data_to_list(data, col, col_names)
        df[col] = list_to_add
    df.to_excel(dest)


def selected_data_to_list(data, key, comp):
    """
    Crée une liste contenant les données que l'on veut écrire dans le fichier excel.

    :param: data: la liste de dictionnaire qui contient les données
    :param: key: La clé du dictionnaire qu'on choppe
    :param: comp: Une liste contenant les compétences que l'on veut évaluer.
    :return: une liste avec juste la clé particulière et le format voulu.
    :raise: Si jamais le paramètre key n'existe pas, lève une exception disant que le paramètre founrnie n'existe pas
    """
    liste = []
    if key == "Date":
        for item in data:
            liste.append(format_date(item["Started on"]))
        return liste
    elif key == "Select":
        for item in data:
            if check_result(item["Grade/10.00"]):
                liste.append("x")
            else:
                liste.append("")
        return liste
    elif key == "Comment":
        for item in data:
            if item["Grade/10.00"] != "-":
                comment = str((float(item["Grade/10.00"]) * 10).__round__(2)) + "%"
                print(comment)
                liste.append(comment)
            else:
                liste.append("")
        return liste
    elif key in comp:
        for item in data:
            if check_result(item["Grade/10.00"]):
                liste.append("x")
            else:
                liste.append("")
        return liste
    else:
        raise MissingParameter("Le paramètre fourni n'existe pas dans le fichier final")


def check_result(result):
    """
    Regarde si la note est plus grande que 75 %

    :param result: le résultat de l'étudiant au test sur /10
    :return: True ou False en fonction de la réussite de l'élève
    """
    if result == "-":
        return False
    return float(result) >= 7.5


def format_date(date_to_strip):
    """
    Traduit et renvoie une date sous le bon format
    :
    :param date_to_strip: une date sous le format 30 September 2020  12:41 PM donné par TLCA
    :return: une chaine de caractère de format :  2020-10-29 16:00
    """
    date_strip = datetime.strptime(date_to_strip, '%d %B %Y  %I:%M %p')
    date_correct_format = date_strip.strftime("%Y-%m-%d %H:%M")
    return date_correct_format


def trad(csv_file):
    """
    Traduit un fichier .csv en dictionnaire contenant les mêmes entrées
    :param csv_file: Fichier contenant les données venant de Moodle :
    :return: une liste de dictionnaires reprenant toute les entrées du fichier csv et gère les cas de doublon
    """
    list_dict = []
    with open(csv_file) as f:
        dialect = csv.Sniffer().sniff(f.read(2048), delimiters=";")
        f.seek(0)
        for line in DictReader(f, dialect=dialect):
            # Gestion des doublons test
            doublon = False
            elem_to_change = 0

            for index, dic in enumerate(list_dict):
                if line['Surname'] in dic['Surname']:
                    if line['First name'] in dic['First name']:
                        doublon = True
                        if line['Grade/10.00'] <= dic['Grade/10.00']:
                            continue
                        # On regarde les deux résultat
                        elem_to_change = index

            if not doublon:
                list_dict.append(line)
            else:
                list_dict[elem_to_change] = line
        return list_dict


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("-s", "--source", required=True,
                        help="Le chemin du fichier source venant de moodle, doit être de format *.csv")
    parser.add_argument("-t", "--tlca_in", required=True,
                        help="Le chemin fichier fourni par TLCA avec le nom de chaque élève et son id")
    parser.add_argument("-d", "--destination", required=True, help="Le chemin +nom du fichier dans lequel on écrira")
    parser.add_argument("-c", "--competence", nargs='*',
                        help="Les compétences que l'on veut validé si l'étudiant à bien réussi son test")
    args = parser.parse_args()

    COMPETENCE = args.competence
    data = trad(args.source)
    COLS.extend(COMPETENCE)

    write_in_xlsx(args.tlca_in, args.destination, data, COLS)
