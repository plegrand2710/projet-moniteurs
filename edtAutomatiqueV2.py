import datetime
import csv
import math
import ezodf
from fusiontest3 import fusionner_cells
from colorama import Fore, Style

def conception_edt(fichier_csv, fichier_sortie):
    moniteurs_heures = {"Anna" : 0, "Heifara" : 0, "Lennon" : 0, "Pauline" : 0}
    # ==== 1. Ouvrir le document ODS ====
    doc = ezodf.opendoc("TEMPLATE.ods")
    sheet = doc.sheets[0]

    # ==== 2. Calcul des dates de la semaine prochaine ====
    today = datetime.date.today()
    days_until_next_monday = (7 - today.weekday()) or 7
    next_monday = today + datetime.timedelta(days=days_until_next_monday)
    # On considère la semaine du lundi au samedi (6 jours) comme dans votre code
    week_dates = [next_monday + datetime.timedelta(days=i) for i in range(6)]
    monday = week_dates[0]
    sunday = week_dates[-1]

    # ==== 3. Mise à jour de la cellule de période ====
    sheet[4, 7].set_value(
        f"Semaine : du {monday.strftime('%d/%m/%Y')} au {sunday.strftime('%d/%m/%Y')}"
    )

    # ==== 4. Écriture des jours de la semaine dans la colonne C ====
    days_of_week = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
    for i, date_val in enumerate(week_dates):
        row_index = 8 + i * 2
        col_index = 2
        sheet[row_index, col_index].set_value(f"{days_of_week[i]}\n{date_val.strftime('%d/%m/%Y')}")

    # ==== 5. Fonction de conversion heure/date vers cellule ====
    def convert_to_cell(dt, cell_type):
        """
        Convertit une date-heure (dt) et un type ("BU" ou "B2-1") en coordonnées (ligne, colonne).
        On utilise 7h30 comme référence pour la colonne 3.
        Chaque tranche de 30 minutes ajoute 1 colonne.
        Le lundi correspond à la ligne 8 pour BU et à la ligne 9 pour B2-1,
        et chaque jour ajoute 2 lignes.
        """
        day_index = dt.weekday()  # lundi = 0, ..., dimanche = 6
        base_row = 8 if cell_type.upper() == "BU" else 9
        row = base_row + (2 * day_index)
        
        base_minutes = 7 * 60 + 30  # 7h30 en minutes
        current_minutes = dt.hour * 60 + dt.minute
        minutes_offset = current_minutes - base_minutes
        col_offset = minutes_offset // 30
        col = 3 + col_offset
        
        return row, col

    # ==== 6. Lecture du CSV et écriture dans les cellules concernées ====
    #csv_path = "Evenements_Travail_Semaine_Prochaine16 (2).csv"

    csv_path = fichier_csv

    print("\n📅 Événements et cellules correspondantes :")
    with open(csv_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile, delimiter=",")
        for row_data in reader:
            debut_str = row_data["Début"]
            fin_str = row_data["Fin"]
            lieu = row_data["Lieu"].strip() if row_data["Lieu"] else ""
            nom = row_data["Nom"]

            # Détermination du type de cellule selon le lieu
            cell_type = "BU" if "bu" in lieu.lower() else "B2-1"

            try:
                dt_start = datetime.datetime.strptime(debut_str, "%Y-%m-%d %H:%M:%S")
                dt_end = datetime.datetime.strptime(fin_str, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                print(f"⚠️ Format de date non reconnu : {debut_str} ou {fin_str}")
                continue

            # Calcul du nombre de cellules nécessaires
            duration_minutes = (dt_end - dt_start).total_seconds() / 60
            num_cells = math.ceil(duration_minutes / 30)
            if num_cells < 1:
                num_cells = 1

            # Calcul de la cellule de départ (coin supérieur gauche)
            row_cell, col_cell = convert_to_cell(dt_start, cell_type)
            print(f"🗂 {nom} – Début : {debut_str} ({lieu or 'B2-1'}) → Cellule ({row_cell}, {col_cell}) sur {num_cells} case(s)")
            if "Anna" in moniteurs_heures:
                moniteurs_heures[nom] += num_cells/2
            else:
                print("Nom non trouvé")

            #fichier_sortie = "./edt/avant_ecriture_" + str(row_cell) + "_" + str(col_cell) + ".ods"
            #doc.saveas(fichier_sortie)
            #print(f"sauvegarder dans {fichier_sortie}")
            # Pour chaque cellule de la plage calculée, écrire la valeur de l'événement
            for offset in range(num_cells):
                # La cellule cible pour l'offset donné
                current_cell = sheet[row_cell, col_cell + offset]
                # Si la cellule contient déjà une valeur, on ajoute une nouvelle ligne
                existing = current_cell.value
                if existing is None or existing.strip() == "":
                    new_val = nom
                else:
                    new_val = existing + "\n" + nom
                current_cell.set_value(new_val)
            
            #fichier_sortie = "./edt/apres_ecriture_" + str(row_cell) + "_" + str(col_cell) + ".ods"
            #doc.saveas(fichier_sortie)
            #print(f"sauvegarder dans {fichier_sortie}")
            doc.saveas("tmp.ods")
            fusionner_cells("tmp.ods", table_index=0, row=row_cell, col_start=col_cell, col_end=col_cell+num_cells-1)
            doc = ezodf.opendoc("tmp.ods")
            sheet = doc.sheets[0]
            

    # ==== 7. Enregistrement du document ODS ====
    doc.saveas(fichier_sortie)
    print(f"Fichier ODS créé : {fichier_sortie}")
    print(moniteurs_heures)
    print(f"\033[1mEDT pour la semaine du {monday.strftime('%d/%m/%Y')} au {sunday.strftime('%d/%m/%Y')}\033[0m")
    print(f"Bonsoir,\n\nVoici notre proposition d\'EDT pour la semaine du {monday.strftime('%d/%m/%Y')} au {sunday.strftime('%d/%m/%Y')}.\n\n\tTotal d'heures :")
    for nom, heures in moniteurs_heures.items():
        h = int(heures)
        m = int((heures - h) * 60)
        print(f"\t- {nom} : {h}h{m:02d} ;")

    print("\nCordialement,")
    print(Fore.BLUE + "\033[1mLEGRAND Pauline")
    print("Étudiante L3 informatique - TP1")
    print("Monitrice DSI 2024-2025\033[0m" + Style.RESET_ALL)

            
conception_edt("Evenements_Travail_Semaine_Prochaine16 (2).csv", "Semaine16 copie.ods")