from odf.opendocument import load
from odf.table import Table, TableRow, TableCell, CoveredTableCell
from odf.text import P
import zipfile, os

def save_correctly(doc, filename):
    temp_file = filename + ".tmp"
    doc.save(temp_file)
    with zipfile.ZipFile(temp_file, 'r') as zin, zipfile.ZipFile(filename, 'w') as zout:
        mimetype_data = zin.read("mimetype")
        zout.writestr("mimetype", mimetype_data, compress_type=zipfile.ZIP_STORED)
        for item in zin.infolist():
            if item.filename == "mimetype":
                continue
            zout.writestr(item, zin.read(item.filename))
    os.remove(temp_file)

def get_cell_text(cell):
    ps = cell.getElementsByType(P)
    return "".join(p.firstChild.data for p in ps if p.firstChild)

def set_cell_text(cell, text):
    for child in list(cell.childNodes):
        cell.removeChild(child)
    cell.addElement(P(text=text))

liste_passage = []
def fusionner_cells(fichier_entree, table_index, row, col_start, col_end):
    if row in liste_passage :
        if row%2 == 0:
            col_start -= 1
            col_end -= 1
        else :
            col_start -= 2
            col_end -= 2
    else :
        if row%2 == 0:
            col_start -= 2
            col_end -= 2
        else :
            col_start -= 3
            col_end -= 3
        
    liste_passage.append(row)

    doc = load(fichier_entree)
    table = doc.spreadsheet.getElementsByType(Table)[table_index]
    ligne = table.getElementsByType(TableRow)[row]
    #print(ligne)
    cellules = ligne.getElementsByType(TableCell)

    # Étendre les cellules (prise en compte de numbercolumnsrepeated)
    expanded_cells = []
    for cell in cellules:
        repeat = int(cell.getAttribute("numbercolumnsrepeated") or 1)
        expanded_cells.extend([cell] * repeat)

    texte = get_cell_text(expanded_cells[col_start])
    if row%2 == 0:
        texte += " - BU"
    else :
        texte += " - B2-1"
    span = col_end - col_start + 1


    # Crée la cellule fusionnée
    merged = TableCell(numbercolumnsspanned=span)
    set_cell_text(merged, texte)

    # Reconstruit la ligne avec la fusion
    new_cells = []
    for i in range(len(expanded_cells)):
        if i == col_start:
            new_cells.append(merged)
        else:
            new_cells.append(expanded_cells[i])

    for cell in new_cells:
        ligne.addElement(cell)

    #print(ligne)

    first_cell = ligne.firstChild
    if first_cell:
        ligne.removeChild(first_cell)

    #print(ligne)
    #fichier_sortie = "./fusions/fusion_test_" + str(row) + "_" + str(col_start) + ".ods"
    #save_correctly(doc, fichier_sortie)
    #print(f"✅ Fusion de ({row}, {col_start}) à ({row}, {col_end}) enregistrée dans {fichier_sortie}")
    fichier_sortie = "tmp.ods"
    save_correctly(doc, fichier_sortie)
    #print(f"✅ Fusion de ({row}, {col_start}) à ({row}, {col_end}) enregistrée dans {fichier_sortie}")


#fusionner_cells("Semaine15 copie.ods", table_index=0, row=17, col_start=3, col_end=6, fichier_sortie="fusion_test.ods")
#fusionner_cells("Semaine15 copie.ods", table_index=0, row=16, col_start=16, col_end=18, fichier_sortie="fusion_test1.ods")
#fusionner_cells("Semaine15 copie.ods", table_index=0, row=17, col_start=13, col_end=14, fichier_sortie="fusion_test2.ods")
#fusionner_cells("Semaine15 copie.ods", table_index=0, row=8, col_start=2, col_end=4, fichier_sortie="fusion_test3.ods")
#fusionner_cells("Semaine15 copie.ods", table_index=0, row=9, col_start=0, col_end=5, fichier_sortie="fusion_test4.ods")