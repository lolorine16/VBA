Sub exemple()
    'Selection d'une cellule
    Range("AB").Select

    'Activation de la feuille 2
    Sheets("Feuil2").Activate

    'Selection de la cellule A8
    Range("A8").Select

    '============================

    'Selection des cellules A1 a A8
    Range("A1:A8").Select

    'Selection des Cellules A8 et C5
    Range("A8, C5").Select

    'Selection des cellules de la plage "ma_plage"
    Range("ma_plage").Select


    '=============================

    'Selection de la cellule de la ligne 8 et de la colonne 1
    Cells(8, 1).Select

    '=============================

    'Selection aleatoire d'une  cellule de la ligne 1 a 10 et de la colonne 1
    Cells(int(Rnd * 10) + 1, 1).Select

    'Traduction:
    'Cells([nombre_aleatoire_entre_1_et_10], 1).Select

End Sub