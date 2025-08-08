Sub couleurs()
    'Nous avons deux possibilites pour definir la couleurs
    'ColorIndex (avec 56 couleurs) ou 
    'Color qui nous permet d'utiliser n'importe couleurs

    '======COLORINDEX======
    'Couleur du texte en A1 : vert(couleur 10)
    Range("A1").Font.ColorIndex - 10

    'Noir = 1
    'Blanc = 2
    'Rouge = 3
    'Bleu = 5
    'Jaune = 6
    'Vert = 10
    'Orange = 46

    '======COLOR======
    'Couleur du Texte en A1 : RGB(0,255,0)
    Range("A1").Font.Color = RGB(0, 255, 0)

    'RGB en francais signifie RVB (Rouge Vert Bleu)
    'Les valeurs vont de 0 a 255 pour chaque couleur

    'RGB(0,0,0) Noir
    'RGB(255,255,255) Blanc
    'RGB(255,0,0) Rouge
    'RGB(0,255,0) Vert
    'RGB(0,0,255) Bleu

    '==================================
    'Creer une bordure coloree

    'Epaisseur de la bordure
    ActiveCell.Borders.Weight = 4 'ActiveCell la cellule active
    'Couleurs de la bordure : rouge
    ActiveCell.Borders.Color = RGB(255,0,0)

    'Colorer le fond des cellules selectionees
    Selection.Interior.Color = RGB(176,242,182)

    'Colorer l'onglet de la feuille "Feuill"
    Sheets("Feuill").Tab.Color = RGB(255, 0, 0)


End Sub