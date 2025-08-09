Sub exerciceBoucle()
    Dim colonne As Integer, ligne As Integer
    'Boucle pour parcourir les lignes de 1 a 10 en remplissant les colonnes de
    '1 a 10

    For ligne = 1 To 10
        'Boucle des colonnes
        For colonne = 1 To 10
            Cells(ligne, colonne) = (ligne - 1) * 10 + colonne
            'Coloration d'une cellule sur 2
            If (ligne + colonne) Mod 2 = 0 Then 'Si le reste de la division
            'par 2 = 0
                Cells(ligne, colonne).Interior.Color = RGB(220, 220, 220)
            End If
        Next
    Next 
En1S10