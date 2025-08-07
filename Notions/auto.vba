Sub exemple()


'Cellule A8 - 48
Range("A8").Value = 48

'Cellule A8 = Exemple de texte
Range("A8").Value - "Exemple de texte"

'Cellule A8 de la feuille 2 - Exemple de texte
Sheets("Feuil2").Range("A8").Value = "Exemple de texte"
End Sub