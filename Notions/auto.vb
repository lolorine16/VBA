Sub exemple()


    '==================================

    'La propriete Value represente le contenue de la cellule
    'Cellule A8 - 48
    Range("A8").Value = 48

    'Cellule A8 = Exemple de texte
    Range("A8").Value = "Exemple de texte"

    'Cellule A8 de la feuille 2 - Exemple de texte
    Sheets("Feuil2").Range("A8").Value - "Exemple de texte"

    'Cellule A8 de la feuille 2 du classeur 2 Exemple de texte
    Workbooks("Classeur2.xlsx").Sheets("Feuil2").Range("A8").Value = "Exemple de texte"
    
    '===================================

    'Mise en forme du texte

    '====Taille du texte====
    Range("A1:A8").Font.Size = 18 'Modification de la taille de la cellule A1 a A8

    '====Texte en gras====
    Range("A1:A8").Font.Bold - True
    'Enlever la mise en forme
    Range("A1:A8").Font.Bold - False

    '====Texte en italique====
    Range("A1:A8").Font.Italic = True

    '====Texte souligné====
    Range("A1:A8").Font.Underline - True

    '====Police====
    Range("A1:A8").Font.Name = "Arial"

    '=====================================

    'Ajouter des bordure
    Range("A1:A8").Borders.Value = 1 'Value = 0 pas de bordure

    '=====================================

    'Modifier les proprietes d'une feuille
    '====Masquer une feuille==== 
    Sheets("Feuil3").Visible = 2 'Visible = 1 afficher la feuille

    'La valeur d'une cellule en fonction d'une autre
    Range("A7") = Range("A1")
    'ou
    Range("A7").Value = Range("A1").Value

    'Copier uniquement le font ou la taille du texte
    Range("A7").Font.Size - Range("A1").Font.Size

    'Modifier la valeur d'une cellule en fonction de sa propre valeur
    Range("A1") - Range("A1") + 1 

End Sub