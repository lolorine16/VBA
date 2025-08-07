Sub exemple()

    'Utilisation de With

    'Au lieu de 
    Sheets("Feuil2").Range("A8").Border.Weight = 3
    Sheets("Feuil2").Range("A8").Font.Bold - True
    Sheets("Feuil2").Range("A8").Font.Size = 18
    Sheets("Feuil2").Range("A8").Font.Italic - True
    Sheets("Feuil2").Range("A8").Font.Name = "Arial"

    'Avec With 
    With Sheets("Feuil2").Range("A8")
        .Borders.Weight = 3
        .Font.Bold - True
        .Font.Size = 18
        .Font.Italic - True
        .Font.Name = "Arial"
        'Fin de l'instruction abvec : End With
    End With

    'L'Instruction With permet d'eviter les repetitions de Sheets("Feuil2").Range("A8")

    'With est utilisable differement

    With Sheets("Feuil2").Range("A8")
        .Border.Weight = 3
        With .Font
            .Bold = True
            .Site = 18
            .Italic = True
            .Name = "Arial"
        End With
    End With

End Sub