Sub NettoyageEtDeplacement()

'Supprimer le contenue de la colone A
    Columns("A:A").Select
    Selection.ClearContents

'Supprimer le contenue de la colone B
    Columns("B:B").Select
    Selection.ClearContents

'Copier le contenu de la colone C dans la colone A
    Columns("C:C").Select
    Selection.Cut Destination:=Columns("A:A")

'Copier le contenu de la colone D dans la colone B
    Columns("D:D").Select
    Selection.Cut Destination:=Columns("B:B")

    
End Sub