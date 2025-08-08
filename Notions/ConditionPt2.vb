Sub Variables()
    'ELSEIF
    'Elle permet d'ajouter plusieurs instructions a la fois

     If condition1 Then '-> Si la condition 1 est vraie ALORS
        'Instruction 1
    ElseIf condition2 Then '-> Si non la condition 2 est vraie ALORS
        'Instructions 2
    Else '-> SINON
        'Instruction 3
    End If

End Sub

Sub commentaire()
    'Variables
    Dim note As Single, commentaire As String
    note = Range("A1")

    'Commentaire en fonction de la note
    If note = 6 Then
        commentaire = "Excellent resultat !"
    ElseIf note >= 5 Then
        commentaire = "Excellent resultat !"
    ElseIf note >= 4 Then
        commentaire = "Resultat satisfaisant"
    ElseIf note >= 3 Then
        commentaire = "Resultat insatisfaisant"
    ElseIf note >= 2 Then
        commentaire = "Mauvais resultat"
    ElseIf note >= 1 Then
        commentaire = "Resultat execrable"
    Else
        commentaire = "Aucun resultat"     
    End If

    'Commentaire en B1
    Range("B1") = commentaire
    
End Sub