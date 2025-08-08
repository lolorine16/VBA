Sub synthaxe1()
    'CONDITIONS PARTIE 1
    'Les conditions sont tres utils en programation, elles nous
    'serviront a effectuer des actions en fonction de criteres precis

    'La principale instruction est If Voici comment elle fonctionne
    If condition Then '-> Si condition est vraie Alors
        'Instructions si vrai
    Else '-> SINON (falcutatif)
        'Instructions si faux
    End If 'Fin de la condition

End Sub

'Exemple
Sub variables1()

    'Si la valeur entre parentheses (cellule F4) est numerique (donc
    'si la condition est vraie) alors on execute les instructions placees
    'entre "Then" et "End If"

    If isNumeric(Range("F4")) Then
        'Declaration des variables
        Dim nom As String, prenom As String, age As Integer
        Dim numeroLigne As Integer

        'Valeurs des variables
        numeroLigne = Range("F4") + 1
        nom = Cells(numeroLigne, 1)
        prenom = Cells(numeroLigne, 2)
        age = Cells(numeroLigne, 3)

        'Boite de dialoge
        MsgBox nom & " " & prenom & ", " & age & " ans"

    Else
        'Boite de dialogue : avertissement
        MsgBox "L'entree """ & .Range("F4") & """ n'est pas valide !"

        'Suppression du contenu de la cellule F4
        Range("F4") = ""        
    End If
        
End Sub

'Utilisation d'operateur logique et de comparaison
Sub exemple1()
    'Si F4 est numerique
    If IsNumeric(Range("F4")) Then

        Dim nom As String, prenom As String, age As Integer
        Dim numeroLigne As Integer
        numeroLigne = Range("F4") + 1

        'Si le numero est dans la bonne plage
        If numeroLigne >= 2 And numeroLigne <= 17 Then
            nom = cells(numeroLigne, 1)
            prenom = Cells(numeroLigne, 2)
            age = Cells(numeroLigne, 3)
            MsgBox nom & " " & prenom & ", " & age & " ans"
        'Si le numero est en dehors de la plage
        Else
            MsgBox "L'entree """ & Range("F4") & """ n'est pas valide !"
            Range("F4") = ""
        End If
End Sub
'========================================================
Sub synthaxe2()
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