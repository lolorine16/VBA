Sub exemple()
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
Sub variables()

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
Sub exemple()
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