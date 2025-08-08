Sub exemple()
    '=========================
    '======LES TABLEAUX=======
    '=========================
    'Les variables permettent de stocker une seule valeure par variable
    'Les tableaux permettent de stocker une multitude de valeurs par tableau
    'Leurs utilisations est proche de celle des variables
    '=========================

    'tableau a une dimension=====================
    'Exemple de declaration de tableau a 1 dimension
    Dim tab1(4) As String 'tableau a une dimension

    ' 4 -> indique le nombre de cases dans le tableau (0 a 4) donc 5 cases
    ' --> tab1(4) est un tableau dont les cases vont de 0 a 4

    'Exemple d'enregistrement de valeurs dans le tableau
    tab1(0) = "Valeur de la case 0"
    tab1(1) = "Valeur de la case 1"
    tab1(2) = "Valeur de la case 2"
    tab1(3) = "Valeur de la case 3"
    tab1(4) = "Valeur de la case 4"

    '===========================
    'Afficher le tableau dans la colonne A (a partir de la ligne 1)
    Dim i As String
    For i=1 To 5
        Cells(i, 1).Value = tab1(i - 1)
    Next

    'tableau a 2 dimension======================
    'Exemple de declaration de tableau a 2 dimension
    Dim tab2(4, 3) As String
    'Exemple d'enregistrement de valeurs dans le tableau a deux dimensions
    tab2(0,0) = "Valeur de la case rouge"
    tab2(4,1) = "Valeur de la case verte"
    tab2(2,3) = "Valeur de la case bleue"

    '===========================
    '======LES CONSTANTES=======
    '===========================
    'Les constantes permettent de stocker des valeurs comme les variables
    'a la difference pres qu'on ne peut pas les modifier (d'ou leur nom)
    'apres les avoir declarees
    '==========================

    'Declaration de la constante + attribution de sa valeur
    Const TAUX_TVA As Double = 0.1234

    Cells(1, 1) = Cells(1, 2) * TAUX_TVA
    Cells(2, 1) = Cells(2, 2) * TAUX_TVA
    Cells(3, 1) = Cells(3, 2) * TAUX_TVA
    Cells(4, 1) = Cells(4, 2) * TAUX_TVA
    Cells(5, 1) = Cells(5, 2) * TAUX_TVA

End Sub
'=============================================
'La portee d'une variable

'Si la variable est declaree au debut d'une procedure (Sub), elle ne peut etre
'utiliser que dans cette meme procedure
'La valeure de la variable n'est pas conservee apres l'execution de la procedure

Sub procedure1()
    Dim var1 As Integer
    '-> Utilisation de la variable dans la procedure uniquement
End Sub

Sub procedure2()
    '-> Impossible d'utiliser var1 ici
End Sub

'Pour pouvoir utiliser une variable dans toutes les procedures d'un module, il suffit
'de la declarer en debut de module. De plus, cela permet de conserver la valeure
'de la variable jusqu'a la fermeture du classeur

Dim var2 As Integer

Sub procedure3()
    '-> Utilisation de var2 possible
End Sub

Sub procedure4()
    '-> Utilisation de var2 possible
End Sub

'==============================================
'Meme principe pour utiliser une variable dans tous les modules, a la difference
'pres que Dim est remplace par Public

Public var3 As Integer

'Pour conserver la valeur d'une variable a la fin d'une procedure, remplacer Dim par
'Static
Sub procedure5()
    Static var4 As Integer
End Sub

'Pour conserver les valeurs de toutes les variables d'une procedure, ajouter
'Static devant sub
Static Sub procedure6()
    Dim var5 As Integer
End Sub

'===========================================
'=====CREER SON PROPRE TYPE DE VARIABLE=====
'===========================================

'Creation d'un type de variable
Type Utilisateur
    Nom As String
    Prenom As String
End Type

Sub utilisationTypeVariables()
    'Declaration
    Dim user1 As Utilisateur

    'Attributions des valeurs a user1
    user1.Nom = "Smith"
    user1.Prenom = "John"

    'Exemple d'utilisation
    MsgBox user1.Nom & " " & user1.Prenom 
End Sub