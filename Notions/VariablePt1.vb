Sub exemple()
    'Les Variables partie 1

    '=======================
    'Declaration de la variable
    Dim maVariable As Integer

    'Attribution d'une valeur a la variable
    maVariable = 12

    'Affichage de la valeur de maVariable dans une MsgBox
    MsgBox maVariable

    '=========================
    'Dim -> Declarationde la variable
    'maVariable -> nom choisi pour cette variable (
    '   sans espace
    '   ne doit pas commencer par un chiffre ou un caractere special
    ')
    'As -> declaration du type de la variable
    'Integer -> type de la variable (entier)
    '=========================

    'Declarer ses variables n'est pas obligatoire mais recommandé
    'le type de la variable indique la nature de son contenue (
    '   texte, nombre, date, etc.
    ')

    '==========================
    'MsgBox est une boite de dialogue

    '==========================
    'Exemple de chaque type de variable

    'Nombre entier
    Dim nbEntier As Integer
    nbEntier = 12345

    'nombre a virgule
    Dim nbVirgule As Single
    nbVirgule = 123.45

    'Texte
    Dim varTexte As String
    varTexte = "Excel-Pratique.com"

    'Date
    Dim VarDate As Date
    varDate = "15/05/2025"

    'Vrai/Faux
    Dim varBoolean as Boolean
    varBoolean = True

    'Objet
    Dim varFeuille As Worksheets
    Set varFeuille = Sheets("Feuil2") 'Set -> attribution d'une valeure
    'a une variable objet

    'Utilisation de la variable objet Activation de la feuille
    varFeuille.Activate

    '=============================
    'On peut utiliser les symboles pour declarer des variables

    Dim exemple As Integer
    Dim exemple%

    'ces deux lignes sont identiques

End Sub

'========================================
'Exemple pratique
'Remplisser votre classeur de Noms a la colonne A, prenoms Colonne B et Age Colonne C
Sub variable()
    'Declaration des variables
    Dim nom As String, prenom As String, age As Integer

    'Valeurs des variables
    nom = Cells(2,1)
    prenom = Cells(2,2)
    age = Cells(2, 3)

    'Boite de dialogue
    MsgBox nom & " " & prenom & ", " & age & " ans"

 
End Sub

'====================================
'effectuer l'operation en fonction d'un numero donner par l'utilisateur
Sub variable2()
    'Declaration des variables
    Dim nom As String, prenom As String, age As Integer, numeroLigne As Integer

    'Valeurs des variables
    numeroLigne = Range("F4") + 1
    nom = Cells(numeroLigne, 1)
    prenom = Cells(numeroLigne, 2)
    age = Cells(numeroLigne, 3)

    'Boite de dialogue
    MsgBox nom & " " & prenom & ", " & age & " ans"
End Sub
