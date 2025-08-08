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