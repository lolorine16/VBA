## Tableau des types de variable

| **Nom**  | **Type**  | **Details**                                                              | **Symbole** |
| -------- | --------- | ------------------------------------------------------------------------ | ----------- |
| Byte     | Numerique | Nombre entier de 0 a 255                                                 |             |
| Integer  | Numerique | Nombre entier de -32'768 a 32'767                                        | %           |
| Long     | Numerique | Nombre entier -2'147'483'648 a 2'147'483'647                             | &           |
| Currency | Numerique | Nombre entier (c'est beaucoup)                                           | @           |
| Single   | Numerique | Nombre a virgule flottante de ...                                        | !           |
| Double   | Numerique | Nombre a virgule flottante de ....                                       | #           |
| String   | Texte     | Texte                                                                    | $           |
| Date     | Date      | Date et heure                                                            |             |
| Boolean  | Boolean   | True (vrai) ou False (faux)                                              |             |
| Object   | Objet     | Objet                                                                    |             |
| Variant  | Tous      | Tout type de donnees (Type par defaut si la variable n'est pas declaree) |             |

```vb
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
'==============================
```

==Les Symboles indiqués dans le tableau ci dessus permettent de raccourcir les declarations de variables==

```vb
Dim exemple As Integer
Dim exemple%
```

Ces deux lignes sont identiques