
' DO WHILE

Sub synthaxe()
    Do While CONDITION
        'Instruction
    Loop
End Sub

Sub exemple1()
    Dim numero As Integer

    numero = 1 'Numero de depart
    Do While numero <= 12 'Tant que la variable numero est <= 12
        'la boucle est repetee
        Cells(numero, 1) = numero 'Numerotation
        numero = numero + 1 'Le numero est augmente de 1 a chaque boucle
    Loop
End Sub

' DO LOOP

Sub synthaxe()

    Do
    'Instruction
    Loop While CONDITION

End Sub

Sub synthaxe()

    Do Until CONDITION
    'Instruction
    Loop 

End Sub


Sub synthaxe()

    Dim i As Integer

    For i = 1 To 5
    'Instruction
    Next 

End Sub
