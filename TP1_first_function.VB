Sub randomnumber()
Dim NbrAleatoire
'Initialiser le générateur de nombres aléatoires
Randomize
NbrAleatoire = Int((1000 * Rnd) + 0)
counter = 0
For i = 0 To 10
counter = counter + 1
nn = InputBox("Entrez la valeur de x")
x = CInt(nn)
If x = NbrAleatoire Then
    MsgBox ("Bravo !")
    MsgBox (counter)
    Exit For
End If
If x > NbrAleatoire Then
    MsgBox ("less")
Else
    MsgBox ("greater")
End If
Next i
End Sub
