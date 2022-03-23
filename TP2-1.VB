
Private Sub Workbook_Open()

Randomize
NbAl = Int((1000 * Rnd) + 0)

Sheets(2).Range("A1").Value = NbAl

Compt = 0
Sheets(2).Range("B1").Value = Compt

Jeu.Show

End Sub





Private Sub CommandButton1_Click()
Dim Nbr As Integer
nbre = Int(Jeu.Proposition.Value)
rand = Sheets(2).Range("A1").Value
If rand = nbre Then
    Jeu.Resultat.Value = "Done"
ElseIf rand < nbre Then
    Jeu.Resultat.Value = "too big"
Else
    Jeu.Resultat.Value = "too little"
End If
End Sub

Private Sub CommandButton2_Click()
Unload Me

End Sub
