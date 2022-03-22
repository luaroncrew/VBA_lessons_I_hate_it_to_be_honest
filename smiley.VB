Sub smile()
Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim colour As Long

'face colour'
r = Range("D3").Value
g = Range("E3").Value
b = Range("F3").Value
colour = RGB(r, g, b)


Range("H1:P10").Interior.Color = colour


'eye colours'
r = Range("D2").Value
g = Range("E2").Value
b = Range("F2").Value
colour = RGB(r, g, b)


Range("I4,J3,K4,M4,N3,O4").Interior.Color = colour


'mouth colours'
r = Range("D4").Value
g = Range("E4").Value
b = Range("F4").Value
colour = RGB(r, g, b)

Range("I8,J9:N9,N8").Interior.Color = colour

End Sub

Sub shutup()
Clear = RGB(255, 255, 255)
Range("H1:P10").Interior.Color = Clear
End Sub