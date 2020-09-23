Attribute VB_Name = "ModTraitement"
Public Function Hexdeconvertor(Hexvalue)
'Convertissuer Hexadecimal
B = Int(Hexvalue / 256)
c = Hexvalue - B * 256
Hexdeconvertor = Chr$(c) & Chr$(B)
End Function

Public Sub FinTraitement()

'rien pour l'instant
End Sub
Public Sub Annulert()
'rien pour l'instant
End Sub
