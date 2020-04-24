Attribute VB_Name = "Módulo1"
Sub macro1()
    Range("b1").Select
    With Selection
        .Interior.Color = 65535
        .Value = "outro texto"
        .Font.Color = -16776961
        .Font.Bold = True
        .Font.Italic = True
    End With
End Sub
