Attribute VB_Name = "Módulo1"
Sub macro1()
    Dim linha As Integer
    linha = 9
    While Cells(linha, 2).Value <> ""
        If Cells(linha, 3).Value > 1500 Then
            Cells(linha, 4).Value = 0.05 * Cells(linha, 3).Value
        End If
        linha = linha + 1
    Wend

End Sub
