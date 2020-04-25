Attribute VB_Name = "Módulo1"
Sub estruturawhile()
    Dim linha As Integer
    
    linha = 2
    Cells(linha, 1).Select
    
    While ActiveCell.Value <= 100 And linha <= 2000
        linha = linha + 1
        Cells(linha, 1).Select
    Wend
    
    Range("C2").Value = linha
    Range("D2").Value = ActiveCell.Value

End Sub
