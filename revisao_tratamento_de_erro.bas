Attribute VB_Name = "Módulo1"
Sub macro1()

    Dim linha As Integer
    Dim percentual As Double
    
    linha = 5
    
    On Error GoTo proximo
    Do Until Cells(linha, 2).Value = ""
    
        percentual = Cells(linha, 3).Value / (Cells(linha, 3).Value + Cells(linha, 4).Value + Cells(linha, 5).Value)
        Cells(linha, 6).Value = percentual
        
proximo:
        linha = linha + 1
        
    Loop


End Sub
