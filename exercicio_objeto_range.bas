Attribute VB_Name = "Módulo1"
Sub exercicio()
    
    Dim faturamento As Range
    Set faturamento = Range("E5:I10")
    faturamento.Value = 5
    
    Cells(1, 1).Value = 10
    faturamento.Cells(1, 1).Value = 12
    
    faturamento.Cells(3, 2).Value = 15
    Cells(3, 2).Value = 20

End Sub
