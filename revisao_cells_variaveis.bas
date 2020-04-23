Attribute VB_Name = "Módulo1"
Sub macro1()

    Dim valor_total As Double
    Dim percentual As Double
    
    valor_total = Cells(4, 3).Value * Cells(4, 4).Value
    Cells(4, 5).Value = valor_total

    valor_total = Cells(5, 3).Value * Cells(5, 4).Value
    Cells(5, 5).Value = valor_total
    
    percentual = Cells(4, 5).Value / (Cells(4, 5).Value + Cells(5, 5).Value)
    Cells(8, 2).Value = percentual
    
End Sub
