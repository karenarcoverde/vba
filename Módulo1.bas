Attribute VB_Name = "Módulo1"
Sub operacoes()

    Dim resultado_soma, resultado_diferenca, resultado_produto, resultado_potencia, resultado_raiz As Double
    resultado_soma = Cells(2, 2).Value + Cells(3, 2).Value + Cells(4, 2).Value + Cells(5, 2).Value + Cells(6, 2).Value + Cells(7, 2).Value
    Cells(8, 2).Value = resultado_soma
    
    resultado_diferenca = Cells(2, 4).Value - Cells(3, 4).Value - Cells(4, 4).Value - Cells(5, 4).Value - Cells(6, 4).Value - Cells(7, 4).Value
    Cells(8, 4).Value = resultado_diferenca
    
    resultado_produto = Cells(2, 6).Value * Cells(3, 6).Value * Cells(4, 6).Value * Cells(5, 6).Value * Cells(6, 6).Value * Cells(7, 6).Value
    Cells(8, 6).Value = resultado_produto
    

End Sub
