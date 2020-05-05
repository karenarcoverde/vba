Attribute VB_Name = "Módulo1"
Sub macro1()

    Dim linha As Integer
    Dim quantidade As Double
    Dim preco As Double
    
    linha = 7
    
    Do Until Cells(linha, 2).Value = ""
    
        quantidade = WorksheetFunction.SumIfs(Range("F:F"), Range("E:E"), Cells(linha, 2).Value)
        preco = WorksheetFunction.VLookup(Cells(linha, 2).Value, Range("I7:J10"), 2, 0)
        
        Cells(linha, 3).Value = quantidade * preco
    
        linha = linha + 1
    Loop

End Sub
