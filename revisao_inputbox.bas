Attribute VB_Name = "Módulo1"
Sub macro1()
    Dim linha As Integer
    linha = Range("B1000").End(xlUp).Row + 1
    
    Cells(linha, 2).Value = InputBox("Nome do Cliente")
    Cells(linha, 3).Value = InputBox("CPF do Cliente")
    Cells(linha, 4).Value = InputBox("Telefone do Cliente")
    Cells(linha, 5).Value = InputBox("Cidade do Cliente")
    Cells(linha, 6).Value = InputBox("Produto do Cliente")

End Sub
