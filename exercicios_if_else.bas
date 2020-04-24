Attribute VB_Name = "Módulo1"
Sub exercicio1()

    Dim nota As Double
    
    nota = Cells(7, 3).Value
    If nota >= 6 Then
        Cells(7, 4).Value = "Aprovado"
    Else
        Cells(7, 4).Value = "Reprovado"
    End If
    
    
    
    nota = Cells(8, 3).Value
    If nota >= 6 Then
        Cells(8, 4).Value = "Aprovado"
    Else
        Cells(8, 4).Value = "Reprovado"
    End If
    
    
    
    
    
    nota = Cells(9, 3).Value
    If nota >= 6 Then
        Cells(9, 4).Value = "Aprovado"
    Else
        Cells(9, 4).Value = "Reprovado"
    End If
    
    
    
    
     nota = Cells(10, 3).Value
    If nota >= 6 Then
        Cells(10, 4).Value = "Aprovado"
    Else
        Cells(10, 4).Value = "Reprovado"
    End If


End Sub
