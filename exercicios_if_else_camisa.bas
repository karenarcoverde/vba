Attribute VB_Name = "Módulo1"
Sub exercicio()
    
    Dim tamanho As String
    
    tamanho = Cells(8, 3).Value
    If tamanho = "P" Then
        Cells(8, 4).Value = 10
    Else
        If tamanho = "M" Then
            Cells(8, 4).Value = 12
        Else
            Cells(8, 4).Value = 13
        End If
    End If
    
    
    
    
    tamanho = Cells(9, 3).Value
    If tamanho = "P" Then
        Cells(9, 4).Value = 10
    Else
        If tamanho = "M" Then
            Cells(9, 4).Value = 12
        Else
            Cells(9, 4).Value = 13
        End If
    End If
    
    
    
    
    
    tamanho = Cells(10, 3).Value
    If tamanho = "P" Then
        Cells(10, 4).Value = 10
    Else
        If tamanho = "M" Then
            Cells(10, 4).Value = 12
        Else
            Cells(10, 4).Value = 13
        End If
    End If
    
    
    
    
    
    tamanho = Cells(11, 3).Value
    If tamanho = "P" Then
        Cells(11, 4).Value = 10
    Else
        If tamanho = "M" Then
            Cells(11, 4).Value = 12
        Else
            Cells(11, 4).Value = 13
        End If
    End If


End Sub
