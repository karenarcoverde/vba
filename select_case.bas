Attribute VB_Name = "Módulo1"
Sub estruturaSelectCase()

    Dim variavel As String
    Dim range1 As Range
    Dim resultado As Double
    
    
    Set range1 = Range("A4:E4")
    variavel = Range("B1").Value
    
    Select Case variavel
        Case "Soma"
            For Each cell In range1
                resultado = resultado + cell.Value
            Next
            
        Case "Diferença"
            For Each cell In range1
                If cell.Column = 1 Then
                    resultado = cell.Value
                Else
                resultado = resultado - cell.Value
                End If
            Next
           
        Case "Multiplicação"
            For Each cell In range1
                If cell.Column = 1 Then
                    resultado = cell.Value
                Else
                resultado = resultado * cell.Value
                End If
            Next
            
        Case "Divisão"
            For Each cell In range1
                If cell.Column = 1 Then
                    resultado = cell.Value
                Else
                    If cell.Value <> "" Then
                        resultado = resultado / cell.Value
                    End If
                End If
            Next
        Case Else
            MsgBox ("Escolha uma operação para realizar")
    End Select
    
    Range("B6").Value = resultado
    
    
End Sub
