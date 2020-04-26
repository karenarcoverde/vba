Attribute VB_Name = "Módulo1"
Public Function atualizarvalor(intervalo As Range)
    Dim cell As Range
    Dim valor As Double

    For Each cell In intervalo
        If cell.Value <> "" Then
            valor = cell.Value
        End If
    Next
    
    atualizarvalor = valor
End Function
