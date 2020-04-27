Attribute VB_Name = "Módulo1"
Option Explicit

Sub ajeitartabela()
    Dim valor_atual As String
    Dim range1, cell As Range
    Dim valor_numero As Double
    
    Set range1 = Range("A2:A239")
    
    For Each cell In range1
        If cell.Value <> "" Then
            valor_atual = cell.Value
        Else
            cell.Value = valor_atual
        End If
    Next
    
    
     For Each cell In range1.Offset(0, 1)
        If cell.Value <> "" Then
            valor_atual = cell.Value
        Else
            cell.Value = valor_atual
        End If
    Next
    
    
    For Each cell In range1.Offset(0, 3)
        If cell.Value <> "" Then
            valor_numero = cell.Value
        Else
            cell.Value = valor_numero
        End If
        cell.Style = "Currency"
    Next
    
    
    For Each cell In range1.Offset(0, 4)
        cell.Value = cell.Offset(0, -1).Value * cell.Offset(0, -2).Value
        cell.Style = "Currency"
    Next

End Sub
