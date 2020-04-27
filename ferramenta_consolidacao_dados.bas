Attribute VB_Name = "Módulo1"
Option Explicit
Dim valor As Double
Dim doc As String
Dim data As Date
Sub AtualizarCompilado()
    Consolidar ("Conta 1")
    Consolidar ("Conta 2")
End Sub
Sub Consolidar(nome_aba As String)

    Dim range1, cell As Range
    Sheets(nome_aba).Activate
    Set range1 = Range("A1:A300")
    
    For Each cell In range1
        If AnalisarLinha(cell) Then
            Call RegistrarLinha
        End If
    Next
End Sub
Function AnalisarLinha(cell As Range) As Boolean
        If cell.Offset(0, 5).Value = "Finalizado" Then
            data = cell.Value
            doc = cell.Offset(0, 4).Value
            Call PegarValor(cell)
            AnalisarLinha = True
        Else
            AnalisarLinha = False
        End If
End Function
Sub PegarValor(cell As Range)
    If cell.Offset(0, 2).Value = "Entrada" Then
        valor = cell.Offset(0, 3).Value
    ElseIf cell.Offset(0, 2).Value = "Saída" Then
        valor = -cell.Offset(0, 3).Value
    Else
        valor = 0
        cell.Offset(0, 8).Value = "Não compilado"
    End If
End Sub
Sub RegistrarLinha()
    Dim nome_aba_conta, nome_aba_consolidacao As String
    Dim range_consolidado, cell As Range
    
    nome_aba_conta = ActiveSheet.Name
    nome_aba_consolidacao = "Consolidação de Contas"
    
    Sheets(nome_aba_consolidacao).Activate
    
    Set range_consolidado = Range("A1", Range("A1").End(xlDown)).Offset(1, 0)
   
    For Each cell In range_consolidado
        If cell.Value = "" Then
            cell.Value = data
            cell.Offset(0, 5).Value = doc
            cell.Offset(0, 4).Value = nome_aba_conta
            
            If valor < 0 Then
                cell.Offset(0, 3).Value = valor
                cell.Offset(0, 1).Value = "Saída"
            ElseIf valor > 0 Then
                cell.Offset(0, 2).Value = valor
                cell.Offset(0, 1).Value = "Entrada"
            End If
            Exit For
        End If
        If cell.Value = data Then
            If cell.Offset(0, 1).Value = "Entrada" And valor > 0 And cell.Offset(0, 4).Value = nome_aba_conta Then
                cell.Offset(0, 2).Value = valor + cell.Offset(0, 2).Value
                cell.Offset(0, 5).Value = cell.Offset(0, 5).Value & ";" & doc
                Exit For
            ElseIf cell.Offset(0, 1).Value = "Saída" And valor < 0 And cell.Offset(0, 4).Value = nome_aba_conta Then
                cell.Offset(0, 3).Value = valor + cell.Offset(0, 3).Value
                cell.Offset(0, 5).Value = cell.Offset(0, 5).Value & ";" & doc
                Exit For
            End If
        End If
    Next
    
    Sheets(nome_aba_conta).Activate
End Sub

