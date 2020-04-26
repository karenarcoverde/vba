Attribute VB_Name = "Módulo1"
Sub macro1()

    Dim linha As Integer
    linha = 14
    
    Do Until Cells(linha, 2).Value = ""
        Cells(linha, 6).Value = Avaliacao(Cells(linha, 3).Value, Cells(linha, 4).Value, Cells(linha, 5).Value)
        linha = linha + 1
    Loop

End Sub


Public Function Avaliacao(notap As Integer, notav As Integer, notaf As Integer) As String
    Dim media_nota As Double
    
    media_nota = (notap + notav + notaf) / 3
    
    If media_nota > 4 Then
        Avaliacao = "Excelente"
    ElseIf media_nota > 3 Then
        Avaliacao = "Muito Bom"
    ElseIf media_nota > 2 Then
        Avaliacao = "Bom"
    ElseIf media_nota > 1 Then
        Avaliacao = "Regular"
    Else
        Avaliacao = "Pode Melhorar"
    End If

End Function
