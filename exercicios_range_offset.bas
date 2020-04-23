Attribute VB_Name = "Módulo1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("B3:G10").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
    Range("B16:B22").Select
    Selection.Copy
    Range("B16:B22").Offset(0, 1).Select
    ActiveSheet.Paste
    Range("B3").Select
    
    
End Sub
