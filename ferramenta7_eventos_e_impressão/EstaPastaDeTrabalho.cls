VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_AfterSave(ByVal Success As Boolean)

    Sheets("Impressão").PageSetup.PrintArea = "$A$1:$B$24"
    Sheets("Impressão").ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        ActiveWorkbook.Path & "\Relatório.pdf", Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        True
        
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    Workbooks.Open Filename:=ActiveWorkbook.Path & ("\Registro.xlsx")
    Workbooks("Registro.xlsx").Activate
    
    Range("A2").Value = Date
    Range("B2").Value = Now()
End Sub
