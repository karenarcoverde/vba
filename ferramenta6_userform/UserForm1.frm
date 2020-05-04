VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Registro de Compras"
   ClientHeight    =   9156.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6216
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Dim objeto As Control

    Call Cadastrar

    UserForm1.Hide
    
    For Each objeto In UserForm1.Controls
        On Error Resume Next
        objeto.Value = ""
    Next
    
End Sub
Sub Cadastrar()
    
    Dim range1 As Range
    
    If RefEdit1.Value <> "" Then
        Set range1 = Range(RefEdit1.Value)
    Else
        If Range("A2").Value = "" Then
            Set range1 = Range("A2")
        Else
            Set range1 = Range("A1").End(xlDown).Offset(1, 0)
        End If
    End If
    
    
    range1.Value = UserForm1.ComboBox1.Value
    range1.Offset(0, 1).Value = UserForm1.ListBox1.Value
    range1.Offset(0, 2).Value = UserForm1.ToggleButton1.Value
    range1.Offset(0, 3).Value = UserForm1.CheckBox1.Value
    range1.Offset(0, 4).Value = UserForm1.CheckBox2.Value
    range1.Offset(0, 5).Value = UserForm1.CheckBox3.Value
    range1.Offset(0, 6).Value = UserForm1.CheckBox4.Value
    
    
    If UserForm1.OptionButton1.Value = True Then
        range1.Offset(0, 7).Value = "Produto"
    Else
        range1.Offset(0, 7).Value = "Serviço"
    End If
    
    If UserForm1.OptionButton3.Value = True Then
        range1.Offset(0, 8).Value = UserForm1.OptionButton3.Caption
    ElseIf UserForm1.OptionButton4.Value = True Then
        range1.Offset(0, 8).Value = UserForm1.OptionButton4.Caption
    Else
        range1.Offset(0, 8).Value = UserForm1.OptionButton5.Caption
    End If
    
    
    range1.Offset(0, 9).Value = CDbl(UserForm1.TextBox2.Value)
    range1.Offset(0, 9).Style = "Currency"
    
    range1.Offset(0, 10).Value = UserForm1.TextBox1.Value
    
End Sub

Private Sub ToggleButton1_Click()
    UserForm1.Frame1.Visible = ToggleButton1.Value
End Sub

Private Sub UserForm_Initialize()

    With UserForm1.ComboBox1
        .AddItem ("Marketing")
        .AddItem ("Operações")
        .AddItem ("Financeiro")
        .AddItem ("Administrativo")
    End With
    
    UserForm1.ToggleButton1.Caption = "Nota emitida?"
    UserForm1.Frame1.Caption = "Impostos"
    UserForm1.Frame1.Visible = False
    
    UserForm1.CheckBox1.Caption = "IR"
    UserForm1.CheckBox2.Caption = "PIS"
    UserForm1.CheckBox3.Caption = "COFINS"
    UserForm1.CheckBox4.Caption = "ISS"
    
    
    UserForm1.OptionButton1.Caption = "Produto"
    UserForm1.OptionButton2.Caption = "Serviço"
    
    
    UserForm1.OptionButton3.Caption = "Antecipado"
    UserForm1.OptionButton4.Caption = "Na entrega"
    UserForm1.OptionButton5.Caption = "30 dias após a entrega"
    
    
    
    UserForm1.MultiPage1.Pages(0).Caption = "Prazo de Pagamento"
    UserForm1.MultiPage1.Pages(1).Caption = "Descrição"
    UserForm1.MultiPage1.Pages(2).Caption = "Valor"
    
    
    UserForm1.CommandButton1.Caption = "Registrar"

End Sub
