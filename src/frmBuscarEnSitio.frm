VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBuscarEnSitio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Búsqueda Avanzada"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8940
   Icon            =   "frmBuscarEnSitio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   315
      Left            =   1620
      TabIndex        =   12
      Top             =   2520
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resultados"
      Height          =   2520
      Left            =   90
      TabIndex        =   5
      Top             =   2940
      Width           =   8775
      Begin MSComctlLib.ListView ListvResultados 
         Height          =   2235
         Left            =   60
         TabIndex        =   6
         Top             =   180
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   3942
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cedula"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cargo"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Numeración"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2205
      Left            =   1980
      TabIndex        =   0
      Top             =   60
      Width           =   4965
      Begin VB.TextBox txtCargo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2700
         TabIndex        =   11
         Top             =   1200
         Width           =   1905
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Por Cargo"
         Height          =   285
         Left            =   420
         TabIndex        =   10
         Top             =   1200
         Width           =   1725
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Por Apellidos:"
         Height          =   285
         Left            =   300
         TabIndex        =   9
         Top             =   3060
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Nombres"
         Height          =   285
         Left            =   450
         TabIndex        =   8
         Top             =   735
         Width           =   1905
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Cédula"
         Height          =   240
         Left            =   450
         TabIndex        =   7
         Top             =   330
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   1680
         Width           =   1140
      End
      Begin VB.TextBox txtApellidos 
         Enabled         =   0   'False
         Height          =   285
         Left            =   660
         TabIndex        =   3
         Top             =   2760
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.TextBox txtNombres 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2700
         TabIndex        =   2
         Top             =   750
         Width           =   1905
      End
      Begin VB.TextBox txtCedula 
         Height          =   285
         Left            =   2700
         TabIndex        =   1
         Top             =   330
         Width           =   1905
      End
   End
End
Attribute VB_Name = "frmBuscarEnSitio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lDestino
Dim lRegClientes As New ADODB.Recordset
Private Sub Form_Load()
   lDestino = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
   If Len(lDestino) = 3 Then
      lDestino = Mid(lDestino, 1, 2)
   End If
   Set lRegClientes = Nothing
   lRegClientes.Open "Select * from Clientes", Modulo.DBConexionSQL, adOpenStatic
End Sub

Private Sub llListvResultados_Click()
Dim lDst As String
Dim lTabla As String
''If lTabla <> "Diario" Then Exit Sub
If ListvResultados.ListItems.Count <= 0 Then Exit Sub
lRegClientes.MoveFirst
lDst = ListvResultados.SelectedItem.ListSubItems(2)
If lTabla = "Diario" Then
   lRegClientes.Find "Codigo=" & lDst
Else
   lDst = Mid(fPersonasAct.cCP.Text, 1, 6)
   lRegClientes.Find "Codigo=" & lDst
End If
If lRegClientes.EOF = False Then
   'lDst = lDestino & "\" & Trim(lRegClientes!Nombre) & "\" & "FOTOS"
   'lDst = lDst & "\" & Replace(ListvResultados.SelectedItem.ListSubItems(6), ".", "") & ".jpg"
   On Error GoTo falla
   'Image1.Picture = LoadPicture(lDst)
Else
   'Image1.Picture = LoadPicture()
End If
falla:
  If Err.Number <> 0 Then
   'Image1.Picture = LoadPicture()
  End If

End Sub

Private Sub ListvResultados_DblClick()
  If ListvResultados.SelectedItem.SubItems(3) <> "" And ListvResultados.SelectedItem.SubItems(3) <> "0" Then
     MsgBox "Esta persona ya ha sido numerada", vbInformation
     Exit Sub
  End If
  frmNroFotos.Show
  frmNroFotos.lblNombre.Caption = Trim(ListvResultados.SelectedItem.SubItems(1))
  frmNroFotos.txtNroFoto.Text = frmFotosEnSitio.lUltimoNumero + 1
  frmNroFotos.txtCedula.Text = Trim(ListvResultados.SelectedItem.Text)
End Sub


Private Sub ListvResultados_KeyPress(KeyAscii As Integer)
  If KeyAscii = 10 Or KeyAscii = 13 Then
   ListvResultados_DblClick
  End If
End Sub

Private Sub Option1_Click()
  If Option1.value = True Then
     txtCedula.Enabled = True
     txtNombres.Enabled = False
     txtCargo.Enabled = False
     txtCedula.SetFocus
  Else
     txtCedula.Enabled = False
  End If
End Sub


Private Sub Option2_Click()
   If Option2.value = True Then
     txtNombres.Enabled = True
     txtCedula.Enabled = False
     txtCargo.Enabled = False
     'txtNroContrato.Enabled = False
     txtNombres.SetFocus
   Else
      txtNombres.Enabled = False
   End If
End Sub

Private Sub Option3_Click()
   If Option3.value = True Then
     'txtApellidos.Enabled = True
     txtNombres.Enabled = False
     txtCedula.Enabled = False
     txtCargo.Enabled = False
     txtCargo.SetFocus
   Else
      txtCargo.Enabled = False
   End If
End Sub

'Private Sub option4_Click()
'   If Option4.Value = True Then
'     txtNroContrato.Enabled = True
'     txtApellidos.Enabled = False
'     txtNombres.Enabled = False
'     txtCedula.Enabled = False
'     txtNroContrato.SetFocus
'   Else
'      txtApartamento.Enabled = False
'   End If
'End Sub

Private Function fFormatearCedula(argCedula) As String
  Dim s As String
  s = Trim(argCedula)
  If s <> "" Then
    Formatear_Cedula s
    fFormatearCedula = s
  End If
End Function


Private Sub cmdBuscar_Click()
 Dim lReg As New ADODB.Recordset
 Dim lReg2 As New ADODB.Recordset
 Dim SqlTxt As String
 Dim lItem As ListItem
 Dim lAuxTabla As String
 Dim lCn As New ADODB.Connection
 Dim NumRegistros As Long
 Dim lReg3 As New ADODB.Recordset
 Dim lReg4 As New ADODB.Recordset
 Dim i As Integer
 Dim lTabla As String
 lTabla = fPersonasAct.lTablas.List(fPersonasAct.lTablas.ListIndex)

 SqlTxt = "Select * from [" & lTabla & "] where "
 
 If Option1.value = True Then
  If txtCedula.Text = "" Then
    MsgBox "Debe Introducir un valor para la búsqueda", vbExclamation
    Exit Sub
  End If

    SqlTxt = SqlTxt & " cedula like '%" & fFormatearCedula(txtCedula.Text) & "%'"
 End If
 
  If Option2.value = True Then
    If txtNombres.Text = "" Then
       MsgBox "Debe Introducir un valor para la búsqueda", vbExclamation
       Exit Sub
    End If
    SqlTxt = SqlTxt & " Nombre like '%" & txtNombres.Text & "%'"
    End If
 
 ListvResultados.ListItems.Clear
 SqlTxt = SqlTxt & " Order by Nombre"
 lCn.ConnectionString = Modulo.DBConexionSQL
 lCn.Open
 Set lReg = lCn.Execute(SqlTxt)

 If lReg.EOF = True Then
     MsgBox "No se encontró ningún registro con el criterio establecido", vbInformation
     Exit Sub
 End If
 Do While lReg.EOF = False
     Set lItem = ListvResultados.ListItems.Add(, , lReg!Cedula)
     lItem.SubItems(1) = Trim(lReg!Nombre)
      For i = 0 To lReg.Fields.Count - 1
         If UCase(lReg.Fields(i).Name) = "CARGO" Then
            lItem.SubItems(2) = lReg!cargo
         End If
      Next i
     lItem.SubItems(3) = IIf(IsNull(lReg!NROFOTO), "", lReg!NROFOTO)

     
     lReg.MoveNext
 Loop
 ListvResultados.SetFocus
 
 'If Option3.Value = True Then
 '   Sqltxt = Sqltxt & " Cargo like '%" & txtApellidos.Text & "%'"
 'End If

 'If Option4.Value = True Then
 '   Sqltxt = Sqltxt & " b.n_contrato like '%" & txtNroContrato.Text & "%'"
 'End If
 
End Sub

Private Sub txtApellidos_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     cmdBuscar_Click
  End If

End Sub

Private Sub txtCodigoEmpleado_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Or KeyAscii = 10 Then
     cmdBuscar_Click
  End If

End Sub

Private Sub txtApellidos_LostFocus()
   txtApellidos.Text = UCase(txtApellidos.Text)
End Sub

Private Sub txtCedula_Click()
   Option1.value = True
   Option2.value = False
   Option1_Click
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Or KeyAscii = 10 Then
     cmdBuscar_Click
  End If
End Sub

Private Sub txtCedula_LostFocus()
 txtCedula.Text = UCase(txtCedula.Text)
End Sub

Private Sub txtNombres_Click()
   Option1.value = False
   Option2.value = True
   Option2_Click
End Sub

Private Sub txtNombres_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Or KeyAscii = 10 Then
     cmdBuscar_Click
  End If

End Sub

Private Sub txtNombres_LostFocus()
  txtNombres.Text = UCase(txtNombres.Text)
End Sub

Private Sub txtNroContrato_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Or KeyAscii = 10 Then
     cmdBuscar_Click
  End If

End Sub
