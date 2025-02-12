VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBuscar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Búsqueda Avanzada"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14265
   Icon            =   "frmBuscar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   14265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   315
      Left            =   2760
      TabIndex        =   12
      Top             =   2520
      Width           =   5895
      _ExtentX        =   10398
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
      Width           =   14055
      Begin MSComctlLib.ListView ListvResultados 
         Height          =   2235
         Left            =   60
         TabIndex        =   6
         Top             =   180
         Width           =   13875
         _ExtentX        =   24474
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha-Mov"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "CodigoCliente"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nombre"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "CodigoSubcliente"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Nombre"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Cedula"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Nombre"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Cargo"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Vence"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Observaciones"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2205
      Left            =   3120
      TabIndex        =   0
      Top             =   60
      Width           =   4965
      Begin VB.TextBox txtCargo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2700
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Por Cargo"
         Height          =   285
         Left            =   420
         TabIndex        =   10
         Top             =   1200
         Visible         =   0   'False
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
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   12240
      Stretch         =   -1  'True
      Top             =   300
      Width           =   1815
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lDestino
Dim lRegClientes As New ADODB.Recordset
Public lTabla As String
Private Sub Form_Load()
   lDestino = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
   If Len(lDestino) = 3 Then
      lDestino = Mid(lDestino, 1, 2)
   End If
   Set lRegClientes = Nothing
   lRegClientes.Open "Select * from Clientes", Modulo.DBConexionSQL, adOpenStatic
End Sub

Private Sub ListvResultados_Click()
Dim lDst As String
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
   lDst = lDestino & "\" & Trim(lRegClientes!Nombre) & "\" & "FOTOS"
   lDst = lDst & "\" & Replace(ListvResultados.SelectedItem.ListSubItems(6), ".", "") & ".jpg"
   On Error GoTo falla
   Image1.Picture = LoadPicture(lDst)
Else
   Image1.Picture = LoadPicture()
End If
falla:
  If Err.Number <> 0 Then
   Image1.Picture = LoadPicture()
  End If

End Sub

Private Sub ListvResultados_DblClick()
   If lTabla <> "Diario" Then
      Dim i As Integer
      fPersonasAct.Adodc1.Recordset.MoveFirst
      fPersonasAct.Adodc1.Recordset.Find "ID=" & ListvResultados.SelectedItem.Text
      fPersonasAct.DataGrid1.Refresh
      Me.Hide
   End If
End Sub

Private Sub ListvResultados_KeyDown(KeyCode As Integer, Shift As Integer)
   ListvResultados_Click
End Sub


Private Sub ListvResultados_KeyUp(KeyCode As Integer, Shift As Integer)
   ListvResultados_Click
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
 SqlTxt = "Select * from " & lTabla & " where "
 
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
    If lTabla = "Diario" Then
       SqlTxt = "Select * from " & lTabla  ''Sqltxt & " Nombre like '%" & txtNombres.Text & "%'"
    Else
      SqlTxt = "Select * From " & lTabla & " Where Nombre like '%" & txtNombres.Text & "%'"
    End If
 End If
 ListvResultados.ListItems.Clear
 SqlTxt = SqlTxt & " Order by ID"
 lCn.ConnectionString = Modulo.DBConexionSQL
 lCn.Open
 Set lReg = lCn.Execute(SqlTxt)

 If lTabla <> "Diario" Then
  If lReg.EOF = True Then
     MsgBox "No se encontró ningún registro con el criterio establecido", vbInformation
     Exit Sub
  End If
   Do While lReg.EOF = False
     Set lItem = ListvResultados.ListItems.Add(, , lReg!ID)
     lItem.SubItems(1) = "NO APLICA" ''Format(lReg!Fecha, "dd/MM/yyyy")
     lItem.SubItems(2) = lTabla
     lItem.SubItems(3) = fPersonasAct.cCP.Text
     lItem.SubItems(4) = fPersonasAct.cSC.Text
     lItem.SubItems(5) = fPersonasAct.cSC.Text
     lItem.SubItems(6) = Trim(lReg!cedula)
     lItem.SubItems(7) = Trim(lReg!Nombre)
     lItem.SubItems(8) = Trim(lReg!cargo)
     lItem.SubItems(9) = Trim(lReg!vence)
     lItem.SubItems(10) = "" 'Trim(lReg!Observaciones)
     lReg.MoveNext
   Loop
   Exit Sub
 End If
 'If Option3.Value = True Then
 '   Sqltxt = Sqltxt & " Cargo like '%" & txtApellidos.Text & "%'"
 'End If

 'If Option4.Value = True Then
 '   Sqltxt = Sqltxt & " b.n_contrato like '%" & txtNroContrato.Text & "%'"
 'End If

 
 If lReg.EOF = False Then
    lReg3.Open "Select * from Clientes", Modulo.DBConexionSQL, adOpenKeyset
    lReg4.Open "Select * from SubClientes order by id", Modulo.DBConexionSQL, adOpenKeyset
    PBar1.Min = 0
    PBar1.value = 0
    NumRegistros = 0
    Do While lReg.EOF = False
       NumRegistros = NumRegistros + 1
       lReg.MoveNext
    Loop
    PBar1.Max = NumRegistros
    lReg.MoveFirst
    
    lAuxTabla = Trim(lReg!tabla)
    If Option2.value = True Then
       lReg2.Open "Select * from [" & Trim(lAuxTabla) & "] Where Nombre like '%" & txtNombres.Text & "%'", Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    Else
       lReg2.Open "Select * from [" & Trim(lAuxTabla) & "]", Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    End If
    
    Do While lReg.EOF = False
       If Option2.value = True Then
          If lAuxTabla <> Trim(lReg!tabla) Then
             lAuxTabla = Trim(lReg!tabla)
          End If
             lReg2.Close
             lReg2.Open "Select * from [" & Trim(lAuxTabla) & "] Where Nombre like '%" & txtNombres.Text & "%' and cedula='" & Trim(lReg!cedula) & "'", Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
          
          'If lReg2.EOF = False Then
          '   lReg2.MoveFirst
          '   'Sqltxt = "Select * from [" & Trim(lReg!tabla) & "] Where Nombre Like '%" & txtNombres.Text & "%'"
          '   lReg2.Find "cedula='" & Trim(lReg!cedula) & "'"
          'End If
       Else
          If lAuxTabla <> Trim(lReg!tabla) Then
             lAuxTabla = Trim(lReg!tabla)
             lReg2.Close
             lReg2.Open "Select * from [" & Trim(lAuxTabla) & "]", Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
          End If
          If lReg2.EOF = False Then
             lReg2.MoveFirst
          
             lReg2.Find "Cedula='" & Trim(lReg!cedula) & "'"
             'Sqltxt = "Select * from [" & Trim(lReg!tabla) & "] Where Cedula='" & fFormatearCedula(lReg!Cedula) & "'"
          End If
       End If
       
       'Set lReg2 = lCn.Execute(Sqltxt)
       If lReg2.EOF = False Then
          lReg3.MoveFirst
          lReg3.Find "Codigo=" & IIf(IsNull(lReg!Cliente), "", lReg!Cliente)
          If lReg4.EOF = False Then
             lReg4.MoveFirst
             lReg4.Find "Id=" & lReg!SubCliente
             lReg4.Find "Cliente=" & lReg!Cliente
          End If
          Set lItem = ListvResultados.ListItems.Add(, , lReg!ID)
          lItem.SubItems(1) = Format(lReg!Fecha, "dd/MM/yyyy")
          lItem.SubItems(2) = Trim(lReg!Cliente)
          If lReg3.EOF = False Then lItem.SubItems(3) = Trim(lReg3!Nombre)
          lItem.SubItems(4) = Trim(IIf(IsNull(lReg!SubCliente), "", lReg!SubCliente))
          If lReg4.EOF = False Then lItem.SubItems(5) = Trim(lReg4!Nombre)
          lItem.SubItems(6) = Trim(lReg2!cedula)
          lItem.SubItems(7) = Trim(lReg2!Nombre)
          For i = 0 To lReg2.Fields.Count - 1
             If UCase(lReg2.Fields(i).Name) = "CARGO" Then lItem.SubItems(8) = Trim(lReg2!cargo)
          Next i
          lItem.SubItems(9) = Trim(lReg2!vence)
          lItem.SubItems(10) = Trim(lReg!Observaciones)
       End If
       lReg.MoveNext
       PBar1.value = PBar1.value + 1
    Loop
 
 
    'Do While lReg.EOF = False
    '   Set lItem = ListvResultados.ListItems.Add(, , lReg!Nombres)
    '   lItem.SubItems(1) = IIf(IsNull(lReg!apellidos), "", lReg!apellidos)
    '   lItem.SubItems(2) = lReg!Cedula_rif
    '   lItem.SubItems(3) = lReg!N_contrato
    '   lItem.SubItems(4) = lReg!ID
    '   ''lItem.SubItems(3) = IIf(IsNull(lReg!NroTarjeta), "Sin Tarjeta", lReg!NroTarjeta)
    '   ''lItem.SubItems(3) = lReg!CodigoBarra
    '   lReg.MoveNext
    'Loop
    ListvResultados.SetFocus
    ListvResultados_Click
 Else
    'If Option4.Value = False Then
    '   Sqltxt = "Select a.Nombres,a.Apellidos,a.Cedula_Rif,a.id from Persona a where "
    'End If
    '   If Option1.Value = True Then
    '      Sqltxt = Sqltxt & " cedula_Rif like '%" & txtCedula.Text & "%'"
    '   End If
'
'       If Option2.Value = True Then
'         Sqltxt = Sqltxt & " nombres like '%" & txtNombres.Text & "%'"
'       End If'''
'
'       If Option3.Value = True Then
'         Sqltxt = Sqltxt & " apellidos like '%" & txtApellidos.Text & "%'"
'       End If'

       'Sqltxt = Sqltxt & " Order by cedula_Rif"
'       Set lReg = gCn.Execute(Sqltxt)
'       If lReg.EOF = False Then
'          Do While lReg.EOF = False
'             Set lItem = ListvResultados.ListItems.Add(, , lReg!Nombres)
'             lItem.SubItems(1) = IIf(IsNull(lReg!apellidos), "", lReg!apellidos)
'             lItem.SubItems(2) = lReg!Cedula_rif
'             lItem.SubItems(3) = "No tiene" 'lReg!N_contrato
'             lItem.SubItems(4) = lReg!ID
'             lReg.MoveNext
'          Loop
'          ListvResultados.SetFocus
'       Else
          MsgBox "No se encontró ningún registro con el criterio establecido", vbInformation
'       End If
 End If
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
