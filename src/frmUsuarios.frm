VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmUsuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2475
      Left            =   360
      TabIndex        =   2
      Top             =   3060
      Width           =   7035
      Begin MSComctlLib.ListView ListVUsuarios 
         Height          =   2055
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   3625
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cedula"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "E-Mail"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Perfil"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2835
      Left            =   300
      TabIndex        =   0
      Top             =   60
      Width           =   7095
      Begin VB.TextBox txtConfirmacion 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5100
         PasswordChar    =   "*"
         TabIndex        =   22
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtContraseña 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5100
         PasswordChar    =   "*"
         TabIndex        =   21
         Top             =   660
         Width           =   1815
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   5640
         TabIndex        =   18
         Top             =   2220
         Width           =   1335
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   5640
         TabIndex        =   17
         Top             =   1740
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   4020
         TabIndex        =   16
         Top             =   1740
         Width           =   1335
      End
      Begin VB.CheckBox Activo 
         Caption         =   "Activo"
         Height          =   315
         Left            =   4140
         TabIndex        =   15
         Top             =   240
         Width           =   1635
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   4020
         TabIndex        =   14
         Top             =   2220
         Width           =   1335
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Left            =   1140
         TabIndex        =   10
         Top             =   2340
         Width           =   2115
      End
      Begin VB.TextBox txtUsuario 
         Height          =   315
         Left            =   1140
         TabIndex        =   9
         Top             =   660
         Width           =   2775
      End
      Begin VB.TextBox txtCedula 
         Height          =   315
         Left            =   1140
         TabIndex        =   7
         Top             =   1920
         Width           =   2115
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1140
         TabIndex        =   6
         Top             =   1500
         Width           =   2775
      End
      Begin MSDataListLib.DataCombo cmbDbPerfiles 
         Bindings        =   "frmUsuarios.frx":0000
         DataSource      =   "AdoPerfiles"
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "Perfiles"
      End
      Begin MSAdodcLib.Adodc AdoUsuarios 
         Height          =   330
         Left            =   6480
         Top             =   180
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   582
         ConnectMode     =   1
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   1
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "AdoUsuarios"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo cmbDbUsuarios 
         Bindings        =   "frmUsuarios.frx":001A
         DataSource      =   "AdoUsuarios"
         Height          =   315
         Left            =   1140
         TabIndex        =   12
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "Usuarios"
      End
      Begin VB.Label Label7 
         Caption         =   "Confirmar:"
         Height          =   255
         Left            =   4200
         TabIndex        =   20
         Top             =   1140
         Width           =   1035
      End
      Begin VB.Label Label6 
         Caption         =   "Contraseña:"
         Height          =   255
         Left            =   4200
         TabIndex        =   19
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Perfil:"
         Height          =   315
         Left            =   300
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Email:"
         Height          =   255
         Left            =   300
         TabIndex        =   11
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Cedula:"
         Height          =   255
         Left            =   300
         TabIndex        =   8
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   300
         TabIndex        =   5
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario:"
         Height          =   315
         Left            =   300
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc AdoPerfiles 
      Height          =   330
      Left            =   3600
      Top             =   4980
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoPerfiles"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lEsNuevo As Boolean

Private Sub cmbDbUsuario_Change()

End Sub

Private Sub cmbDbUsuarios_Change()
If cmbDbUsuarios.Text <> "" Then sCargarUsuario cmbDbUsuarios.BoundText
End Sub

Private Sub cmdCerrar_Click()
   Unload Me
End Sub

Private Sub cmdNuevo_Click()
   LimpiarFormulario
   lEsNuevo = True
   txtUsuario.SetFocus
End Sub

Private Sub Command1_Click()
   Dim lCn As New ADODB.Connection
   Dim lSqltxt As String
   If txtContraseña.Text <> txtConfirmacion.Text Then
      MsgBox "Las contraseñas no coinciden", vbExclamation
      txtContraseña.SetFocus
   End If
   lCn.Open Modulo.DBConexionSQL
   If lEsNuevo = False Then
      lSqltxt = UCase("Update Usuarios2 set Nombre='" & txtNombre.Text & "'," _
      & "Cedula='" & txtCedula.Text & "'," _
      & "Email='" & txtEmail.Text & "'," _
      & "CodigoPerfil=" & cmbDbPerfiles.BoundText & "," _
      & "Activo=" & Activo.value & " where " _
      & "ID= " & cmbDbUsuarios.BoundText)
   Else
      lSqltxt = UCase("Insert into Usuarios2 (Usuario,Nombre,Cedula,Email,Contraseña,codigoPerfil,Activo) Values('" _
         & txtUsuario.Text & "','" & txtNombre.Text & "','" & txtCedula.Text & "','" & txtEmail.Text & "','" & txtContraseña.Text & "','" _
         & cmbDbPerfiles.BoundText & "'," & Activo.value & ")")
         lEsNuevo = False
   End If
   lCn.Execute lSqltxt
   Form_Load
End Sub

Private Sub LimpiarFormulario()
   txtNombre.Text = ""
   txtCedula.Text = ""
   txtEmail.Text = ""
   txtUsuario.Text = ""
   txtContraseña.Text = ""
   txtConfirmacion.Text = ""
   Activo.value = 0
   cmbDbPerfiles.Text = ""
   cmbDbUsuarios.Text = ""
   cmbDbPerfiles.Text = ""
   
End Sub

Private Sub Form_Load()
   sCargarAdoUsuarios
   sCargarListaUsuarios
   LimpiarFormulario
   lEsNuevo = False
End Sub

Private Sub sCargarAdoUsuarios()
    AdoUsuarios.ConnectionString = Modulo.DBConexionSQL
    AdoUsuarios.RecordSource = "Select * from Usuarios2 order by Nombre" 'where codigo > 0
    AdoUsuarios.Refresh
    cmbDbUsuarios.DataField = "Nombre"
    cmbDbUsuarios.BoundColumn = "id"
    cmbDbUsuarios.ListField = "Nombre"
    cmbDbUsuarios.Refresh
    lEsNuevo = False
    
    AdoPerfiles.ConnectionString = Modulo.DBConexionSQL
    AdoPerfiles.RecordSource = "Select * from Perfiles where codigo > 0 order by Nombre"
    AdoPerfiles.Refresh
    cmbDbPerfiles.DataField = "Nombre"
    cmbDbPerfiles.BoundColumn = "codigo"
    cmbDbPerfiles.ListField = "Nombre"
    cmbDbPerfiles.Refresh
    
    
End Sub


Sub sCargarListaUsuarios()
   Dim lReg As New ADODB.Recordset
   Dim lCn As New ADODB.Connection
   Dim lItem As ListItem
   Dim i As Integer
   Dim j As Integer
   j = 1
   lCn.Open Modulo.DBConexionSQL
   Set lReg = lCn.Execute("Select * from Usuarios2 order by nombre")
   ListVUsuarios.ListItems.Clear
   Do While lReg.EOF = False
      Set lItem = ListVUsuarios.ListItems.Add(, , lReg!ID)
          lItem.SubItems(1) = lReg!Usuario
          lItem.SubItems(2) = lReg!Nombre
          lItem.SubItems(3) = lReg!Cedula
          lItem.SubItems(4) = lReg!Email
          cmbDbPerfiles.BoundText = lReg!CodigoPerfil
          lItem.SubItems(5) = cmbDbPerfiles.Text
      lReg.MoveNext
   Loop
   
   
End Sub

Sub sCargarUsuario(argIdUsuario As Integer)
   Dim lReg As New ADODB.Recordset
   Dim lCn As New ADODB.Connection
   lCn.Open Modulo.DBConexionSQL
   Set lReg = lCn.Execute("Select * from Usuarios2 Where id=" & argIdUsuario)
   If lReg.EOF = False Then
      txtUsuario.Text = lReg!Usuario
      txtNombre.Text = lReg!Nombre
      txtCedula.Text = lReg!Cedula
      txtEmail.Text = lReg!Email
      cmbDbPerfiles.BoundText = lReg!CodigoPerfil
      Activo.value = lReg!Activo
      txtContraseña.Text = lReg!contraseña & "     "
      txtConfirmacion.Text = lReg!contraseña & "     "
      cmbDbUsuarios.BoundText = lReg!ID
   Else
      MsgBox "No se encontro informacion del usuario seleccionado", vbExclamation
   End If
   
   
End Sub

Private Sub ListVUsuarios_DblClick()
If ListVUsuarios.ListItems.Count > 0 Then sCargarUsuario ListVUsuarios.SelectedItem.Text

End Sub
