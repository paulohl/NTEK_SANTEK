VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPerfiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Perfiles de Usuario"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoPerfiles 
      Height          =   330
      Left            =   5040
      Top             =   7740
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
   Begin VB.Frame Frame2 
      Caption         =   "Accesos"
      Height          =   5295
      Left            =   360
      TabIndex        =   1
      Top             =   2400
      Width           =   4875
      Begin MSComctlLib.ListView ListVPerfiles 
         Height          =   4815
         Left            =   300
         TabIndex        =   6
         Top             =   300
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   8493
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Accesos"
            Object.Width           =   6174
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   60
      Width           =   4875
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   1740
         Width           =   1335
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   4800
         TabIndex        =   8
         Top             =   1380
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtPerfil 
         Height          =   315
         Left            =   420
         TabIndex        =   7
         Top             =   1080
         Width           =   3975
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   3420
         TabIndex        =   5
         Top             =   1740
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   180
         TabIndex        =   4
         Top             =   1740
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo cmbDbPerfiles 
         Bindings        =   "frmPerfiles.frx":0000
         DataSource      =   "AdoPerfiles"
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   600
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre Perfil:"
         Height          =   315
         Left            =   60
         TabIndex        =   2
         Top             =   660
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmPerfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lEsNuevo As Boolean

Private Sub cmbDbPerfiles_Change()
   If cmbDbPerfiles.BoundText <> "" Then
      txtPerfil.Text = cmbDbPerfiles.Text
      lEsNuevo = False
      sCargarPerfiles cmbDbPerfiles.BoundText
   End If
End Sub

Private Sub cmdCerrar_Click()
   Unload Me
End Sub

Private Sub cmdEliminar_Click()
   
   If MsgBox("¿Seguro que desea Eliminar el Perfil '" & txtPerfil.Text & "'?", vbYesNo + vbQuestion) = vbYes Then
      Dim lCn As New ADODB.Connection
      lCn.Open Modulo.DBConexionSQL
      lCn.Execute "Delete from Perfiles Where Codigo=" & cmbDbPerfiles.BoundText
       Form_Load
   End If
End Sub

Private Sub cmdNuevo_Click()
   Dim lCn As New ADODB.Connection
   lCn.Open Modulo.DBConexionSQL
   
   lEsNuevo = True
   cmbDbPerfiles.Text = ""
   txtPerfil.Text = ""
   sCargarPerfiles 0
   txtPerfil.SetFocus
End Sub

Private Sub Command1_Click()
   Dim i As Integer
   Dim lCn As New ADODB.Connection
   Dim lSqltxt As String
   Dim lSqltxt2 As String
   If lEsNuevo = False Then
      lSqltxt = ""
      lSqltxt = "Update Perfiles set Nombre='" & UCase(txtPerfil.Text) & "', "
      For i = 1 To ListVPerfiles.ListItems.Count
         lSqltxt = lSqltxt & ListVPerfiles.ListItems(i).Text & " = " & Abs(CInt(ListVPerfiles.ListItems(i).Checked)) & ", "
      Next i
      lSqltxt = Mid(lSqltxt, 1, Len(lSqltxt) - 2) & " where Codigo=" & cmbDbPerfiles.BoundText
      lCn.Open Modulo.DBConexionSQL
      lCn.Execute lSqltxt
   Else
      lSqltxt = ""
      lSqltxt2 = ""
      lSqltxt = "Insert Into Perfiles (Nombre,"
      For i = 1 To ListVPerfiles.ListItems.Count
         lSqltxt = lSqltxt & ListVPerfiles.ListItems(i).Text & ","
         lSqltxt2 = lSqltxt2 & Abs(CInt(ListVPerfiles.ListItems(i).Checked)) & ","
      Next i
      lSqltxt = Mid(lSqltxt, 1, Len(lSqltxt) - 1) & ") values ('"
      lSqltxt2 = UCase(txtPerfil.Text) & "'," & Mid(lSqltxt2, 1, Len(lSqltxt2) - 1) & ")"
      lSqltxt = lSqltxt & lSqltxt2
      Clipboard.Clear
      Clipboard.SetText lSqltxt
      lCn.Open Modulo.DBConexionSQL
      lCn.Execute lSqltxt
     
   End If
    Form_Load
End Sub

Private Sub Form_Load()
 AdoPerfiles.ConnectionString = Modulo.DBConexionSQL
 AdoPerfiles.RecordSource = "Select * from Perfiles where codigo > 0 order by Nombre"
 AdoPerfiles.Refresh
 cmbDbPerfiles.DataField = "Nombre"
 cmbDbPerfiles.BoundColumn = "codigo"
 cmbDbPerfiles.ListField = "Nombre"
 cmbDbPerfiles.Refresh
 lEsNuevo = False
 txtPerfil.Text = ""
End Sub


Sub sCargarPerfiles(argCodigoPerfil As Integer)
   Dim lReg As New ADODB.Recordset
   Dim lCn As New ADODB.Connection
   Dim lItem As ListItem
   Dim i As Integer
   Dim j As Integer
   j = 1
   lCn.Open Modulo.DBConexionSQL
   Set lReg = lCn.Execute("Select * from Perfiles Where Codigo=" & argCodigoPerfil)
   ''Activo.value = lReg!Activo
   If lReg.EOF = False Then
      ListVPerfiles.ListItems.Clear
      For i = 2 To lReg.Fields.Count - 1
         Set lItem = ListVPerfiles.ListItems.Add(, , UCase(lReg.Fields(i).Name))
         If lReg.Fields(i).value = 1 Then ListVPerfiles.ListItems(j).Checked = True
         j = j + 1
      Next i
   End If
   
   
End Sub
