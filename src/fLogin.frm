VERSION 5.00
Begin VB.Form fLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SANTEK - Seguridad"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4620
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bConexion 
      Height          =   500
      Left            =   3960
      Picture         =   "fLogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "IP Conexión"
      Top             =   420
      UseMaskColor    =   -1  'True
      Width           =   500
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   500
      Left            =   2400
      Picture         =   "fLogin.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton bAceptar 
      Caption         =   "Aceptar"
      Height          =   500
      Left            =   1290
      Picture         =   "fLogin.frx":0F8C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Editar"
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Height          =   1245
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4605
      Begin VB.TextBox ePass 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   690
         Width           =   2115
      End
      Begin VB.TextBox eUsu 
         Height          =   315
         Left            =   1710
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   270
         Width           =   2115
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña:"
         Height          =   195
         Left            =   810
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   1050
         TabIndex        =   5
         Top             =   300
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   210
         Picture         =   "fLogin.frx":1516
         Top             =   300
         Width           =   480
      End
   End
End
Attribute VB_Name = "fLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bAceptar_Click()
 Dim lReg As New ADODB.Recordset
 Dim lCn As New ADODB.Connection
 If Modulo.DBConexionSQL.State = adStateClosed Then Modulo.Abrir_BD
 lCn.Open Modulo.DBConexionSQL
 Set lReg = lCn.Execute("Select * from Usuarios2 where usuario='" & LCase(eUsu.Text) & "' and Contraseña='" & LCase(ePass.Text) & "' and activo=1")
 If lReg.EOF = False Then
    sCargarPerfilUsuario lReg!CodigoPerfil
    Modulo.USUARIO_ACTUAL = eUsu.Text
    Unload Me
 Else
    MsgBox "Nombre de usuario y/o contraseña incorrecto", vbCritical
    eUsu.SetFocus
 End If
  '''If Modulo.Usuario_VALIDO2(eUsu.Text, ePass.Text, Modulo.NIVEL_ACTUAL, Modulo.PERMISOS_ACTUAL) Then
  
''    Modulo.USUARIO_ACTUAL = eUsu.Text
''    Unload Me
    
''  Else
  
''    MsgBox "Usuario No Válido...", vbCritical, "Información"
''    eUsu.SetFocus
    
'''  End If
  
End Sub

Private Sub bCancelar_Click()
  Modulo.USUARIO_ACTUAL = ""
  Modulo.NIVEL_ACTUAL = ""
  Modulo.PERMISOS_ACTUAL = ""
  Unload Me
  End
End Sub

Private Sub bConexion_Click()
  Dim s As String
  's = Modulo.IP_Servidor
  
  s = GetSetting(Modulo.APPNAME, "Opciones", "IP", s)
 
  s = InputBox("Indique el Nombre del SERVIDOR: (Ejm. PC-01\SQLEXPRESS)", "CONEXIÓN CON EL SERVIDOR", s)
  
  If s <> "" Then
    SaveSetting Modulo.APPNAME, "Opciones", "IP", s
  End If
  
  Modulo.IP_Servidor = s
  
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

Private Sub Form_Load()
  eUsu.Text = ""
  ePass.Text = ""
  
  Modulo.IP_Servidor = GetSetting(Modulo.APPNAME, "Opciones", "IP", "")
 
End Sub
