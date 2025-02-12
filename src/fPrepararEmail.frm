VERSION 5.00
Begin VB.Form fPrepararEmail 
   Caption         =   "Preparar Email Listado de Errores"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bEnviar 
      Caption         =   "Enviar"
      Height          =   500
      Left            =   2640
      Picture         =   "fPrepararEmail.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   1000
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   500
      Left            =   3990
      Picture         =   "fPrepararEmail.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Height          =   4275
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.ListBox ListAdjunto 
         Height          =   255
         Left            =   1860
         TabIndex        =   16
         Top             =   3540
         Width           =   5535
      End
      Begin VB.CommandButton bVerAdjunto 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Ver Adjunto con EXCEL"
         Height          =   315
         Left            =   5550
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3870
         Width           =   1875
      End
      Begin VB.TextBox eAdjunto 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2100
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   4140
         Visible         =   0   'False
         Width           =   5600
      End
      Begin VB.TextBox eMensaje 
         BackColor       =   &H00FFFFFF&
         Height          =   1545
         Left            =   1860
         MaxLength       =   254
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "fPrepararEmail.frx":0B14
         Top             =   1860
         Width           =   5600
      End
      Begin VB.TextBox eTitulo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1860
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1470
         Width           =   5600
      End
      Begin VB.TextBox eEmail 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1860
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1050
         Width           =   5600
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Width           =   5600
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   5600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Adjunto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   11
         Top             =   3510
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   940
         TabIndex        =   9
         Top             =   1890
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titulo del Mensaje:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   7
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1170
         TabIndex        =   5
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Principal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   330
         TabIndex        =   4
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Sub-Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   3
         Top             =   660
         Width           =   1080
      End
   End
End
Attribute VB_Name = "fPrepararEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oMail As clsCDOmail
Attribute oMail.VB_VarHelpID = -1
Dim s4 As String
Dim s5 As String
Dim s6 As String

Private Sub bCancelar_Click()
  Unload Me
End Sub

Private Sub bEnviar_Click()
    'Set oMail = New clsCDOmail
    'With oMail
    '     'datos para enviar
    '    .servidor = "mail.cantv.net"
    '    .puerto = 25
    '    '.UseAuntentificacion = True
    '    '.ssl = True
    '    .Usuario = s4
    '    .PassWord = s5
    '
    '    .Asunto = eTitulo.Text
    '    .Adjunto = eAdjunto.Text
    '    .de = s6
    '    .para = eEmail.Text & ";" & s6
    '    .Mensaje = eMensaje.Text
    '
    '    .Enviar_Backup ' manda el mail
   '
   ' End With
    
   ' Set oMail = Nothing
  
  If MandaMail(Trim(eEmail.Text), "", "", Trim(eTitulo.Text), Trim(eMensaje.Text), ListAdjunto) = True Then
    Unload Me
  End If
   Unload Me
End Sub

' envio completo
Private Sub oMail_EnvioCompleto()
    MsgBox "Mensaje enviado", vbInformation
End Sub
' error al enviar
Private Sub oMail_Error(Descripcion As String, Numero As Variant)
    MsgBox Descripcion, vbCritical, Numero
End Sub


Private Sub Form_Load()
  Dim i As Integer
  Dim s1 As String, s2 As String, s3 As String
  
  Combo1.Clear
  For i = 0 To fPersonasAct.cCP.ListCount - 1
    Combo1.AddItem fPersonasAct.cCP.List(i)
  Next i
  
  Combo1.ListIndex = fPersonasAct.cCP.ListIndex
  
  Combo2.Clear
  For i = 0 To fPersonasAct.cSC.ListCount - 1
    Combo2.AddItem fPersonasAct.cSC.List(i)
  Next i
  
  Combo2.ListIndex = fPersonasAct.cSC.ListIndex
  
  Combo1.Enabled = False
  Combo2.Enabled = False
  
  s1 = Mid(fPersonasAct.cCP.Text, 1, 6)
  s2 = Mid(fPersonasAct.cSC.Text, 1, 6)
  eEmail.Text = LCase(Modulo.Correo_E(s1, s2))
  
  s1 = ""
  s2 = ""
  s3 = ""
  s4 = ""
  s5 = ""
  s6 = ""
  
  CargarOpcionesCorreo s1, s2, s3, s4, s5, s6
  eTitulo.Text = s1
  eMensaje.Text = s2
  eAdjunto.Text = ""
  bVerAdjunto.Enabled = False
  
  
End Sub
