VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fDatosBoton 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATOS DEL BOTÓN"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   2130
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   500
      Left            =   4140
      Picture         =   "fDatosBoton.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2130
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton bAceptar 
      Caption         =   "Aceptar"
      Height          =   500
      Left            =   3000
      Picture         =   "fDatosBoton.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2130
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Height          =   2025
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7995
      Begin VB.CommandButton bBuscar 
         Height          =   345
         Left            =   7230
         Picture         =   "fDatosBoton.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Examinar"
         Top             =   1530
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.TextBox eImagen 
         Height          =   285
         Left            =   1830
         MaxLength       =   100
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1560
         Width           =   5325
      End
      Begin VB.ComboBox cProductos 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1110
         Width           =   5325
      End
      Begin VB.TextBox eTitulo 
         Height          =   285
         Left            =   1830
         MaxLength       =   20
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   660
         Width           =   5295
      End
      Begin VB.Label lNumero 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1830
         TabIndex        =   10
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Imagen del Botón:"
         Height          =   195
         Left            =   450
         TabIndex        =   9
         Top             =   1590
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Producto (Inventario):"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1170
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Título del Botón:"
         Height          =   195
         Left            =   570
         TabIndex        =   7
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   1140
         TabIndex        =   6
         Top             =   300
         Width           =   600
      End
   End
End
Attribute VB_Name = "fDatosBoton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bAceptar_Click()
  If Me.Caption = "DATOS DEL BOTÓN" Then
    With fConfigurarBotones
      .FG.TextMatrix(.FG.Row, 1) = Trim(eTitulo.Text)
      .FG.TextMatrix(.FG.Row, 2) = cProductos.Text
      .FG.TextMatrix(.FG.Row, 3) = eImagen.Text
    End With
  Else
    With fConfigurarBotones2
      .FG2.TextMatrix(.FG2.Row, 1) = Trim(eTitulo.Text)
      .FG2.TextMatrix(.FG2.Row, 2) = cProductos.Text
      .FG2.TextMatrix(.FG2.Row, 3) = eImagen.Text
    End With
  End If
  Unload Me
End Sub

Private Sub bBuscar_Click()
  Dim s As String
  CommonDialog1.ShowOpen
  If Err.Number = 0 Then
    s = CommonDialog1.FileName
    eImagen.Text = s
  End If
End Sub

Private Sub bCancelar_Click()
  Unload Me
End Sub

Private Sub eTitulo_LostFocus()
  eTitulo.Text = UCase(eTitulo.Text)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

