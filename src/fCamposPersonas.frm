VERSION 5.00
Begin VB.Form fCamposPersonas 
   Caption         =   "Definición de Campo"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bAceptar 
      Caption         =   "Aceptar"
      Height          =   500
      Left            =   1740
      Picture         =   "fCamposPersonas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   500
      Left            =   2880
      Picture         =   "fCamposPersonas.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5505
      Begin VB.TextBox eCampo 
         Height          =   315
         Left            =   1590
         MaxLength       =   20
         TabIndex        =   0
         Text            =   "eCampo"
         Top             =   270
         Width           =   3630
      End
      Begin VB.TextBox eAncho 
         Height          =   315
         Left            =   1590
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   660
         Width           =   400
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Nombre del Campo:"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   330
         Width           =   1395
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Ancho:"
         Height          =   195
         Left            =   1050
         TabIndex        =   5
         Top             =   720
         Width           =   510
      End
   End
End
Attribute VB_Name = "fCamposPersonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bAceptar_Click()
  Modulo.vTemporal1 = Trim(eCampo.Text)
  Modulo.vTemporal2 = Trim(eAncho.Text)
  Modulo.fModalResult = Modulo.fModalResultOK
  Unload Me
End Sub

Private Sub bCancelar_Click()
  Modulo.vTemporal1 = ""
  Modulo.vTemporal2 = ""
  Modulo.fModalResult = Modulo.fModalResultCANCEL
  Unload Me
End Sub

Private Sub eCampo_Change()
  EnMayusculas eCampo
End Sub

Private Sub eCampo_LostFocus()
  If Trim(eCampo.Text) <> "" Then eCampo.Text = UCase(Modulo.DepurarStr(eCampo.Text, " "))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

