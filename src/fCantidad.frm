VERSION 5.00
Begin VB.Form fCantidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta Cliente"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   500
      Left            =   1560
      Picture         =   "fCantidad.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1380
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4035
      Begin VB.TextBox ePre 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1650
         TabIndex        =   0
         Text            =   "Text3"
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Monto Bs:"
         Height          =   195
         Left            =   870
         TabIndex        =   4
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Aperturar Cuenta con Monto Deuda del Cliente:"
         Height          =   195
         Left            =   345
         TabIndex        =   3
         Top             =   240
         Width           =   3360
      End
   End
End
Attribute VB_Name = "fCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ePre_GotFocus()
  ePre.SelStart = 0
  ePre.SelLength = Len(ePre.Text)
End Sub

Private Sub ePre_KeyPress(KeyAscii As Integer)
  If KeyAscii <> vbKeyDelete And _
     KeyAscii <> vbKeyBack And _
     KeyAscii <> vbKeyReturn And _
     KeyAscii <> Asc(".") Then
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0
      Beep
    End If
  Else
    If KeyAscii = vbKeyReturn Then eAD.SetFocus
    If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
  End If
End Sub

Private Sub ePre_LostFocus()
  Dim d As Double
  Dim s As String
  s = Trim(ePre.Text)
  If s = "" Then s = "0,00"
  On Error Resume Next
  d = CDbl(s)
  If Err.Number <> 0 Then
    MsgBox "El monto No es válido, Revise...", vbCritical, "Información"
    ePre.SetFocus
  Else
    ePre.Text = Format(d, "#,0.00")
  End If
End Sub

Private Sub bAceptar_Click()
  'Modulo.Crear_Cuenta_Cliente
  Unload Me
End Sub

