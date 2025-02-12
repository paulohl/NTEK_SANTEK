VERSION 5.00
Begin VB.Form frmCantidad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cantidad a descontar"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   2340
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   1140
      TabIndex        =   0
      Top             =   180
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   1140
      TabIndex        =   2
      Top             =   660
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   660
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "CANTIDAD"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   180
      Width           =   915
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtCantidad_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then ' tecla escape
   Unload Me
End If

If KeyAscii = 10 Or KeyAscii = 13 Then ' tecla enter
   Command1.SetFocus
End If

If (KeyAscii <> 8) And (KeyAscii < 48) Or (KeyAscii > 57) Then
   KeyAscii = 0
End If

End Sub

Private Sub Command1_Click()
  fSistema.lCantidad = txtCantidad.Text
  Unload Me
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then ' tecla escape
   Unload Me
End If

End Sub

Private Sub Command2_Click()
  fSistema.lCantidad = 0
  Unload Me
End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then ' tecla escape
   Unload Me
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then ' tecla escape
   Unload Me
End If

End Sub


