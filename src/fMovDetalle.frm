VERSION 5.00
Begin VB.Form fMovDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item de Producto"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   500
      Left            =   2310
      Picture         =   "fMovDetalle.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2190
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton bAceptar 
      Caption         =   "Aceptar"
      Height          =   500
      Left            =   1170
      Picture         =   "fMovDetalle.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2190
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton bPrecio 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         TabIndex        =   13
         ToolTipText     =   "Editar Precio"
         Top             =   870
         Width           =   345
      End
      Begin VB.TextBox eST 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1170
         TabIndex        =   4
         Text            =   "Text3"
         Top             =   1650
         Width           =   1065
      End
      Begin VB.TextBox eCan 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1170
         TabIndex        =   3
         Text            =   "Text3"
         Top             =   1260
         Width           =   1065
      End
      Begin VB.TextBox ePre 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1170
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   870
         Width           =   1065
      End
      Begin VB.Label lPrecioEspecial 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Atención: Precio Especial Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   585
         Left            =   2490
         TabIndex        =   14
         Top             =   1290
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sub-Total:"
         Height          =   195
         Left            =   330
         TabIndex        =   12
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   390
         TabIndex        =   11
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Precio Bs:"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   930
         Width           =   720
      End
      Begin VB.Label lDes 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Left            =   1170
         TabIndex        =   1
         Top             =   540
         Width           =   2955
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   570
         Width           =   885
      End
      Begin VB.Label lCod 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Left            =   1170
         TabIndex        =   0
         Top             =   210
         Width           =   2955
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   570
         TabIndex        =   8
         Top             =   240
         Width           =   540
      End
   End
End
Attribute VB_Name = "fMovDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bAceptar_Click()
  If Trim(eCan.Text) = "" Then eCan.Text = "1"
  If Not IsNumeric(eCan.Text) Then
    MsgBox "Debe Ingresar la Cantidad...", vbCritical, "Información"
    eCan.SetFocus
    Exit Sub
  End If
  
  If CDbl(eCan.Text) <= 0# Then
    MsgBox "La Cantidad debe ser Mayor que Cero...", vbCritical, "Información"
    eCan.SetFocus
    Exit Sub
  End If
  
  Modulo.fModalResult = Modulo.fModalResultOK
  fMovDetalle.Hide
End Sub

Private Sub bCancelar_Click()
  Modulo.fModalResult = Modulo.fModalResultCANCEL
  Unload Me
End Sub

Private Sub bPrecio_Click()
  ePre.Enabled = Not ePre.Enabled
  If ePre.Enabled Then ePre.SetFocus
End Sub

Private Sub eCan_Change()
  Dim dST As Double
  If Trim(eCan.Text) = "" Then eCan.Text = "0"
  
  If Not IsNumeric(eCan.Text) Then eCan.Text = "0"
  If Not IsNumeric(ePre.Text) Then ePre.Text = "0,00"
  
  dST = CDbl(eCan.Text) * CDbl(ePre.Text)
  eST.Text = Format(dST, "#,0.00")
End Sub

Private Sub eCan_GotFocus()
  eCan.SelStart = 0
  eCan.SelLength = Len(eCan.Text)
End Sub

Private Sub eCan_KeyPress(KeyAscii As Integer)
  If KeyAscii <> vbKeyDelete And _
     KeyAscii <> vbKeyBack And _
     KeyAscii <> vbKeyReturn And _
     KeyAscii <> Asc("-") And _
     KeyAscii <> Asc(".") Then
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0
      Beep
    End If
  Else
    If KeyAscii = vbKeyReturn Then bAceptar.SetFocus
    If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
  End If
End Sub

Private Sub eCan_LostFocus()
  Dim d As Double
  Dim s As String
  s = Trim(eCan.Text)
  If s = "" Then s = "0,00"
  On Error Resume Next
  d = CDbl(s)
  If Err.Number <> 0 Then
    MsgBox "El monto No es válido, Revise...", vbCritical, "Información"
    eCan.SetFocus
  Else
    eCan.Text = Format(d, "#,0.00")
  End If
End Sub

Private Sub ePre_Change()
  Dim dST As Double
  If Trim(eCan.Text) = "" Then eCan.Text = "0"
  
  If Not IsNumeric(eCan.Text) Then eCan.Text = "0"
  If Not IsNumeric(ePre.Text) Then ePre.Text = "0,00"
  
  dST = CDbl(eCan.Text) * CDbl(ePre.Text)
  eST.Text = Format(dST, "#,0.00")
End Sub

Private Sub ePre_GotFocus()
  ePre.SelStart = 0
  ePre.SelLength = Len(ePre.Text)
End Sub

Private Sub ePre_KeyPress(KeyAscii As Integer)

  If KeyAscii = Asc(".") Then
    
    KeyAscii = Asc(",")
  
  Else

    If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyDelete And _
       KeyAscii <> vbKeyBack Then
     
      If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
      End If
    End If
    
  End If
End Sub

Private Sub ePre_LostFocus()
  Dim dST As Double, dP As Double
  If Trim(eCan.Text) = "" Then eCan.Text = "0"
  
  If Not IsNumeric(eCan.Text) Then eCan.Text = "0"
  If Not IsNumeric(ePre.Text) Then ePre.Text = "0,00"
  
  dP = CDbl(ePre.Text)
  dST = CDbl(eCan.Text) * CDbl(ePre.Text)
  eST.Text = Format(dST, "#,0.00")
  ePre.Text = Format(dP, "#,0.00")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    SendKeys "{TAB}"
  End If
End Sub

