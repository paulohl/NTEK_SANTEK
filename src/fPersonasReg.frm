VERSION 5.00
Begin VB.Form fPersonasReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Persona"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   17
      Left            =   2040
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   6840
      Width           =   5500
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   16
      Left            =   2040
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   6420
      Width           =   5500
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   15
      Left            =   2040
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   6000
      Width           =   5500
   End
   Begin VB.CommandButton bAceptar 
      Caption         =   "Aceptar"
      Height          =   500
      Left            =   2790
      Picture         =   "fPersonasReg.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   500
      Left            =   3960
      Picture         =   "fPersonasReg.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   14
      Left            =   2040
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   5580
      Width           =   5500
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   13
      Left            =   2040
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   5190
      Width           =   5500
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   12
      Left            =   2040
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   4800
      Width           =   5500
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   11
      Left            =   2040
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   4410
      Width           =   5500
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   10
      Left            =   2040
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   4020
      Width           =   5500
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   9
      Left            =   2040
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   3630
      Width           =   5500
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   8
      Left            =   2040
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3240
      Width           =   5500
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   7
      Left            =   2040
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2850
      Width           =   5500
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2460
      Width           =   5500
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2070
      Width           =   5500
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1680
      Width           =   5500
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1290
      Width           =   5500
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   900
      Width           =   5500
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   510
      Width           =   5500
   End
   Begin VB.TextBox Texts 
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   5500
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   17
      Left            =   60
      TabIndex        =   37
      Top             =   6900
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   16
      Left            =   60
      TabIndex        =   35
      Top             =   6480
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   15
      Left            =   60
      TabIndex        =   33
      Top             =   6060
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   14
      Left            =   60
      TabIndex        =   28
      Top             =   5640
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   13
      Left            =   60
      TabIndex        =   26
      Top             =   5250
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   12
      Left            =   60
      TabIndex        =   24
      Top             =   4860
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   11
      Left            =   60
      TabIndex        =   22
      Top             =   4470
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   10
      Left            =   60
      TabIndex        =   20
      Top             =   4080
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   9
      Left            =   60
      TabIndex        =   18
      Top             =   3690
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   8
      Left            =   60
      TabIndex        =   16
      Top             =   3300
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   7
      Left            =   60
      TabIndex        =   14
      Top             =   2910
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   6
      Left            =   60
      TabIndex        =   12
      Top             =   2520
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   5
      Left            =   60
      TabIndex        =   10
      Top             =   2130
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   8
      Top             =   1740
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   6
      Top             =   1350
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   4
      Top             =   960
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   570
      Width           =   1905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Columna...0"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   1905
   End
End
Attribute VB_Name = "fPersonasReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MAXCampos = 15
Const CLRFnd = vbBlack
Const CLRTxt = vbGreen

Private Sub Limpiar()
  Dim i As Integer
  For i = 0 To MAXCampos - 1
    Labels(i).Caption = ""
    Texts(i).Text = ""
  Next i
End Sub

Private Sub ActivarTexts(bOnOff As Boolean)
  Dim i As Integer
  For i = 0 To MAXCampos - 1
    'Labels(i).Caption = ""
    'Texts(i).Text = ""
    Texts(i).Enabled = bOnOff
  Next i
End Sub


Private Sub bAceptar_Click()
  
    



  Modulo.fModalResult = Modulo.fModalResultOK
  'Unload Me
  
  
  
  Me.Hide
End Sub

Private Sub bCancelar_Click()
  Modulo.fModalResult = Modulo.fModalResultCANCEL
  'Unload Me
  Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

Private Sub Form_Load()
  Limpiar
  ActivarTexts False
End Sub

Private Sub Formatear_Cedula(xTexto As TextBox)
  Dim i As Integer
  Dim d As Double
  Dim s As String
  If IsNumeric(xTexto.Text) Then
    d = CDbl(xTexto.Text)
    s = Format(d, "#,0")
    xTexto.Text = s
  End If
End Sub

Private Function Hacer_Cadena_Foto(sCedula As String) As String
  Const C1 As String = "."
  Const C2 As String = ","
  Dim i As Integer
  Dim s As String
  s = ""
  For i = 1 To Len(sCedula)
    If (Mid(sCedula, i, 1) <> C1) And (Mid(sCedula, i, 1) <> C2) Then
      s = s & Mid(sCedula, i, 1)
    End If
  Next i
  If s = "" Then s = sCedula
  Hacer_Cadena_Foto = s
End Function

Private Function Existe_Cedula(sCed As String, lID As Long) As Boolean
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim EC As Boolean
  EC = False
  If fPersonasAct.lTablas.ListIndex >= 0 Then
    If Trim(sCed) = "" Then
      EC = False
      Set r = Nothing
      Existe_Cedula = EC
      Exit Function
    End If
    s = "select * from [" & fPersonasAct.lTablas.List(fPersonasAct.lTablas.ListIndex) & "] " & _
        "where cedula = '" & sCed & "'"
    r.Open s, Modulo.DBConexionSQL, adOpenKeyset, adLockReadOnly
    If Not r.EOF Then
      If lID = r.Fields("ID").value Then EC = False Else EC = True
    End If
    r.Close
    Set r = Nothing
  End If
  Existe_Cedula = EC
End Function

Private Sub Texts_Change(Index As Integer)
  Modulo.EnMayusculas Texts(Index)
  If Labels(Index).Caption = "CEDULA" Then  '--Cedula:
    If Texts(Index).Enabled Then
      Formatear_Cedula Texts(Index)
      If Modulo.vIndiceFoto <> -1 Then
        Texts(Modulo.vIndiceFoto).Text = Hacer_Cadena_Foto(Texts(Index).Text)
      End If
      If Existe_Cedula(Texts(Index).Text, vID) = True Then
        MsgBox "Cédula [" & Texts(Index).Text & "] ya Existe, Revise...", vbCritical, "Información"
        Texts(Index).Text = ""
      End If
    End If
  End If
End Sub

Private Sub Texts_DblClick(Index As Integer)
  If Labels(Index).Caption = "VENCE" Then  '--VENCE
     frmVencimiento.Show vbModal
     Texts(Index).Text = frmVencimiento.lVencimiento
  End If
End Sub

Private Sub Texts_GotFocus(Index As Integer)
  If Texts(Index).Enabled Then
    Texts(Index).BackColor = CLRFnd
    Texts(Index).ForeColor = CLRTxt
  End If
    
  If Labels(Index).Caption = "CEDULA" Then  '--Cedula
    Texts(Index).Alignment = vbRightJustify
  End If
  If Labels(Index).Caption = "VENCE" Then  '--Cedula
     Texts(Index).ToolTipText = "Haga doble click si desea modificar la fecha con un ComboBox"
  Else
     Texts(Index).ToolTipText = ""
  End If

End Sub

Private Sub Texts_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 0 Then '--Cedula
    If KeyAscii = Asc(" ") Then KeyAscii = 0
  End If
End Sub

Private Sub Texts_LostFocus(Index As Integer)
  If Texts(Index).Enabled Then
    Texts(Index).BackColor = vbWhite
    Texts(Index).ForeColor = vbBlack
  End If
  
  If Labels(Index).Caption = "CEDULA" Then  '--Cedula
    Texts(Index).Alignment = vbLeftJustify
    Hacer_Cadena_Foto Texts(0).Text
  End If
  
  If Existe_Cedula(Texts(Index).Text, Modulo.vID) = True Then
    MsgBox "Cédula [" & Texts(Index).Text & "] ya Existe, Revise...", vbCritical, "Información"
    Texts(Index).Text = ""
  End If
  Texts(Index).ToolTipText = ""
  
End Sub
