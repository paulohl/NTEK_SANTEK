VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fHEtiquetador 
   Caption         =   "Etiquetador de Fotos"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Carpeta Origen"
      Height          =   9915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15195
      Begin VB.CommandButton bRefresh 
         Height          =   375
         Left            =   4440
         Picture         =   "fHEtiquetador.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Refresh"
         Top             =   7860
         Width           =   405
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFC0&
         Height          =   795
         Left            =   4260
         ScaleHeight     =   735
         ScaleWidth      =   8565
         TabIndex        =   9
         Top             =   7650
         Width           =   8625
         Begin VB.CommandButton bRenombrar 
            Caption         =   "Renombrar Todo"
            Height          =   375
            Left            =   4860
            TabIndex        =   12
            Top             =   150
            Width           =   1500
         End
         Begin VB.TextBox eTotal 
            Height          =   345
            Left            =   3390
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   180
            Width           =   700
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Archivos en Listado:"
            Height          =   195
            Left            =   1410
            TabIndex        =   11
            Top             =   240
            Width           =   1845
         End
      End
      Begin VB.TextBox eCed 
         Height          =   345
         Left            =   8580
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   210
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid FG 
         Height          =   6975
         Left            =   4260
         TabIndex        =   6
         Top             =   600
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   12303
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         FormatString    =   "Nº    | ARCHIVO                                           | CÉDULA                                                "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton bCancelar 
         Caption         =   "Cancelar"
         Height          =   500
         Left            =   7170
         Picture         =   "fHEtiquetador.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8820
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2340
         Left            =   510
         TabIndex        =   3
         Top             =   8970
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   90
         TabIndex        =   2
         Top             =   600
         Width           =   3930
      End
      Begin VB.DriveListBox Drive1 
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
         Left            =   90
         TabIndex        =   1
         Top             =   210
         Width           =   3930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Indique el Número de Cédula:"
         Height          =   195
         Left            =   6390
         TabIndex        =   7
         Top             =   270
         Width           =   2100
      End
      Begin VB.Label lRutaActual 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
         Height          =   795
         Left            =   90
         TabIndex        =   4
         Top             =   7650
         Width           =   3930
      End
   End
End
Attribute VB_Name = "fHEtiquetador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const cFG = "Nº      | ARCHIVO                                           | CÉDULA                                                "

Private Sub bCancelar_Click()
  Unload Me
End Sub

Private Sub bRefresh_Click()
  File1.Refresh
  CargarArchivosDeDisco
End Sub

Private Sub bRenombrar_Click()
  Dim sAO As String  'Archivo Original
  Dim sAN As String  'Archivo Nuevo
  Dim i As Integer
  Dim bHuboRenombre As Boolean
  bHuboRenombre = False
  For i = 1 To FG.Rows - 1
    If Trim(FG.TextMatrix(i, 1)) <> "" And Trim(FG.TextMatrix(i, 2)) <> "" Then
      sAO = lRutaActual.Caption & "\" & FG.TextMatrix(i, 1)
      If Dir(sAO) <> "" Then 'Archivo Existe
        sAN = lRutaActual.Caption & "\" & FG.TextMatrix(i, 2) & ".JPG"
        Name sAO As sAN
        bHuboRenombre = True
      End If
    End If
  Next i
  If bHuboRenombre Then
    MsgBox "Archivos han sido Renombrados Exitosamente.", vbCritical, "Información"
    bRefresh_Click
    eCed.Text = ""
  End If
End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
End Sub

Private Sub Dir1_Click()
  'lRutaActual.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
  On Error Resume Next
  Dir1.Path = Drive1.Drive
  If Err.Number <> 0 Then
    MsgBox "Unidad No Está Disponible...", vbCritical, "Información"
    Drive1.Drive = "C"
    'Dir1.Path = "C"
  End If
End Sub

Private Sub CargarArchivosDeDisco()
  Dim i As Integer, k As Integer
  If File1.ListCount > 0 Then
    FG.Clear
    FG.Rows = 2
    FG.FormatString = cFG
    k = 0
    For i = 0 To File1.ListCount - 1
      FG.Row = i + 1: FG.Col = 0: FG.CellAlignment = flexAlignLeftCenter
      FG.TextMatrix(i + 1, 0) = CStr(i + 1) 'Zeros(CStr(i), 4)
      
      FG.Row = i + 1: FG.Col = 1: FG.CellAlignment = flexAlignLeftCenter
      FG.TextMatrix(i + 1, 1) = UCase(File1.List(i))
      
      k = k + 1
      
      If i < File1.ListCount - 1 Then FG.Rows = FG.Rows + 1
        
      
      eTotal.Text = CStr(k)
    Next i
    
    If Trim(FG.TextMatrix(1, 1)) <> "" Then
      Analizar_Archivo Trim(FG.TextMatrix(1, 1))
    End If
    
  End If
End Sub

Private Sub SeleccionarCampo(eTexto As TextBox)
  eTexto.SelStart = 0
  eTexto.SelLength = Len(eTexto.Text)
End Sub

Private Sub eCed_GotFocus()
  Color_Fila_Disponible
  'SeleccionarCampo eCed
  eCed.Text = ""

End Sub

Private Sub eCed_KeyPress(KeyAscii As Integer)
  Dim i As Integer, j As Integer
  i = 1
  j = -1
  If KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
    KeyAscii = 0
    Exit Sub
  End If
    
  If KeyAscii = vbKeyReturn Then
  
    KeyAscii = 0
  
    If Existe_Valor_FG(Trim(eCed.Text), 2) Then
      MsgBox "Cédula ya Existe en el Listado...", vbCritical, "Información"
      Exit Sub
    End If
  
    Do While (i < FG.Rows) And (j = -1)
      If Trim(FG.TextMatrix(i, 2)) = "" Then j = i
      i = i + 1
    Loop
    If j <> -1 Then
      FG.Row = j: FG.Col = 2: FG.CellAlignment = flexAlignLeftCenter
      FG.TextMatrix(j, 2) = Trim(eCed.Text)
      'SeleccionarCampo eCed
      eCed.Text = ""
      eCed.SetFocus
      
      Color_Fila_Disponible
      
    Else
      MsgBox "No hay Fila Vacía donde copiar la cédula...", vbCritical, "Información"
    End If
  End If
  
End Sub

Private Sub FG_DblClick()
  Dim s As String
  If FG.Col = 2 Then 'está en posición de cédula
    s = FG.TextMatrix(FG.Row, FG.Col)
    s = InputBox("Indique Cedula:", "Editar", s)
    FG.TextMatrix(FG.Row, FG.Col) = s
    eCed.SetFocus
  End If
End Sub

Private Sub File1_PathChange()
  CargarArchivosDeDisco
  lRutaActual.Caption = File1.Path
  
  'If Trim(FG.TextMatrix(1, 1)) <> "" Then
  '  Analizar_Archivo Trim(FG.TextMatrix(1, 1))
  'End If
  
End Sub

Private Function Existe_Valor_FG(sValor As String, iColumna As Integer) As Boolean
  Dim i As Integer
  Dim e As Boolean
  i = 0
  e = False
  Do While i < FG.Rows And Not e
    If FG.TextMatrix(i, iColumna) = sValor Then
      e = True
    Else
      i = i + 1
    End If
  Loop
  Existe_Valor_FG = e
End Function

Private Sub Analizar_Archivo(sNombre As String)
  Dim i As Integer
  Dim s As String
  Dim Caracter As String
  Dim sLetras As String
  Dim sDigitos As String
  
  'Letras : A..65 / Z..90
  'Digitos: 0..48 / 9..57
  s = ""
  sLetras = ""
  sDigitos = ""
  i = 1
  Do While i <= Len(sNombre)
    Caracter = Mid(sNombre, i, 1)
    If Caracter = "." Then
      i = Len(sNombre) + 1
    Else
      Select Case Caracter
        Case "A", "B", "C", "D", _
             "E", "F", "G", "H", _
             "I", "J", "K", "L", _
             "M", "N", "O", "P", _
             "Q", "R", "S", "T", _
             "U", "V", "W", "X", _
             "Y", "Z", "-", "_":
          sLetras = sLetras & Caracter
          
        Case "0", "1", "2", "3", "4", _
             "5", "6", "7", "8", "9":
          sDigitos = sDigitos & Caracter
      End Select
      
      i = i + 1
    End If
      
  Loop
  
  'eTexto.Text = sLetras
  'eDesde.Text = ""
  'eHasta.Text = ""
  
  'If Len(sDigitos) > 0 Then
  '  eDesde.Text = sDigitos
  'End If
  
  
  
   
End Sub

Private Sub Color_Fila_Disponible()
  Dim i As Integer, j As Integer
  
  For i = 1 To FG.Rows - 1
    For j = 0 To FG.Cols - 1
      FG.Row = i
      FG.Col = j
      FG.CellBackColor = vbWhite
    Next j
  Next i
  
  For i = 1 To FG.Rows - 1
    If Trim(FG.TextMatrix(i, 2)) = "" Then
      For j = 0 To FG.Cols - 1
        FG.Row = i
        FG.Col = j
        FG.CellBackColor = vbGreen
      Next j
      Exit Sub
    End If
  Next i
End Sub

Private Sub Form_Load()
  Dim UI As String
  
  UI = Mid(App.Path, 1, 2)
    
  eCed.Text = ""
    
  FG.Clear
  FG.Rows = 2
  FG.FormatString = cFG
  
  Dir1.Path = UI
  
  File1.Refresh
  CargarArchivosDeDisco
  

  
End Sub
