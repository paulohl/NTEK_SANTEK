VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fHCopiarOD 
   Caption         =   "Importar Fotos desde Origen a Destino"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8130
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Carpeta Origen"
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9555
      Begin VB.CommandButton bCopiar 
         Caption         =   "Copiar"
         Height          =   500
         Left            =   5070
         Picture         =   "fHCopiarOD.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   8220
         Width           =   900
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Crear carpeta [Fecha] en la carpeta Destino."
         Height          =   225
         Left            =   930
         TabIndex        =   11
         Top             =   7980
         Value           =   1  'Checked
         Width           =   4065
      End
      Begin VB.Frame Frame3 
         Caption         =   "Carpeta Destino"
         Height          =   3585
         Left            =   30
         TabIndex        =   7
         Top             =   3600
         Width           =   9345
         Begin VB.DriveListBox Drive2 
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
            Left            =   150
            TabIndex        =   10
            Top             =   270
            Width           =   5250
         End
         Begin VB.DirListBox Dir2 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Left            =   150
            TabIndex        =   9
            Top             =   660
            Width           =   5250
         End
         Begin VB.FileListBox File2 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2790
            Left            =   5520
            TabIndex        =   8
            Top             =   660
            Width           =   3705
         End
      End
      Begin VB.CommandButton bCancelar 
         Caption         =   "Cancelar"
         Height          =   500
         Left            =   8370
         Picture         =   "fHCopiarOD.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8220
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
         Height          =   2790
         Left            =   5580
         TabIndex        =   3
         Top             =   600
         Width           =   3705
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
         Height          =   2895
         Left            =   210
         TabIndex        =   2
         Top             =   600
         Width           =   5250
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
         Left            =   210
         TabIndex        =   1
         Top             =   210
         Width           =   5250
      End
      Begin VB.Label lRutaDestino 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
         Height          =   315
         Left            =   90
         TabIndex        =   6
         Top             =   7620
         Width           =   9240
      End
      Begin VB.Label lRutaActual 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
         Height          =   315
         Left            =   90
         TabIndex        =   4
         Top             =   7260
         Width           =   9240
      End
   End
End
Attribute VB_Name = "fHCopiarOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cFG = "Nº      | ARCHIVO                                           | CÉDULA                                                "

Private Sub bCancelar_Click()
  Unload Me
End Sub

Private Sub bCopiar_Click()
  Dim s1 As String, s2 As String
  Dim i As Integer, sa As String, sb As String
  '--Crear subcarpeta de Fecha (si se requiere)
  
  Load fMensaje2
  fMensaje2.Label1.Caption = "Copiando " & CStr(File1.ListCount) & " archivos, Espere..."
  fMensaje2.ProgressBar1.Min = 0
  fMensaje2.ProgressBar1.Max = File1.ListCount
  fMensaje2.ProgressBar1.Value = 0
  fMensaje2.Show
  DoEvents
   
  s1 = lRutaActual.Caption
  s2 = lRutaDestino.Caption
  
  If Check3.Value = vbChecked Then MkDir s2
  
  For i = 0 To File1.ListCount - 1
    fMensaje2.ProgressBar1.Value = fMensaje2.ProgressBar1.Value + 1
    fMensaje2.Show
    DoEvents
    
    sa = s1 & "\" & File1.List(i)
    sb = s2 & "\" & File1.List(i)
        
    FileCopy sa, sb
  Next i
  
  Unload fMensaje2

End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
End Sub

Private Sub Dir1_Click()
  'lRutaActual.Caption = Dir1.Path
End Sub

Private Sub Dir2_Change()
  File2.Path = Dir2.Path
  If Check3.Value = vbChecked Then
    lRutaDestino = Dir2.Path & "\" & Format(Date, "dd-mm-yyyy")
  Else
    lRutaDestino = Dir2.Path
  End If
End Sub


Private Sub Dir2_Click()
  If Check3.Value = vbChecked Then
    lRutaDestino = Dir2.Path & "\" & Format(Date, "dd-mm-yyyy")
  Else
    lRutaDestino = Dir2.Path
  End If
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

Private Sub SeleccionarCampo(eTexto As TextBox)
  eTexto.SelStart = 0
  eTexto.SelLength = Len(eTexto.Text)
End Sub


Private Sub Drive2_Change()
  On Error Resume Next
  Dir2.Path = Drive2.Drive
  If Err.Number <> 0 Then
    MsgBox "Unidad No Está Disponible...", vbCritical, "Información"
    Drive2.Drive = "C"
    'Dir1.Path = "C"
  End If
End Sub

Private Sub File1_PathChange()
  'CargarArchivosDeDisco
  lRutaActual.Caption = File1.Path
  
  'If Trim(FG.TextMatrix(1, 1)) <> "" Then
  '  Analizar_Archivo Trim(FG.TextMatrix(1, 1))
  'End If
  
End Sub

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




Private Sub File2_Click()
  lRutaDestino.Caption = File2.Path
End Sub

Private Sub Form_Load()
  Dim UI As String
  
  'Frame2.Visible = False
  'List1.Clear
  'cCP.Clear
  'cSC.Clear
  
  UI = Mid(App.Path, 1, 2)
    
  Dir1.Path = UI
  
  File1.Refresh
  'CargarArchivosDeDisco
  
  'Cargar_Clientes
  
  'Check2.Value = vbUnchecked
  Check3.Value = vbChecked
  
  Dir2.Path = UI
  File2.Refresh
  
End Sub

