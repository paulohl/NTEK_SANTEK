VERSION 5.00
Begin VB.Form fDir2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Directorios"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bAceptar 
      Caption         =   "Aceptar"
      Height          =   500
      Left            =   1590
      Picture         =   "fDir2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4410
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   500
      Left            =   2760
      Picture         =   "fDir2.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4410
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5235
      Begin VB.FileListBox File1 
         Height          =   3015
         Left            =   3000
         TabIndex        =   7
         Top             =   570
         Width           =   2085
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   3900
         Width           =   5000
      End
      Begin VB.DirListBox Dir1 
         Height          =   3015
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Pulse Doble-Clic para Seleccionar..."
         Top             =   570
         Width           =   2865
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Seleccionada:"
         Height          =   195
         Left            =   1920
         TabIndex        =   4
         Top             =   3660
         Width           =   1410
      End
   End
End
Attribute VB_Name = "fDir2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bAceptar_Click()
  If Trim(Text1.Text) = "" Then
    MsgBox "Debe Indicar la Carpeta o ruta del Directorio...", vbCritical, "Información"
  Else
    vTemporal1 = Text1.Text
    Unload Me
  End If
End Sub

Private Sub bCancelar_Click()
  vTemporal1 = ""
  Unload Me
End Sub

Private Sub Dir1_Change()
  Text1.Text = Dir1.Path
  File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
  File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
  If File1.ListIndex >= 0 Then
    Text1.Text = Dir1.Path & "\" & File1.List(File1.ListIndex)
  End If
End Sub

Private Sub Form_Load()
  Text1.Text = ""
  vTemporal1 = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'vTemporal1 = ""
End Sub
