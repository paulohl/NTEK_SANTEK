VERSION 5.00
Begin VB.Form frmUtilFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Utilidades"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   300
      TabIndex        =   7
      Top             =   1020
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   4380
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   3060
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuitarPuntos 
      Caption         =   "Quitar Puntos"
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   3795
      Left            =   3360
      TabIndex        =   1
      Top             =   1020
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   360
      TabIndex        =   0
      Top             =   1620
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   7440
      Stretch         =   -1  'True
      Top             =   180
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Archivos"
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Unidades"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Carpetas"
      Height          =   195
      Left            =   300
      TabIndex        =   6
      Top             =   1380
      Width           =   2595
   End
End
Attribute VB_Name = "frmUtilFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuitarPuntos_Click()
 If File1.ListCount <= 0 Then Exit Sub
    
 If MsgBox("¿Esta seguro de eliminar los puntos en los archivos listados?", vbYesNo + vbQuestion) = vbYes Then
    Dim i As Integer
    Dim fileAux As String
    For i = 0 To File1.ListCount - 1
       fileAux = Replace(File1.List(i), ".", "")
       fileAux = Mid(fileAux, 1, Len(fileAux) - 3) & "." & Mid(fileAux, Len(fileAux) - 2, 3)
       'CopyFile Dir1 & "\" & File1.List(i), Dir1 & "\" & fileAux, True
       'If fileAux <> File1.List(i) Then Kill Dir1 & "\" & File1.List(i)
       Name Dir1 & "\" & File1.List(i) As Dir1 & "\" & fileAux
    Next i
 End If
 File1.Refresh
 End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
   File1.Refresh
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

Private Sub File1_Click()
   On Error GoTo falla
   Image1.Picture = LoadPicture(Dir1 & "\" & File1.List(File1.ListIndex))
falla:
   If Err.Number <> 0 Then Image1.Picture = LoadPicture()
End Sub

Private Sub Form_Load()
Dir1.Path = Drive1.Drive
Dir1.Refresh
File1.Path = Dir1.Path
File1.Refresh
End Sub

