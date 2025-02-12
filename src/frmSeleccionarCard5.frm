VERSION 5.00
Begin VB.Form frmSeleccionarCard5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selecionar el Archivo de CardFive"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9645
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   9435
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   2280
      Width           =   1515
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   2940
      TabIndex        =   1
      Top             =   2280
      Width           =   1515
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmSeleccionarCard5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
Dim r As String
If List1.ListIndex < 0 Then Exit Sub
r = GetSetting(APPNAME, "Opciones", "RutaCard5", "")
If Shell(r & " " & List1.List(List1.ListIndex), vbMaximizedFocus) = 0# Then
    MsgBox "Error: No se pudo Iniciar Card-5" & vbCrLf & CStr(Err.Number) & ":" & Err.Description, "Información"
End If
Unload Me
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Public Function sCargarArchivosCard5(argDir As String) As String
   Dim i As Integer
   Dim j As Integer
   File1.Path = argDir
   j = 0
   List1.Clear
   For i = 0 To File1.ListCount - 1
      If UCase(Mid(File1.List(i), Len(File1.List(i)) - 2, 3)) = "CAR" Then
        j = j + 1
        List1.AddItem File1.Path & "\" & File1.List(i)
      End If
   Next i
   If List1.ListCount > 1 Then
      sCargarArchivosCard5 = ""
   Else
      sCargarArchivosCard5 = argDir
   End If
End Function

