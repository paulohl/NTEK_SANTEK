VERSION 5.00
Begin VB.Form fMostrarCampos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4485
   End
   Begin VB.CommandButton bOK 
      Caption         =   "OK"
      Height          =   405
      Left            =   1110
      TabIndex        =   1
      Top             =   3510
      Width           =   945
   End
   Begin VB.CommandButton bCancela 
      Caption         =   "Cancelar"
      Height          =   405
      Left            =   2460
      TabIndex        =   2
      Top             =   3510
      Width           =   945
   End
End
Attribute VB_Name = "fMostrarCampos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bCancela_Click()
  Modulo.vTemporal1 = ""
  Unload Me
End Sub

Private Sub bOK_Click()
  Modulo.vTemporal1 = ""
  If List1.ListIndex >= 0 Then
    Modulo.vTemporal1 = List1.List(List1.ListIndex)
  End If
  Unload Me
End Sub

Private Sub List1_DblClick()
  Modulo.vTemporal1 = ""
  If List1.ListIndex >= 0 Then
    Modulo.vTemporal1 = List1.List(List1.ListIndex)
  End If
  Unload Me
End Sub
