VERSION 5.00
Begin VB.Form fIndicaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indicaciones Especiales del Cliente"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6990
      Begin VB.CommandButton bAceptar 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   500
         Left            =   2995
         Picture         =   "fIndicaciones.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5370
         UseMaskColor    =   -1  'True
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   5085
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "fIndicaciones.frx":058A
         Top             =   210
         Width           =   6735
      End
   End
End
Attribute VB_Name = "fIndicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bAceptar_Click()
  Unload Me
End Sub
