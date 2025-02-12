VERSION 5.00
Begin VB.Form fAbrirTablaMDB 
   Caption         =   "Abrir Tabla de .MDB"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   3315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   500
      Left            =   1800
      Picture         =   "fAbrirTablaMDB.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton bAceptar 
      Caption         =   "Aceptar"
      Height          =   500
      Left            =   660
      Picture         =   "fAbrirTablaMDB.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione la Tabla que va Utilizar"
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3285
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2160
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   3075
      End
   End
End
Attribute VB_Name = "fAbrirTablaMDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bAceptar_Click()
  If List1.ListIndex >= 0 Then
    Modulo.vTemporal1 = List1.List(List1.ListIndex)
  Else
    MsgBox "Debe Marcar(Seleccionar) la Tabla...", vbCritical, "Información"
    Exit Sub
  End If
  Modulo.fModalResult = Modulo.fModalResultOK
  Unload Me
End Sub

Private Sub bCancelar_Click()
  Modulo.fModalResult = Modulo.fModalResultCANCEL
  Unload Me
End Sub


Private Sub List1_DblClick()
bAceptar_Click
End Sub
