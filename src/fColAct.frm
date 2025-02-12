VERSION 5.00
Begin VB.Form fColAct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualización de Datos"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   500
      Left            =   2430
      Picture         =   "fColAct.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2730
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton bAceptar 
      Caption         =   "Aceptar"
      Height          =   500
      Left            =   1290
      Picture         =   "fColAct.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2730
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Caption         =   "Indique los Campos que desee Actualizar "
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   150
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   270
         Width           =   4245
      End
   End
End
Attribute VB_Name = "fColAct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bAceptar_Click()
  If Not Modulo.Hay_Seleccion(List1) = True Then
    MsgBox "Debe Seleccionar al menos un Campo para Actualizar...", vbCritical, "Información"
    List1.SetFocus
  Else
    Modulo.vTemporal1 = "OK"
    Me.Hide
  End If
End Sub

Private Sub bCancelar_Click()
  Modulo.vTemporal1 = ""
  Me.Hide
End Sub
