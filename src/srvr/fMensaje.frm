VERSION 5.00
Begin VB.Form fMensaje 
   BorderStyle     =   0  'None
   Caption         =   "Información"
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7700
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Procesando, Espere..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   7425
      End
   End
End
Attribute VB_Name = "fMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
