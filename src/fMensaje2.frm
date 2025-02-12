VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fMensaje2 
   BorderStyle     =   0  'None
   Caption         =   "Información"
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   443
         TabIndex        =   2
         Top             =   420
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
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
         Left            =   250
         TabIndex        =   1
         Top             =   120
         Width           =   5500
      End
   End
End
Attribute VB_Name = "fMensaje2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
