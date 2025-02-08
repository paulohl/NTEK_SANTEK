VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSetupPrinter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   Icon            =   "frmSetupPrinter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame_Seleccionar 
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CommandButton Command3 
         Caption         =   "&Cancelar"
         Height          =   615
         Left            =   480
         Picture         =   "frmSetupPrinter.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Seleccionar 
         Caption         =   "&Seleccionar"
         Height          =   615
         Left            =   480
         Picture         =   "frmSetupPrinter.frx":0FCC
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame_imprimir 
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2295
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   735
         Left            =   600
         Picture         =   "frmSetupPrinter.frx":16CE
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Imprimir"
         Height          =   735
         Left            =   600
         Picture         =   "frmSetupPrinter.frx":1DD0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1935
      Left            =   2520
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
      Begin VB.TextBox txt_NumeroCopias 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "1"
         Top             =   840
         Width           =   375
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         Max             =   20
         Enabled         =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Número de Copias"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   900
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Impresora"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      Begin VB.ComboBox cmb_Impresoras 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmSetupPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Caso As String
Dim lImpresora As Printer

Sub sCargarImpresoras()
 Dim lImpresora_Predeterminada As String
 lImpresora_Predeterminada = Printer.DeviceName
 cmb_Impresoras.Clear
 For Each lImpresora In Printers
   cmb_Impresoras.AddItem lImpresora.DeviceName
 Next
 cmb_Impresoras.Text = lImpresora_Predeterminada
End Sub



Private Sub cmd_Seleccionar_Click()
   For Each lImpresora In Printers
      If lImpresora.DeviceName = cmb_Impresoras.Text Then
         Set Printer = lImpresora
         ''frmPreviewPedidos.Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
         ''frmPrevia.CRViewer1.PrintReport
         Exit For
      End If
   Next
         
   Unload Me
End Sub
Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub Command3_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    sCargarImpresoras
End Sub

Private Sub UpDown1_Change()
txt_NumeroCopias.Text = UpDown1.value
If txt_NumeroCopias.Text <= 0 Then
   txt_NumeroCopias.Text = "1"
End If
End Sub
