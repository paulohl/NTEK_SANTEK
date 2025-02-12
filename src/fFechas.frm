VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form fFechas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3900
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3900
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSinVencimiento 
      Caption         =   "Sin Vencimiento"
      CausesValidation=   0   'False
      Height          =   500
      Left            =   2220
      Picture         =   "fFechas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1260
      UseMaskColor    =   -1  'True
      Width           =   1320
   End
   Begin VB.CommandButton bOK 
      Caption         =   "OK"
      CausesValidation=   0   'False
      Height          =   500
      Left            =   300
      Picture         =   "fFechas.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1260
      UseMaskColor    =   -1  'True
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3885
      Begin MSComCtl2.DTPicker eFechaVence 
         Height          =   345
         Left            =   1103
         TabIndex        =   1
         Top             =   540
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57081857
         CurrentDate     =   40028
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Vencimiento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   990
         TabIndex        =   2
         Top             =   210
         Width           =   1920
      End
   End
End
Attribute VB_Name = "fFechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bOK_Click()
  Modulo.vTemporal1 = Format(eFechaVence.value, "dd/mm/yyyy")
  Modulo.fModalResult = Modulo.fModalResultOK
  Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSinVencimiento_Click()
  If MsgBox("¿Está seguro que desea dejar el formato sin fecha de vencimiento?", vbYesNo + vbQuestion) = vbYes Then
     'Modulo.vTemporal1 = Format(eFechaVence.value, "dd/mm/yyyy")
     Modulo.fModalResult = Modulo.fModalResultOK
     Unload Me
  End If
End Sub

Private Sub Form_Load()
  eFechaVence.value = Date
End Sub
