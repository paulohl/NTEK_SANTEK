VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVencimiento 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fecha Vencimiento de los Carnets"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   1140
      TabIndex        =   3
      Top             =   540
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   62259201
      CurrentDate     =   40459
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   3000
      TabIndex        =   2
      Top             =   3060
      Width           =   1515
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   435
      Left            =   600
      TabIndex        =   1
      Top             =   3060
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccione la Fecha de Vencimiento que va a llevar este registro"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   4755
   End
End
Attribute VB_Name = "frmVencimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lVencimiento As String
Public lColumna As Integer

Private Sub cmdAceptar_Click()
  fMov.GRID.TextMatrix(fMov.GRID.Row, lColumna) = MonthView1.value
  Unload Me
End Sub

Private Sub cmdCancelar_Click()
  fMov.GRID.TextMatrix(fMov.GRID.Row, lColumna) = UCase(MonthName(Month(Date))) & " " & Year(Date) + 1
  Modulo.fModalResult = Modulo.fModalResultCANCEL
  Unload Me
End Sub

' Select Case GRID.TextMatrix(GRID.Row, Col)
'        Case "2 AÑOS"
'           GRID.TextMatrix(GRID.Row, Col) = UCase(MonthName(Month(Date))) & " " & Year(Date) + 2
'        Case "1 AÑO"
'           GRID.TextMatrix(GRID.Row, Col) = UCase(MonthName(Month(Date))) & " " & Year(Date) + 1
'        Case "6 MESES"
          ' GRID.TextMatrix(GRID.Row, Col) = UCase(MonthName(Month(Date) + 5)) & " " & Year(Date)
'        Case "3 MESES"
'           GRID.TextMatrix(GRID.Row, Col) = UCase(MonthName(Month(Date) + 2)) & " " & Year(Date)
'      End Select
'      ComboVenceListo = True
'   End If

