VERSION 5.00
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Begin VB.Form Form1 
   Caption         =   "Ejemplo"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   Begin ubGridControl.ubGrid ubGrid1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7223
      Rows            =   1
      Cols            =   5
      Redraw          =   -1  'True
      ShowGrid        =   -1  'True
      GridSolid       =   -1  'True
      GridLineColor   =   12632256
      UseBackColorAlt =   -1  'True
      BackColorFixed  =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontEdit {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoNewRow      =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

' configura los encabezdos
ubGrid1.AutoSetup 20, 10, True, True, "Sin máscara     |Solo mayusculas     |Solo números     |Solo fechas     |Drop Down     |Checkbox     |Columna no editable     "


'For cnt = 1 To 10
'  ubGrid1.TextMatrix(cnt, 0) = cnt
'Next

' para la columna que tiene el DropDown
ubGrid1.AddLookup 5, "Item 1", "data1"
ubGrid1.AddLookup 5, "Item 2", "data2"
ubGrid1.AddLookup 5, "Item 3", "data3"
ubGrid1.AddLookup 5, "Item 4", "data4"
ubGrid1.AddLookup 5, "Item 5", 5

'ubGrid1.RemoveLookup 5

' establece las máscaras para las columnas
ubGrid1.ColMask(2) = 1
ubGrid1.ColMask(3) = 2
ubGrid1.ColMask(4) = 3
ubGrid1.ColMask(6) = 4

' Columna 7 sin edicion
ubGrid1.ColAllowEdit(7) = False
ubGrid1.AddButton 8

ubGrid1.ColEditWidth(1) = 5

ubGrid1.AutoRedraw = False

' Agrega elementos copn el método textMAtrix
For Row = 1 To 10
  ubGrid1.TextMatrix(Row, 1) = "Abcd" & Row
  ubGrid1.TextMatrix(Row, 2) = "ABCD" & Row
  ubGrid1.TextMatrix(Row, 3) = "12345" & Row
  ubGrid1.TextMatrix(Row, 4) = Date + Row
Next

ubGrid1.AutoRedraw = True

End Sub

'Dim txt As String

Private Sub ubGrid1_BtnClick(ByVal Row As Long, ByVal Col As Long)
  'txt = Me.ubGrid1.TextMatrix(Row, 1)
  'Me.Hide

End Sub

'Private Sub ubGrid1_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
'  If KeyCode = 46 Then
'    ubGrid1.RemoveItem ubGrid1.Row
'  End If
'
'End Sub

'Private Sub ubGrid1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
'  If ubGrid1.MouseRow > 0 And ubGrid1.MouseCol > 0 Then
'    ubGrid1.ToolTipText = ubGrid1.TextMatrix(ubGrid1.MouseRow, ubGrid1.MouseCol)
'  End If
'End Sub
Private Sub ubGrid1_Click()
Me.Caption = ubGrid1.Text
End Sub
