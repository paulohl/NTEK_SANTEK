VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmColumnasExportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecione Las Columnas a Exportar"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   2580
      TabIndex        =   2
      Top             =   3060
      Width           =   1335
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   435
      Left            =   840
      TabIndex        =   1
      Top             =   3060
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListVExportar 
      Height          =   2655
      Left            =   780
      TabIndex        =   0
      Top             =   180
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Columnas"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nro"
         Object.Width           =   1058
      EndProperty
   End
End
Attribute VB_Name = "frmColumnasExportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Sub sExportar(dgExportar As DataGrid)
On Error GoTo ErrorExcel

Dim objExcel As Excel.Application
Dim HNom As Integer 'Horizontal
Dim VNom As Integer 'Vertical
Dim Hdatos As Integer 'Horizontal
Dim Vdatos As Integer 'Vertical
Dim cuentaNombres As Integer
Dim cuentadatos As Integer
Dim i As Integer
Dim n As Integer
Dim j As Integer

If fPersonasAct.Adodc1.Recordset.RecordCount <> 0 Then
'Crear un objeto del tipo excel.application
Set objExcel = New Excel.Application
objExcel.Visible = True
objExcel.SheetsInNewWorkbook = 1
objExcel.Workbooks.Add

'PONER UN TITULO
objExcel.ActiveSheet.Cells(1, 1) = fPersonasAct.cCP.Text
With objExcel.ActiveSheet.Cells(1, 1).Font
.Color = vbBlack
.Size = 9
.Bold = True
End With
'AHORA UN SUBTITULO
objExcel.ActiveSheet.Cells(3, 1) = "LISTADO DE CLIENTES TABLA " & fPersonasAct.lTablas.List(fPersonasAct.lTablas.ListIndex)
With objExcel.ActiveSheet.Cells(3, 1).Font
.Color = vbBlack
.Size = 9
.Bold = True
End With

'UTILIZAMOS LAS VARIABLES PARA LA UBICACION DE NUESTROS TEXTOS
HNom = 1
VNom = 7
Vdatos = 8
Hdatos = 1

cuentaNombres = fPersonasAct.Adodc1.Recordset.Fields.Count
cuentadatos = fPersonasAct.Adodc1.Recordset.RecordCount

'AGREGAMOS LOS REGISTROS (RECUERDEN QUE NO IMPORTA CUANTAS COLUMNAS O REGISTROS TENGAMOS EL BUCLE_
'FUNCIONA SEGUN EL NUMERO DE CABECERAS Y REGISTROS
For i = 1 To ListVExportar.ListItems.Count      ''(cuentaNombres - 1)
 If ListVExportar.ListItems(i).Checked = True Then
   objExcel.ActiveSheet.Cells(VNom, HNom) = ListVExportar.ListItems(i).Text  'rstFacturas.Fields(i).Name
   objExcel.ActiveSheet.Range(objExcel.ActiveSheet.Cells(VNom, HNom), objExcel.ActiveSheet.Cells(VNom, HNom)).HorizontalAlignment = xlHAlignCenterAcrossSelection
   With objExcel.ActiveSheet.Cells(VNom, HNom).Font
   .Size = 9
   .Color = vbRed
   .Bold = True
   End With

   fPersonasAct.Adodc1.Recordset.MoveFirst
   For n = 1 To fPersonasAct.Adodc1.Recordset.RecordCount
      objExcel.ActiveSheet.Cells(Vdatos, Hdatos) = fPersonasAct.Adodc1.Recordset.Fields(i - 1).value
      objExcel.ActiveSheet.Cells(Vdatos, Hdatos).Font.Size = 8
      Vdatos = Vdatos + 1
      fPersonasAct.Adodc1.Recordset.MoveNext
   Next
  HNom = HNom + 1
  Hdatos = Hdatos + 1
  Vdatos = 8
  fPersonasAct.Adodc1.Recordset.MoveFirst
 End If
Next i
'AHORA LE ASIGNAMOS UN TAMAÑO A CADA COLUMNA SEGUN NESECITEMOS
objExcel.Columns("B").ColumnWidth = 19.43
objExcel.Columns("C").ColumnWidth = 19.43
objExcel.Columns("D").ColumnWidth = 16.86
objExcel.Columns("E").ColumnWidth = 10.83
End If
Exit Sub
ErrorExcel:
MsgBox Err.Description
End Sub

Private Sub cmdExportar_Click()
   sExportar fPersonasAct.DataGrid1
End Sub
