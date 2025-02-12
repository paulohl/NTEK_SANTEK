VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form fImportar 
   Caption         =   "Importar Datos de Personas desde EXCEL"
   ClientHeight    =   10125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10125
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   405
      Left            =   3360
      TabIndex        =   20
      Top             =   9600
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton bGenerarArchivoErrores 
      Caption         =   "Generar Archivo .XLS de Errores"
      Height          =   405
      Left            =   9360
      TabIndex        =   17
      Top             =   9000
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Archivo .XLS a importar:"
      Height          =   825
      Left            =   7050
      TabIndex        =   12
      Top             =   0
      Width           =   8150
      Begin VB.CommandButton bExaminar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Examinar"
         Top             =   330
         Width           =   435
      End
      Begin VB.TextBox Text1 
         Height          =   550
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "fLeer.frx":0000
         Top             =   210
         Width           =   7050
      End
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "CANCELAR"
      Height          =   500
      Left            =   7800
      Picture         =   "fLeer.frx":0006
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9540
      UseMaskColor    =   -1  'True
      Width           =   1000
   End
   Begin VB.CommandButton bAceptar 
      Caption         =   "GUARDAR"
      Height          =   500
      Left            =   6480
      Picture         =   "fLeer.frx":0590
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9540
      UseMaskColor    =   -1  'True
      Width           =   1000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6210
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contenido del Archivo XLS"
      Height          =   8625
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   15195
      Begin VB.ListBox ListCols 
         Height          =   1035
         Left            =   5700
         TabIndex        =   21
         Top             =   7290
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ListBox lRepetidas 
         Height          =   1425
         Left            =   1200
         TabIndex        =   19
         Top             =   6420
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.ListBox lFilas 
         Height          =   1230
         Left            =   3000
         TabIndex        =   18
         Top             =   6420
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CommandButton bLeerXLS 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Leer Datos desde XLS"
         Height          =   345
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   8220
         UseMaskColor    =   -1  'True
         Width           =   2040
      End
      Begin VB.CommandButton bPegar 
         Caption         =   "PEGAR Campos"
         Height          =   375
         Left            =   13500
         TabIndex        =   5
         Top             =   8190
         Width           =   1545
      End
      Begin VB.CommandButton bCopiar 
         Caption         =   "COPIAR Campos"
         Height          =   375
         Left            =   11970
         TabIndex        =   6
         Top             =   8190
         Width           =   1545
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   7905
         Left            =   60
         TabIndex        =   2
         Top             =   180
         Width           =   15060
         _ExtentX        =   26564
         _ExtentY        =   13944
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         FillStyle       =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0000"
         Height          =   285
         Left            =   4740
         TabIndex        =   11
         Top             =   8250
         Width           =   555
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Filas de Datos Leidas:"
         Height          =   285
         Left            =   3060
         TabIndex        =   10
         Top             =   8250
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7035
      Begin VB.CommandButton bPreparar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Preparar Archivo de Datos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   220
         Width           =   3225
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "==>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2670
         TabIndex        =   16
         Top             =   330
         Width           =   330
      End
      Begin VB.Label Label4 
         Caption         =   "Para Crear Listado de Personas la Primera Vez."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   15
         Top             =   180
         Width           =   2445
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Archivo .XLS a importar:"
      Height          =   195
      Left            =   8130
      TabIndex        =   9
      Top             =   240
      Width           =   1710
   End
End
Attribute VB_Name = "fImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' -- Variables para acceder a la hoja excel
Private obj_Excel       As Object
Private obj_Workbook    As Object
Private obj_Worksheet   As Object

Dim aCampos() As String
Dim lErrores As Boolean

' ----------------------------------------------------------------------------------
' \\ -- Función para leer los datos del Excel y cargarlos en el Flex
' ----------------------------------------------------------------------------------
Private Function Estampar_Cabecera_Excel_FlexGrid(ByRef sCampos() As String, iCampos As Integer, sPath As String, FlexGrid As Object, Optional sSheetName As String = vbNullString) As Boolean
    Dim i As Long
    Dim n As Long
    Dim HayDatos As Boolean
    Dim j As Long
    Dim s1 As String
    'Dim excel As New excel.Worksheet
    'excel.Cells.Font.ColorIndex
    On Error GoTo error_sub
    
    Estampar_Cabecera_Excel_FlexGrid = False
    
    ' -- Comproba si existe l archivo
    If Len(Dir(sPath)) = 0 Then
       MsgBox "No se ha encontrado el archivo: " & sPath, vbCritical
       Exit Function
    End If
    
    Me.MousePointer = vbHourglass
    ' -- crea rnueva instancia de Excel
    Set obj_Excel = CreateObject("Excel.Application")
    'obj_Excel.Visible = True
    
    ' -- Abrir el libro
    Set obj_Workbook = obj_Excel.Workbooks.Open(sPath)
    ' -- referencia la Hoja, por defecto la hoja activa
    If sSheetName = vbNullString Then
        Set obj_Worksheet = obj_Workbook.ActiveSheet
    Else
        Set obj_Worksheet = obj_Workbook.Sheets(sSheetName)
    End If
    
    HayDatos = False
    i = 1
    Do While i < iCampos And Not HayDatos
      s1 = obj_Worksheet.Cells(1, i).value
      If Trim(s1) <> "" Then
        HayDatos = True
      Else
        i = i + 1
      End If
    Loop
    
    If HayDatos Then
      MsgBox "No se puede grabar la cabecera de la tabla porque el " & vbCrLf & _
             "Archivo EXCEL posee datos en Celdas de la primera Fila...", vbCritical, "Información"
      Call Descargar
      Me.MousePointer = vbDefault
      Estampar_Cabecera_Excel_FlexGrid = False
      Exit Function
      
    Else
      j = 1
      For i = 0 To iCampos - 1
        If Trim(sCampos(i)) <> "" Then
          obj_Worksheet.Cells(1, j) = sCampos(i)  'Escribe los campos en las celdas del XLS!
          j = j + 1
        End If
      Next i
      
      '-- Guardar los cambios al libro (22/06/09)
      obj_Workbook.Save
    End If
    ' -- Cerrar libro
    obj_Workbook.Close
    ' -- Cerrar Excel
    obj_Excel.Quit
    ' -- Descargar objetos para liberar recursos
    Call Descargar
' -- Errores
    Estampar_Cabecera_Excel_FlexGrid = True
Exit Function
error_sub:
    MsgBox Err.Description
    Call Descargar
    Me.MousePointer = vbDefault
    Estampar_Cabecera_Excel_FlexGrid = False
End Function


Private Function Leer_Excel_FlexGrid(sPath As String, FlexGrid As Object, Optional sSheetName As String = vbNullString) As Boolean
    Dim i As Long
    Dim n As Long
    Dim HayDatos As Boolean
    Dim j As Long
    Dim s1 As String
    Dim bRecorrer As Boolean, bEnBlanco As Boolean
    
    Dim m As Integer
    Dim bHayDato As Boolean
    Dim Fila As Integer
    
    On Error GoTo error_sub
    
    Leer_Excel_FlexGrid = False
    
    ' -- Comproba si existe l archivo
    If Len(Dir(sPath)) = 0 Then
       MsgBox "No se ha encontrado el archivo: " & sPath, vbCritical
       Exit Function
    End If
    
    Load fMensaje
    fMensaje.Label1.Caption = "Leyendo en archivo .XLS"
    fMensaje.Show
    DoEvents
    
    
    
    Me.MousePointer = vbHourglass
    ' -- crea rnueva instancia de Excel
    Set obj_Excel = CreateObject("Excel.Application")
    'obj_Excel.Visible = True
    
    ' -- Abrir el libro
    Set obj_Workbook = obj_Excel.Workbooks.Open(sPath)
    ' -- referencia la Hoja, por defecto la hoja activa
    If sSheetName = vbNullString Then
        Set obj_Worksheet = obj_Workbook.ActiveSheet
    Else
        Set obj_Worksheet = obj_Workbook.Sheets(sSheetName)
    End If
    
    '--Contar las Filas para ajustar el FLEXGRID "dinamicamente"
    i = 1
    Fila = 1
    '--Contar hasta que TODAS las columnas esten en "blancos":
    '--Es decir una FILA completa en "blanco"
    j = 1
    bRecorrer = True
    Do While bRecorrer    'parecido al .T.
      j = 1
      bHayDato = False
      Do While j < 256 And Not bHayDato '--Asumir 256 columnas para registro en blanco
        s1 = obj_Worksheet.Cells(i, j).value
        If Trim(s1) <> "" Then bHayDato = True
        j = j + 1
      Loop
      
      If bHayDato Then  'Es Fila de Datos
        If Fila > 1 Then MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        Fila = Fila + 1
        For m = 0 To MSFlexGrid1.Cols - 1
          s1 = obj_Worksheet.Cells(i + 1, m + 1).value
          MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, m) = s1
        Next m
        i = i + 1
        fMensaje.Label1.Caption = "Leyendo en archivo .XLS / Fila " & CStr(Fila)
      Else
        bRecorrer = False 'Es Fila Vacia COMPLETAMENTE, se asume FINAL DEL LISTADO.
      End If
    Loop
      
    ' -- Cerrar libro
    obj_Workbook.Close
    ' -- Cerrar Excel
    obj_Excel.Quit
    ' -- Descargar objetos para liberar recursos
    Call Descargar
' -- Errores
    Leer_Excel_FlexGrid = True
Exit Function
error_sub:
    MsgBox Err.Description
    Call Descargar
    Me.MousePointer = vbDefault
    Leer_Excel_FlexGrid = False
End Function





' ----------------------------------------------------------------------------------
' \\ -- Función para leer los datos del Excel y cargarlos en el Flex
' ----------------------------------------------------------------------------------
Private Sub Excel_FlexGrid(sPath As String, FlexGrid As Object, Filas As Integer, Columnas As Integer, Optional sSheetName As String = vbNullString)

    Dim i As Long
    Dim n As Long
    
    On Error GoTo error_sub
    ' -- Comproba si existe l archivo
    If Len(Dir(sPath)) = 0 Then
       MsgBox "No se ha encontrado el archivo: " & sPath, vbCritical
       Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    ' -- crea rnueva instancia de Excel
    Set obj_Excel = CreateObject("Excel.Application")
    'obj_Excel.Visible = True
    
    ' -- Abrir el libro
    Set obj_Workbook = obj_Excel.Workbooks.Open(sPath)
    ' -- referencia la Hoja, por defecto la hoja activa
    If sSheetName = vbNullString Then
        Set obj_Worksheet = obj_Workbook.ActiveSheet
    Else
        Set obj_Worksheet = obj_Workbook.Sheets(sSheetName)
    End If
    
    ' -- Setear Grid
    With MSFlexGrid1
        ' -- Especificar  la cantidad de filas y columnas
        .Cols = Columnas
        .Rows = Filas
        ' -- Recorrer las filas del FlexGrid para agregar los datos
        For i = 0 To .Rows - 1
            ' -- Establecer la fila activa
            .Row = i
            ' -- Recorrer las columnas del FlexGrid
            For n = 0 To .Cols - 1
                ' -- Establecer columna activa
                .Col = n
                ' -- Asignar a la celda del Flex el contenido de la celda del excel
                .Text = IIf(i = 0, CStr(n) & " - ", "") & obj_Worksheet.Cells(i + 1, n + 1).value
                
                'If (i = 0) Then lColumnas.AddItem .Text
            Next
        Next
        
        
    End With
    
    ' -- Cerrar libro
    obj_Workbook.Close
    ' -- Cerrar Excel
    obj_Excel.Quit
    ' -- Descargar objetos para liberar recursos
    Call Descargar
' -- Errores
Exit Sub
error_sub:
    MsgBox Err.Description
    Call Descargar
    Me.MousePointer = vbDefault
End Sub
' ----------------------------------------------------------------------------------
' \\ -- Función para descargar los objetos
' ----------------------------------------------------------------------------------
Private Sub Descargar()
    On Local Error Resume Next
    Set obj_Workbook = Nothing
    Set obj_Excel = Nothing
    Set obj_Worksheet = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Function Guardar_ERRORES_FlexGrid_Excel(sPath As String, ByRef FlexGrid As MSFlexGrid, Optional sSheetName As String = vbNullString) As Boolean
    Dim i As Long, j As Long
    Dim n As Long
    Dim HayDatos As Boolean
    Dim s1 As String
    
    On Error GoTo error_sub
    
    Guardar_ERRORES_FlexGrid_Excel = False
    
    ' -- Comproba si existe l archivo
    If Len(Dir(sPath)) = 0 Then
       MsgBox "No se ha encontrado el archivo: " & sPath, vbCritical
       Exit Function
    End If
    
    Me.MousePointer = vbHourglass
    ' -- crea rnueva instancia de Excel
    Set obj_Excel = CreateObject("Excel.Application")
    'obj_Excel.Visible = True
    
    ' -- Abrir el libro
    Set obj_Workbook = obj_Excel.Workbooks.Open(sPath)
    ' -- referencia la Hoja, por defecto la hoja activa
    If sSheetName = vbNullString Then
        Set obj_Worksheet = obj_Workbook.ActiveSheet
    Else
        Set obj_Worksheet = obj_Workbook.Sheets(sSheetName)
    End If
        
    For i = 0 To FlexGrid.Rows - 1
    
      For j = 0 To FlexGrid.Cols - 1
      
        obj_Worksheet.Cells(i + 1, j + 1) = FlexGrid.TextMatrix(i, j)
        FlexGrid.Col = j
        FlexGrid.Row = i
        If FlexGrid.CellBackColor = vbRed Then
           obj_Worksheet.Cells(i + 1, j + 1).Font.Color = 255
        Else
           obj_Worksheet.Cells(i + 1, j + 1).Font.Color = 0
        End If
      Next j
      obj_Worksheet.Columns().AutoFit
      Dim hoja As New excel.Worksheet
      'hoja.au
      'FlexGrid.Row = i
      'FlexGrid.Col = j - 1
      'If FlexGrid.CellBackColor = vbRed Then
      '  obj_Worksheet.Cells(i + 1, FlexGrid.Cols + 1) = "*"
      'End If
      
    Next i
    
'    'Introducir los Errores:
'    obj_Worksheet.Cells(1, FlexGrid.Cols + 1) = "ERROR"
'
'    'Determinar si la fila tiene errores, adjuntarlos al Excel:
'    For i = 1 To FlexGrid.Rows - 1
'      FlexGrid.Row = i
'      FlexGrid.Col = 0
'      If FlexGrid.CellBackColor = vbRed Then
'        c = FlexGrid.Cols + 1
'        'Buscarlo en la lista de Logs:
'        For j = 0 To lLog.ListCount - 1
'          s = lLog.List(j)
'          p = InStr(s, "Nº")  'Fila Nº
'          If p > 0 Then
'            s1 = Trim(Mid(s, p + 2)) 'Toma solo la parte del numero de Fila:
'            If IsNumeric(s1) Then
'              f = CLng(s1)
'              If f = i Then 'El mensaje de Error es de la misma fila:
'
'                obj_Worksheet.Cells(i + 1, c) = s
'                c = c + 1
'
'              End If
'            End If
'          End If
'        Next j
'      End If
'    Next i
           
      
    '-- Guardar los cambios al libro (22/06/09)
    obj_Workbook.Save
    ' -- Cerrar libro
    obj_Workbook.Close
    ' -- Cerrar Excel
    obj_Excel.Quit
    ' -- Descargar objetos para liberar recursos
    Call Descargar
' -- Errores
    Guardar_ERRORES_FlexGrid_Excel = True
Exit Function
error_sub:
    MsgBox Err.Description
    Call Descargar
    Me.MousePointer = vbDefault
    Guardar_ERRORES_FlexGrid_Excel = False
End Function


'''Private Sub bAbrir_Click()
'''  Dim a As String
'''  Dim l As String
'''  Dim aCampos() As String
'''  Dim i As Integer, k As Integer, j As Integer
'''  Dim sNom As String, sOrigen As String
'''  Dim Fila As Integer
'''  Dim bHayCedula As Boolean
'''
'''  bHayCedula = False
'''  If Trim(Text1.Text) <> "" Then
'''    a = Trim(Text1.Text)
'''    sOrigen = a
'''    'Call Excel_FlexGrid(a, MSFlexGrid1, 1000, 20, "Sheet1")
'''    Load fMensaje
'''    fMensaje.Label1.Caption = "Abriendo Documento EXCEL, Espere..."
'''    fMensaje.Show
'''
'''    'Introducir los campos en un arreglo:
'''    k = fPersonasAct.Adodc1.Recordset.Fields.Count
'''    If k > 0 Then
'''      ReDim aCampos(k)
'''      For i = 0 To k - 1
'''        aCampos(i) = ""
'''        If fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "ID" And _
'''          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "FOTO" And _
'''          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "FECHA" And _
'''          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "CONTADOR" And _
'''          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "TIENE_FOTO" And _
'''          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "CREACION" Then
'''          aCampos(i) = fPersonasAct.Adodc1.Recordset.Fields(i).Name
'''        End If
'''      Next i
'''
'''      If Estampar_Cabecera_Excel_FlexGrid(aCampos(), k, a, MSFlexGrid1) = True Then
'''        '-- Anexarle la cadena "OFICINA" al nombre del archivo XLS:
'''        sNom = Modulo.ExtraerFilePath(a, "\")
'''        If InStr(UCase(sNom), ".XLS") > 0 Then sNom = Mid(sNom, 1, InStr(UCase(sNom), ".XLS") - 1)
'''        If InStr(sNom, "_OFICINA") <= 0 Then
'''          sNom = GetPath(a, "\") & sNom & "_OFICINA.XLS"
'''          FileCopy sOrigen, sNom
'''          MsgBox "Operación Efectuada Correctamente..." & vbCrLf & _
'''                 "Archivo excel copiado con el nombre:" & vbCrLf & _
'''                 "[" & sNom & "]" & vbCrLf, vbInformation, "Información"
'''
'''          MSFlexGrid1.Clear
'''          MSFlexGrid1.Rows = 2
'''          MSFlexGrid1.FixedRows = 1
'''          MSFlexGrid1.Cols = 1
'''          Fila = 1
'''          j = 0
'''          For i = 0 To k - 1
'''            If Trim(aCampos(i)) <> "" Then
'''              MSFlexGrid1.TextMatrix(0, j) = aCampos(i)
'''              If UCase(aCampos(i)) = "CEDULA" Then bHayCedula = True
'''              MSFlexGrid1.Cols = MSFlexGrid1.Cols + 1
'''              j = j + 1
'''            End If
'''          Next i
'''
'''          If bHayCedula Then '--Lleva FOTO!
'''            'MSFlexGrid1.Cols = MSFlexGrid1.Cols + 1
'''            MSFlexGrid1.TextMatrix(0, MSFlexGrid1.Cols - 1) = "FOTO"
'''          Else
'''            MSFlexGrid1.Cols = MSFlexGrid1.Cols - 1
'''          End If
'''
'''
'''          Text1.Text = a
'''
'''        End If
'''      End If
'''
'''    Else
'''      MsgBox "No Hay Campos para asignar en XLS.", vbCritical, "Información"
'''    End If
'''
'''    'Enlace_Estimado
'''    Unload fMensaje
'''  Else
'''    MsgBox "Debe Indicar el archivo en formato EXCEL a procesar, Revise...", vbCritical, "Información"
'''  End If
'''End Sub

Private Sub bAceptar_Click()
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim st As String
  Dim i As Integer, j As Integer
  Dim cr As Long, agregando As Boolean
  Dim sCampo As String, bHayUpdate As Boolean
  Dim sSQL As String, sValores As String, sSQL2 As String
  Dim db_campo As String
  Dim flx_campo As String
  Dim Filas As Integer
  Dim sV As String
  Dim bHayCedula As Boolean, bSePuedeAgregar As Boolean
  Dim sValorCedula As String
  Dim sCedExistencia As String
  Dim bExisteCedReg As Boolean
  Dim iCedExistencia As Integer
  
  Dim ii As Integer, jj As Integer
  
  Dim sCampoactualizar As String
  If lErrores = True Then
     If MsgBox("Los Datos contienen al menos un error. ¿Desea continuar de todas formas?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
  End If
  
  bHayCedula = False
  
  bExisteCedReg = False
  sCedExistencia = ""
  iCedExistencia = 0
  
  If fPersonasAct.lTablas.ListIndex >= 0 Then
    st = fPersonasAct.lTablas.List(fPersonasAct.lTablas.ListIndex)
    
    If MsgBox("Se agregarán los [" & CStr(MSFlexGrid1.Rows - 1) & "] registros a la Base de datos del Cliente." & vbCrLf & _
              "¿Está Seguro de Iniciar la Operación?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
              
      '--Cargar la data de personas del cliente:
      s = "select * from [" & st & "]"
      r.Open s, Modulo.DBConexionSQL, adOpenKeyset, adLockReadOnly
      
      cr = Modulo.Total_Registros(st)
      
      If cr > 0 Then
        s = "Se ha detectado que ya existen [" & CStr(cr) & "] personas registradas" & vbCrLf & _
            "en la base de datos del cliente. " & vbCrLf
            
        bHayCedula = Modulo.EXISTE_CAMPO(r, "CEDULA")
        
        If bHayCedula Then
          s = s & " Se tomará en cuenta la [CEDULA] como campo clave exclusivo."
        End If
        
        MsgBox s, vbCritical, "Atención"
        
        Load fColAct
        With fColAct
          .List1.Clear
          For i = 0 To MSFlexGrid1.Cols - 1
            .List1.AddItem MSFlexGrid1.TextMatrix(0, i)
          Next i
        End With
        Modulo.vTemporal1 = ""
        fColAct.Show vbModal
        
        ListCols.Clear
        For i = 0 To fColAct.List1.ListCount - 1
          If fColAct.List1.Selected(i) Then
            ListCols.AddItem fColAct.List1.List(i)
          End If
        Next i
        
        Unload fColAct
                             
        
      End If
      
      Load fMensaje
      
      sCedExistencia = ""
      
      For Filas = 1 To MSFlexGrid1.Rows - 1
      
        bSePuedeAgregar = True
              
        If Not Fila_En_Blanco_En_FlexGrid(MSFlexGrid1, Filas) Then
      
          fMensaje.Label1.Caption = "Añadiendo Fila [" & CStr(Filas) & "] / [" & CStr(MSFlexGrid1.Rows) & "] , Espere..."
          fMensaje.Show
          DoEvents
        
          'If cr <= 0 Then
            'Agregación en Lote!
            sSQL = "insert into [" & st & "] ("
            sValores = ""
              
            For i = 0 To fPersonasAct.Adodc1.Recordset.Fields.Count - 1
            
              db_campo = UCase(fPersonasAct.Adodc1.Recordset.Fields(i).Name)
              If db_campo <> "ID" And db_campo <> "FECHA" Then
                sSQL = sSQL & db_campo & ","
                  
                If Modulo.EXISTE_CAMPO_EN_FLEXGRID(MSFlexGrid1, db_campo) Then
                  
                  'Select Case fPersonasAct.Adodc1.Recordset.Fields(i).Type
                  '  Case adChar: s = "CHAR"
                  '  Case adInteger: s = "INTEGER"
                  '  Case adDBTimeStamp: s = "DATETIME"
                  '  Case adDouble: s = "FLOAT"
                  'End Select
                  
                  If fPersonasAct.Adodc1.Recordset.Fields(i).Type = adChar Or _
                     fPersonasAct.Adodc1.Recordset.Fields(i).Type = adDBDate Then
                    sValores = sValores & "'" & MSFlexGrid1.TextMatrix(Filas, Modulo.Nro_Columna_FlexGrid(MSFlexGrid1, db_campo)) & "',"
                  Else
                    If fPersonasAct.Adodc1.Recordset.Fields(i).Type = adInteger Or _
                       fPersonasAct.Adodc1.Recordset.Fields(i).Type = adDouble Then
                      sValores = sValores & " " & MSFlexGrid1.TextMatrix(Filas, Modulo.Nro_Columna_FlexGrid(MSFlexGrid1, db_campo)) & " ,"
                    End If
                  End If
                End If
              End If
            Next i
              
            '--Completar los Predeterminados:
            If Modulo.EXISTE_CAMPO(fPersonasAct.Adodc1.Recordset, "TIENE_FOTO") Then
              sValores = sValores & "'N',"
            End If
            
            If Modulo.EXISTE_CAMPO(fPersonasAct.Adodc1.Recordset, "MARCA") Then
              sValores = sValores & "'I',"
            End If
            
            
            'If Modulo.EXISTE_CAMPO(fPersonasAct.Adodc1.Recordset, "FECHA") Then
            '  sValores = sValores & "'" & Format(Date, "yyyymmdd hh:mm:ss") & "',"
            'End If
              
            If Modulo.EXISTE_CAMPO(fPersonasAct.Adodc1.Recordset, "CONTADOR") Then
              sValores = sValores & "0,"
            End If
             
            If Modulo.EXISTE_CAMPO(fPersonasAct.Adodc1.Recordset, "CREACION") Then
              sValores = sValores & "'" & Format(Date, "yyyymmdd hh:mm:ss") & "',"
            End If
              
            '--Finalizar el comando SQL:
              
            If Mid(sSQL, Len(sSQL), 1) = "," Then Mid(sSQL, Len(sSQL), 1) = " "
            If Mid(sValores, Len(sValores), 1) = "," Then Mid(sValores, Len(sValores), 1) = " "
            sSQL = sSQL & ") VALUES (" & sValores & ")"
            
            '--------------------------------------------------------------
            '--Antes de agregarlo, verificar que NO exista, si hay CEDULA:
            '--------------------------------------------------------------
            If bHayCedula And cr > 0 Then
              sValorCedula = MSFlexGrid1.TextMatrix(Filas, Modulo.Nro_Columna_FlexGrid(MSFlexGrid1, "CEDULA"))
              s = "Cedula = '" & sValorCedula & "'"
              If cr > 0 Then r.MoveFirst
              r.Find s
              If r.EOF Then bExisteCedReg = False Else bExisteCedReg = True
            End If
            
            If bExisteCedReg Then
              'Actualizar solo las columnas seleccionadas:
              sSQL2 = "update [" & st & "] set "
              
              For ii = 0 To ListCols.ListCount - 1
                sCampoactualizar = ListCols.List(ii)
                
                sSQL2 = sSQL2 & sCampoactualizar & " = '"
                
                For jj = 0 To MSFlexGrid1.Cols - 1
                  If UCase(sCampoactualizar) = MSFlexGrid1.TextMatrix(0, jj) Then
                    sSQL2 = sSQL2 & MSFlexGrid1.TextMatrix(Filas, jj) & "'"
                  End If
                Next jj
                
                If ListCols.ListCount > 1 Then sSQL2 = sSQL2 & ","
              Next ii
              
              sSQL2 = sSQL2 & " MARCA = 'I' WHERE CEDULA = '" & sValorCedula & "'"
                
              'sSQL2 = "DELETE FROM [" & st & "] WHERE CEDULA = '" & sValorCedula & "'"
              On Error Resume Next
              Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
              Modulo.DBComandoSQL.CommandText = sSQL2
              Modulo.DBComandoSQL.Execute
              If Err.Number <> 0 Then
                MsgBox "Error: Imposible Actualizar el Registro..." & vbCrLf & Err.Description, vbCritical, "Información"
              End If
              
              sCedExistencia = sCedExistencia & sValorCedula & "  "
              iCedExistencia = iCedExistencia + 1
            Else
              Call Modulo.ExecSQL(sSQL)
            End If
        End If
        
      Next Filas
      Unload fMensaje
      MsgBox "Operación Efectuada Correctamente.", vbInformation, "Información"
      
      If iCedExistencia > 0 Then
        MsgBox "Hubieron [" & CStr(iCedExistencia) & "] Cédulas que estaban registradas en la Base de Datos." & vbCrLf & _
               "Fueron actualizados todos sus Datos...", vbCritical, "Atención"
      End If
      Unload fImportar
      fPersonasAct.RefrescarPersonas
      fPersonasAct.bAuditarFotos_Click
      'fpersonas.
    End If
  End If
        
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'        If MsgBox("Se ha detectado que existen [" & cr & "] en la Base de datos del Cliente..." & vbCrLf & _
'                  "¿Desea BORRAR todos los registros existentes antes de iniciar la Agregación?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
'          s = "delete from " & st
'          Call Modulo.ExecSQL(s)
'        End If
'
'        agregando = True
'        Load fMensaje
'        fMensaje.Caption = "Agregando, Espere..."
'        fMensaje.Show
'        DoEvents
'
'        bHayUpdate = False
'
'        cr = Modulo.Total_Registros(st)
'        If cr = 0 Then 'Agregación en masa!
'          s = "select * from " & st
'          r.Open s, Modulo.DBConexionSQL, adOpenDynamic, adLockOptimistic
'          For i = 1 To MSFlexGrid1.Rows - 1
'            fMensaje.Caption = "Agregando Fila [" & CStr(i) & "], Espere..."
'            fMensaje.Show
'            DoEvents
'            For j = 1 To MSFlexGrid1.Cols - 1
'              sCampo = MSFlexGrid1.TextMatrix(0, j)
'              If sCampo <> "ERRORES" Then
'                If Modulo.EXISTE_CAMPO(r, s) = True Then
'                  If agregando Then
'                    r.AddNew
'                    agregando = False
'                    bHayUpdate = True
'                  End If
'                  s = Trim(MSFlexGrid1.TextMatrix(i, j))
'                  r.Fields(sCampo).Value = s
'                End If
'              End If
'            Next j
'          Next i
'          Unload fMensaje
'          r.Update
'          r.Close
'        Else
'          bHayUpdate = False
'          s = "select * from " & st 'Buscar primero, agregacion con registros
'          r.Open s, Modulo.DBConexionSQL, adOpenDynamic, adLockOptimistic
'          For i = 1 To MSFlexGrid1.Rows - 1
'            fMensaje.Caption = "Agregando Fila [" & CStr(i) & "], Espere..."
'            fMensaje.Show
'            DoEvents
'
'            If Modulo.EXISTE_CAMPO(r, "CEDULA") Then
'              s = MSFlexGrid1.TextMatrix(i, Modulo.Nro_Columna_FlexGrid(MSFlexGrid1, "CEDULA"))
'              If Modulo.Total_Registros(st) > 0 Then
'                r.MoveFirst
'                s = "CEDULA = '" & MSFlexGrid1.TextMatrix(i, Modulo.Nro_Columna_FlexGrid(MSFlexGrid1, "CEDULA")) & "'"
'                r.Find s
'                If Not r.EOF Then
'                  If MsgBox("La Cédula [" & s & "] existe en la base de datos," & vbCrLf & _
'                             "¿Desea solo actualizar los datos del registro?", vbCritical + vbYesNo, "Confirme") = vbYes Then
'
'                    s = MSFlexGrid1.TextMatrix(i, Modulo.Nro_Columna_FlexGrid(MSFlexGrid1, "CEDULA"))
'                    For j = 1 To MSFlexGrid1.Cols - 1
'                      If s <> "ERRORES" Then
'                        If Modulo.EXISTE_CAMPO(r, s) = True Then
'                          s = Trim(MSFlexGrid1.TextMatrix(i, j))
'                          r.Fields(s).Value = s
'                          bHayUpdate = True
'                        End If
'                      End If
'                    Next j
'                  End If
'                End If
'              End If
'            End If
'          Next i
'          Unload fMensaje
'        End If
'
'        Unload fMensaje
'
'      Else
'        r.Close
'        Unload fMensaje
'      End If
'    End If
'  End If
'  Set r = Nothing
End Sub

Private Sub bCancelar_Click()
  Unload Me
End Sub



Private Sub bCopiar_Click()
  Dim s As String
  Dim i As Integer
  s = ""
  Clipboard.Clear
  For i = 0 To fPersonasAct.Adodc1.Recordset.Fields.Count - 1
    If fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "ID" And _
       fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "FOTO" And _
       fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "FECHA" And _
       fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "CONTADOR" And _
       fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "TIENE_FOTO" And _
       fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "CREACION" And _
       fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "MARCA" Then
      s = s & fPersonasAct.Adodc1.Recordset.Fields(i).Name & vbTab
    End If
  Next i
  
  Clipboard.SetText s
  
  MsgBox "Estructura:" & vbCrLf & "[" & s & "]" & vbCrLf & "copiada al portapapeles, recuerde que puede «PEGAR» en Excel...", vbInformation, "Información"
End Sub

Private Sub bExaminar_Click()
  Dim p As Integer, k As Integer, i As Integer
  Dim s As String
  
  On Error Resume Next
  CommonDialog1.CancelError = True
  CommonDialog1.InitDir = Modulo.vRUTAINICIAL
  CommonDialog1.FileName = ""
  CommonDialog1.ShowOpen
  If Err.Number = 0 Then
    If CommonDialog1.FileName <> "" Then
      Text1.Text = CommonDialog1.FileName
    
      'Introducir los campos en un arreglo:
      k = fPersonasAct.Adodc1.Recordset.Fields.Count
      If k > 0 Then
        ReDim aCampos(k)
        For i = 0 To k - 1
          aCampos(i) = ""
          If fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "ID" And _
            fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "FOTO" And _
            fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "FECHA" And _
            fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "CONTADOR" And _
            fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "TIENE_FOTO" And _
            fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "CREACION" And _
            fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "MARCA" Then
            aCampos(i) = fPersonasAct.Adodc1.Recordset.Fields(i).Name
          End If
        Next i
      End If
    End If
  End If
End Sub

Private Sub CargarCOLUMNAS()
  Dim i As Integer
  'lCampos.Clear
  'lColumnas.Clear
  'lVinculos.Clear
  For i = 0 To fPersonasAct.lCampos.ListCount - 1
    If fPersonasAct.lCampos.List(i) <> "ID" Then
      'lCampos.AddItem fPersonasAct.lCampos.List(i)
      'lVinculos.AddItem "NO ASIGNADO"
    End If
  Next i
End Sub

Private Sub Quitar_Puntos_Comas_Columna(iColumna As Integer, iColumnaFoto As Integer)
  Dim i As Integer, j As Integer, k As Integer
  Dim c As String
  Dim s As String
  i = iColumna
  If i >= 0 And i < MSFlexGrid1.Cols Then
    'If MsgBox("Columna " & MSFlexGrid1.TextMatrix(0, i) & vbCrLf & _
    '          "¿Está Seguro de quitar PUNTOS(.) y COMAS(,) ?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
      For j = 1 To MSFlexGrid1.Rows - 1
        s = MSFlexGrid1.TextMatrix(j, i)
        c = ""
        For k = 1 To Len(s)
          If Mid(s, k, 1) <> "." And Mid(s, k, 1) <> "," Then
            c = c & Mid(s, k, 1)
          End If
        Next k
        If c = "" Then c = s
        MSFlexGrid1.TextMatrix(j, i) = c
        If iColumnaFoto >= 0 Then MSFlexGrid1.TextMatrix(j, iColumnaFoto) = c
      Next j
    'End If
  End If
End Sub

Private Function Es_Fila_Blanco(iFila As Integer) As Boolean
  Dim i As Integer, bEnBlanco As Boolean
  bEnBlanco = False
  i = 0
  Do While i < MSFlexGrid1.Cols
    If Trim(MSFlexGrid1.TextMatrix(iFila, i)) = "" Then bEnBlanco = True Else bEnBlanco = False
    i = i + 1
  Loop
  Es_Fila_Blanco = bEnBlanco
End Function

Private Sub bGenerarArchivoErrores_Click()
  Dim a As String
  Dim l As String
  
  Dim i As Integer, k As Integer, j As Integer
  Dim sNom As String, sOrigen As String
  Dim Fila As Integer
  Dim bHayCedula As Boolean
  Dim s1 As String, s2 As String, s3 As String
  Dim sCadenaEXE As String, sRutaExe As String
  Dim sReemplazo As Boolean
  
  sReemplazo = False
  
  If Dir(App.Path & "\" & "vacio.xls") = "" Then
    MsgBox "Falta el archivo [vacio.xls] en la carpeta del programa para hacer la copia. Revise...", vbCritical, "Información"
    Exit Sub
  End If
 
  s1 = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
  s2 = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
    
  sOrigen = App.Path & "\" & "vacio.xls"
    
  a = "LISTADO " & Trim(Mid(fPersonasAct.cCP.Text, 9))
  
  If Trim(fPersonasAct.cSC.Text) <> "" Then
    a = a & "_" & Trim(Mid(fPersonasAct.cSC.Text, 9))
  End If
  
  a = a & "_" & Format(Date, "dd") & "_" & _
                Format(Date, "mm") & "_" & _
                Format(Date, "yy") & "_OFICINA_ERRORES.XLS"
  
  s3 = s2 & "\" & Trim(Mid(fPersonasAct.cCP.Text, 9)) & "\" & _
                  IIf(Trim(fPersonasAct.cSC.Text) <> "", Trim(Mid(fPersonasAct.cSC.Text, 9)) & "\", "") & _
                  a
                  
  If Dir(s3) <> "" Then
    If MsgBox("Se ha Detectado que ya existe un archivo de ERRORES EXCEL de fecha " & Format(Date, "dd/mm/yy") & " Existe." & vbCrLf & _
               "¿Desea Reemplazarlo en este momento ?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
      sReemplazo = True
      Kill s3
      'sRutaExe = "C:\Archivos de programa\Microsoft Office\Office12\EXCEL.EXE"
        
      'If Dir(sRutaExe) <> "" Then
      '  sCadenaEXE = " " & Chr(34) & s3 & Chr(34) & " "
      '  Shell sRutaExe & sCadenaEXE, vbMaximizedFocus
      'End If
    End If
  Else
    sReemplazo = True
  End If
  
  If sReemplazo Then
    FileCopy sOrigen, s3
  End If
  
  MsgBox "Archivo de ERRORES XLS [" & a & "] ha sido Creado...", vbInformation, "Información"
         
  '-- Introducir TODA la informacion del FLEXGRID:
    
  Load fMensaje
  fMensaje.Label1.Caption = "Guardando Información en EXCEL, Espere..."
  fMensaje.Show
  
  If Guardar_ERRORES_FlexGrid_Excel(s3, MSFlexGrid1) = True Then
    MsgBox "Archivo de Errores Generado Correctamente.", vbInformation, "Información"
    
    Load fPrepararEmail
    fPrepararEmail.eAdjunto.Text = s3
    fPrepararEmail.ListAdjunto.AddItem s3
    If s3 <> "" Then fPrepararEmail.bVerAdjunto.Enabled = True
    fPrepararEmail.Show
    
  End If
    
  Unload fMensaje

End Sub

Private Sub Preparar_Email_Listado_ERRORES()
  Dim sContenido As String
  Dim sCorreo As String
  Dim s1 As String, s2 As String
  sContenido = "Estimados Señores:" & vbCrLf & _
  "Les enviamos el correo con el archivo del listado con errores...."
  
  s1 = Trim(fPersonasAct.cCP.Text)
  s2 = Trim(fPersonasAct.cSC.Text)
  If s1 <> "" Then s1 = Mid(s1, 1, 6)
  If s2 <> "" Then s2 = Mid(s2, 1, 6)
  Dim lists1 As ListBox
  lists1.AddItem s1
  sCorreo = Correo_E(s1, s2)
  sCorreo = LCase(sCorreo)
  'sCorreo = "rangelobb@gmail.com"
  
  s1 = "C:\Clientes\TAXI RAPIDO Y FURIOSO\LISTADO TAXI RAPIDO Y FURIOSO_26_06_09_OFICINA_ERRORES.XLS"
      
  Call MandaMail(sCorreo, "", "", "LISTADO CON ERRORES", sContenido, lists1)



End Sub


Private Sub bLeerXLS_Click()
  Dim a As String, k As Integer, i As Integer, Fila As Integer, j As Integer
  Dim bHayCedula As Boolean
  On Error GoTo falla
  If Trim(Text1.Text) = "" Then
    MsgBox "Debe Indicar archivo en formato EXCEL a procesar...", vbCritical, "Información"
  Else
   lErrores = False
    'Introducir los campos en un arreglo:
    k = fPersonasAct.Adodc1.Recordset.Fields.Count
    If k > 0 Then
      ReDim aCampos(k)
      For i = 0 To k - 1
        aCampos(i) = ""
        If fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "ID" And _
          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "FOTO" And _
          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "FECHA" And _
          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "CONTADOR" And _
          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "TIENE_FOTO" And _
          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "CREACION" And _
          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "MARCA" Then
          
          aCampos(i) = fPersonasAct.Adodc1.Recordset.Fields(i).Name
        End If
      Next i
    End If
 
    a = Trim(Text1.Text)
    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.FixedRows = 1
    MSFlexGrid1.Cols = 1
    Fila = 1
    j = 0
    For i = 0 To fPersonasAct.Adodc1.Recordset.Fields.Count - 1
      If Trim(aCampos(i)) <> "" Then
        MSFlexGrid1.TextMatrix(0, j) = aCampos(i)
        If UCase(aCampos(i)) = "CEDULA" Then bHayCedula = True
        MSFlexGrid1.Cols = MSFlexGrid1.Cols + 1
        j = j + 1
      End If
    Next i

    If bHayCedula Then '--Lleva FOTO!
      'MSFlexGrid1.Cols = MSFlexGrid1.Cols + 1
      MSFlexGrid1.TextMatrix(0, MSFlexGrid1.Cols - 1) = "FOTO"
    Else
      MSFlexGrid1.Cols = MSFlexGrid1.Cols - 1
    End If
    
       
    
    
    If MsgBox("¿Está Seguro de Leer datos desde" & vbCrLf & _
               a, vbQuestion + vbYesNo, "Confirme") = vbYes Then
      
      If Leer_Excel_FlexGrid(a, MSFlexGrid1) = True Then
      
        If MSFlexGrid1.Rows > 1 Then
          If Es_Fila_Blanco(MSFlexGrid1.Rows - 1) Then
            MSFlexGrid1.RemoveItem (MSFlexGrid1.Rows - 1)
          End If
        End If
        
        'MsgBox "Datos Leidos Correctamente, Por favor verifique los mismos antes de Grabar toda la Información a la Base de Datos.", vbInformation, "Información"
        If MSFlexGrid1.Rows >= 1 Then Label3.Caption = CStr(MSFlexGrid1.Rows - 1)
        
        'Agregar la columna de ERRORES:
        MSFlexGrid1.Cols = MSFlexGrid1.Cols + 1
        MSFlexGrid1.TextMatrix(0, MSFlexGrid1.Cols - 1) = "ERRORES"
        MSFlexGrid1.ColWidth(MSFlexGrid1.Cols - 1) = 4600
        
        j = Nro_Columna_FlexGrid(MSFlexGrid1, "NOMBRE")
        If j >= 0 Then MSFlexGrid1.ColWidth(j) = 3000
          
                 
       
        If bHayCedula Then
          Load fMensaje
          fMensaje.Label1.Caption = "Formateando la columna de Cedula, Espere..."
          fMensaje.Show
          DoEvents
          
          
          
          'lLog.Clear
          
          Call QuitarPuntos_CEDULA
          Call ColocarPuntos_CEDULA
          Load fMensaje
          fMensaje.Label1.Caption = "Colocando la fecha de vencimiento segun el cliente, Espere..."
          fMensaje.Show
          DoEvents
          sColocarVencimiento
          
          
          Unload fMensaje
        End If
        
            sChequearCampos
      End If
    End If
  End If
falla:
  If Err.Number <> 0 Then
     MsgBox "Ocurrió un error al intentar leer los datos desde el archivo seleccionado.", vbCritical
     Unload fMensaje
  End If
End Sub

Private Sub sColocarVencimiento()
 Dim lReg As New ADODB.Recordset
 Dim lReg1 As New ADODB.Recordset
 Dim i As Integer
 Dim lColumna As Integer
 Dim lCliente As String
 lCliente = Mid(fPersonasAct.cCP.Text, 1, 6)
 For i = 0 To MSFlexGrid1.Rows
    If MSFlexGrid1.TextMatrix(0, i) = "VENCE" Then
           lColumna = i
           Exit For
    End If
 Next i
  'lReg.Close
  'BUSCAR EL ID DEL DISEÑO ACTIVO EN EL CLIENTE
 lReg1.Open "Select IDDIsenoActivo from clientes where codigo=" & lCliente, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
 If lReg1.EOF = False Then
    lReg.Open "Select Vencimiento from formatodisenoDetalle where cliente=" & lCliente & " And ID=" & lReg1!IdDisenoActivo, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    If lReg.EOF = False Then
       If (IsNull(lReg!vencimiento) = True) Then
         Exit Sub
       End If
       
       
       If ((lReg!vencimiento) <= CDate(Now)) Then
          MsgBox "La Fecha de Vencimiento en el formato de diseño es menor o igual a la fecha actual (" & lReg!vencimiento & ")", vbCritical
          Exit Sub
       End If
       For i = 1 To MSFlexGrid1.Rows - 1
          If MSFlexGrid1.TextMatrix(i, lColumna) = "" Then
             MSFlexGrid1.TextMatrix(i, lColumna) = IIf(IsDate(lReg!vencimiento), MonthName(Month(lReg!vencimiento)) & " " & Year(lReg!vencimiento), lReg!vencimiento)
          End If
        Next i
    End If
 End If
End Sub

Private Sub Revisar_Valores_Blanco()
  Dim i As Integer, j As Integer
  Dim CE As Integer
  
  CE = Nro_Columna_FlexGrid(MSFlexGrid1, "ERRORES")
  lErrores = False
  If CE > 0 Then
     For i = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.TextMatrix(i, CE) = ""
        Call Pintar_Fila(i, vbWhite, &H8000&)
     Next i
     
  End If
  i = 0
  Do While (i < MSFlexGrid1.Rows)
    fMensaje.Label1.Caption = "Revisando Valores en Blanco, Espere... Fila Nº " & CStr(i) & " "
    fMensaje.Show
    DoEvents

    j = 0
    Do While (j < MSFlexGrid1.Cols)
    
      If Trim(MSFlexGrid1.TextMatrix(0, j)) <> "ERRORES" And Trim(MSFlexGrid1.TextMatrix(0, j)) <> "FOTO" Then
    
        If Trim(MSFlexGrid1.TextMatrix(i, j)) = "" Then
          'lLog.AddItem "[" & MSFlexGrid1.TextMatrix(0, j) & "] en blanco. Fila Nº " & CStr(i)
        
          Call Pintar_Fila(i, vbWhite, vbRed)
        
          MSFlexGrid1.TextMatrix(i, CE) = "[" & MSFlexGrid1.TextMatrix(0, j) & "] en blanco. " & MSFlexGrid1.TextMatrix(i, CE)
          lErrores = True
        End If
      End If
      j = j + 1
    Loop
    i = i + 1
  Loop
End Sub

Function Nro_Columna_Cedula() As Integer
  Dim cc As Integer
  Dim i As Integer
  i = 0
  cc = -1
  Do While i < MSFlexGrid1.Cols And cc = -1
    If UCase(MSFlexGrid1.TextMatrix(0, i)) = "CEDULA" Then cc = i
    i = i + 1
  Loop
  Nro_Columna_Cedula = cc
End Function

Function Nro_Columna_Nombre() As Integer
  Dim cc As Integer
  Dim i As Integer
  i = 0
  cc = -1
  Do While i < MSFlexGrid1.Cols And cc = -1
    If UCase(MSFlexGrid1.TextMatrix(0, i)) = "NOMBRE" Then cc = i
    i = i + 1
  Loop
  Nro_Columna_Nombre = cc
End Function


Private Function Existe_Repetidas(sCedula As String) As Boolean
  Dim i As Integer
  Dim ER As Boolean
  i = 0
  ER = False
  Do While i < lRepetidas.ListCount And Not ER
    If sCedula = lRepetidas.List(i) Then ER = True Else i = i + 1
  Loop
  Existe_Repetidas = ER
End Function


Private Sub Revisar_Cedula_Repetidas()
  Dim aRepetidas() As String        'Arreglo de Cedulas no definido por defecto.
  Dim i As Integer, j As Integer, k As Integer
  Dim iColCed As Integer
  Dim s As String
  Dim CE As Integer
  
  CE = Nro_Columna_FlexGrid(MSFlexGrid1, "ERRORES")
   
  iColCed = Nro_Columna_Cedula()
  
  If iColCed < 0 Then
    MsgBox "No Existe el Campo CEDULA, Revise...", vbCritical, "Información"
    Exit Sub
  End If
  
  
  
  lRepetidas.Clear
  lFilas.Clear
     
  'Recorrido Primario: Las Cedulas en el FLEXGRID
  For i = 1 To MSFlexGrid1.Rows - 1
    
    fMensaje.Label1.Caption = "Verificando CEDULAS, Espere... [" & CStr(i) & " / " & CStr(MSFlexGrid1.Rows - 1) & "]"
    fMensaje.Show
    DoEvents
    
    s = Trim(MSFlexGrid1.TextMatrix(i, iColCed))
    If Len(s) > 10 Then
       MSFlexGrid1.TextMatrix(i, CE) = MSFlexGrid1.TextMatrix(i, CE) & "CEDULA MAYOR A 8 DIGITOS. "
       Call Pintar_Fila(i, vbWhite, vbRed)
    End If
    
    If Len(s) < 7 Then
       MSFlexGrid1.TextMatrix(i, CE) = MSFlexGrid1.TextMatrix(i, CE) & "CEDULA MENOR A 6 DIGITOS. "
       Call Pintar_Fila(i, vbWhite, vbRed)
    End If


    lErrores = False

    If s <> "" Then
      If Not Existe_Repetidas(s) Then
        'Recorrido por todos los datos de la columnas:
        For j = 1 To MSFlexGrid1.Rows - 1
          If s = Trim(MSFlexGrid1.TextMatrix(j, iColCed)) And i <> j Then
            lRepetidas.AddItem s
            lFilas.AddItem CStr(j)
            
            Call Pintar_Fila(j, vbWhite, vbRed)
            If s <> "" Then
              If CE >= 0 Then
                MSFlexGrid1.TextMatrix(j, CE) = MSFlexGrid1.TextMatrix(j, CE) & "CEDULA " & s & " REPETIDA. "
                lErrores = True
              End If
            End If
          End If
        Next j
      End If
    End If
  Next i
  
  fMensaje.Label1.Caption = "Escribiendo LOG, Espere..."
  fMensaje.Show
  DoEvents
  
  'Vaciar el arreglo para formar el LOG:
  If lRepetidas.ListCount > 0 Then
    For i = 0 To lRepetidas.ListCount - 1
      'lLog.AddItem "Cedula [" & lRepetidas.List(i) & "] aparece en la Fila Nº " & Str(lFilas.List(i))
    Next i
  End If
  'Dim j As Integer
  For i = 1 To MSFlexGrid1.Rows - 1
     For j = 2 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(i, iColCed) = MSFlexGrid1.TextMatrix(j, iColCed) And i <> j Then
           Call Pintar_Fila(j, vbWhite, vbRed)
           Call Pintar_Fila(i, vbWhite, vbRed)
        End If
     Next j
  Next i
  
  
  
End Sub

Private Sub Revisar_Nombres_Repetidos()
  Dim aRepetidas() As String        'Arreglo de Cedulas no definido por defecto.
  Dim i As Integer, j As Integer, k As Integer
  Dim iColNomb As Integer
  Dim s As String
  Dim CE As Integer
  
  CE = Nro_Columna_FlexGrid(MSFlexGrid1, "ERRORES")
   
  iColNomb = Nro_Columna_Nombre()
  
  If iColNomb < 0 Then
    MsgBox "No Existe el Campo CEDULA, Revise...", vbCritical, "Información"
    Exit Sub
  End If
  
  
  
  lRepetidas.Clear
  lFilas.Clear
     
  'Recorrido Primario: Las Cedulas en el FLEXGRID
  For i = 1 To MSFlexGrid1.Rows - 1
    
    fMensaje.Label1.Caption = "Verificando NOMBRES, Espere... [" & CStr(i) & " / " & CStr(MSFlexGrid1.Rows - 1) & "]"
    fMensaje.Show
    DoEvents
    
    s = Trim(MSFlexGrid1.TextMatrix(i, iColNomb))
    
    If s <> "" Then
      If Not Existe_Repetidas(s) Then
        'Recorrido por todos los datos de la columnas:
        For j = 1 To MSFlexGrid1.Rows - 1
          If s = Trim(MSFlexGrid1.TextMatrix(j, iColNomb)) And i <> j Then
            lRepetidas.AddItem s
            lFilas.AddItem CStr(j)
            
            Call Pintar_Fila(j, vbBlack, vbYellow)
            If s <> "" Then
              If CE >= 0 Then
                MSFlexGrid1.TextMatrix(j, CE) = MSFlexGrid1.TextMatrix(j, CE) & "NOMBRE " & s & " REPETIDO. "
                lErrores = True
              End If
            End If
          End If
        Next j
      End If
    End If
  Next i
  
  fMensaje.Label1.Caption = "Escribiendo LOG, Espere..."
  fMensaje.Show
  DoEvents
  
  'Vaciar el arreglo para formar el LOG:
  If lRepetidas.ListCount > 0 Then
    For i = 0 To lRepetidas.ListCount - 1
      'lLog.AddItem "Cedula [" & lRepetidas.List(i) & "] aparece en la Fila Nº " & Str(lFilas.List(i))
    Next i
  End If
  'Dim j As Integer
  For i = 1 To MSFlexGrid1.Rows - 1
     For j = 2 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(i, iColNomb) = MSFlexGrid1.TextMatrix(j, iColNomb) And i <> j Then
           Call Pintar_Fila(j, vbBlack, vbYellow)
           Call Pintar_Fila(i, vbBlack, vbYellow)
        End If
     Next j
  Next i
  
  
  
End Sub


Private Sub Pintar_Fila(iFila As Integer, lColorTxt As Long, lColorFnd As Long)
  Dim c As Integer
  If iFila > 0 Then  ' And iFila < MSFlexGrid1.Rows
    MSFlexGrid1.Row = iFila
    For c = 0 To MSFlexGrid1.Cols - 1
      MSFlexGrid1.Col = c
      MSFlexGrid1.CellBackColor = lColorFnd
      MSFlexGrid1.CellForeColor = lColorTxt
      MSFlexGrid1.Refresh
    Next c
  End If
End Sub

Private Sub bPreparar_Click()
  Dim a As String
  Dim l As String
  
  Dim i As Integer, k As Integer, j As Integer
  Dim sNom As String, sOrigen As String
  Dim Fila As Integer
  Dim bHayCedula As Boolean
  Dim s1 As String, s2 As String, s3 As String
  Dim sCadenaEXE As String, sRutaExe As String
  
  If Dir(App.Path & "\" & "vacio.xls") = "" Then
    MsgBox "Falta el archivo [vacio.xls] en la carpeta del programa para hacer la copia. Revise...", vbCritical, "Información"
    Exit Sub
  End If
 
  s1 = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
  s2 = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
    
  sOrigen = App.Path & "\" & "vacio.xls"
    
  a = "LISTADO " & Trim(Mid(fPersonasAct.cCP.Text, 9))
  
  If Trim(fPersonasAct.cSC.Text) <> "" Then
    a = a & "_" & Trim(Mid(fPersonasAct.cSC.Text, 9))
  End If
  
  a = a & "_" & Format(Date, "dd") & "_" & _
                Format(Date, "mm") & "_" & _
                Format(Date, "yy") & "_OFICINA.XLS"
  
  s3 = s2 & "\" & Trim(Mid(fPersonasAct.cCP.Text, 9)) & "\" & _
                  IIf(Trim(fPersonasAct.cSC.Text) <> "", Trim(Mid(fPersonasAct.cSC.Text, 9)) & "\", "") & _
                  a
                  
  If Dir(s3) <> "" Then
    If MsgBox("Se ha Detectado que ya existe un archivo de Listado EXCEL de fecha " & Format(Date, "dd/mm/yy") & " Existe." & vbCrLf & _
               "¿Desea Visualizarlo en este momento ?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
               
      sRutaExe = "C:\Archivos de programa\Microsoft Office\Office12\EXCEL.EXE"
        
      If Dir(sRutaExe) <> "" Then
        sCadenaEXE = " " & Chr(34) & s3 & Chr(34) & " "
        Shell sRutaExe & sCadenaEXE, vbMaximizedFocus
      End If
    End If
    Exit Sub
  End If
    
  FileCopy sOrigen, s3
  
  MsgBox "Archivo de Excel [" & a & "] ha sido Creado...", vbInformation, "Información"
         
  '-- Estampar la estructura de datos:
  bHayCedula = False
  If Trim(a) <> "" Then
    Load fMensaje
    fMensaje.Label1.Caption = "Preparando Documento EXCEL, Espere..."
    fMensaje.Show
      
    'Introducir los campos en un arreglo:
    k = fPersonasAct.Adodc1.Recordset.Fields.Count
    If k > 0 Then
      ReDim aCampos(k)
      For i = 0 To k - 1
        aCampos(i) = ""
        If fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "ID" And _
          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "FOTO" And _
          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "FECHA" And _
          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "CONTADOR" And _
          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "TIENE_FOTO" And _
          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "CREACION" And _
          fPersonasAct.Adodc1.Recordset.Fields(i).Name <> "MARCA" Then
          aCampos(i) = fPersonasAct.Adodc1.Recordset.Fields(i).Name
        End If
      Next i
        
      If Estampar_Cabecera_Excel_FlexGrid(aCampos(), k, s3, MSFlexGrid1) = True Then
        '--Llamar a EXCEL con el archivo abierto para tratamiento manual de la data
        
        '--Determinar segun Registro de Windows donde está instalado
        '--Excel:
        
        'Para Office 2007:
        sRutaExe = "C:\Archivos de programa\Microsoft Office\Office12\EXCEL.EXE"
                
        'Para Office 2003:
        'sRutaExe = "C:\Archivos de programa\Microsoft Office\Office9\EXCEL.EXE"
        
        If Dir(sRutaExe) <> "" Then
          sCadenaEXE = " " & Chr(34) & s3 & Chr(34) & " "
          Shell sRutaExe & sCadenaEXE, vbMaximizedFocus
          Text1.Text = s3
        End If
                
      End If
      
    Else
      MsgBox "No hay Campos de la Base de Datos a Preparar en EXCEL.", vbCritical, "Información"
    End If
    
    Unload fMensaje
    
  End If
  
End Sub

Private Sub QuitarPuntos_CEDULA()
  Dim i As Integer, j As Integer, k As Integer
  Dim c As String
  Dim s As String
  Dim e As Boolean, e1 As Boolean
  Dim fotocol As Integer
  i = 0
  e = False
  Do While i < MSFlexGrid1.Cols And Not e
    If MSFlexGrid1.TextMatrix(0, i) = "CEDULA" Then e = True Else i = i + 1
  Loop
  
  j = 0
  e1 = False
  Do While i < MSFlexGrid1.Cols And Not e1
    If MSFlexGrid1.TextMatrix(0, j) = "FOTO" Then e1 = True Else j = j + 1
  Loop
  
  If e Then Quitar_Puntos_Comas_Columna i, j
    
End Sub

Private Sub ColocarPuntos_CEDULA()
  Dim i As Integer, j As Integer, k As Integer
  Dim c As String
  Dim s As String
  Dim d As Double
    
  Dim e As Boolean
  i = 0
  e = False
  Do While i < MSFlexGrid1.Cols And Not e
    If MSFlexGrid1.TextMatrix(0, i) = "CEDULA" Then e = True Else i = i + 1
  Loop
  
  If e Then
    'If MsgBox("Columna " & MSFlexGrid1.TextMatrix(0, i) & vbCrLf & _
              "¿Está Seguro de colocar PUNTOS(.) en Cantidades?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
      For j = 1 To MSFlexGrid1.Rows - 1
        s = MSFlexGrid1.TextMatrix(j, i)
        If IsNumeric(s) Then
          d = CDbl(s)
          s = Format(d, "#,0")
          MSFlexGrid1.TextMatrix(j, i) = s
        End If
      Next j
    'End If
  End If

End Sub



Private Sub Command1_Click()
  Dim sContenido As String
  Dim sCorreo As String
  Dim s1 As String, s2 As String
  sContenido = "Estimados Señores:" & vbCrLf & _
  "Les enviamos el correo con el archivo del listado con errores...."
  
  s1 = Trim(fPersonasAct.cCP.Text)
  s2 = Trim(fPersonasAct.cSC.Text)
  If s1 <> "" Then s1 = Mid(s1, 1, 6)
  If s2 <> "" Then s2 = Mid(s2, 1, 6)
  Dim lists1 As ListBox
  lists1.AddItem s1
  sCorreo = Correo_E(s1, s2)
  sCorreo = LCase(sCorreo)
  sCorreo = "rangelobb@gmail.com"
  
  s1 = "C:\Clientes\TAXI RAPIDO Y FURIOSO\LISTADO TAXI RAPIDO Y FURIOSO_26_06_09_OFICINA_ERRORES.XLS"
                                          
      
  Call MandaMail(sCorreo, "", "", "LISTADO CON ERRORES", sContenido, lists1)
    
End Sub

Private Sub Form_Load()
  
  Text1.Text = ""
  Label3.Caption = "-"
    
  CargarCOLUMNAS
  lErrores = False
End Sub

Private Sub sChequearCampos()
  Load fMensaje
  fMensaje.Label1.Caption = "Revisando Valores en Blanco, Espere..."
  fMensaje.Show
  DoEvents
     Call Revisar_Valores_Blanco
        
  fMensaje.Label1.Caption = "Revisando Cédulas Repetidas, Espere..."
  fMensaje.Show
  DoEvents
     Call Revisar_Cedula_Repetidas
     
  fMensaje.Label1.Caption = "Revisando Nombres Repetidos, Espere..."
  fMensaje.Show
  DoEvents
     Call Revisar_Nombres_Repetidos
     
        
  Unload fMensaje

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Descargar
End Sub

Private Sub MSFlexGrid1_Click()
MSFlexGrid1.ToolTipText = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Cols - 1)

End Sub

Private Sub MSFlexGrid1_DblClick()
  Dim s As String
  
  If MSFlexGrid1.Col < 0 Then Exit Sub
  
  s = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
  s = InputBox("Editar Valor:", "Edición", s)
  If s <> "" Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col) = s
    If UCase(MSFlexGrid1.TextMatrix(0, MSFlexGrid1.Col)) = "CEDULA" Then
      Call QuitarPuntos_CEDULA
      Call ColocarPuntos_CEDULA
    End If
  End If
  sChequearCampos
End Sub


'Public Function MandaMail(Para As String, _
'                          CC As String, _
'                          BCC As String, _
'                          Titulo As String, _
'                          Cuerpo As String, _
'                          ArchivoAdj As String) As Boolean
'
'  Dim correo As Outlook.Application
'  Dim item As Outlook.MailItem
'  Dim sClase As String
'
'  Load fMensaje
'  fMensaje.Label1.Caption = "Generando Correo MS-OUTLOOK, Espere..."
'  fMensaje.Show
'  DoEvents
'
'  sClase = "Outlook.Application" '"Microsoft.Office.Interop.Outlook.Application"  '"Outlook.Application"
'
'  Set correo = New Outlook.Application  'GetObject("", sClase)
'
'  'If correo Is Nothing Then
'  '  Set correo = CreateObject("", sClase)
'  'End If
'
'  Set item = correo.CreateItem(olMailItem)
'
'  On Error GoTo ErrorMAPI
'
'  DoEvents
'
'  If Trim(Para) <> "" Then item.To = Para
'  If Trim(BCC) <> "" Then item.BCC = BCC
'  If Trim(CC) <> "" Then item.CC = CC
'
'  item.Subject = Titulo
'  item.Body = Cuerpo
'
'  item.Attachments.Add ArchivoAdj
'
'  Unload fMensaje
'  'item.Display True
'
'  'item.Send
'
'  correo.Session.SendAndReceive True
'
'
'  Set item = Nothing
'  Set correo = Nothing
'  MandaMail = True
'
'Salir:
'  Exit Function
'ErrorMAPI:
'  MandaMail = False
'  Screen.MousePointer = vbDefault
'  MsgBox "Hubo un error al enviar el correo electrónico, el error fue debido a: " + Err.Description
'  Resume Salir
'End Function



Private Sub MSFlexGrid1_GotFocus()
MSFlexGrid1.ToolTipText = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Cols - 1)
End Sub
