VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fHXLS 
   Caption         =   "Generar .XLS con Listado de Fotos"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8130
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Carpeta Origen"
      Height          =   9915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15195
      Begin VB.CheckBox Check2 
         Caption         =   "Copiar en XLS sólo el Nombre de Archivo sin la Extension"
         Height          =   435
         Left            =   12270
         TabIndex        =   16
         Top             =   2450
         Width           =   2715
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Copiar la Cabecera de Datos"
         Height          =   195
         Left            =   9570
         TabIndex        =   13
         Top             =   2550
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Caption         =   "Generar en XLS"
         Height          =   4545
         Left            =   9450
         TabIndex        =   7
         Top             =   2910
         Width           =   5650
         Begin VB.ListBox List1 
            Height          =   2400
            Left            =   1320
            TabIndex        =   15
            Top             =   1620
            Width           =   3165
         End
         Begin VB.CommandButton bBuscarCP 
            Height          =   345
            Left            =   5100
            Picture         =   "fHXLS.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1140
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.ComboBox cCP 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   300
            Width           =   4200
         End
         Begin VB.ComboBox cSC 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   720
            Width           =   4200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cabecera:"
            Height          =   195
            Left            =   2490
            TabIndex        =   14
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Cliente Principal:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1170
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Cliente:"
            Height          =   195
            Left            =   420
            TabIndex        =   11
            Top             =   780
            Width           =   855
         End
      End
      Begin VB.CommandButton bGenerarXLS 
         Caption         =   "Generar XLS con el Listado"
         Height          =   555
         Left            =   11430
         TabIndex        =   6
         Top             =   1500
         Width           =   1500
      End
      Begin VB.CommandButton bCancelar 
         Caption         =   "Cancelar"
         Height          =   500
         Left            =   11850
         Picture         =   "fHXLS.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7770
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6840
         Left            =   5460
         TabIndex        =   3
         Top             =   600
         Width           =   3855
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   90
         TabIndex        =   2
         Top             =   600
         Width           =   5250
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   1
         Top             =   210
         Width           =   5250
      End
      Begin VB.Label lRutaActual 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
         Height          =   315
         Left            =   90
         TabIndex        =   4
         Top             =   7650
         Width           =   9240
      End
   End
End
Attribute VB_Name = "fHXLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cFG = "Nº      | ARCHIVO                                           | CÉDULA                                                "
' -- Variables para acceder a la hoja excel
Private obj_Excel       As Object
Private obj_Workbook    As Object
Private obj_Worksheet   As Object


Private Sub bBuscarCP_Click()
  Dim f As Integer, c As Integer, i As Integer
  Dim e As Boolean
  
  Modulo.vTemporal1 = ""
  Load fBuscarSimple
  
  With fBuscarSimple
    .Adodc1.ConnectionString = Modulo.DBConexionSQL.ConnectionString
    .Adodc1.Caption = "CLIENTES"
    .Adodc1.RecordSource = "select codigo, nombre, direccion, telefonos from clientes order by codigo"
    .Adodc1.Refresh
    .DataGrid1.Refresh
    
    .Combo1.Clear
    .Combo1.AddItem "CODIGO"
    .Combo1.AddItem "NOMBRE"
    .Combo1.ListIndex = 1
        
    Fields_DataGrid_En_Mayusculas .DataGrid1
    
    .DataGrid1.Columns(0).width = 800 'codigo
    .DataGrid1.Columns(1).width = 4000 'nombre
    .DataGrid1.Columns(2).width = 3000 'direccion
    .DataGrid1.Columns(3).width = 2000 'telefonos
  End With
   
  'fBuscarSimple.BuscarSimple fClientes.FG
  Modulo.vTemporal1 = ""
  'fBuscarSimple.Option2.Value = True 'Buscar por nombre por defecto
  fBuscarSimple.Show vbModal
  
  If Modulo.vTemporal1 <> "" Then
    cCP.ListIndex = Modulo.Buscar_ComboLen(cCP, Mid(Modulo.vTemporal1, 1, 6), 6)
  End If

End Sub

Private Sub bCancelar_Click()
  Unload Me
End Sub

Private Sub bGenerarXLS_Click()
  Dim o As String
  Dim s As String
  s = "Listado_Fotos_" & Format(Date, "dd_mm_yyyy") & ".xls"
  CommonDialog1.FileName = s
  CommonDialog1.CancelError = True
  CommonDialog1.ShowSave
  s = CommonDialog1.FileName
  If Err.Number = 0 Then
    o = App.Path & "\VACIO.xls"
    FileCopy o, s
    
    Guardar_LISTADO_FOTOS_Excel s
  End If
      
End Sub


Private Function Guardar_LISTADO_FOTOS_Excel(sPath As String, Optional sSheetName As String = vbNullString) As Boolean
    Dim i As Long, j As Long
    Dim n As Long
    Dim HayDatos As Boolean
    Dim s1 As String, s As String
    
    On Error GoTo error_sub
    
    Guardar_LISTADO_FOTOS_Excel = False
    
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
    
    obj_Worksheet.Cells(1, 2) = "FOTO"
        
    For i = 0 To File1.ListCount - 1
      
      s = File1.List(i)
      If Check2.Value = vbChecked Then
        s = Modulo.QuitarExtension(s)
      End If
           
      j = 2 + i
      obj_Worksheet.Cells(j, 2) = s
        
    Next i
    
    If Check1.Value = vbChecked Then
    
      If List1.ListCount > 0 Then
      
        For i = 0 To List1.ListCount - 1
        
          obj_Worksheet.Cells(1, 3 + i) = List1.List(i)
          
        Next i
          
      End If
      
    End If
    
    
      
    '-- Guardar los cambios al libro (22/06/09)
    obj_Workbook.Save
    ' -- Cerrar libro
    obj_Workbook.Close
    ' -- Cerrar Excel
    obj_Excel.Quit
    ' -- Descargar objetos para liberar recursos
    Call Descargar
' -- Errores
    Guardar_LISTADO_FOTOS_Excel = True
    MsgBox "Archivo XLS generado correctamente...", vbInformation, "Información"
Exit Function
error_sub:
    MsgBox Err.Description
    Call Descargar
    Me.MousePointer = vbDefault
    Guardar_LISTADO_FOTOS_Excel = False
End Function

Private Sub Descargar()
    On Local Error Resume Next
    Set obj_Workbook = Nothing
    Set obj_Excel = Nothing
    Set obj_Worksheet = Nothing
    Me.MousePointer = vbDefault
End Sub



Private Sub cCP_Click()
  If Trim(cCP.Text) <> "" Then
    Cargar_SubClientes
    MostrarCabecera
  End If
End Sub

Private Sub Check1_Click()
  If Check1.Value = vbChecked Then
    Frame2.Visible = True
  Else
    Frame2.Visible = False
  End If
End Sub

Private Sub MostrarCabecera()
  Dim c As String, SC As String
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim i As Integer
  
  
  'aCamposOcul(0) = ""
  'aCamposOcul(1) = ""
  'aCamposOcul(2) = ""
  'aCamposOcul(3) = ""
  'aCamposOcul(4) = "FOTO"
  'aCamposOcul(5) = "TIENE_FOTO"
  'aCamposOcul(6) = "MARCA"
  'aCamposOcul(7) = "FECHA"
  'aCamposOcul(8) = "CONTADOR"
 
  c = Mid(cCP.Text, 1, 6)
  SC = Mid(cSC.Text, 1, 6)
  
  If Trim(c) = "" Then Exit Sub
  
  If Trim(SC) = "" Then SC = "-"
  
  s = Modulo.La_Tabla_Actual_Personas(c, SC)
  
  If s <> "" Then
    
    s = "select * from [" & s & "]"
    r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    
    List1.Clear
    
    For i = 0 To r.Fields.Count - 1
    
      s = Trim(UCase(r.Fields(i).Name))
    
      If s <> "ID" And s <> "FOTO" And s <> "TIENE_FOTO" And _
         s <> "MARCA" And s <> "FECHA" And s <> "CONTADOR" And _
         s <> "CREACION" Then
              
        List1.AddItem s
        
      End If
      
    Next i
    
    r.Close
    Set r = Nothing
    
  End If

End Sub

Private Sub cSC_Click()
  MostrarCabecera
End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
End Sub

Private Sub Dir1_Click()
  'lRutaActual.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
  On Error Resume Next
  Dir1.Path = Drive1.Drive
  If Err.Number <> 0 Then
    MsgBox "Unidad No Está Disponible...", vbCritical, "Información"
    Drive1.Drive = "C"
    'Dir1.Path = "C"
  End If
End Sub

Private Sub SeleccionarCampo(eTexto As TextBox)
  eTexto.SelStart = 0
  eTexto.SelLength = Len(eTexto.Text)
End Sub


Private Sub File1_PathChange()
  'CargarArchivosDeDisco
  lRutaActual.Caption = File1.Path
  
  'If Trim(FG.TextMatrix(1, 1)) <> "" Then
  '  Analizar_Archivo Trim(FG.TextMatrix(1, 1))
  'End If
  
End Sub

Private Sub Analizar_Archivo(sNombre As String)
  Dim i As Integer
  Dim s As String
  Dim Caracter As String
  Dim sLetras As String
  Dim sDigitos As String
  
  'Letras : A..65 / Z..90
  'Digitos: 0..48 / 9..57
  s = ""
  sLetras = ""
  sDigitos = ""
  i = 1
  Do While i <= Len(sNombre)
    Caracter = Mid(sNombre, i, 1)
    If Caracter = "." Then
      i = Len(sNombre) + 1
    Else
      Select Case Caracter
        Case "A", "B", "C", "D", _
             "E", "F", "G", "H", _
             "I", "J", "K", "L", _
             "M", "N", "O", "P", _
             "Q", "R", "S", "T", _
             "U", "V", "W", "X", _
             "Y", "Z", "-", "_":
          sLetras = sLetras & Caracter
          
        Case "0", "1", "2", "3", "4", _
             "5", "6", "7", "8", "9":
          sDigitos = sDigitos & Caracter
      End Select
      
      i = i + 1
    End If
      
  Loop
  
  'eTexto.Text = sLetras
  'eDesde.Text = ""
  'eHasta.Text = ""
  
  'If Len(sDigitos) > 0 Then
  '  eDesde.Text = sDigitos
  'End If
  
  
  
   
End Sub


Private Sub Cargar_Clientes()
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim l As Integer
  Dim sCod As String
  
  s = "SELECT * FROM Clientes ORDER BY Codigo"
  
  r.Open s, DBConexionSQL, adOpenKeyset, adLockReadOnly
  
  cCP.Clear
  
  Do While Not r.EOF
    s = Zeros(r.Fields("codigo").Value, 6) & " : " & Trim(r.Fields("nombre").Value)
    cCP.AddItem s
    r.MoveNext
  Loop
  
  r.Close
  Set r = Nothing
End Sub


Private Sub Cargar_SubClientes()
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim l As Integer
  Dim sCod As String
  
  cSC.Clear
  
  'Viendo = False
  
  sCod = "000000"
  If Trim(cCP.Text) <> "" Then sCod = Mid(cCP.Text, 1, 6)
      
  s = "SELECT * FROM SubClientes WHERE Cliente = " & sCod & " ORDER BY Id"
  
  r.Open s, DBConexionSQL, adOpenDynamic, adLockOptimistic
  l = 1
  Do While Not r.EOF
    s = Zeros(r.Fields("id").Value, 6) & " : " & Trim(r.Fields("nombre").Value)
    cSC.AddItem s
    r.MoveNext
  Loop
  r.Close
  
  If cSC.ListCount > 0 Then cSC.ListIndex = 0
  
End Sub


Private Sub Form_Load()
  Dim UI As String
  
  Frame2.Visible = False
  List1.Clear
  cCP.Clear
  cSC.Clear
  
  UI = Mid(App.Path, 1, 2)
    
  Dir1.Path = UI
  
  File1.Refresh
  'CargarArchivosDeDisco
  
  Cargar_Clientes
  
  Check2.Value = vbUnchecked
  

  
End Sub
