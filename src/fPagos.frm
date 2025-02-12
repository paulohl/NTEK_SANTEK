VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form fPagos 
   Caption         =   "Registro de Pagos"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Pagos Registrados por Cliente"
      Height          =   9435
      Left            =   9570
      TabIndex        =   7
      Top             =   0
      Width           =   5655
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00808000&
         Height          =   885
         Left            =   120
         ScaleHeight     =   825
         ScaleWidth      =   5355
         TabIndex        =   25
         Top             =   8160
         Width           =   5415
         Begin VB.Label lTC 
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1230
            TabIndex        =   31
            Top             =   90
            Width           =   930
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Débitos Bs:"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   30
            TabIndex        =   30
            Top             =   90
            Width           =   1215
         End
         Begin VB.Label lTA 
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   4230
            TabIndex        =   29
            Top             =   90
            Width           =   930
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Créditos Bs:"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   2940
            TabIndex        =   28
            Top             =   90
            Width           =   1245
         End
         Begin VB.Label lTS 
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   2790
            TabIndex        =   27
            Top             =   570
            Width           =   930
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Saldo Bs:"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   1650
            TabIndex        =   26
            Top             =   570
            Width           =   1080
         End
      End
      Begin VB.CommandButton bRefrescar 
         Caption         =   "Mostrar Pagos"
         Height          =   525
         Left            =   3900
         Picture         =   "fPagos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1140
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   345
         Left            =   2160
         Top             =   5610
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   609
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Pagos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00808000&
         Height          =   465
         Left            =   1350
         ScaleHeight     =   405
         ScaleWidth      =   2235
         TabIndex        =   21
         Top             =   7320
         Width           =   2295
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1230
            TabIndex        =   23
            Top             =   90
            Width           =   930
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Pagos Bs:"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   30
            TabIndex        =   22
            Top             =   90
            Width           =   1125
         End
      End
      Begin MSComCtl2.DTPicker eDesde 
         Height          =   285
         Left            =   2430
         TabIndex        =   18
         Top             =   1110
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Format          =   55640065
         CurrentDate     =   40015
      End
      Begin VB.CheckBox cFiltrar 
         Caption         =   "Filtrar Fecha:"
         Height          =   195
         Left            =   420
         TabIndex        =   16
         Top             =   1320
         Width           =   1245
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   630
         Width           =   3975
      End
      Begin VB.CommandButton bAgregar 
         Caption         =   "Agregar Pagos"
         Height          =   525
         Left            =   4650
         TabIndex        =   10
         Top             =   2370
         Width           =   885
      End
      Begin VB.CommandButton bEditarPagos 
         Caption         =   "Editar"
         Height          =   525
         Left            =   4650
         TabIndex        =   9
         Top             =   3180
         Width           =   885
      End
      Begin VB.CommandButton bEliminar 
         Caption         =   "Eliminar"
         Height          =   525
         Left            =   4650
         TabIndex        =   8
         Top             =   3990
         Width           =   885
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "fPagos.frx":0A02
         Height          =   5445
         Left            =   210
         TabIndex        =   11
         Top             =   1800
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   9604
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker eHasta 
         Height          =   285
         Left            =   2430
         TabIndex        =   20
         Top             =   1410
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Format          =   55640065
         CurrentDate     =   40015
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   1920
         TabIndex        =   19
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   1890
         TabIndex        =   17
         Top             =   1140
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   420
         TabIndex        =   15
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lCliente 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         Height          =   255
         Left            =   1050
         TabIndex        =   14
         Top             =   330
         Width           =   4005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Subcliente:"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   690
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9555
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   345
         Left            =   240
         Top             =   9060
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Clientes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "fPagos.frx":0A17
         Height          =   8055
         Left            =   90
         TabIndex        =   6
         Top             =   990
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   14208
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton bBuscar 
         Caption         =   "buscar"
         Height          =   465
         Left            =   6120
         Picture         =   "fPagos.frx":0A2C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   1635
      End
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   2640
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   300
         Width           =   3345
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "=>"
         Height          =   195
         Left            =   2340
         TabIndex        =   4
         Top             =   300
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buscar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Visible         =   0   'False
         Width           =   810
      End
   End
End
Attribute VB_Name = "fPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bDetallando As Boolean
Dim Viendo As Boolean


Public Sub Clientes_Format_DataGrid()
  Dim i As Integer
  With fPagos
    For i = 0 To .DataGrid1.Columns.Count - 1
      .DataGrid1.Columns(i).Caption = UCase(.DataGrid1.Columns(i).Caption)
    Next i
    .DataGrid1.Columns(1).Visible = False
    .DataGrid1.Columns(2).Visible = False
    .DataGrid1.Columns(4).Visible = False
    .DataGrid1.Columns(6).Visible = False
    .DataGrid1.Columns(7).Visible = False
    .DataGrid1.Columns(8).Visible = False
    .DataGrid1.Columns(9).Visible = False
    .DataGrid1.Columns(10).Visible = False
    .DataGrid1.Columns(11).Visible = False
    .DataGrid1.Columns(12).Visible = False
    .DataGrid1.Columns(13).Visible = False
    .DataGrid1.Columns(14).Visible = False
    .DataGrid1.Columns(15).Visible = False
    '.DataGrid1.Columns(16).Visible = False
        
    .DataGrid1.Columns(0).width = 700 'codigo
    .DataGrid1.Columns(3).width = 5000 'nombre
    .DataGrid1.Columns(5).width = 3000  'telefono
    
    
    
'    .DataGrid1.Columns(3).width = 900  'precio
'    .DataGrid1.Columns(4).width = 1000 'creado
'
'    .DataGrid1.Columns(0).NumberFormat = "#,0.00"
'    .DataGrid1.Columns(3).NumberFormat = "#,0.00"
'
'    .DataGrid1.Columns(2).Alignment = dbgRight
'    .DataGrid1.Columns(3).Alignment = dbgRight
  End With
End Sub

Public Sub Pagos_Format_DataGrid()
  Dim i As Integer
  With fPagos
    For i = 0 To .DataGrid2.Columns.Count - 1
      .DataGrid2.Columns(i).Caption = UCase(.DataGrid2.Columns(i).Caption)
    Next i
    .DataGrid2.Columns(1).Visible = False
    .DataGrid2.Columns(2).Visible = False
        
    .DataGrid2.Columns(0).width = 500  'id
    .DataGrid2.Columns(3).width = 1000 'fecha
    .DataGrid2.Columns(4).width = 1200 'monto
    .DataGrid2.Columns(5).width = 1000 'monto
    
    .DataGrid2.Columns(4).NumberFormat = "#,0.00"
    .DataGrid2.Columns(4).Alignment = dbgRight
    
    '.DataGrid2.Columns(5).NumberFormat = "#,0.00"
    '.DataGrid2.Columns(5).Alignment = dbgRight
    
    
    .DataGrid2.Columns(5).Visible = False
    .DataGrid2.Columns(6).Visible = False
    .DataGrid2.Columns(7).Visible = False
    '.DataGrid2.Columns(8).Visible = False
 
  End With
End Sub

Private Sub Listar_Pagos(bSubClienteSeleccionado As Boolean)
  Dim s As String
  Dim sFD As String, sFH As String
  Dim lSC As Long
  
  
    sFD = Format(eDesde.value, "yyyymmdd")
    sFH = Format(eHasta.value, "yyyymmdd")
    
    If Not Modulo.TIENE_Subcliente(Adodc1.Recordset.Fields("codigo").value) Then   'No tiene subcliente:
      
      If cFiltrar.value = vbChecked Then
      
        s = "select * from pagos where " & _
            "cliente = " & CStr(Adodc1.Recordset.Fields("codigo").value) & " and " & _
            "fecha  >= '" & sFD & "' and fecha <= '" & sFH & "' " & _
            "order by id"
            
      Else
      
        s = "select * from pagos where " & _
            "cliente = " & CStr(Adodc1.Recordset.Fields("codigo").value) & " " & _
            "order by id"
            
      End If
      
    Else  'Tiene al menos un (01) subcliente:
    
      If cFiltrar.value = vbChecked Then
      
        If Not bSubClienteSeleccionado Then
        
          s = "select * from pagos where " & _
              "cliente    = " & CStr(Adodc1.Recordset.Fields("codigo").value) & " and " & _
              "subcliente = " & CStr(lSC) & " and " & _
              "fecha     >= '" & sFD & "' and fecha <= '" & sFH & "' " & _
              "order by id"
              
        Else
          
          s = "select * from pagos where " & _
              "cliente    = " & CStr(Adodc1.Recordset.Fields("codigo").value) & " and " & _
              "subcliente = " & Mid(Combo2.Text, 1, 6) & " and " & _
              "fecha     >= '" & sFD & "' and fecha <= '" & sFH & "' " & _
              "order by id"
              
        End If
        
            
      Else
      
        If bSubClienteSeleccionado = False Then
      
          s = "select * from pagos where " & _
              "cliente    = " & CStr(Adodc1.Recordset.Fields("codigo").value) & " and " & _
              "subcliente = " & CStr(lSC) & " " & _
              "order by id"
              
        Else
          
          s = "select * from pagos where " & _
              "cliente    = " & CStr(Adodc1.Recordset.Fields("codigo").value) & _
               " and subcliente = " & Mid(Combo2.Text, 1, 6) & " " & _
               "order by id"
               

        End If
            
      End If
      
    End If
    
    Adodc2.RecordSource = s
    Adodc2.Refresh
    DataGrid2.Refresh
    Pagos_Format_DataGrid
  
End Sub

Private Sub CargarSubclientes()
  Dim r As New ADODB.Recordset
  Dim s As String
  
  Combo2.Clear
  
  If Not Adodc1.Recordset.EOF Then
    lCliente.Caption = Zeros(Adodc1.Recordset.Fields("codigo").value, 6) & " " & Trim(Adodc1.Recordset.Fields("nombre").value)
    
    lTC.Caption = Format(Adodc1.Recordset.Fields("deuda").value, "#,0.00")
    lTA.Caption = Format(Adodc1.Recordset.Fields("pagos").value, "#,0.00")
    lTS.Caption = Format(Adodc1.Recordset.Fields("abonos").value, "#,0.00")
    
    If Modulo.TIENE_Subcliente(Adodc1.Recordset.Fields("codigo").value) Then
      s = "select * from subclientes where cliente = " & CStr(Adodc1.Recordset.Fields("codigo").value) & " order by id"
      r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
      Combo2.Clear
      Do While Not r.EOF
        Combo2.AddItem Zeros(r.Fields("id").value, 6) & " " & Trim(r.Fields("nombre").value)
        r.MoveNext
      Loop
      If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
      r.Close
    End If
  End If
  
  Set r = Nothing
End Sub


'Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  If bDetallando Then Listar_Pagos False
'  CargarSubclientes
'  Totalizar_Pagos
'End Sub

Private Sub bAgregar_Click()
  Dim r As New ADODB.Recordset
  Dim s As String
  
  bRefrescar_Click
  
  If Adodc1.Recordset.RecordCount <= 0 Then Exit Sub
  If Adodc1.Recordset.EOF Then Exit Sub
  
  Load fPagos2
  With fPagos2
    .Caption = "AGREGAR PAGO DE CLIENTE"
    .lCliente.Caption = Zeros(Adodc1.Recordset.Fields("codigo").value, 6) & " " & Trim(Adodc1.Recordset.Fields("nombre").value)
    .Combo1.Clear
    s = "select id, cliente, nombre from subclientes where cliente = " & CStr(Adodc1.Recordset.Fields("codigo").value) & " order by id"
    r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    Do While Not r.EOF
      .Combo1.AddItem Zeros(r.Fields("id").value, 6) & " " & Trim(r.Fields("nombre").value)
      r.MoveNext
    Loop
    If .Combo1.ListCount > 0 Then .Combo1.ListIndex = 0 Else .Combo1.Enabled = False
    .eFecha.value = Date
    .ePre.Text = "0,00"
    .Option1.value = False
    .Option2.value = True
    .eNumero.Text = ""
    .eBanco.Text = ""
    '.eObservaciones.Text = ""
    
    If Combo2.ListCount > 0 Then .Combo1.ListIndex = Combo2.ListIndex
    
  End With
  fPagos2.Show vbModal
End Sub

Private Sub bBuscar_Click()
  Dim s As String, sSQL As String
  Text1.Text = UCase(Text1.Text)
  s = Trim(Text1.Text)
  If s <> "" Then
    If InStr(s, "*") > 0 Then Mid(s, InStr(s, "*"), 1) = "%"
    If InStr(s, "*") > 0 Then Mid(s, InStr(s, "*"), 1) = "%"
    
    If Combo1.Text = "CÓDIGO" And IsNumeric(s) = True Then
      sSQL = "select * from Clientes where Codigo = '" & s & "' order by codigo"
    Else
      Combo1.Text = "NOMBRE"
      sSQL = "select * from Clientes where Nombre LIKE '%" & s & "%' order by nombre"
    End If
    
  Else
  
    If Combo1.Text = "CÓDIGO" Then
      sSQL = "select * from Clientes order by Codigo"
    Else
      sSQL = "select * from Clientes order by Nombre"
    End If
    
  End If
  
  Adodc1.RecordSource = sSQL
  Adodc1.Refresh
  DataGrid1.Refresh
  Clientes_Format_DataGrid
  
End Sub

Private Sub bRefrescar_Click()
  Dim dTC As Double, dTA As Double, dTS As Double
  If Combo2.Text = "" Then
     Listar_Pagos False
  Else
     Listar_Pagos True
  End If
  Totalizar_Pagos
  
  
  dTC = 0#
  dTA = 0#
  dTS = 0#
  Modulo.Resumen_Cuenta_Cliente lCliente.Caption, Combo2.Text, _
                                dTC, dTA, dTS

  lTC.Caption = Format(dTC, "#,0.00")
  lTA.Caption = Format(dTA, "#,0.00")
  lTS.Caption = Format(dTS, "#,0.00")
  'Combo2.ListIndex = -1
  
End Sub

Private Sub bEditarPagos_Click()
  Dim r As New ADODB.Recordset
  Dim s As String
  
  If Adodc1.Recordset.RecordCount <= 0 Then Exit Sub
  If Adodc1.Recordset.EOF Then Exit Sub
  
  Load fPagos2
  With fPagos2
    .Caption = "EDITAR PAGO DE CLIENTE - ID " & CStr(Adodc2.Recordset.Fields("id").value)
    .lCliente.Caption = Zeros(Adodc1.Recordset.Fields("codigo").value, 6) & " " & Trim(Adodc1.Recordset.Fields("nombre").value)
    .Combo1.Clear
    s = "select id, cliente, nombre from subclientes where cliente = " & CStr(Adodc1.Recordset.Fields("codigo").value) & " order by id"
    r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    Do While Not r.EOF
      .Combo1.AddItem Zeros(r.Fields("id").value, 6) & " " & Trim(r.Fields("nombre").value)
      r.MoveNext
    Loop
    'If .Combo1.ListCount > 0 Then .Combo1.ListIndex = 0 Else .Combo1.Enabled = False
    '.Combo1.ListIndex = Modulo.Buscar_ComboLen(.Combo1, Mid(lCliente.Caption, 1, 6), 6)
    .Label4.Caption = CStr(Adodc2.Recordset.Fields("id").value)
    
    .eFecha.value = Adodc2.Recordset.Fields("fecha").value
    .ePre.Text = Format(Adodc2.Recordset.Fields("monto").value, "#,0.00")
    
    Modulo.vMontoAnterior = Adodc2.Recordset.Fields("monto").value
    
    .Option1.value = False
    .Option2.value = False
    If Adodc2.Recordset.Fields("tipo").value = "C" Then .Option2.value = True Else .Option1.value = True
    .eNumero.Text = Trim(Adodc2.Recordset.Fields("numero").value)
    .eBanco.Text = Trim(Adodc2.Recordset.Fields("banco").value)
'    .eObservaciones.Text = Trim(Adodc2.Recordset.Fields("observaciones").Value)
  End With
  fPagos2.Show
End Sub

Private Sub bEliminar_Click()
  Dim s As String, s1 As String, s2 As String, s3 As String
  Dim dMonto As Double
  
  If Adodc2.Recordset.RecordCount > 0 Then
    If Not Adodc2.Recordset.EOF Then
    
      dMonto = Adodc2.Recordset.Fields("monto").value
      
      If MsgBox("¿Está Seguro de Borrar el Pago ID#" & CStr(Adodc2.Recordset.Fields("id").value) & "?" & vbCrLf & _
                " Por un Monto de Bs." & Format(dMonto, "#,0.00"), vbQuestion + vbYesNo, "Confirme") = vbYes Then
                
                
        s1 = Mid(lCliente.Caption, 1, 6)
    
        s2 = Trim(Combo2.Text)
        If Trim(s2) = "" Then s2 = "0"
    
        If s2 <> "0" Then s2 = Mid(s2, 1, 6)
                
        Modulo.Actualizar_Pago_Cliente CLng(s1), CLng(s2), (dMonto * -1#)
               
        s = "delete from pagos where id = " & CStr(Adodc2.Recordset.Fields("id").value)
        Modulo.ExecSQL s
        Adodc2.Refresh
        DataGrid2.Refresh
        
        Pagos_Format_DataGrid
      End If
    End If
  End If
End Sub

Private Sub CargarSubCliente(sCliente As String, xCombo As ComboBox)
  Dim s As String
  Dim r As New ADODB.Recordset
  s = "select * from subclientes where cliente = " & sCliente & " order by id"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  xCombo.Clear
  Do While Not r.EOF
    s = Zeros(r.Fields("id").value, 6) & " " & Trim(r.Fields("nombre").value)
    xCombo.AddItem s
    r.MoveNext
  Loop
  If xCombo.ListCount > 0 Then xCombo.ListIndex = -1
  r.Close
  Set r = Nothing
End Sub

Private Sub Combo2_Change()
  bRefrescar_Click
End Sub

Public Sub Combo2_Click()
  bRefrescar_Click
  
    Dim dTC As Double, dTA As Double, dTS As Double
  Modulo.Resumen_Cuenta_Cliente lCliente.Caption, Combo2.Text, _
                                dTC, dTA, dTS

  lTC.Caption = Format(dTC, "#,0.00")
  lTA.Caption = Format(dTA, "#,0.00")
  lTS.Caption = Format(dTS, "#,0.00")

End Sub

Public Sub DataGrid1_Click()
  Dim dTC As Double, dTA As Double, dTS As Double
  'Modulo.Resumen_Cuenta_Cliente lCliente.Caption, Combo2.Text, _
                                dTC, dTA, dTS
 
  Resumen_Cuenta_Cliente_Codigo Adodc1.Recordset.Fields("codigo").value, dTC, dTA, dTS
  
  lTC.Caption = Format(dTC, "#,0.00")
  lTA.Caption = Format(dTA, "#,0.00")
  lTS.Caption = Format(dTS, "#,0.00")
  bRefrescar_Click
End Sub

Private Sub DataGrid1_DblClick()
  Dim dTC As Double, dTA As Double, dTS As Double
  Modulo.Resumen_Cuenta_Cliente lCliente.Caption, Combo2.Text, _
                                dTC, dTA, dTS

  lTC.Caption = Format(dTC, "#,0.00")
  lTA.Caption = Format(dTA, "#,0.00")
  lTS.Caption = Format(dTS, "#,0.00")

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim dTC As Double, dTA As Double, dTS As Double
  Modulo.Resumen_Cuenta_Cliente lCliente.Caption, Combo2.Text, _
                                dTC, dTA, dTS

  lTC.Caption = Format(dTC, "#,0.00")
  lTA.Caption = Format(dTA, "#,0.00")
  lTS.Caption = Format(dTS, "#,0.00")

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
  Dim dTC As Double, dTA As Double, dTS As Double
  Modulo.Resumen_Cuenta_Cliente lCliente.Caption, Combo2.Text, _
                                dTC, dTA, dTS

  lTC.Caption = Format(dTC, "#,0.00")
  lTA.Caption = Format(dTA, "#,0.00")
  lTS.Caption = Format(dTS, "#,0.00")

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

  Dim dTC As Double, dTA As Double, dTS As Double
  'Modulo.Resumen_Cuenta_Cliente lCliente.Caption, Combo2.Text, _
                                dTC, dTA, dTS

  
  
  
  Resumen_Cuenta_Cliente_Codigo Adodc1.Recordset.Fields("codigo").value, dTC, dTA, dTS
  
  lTC.Caption = Format(dTC, "#,0.00")
  lTA.Caption = Format(dTA, "#,0.00")
  lTS.Caption = Format(dTS, "#,0.00")


  If Not Adodc1.Recordset.EOF Then
    lCliente.Caption = Zeros(Adodc1.Recordset.Fields("codigo").value, 6) & " " & Trim(Adodc1.Recordset.Fields("nombre").value)
       
    Combo2.Clear
    If Modulo.TIENE_Subcliente(Adodc1.Recordset.Fields("codigo").value) Then
      CargarSubCliente Adodc1.Recordset.Fields("codigo").value, Combo2
    End If
  End If
End Sub

Private Sub Form_Load()

 Dim s As String
  Load fPagos
  With fPagos
    .Combo1.Clear
    .Combo1.AddItem "CÓDIGO"
    .Combo1.AddItem "NOMBRE"
    .Combo1.ListIndex = 1
    
    s = "Select * from Clientes order by codigo"
    
    .Adodc1.ConnectionString = Modulo.DBConexionSQL.ConnectionString
    .Adodc1.RecordSource = s
    .Adodc1.Refresh
    .Clientes_Format_DataGrid
    '.bBuscar_Click
  End With

  bDetallando = False
  Text1.Text = ""
  Adodc2.ConnectionString = Modulo.DBConexionSQL.ConnectionString
  bDetallando = True
  eDesde.value = Date
  eHasta.value = Date
  
  lTC.Caption = "0,00"
  lTA.Caption = "0,00"
  lTS.Caption = "0,00"
  
  'Adodc1.Refresh
  
End Sub


Private Sub lCliente_Click()
  Dim dTC As Double, dTA As Double, dTS As Double
  Modulo.Resumen_Cuenta_Cliente lCliente.Caption, Combo2.Text, _
                                dTC, dTA, dTS

  lTC.Caption = Format(dTC, "#,0.00")
  lTA.Caption = Format(dTA, "#,0.00")
  lTS.Caption = Format(dTS, "#,0.00")

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 10 Or KeyAscii = 13 Then bBuscar_Click
End Sub

Private Sub Text1_LostFocus()
  Text1.Text = UCase(Text1.Text)
End Sub

Public Sub Totalizar_Pagos()
  Dim idactual As Long, dT As Double
  dT = 0#
  If Adodc2.Recordset.RecordCount > 0 Then
    idactual = Adodc2.Recordset.Fields("id").value
    Adodc2.Recordset.MoveFirst
    dT = 0#
    Do While Not Adodc2.Recordset.EOF
      dT = dT + Adodc2.Recordset.Fields("monto").value
      Adodc2.Recordset.MoveNext
    Loop
    Adodc2.Recordset.MoveFirst
    Adodc2.Recordset.Find "id = " & CStr(idactual)
  End If
  Label7.Caption = Format(dT, "#,0.00")
End Sub
