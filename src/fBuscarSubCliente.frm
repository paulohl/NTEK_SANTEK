VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form fBuscarSubCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Sub-Cliente"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bAceptar 
      Caption         =   "Aceptar"
      Height          =   500
      Left            =   3240
      Picture         =   "fBuscarSubCliente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6720
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   500
      Left            =   4410
      Picture         =   "fBuscarSubCliente.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6720
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "fBuscarSubCliente.frx":0B14
      Height          =   5475
      Left            =   0
      TabIndex        =   5
      Top             =   1110
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9657
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   150
      Top             =   6750
      Width           =   2145
      _ExtentX        =   3784
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
      Caption         =   "Tabla"
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
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8565
      Begin VB.CommandButton bBuscar 
         Height          =   345
         Left            =   5910
         Picture         =   "fBuscarSubCliente.frx":0B29
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   510
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   150
         Width           =   1755
      End
      Begin VB.TextBox eBuscar 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   540
         Width           =   4500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Buscar SUB-CLIENTE"
         Height          =   255
         Left            =   6720
         TabIndex        =   10
         Top             =   270
         Width           =   1650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Buscar Por:"
         Height          =   195
         Left            =   420
         TabIndex        =   3
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Texto a Buscar:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   570
         Width           =   1125
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   195
      Left            =   7140
      TabIndex        =   4
      Top             =   6570
      Width           =   45
   End
End
Attribute VB_Name = "fBuscarSubCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EncabezadoFG As String

Private Sub BuscarValor()
  Dim f As Integer, c As Integer
  Dim Hay As Boolean
  Dim s As String, sTabla As String
  Dim sSQL As String
  Dim sBuscar As String
      
  sBuscar = Trim(eBuscar.Text)
  If sBuscar = "" Then
    s = "select codigo, nombre, direccion, telefonos from " & Adodc1.Caption & " order by " & Combo1.Text
  Else
    If InStr(sBuscar, "*") > 0 Then Mid(sBuscar, InStr(sBuscar, "*"), 1) = "%"
    If InStr(sBuscar, "*") > 0 Then Mid(sBuscar, InStr(sBuscar, "*"), 1) = "%"
    s = "select codigo, nombre, direccion, telefonos from " & Adodc1.Caption & " where " & Combo1.Text & " like '" & sBuscar & "'"
  End If
  
  Adodc1.RecordSource = s
  Adodc1.Refresh
  DataGrid1.Refresh
  AjustaColumnaDataGrid DataGrid1, "NOMBRE", 4000
End Sub

Private Sub bAceptar_Click()
  If Not Adodc1.Recordset.EOF Then
    Modulo.vTemporal1 = Zeros(Adodc1.Recordset.Fields("id").Value, 6)
    Unload Me
  End If
End Sub

Private Sub bBuscar_Click()
  BuscarValor
End Sub

Private Sub bCancelar_Click()
  Unload Me
End Sub

Private Sub DataGrid1_DblClick()
  If Not Adodc1.Recordset.EOF Then
    Modulo.vTemporal1 = Zeros(Adodc1.Recordset.Fields("id").Value, 6)
    Unload Me
  End If
End Sub

Private Sub eBuscar_Change()
  EnMayusculas eBuscar
  'Call BuscarSimple
End Sub

Private Sub eBuscar_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then BuscarValor
End Sub

Private Sub Form_Load()
  eBuscar.Text = ""
  EncabezadoFG = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Unload Me
End Sub

Function Total_Clientes() As Long
  Dim r As New ADODB.Recordset
  Dim ccn As Long
  ccn = 0
  r.Open "SELECT count(*) FROM clientes", DBConexionSQL, adOpenKeyset, adLockReadOnly
  If Not r.EOF Then
    If Not IsNull(r.Fields(0).Value) Then
      ccn = r.Fields(0).Value
    End If
  End If
  r.Close
  Set r = Nothing
  Total_Clientes = ccn
End Function


Sub Cargar_Clientes()
  Dim r As New ADODB.Recordset
  Dim l As Integer
  Dim s As String
  Dim sorden As String
  
  sorden = "codigo"
  If Option1.Value Then sorden = "codigo" Else
  If Option2.Value Then sorden = "nombre" Else
  If Option3.Value Then sorden = "rif"
  
  FG.Clear
  FG.FormatString = EncabezadoFG
  FG.Rows = 2
  FG.FixedRows = 1
   
  s = "SELECT * FROM clientes ORDER BY " & sorden
  r.Open s, DBConexionSQL, adOpenDynamic, adLockOptimistic
  l = 1
  Do While Not r.EOF
    FG.TextMatrix(l, 0) = Zeros(r.Fields("codigo").Value, 6)
    FG.TextMatrix(l, 1) = Trim(r.Fields("nombre").Value)
    FG.TextMatrix(l, 2) = Trim(r.Fields("rif").Value)
    FG.Row = l: FG.Col = 2: FG.CellAlignment = flexAlignLeftCenter
    FG.TextMatrix(l, 3) = Trim(r.Fields("telefonos").Value)
    FG.Row = l: FG.Col = 3: FG.CellAlignment = flexAlignLeftCenter
    
    r.MoveNext
    
    FG.Col = 0
    
    If Not r.EOF Then
      l = l + 1
      FG.Rows = FG.Rows + 1
    End If
  Loop
  
  FG.Refresh

  
  Label3.Caption = "[" & CStr(Total_Clientes()) & " Regs.]"
  'RClientes.Close
  'Set r = Nothing
  Viendo = True
End Sub

Private Sub Option1_Click()
  Cargar_Clientes
  If fBuscarSimple.Visible Then eBuscar.SetFocus
  eBuscar.Text = ""
End Sub

Private Sub Option2_Click()
  Cargar_Clientes
  If fBuscarSimple.Visible Then eBuscar.SetFocus
  eBuscar.Text = ""
End Sub

Private Sub Option3_Click()
  Cargar_Clientes
  If fBuscarSimple.Visible Then eBuscar.SetFocus
  eBuscar.Text = ""
End Sub

