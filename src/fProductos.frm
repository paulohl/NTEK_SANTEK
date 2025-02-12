VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form fProductos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   7155
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9795
      Begin VB.CommandButton bEliminar 
         Caption         =   "Eliminar"
         Height          =   525
         Left            =   8610
         TabIndex        =   9
         Top             =   2790
         Width           =   885
      End
      Begin VB.CommandButton bEditar 
         Caption         =   "Editar"
         Height          =   525
         Left            =   8610
         TabIndex        =   8
         Top             =   1860
         Width           =   885
      End
      Begin VB.CommandButton bAgregar 
         Caption         =   "Agregar"
         Height          =   525
         Left            =   8610
         TabIndex        =   7
         Top             =   990
         Width           =   885
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   345
         Left            =   330
         Top             =   6690
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
         Caption         =   "Productos"
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
         Bindings        =   "fProductos.frx":0000
         Height          =   5805
         Left            =   60
         TabIndex        =   6
         Top             =   750
         Width           =   8325
         _ExtentX        =   14684
         _ExtentY        =   10239
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
         Height          =   345
         Left            =   6120
         Picture         =   "fProductos.frx":0015
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   2640
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   3345
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "=>"
         Height          =   195
         Left            =   2340
         TabIndex        =   4
         Top             =   300
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buscar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   810
      End
   End
End
Attribute VB_Name = "fProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Productos_Format_DataGrid()
  With fProductos
    .DataGrid1.Columns(0).width = 1000 'codigo
    .DataGrid1.Columns(1).width = 3000 'descripcion
    .DataGrid1.Columns(2).width = 900  'existencia
    .DataGrid1.Columns(3).width = 900  'precio
    .DataGrid1.Columns(4).width = 1000 'creado
    
    .DataGrid1.Columns(2).NumberFormat = "#,0.00"
    .DataGrid1.Columns(3).NumberFormat = "#,0.00"
    
    .DataGrid1.Columns(2).Alignment = dbgRight
    .DataGrid1.Columns(3).Alignment = dbgRight
  End With
End Sub

Private Sub bAgregar_Click()
  Load fProductosAg
  With fProductosAg
    .Caption = "AGREGAR PRODUCTO"
    .eCod.Text = ""
    .eDes.Text = ""
    .ePre.Text = "0,00"
    .eExi.Text = "0"
    .eCre.Text = Format(Now, "dd/mm/yyyy hh:mm ampm")
    .Frame2.Visible = False
  End With
  fProductosAg.Show
End Sub

Private Sub bBuscar_Click()
  Dim s As String, sSQL As String
  Text1.Text = UCase(Text1.Text)
  s = Trim(Text1.Text)
  If s <> "" Then
    If InStr(s, "*") > 0 Then Mid(s, InStr(s, "*"), 1) = "%"
    If InStr(s, "*") > 0 Then Mid(s, InStr(s, "*"), 1) = "%"
    
    If Combo1.Text = "CÓDIGO" Then
      sSQL = "select * from Productos where Codigo = '" & s & "' order by codigo"
    Else
      sSQL = "select * from Productos where Descripcion LIKE '" & s & "' order by Descripcion"
    End If
    
  Else
  
    If Combo1.Text = "CÓDIGO" Then
      sSQL = "select * from Productos order by Codigo"
    Else
      sSQL = "select * from Productos order by Descripcion"
    End If
    
  End If
  
  Adodc1.RecordSource = sSQL
  Adodc1.Refresh
  DataGrid1.Refresh
  Productos_Format_DataGrid
  
End Sub

Private Sub bEditar_Click()
  If Not Adodc1.Recordset.EOF Then
    Load fProductosAg
    With fProductosAg
      .Caption = "EDITAR PRODUCTO"
      .eCod.Text = Trim(Adodc1.Recordset.Fields("codigo").Value)
      .eDes.Text = Trim(Adodc1.Recordset.Fields("descripcion").Value)
      .ePre.Text = Format(Adodc1.Recordset.Fields("precio").Value, "#,0.00")
      .eExi.Text = Format(Adodc1.Recordset.Fields("existencia").Value, "#,0.00")
      .eCre.Text = Format(Adodc1.Recordset.Fields("creado").Value, "dd/mm/yyyy hh:mm ampm")
      .eAD.Text = "0"
      .Frame2.Visible = True
      .eExi.Enabled = False
      .eCod.Enabled = False
    End With
    fProductosAg.Show
  End If
End Sub

Private Sub bEliminar_Click()
  Dim s As String, sSQL As String
  If Not Adodc1.Recordset.EOF Then
    s = Trim(Adodc1.Recordset.Fields("codigo").Value)
    If s <> "" Then
      If MsgBox("¿Está Seguro de Eliminar el Artículo:" & vbCrLf & _
                 Trim(Adodc1.Recordset.Fields("codigo").Value) & vbCrLf & _
                 Trim(Adodc1.Recordset.Fields("descripcion").Value), vbQuestion + vbYesNo, "Confirme") = vbYes Then
        sSQL = "delete from productos where codigo = '" & s & "'"
        Modulo.ExecSQL sSQL
        
        MsgBox "Artículo Borrado Correctamente...", vbInformation, "Información"
        
        AgregarLogs "Borra Producto [" & s & "]"
        
        Adodc1.Refresh
        DataGrid1.Refresh
        Productos_Format_DataGrid
      End If
    End If
  End If
End Sub

Private Sub Form_Load()
  Text1.Text = ""
End Sub

Private Sub Text1_LostFocus()
  Text1.Text = UCase(Text1.Text)
End Sub
