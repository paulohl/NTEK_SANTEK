VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form fFD 
   Caption         =   "Formato de Diseño"
   ClientHeight    =   10290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Caption         =   "Datos del Cliente"
      Height          =   1875
      Left            =   0
      TabIndex        =   27
      Top             =   930
      Width           =   7095
      Begin VB.TextBox tRif 
         Height          =   315
         Left            =   4560
         MaxLength       =   20
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1020
         Width           =   2400
      End
      Begin VB.TextBox tdir 
         Height          =   315
         Left            =   870
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   6100
      End
      Begin VB.TextBox ttel 
         Height          =   315
         Left            =   870
         MaxLength       =   50
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   630
         Width           =   3150
      End
      Begin VB.TextBox tfax 
         Height          =   315
         Left            =   4560
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   630
         Width           =   2400
      End
      Begin VB.TextBox temail 
         Height          =   315
         Left            =   870
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1020
         Width           =   3150
      End
      Begin VB.TextBox tcon 
         Height          =   315
         Left            =   870
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1410
         Width           =   3150
      End
      Begin VB.TextBox tcontelf 
         Height          =   315
         Left            =   4560
         MaxLength       =   50
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1410
         Width           =   2400
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "RIF:"
         Height          =   195
         Left            =   4200
         TabIndex        =   36
         Top             =   1050
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   90
         TabIndex        =   33
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Teléfonos:"
         Height          =   195
         Left            =   30
         TabIndex        =   32
         Top             =   660
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fax:"
         Height          =   195
         Left            =   4200
         TabIndex        =   31
         Top             =   660
         Width           =   300
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "EMail:"
         Height          =   195
         Left            =   330
         TabIndex        =   30
         Top             =   1050
         Width           =   435
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Contacto:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Telf:"
         Height          =   195
         Left            =   4170
         TabIndex        =   28
         Top             =   1440
         Width           =   315
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   -30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Caption         =   "Productos y Precios según Negociación"
      Height          =   3705
      Left            =   0
      TabIndex        =   20
      Top             =   6450
      Width           =   7065
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "0"
         Top             =   2520
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListvProductos 
         Height          =   2055
         Left            =   60
         TabIndex        =   37
         Top             =   300
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Normal"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Acordado"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "SubTotal"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Entregado"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Pagado"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton bAceptar 
         Caption         =   "Guardar"
         Height          =   500
         Left            =   4920
         Picture         =   "fFD.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3120
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton bCancelar 
         Caption         =   "Cancelar"
         Height          =   500
         Left            =   6090
         Picture         =   "fFD.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3120
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.ListBox aProductos 
         Height          =   1230
         Left            =   4680
         TabIndex        =   24
         Top             =   840
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton bVerSoloPrecios 
         Caption         =   "Ver Sólo Productos Seleccionados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   1500
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2910
         Width           =   1860
      End
      Begin VB.CommandButton bVerTodo 
         Caption         =   "Ver Todos los Productos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2910
         Width           =   1110
      End
      Begin VB.ListBox lProductos 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   3630
         Style           =   1  'Checkbox
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   3345
      End
      Begin VB.Label Label3 
         Caption         =   "Total:"
         Height          =   195
         Left            =   4680
         TabIndex        =   38
         Top             =   2520
         Width           =   675
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Imagen Formato Diseño"
      Height          =   10155
      Left            =   7110
      TabIndex        =   19
      Top             =   0
      Width           =   8130
      Begin VB.Image Image1 
         Height          =   9855
         Left            =   60
         Stretch         =   -1  'True
         Top             =   240
         Width           =   8025
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ID          Formatos Registrados / Vencimiento"
      Height          =   3555
      Left            =   3660
      TabIndex        =   15
      Top             =   2850
      Width           =   3435
      Begin VB.ComboBox cmbIDdisenoActivo 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   3180
         Width           =   1215
      End
      Begin VB.ListBox ListID 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   60
         TabIndex        =   40
         Top             =   210
         Width           =   615
      End
      Begin VB.ListBox aImagenes 
         Height          =   645
         Left            =   840
         TabIndex        =   25
         Top             =   1680
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.ListBox lVence 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   2100
         TabIndex        =   26
         Top             =   210
         Width           =   1275
      End
      Begin VB.CommandButton bAg 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   150
         TabIndex        =   18
         ToolTipText     =   "Agregar Nuevo Formato Escaneado"
         Top             =   3060
         Width           =   450
      End
      Begin VB.CommandButton bBo 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   750
         TabIndex        =   17
         ToolTipText     =   "Borrar Formato"
         Top             =   3060
         Width           =   450
      End
      Begin VB.ListBox lFormatos 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   720
         TabIndex        =   16
         Top             =   210
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "ID Diseño Activo"
         Height          =   195
         Left            =   2040
         TabIndex        =   41
         Top             =   2940
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Indicaciones Especiales de Trabajo"
      Height          =   3555
      Left            =   0
      TabIndex        =   12
      Top             =   2850
      Width           =   3615
      Begin VB.TextBox eIndicaciones 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Left            =   90
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "fFD.frx":0B14
         Top             =   210
         Width           =   3465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "* Pulsar <CTRL-ENTER> para nueva línea."
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   3300
         Width           =   3090
      End
   End
   Begin VB.CommandButton bBuscarCP 
      Height          =   345
      Left            =   6660
      Picture         =   "fFD.frx":0B1A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.ComboBox cCP 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1650
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4905
   End
   Begin VB.ComboBox cSC 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1650
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   540
      Width           =   4905
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Cliente Principal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   180
      Width           =   1485
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Sub-Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   540
      TabIndex        =   10
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "fFD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ProductosNegociados(1 To 50) As Productos
Dim V_DISENO_BS As Double
Dim V_FOTOGRAFIA_BS As Double
Public HayCambios As Boolean
Dim iCliente
Dim iSCliente
Dim lNegociacionActual As Currency
Dim lCantidadProductosNegociados As Integer ' para saber cuantos productos ya tiene negociado el cliente por si se hace algun cambio
Private Sub sGuardarDatos()
  Dim s As String, s1 As String, i As Integer
  Dim sNomCliente As String, sNomSubCliente As String
  Dim r As New ADODB.Recordset
  Dim sNomArchivoInicial As String, sNomArchivoNuevo As String
  Dim dP As Double, p As Integer, e As Boolean
  Dim s2 As String
  Dim lDeuda As Double
  Dim lPrecioAcordado As String
  Dim sCD As String, sMCD As String, dMCD As Double
  Dim j As Integer
  
  
  If Trim(cCP.Text) = "" Then Exit Sub
  
  s = Mid(cCP.Text, 1, 6)
  iCliente = CLng(s)
  
  If Trim(cSC.Text) = "" Then s = "0" Else s = Mid(cSC.Text, 1, 6)
  iSCliente = CLng(s)
  
  s = "select * from FormatoDiseno where " & _
      "cliente    = " & CStr(iCliente) & " and " & _
      "subcliente = " & CStr(iSCliente) & " "
  
  r.Open s, Modulo.DBConexionSQL, adOpenStatic
  If r.EOF Then 'NO EXISTE... Agregarlo:
    s = "insert into FormatoDiseno " & _
        "(cliente, subcliente, especificaciones) values (" & _
        CStr(iCliente) & "," & CStr(iSCliente) & ",'" & _
        Trim(eIndicaciones.Text) & "')"
    
  Else
    'YA EXISTE... Actualizarlo:
    s = "update FormatoDiseno set " & _
        "especificaciones = '" & Trim(eIndicaciones.Text) & "' " & _
        "where " & _
        "cliente    = " & CStr(iCliente) & " and " & _
        "subcliente = " & CStr(iSCliente) & " "
        
  End If
  
  Modulo.ExecSQL s
    
  
    s = "delete from PreciosEspeciales where " & _
      "cliente    = " & CStr(iCliente) & " and " & _
      "subcliente = " & CStr(iSCliente)
      If ListID.List(ListID.ListIndex) <> "" Then s = s & " And IDDiseño= " & ListID.List(ListID.ListIndex)
      
  Modulo.ExecSQL s

  
  
  '.Ahora procesar los productos con precios acordados:
  '1.borrar los ya registrados:
  lDeuda = 0
  For j = 1 To lCantidadProductosNegociados
    Actualizar_Existencia_Producto ProductosNegociados(j).CodigoProducto, Val(ProductosNegociados(j).Cantidad)
  Next j

  For i = 1 To ListvProductos.ListItems.Count
  
    
    If ListvProductos.ListItems(i).Checked = True Then
       lPrecioAcordado = Replace(ListvProductos.ListItems(i).SubItems(4), ",", ".")
      
       sInsertarProductos ListvProductos.ListItems(i).Text, _
       Val(lPrecioAcordado), _
       ListvProductos.ListItems(i).SubItems(2), _
       ListvProductos.ListItems(i).SubItems(6), 0, _
       ListvProductos.ListItems(i).SubItems(7), Val(ListID.List(ListID.ListIndex))
       If ListvProductos.ListItems(i).SubItems(7) = "SI" Then  ' SI SE PAGÓ EL ITEM
          lDeuda = lDeuda + (Val(ListvProductos.ListItems(i).SubItems(2)) * CDbl(ListvProductos.ListItems(i).SubItems(4)))
       End If
       
       If ListvProductos.ListItems(i).SubItems(6) = "SI" Then 'si se esta entregando en item
          Actualizar_Existencia_Producto ListvProductos.ListItems(i).Text, ListvProductos.ListItems(i).SubItems(2) * -1
       End If

        'Modulo.Actualizar_Deuda_Cliente CLng(s1), CLng(s2), -1# * V_FOTOGRAFIA_BS

    End If
  Next i
  Modulo.Actualizar_Deuda_Cliente CLng(iCliente), CLng(iSCliente), lNegociacionActual * (-1)
  Modulo.Actualizar_Deuda_Cliente CLng(iCliente), CLng(iSCliente), lDeuda
  On Error GoTo falla
  If iSCliente = "0" Then
      If cmbIDdisenoActivo.Text <> "" Then
         s = "Update Clientes Set IDDisenoActivo=" & cmbIDdisenoActivo.Text & " Where Codigo=" & CStr(iCliente)
         Modulo.ExecSQL s
      Else
         MsgBox "Debe seleccionar un diseño activo", vbExclamation
      End If
  Else
      If cmbIDdisenoActivo.Text <> "" Then
         s = "Update subClientes Set IDDisenoActivo=" & cmbIDdisenoActivo.Text & " Where id=" & CStr(iSCliente) & " And  cliente=" & CStr(iCliente)
         Modulo.ExecSQL s
      Else
         MsgBox "Debe seleccionar un diseño activo", vbExclamation
      End If
  End If
falla:
  If Err.Number <> 0 Then
      If cmbIDdisenoActivo.Text <> "" Then
         MsgBox Err.Number & "::" & Err.Description, vbCritical
      Else
         MsgBox "Debe seleccionar un diseño activo", vbExclamation
      End If
  End If
End Sub

Private Sub sbAceptar_Click()
  Dim s As String, s1 As String, i As Integer
  Dim sNomCliente As String, sNomSubCliente As String
  Dim r As New ADODB.Recordset
  Dim sNomArchivoInicial As String, sNomArchivoNuevo As String
  Dim dP As Double, p As Integer, e As Boolean
  Dim s2 As String
  
  Dim sCD As String, sMCD As String, dMCD As Double
  
  
  If Trim(cCP.Text) = "" Then Exit Sub
  
  s = Mid(cCP.Text, 1, 6)
  iCliente = CLng(s)
  
  If Trim(cSC.Text) = "" Then s = "0" Else s = Mid(cSC.Text, 1, 6)
  iSCliente = CLng(s)
  
  s = "select * from FormatoDiseno where " & _
      "cliente    = " & CStr(iCliente) & " and " & _
      "subcliente = " & CStr(iSCliente) & " "
  
  r.Open s, Modulo.DBConexionSQL, adOpenStatic
  If r.EOF Then 'NO EXISTE... Agregarlo:
    s = "insert into FormatoDiseno " & _
        "(cliente, subcliente, especificaciones) values (" & _
        CStr(iCliente) & "," & CStr(iSCliente) & ",'" & _
        Trim(eIndicaciones.Text) & "')"
    
  Else
    'YA EXISTE... Actualizarlo:
    s = "update FormatoDiseno set " & _
        "especificaciones = '" & Trim(eIndicaciones.Text) & "' " & _
        "where " & _
        "cliente    = " & CStr(iCliente) & " and " & _
        "subcliente = " & CStr(iSCliente) & " "
        
  End If
  
  Modulo.ExecSQL s
  
  '.Ahora procesar los productos con precios acordados:
  '1.borrar los ya registrados:
  
 ''sInsertarProductos
  
  
  s = "delete from PreciosEspeciales where " & _
      "cliente    = " & CStr(iCliente) & " and " & _
      "subcliente = " & CStr(iSCliente) & " "
      
  Modulo.ExecSQL s
      
  '2.agregar los actuales:
  For i = 0 To lProductos.ListCount - 1
    If lProductos.Selected(i) Then
      s1 = Trim(Mid(lProductos.List(i), 41))
      dP = CDbl(s1)
      s1 = Format(dP, "#0.00")
      p = InStr(s1, ",")
      If p > 0 Then Mid(s1, p, 1) = "."
                  
      s = "insert into PreciosEspeciales " & _
          "(cliente, subcliente, codigoproducto, precio, fecha) values (" & _
          CStr(iCliente) & "," & _
          CStr(iSCliente) & ",'" & _
          aProductos.List(i) & "'," & _
          s1 & ",'" & Format(Date, "yyyymmdd") & "')"
      
      Modulo.ExecSQL s
    End If
  Next i
  
  '3.Actualizar el Cliente / SubCliente:
  If iSCliente = 0 Then
    s = "update Clientes Set " & _
        "direccion   = '" & Trim(tdir.Text) & "'," & _
        "telefonos   = '" & Trim(ttel.Text) & "'," & _
        "fax         = '" & Trim(tfax.Text) & "'," & _
        "email       = '" & Trim(temail.Text) & "'," & _
        "rif         = '" & Trim(trif.Text) & "'," & _
        "contacto    = '" & Trim(tcon.Text) & "'," & _
        "contactotlf = '" & Trim(tcontelf) & "' " & _
        "where " & _
        "codigo      = " & CStr(iCliente) & " "
  Else
    s = "update SubClientes Set " & _
        "direccion   = '" & Trim(tdir.Text) & "'," & _
        "telefonos   = '" & Trim(ttel.Text) & "'," & _
        "fax         = '" & Trim(tfax.Text) & "'," & _
        "email       = '" & Trim(temail.Text) & "'," & _
        "rif         = '" & Trim(trif.Text) & "'," & _
        "contacto    = '" & Trim(tcon.Text) & "'," & _
        "contactotlf = '" & Trim(tcontelf) & "' " & _
        "where " & _
        "cliente     = " & CStr(iCliente) & " and " & _
        "id          = " & CStr(iSCliente) & " "
  End If
  
  Modulo.ExecSQL s
  
  AgregarLogs "Agrega/Modifica Formato de Diseno de [" & Mid(cCP.Text, 1, 20) & "...]"
  
  
  '--Actualizar Monto Deuda de Cliente/Subcliente si existe "DISENO"
  '--"D0001"
  
  'Al guardar: 1.restar la q tenia antes.
  '            2.asignar nueva, si la hay.
  '1-->
  s1 = Trim(Mid(cCP.Text, 1, 6))
  s2 = Trim(Mid(cSC.Text, 1, 6))
  If s2 = "" Or s2 = "-" Then s2 = "0"
  Modulo.Actualizar_Deuda_Cliente CLng(s1), CLng(s2), -1# * V_DISENO_BS
  Modulo.Actualizar_Deuda_Cliente CLng(s1), CLng(s2), -1# * V_FOTOGRAFIA_BS
  '<--
  
  '2-->
  sCD = GetSetting(APPNAME, "Opciones", "CodigoDiseno", "")
  If sCD <> "" Then
    i = 0
    e = False
    Do While i < lProductos.ListCount And Not e
      If aProductos.List(i) = sCD Then
        e = True
      Else
        i = i + 1
      End If
    Loop
    
    If e Then
      sMCD = Mid(lProductos.List(i), 41)
      dMCD = CDbl(sMCD)
      
      s1 = Trim(Mid(cCP.Text, 1, 6))
      s2 = Trim(Mid(cSC.Text, 1, 6))
      
      If s2 = "" Or s2 = "-" Then s2 = "0"
      
      Modulo.Actualizar_Deuda_Cliente CLng(s1), CLng(s2), dMCD
    End If
  End If
  '2.1-->Servicio Foto:
  sCD = GetSetting(APPNAME, "Opciones", "CodigoServicioFoto", "")
  
  If sCD <> "" Then
    i = 0
    e = False
    Do While i < lProductos.ListCount And Not e
      If aProductos.List(i) = sCD Then
        e = True
      Else
        i = i + 1
      End If
    Loop
    
    If e Then
      sMCD = Mid(lProductos.List(i), 41)
      dMCD = CDbl(sMCD)
      
      s1 = Trim(Mid(cCP.Text, 1, 6))
      s2 = Trim(Mid(cSC.Text, 1, 6))
      
      If s2 = "" Or s2 = "-" Then s2 = "0"
      
      Modulo.Actualizar_Deuda_Cliente CLng(s1), CLng(s2), dMCD
    End If
  End If
  '<---2

     
  
  MsgBox "Información ha sido almacenada correctamente...", vbInformation, "Información"
  
  r.Close
  Set r = Nothing
  
End Sub
Private Sub sInsertarProductos(argCodigoProducto As String, argPrecio As Currency, argCantidad As Integer, argEntregado As String, argTipoPago As Integer, argPagado As String, argIDDiseño As Integer)
   Dim SqlTxt As String
   Dim lCn  As New ADODB.Connection
   On Error GoTo falla
      SqlTxt = "insert into PreciosEspeciales " & _
          "(cliente, subcliente, codigoproducto, precio, fecha,Cantidad,Entregado,TipoPago,Pagado,IDDiseño) values (" & _
          CStr(iCliente) & "," & _
          CStr(iSCliente) & ",'" & _
          argCodigoProducto & "'," & _
          Replace(argPrecio, ",", ".") & ",'" & Format(Date, "yyyymmdd") & "'," & _
          argCantidad & ",'" & _
          argEntregado & "'," & _
          argTipoPago & ",'" & _
          argPagado & "'," & argIDDiseño & ")"
      
      lCn.ConnectionString = Modulo.DBConexionSQL
      lCn.Open
      lCn.Execute SqlTxt
falla:
   If Err.Number <> 0 Then
      'MsgBox Err.Number & "::" & Err.Description, vbCritical
      sActualizarProducto argCodigoProducto, argPrecio, argCantidad, argEntregado, argTipoPago, argPagado, argIDDiseño
   End If
End Sub

Private Sub sActualizarProducto(argCodigoProducto As String, argPrecio As Currency, argCantidad As Integer, argEntregado As String, argTipoPago As Integer, argPagado As String, argIDDiseño As Integer)
   Dim SqlTxt As String
   Dim lCn As New ADODB.Connection
   On Error GoTo falla
      SqlTxt = "Update PreciosEspeciales " & _
          "set Precio = " & argPrecio & "," & _
          "Cantidad = " & argCantidad & "," & _
          "Entregado='" & argEntregado & "'," & _
          "TipoPago=" & argTipoPago & "," & _
          "Pagado='" & argPagado & "'" & _
          "Where Cliente=" & CStr(iCliente) & _
          " And SubCliente= " & CStr(iSCliente) & _
          " And CodigoProducto='" & argCodigoProducto & "'" & _
          " And IDDiseño=" & argIDDiseño
          
      lCn.ConnectionString = Modulo.DBConexionSQL
      lCn.Open
      lCn.Execute SqlTxt
falla:
   If Err.Number <> 0 Then
      MsgBox Err.Number & "::" & Err.Description, vbCritical
      ''sActualizarProducto argCodigoProducto, argPrecio, argCantidad
   End If

End Sub


Private Sub bAceptar_Click()
   sGuardarDatos
   MsgBox "Listo", vbInformation
   ''cCP_Click
End Sub

Private Sub bAg_Click()
  Dim s As String, sRutaCliente As String, sRutaSubCliente As String
  Dim sNomCliente As String, sNomSubCliente As String
  Dim r As New ADODB.Recordset, sF As String
  Dim sNomArchivoInicial As String, sNomArchivoNuevo As String
  
  'sRutaCliente = ""
  's = "select RutaDestinoDatosCliente from Opciones"
  'r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  'If Not r.EOF Then
  '  If Not IsNull(r.Fields("RutaDestinoDatosCliente").Value) Then
  '    sRutaCliente = Trim(r.Fields("RutaDestinoDatosCliente").Value)
  '  End If
  'End If
  'r.Close
  'Set r = Nothing
  
  sRutaCliente = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
  
  If sRutaCliente = "" Then
    MsgBox "Falta Configurar Ruta de Datos del Cliente, Revise...", vbCritical, "Información"
    Exit Sub
  End If
  
  
  If Trim(cCP.Text) = "" Then Exit Sub
  
  s = Mid(cCP.Text, 1, 6)
  iCliente = CLng(s)
  
  If Trim(cSC.Text) = "" Then s = "0" Else s = Mid(cSC.Text, 1, 6)
  iSCliente = CLng(s)
  
  sNomCliente = ""
  sNomSubCliente = ""
  
  '1.Obtener Nombre del Cliente:
  s = "select nombre from clientes where codigo = " & CStr(iCliente) & " "
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If Not r.EOF Then
    If Not IsNull(r.Fields("nombre").value) Then
      sNomCliente = Trim(r.Fields("nombre").value)
    End If
  End If
  r.Close
  
  If sNomCliente = "" Then Exit Sub
  
  '2.Obtener Nombre del Sub-cliente:
  sNomSubCliente = ""
  If iSCliente > 0 Then
    s = "select nombre from subclientes where cliente = " & CStr(iCliente) & " and id = " & iSCliente & " "
    r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    If Not r.EOF Then
      If Not IsNull(r.Fields("nombre").value) Then
        sNomSubCliente = Trim(r.Fields("nombre").value)
      End If
    End If
    r.Close
  
    If sNomSubCliente = "" Then Exit Sub
  End If
     
  sRutaCliente = sRutaCliente & "\" & sNomCliente & IIf(sNomSubCliente <> "", "\" & sNomSubCliente, "") & "\IMAGENES"
  
  Set r = Nothing
      
  CommonDialog1.InitDir = sRutaCliente
  'CommonDialog1.FileName = "FORMATO_DISENO_" & Format(Date, "dd_mm_yyyy") & ".JPG"
  CommonDialog1.CancelError = False
  CommonDialog1.ShowOpen
  If Err.Number = 0 Then
    sNomArchivoInicial = CommonDialog1.FileName
    If sNomArchivoInicial <> "" Then
    
      Load fFechas
      fFechas.eFechaVence.value = Date + 365   'un año por default!
      Modulo.vTemporal1 = ""
      Modulo.fModalResult = Modulo.fModalResultCANCEL
      fFechas.Show vbModal
      If Modulo.fModalResult = Modulo.fModalResultOK Then
      
        'Modulo.vTemporal1 = Format(eFechaVence.Value, "dd/mm/yyyy")
        
        sF = Mid(Modulo.vTemporal1, 7, 4) & _
             Mid(Modulo.vTemporal1, 4, 2) & _
             Mid(Modulo.vTemporal1, 1, 2)
             
        sNomArchivoNuevo = sRutaCliente & "\" & "FORMATO_DISENO_" & Mid(CommonDialog1.FileTitle, 1, Len(CommonDialog1.FileTitle) - 4) & "_" & Format(Date, "dd_mm_yyyy") & ".JPG"
        
        'If sNomSubCliente = "" Then sNomSubCliente = sNomCliente
        'sNomArchivoNuevo = sRutaCliente & "\" & "FORMATO_DISENO_" & sNomSubCliente & "_" & Format(Date, "dd_mm_yyyy") & ".JPG"
        
        
        FileCopy sNomArchivoInicial, sNomArchivoNuevo
        If sF <> "" Then
           s = "insert into FormatoDisenoDetalle " & _
               "(cliente, subcliente, fecha, vencimiento, imagen) values (" & _
               CStr(iCliente) & "," & CStr(iSCliente) & ",'" & _
               Format(Date, "yyyymmdd") & "','" & _
               sF & "','" & _
               "FORMATO_DISENO_" & Mid(CommonDialog1.FileTitle, 1, Len(CommonDialog1.FileTitle) - 4) & "_" & Format(Date, "dd_mm_yyyy") & ".JPG')"
        Else
           s = "insert into FormatoDisenoDetalle " & _
               "(cliente, subcliente, fecha,imagen) values (" & _
               CStr(iCliente) & "," & CStr(iSCliente) & ",'" & _
               Format(Date, "yyyymmdd") & "','" & _
               "FORMATO_DISENO_" & Mid(CommonDialog1.FileTitle, 1, Len(CommonDialog1.FileTitle) - 4) & "_" & Format(Date, "dd_mm_yyyy") & ".JPG')"
           
        End If
            
        Modulo.ExecSQL s
      
        Cargar_Formatos
      End If
    End If
  End If
  
  
End Sub


Private Sub bBo_Click()
  Dim i As Integer
  Dim s As String, X As String
  Dim iCliente As Long, iSCliente As Long
  Dim sRuta As String, sArchivo As String
  
  s = Mid(cCP.Text, 1, 6)
  iCliente = CLng(s)
  
  If Trim(cSC.Text) = "" Then iSCliente = 0 Else iSCliente = CLng(Mid(cSC.Text, 1, 6))
   
  i = lFormatos.ListIndex
  If i >= 0 Then
    If MsgBox("¿Está Seguro de Borrar esta Hoja de Formato Diseño?", vbQuestion + vbYesNo) = vbYes Then
      ' dd/mm/yyyy
      ' 1234567890
      X = Trim(lFormatos.List(i))
      X = Mid(X, 7, 4) & Mid(X, 4, 2) & Mid(X, 1, 2)  'yyyymmdd
        
      'Borrar de la Tabla de datos:
      
      s = "delete from FormatoDisenoDetalle where " & _
          "cliente    = " & CStr(iCliente) & " and " & _
          "subcliente = " & CStr(iSCliente) & " and " & _
          "fecha      = '" & X & "'"
        
      Modulo.ExecSQL s
      
      'Borrar de su carpeta el archivo .JPG:
      sRuta = Modulo.LA_RUTA_DEL_CLIENTE(cCP.Text, cSC.Text)
      sArchivo = sRuta & "\" & aImagenes.List(i)
      If Dir(sArchivo) <> "" Then Kill sArchivo
            
      'Borrar del Item de formatos:
      lFormatos.RemoveItem i
      lVence.RemoveItem i
      aImagenes.RemoveItem i
      ListID.RemoveItem i
      Set Image1.Picture = Nothing
           
      
    End If
  End If
 
End Sub

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
    If Len(Modulo.vTemporal1) < 6 Then Modulo.vTemporal1 = Zeros(CLng(Modulo.vTemporal1), 6)
    cCP.ListIndex = Modulo.Buscar_ComboLen(cCP, Modulo.vTemporal1, 6)
  End If

End Sub

Private Sub bCancelar_Click()
  Unload Me
End Sub
Private Sub bVerSoloPrecios_Click()
If HayCambios = True Then
   If MsgBox("Se han hecho cambios a los datos de los productos, si continua perderá dichos cambios.¿Continuar?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
End If
CargarProductosAcordados
HayCambios = False
End Sub


Private Sub bVerSoloPrecios_Click_()
  Dim i As Integer
  Dim se_borro As Boolean
  
  Do While Modulo.No_Hay_Seleccion(lProductos)
    i = 0
    se_borro = False
    Do While i < lProductos.ListCount And Not se_borro
      If Not lProductos.Selected(i) Then
        lProductos.RemoveItem (i)
        aProductos.RemoveItem (i)
        se_borro = True
      End If
      i = i + 1
    Loop
  Loop
  
  If lProductos.ListCount > 0 Then lProductos.ListIndex = 0
End Sub

Private Sub bVerTodo_Click()
Dim lCn As New ADODB.Connection
Dim lCmd As New ADODB.Command
Dim lReg As New ADODB.Recordset
Dim lItem As ListItem
Dim lCliente As Integer
Dim s As String
Dim i As Integer
Dim j As Integer
Dim lTotal As Currency
If HayCambios = True Then
   If MsgBox("Se han hecho cambios a los datos de los productos, si continua perderá dichos cambios.¿Continuar?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
End If
HayCambios = False
lTotal = 0
  s = Mid(cCP.Text, 1, 6)
  lCliente = CInt(s)
  lCn.ConnectionString = Modulo.DBConexionSQL
  lCn.Open

  Set lReg = lCn.Execute("Select * from Productos Order By codigo")
  ListvProductos.ListItems.Clear
  If lReg.EOF = False Then
     Do While lReg.EOF = False
        Set lItem = ListvProductos.ListItems.Add(, , Trim(lReg!Codigo))
         lItem.SubItems(1) = UCase(Trim(lReg!Descripcion))
         lItem.SubItems(2) = 0
         lItem.SubItems(3) = Trim(lReg!precio)
         lItem.SubItems(4) = Trim(lReg!precio)
         lItem.SubItems(5) = 0
         lItem.SubItems(6) = "NO"
         lItem.SubItems(7) = "NO"
         lItem.SubItems(8) = "0"
        lReg.MoveNext
     Loop
  End If
  Set lReg = lCn.Execute("Select * from V_PreciosEspeciales where Cliente=" & lCliente & " and IDDiseño=" & Val(ListID.List(ListID.ListIndex)))
  If lReg.EOF = False Then
     lNegociacionActual = 0
     lCantidadProductosNegociados = 0
     Do While lReg.EOF = False
        j = 1
        For i = 1 To ListvProductos.ListItems.Count
           If ListvProductos.ListItems(i).Text = Trim(lReg!CodigoProducto) Then
              If lReg!ENTREGADO = "SI" Then
                 ProductosNegociados(j).CodigoProducto = Trim(lReg!CodigoProducto)
                 ProductosNegociados(j).Cantidad = IIf(IsNull(Trim(lReg!Cantidad)), 0, Trim(lReg!Cantidad))
                 lCantidadProductosNegociados = lCantidadProductosNegociados + 1
              End If
              ListvProductos.ListItems(i).Checked = True
              ListvProductos.ListItems(i).SubItems(2) = IIf(IsNull(Trim(lReg!Cantidad)), 0, Trim(lReg!Cantidad))
              ListvProductos.ListItems(i).SubItems(4) = lReg!PrecioAcordado
              ListvProductos.ListItems(i).SubItems(5) = lReg!PrecioAcordado * IIf(IsNull(Trim(lReg!Cantidad)), 0, Trim(lReg!Cantidad))
              ListvProductos.ListItems(i).SubItems(6) = lReg!ENTREGADO
              ListvProductos.ListItems(i).SubItems(7) = lReg!Pagado
              ListvProductos.ListItems(i).SubItems(8) = lReg!ID
              If lReg!Pagado = "SI" Then lNegociacionActual = lNegociacionActual + Val(lItem.SubItems(5))
              lTotal = lTotal + Val(ListvProductos.ListItems(i).SubItems(5))
              Exit For
           End If
        Next i
        j = j + 1
        lReg.MoveNext
     Loop
     
  End If
  txtTotal.Text = FormatNumber(lTotal, 2)
  lNegociacionActual = lTotal
   
End Sub

Private Sub bVerTodo_Click_()
  Dim s As String
  Dim r As New ADODB.Recordset
  
  s = "select * from Productos Order By Descripcion"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  
  aProductos.Clear
  lProductos.Clear
  
  Do While Not r.EOF
  
    s = r.Fields("codigo").value
    If Not Modulo.EXISTE_LIST(aProductos, s) Then
      aProductos.AddItem s
      
      s = Trim(r.Fields("descripcion").value)
      If Len(s) > 40 Then s = Mid(s, 1, 40)
      If Len(s) < 40 Then s = Modulo.BlancosDER(s, 40)
      s = s & " " & Modulo.BlancosIZQ(Format(r.Fields("precio").value, "#,0.00"), 10)
      
      lProductos.AddItem s
    End If
    r.MoveNext
  Loop
  r.Close
  Set r = Nothing
  
End Sub

Private Sub cCP_Change()
  Cargar_SubClientes
End Sub

Private Sub cCP_Click()
  Cargar_SubClientes
  'bVerTodo_Click
  Set Image1.Picture = Nothing
  HayCambios = False
End Sub
Public Sub sCargarDatosCliente()
  cCP_Click
End Sub
Private Sub Cargar_SubClientes()
  'Dim r As New ADODB.Recordset
  Dim s As String
  Dim l As Integer
  Dim sCod As String
  Dim RSubClientes As New ADODB.Recordset
  
  cSC.Clear
  
  Viendo = False
  
  If RSubClientes.State <> adStateClosed Then RSubClientes.Close
  
  sCod = "000000"
  If Trim(cCP.Text) <> "" Then sCod = Mid(cCP.Text, 1, 6)
      
  s = "SELECT * FROM SubClientes WHERE Cliente = " & sCod & " ORDER BY Id"
  
  RSubClientes.Open s, DBConexionSQL, adOpenDynamic, adLockOptimistic
  l = 1
  If RSubClientes.EOF = False Then
     cSC.BackColor = &H80FFFF
  Else
    cSC.BackColor = &HC0FFFF
  End If
  Do While Not RSubClientes.EOF
    s = Zeros(RSubClientes.Fields("id").value, 6) & " : " & Trim(RSubClientes.Fields("nombre").value)
    cSC.AddItem s
    RSubClientes.MoveNext
  Loop
  RSubClientes.Close
  Set RSubClientes = Nothing
  
  If cSC.ListCount > 0 Then cSC.ListIndex = 0
  cSC.ListIndex = -1


'  CargarProductosAcordados
  
  Cargar_Datos_Cliente
  Cargar_Formatos
  
  'ListID_Click
  
End Sub

Private Sub Cargar_Clientes()
  'Dim r As New ADODB.Recordset
  Dim s As String
  Dim l As Integer
  Dim sCod As String
  Dim RClientes As New ADODB.Recordset
  
  If RClientes.State <> adStateClosed Then RClientes.Close
  
  s = "SELECT * FROM Clientes ORDER BY Codigo"
  
  RClientes.Open s, DBConexionSQL, adOpenKeyset, adLockReadOnly
  
  cCP.Clear
  
  Do While Not RClientes.EOF
    s = Zeros(RClientes.Fields("codigo").value, 6) & " : " & Trim(RClientes.Fields("nombre").value)
    cCP.AddItem s
    RClientes.MoveNext
  Loop
  
  If cCP.ListCount > 0 Then cCP.ListIndex = 0
  
  RClientes.Close
  Set RClientes = Nothing
End Sub

Private Sub CargarProductosAcordados()
Dim lCn As New ADODB.Connection
Dim lCmd As New ADODB.Command
Dim lReg As New ADODB.Recordset
Dim lItem As ListItem
Dim lCliente As Integer
Dim s As String
Dim lTotal As Currency
Dim i As Integer
lTotal = 0
i = 1
  s = Mid(cCP.Text, 1, 6)
  lCliente = CInt(s)
  lCn.ConnectionString = Modulo.DBConexionSQL
  lCn.Open
  Set lReg = lCn.Execute("Select * from V_PreciosEspeciales where Cliente=" & lCliente & " And IDDiseño=" & Val(ListID.List(ListID.ListIndex)))
  ListvProductos.ListItems.Clear
  lNegociacionActual = 0
  lCantidadProductosNegociados = 0
  If lReg.EOF = False Then
     Do While lReg.EOF = False
        Set lItem = ListvProductos.ListItems.Add(, , Trim(lReg!CodigoProducto))
         If lReg!ENTREGADO = "SI" Then
            ProductosNegociados(i).CodigoProducto = Trim(lReg!CodigoProducto)
            ProductosNegociados(i).Cantidad = IIf(IsNull(Trim(lReg!Cantidad)), 0, Trim(lReg!Cantidad))
            lCantidadProductosNegociados = lCantidadProductosNegociados + 1
         End If
         lItem.SubItems(1) = UCase(Trim(lReg!Descripcion))
         lItem.SubItems(2) = IIf(IsNull(Trim(lReg!Cantidad)), 0, Trim(lReg!Cantidad))
         lItem.SubItems(3) = Trim(lReg!precioNormal)
         lItem.SubItems(4) = Trim(lReg!PrecioAcordado)
         lItem.SubItems(5) = lReg!PrecioAcordado * IIf(IsNull(Trim(lReg!Cantidad)), 0, Trim(lReg!Cantidad))
         lItem.SubItems(6) = lReg!ENTREGADO
         lItem.SubItems(7) = lReg!Pagado
         lItem.SubItems(8) = lReg!ID
         lItem.Checked = True
         If lReg!Pagado = "SI" Then lNegociacionActual = lNegociacionActual + Val(lItem.SubItems(5))
         lTotal = lTotal + Val(lItem.SubItems(5))
        i = i + 1
        lReg.MoveNext
     Loop
     
  End If
  txtTotal.Text = FormatNumber(lTotal, 2)
  
End Sub

Private Sub CargarProductosAcordados_()
  Dim s As String, s1 As String
  Dim r As New ADODB.Recordset
  Dim rP As New ADODB.Recordset
  Dim iCliente As Long
  Dim iSCliente As Long
  Dim sP As String 'Productos
  Dim sD As String 'Descripcion
  Dim s10 As String
  Dim lItem As ListItem
  If Trim(cCP.Text) = "" Then Exit Sub
  
  s = Mid(cCP.Text, 1, 6)
  iCliente = CLng(s)
  
  If Trim(cSC.Text) = "" Then s = "0" Else s = Mid(cSC.Text, 1, 6)
  iSCliente = CLng(s)
  
  eIndicaciones.Text = ""
  s1 = "select Especificaciones from FormatoDiseno where cliente = " & CStr(iCliente) & " and subcliente = " & CStr(iSCliente)
  r.Open s1, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If Not r.EOF Then eIndicaciones.Text = Trim(r.Fields("Especificaciones").value)
  r.Close
   
  s1 = "select * from PreciosEspeciales where cliente = " & CStr(iCliente) & " and subcliente = " & CStr(iSCliente)
  sP = "select * from Productos"
    
  r.Open s1, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  rP.Open sP, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    
  lProductos.Clear
  aProductos.Clear
  
  s10 = GetSetting(APPNAME, "Opciones", "CodigoDiseno", "")
  V_DISENO_BS = 0#
  
  s11 = GetSetting(APPNAME, "Opciones", "CodigoServicioFoto", "")
  V_FOTOGRAFIA_BS = 0#
  
  Do While Not r.EOF 'Productos acordados
    s = Trim(r.Fields("codigoproducto").value)
    sD = Space(50)
    
    If rP.RecordCount > 0 Then rP.MoveFirst
    rP.Find "Codigo = '" & s & "'"
    If Not rP.EOF Then
      sD = Trim(rP.Fields("descripcion").value)
      If Len(sD) > 40 Then sD = Mid(sD, 1, 40)
      If Len(sD) < 40 Then sD = Modulo.BlancosDER(sD, 40)
      sD = sD & Modulo.BlancosIZQ(Format(r.Fields("precio").value, "#,0.00"), 10)
    End If
    Set lItem = ListvProductos.ListItems.Add(, , Trim(rP!Codigo))
     lItem.SubItems(1) = Trim(rP!Descripcion)
     lItem.SubItems(3) = Trim(r!precio)
    aProductos.AddItem s
    lProductos.AddItem sD
    lProductos.Selected(lProductos.ListCount - 1) = True
    
    If s = s10 Then 'Tiene Diseno (arte)
      V_DISENO_BS = r.Fields("precio").value
    End If
    
    If s = s11 Then 'Servicio Foto
      V_FOTOGRAFIA_BS = r.Fields("precio").value
    End If
    
    
    r.MoveNext
  Loop
  
  If lProductos.ListCount > 0 Then lProductos.ListIndex = 0
  
  r.Close
  rP.Close
  Set r = Nothing
  Set rP = Nothing
  
End Sub

Private Sub Cargar_Formatos()
  Dim s As String, s1 As String
  Dim r As New ADODB.Recordset
  Dim iCliente As Long
  Dim iSCliente As Long
  
  If Trim(cCP.Text) = "" Then Exit Sub
  
  s = Mid(cCP.Text, 1, 6)
  iCliente = CLng(s)
  
  If Trim(cSC.Text) = "" Then s = "0" Else s = Mid(cSC.Text, 1, 6)
  iSCliente = CLng(s)
  
  s1 = "Select ID,Fecha, Vencimiento, Imagen from FormatoDisenoDetalle where cliente = " & CStr(iCliente) & " and subcliente = " & CStr(iSCliente)
  r.Open s1, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    
  lFormatos.Clear
  lVence.Clear
  aImagenes.Clear
  ListID.Clear
  cmbIDdisenoActivo.Clear
  Do While Not r.EOF 'Imagenes ya guardadas... escaneadas
    s = Format(r.Fields("fecha").value, "dd/mm/yyyy")
    lFormatos.AddItem s
    
    s = Format(r.Fields("vencimiento").value, "dd/mm/yyyy")
    lVence.AddItem s
    
    s = Trim(r.Fields("imagen").value)
    aImagenes.AddItem s
    
    ListID.AddItem Trim(r!ID)
    cmbIDdisenoActivo.AddItem r!ID
    r.MoveNext
  Loop
  
  r.Close
  Set r = Nothing
On Error GoTo falla:
  If iSCliente = "0" Then
     s1 = "Select * from Clientes where  codigo= " & CStr(iCliente)
     r.Open s1, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
     If IsNull(r!IdDisenoActivo) = False Then
        cmbIDdisenoActivo.Text = r!IdDisenoActivo
     End If
  Else   'subcliente
     s1 = "Select * from SubClientes where  cliente= " & CStr(iCliente) & " And ID=" & CStr(iSCliente)
     r.Open s1, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
     If IsNull(r!IdDisenoActivo) = False Then
        cmbIDdisenoActivo.Text = r!IdDisenoActivo
     End If

  End If
falla:
   If Err.Number <> 0 Then
      If Err.Number <> 383 Then
         MsgBox Err.Number & "::" & Err.Description, vbCritical
      Else
         MsgBox "Debe seleccionar un Id de diseño valido", vbExclamation
      End If
   End If
End Sub

Private Sub Cargar_Datos_Cliente()
  Dim s As String, s1 As String
  Dim r As New ADODB.Recordset
  Dim iCliente As Long
  Dim iSCliente As Long
  
  If Trim(cCP.Text) = "" Then Exit Sub
  
  s = Mid(cCP.Text, 1, 6)
  iCliente = CLng(s)
  
  If Trim(cSC.Text) = "" Then s = "0" Else s = Mid(cSC.Text, 1, 6)
  iSCliente = CLng(s)
  
  If iSCliente <= 0 Then
    s1 = "select * from Clientes where codigo = " & CStr(iCliente)
  Else
    s1 = "select * from Subclientes where cliente = " & CStr(iCliente) & " and id = " & CStr(iSCliente)
  End If
  
  r.Open s1, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  
  tdir.Text = ""
  ttel.Text = ""
  tfax.Text = ""
  temail.Text = ""
  trif.Text = ""
  tcon.Text = ""
  tcontelf.Text = ""
  
  If Not r.EOF Then
    If Not IsNull(r.Fields("direccion").value) Then tdir.Text = Trim(r.Fields("direccion").value)
    If Not IsNull(r.Fields("telefonos").value) Then ttel.Text = Trim(r.Fields("telefonos").value)
    If Not IsNull(r.Fields("fax").value) Then tfax.Text = Trim(r.Fields("fax").value)
    If Not IsNull(r.Fields("email").value) Then temail.Text = Trim(r.Fields("email").value)
    If Not IsNull(r.Fields("rif").value) Then trif.Text = Trim(r.Fields("rif").value)
    If Not IsNull(r.Fields("contacto").value) Then tcon.Text = Trim(r.Fields("contacto").value)
    If Not IsNull(r.Fields("contactotlf").value) Then tcontelf.Text = Trim(r.Fields("contactotlf").value)
  End If
  
  r.Close
  Set r = Nothing

End Sub


Private Sub Cargar_Datos_SubCliente()
  Dim s As String, s1 As String
  Dim r As New ADODB.Recordset
  Dim iCliente As Long
  Dim iSCliente As Long
  
  If Trim(cCP.Text) = "" Then Exit Sub
  
  s = Mid(cCP.Text, 1, 6)
  iCliente = CLng(s)
  
  If Trim(cSC.Text) = "" Then s = "0" Else s = Mid(cSC.Text, 1, 6)
  iSCliente = CLng(s)
  
  If iSCliente <= 0 Then
    s1 = "select * from Clientes where codigo = " & CStr(iCliente)
  Else
    s1 = "select * from Subclientes where cliente = " & CStr(iCliente) & " and id = " & CStr(iSCliente)
  End If
  
  r.Open s1, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  
  tdir.Text = ""
  ttel.Text = ""
  tfax.Text = ""
  temail.Text = ""
  trif.Text = ""
  tcon.Text = ""
  tcontelf.Text = ""
  
  If Not r.EOF Then
    If Not IsNull(r.Fields("direccion").value) Then tdir.Text = Trim(r.Fields("direccion").value)
    If Not IsNull(r.Fields("telefonos").value) Then ttel.Text = Trim(r.Fields("telefonos").value)
    If Not IsNull(r.Fields("fax").value) Then tfax.Text = Trim(r.Fields("fax").value)
    If Not IsNull(r.Fields("email").value) Then temail.Text = Trim(r.Fields("email").value)
    If Not IsNull(r.Fields("rif").value) Then trif.Text = Trim(r.Fields("rif").value)
    If Not IsNull(r.Fields("contacto").value) Then tcon.Text = Trim(r.Fields("contacto").value)
    If Not IsNull(r.Fields("contactotlf").value) Then tcontelf.Text = Trim(r.Fields("contactotlf").value)
  End If
  
  r.Close
  Set r = Nothing

End Sub




Private Sub cSC_Change()
  If Trim(cSC.Text) <> "" Then Cargar_Datos_Cliente
End Sub

Private Sub cSC_Click()
  If Trim(cSC.Text) <> "" Then
     Cargar_Datos_Cliente
     Cargar_Formatos
  End If
End Sub

Private Sub lFormatos_Click()
  Dim i As Integer
  Dim s As String, sRutaCliente As String, sRutaSubCliente As String
  Dim sNomCliente As String, sNomSubCliente As String
  'Dim r As New ADODB.Recordset
  Dim sNomArchivo As String
   On Error GoTo falla
  i = lFormatos.ListIndex
  If i >= 0 Then
    lVence.ListIndex = i
    ListID.ListIndex = i
    s = aImagenes.List(i)
    sRutaCliente = Modulo.LA_RUTA_DEL_CLIENTE(cCP.Text, cSC.Text)
    CargarProductosAcordados
    If sRutaCliente = "" Then Exit Sub
              
    sNomArchivo = sRutaCliente & "\" & s
    
    If Dir(sNomArchivo) = "" Then
      MsgBox "Archivo de Imagen [" & s & "] No Existe...", vbCritical, "Información"
    Else
      'Picture1.Picture = LoadPicture(sNomArchivo)
      Image1.Picture = LoadPicture(sNomArchivo)
    End If
  End If
falla:
    If Err.Number <> 0 Then
       MsgBox Err.Number & "::" & Err.Description, vbCritical
    End If
End Sub


Private Sub ListID_Click()
  Dim i As Integer
  i = ListID.ListIndex
  If i >= 0 Then
    lFormatos.ListIndex = i
    lVence.ListIndex = i
    lFormatos_Click
    CargarProductosAcordados
  End If

End Sub

Private Sub ListvProductos_DblClick()
   frmSelProducto.Show vbModal
   sTotalizar
End Sub

Private Sub ListvProductos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
   sTotalizar
End Sub

Private Sub ListvProductos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 10 Then
   frmSelProducto.Show vbModal
   sTotalizar

End If
End Sub

Private Sub lProductos_DblClick()
  Dim i As Integer, j As Integer, longitud As Long
  Dim s As String, X As String, s1 As String, s2 As String
  Dim dP As Double
  i = lProductos.ListIndex
  
  If i >= 0 Then
    s = lProductos.List(i)
    s = Mid(s, 40 + 1)
    s = Trim(s)
    s = Mid(s, Len(s) - 4, 2)
    longitud = Len(lProductos.List(i))
    Do
      X = InputBox("Indique Monto de Precio Acordado:", "Confirme", s)
      If Not IsNumeric(X) Then
        MsgBox "ERROR: El Monto del Precio es Inválido...", vbCritical, "Información"
      End If
    Loop Until IsNumeric(X)
    dP = Val(X)
    
    s1 = Mid(lProductos.List(lProductos.ListIndex), 1, 40)
    s2 = Format(X, "#,0.00")
    j = Len(s2)
    
    s = s1 & Space(longitud - 40 - j) & s2
    lProductos.List(lProductos.ListIndex) = s
    X = InputBox("Indique la Cantidad:", "Confirme", Mid(lProductos.List(lProductos.ListIndex), Len(lProductos.List(lProductos.ListIndex)) - 2, 2))
    lProductos.List(lProductos.ListIndex) = s & " " & X
  End If
End Sub

Private Sub sTotalizar()
   Dim i As Integer
   Dim lTotal As Currency
   lTotal = 0
   For i = 1 To ListvProductos.ListItems.Count
      If ListvProductos.ListItems(i).Checked = True Then
         lTotal = lTotal + Val(Replace(ListvProductos.ListItems(i).SubItems(5), ",", "."))
      End If
   Next i
   txtTotal.Text = FormatNumber(lTotal, 2)
   
End Sub
Private Sub Form_Load()
  eIndicaciones.Text = ""
  lFormatos.Clear
  lProductos.Clear
  aImagenes.Clear
  Cargar_Clientes
  HayCambios = False
End Sub

Private Sub lVence_Click()
  Dim i As Integer
  i = lVence.ListIndex
  If i >= 0 Then
    lFormatos.ListIndex = i
    ListID.ListIndex = i
    CargarProductosAcordados
  End If
End Sub

Private Sub lVence_DblClick()
   Dim lVencimiento As String
   Dim SqlTxt As String
   Load fFechas
   If lVence.List(lVence.ListIndex) <> "" Then
      fFechas.eFechaVence.value = lVence.List(lVence.ListIndex)
   Else
      fFechas.eFechaVence.value = Date + 365
   End If
   fFechas.Show vbModal
   If Modulo.fModalResult = Modulo.fModalResultOK And lVence.List(lVence.ListIndex) <> Modulo.vTemporal1 Then
      'actualizar
      lVencimiento = Mid(Modulo.vTemporal1, 7, 4) & _
           Mid(Modulo.vTemporal1, 4, 2) & _
           Mid(Modulo.vTemporal1, 1, 2)
      
      
      SqlTxt = "Update FormatoDisenoDetalle " _
              & "Set Vencimiento='" & lVencimiento & "' " _
              & " Where cliente= " & CStr(Mid(cCP.Text, 1, 6)) _
              & " And SubCliente= " & CStr(IIf(Trim(cSC.Text) = "", "0", Mid(cSC.Text, 1, 6)) & " and ID=" & ListID.List(ListID.ListIndex))
            
        Modulo.ExecSQL SqlTxt
      
        Cargar_Formatos


   End If
   
End Sub
