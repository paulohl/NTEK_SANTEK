VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form fObservaciones 
   Caption         =   "Observaciones y Registro de Pago"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808000&
      Height          =   945
      Left            =   120
      ScaleHeight     =   885
      ScaleWidth      =   2895
      TabIndex        =   29
      Top             =   8730
      Width           =   2955
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0FF&
         X1              =   30
         X2              =   2790
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label lTC 
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
         Left            =   1440
         TabIndex        =   35
         Top             =   90
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deuda Bs:"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   34
         Top             =   90
         Width           =   1155
      End
      Begin VB.Label lTA 
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
         Left            =   1440
         TabIndex        =   33
         Top             =   300
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pagos Bs:"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   32
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label lTS 
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
         Left            =   1440
         TabIndex        =   31
         Top             =   630
         Width           =   930
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Saldo Bs:"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   30
         Top             =   630
         Width           =   1080
      End
   End
   Begin VB.CommandButton bDesmarcar 
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
      Left            =   4140
      Picture         =   "fObservaciones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6960
      Width           =   450
   End
   Begin VB.CommandButton bMarcar 
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
      Left            =   3690
      Picture         =   "fObservaciones.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6960
      Width           =   450
   End
   Begin VB.ListBox lC2 
      Height          =   1620
      Left            =   6030
      TabIndex        =   11
      Top             =   4350
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.ListBox lC1 
      Height          =   1620
      Left            =   4680
      TabIndex        =   10
      Top             =   4350
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   6825
      Left            =   3660
      TabIndex        =   26
      Top             =   90
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   12039
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      FocusRect       =   2
      AllowUserResizing=   3
      FormatString    =   $"fObservaciones.frx":1404
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   1740
      Top             =   6480
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
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar Combos"
      Height          =   855
      Left            =   3690
      TabIndex        =   12
      Top             =   7470
      Width           =   1725
      Begin VB.CheckBox C2 
         Caption         =   "Combo-2"
         Height          =   225
         Left            =   270
         TabIndex        =   14
         Top             =   510
         Width           =   1125
      End
      Begin VB.CheckBox C1 
         Caption         =   "Combo-1"
         Height          =   225
         Left            =   270
         TabIndex        =   13
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del Pago"
      Height          =   1905
      Left            =   5820
      TabIndex        =   8
      Top             =   7020
      Width           =   9045
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0E0FF&
         Height          =   1545
         Left            =   240
         ScaleHeight     =   1485
         ScaleWidth      =   5025
         TabIndex        =   9
         Top             =   240
         Width           =   5085
         Begin VB.TextBox ePago 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2700
            TabIndex        =   18
            Text            =   "0,00"
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Total Operación Bs:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   420
            TabIndex        =   23
            Top             =   120
            Width           =   2235
         End
         Begin VB.Label lTO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2700
            TabIndex        =   22
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Su Pago Bs:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1350
            TabIndex        =   21
            Top             =   510
            Width           =   1290
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   4800
            Y1              =   870
            Y2              =   870
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Su Cambio Bs:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1110
            TabIndex        =   20
            Top             =   1020
            Width           =   1530
         End
         Begin VB.Label lCambio 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2700
            TabIndex        =   19
            Top             =   990
            Width           =   1335
         End
      End
      Begin VB.Label lTotal 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.999,99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6570
         TabIndex        =   25
         Top             =   810
         Width           =   1995
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Bs."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6570
         TabIndex        =   24
         Top             =   420
         Width           =   1995
      End
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
      Left            =   1170
      TabIndex        =   6
      ToolTipText     =   "Borrar Item Observación"
      Top             =   7050
      Width           =   450
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   2250
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton bEd 
      Caption         =   ":-:"
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
      Left            =   630
      TabIndex        =   2
      ToolTipText     =   "Editar Observación"
      Top             =   7050
      Width           =   450
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
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "Agregar Nueva Observación"
      Top             =   7050
      Width           =   450
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6540
      ItemData        =   "fObservaciones.frx":148E
      Left            =   90
      List            =   "fObservaciones.frx":1490
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton bAceptar 
      Caption         =   "Guardar"
      Height          =   500
      Left            =   6600
      Picture         =   "fObservaciones.frx":1492
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9090
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   500
      Left            =   7770
      Picture         =   "fObservaciones.frx":1A1C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9090
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Resumen de Cuenta:"
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   8490
      Width           =   1500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   7620
      Width           =   525
   End
   Begin VB.Label lSubCliente 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   8160
      Width           =   2955
   End
   Begin VB.Label lCliente 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   7830
      Width           =   2955
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   7
      Top             =   90
      Width           =   3500
   End
End
Attribute VB_Name = "fObservaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Const cFG = "*     | Código                | Descripción                                                          | Precio Bs | Cantidad | SubTotal Bs"
Const cFG = "*    | Código           | Descripción                                                        |Precio Bs|Cantidad|SubTotal Bs|Entregado"

Const CLR_HIGH_FND = vbGreen
Const CLR_HIGH_TXT = vbBlack

Const CLR_NORMAL_FND = vbWhite
Const CLR_NORMAL_TXT = vbBlack

Dim cellFontNameOld As String
Dim cellAlignOld As Integer
Dim cellFontSizeOld As Integer

Const CMARCA = "P"
Const CDESMARCA = ""

Public bPedirCantidad As Boolean
Dim lObs(100) As Observaciones
Dim FaltaPago As Boolean
Dim EsPago As Boolean
Dim EsPagoPorAdelantado As Boolean
Public lRegDetalles As New ADODB.Recordset
Public EsActualizar As Boolean
Public sLocalizador  As String
Public EsGuardarPorLotes As Boolean


Private Sub sGuardarPorLotes(argTabla As String, argTipoPago As Integer)
  Dim lReg As New ADODB.Recordset
  Dim lReg2 As New ADODB.Recordset
  Dim Localizador As String
  Dim CodigoPVC As String
  Dim lLocalizadorOK As Boolean
  Dim sEntregado As String
  Dim s As String
  Dim i As Integer
  Dim DeudaMov As Double, dST As Double
  Dim PagoMov As Double
  Dim lCn As New ADODB.Connection
  Dim pr As Double
  Dim SqlTxt As String
  Dim lPrecio As String
  Dim lSubTotal As String
  lCn.ConnectionString = Modulo.DBConexionSQL
  lCn.Open
  CodigoPVC = GetSetting(APPNAME, "Opciones", "CodigoCarnet", "")
  If CodigoPVC = "" Then
     lReg.Open "select CodigoProductoPVC from opciones", Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
     If Not lReg.EOF Then
        If Not IsNull(lReg.Fields("CodigoProductoPVC").value) Then
          CodigoPVC = Trim(lReg.Fields("CodigoProductoPVC").value)
        End If
     End If
     lReg.Close
     Set lReg = Nothing
  End If
  Set lReg = New ADODB.Recordset
  lReg.Open "Select * From [" & argTabla & "] Where Marca='I'", Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If lReg.EOF = False Then
     Do While lReg.EOF = False
     Load fMensaje
     fMensaje.Label1.Caption = "Procesando Carnet con Cedula=" & Trim(lReg!Cedula) & ", Espere..."
     fMensaje.Show
     DoEvents
        'Generar Localizador**************************************************
        lLocalizadorOK = False
        Do While Not lLocalizadorOK
          Localizador = Modulo.GenerarLocalizador()
          Do While Modulo.DBExiste("Diario", "Localizador", Localizador)
            Localizador = Modulo.GenerarLocalizador()
          Loop
          If Not Modulo.DBExiste("DiarioDetalle", "Localizador", Localizador) Then
             lLocalizadorOK = True
          End If
       Loop
     '**********************************************************************
     DeudaMov = 0#
     Modulo.vMOVPAGO = 0#
     For i = 1 To FG.Rows - 1
        
     
          If FG.TextMatrix(i, 0) = CMARCA Then  'Producto Seleccionado
             dST = CDbl(FG.TextMatrix(i, 3)) * CDbl(FG.TextMatrix(i, 4))
             pr = CDbl(FG.TextMatrix(i, 3))
             FG.TextMatrix(i, 3) = pr
             pr = CDbl(FG.TextMatrix(i, 5))
             FG.TextMatrix(i, 5) = pr
             sEntregado = Mid(FG.TextMatrix(i, 6), 1, 1)
          
           'Insertar en DiarioDetalle
             lPrecio = FG.TextMatrix(i, 3)
             lPrecio = Trim(Replace(lPrecio, ",", "."))
             lSubTotal = FG.TextMatrix(i, 5)
             lSubTotal = Trim(Replace(lSubTotal, ",", ".")) 'SE DEBE FORZAR A QUE SEA PAGADO POR ADELANTADO
             
             s = "insert into DiarioDetalle " & _
                 "(Localizador,CodigoProducto,Cantidad,Precio,SubTotal,Entregado,Estacion) values " & _
                 "('" & _
                 Localizador & "','" & FG.TextMatrix(i, 1) & "'," & FG.TextMatrix(i, 4) & "," & _
                 lPrecio & "," & _
                 lSubTotal & ",'" & _
                 sEntregado & "','" & _
                 Modulo.ESTACION & "')"
             Modulo.ExecSQL s
             'FG.TextMatrix(i, 3) = pr
             If Trim(FG.TextMatrix(i, 1)) <> Trim(CodigoPVC) Then
                 ' si el producto no ha sido entregado no generar deuda ni actualizar existencia
                 'Aumentar la Deuda del Cliente siempre y cuando no sea
                 'CodigoProductoPVC:
                If sEntregado <> "N" Then
                   DeudaMov = DeudaMov + dST
                   Actualizar_Existencia_Producto FG.TextMatrix(i, 1), CLng(FG.TextMatrix(i, 4)) * -1
                End If
             End If
             If EsPago = True Then
                Modulo.vMOVPAGO = Modulo.vMOVPAGO + dST
             End If
             
          End If
     Next i
     '**********************************************************************
    s = lSubCliente.Caption
    If DeudaMov <> 0# Then
      If Trim(lSubCliente.Caption) = "" Or Trim(lSubCliente.Caption) = "-" Then
        s = "0"
      Else
        s = Mid(lSubCliente.Caption, 1, 6)
      End If
      'Modulo.Actualizar_Deuda_Cliente CLng(Mid(lCliente.Caption, 1, 6)), CLng(s), DeudaMov
      
    End If
    
    PagoMov = CDbl(lTotal.Caption)
    
    If PagoMov > 0# Then
      s = lSubCliente.Caption
      If Trim(lSubCliente.Caption) = "" Or Trim(lSubCliente.Caption) = "-" Then
        s = "0"
      Else
        s = Mid(lSubCliente.Caption, 1, 6)
      End If
      If EsPago = True Then Modulo.Actualizar_Pago_Cliente CLng(Mid(lCliente.Caption, 1, 6)), CLng(s), PagoMov
    End If
'***************************************************************************************
     'Insertar en Diario
     Load fMensaje
     fMensaje.Label1.Caption = "Insertando Carnet " & Trim(lReg!Cedula) & " en Diario y DiarioDetalle, Espere..."
     fMensaje.Show
     DoEvents
    'SE DEBE FORZAR A QUE SEA PAGADO POR ADELANTADO
    Set lReg2 = lCn.Execute("Select id from [" & argTabla & "] where cedula='" & lReg!Cedula & "'")
    SqlTxt = "insert into diario (localizador,fecha,hora,cliente,subcliente,tabla,cedula,observaciones,pago,monto,idCarnet,TipoPago) values ('" & _
           Localizador & "','" & _
           Format(Now, "yyyy/MM/dd") & "','" & _
           Format(Now, "hh:mm am/pm") & "'," & _
           CLng(Mid(lCliente.Caption, 1, 6)) & "," & _
           s & ",'" & _
           argTabla & "','" & _
           Trim(lReg!Cedula) & "','" & _
           "PORLOTES" & "','" & IIf(Modulo.vMOVPAGO > 0#, "S", "N") & "'," & Replace(CStr(PagoMov), ",", ".") & "," & (lReg2!ID + 1) & "," & argTipoPago & ")"
         ''Modulo.vTemporal1
             '(GRID.TextMatrix(GRID.Rows, 1) + 1) = para el id del carnet y poder ubicar el registro
    ''Clipboard.Clear
    ''Clipboard.SetText SqlTxt
    'On Error Resume Next
    Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
    Modulo.DBComandoSQL.CommandText = SqlTxt
    Modulo.DBComandoSQL.Execute
'***********************************************************************************
    
        
        lReg.MoveNext
     Loop
     Unload fMensaje
     'MsgBox "Proceso Carnets por Lotes Finalizado", vbInformation
  Else
     MsgBox "No existe ningun registro en la tabla del cliente"
  End If
 lReg.Close
End Sub


Private Sub bAceptar_Click()
  Dim i As Integer
  Dim s As String
  Dim r As New ADODB.Recordset
  Dim cp As String
  Dim DeudaMov As Double, dST As Double
  Dim PagoMov As Double
  'Dim sLocalizador As String
  Dim LocalizadorOK As Boolean
  Dim pr As Double
  Dim sEntregado As String
  fMov.lTipoPago = 0
  If Not Modulo.Hay_Seleccion(List1) Then
    MsgBox "Debe Seleccionar al menos una Observación...", vbCritical, "Información"
    Exit Sub
  Else
    Modulo.vTemporal1 = ""
    For i = 0 To List1.ListCount - 1
      If List1.Selected(i) Then
        If UCase(List1.List(i)) = "PAGO" Then
           EsPago = True
           fMov.lTipoPago = 1
        End If
        If UCase(List1.List(i)) = "PAGADO POR ADELANTADO" Then
           EsPagoPorAdelantado = True
           fMov.lTipoPago = 2
        End If
        Modulo.vTemporal1 = Modulo.vTemporal1 & " " & List1.List(i)
      End If
    Next i
  End If
  'en modulo.vTemporal1 es donde se estan guardando las observaciones que se seleccionan
  If EsPagoPorAdelantado = True And lTS >= 0 Then
     MsgBox "Usted Seleccionó Pago Por adelantado y el Cliente No tiene suficiente saldo para efectuar la transacción", vbCritical
     Exit Sub
  End If
  Modulo.vTemporal1 = Trim(Modulo.vTemporal1)
  
  Modulo.vMOVPAGO = 0#
    
  'If Not Modulo.HAY_SELECCION_FG(FG, 0, CMARCA) Then
  '  MsgBox "Debe Seleccionar al menos un Producto...", vbCritical, "Información"
  'Else
  
  If EsGuardarPorLotes = True Then  'FORZAR A QUE TIPO PAGO DEBA SER 2(PAGADO PORA DELANTADO)
     sGuardarPorLotes fPersonasAct.lTablas.List(fPersonasAct.lTablas.ListIndex), fMov.lTipoPago
     sLlamarCard5
     Unload fPersonasAct
     Unload Me
  End If
  
    cp = GetSetting(APPNAME, "Opciones", "CodigoCarnet", "")
   
    r.Open "select CodigoProductoPVC from opciones", Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    If Not r.EOF Then
      If Not IsNull(r.Fields("CodigoProductoPVC").value) Then
        cp = Trim(r.Fields("CodigoProductoPVC").value)
      End If
    End If
    r.Close
    Set r = Nothing
  
    'Load fMensaje
    'fMensaje.Label1.Caption = "Guardando Detalle de Movimiento, Espere..."
    'fMensaje.Show
    'DoEvents
    
    If EsActualizar = False Then
       LocalizadorOK = False
        
       Do While Not LocalizadorOK
        
         sLocalizador = Modulo.GenerarLocalizador()
        
         Do While Modulo.DBExiste("Diario", "Localizador", sLocalizador)
           
           sLocalizador = Modulo.GenerarLocalizador()
           
         Loop
          
         If Not Modulo.DBExiste("DiarioDetalle", "Localizador", sLocalizador) Then
           LocalizadorOK = True
         End If
          
       Loop
     Else
       LocalizadorOK = True
       
     End If
    Modulo.vLocalizador = sLocalizador
    
    DeudaMov = 0#
    
    For i = 1 To FG.Rows - 1
      
          dST = CDbl(Val(FG.TextMatrix(i, 3))) * CDbl(FG.TextMatrix(i, 4))
        
          pr = CDbl(Val(FG.TextMatrix(i, 3)))
          FG.TextMatrix(i, 3) = Str(pr)
          
          pr = CDbl(Val(FG.TextMatrix(i, 5)))
          FG.TextMatrix(i, 5) = Str(pr)
        
          sEntregado = Mid(FG.TextMatrix(i, 6), 1, 1)
        

      If EsActualizar = False Then
        If FG.TextMatrix(i, 0) = CMARCA Then  'Producto Seleccionado
        
          s = "insert into DiarioDetalle " & _
               "(Localizador,CodigoProducto,Cantidad,Precio,SubTotal,Entregado,Estacion) values " & _
               "('" & _
               sLocalizador & "','" & FG.TextMatrix(i, 1) & "'," & FG.TextMatrix(i, 4) & "," & _
               FG.TextMatrix(i, 3) & "," & _
               FG.TextMatrix(i, 5) & ",'" & _
               sEntregado & "','" & _
               Modulo.ESTACION & "')"
          Modulo.ExecSQL s
               
          If Trim(FG.TextMatrix(i, 1)) <> Trim(cp) Then
            ' si el producto no ha sido entregado no generar deuda ni actualizar existencia
          
        
             'Aumentar la Deuda del Cliente siempre y cuando no sea
             'CodigoProductoPVC:
             If sEntregado <> "N" Then
                DeudaMov = DeudaMov + dST
                Actualizar_Existencia_Producto FG.TextMatrix(i, 1), CLng(FG.TextMatrix(i, 4)) * -1
             End If
          End If
          If EsPago = True Then
             Modulo.vMOVPAGO = Modulo.vMOVPAGO + dST
          End If
        End If
      Else
         If FG.TextMatrix(i, 0) = CMARCA Then  'Producto Seleccionado
           r.Open "Select * from DiarioDetalle Where Localizador='" & sLocalizador & "' And CodigoProducto='" & FG.TextMatrix(i, 1) & "'", Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
           If Not r.EOF Then
              s = "Update DiarioDetalle Set Cantidad=" & FG.TextMatrix(i, 4) & ",Precio=" & FG.TextMatrix(i, 4) & ",SubTotal=" & FG.TextMatrix(i, 5) & ",Entregado='" & sEntregado & "' Where " _
                 & " Localizador='" & sLocalizador & "' And CodigoProducto='" & FG.TextMatrix(i, 1) & "'"
                     
           Else
              s = "insert into DiarioDetalle " & _
                 "(Localizador,CodigoProducto,Cantidad,Precio,SubTotal,Entregado,Estacion) values " & _
                 "('" & _
                 sLocalizador & "','" & FG.TextMatrix(i, 1) & "'," & FG.TextMatrix(i, 4) & "," & _
                 FG.TextMatrix(i, 3) & "," & _
                 FG.TextMatrix(i, 5) & ",'" & _
                 sEntregado & "','" & _
                 Modulo.ESTACION & "')"
              
           End If
           r.Close
           Modulo.ExecSQL s
          If Trim(FG.TextMatrix(i, 1)) <> Trim(cp) Then
            ' si el producto no ha sido entregado no generar deuda ni actualizar existencia
             'Aumentar la Deuda del Cliente siempre y cuando no sea
             'CodigoProductoPVC:
             If sEntregado <> "N" Then
                DeudaMov = DeudaMov + dST
                Actualizar_Existencia_Producto FG.TextMatrix(i, 1), CLng(FG.TextMatrix(i, 4)) * -1
             End If
          End If
          If EsPago = True Then
             Modulo.vMOVPAGO = Modulo.vMOVPAGO + dST
          End If
           
         Else ' si no esta marcado
           r.Open "Select * from DiarioDetalle Where Localizador='" & sLocalizador & "' And CodigoProducto='" & FG.TextMatrix(i, 1) & "'", Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
           If Not r.EOF Then
              s = "Delete from DiarioDetalle Where " _
                 & " Localizador='" & sLocalizador & "' And CodigoProducto='" & FG.TextMatrix(i, 1) & "'"
              Modulo.ExecSQL s
           End If
           r.Close
         End If
      End If
      
    Next i
    'Actualizar las observaciones en tabla Diario
   If EsActualizar = True Then
       s = "Update Diario Set Observaciones=Observaciones + ', " & Modulo.vTemporal1 & "(" & Format(Now, "dd/MM/yyyy") & ":Bs" & Replace(CDbl(lTotal.Caption), ",", ".") & ")', TipoPago=" & fMov.lTipoPago & ", Pago='" & IIf(fMov.lTipoPago > 0#, "S", "N") & "', Monto=" & IIf(fMov.lTipoPago > 0#, Replace(CDbl(lTotal.Caption), ",", "."), 0) & " Where Localizador='" & sLocalizador & "'"
       Modulo.ExecSQL s
   End If
    
    s = lSubCliente.Caption
    If DeudaMov <> 0# Then
      If Trim(lSubCliente.Caption) = "" Or Trim(lSubCliente.Caption) = "-" Then
        s = "0"
      Else
        s = Mid(lSubCliente.Caption, 1, 6)
      End If
      'Modulo.Actualizar_Deuda_Cliente CLng(Mid(lCliente.Caption, 1, 6)), CLng(s), DeudaMov
      
    End If
    
    PagoMov = CDbl(lTotal.Caption)
    
    If PagoMov > 0# Then
      s = lSubCliente.Caption
      If Trim(lSubCliente.Caption) = "" Or Trim(lSubCliente.Caption) = "-" Then
        s = "0"
      Else
        s = Mid(lSubCliente.Caption, 1, 6)
      End If
      If EsPago = True Then Modulo.Actualizar_Pago_Cliente CLng(Mid(lCliente.Caption, 1, 6)), CLng(s), PagoMov
    End If
    
    Auditar_Fotos
    Modulo.fModalResult = Modulo.fModalResultOK
    Unload Me
  'End If
End Sub

Private Sub sLlamarCard5()
  Dim sOri As String
  Dim sDes As String
  Dim SC As String
  Dim sSC As String
  Dim sRuta As String
  Dim s As String
  Dim sCard5 As String
  sOri = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
  sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
  
  If sOri = "" Then
     MsgBox "No es posible determinar la carpeta de destino. Diríjase al menú de Opciones y verifique"
  End If
  SC = Trim(lCliente.Caption)
  sSC = Trim(lSubCliente.Caption)
      
  SC = Trim(Mid(SC, 9))
       
  If Trim(sSC) = "" Or Trim(sSC) = "-" Then
    sRuta = sDes & "\" & SC & "\CARNET" & "\BASE CARNET " & SC & ".car"
  Else
    sSC = Trim(Mid(sSC, 7))
    sRuta = sDes & "\" & SC & "\" & sSC & "\CARNET" & "\BASE CARNET " & sSC & ".car"
  End If

  s = GetSetting(APPNAME, "Opciones", "RutaCard5", "")

  sCard5 = s & " " & Chr(34) & sRuta & Chr(34)

  If Shell(sCard5, vbMaximizedFocus) = 0# Then
    MsgBox "Error: No se pudo Iniciar Card-5" & vbCrLf & CStr(Err.Number) & ":" & Err.Description, "Información"
  End If
  
  Unload fMensaje

End Sub

Private Sub bAg_Click()
  Dim s As String, sql As String
  s = ""
  s = InputBox("Indique la Observación:", "Agregar Nueva Observación")
  If Trim(s) <> "" Then
    sql = "insert into observaciones (observacion) values ('" & s & "')"
    Modulo.ExecSQL sql
    CargarObservaciones
  End If
End Sub

Private Sub bBo_Click()
  Dim s As String, lID As Long, sSQL As String
  If List1.ListIndex < 0 Then
    MsgBox "Debe Marcar un item de Observación a Editar...", vbCritical, "Información"
  Else
    s = List1.List(List1.ListIndex)
    lID = CLng(List2.List(List1.ListIndex))
    If MsgBox("¿Está Seguro de Borrar la Observación:" & vbCrLf & _
              "[" & s & "]", vbQuestion + vbYesNo, "Confirme") = vbYes Then
      sSQL = "delete from observaciones where id = " & CStr(lID) & " "
      Modulo.ExecSQL sSQL
      CargarObservaciones
    End If
  End If
End Sub

Private Sub bCancelar_Click()
  Modulo.fModalResult = Modulo.fModalResultCANCEL
  Unload Me
End Sub

Private Sub DataGrid1_DblClick()
  If Not Adodc1.Recordset.EOF Then
    Modulo.vTemporal1 = Zeros(Adodc1.Recordset.Fields("observacion").value, 6)
    Unload Me
  End If
End Sub

Private Sub CargarObservaciones()
  Adodc1.ConnectionString = Modulo.DBConexionSQL.ConnectionString
  Adodc1.RecordSource = "select * from observaciones order by observacion"
  Adodc1.Refresh
  List1.Clear
  List2.Clear
  Do While Not Adodc1.Recordset.EOF
    List1.AddItem Trim(Adodc1.Recordset.Fields("observacion").value)
    List2.AddItem CStr(Adodc1.Recordset.Fields("id").value)
    lObs(List1.ListCount).Codigo = Adodc1.Recordset!ID
    lObs(List1.ListCount).Observacion = Trim(Adodc1.Recordset!Observacion)
    Adodc1.Recordset.MoveNext
  Loop
End Sub

Private Sub bDesmarcar_Click()
  Dim i As Integer
  Dim dST As Double
  
  For i = 1 To FG.Rows - 1
  
    FG.Row = i
    FG.Col = 0
  
    FG.CellFontName = cellFontNameOld
    FG.CellAlignment = cellAlignOld
    FG.CellFontSize = cellFontSizeOld
    
    FG.TextMatrix(i, 0) = ""
    
    FG.TextMatrix(i, 4) = "0"

    If Trim(FG.TextMatrix(i, 4)) <> "" Then
      If CLng(Trim(FG.TextMatrix(i, 4))) = 0 Then
        FG.TextMatrix(i, 4) = "0"
      End If
    End If
 
    FG.CellFontName = cellFontNameOld
    FG.CellAlignment = cellAlignOld
    FG.CellFontSize = cellFontSizeOld

    FG.Row = i: FG.Col = 6: FG.CellAlignment = flexAlignCenterCenter
    FG.TextMatrix(i, 6) = "NO"
    
    Calcular_SubTotal_En_Linea i
    
    Color_NORMAL_FILA i

  Next i
  
  TotalizarTODO
End Sub

Private Sub bEd_Click()
  Dim s As String, lID As Long, sSQL As String
  If List1.ListIndex < 0 Then
    MsgBox "Debe Marcar un item de Observación a Editar...", vbCritical, "Información"
  Else
    s = List1.List(List1.ListIndex)
    lID = CLng(List2.List(List1.ListIndex))
    s = InputBox("Observación:", "Editar Observación", s)
    If Trim(s) <> "" Then
      sSQL = "update observaciones set observacion = '" & s & "' where id = " & CStr(lID) & " "
      Modulo.ExecSQL sSQL
      CargarObservaciones
    End If
  End If
End Sub

Private Sub TotalizarSinPVC()
'  Dim dTotal As Double
'  Dim dST As Double
'  Dim i As Integer
'  Dim sCodPVC As String
'  Dim r As New ADODB.Recordset
'  Dim s As String
'
'  s = "select CodigoProductoPVC from opciones"
'  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
'  sCodPVC = ""
'  If Not r.EOF Then
'    If Not IsNull(r.Fields("CodigoProductoPVC").Value) Then
'      sCodPVC = Trim(r.Fields("CodigoProductoPVC").Value)
'    End If
'  End If
'  r.Close
'  Set r = Nothing
'
'  lTotal.Caption = "0,00"
'  lTO.Caption = "0,00"
'  ePago.Text = "0,00"
'
'  dTotal = 0#
'  dST = 0#
'
'  For i = 0 To aPrecio.ListCount - 1
'    If aCodigo.List(i) <> sCodPVC Then
'      If IsNumeric(aPrecio.List(i)) And IsNumeric(aCantidad.List(i)) Then
'        dST = Val(aPrecio.List(i)) * Val(aCantidad.List(i))
'        dTotal = dTotal + dST
'      End If
'    End If
'  Next i
'
'  lTotal.Caption = Format(dTotal, "#,0.00")
'  lTO.Caption = lTotal.Caption
'  ePago.Text = lTotal.Caption
End Sub

Private Sub TotalizarTODO()
  Dim dTotal As Double
  Dim dST As Double
  Dim i As Integer
  Dim s As String
  
  lTotal.Caption = "0,00"
  lTO.Caption = "0,00"
  ePago.Text = "0,00"
  
  dTotal = 0#
  dST = 0#
  
  For i = 1 To FG.Rows - 1
    If Trim(FG.TextMatrix(i, 0)) <> "" Then
      'Precio y Cantidad:
      If Trim(FG.TextMatrix(i, 3)) <> "" And _
         Trim(FG.TextMatrix(i, 4)) <> "" Then
       
        dST = CDbl(FG.TextMatrix(i, 3)) * CDbl(FG.TextMatrix(i, 4))
        dTotal = dTotal + dST
      Else
        FG.TextMatrix(i, 5) = "0,00"
      End If
    Else
      FG.TextMatrix(i, 5) = "0,00"
    End If
  Next i
  
  lTotal.Caption = Format(dTotal, "#,0.00")
  lTO.Caption = Format(dTotal, "#,0.00")
  ePago.Text = lTO.Caption
End Sub

Private Sub bMarcar_Click()
  Dim i As Integer
  Dim dST As Double
  
  For i = 1 To FG.Rows - 1
  
    'Asignar 1 a cantidad en caso de ser cero.
    If Trim(FG.TextMatrix(i, 4)) <> "" Then
      If CLng(Trim(FG.TextMatrix(i, 4))) = 0 Then
        FG.TextMatrix(i, 4) = "1"
      End If
    End If

  
    FG.Row = i
    FG.Col = 0
    FG.CellFontName = "Wingdings 2"
    FG.CellAlignment = flexAlignCenterCenter
    FG.CellFontSize = 12
    FG.TextMatrix(i, 0) = "P" 'dibuja un Check
    
    Calcular_SubTotal_En_Linea i
    
    Color_HIGH_FILA i

  Next i
  
  TotalizarTODO
 
End Sub

Private Sub C1_Click()
  Dim i As Integer, j As Integer
  Dim s As String, s1 As String
  Dim dST As Double
 
  For i = 0 To lC1.ListCount - 1
    s = Trim(lC1.List(i))
    For j = 1 To FG.Rows - 1
      s1 = Trim(FG.TextMatrix(j, 1)) 'Trim(Mid(lProductos.List(j), 1, 20))
      If s1 = s Then
          
        FG.Row = j
        FG.Col = 0
        bPedirCantidad = False
        Call FG_Click
        bPedirCantidad = True
                    
      End If
    Next j
  Next i
    
  TotalizarTODO
  
End Sub

Private Sub C2_Click()
  Dim i As Integer, j As Integer
  Dim s As String, s1 As String
  Dim dST As Double
 
  For i = 0 To lC2.ListCount - 1
    s = Trim(lC2.List(i))
    For j = 1 To FG.Rows - 1
      s1 = Trim(FG.TextMatrix(j, 1)) 'Trim(Mid(lProductos.List(j), 1, 20))
      If s1 = s Then
          
        FG.Row = j
        FG.Col = 0
        bPedirCantidad = False
        Call FG_Click
        bPedirCantidad = True
                    
      End If
    Next j
  Next i
    
  TotalizarTODO
  
End Sub


Private Sub ePago_Change()
  CalcularPago
End Sub

Private Sub CalcularPago()
  Dim dCambio As Double
  dCambio = 0#
  
  If Trim(ePago.Text) = "" Then
    lCambio.Caption = "0,00"
    Exit Sub
  End If
  
  If IsNumeric(ePago.Text) Then
    If IsNumeric(lTO.Caption) Then
      dCambio = CDbl(ePago.Text) - CDbl(lTO.Caption)
      lCambio.Caption = Format(dCambio, "#,0.00")
    End If
  End If
End Sub

Private Sub ePago_GotFocus()
  ePago.SelStart = 0
  ePago.SelLength = Len(ePago.Text)
End Sub

Private Sub CargarCombo(sNumero As String)
  Dim r As New ADODB.Recordset
  Dim s As String
  s = "select * from CombosProgramados where ComboNumero = '" & sNumero & "' order by id"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If sNumero = "1" Then lC1.Clear Else
  If sNumero = "2" Then lC2.Clear
  
  Do While Not r.EOF
    s = r.Fields("codigoproducto").value
    If sNumero = "1" Then lC1.AddItem s Else
    If sNumero = "2" Then lC2.AddItem s
    
    r.MoveNext
  Loop
  r.Close
  Set r = Nothing
End Sub

Private Sub Color_HIGH_FILA(Fila As Integer)
  Dim j As Integer
  For j = 0 To FG.Cols - 1
    FG.Row = Fila
    FG.Col = j
    FG.CellBackColor = CLR_HIGH_FND
    FG.CellForeColor = CLR_HIGH_TXT
  Next j
End Sub

Private Sub Color_NORMAL_FILA(Fila As Integer)
  Dim j As Integer
  For j = 0 To FG.Cols - 1
    FG.Row = Fila
    FG.Col = j
    FG.CellBackColor = CLR_NORMAL_FND
    FG.CellForeColor = CLR_NORMAL_TXT
  Next j
End Sub

Private Sub Calcular_SubTotal_En_Linea(iLinea As Integer)
  Dim dST As Double
  
  If Trim(FG.TextMatrix(iLinea, 3)) <> "" Then 'Precio
    If Trim(FG.TextMatrix(iLinea, 4)) <> "" Then 'Cantidad
      
      dST = CDbl(Trim(FG.TextMatrix(iLinea, 3))) * _
            CDbl(Trim(FG.TextMatrix(iLinea, 4)))
            
      FG.Row = iLinea
      FG.Col = 5
      FG.CellFontName = cellFontNameOld
      FG.CellAlignment = cellAlignOld
      FG.CellFontSize = cellFontSizeOld
      
      FG.CellAlignment = flexAlignRightCenter
      
      FG.TextMatrix(iLinea, 5) = Format(dST, "#,0.00")
    End If
  End If

End Sub


Public Sub FG_Click()
  Dim dST As Double
  Dim PVC As String
  Dim r As New ADODB.Recordset
    r.Open "select CodigoProductoPVC from opciones", Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    If Not r.EOF Then
      If Not IsNull(r.Fields("CodigoProductoPVC").value) Then
        PVC = Trim(r.Fields("CodigoProductoPVC").value)
      End If
    End If
    r.Close
    Set r = Nothing
  
  
  'If FG.Col <> 0 Then Exit Sub
  If FaltaPago = True Then
     MsgBox "No se puede entregar el producto porque seleccionó Falta Pago", vbExclamation
     Exit Sub
  End If '' And FaltaPago = False
  If FG.Col = 6 Then  'Directo en Entregado:
    If Trim(FG.TextMatrix(FG.Row, 0)) <> CDESMARCA Then
      If FG.TextMatrix(FG.Row, 6) = "NO" Then
        FG.TextMatrix(FG.Row, 6) = "SI"
      Else
        FG.TextMatrix(FG.Row, 6) = "NO"
      End If
    End If
    Exit Sub
  End If
  
  If FG.Col = 0 Then
  
    If Trim(FG.TextMatrix(FG.Row, 0)) <> CDESMARCA Then
    
      FG.CellFontName = cellFontNameOld
      FG.CellAlignment = cellAlignOld
      FG.CellFontSize = cellFontSizeOld
      
      FG.TextMatrix(FG.Row, 0) = CDESMARCA '""
      
      FG.TextMatrix(FG.Row, 4) = "0"
      
      FG.Col = 6: FG.CellAlignment = flexAlignCenterCenter
      FG.TextMatrix(FG.Row, 6) = "NO"
      
      Color_NORMAL_FILA FG.Row
    Else
      
      FG.TextMatrix(FG.Row, 6) = "SI"    'Entregado
       ' el pvc no se entrega hasta que se imprime en el cardfive
      If FG.TextMatrix(FG.Row, 1) = PVC Then FG.TextMatrix(FG.Row, 6) = "NO"
      '********************************************************************
      cellFontNameOld = FG.CellFontName
      cellAlignOld = FG.CellAlignment
      cellFontSizeOld = FG.CellFontSize
      
      FG.Col = 0
      FG.CellFontName = "Wingdings 2"
      FG.CellAlignment = flexAlignCenterCenter
      FG.CellFontSize = 12
      FG.TextMatrix(FG.Row, 0) = CMARCA '"P" 'dibuja un Check
      
      'Asignar 1 a cantidad en caso de ser cero.
      If Trim(FG.TextMatrix(FG.Row, 4)) <> "" Then
        If CLng(Trim(FG.TextMatrix(FG.Row, 4))) = 0 Then
          FG.TextMatrix(FG.Row, 4) = "1"
        End If
      End If
      
      'FG.CellFontName = cellFontNameOld
      'FG.CellAlignment = cellAlignOld
      'FG.CellFontSize = cellFontSizeOld
     
      'FG.TextMatrix(FG.Row, 6) = "SI"    'Entregado
      
      Color_HIGH_FILA FG.Row
      
      'Mostrar un Mensaje en Linea si tiene PRECIO ESPECIAL:
      If EsGuardarPorLotes = False Then EditarPrecioYCantidad
      
      
    End If
  End If

  Calcular_SubTotal_En_Linea FG.Row
  TotalizarTODO
  
End Sub

Private Sub FG_DblClick()

  If Trim(FG.TextMatrix(FG.Row, 0)) <> "" Then
  
    EditarPrecioYCantidad
    
    Calcular_SubTotal_En_Linea FG.Row
  End If
    
   
    
End Sub

Private Sub Form_Load()
  bPedirCantidad = True
  cellFontNameOld = FG.CellFontName
  cellAlignOld = FG.CellAlignment
  cellFontSizeOld = FG.CellFontSize
  
  lCliente.Caption = Modulo.vTemporal2
  lSubCliente.Caption = Modulo.vTemporal3

  CargarObservaciones
  CargarProductos
  TotalizarTODO
  CargarCombo "1"
  CargarCombo "2"
  EsPago = False
  EsPagoPorAdelantado = False
  FaltaPago = False
  EsGuardarPorLotes = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Adodc1.Recordset.Close
  Unload Me
End Sub
Private Sub ExcluirObservacion()
   Dim i As Integer
   If FaltaPago = True Then
      For i = 0 To List1.ListCount - 1
         If UCase(List1.List(i)) = "FALTA PAGO" Then
            List1.Selected(i) = True
         End If
         If UCase(List1.List(i)) = "PAGO" Then
            List1.Selected(i) = False
         End If
         If UCase(List1.List(i)) = "PAGADO POR ADELANTADO" Then
            List1.Selected(i) = False
         End If
         
      Next i
   End If
   
      If EsPago = True Then
      For i = 0 To List1.ListCount - 1
         If UCase(List1.List(i)) = "PAGO" Then
            List1.Selected(i) = True
         End If
         If UCase(List1.List(i)) = "FALTA PAGO" Then
            List1.Selected(i) = False
         End If
         If UCase(List1.List(i)) = "PAGADO POR ADELANTADO" Then
            List1.Selected(i) = False
         End If
      Next i
   End If
   
   If EsPagoPorAdelantado = True Then
      For i = 0 To List1.ListCount - 1
         If UCase(List1.List(i)) = "PAGADO POR ADELANTADO" Then
            List1.Selected(i) = True
         End If
         If UCase(List1.List(i)) = "PAGO" Then
            List1.Selected(i) = False
         End If
         If UCase(List1.List(i)) = "FALTA PAGO" Then
            List1.Selected(i) = False
         End If
      Next i
   End If


End Sub


Private Sub List1_Click()
 Dim i As Integer
 'FaltaPago = False
 'EsPago = False
 'EsPagoPorAdelantado = False
 
 For i = 0 To List1.ListCount - 1
    ''If List1.Selected(i) = True Then
       
       Select Case UCase(List1.List(i))
          Case Is = "FALTA PAGO"
             If List1.Selected(i) = True Then
                FaltaPago = True
                ExcluirObservacion
             Else
                FaltaPago = False
             End If
          Case Is = "PAGO"
             If List1.Selected(i) = True Then
                EsPago = True
                ExcluirObservacion
             Else
                EsPago = False
             End If
             
          Case Is = "PAGADO POR ADELANTADO"
             If List1.Selected(i) = True Then
                EsPagoPorAdelantado = True
                ExcluirObservacion
             Else
                EsPagoPorAdelantado = False
             End If
             
       End Select
       
    ''End If
 Next i
    If FaltaPago = True Then
       For i = 1 To FG.Rows - 1
          FG.TextMatrix(i, 6) = "NO"
       Next i
    End If
End Sub

Private Sub List1_DblClick()
  If List1.ListIndex >= 0 Then
    Modulo.vTemporal1 = List1.List(List1.ListIndex)
    Unload Me
  End If
End Sub

Public Sub CargarProductos()
  Dim s As String
  Dim r As New ADODB.Recordset
  
  Dim SC As String, sSC As String
  Dim dPrecioEspecial As Double
  Dim Fila As Integer
  
  SC = Trim(Mid(lCliente.Caption, 1, 6))
  sSC = Trim(lSubCliente.Caption)
  
  If SC = "-" Or SC = "" Then SC = "0"
  
  If sSC = "-" Or sSC = "" Then sSC = "0" Else sSC = Mid(sSC, 1, 6)
       
  s = "select * from Productos order by descripcion"
  
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  
  'lProductos.Clear
  'lPrecio.Clear
  'lCantidad.Clear
  'lSubTotal.Clear
  
  
  FG.Clear
  FG.Rows = 2
  FG.FormatString = cFG
  Fila = 1
    
  'aCodigo.Clear
  'aProducto.Clear
  'aPrecio.Clear
  'aCantidad.Clear
    
  Do While Not r.EOF
  
      'FG.CellFontName = cellFontNameOld
      'FG.CellAlignment = cellAlignOld
      'FG.CellFontSize = cellFontSizeOld

    FG.TextMatrix(Fila, 0) = ""
    FG.TextMatrix(Fila, 1) = Trim(r.Fields("codigo").value)
    FG.TextMatrix(Fila, 2) = Trim(r.Fields("descripcion").value)
    FG.TextMatrix(Fila, 3) = ""
    FG.TextMatrix(Fila, 4) = "0"
    FG.TextMatrix(Fila, 5) = "0,00"
    
    FG.Row = Fila: FG.Col = 6: FG.CellAlignment = flexAlignCenterCenter
    FG.TextMatrix(Fila, 6) = "NO"
  
    'Buscar primero si tiene precio especial en la tabla de precios acordados:
    'dPrecioEspecial = DBPrecioEspecial(CLng(sC), CLng(sSC), Trim(r.Fields("codigo").Value))
    dPrecioEspecial = -1#
    dPrecioEspecial = Modulo.DBPrecioEspecial(CLng(SC), CLng(sSC), Trim(r.Fields("codigo").value))
    
    'si no tiene precio especial, tomar precio del producto.
    If dPrecioEspecial = -1# Then
       dPrecioEspecial = r.Fields("precio").value
       FG.TextMatrix(Fila, 3) = Format(dPrecioEspecial, "#,0.00")
    Else
      'cellFontNameOld = FG.CellFontName
      'cellAlignOld = FG.CellAlignment
      'cellFontSizeOld = FG.CellFontSize
      
       'FG.Col = 0
       'FG.CellFontName = "Wingdings 2"
       'FG.CellAlignment = flexAlignCenterCenter
       'FG.CellFontSize = 12
       'FG.TextMatrix(Fila, 0) = CMARCA
       FG.TextMatrix(Fila, 3) = Format(dPrecioEspecial, "#,0.00")
       FG.Row = Fila
       FG.Col = 0
        If EsGuardarPorLotes = True And FG.TextMatrix(Fila, 1) <> "D001" And FG.TextMatrix(Fila, 1) <> "SF001" Then FG_Click
    End If
    
    
    r.MoveNext
    Fila = Fila + 1
    If Not r.EOF Then FG.Rows = FG.Rows + 1
  Loop
  
  r.Close
  Set r = Nothing
  
End Sub

Private Sub lProductos_Click()
'  Dim i As Long
'  Dim dST As Double
'  Dim s As String
'
'  i = lProductos.ListIndex
'
'  If i < 0 Then Exit Sub
'
'  If lProductos.Selected(i) Then
'    aCantidad.List(i) = "1"
'  Else
'    aCantidad.List(i) = "0"
'  End If
'
'  dST = Val(aPrecio.List(i)) * Val(aCantidad.List(i))
'
'  s = Modulo.BlancosDER(aProducto.List(i), 50) & " " & _
'      Modulo.BlancosIZQ(Format(Val(aPrecio.List(i)), "#,0.00"), 10) & " " & _
'      Modulo.BlancosIZQ(aCantidad.List(i), 10) & " " & _
'      Modulo.BlancosIZQ(Format(dST, "#,0.00"), 15)
'
'  lProductos.List(i) = s
'
'  TotalizarTODO
End Sub

Private Sub lProductos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Button = vbRightButton Then
    ' Depliega el menú PopUP
    'Me.PopupMenu mnucontextual
  End If

End Sub

Private Sub EditarPrecioYCantidad()
  Dim i As Long
  Dim dST As Double
  Dim s As String
  Dim PP As Double, PA As Double  'Precio Producto  / Precio Acordado
  Dim dPrecio As Double
  Dim sCod As String, sCan As String
  Dim iC As Long, iSC As Long
  
  If Not bPedirCantidad Then Exit Sub
    
  i = FG.Row 'Productos.ListIndex
  If i <= 0 Then Exit Sub
 
  sCod = FG.TextMatrix(FG.Row, 1)   'Trim(aCodigo.List(lProductos.ListIndex))
  PP = 0#
  PA = 0#
  
  PP = Modulo.DBValorDouble("productos", "codigo", sCod, "Precio")
    
  dPrecio = PP
  
  s = Trim(Mid(lCliente.Caption, 1, 6))
  If IsNumeric(s) Then iC = CLng(s)
  
  s = Trim(Mid(lSubCliente.Caption, 1, 6))
  If IsNumeric(s) Then iSC = CLng(s) Else iSC = 0
    
    
  PA = Modulo.DBPrecioEspecial(iC, iSC, sCod)
    
  If PA > 0# Then
    MsgBox "Precio Venta ACTUAL este Producto = " & Format(PP, "#,0.00") & vbCrLf & _
           "Precio Venta ESPECIAL este Cliente = " & Format(PA, "#,0.00"), vbCritical, "ATENCIÓN"
             
    dPrecio = PA
  End If
  
  Load fMovDetalle
  With fMovDetalle
    .lcod.Caption = sCod 'aCodigo.List(i)
    .lDes.Caption = FG.TextMatrix(i, 2)   'aProducto.List(i)
    .ePre.Text = FG.TextMatrix(i, 3)      'Format(dPrecio, "#,0.00")
      
    .eCan.Text = Trim(FG.TextMatrix(i, 4))
    If .eCan.Text = "" Then .eCan.Text = "0"
     
    .eST.Text = Format(CDbl(.ePre.Text) * CDbl(.eCan.Text), "#,0.00")
    
    .ePre.Enabled = False
    .eST.Enabled = False
      
    If PA > 0 Then .lPrecioEspecial.Visible = True
  End With
  
  Modulo.fModalResult = Modulo.fModalResultCANCEL
  
  fMovDetalle.Show vbModal
  
  If Modulo.fModalResult = Modulo.fModalResultOK Then
  
    'Precio:
    FG.TextMatrix(FG.Row, 3) = Format(CDbl(fMovDetalle.ePre.Text), "#,0.00")
    
    'Cantidad:
    FG.TextMatrix(FG.Row, 4) = Format(CDbl(fMovDetalle.eCan.Text), "#,0")
    
    'SubTotal:
    Calcular_SubTotal_En_Linea FG.Row
  End If
  
  Unload fMovDetalle
  TotalizarTODO

End Sub

Private Sub mEdi_Click()
  EditarPrecioYCantidad
End Sub

Private Sub mMarDes_Click()
'  Dim i As Integer
'  i = lProductos.ListIndex
'  If i >= 0 Then
'    lProductos.Selected(i) = Not lProductos.Selected(i)
'  End If
End Sub

