VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form fCargaClientes 
   Caption         =   "Carga de Clientes (Codificación)"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10080
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Carpeta Origen"
      Height          =   9915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15195
      Begin MSComctlLib.ListView ListViewArchivo 
         Height          =   3015
         Left            =   5280
         TabIndex        =   41
         Top             =   600
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "STRING"
            Text            =   "Archivo"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "NUMBER"
            Text            =   "Tamaño"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "DATE"
            Text            =   "Ultima Modificacion"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   12240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4620
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.FileListBox File3 
         Height          =   870
         Left            =   13740
         TabIndex        =   40
         Top             =   3840
         Width           =   1155
      End
      Begin MSComDlg.CommonDialog Dialog1 
         Left            =   13860
         Top             =   4440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.DirListBox Dir2 
         Height          =   765
         Left            =   7800
         TabIndex        =   37
         Top             =   3840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.FileListBox File2 
         Height          =   480
         Left            =   6240
         TabIndex        =   35
         Top             =   3900
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   11280
         Top             =   1500
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   1
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
         Caption         =   "Adodc3"
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   11280
         Top             =   1110
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   1
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
         Caption         =   "Adodc2"
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
      Begin VB.CommandButton bBuscarCP 
         Height          =   345
         Left            =   5190
         Picture         =   "fCargaClientes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4650
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cCP 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   4680
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CheckBox cEsSubCliente 
         Caption         =   "Es Subcliente."
         Height          =   195
         Left            =   60
         TabIndex        =   23
         Top             =   4440
         Width           =   1335
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFC0FF&
         Height          =   700
         Left            =   120
         ScaleHeight     =   645
         ScaleWidth      =   6795
         TabIndex        =   17
         Top             =   9150
         Width           =   6855
         Begin VB.CommandButton bNuevaCol 
            Caption         =   "Nueva Columna"
            Height          =   345
            Left            =   60
            TabIndex        =   21
            Top             =   150
            Width           =   1365
         End
         Begin VB.CommandButton bIzq 
            Caption         =   "< Mover Columna IZQ."
            Height          =   375
            Left            =   2940
            TabIndex        =   20
            Top             =   120
            Width           =   1845
         End
         Begin VB.CommandButton bDer 
            Caption         =   "Mover Columna DER. >"
            Height          =   375
            Left            =   4860
            TabIndex        =   19
            Top             =   120
            Width           =   1845
         End
         Begin VB.CommandButton bBorrarCOL 
            Caption         =   "Borrar Columna"
            Height          =   345
            Left            =   1500
            TabIndex        =   18
            Top             =   150
            Width           =   1365
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFC0FF&
         Height          =   700
         Left            =   7470
         ScaleHeight     =   645
         ScaleWidth      =   6420
         TabIndex        =   15
         Top             =   9150
         Width           =   6480
         Begin VB.CommandButton cmdCopiarFotosEnSitio 
            Caption         =   "Copiar Fotos    (en sitio)"
            Height          =   465
            Left            =   2940
            TabIndex        =   39
            Top             =   60
            Width           =   1365
         End
         Begin VB.CommandButton bCrearPersonasCopiar 
            Caption         =   "Copiar Registros de Personas"
            Height          =   465
            Left            =   4500
            TabIndex        =   38
            Top             =   60
            Width           =   1695
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Copiar Carpetas"
            Height          =   495
            Left            =   1620
            TabIndex        =   36
            Top             =   60
            Width           =   1035
         End
         Begin VB.CommandButton bCrearCliente 
            Caption         =   "Crear Cliente"
            Height          =   465
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   1365
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Ver SÓLO Archivos .MDB"
         Height          =   225
         Left            =   9450
         TabIndex        =   14
         Top             =   300
         Width           =   2355
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ver TODOS los Archivos"
         Height          =   225
         Left            =   6180
         TabIndex        =   13
         Top             =   300
         Value           =   -1  'True
         Width           =   2355
      End
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   13770
         TabIndex        =   11
         Top             =   2610
         Width           =   1275
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00FFFFC0&
         Height          =   1815
         Left            =   60
         ScaleHeight     =   1755
         ScaleWidth      =   15015
         TabIndex        =   9
         Top             =   7260
         Width           =   15075
         Begin VB.ListBox List3 
            Height          =   1035
            Left            =   6600
            TabIndex        =   34
            Top             =   450
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ListBox List2 
            Height          =   1035
            Left            =   5010
            TabIndex        =   28
            Top             =   450
            Visible         =   0   'False
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid FG2 
            Height          =   1455
            Left            =   90
            TabIndex        =   10
            Top             =   210
            Width           =   14835
            _ExtentX        =   26167
            _ExtentY        =   2566
            _Version        =   393216
            Rows            =   3
            FixedCols       =   0
            AllowUserResizing=   3
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "Estructura Predeterminada de Tabla para Nuevo Cliente"
            Height          =   195
            Left            =   5430
            TabIndex        =   26
            Top             =   15
            Width           =   3960
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   11280
         Top             =   720
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
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
         Caption         =   "Adodc1"
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
      Begin VB.CommandButton bLeerMDB 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Leer Estructura/Registros"
         Height          =   435
         Left            =   10770
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3840
         Width           =   2925
      End
      Begin VB.CommandButton bFijarIO 
         BackColor       =   &H00FFC0FF&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5190
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fijar Nombre de Cliente"
         Top             =   4020
         Width           =   375
      End
      Begin VB.TextBox eCliente 
         Height          =   345
         Left            =   60
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   4020
         Width           =   5085
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
         TabIndex        =   4
         Top             =   180
         Width           =   5070
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
         Height          =   3150
         Left            =   90
         TabIndex        =   3
         Top             =   600
         Width           =   5040
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
         Height          =   3015
         Left            =   5280
         System          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton bCancelar 
         Caption         =   "Cancelar"
         Height          =   500
         Left            =   14160
         Picture         =   "fCargaClientes.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   9240
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   1815
         Left            =   60
         TabIndex        =   27
         Top             =   5370
         Width           =   14835
         _ExtentX        =   26167
         _ExtentY        =   3201
         _Version        =   393216
         Rows            =   3
         FixedCols       =   0
         HighLight       =   0
         AllowUserResizing=   3
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13740
         TabIndex        =   33
         Top             =   5100
         Width           =   705
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Registros:"
         Height          =   255
         Left            =   12570
         TabIndex        =   32
         Top             =   5100
         Width           =   1170
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9870
         TabIndex        =   31
         Top             =   5100
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9150
         TabIndex        =   30
         Top             =   5100
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código Asignado:"
         Height          =   255
         Left            =   7830
         TabIndex        =   29
         Top             =   5100
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Indique el Cliente Principal:"
         Height          =   195
         Left            =   1440
         TabIndex        =   22
         Top             =   4440
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TABLA"
         Height          =   255
         Left            =   13770
         TabIndex        =   12
         Top             =   2370
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre del Cliente o SubCliente a CREAR:"
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   3780
         Width           =   3060
      End
   End
End
Attribute VB_Name = "fCargaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const cFG = "Nº      | ARCHIVO                                           | CÉDULA                                                "

Const FGROWS = 4

Dim ConexionAccess As String
Dim RutaAccess As String

Const MAXCAMPOSPRE = 10

Dim aCamposPre(MAXCAMPOSPRE) As String
Dim aCamposOcul(MAXCAMPOSPRE) As String
Dim aTiposPre(MAXCAMPOSPRE) As String
Dim aAnchosPre(MAXCAMPOSPRE) As Integer

Const MAXLONGCLIENTE = 100
Dim lCodigoCliente As String
Dim lTablaVacia As Boolean
Public lDirCarnet As String
Public lDirFotos As String
Public lDirImagenes As String
Dim FSO As New FileSystemObject


Private Sub CargarCamposPre()  '--Predeterminados
  'aCamposPre(0) = "ID"
  
  aCamposPre(0) = "CEDULA"
  aCamposPre(1) = "NOMBRE"
  aCamposPre(2) = "CARGO"
  aCamposPre(3) = "VENCE"
  
  aCamposPre(7) = "FECHA"
  aCamposPre(8) = "CONTADOR"
  
  aCamposPre(4) = "FOTO"
  aCamposPre(5) = "TIENE_FOTO"
  aCamposPre(6) = "MARCA"
  aCamposPre(9) = "CREACION"
    
  'aTiposPre(0) = "INTEGER"
  aTiposPre(0) = "CHAR"
  aTiposPre(1) = "CHAR"
  aTiposPre(2) = "CHAR"
  aTiposPre(3) = "CHAR"
  
  aTiposPre(7) = "DATETIME"
  aTiposPre(8) = "INTEGER"
  
  aTiposPre(4) = "CHAR"
  aTiposPre(5) = "CHAR"
  aTiposPre(6) = "CHAR"
  aTiposPre(9) = "DATETIME"
   
  'aAnchosPre(0) = 0
  aAnchosPre(0) = 20
  aAnchosPre(1) = 50
  aAnchosPre(2) = 50
  aAnchosPre(3) = 30
  aAnchosPre(7) = 0
  aAnchosPre(8) = 0
  aAnchosPre(4) = 20 'LONGITUD DEL CAMPO FOTO
  aAnchosPre(5) = 1
  aAnchosPre(6) = 1
  aAnchosPre(9) = 0
   
  'aCamposOcul(0) = ""
  aCamposOcul(0) = ""
  aCamposOcul(1) = ""
  aCamposOcul(2) = ""
  aCamposOcul(3) = ""
  aCamposOcul(4) = "FOTO"
  aCamposOcul(5) = "TIENE_FOTO"
  aCamposOcul(6) = "MARCA"
  aCamposOcul(9) = "CREACION"
  
  aCamposOcul(7) = "FECHA"
  aCamposOcul(8) = "CONTADOR"
  
 
End Sub

Private Function ELANCHOPRE(sCampo As String) As Integer
  Dim i As Integer
  Dim ea As Integer
  Dim e As Boolean
  i = 0
  e = False
  Do While i < MAXCAMPOSPRE And Not e
    If sCampo = aCamposPre(i) Then
      ea = i
      e = True
    Else
      i = i + 1
    End If
  Loop
  ea = 0
  If e Then ea = aAnchosPre(i)
  ELANCHOPRE = ea
End Function

Private Function ELTIPOPRE(sCampo As String) As String
  Dim i As Integer
  Dim ea As String
  Dim e As Boolean
  
  If sCampo = "FECHA" Then ELTIPOPRE = "DATETIME"
  If sCampo = "CONTADOR" Then ELTIPOPRE = "INTEGER"
  If sCampo = "FECHA" Or sCampo = "CONTADOR" Then Exit Function
   
  i = 0
  e = False
  Do While i < MAXCAMPOSPRE And Not e
    If sCampo = aCamposPre(i) Then
      ea = i
      e = True
    Else
      i = i + 1
    End If
  Loop
  ea = ""
  If e Then ea = aTiposPre(i)
  ELTIPOPRE = ea
End Function


Private Function EsCampoPre(sCualCampo As String) As Boolean
  Dim i As Integer
  Dim ECP As Boolean
  ECP = False
  i = 0
  Do While (i < MAXCAMPOSPRE) And Not ECP
    If aCamposPre(i) = sCualCampo Then ECP = True Else i = i + 1
  Loop
  EsCampoPre = ECP
End Function

Private Function EsCampoOcul(sCualCampo As String) As Boolean
  Dim i As Integer
  Dim ECP As Boolean
  ECP = False
  i = 0
  Do While (i < MAXCAMPOSPRE) And Not ECP
    If aCamposOcul(i) = sCualCampo Then ECP = True Else i = i + 1
  Loop
  EsCampoOcul = ECP
End Function



Private Sub bBorrarCOL_Click()
  Dim c As Integer, i As Integer, e As Boolean
  
  i = 0
  e = False
  Do While i < FG1.Cols And Not e
    FG1.Row = 1
    FG1.Col = i
    If FG1.CellBackColor = vbRed Then e = True Else i = i + 1
  Loop
  
  If Not e Then Exit Sub
    
  c = i
  
  If EsCampoPre(FG1.TextMatrix(0, c)) Then
    MsgBox "El Campo [" & FG1.TextMatrix(0, c) & "] no se puede Borrar...", vbCritical, "Información"
    Exit Sub
  End If

  If MsgBox("¿Está Seguro de Borrar la COLUMNA Nº " & CStr(c + 1) & "?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
  
    FG1.TextMatrix(0, c) = ""
    FG1.TextMatrix(1, c) = ""
        
    If c < FG1.Cols - 1 Then
      For i = c To FG1.Cols - 2
        FG1.TextMatrix(0, i) = FG1.TextMatrix(0, i + 1)
        FG1.TextMatrix(1, i) = FG1.TextMatrix(1, i + 1)
        FG1.TextMatrix(2, i) = FG1.TextMatrix(2, i + 1)
      Next i
      FG1.Cols = FG1.Cols - 1
      Modulo.Ajustar_Columnas_FLEXGRID FG1
    Else
      FG1.Cols = FG1.Cols - 1
      Modulo.Ajustar_Columnas_FLEXGRID FG1
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
    cCP.ListIndex = Modulo.Buscar_ComboLen(cCP, Mid(Modulo.vTemporal1, 1, 6), 6)
  End If

End Sub

Private Sub bCancelar_Click()
  Adodc1.ConnectionString = ""
  Unload Me
End Sub

Private Sub bCarpeta_Click()
  If Dir1.ListIndex >= 0 Then eCliente.Text = Dir1.List(Dir1.ListIndex)
End Sub


Private Sub bCrearCliente_Click()
  Dim s As String, s1 As String
  Dim lcod As Long
  Dim lCn As New ADODB.Connection
  s = Trim(eCliente.Text)
  If Modulo.DBExiste("clientes", "nombre", s) Then
    lcod = CLng(Modulo.DBValorStr("clientes", "nombre", s, "codigo"))
    s1 = Zeros(lcod, 6)
    MsgBox "Ya Existe un Cliente con el Nombre [" & s & "] " & vbCrLf & _
           "tiene Código [" & s1 & "], Revise...", vbCritical, "Información"
    Label6.Caption = s1
    Exit Sub
  End If
      
  ''Label6.Caption = s1
  
  
  
  If Not eCliente.Enabled Then
  
    eCliente.Text = UCase(eCliente.Text)
    
    If MsgBox("¿Está Seguro de Crear Cliente/SubCliente [" & eCliente.Text & "]?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
      Load fMensaje
      fMensaje.Caption = "Creando Cliente, Espere..."
      If cEsSubCliente.value = vbChecked Then
        fMensaje.Caption = "Creando Subcliente, Espere..."
      End If
      
      fMensaje.Show
      DoEvents
      lCn.ConnectionString = Modulo.DBConexionSQL.ConnectionString
      lCn.Open
      Adodc2.ConnectionString = Modulo.DBConexionSQL.ConnectionString
      Adodc2.RecordSource = "select * from clientes order by codigo"
      
      If cEsSubCliente.value = vbChecked Then
        Adodc2.RecordSource = "select * from subclientes order by id"
      End If
      
      Adodc2.Refresh
        If cEsSubCliente.value = vbChecked Then
          lCn.Execute "Insert Into SubClientes (Cliente,Nombre,FechaInicio,Activo) Values(" & CLng(Label6.Caption) & ",'" _
           & Trim(eCliente.Text) & "','" & Format(Date, "yyyyMMdd") & "','S')"
        Else
          lCn.Execute "Insert Into Clientes (Nombre,FechaInicio,Activo) Values('" _
           & Trim(eCliente.Text) & "','" & Format(Date, "yyyyMMdd") & "','S')"
        End If
      'On Error Resume Next
      
      'With Adodc2.Recordset
      '  .AddNew
      '
      '  If cEsSubCliente.Value = vbChecked Then
      '    .Fields("cliente").Value = CLng(Label6.Caption)
      '  End If
      '  '.Fields("Codigo").Value = 99
      '  .Fields("rif").Value = ""
      '  .Fields("nit").Value = ""
      '  .Fields("nombre").Value = Trim(eCliente.Text)
      '  .Fields("direccion").Value = ""
      '  .Fields("telefonos").Value = ""
      '  .Fields("fax").Value = ""
      '  .Fields("email").Value = ""
      '  .Fields("website").Value = ""
      '  .Fields("contacto").Value = ""
      '  .Fields("contactotlf").Value = ""
      '  .Fields("fechainicio").Value = Date
      '  .Fields("activo").Value = "S"
              
      '  .Fields("deuda").Value = 0#
      '  .Fields("pagos").Value = 0#
      '  .Fields("saldo").Value = 0#
     '
     '   .Fields("cedulaauto1").Value = ""
     '   .Fields("nombreauto1").Value = ""
     '   .Fields("cargoauto1").Value = ""
     '   .Fields("telefauto1").Value = ""
     '   .Fields("fotoauto1").Value = ""
     '
     '   .Fields("cedulaauto2").Value = ""
     '   .Fields("nombreauto2").Value = ""
     '   .Fields("cargoauto2").Value = ""
     '   .Fields("telefauto2").Value = ""
     '   .Fields("fotoauto2").Value = ""
     '
     '   .Fields("cedulaauto3").Value = ""
     '   .Fields("nombreauto3").Value = ""
     '   .Fields("cargoauto3").Value = ""
     '   .Fields("telefauto3").Value = ""
     '   .Fields("fotoauto3").Value = ""
     '
     '   .Update
     ' End With
      
      Unload fMensaje
      
      'If Err.Number <> 0 Then
      '  MsgBox "Error Agregando Cliente Nuevo...", vbCritical, "Información"
      '  Exit Sub
      'End If
      
      Adodc2.Recordset.Close
      
      Adodc2.RecordSource = "select * from clientes where nombre = '" & Trim(eCliente.Text) & "'"
           
      If cEsSubCliente.value = vbChecked Then
        Adodc2.RecordSource = "select * from subclientes where cliente = " & Label6.Caption & " and nombre = '" & Trim(eCliente.Text) & "'"
      End If
        
      Adodc2.Refresh
      
      
      
      If Adodc2.Recordset.EOF Then
        MsgBox "El Cliente o Subcliente No Fue Agregado no puede Ser Localizado...", vbCritical, "Error"
        Exit Sub
      End If
      
      If Not Adodc2.Recordset.EOF Then
        If cEsSubCliente.value = vbChecked Then
          s = Zeros(Adodc2.Recordset.Fields("cliente").value, 6) & "-" & Zeros(Adodc2.Recordset.Fields("id").value, 6)
          MsgBox "Código Asignado [ " & s & " ] para el Subcliente: " & vbCrLf & _
                 Trim(eCliente.Text), vbInformation, "Información"
          Label7.Caption = Zeros(Adodc2.Recordset.Fields("id").value, 6)
          fSubClientes.SCargarfSubClientes
          fSubClientes.cCP.Text = cCP.Text

          fSubClientes.FG.Row = fSubClientes.FG.Rows - 1
          fSubClientes.Show
          fSubClientes.bEditar_Click
          
        Else
          s = Zeros(Adodc2.Recordset.Fields("codigo").value, 6)
          MsgBox "Código Asignado [ " & s & " ] para el Cliente: " & vbCrLf & _
                  Trim(eCliente.Text), vbInformation, "Información"
          Label6.Caption = s
          lCodigoCliente = s
          Load fClientes
          fClientes.FG.Row = fClientes.FG.Rows - 1
          fClientes.bEditar_Click
          fClientes.Show
        End If
      End If
    End If
  End If
End Sub




Function Numero_Siguiente_Tabla_Personas() As Integer
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim SC As String, sSC As String
  Dim k As Integer, n As Integer
  
  If Label6.Caption = "" Or Label6.Caption = "-" Then
    MsgBox "Debe Seleccionar el Cliente antes de Realizar esta Operación...", vbCritical, "Información"
    Exit Function
  End If
    
  SC = Label6.Caption   'Trim(Mid(cCP.Text, 1, 6))
  sSC = Label7.Caption  'Trim(Mid(cSC.Text, 1, 6))
  If sSC = "" Or sSC = "-" Then sSC = "0"
  
  s = "select * from Personas where " & _
      "cliente    = " & SC & " and " & _
      "subcliente = " & sSC & " order by id "
      
  k = 0

  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  Do While Not r.EOF
    s = Trim(r.Fields("tabla").value)
    s = Trim(Mid(s, InStrRev(s, "-") + 1))
    k = CInt(s)
    r.MoveNext
  Loop
  r.Close
  Set r = Nothing
  Numero_Siguiente_Tabla_Personas = k + 1
End Function








'Crear Tabla de Datos Personas y el trigger respectivo

Private Sub bCrearPersonasCopiar_Click()
  Dim r As New ADODB.Recordset
  Dim cp As String, SC As String
  Dim s As String, sComando As String, sTablaNueva As String
  Dim i As Integer, iNumeroTablaNueva As Integer
  Dim sCampo As String, sTipo As String, sAncho As String
  Dim sComando2 As String
  Dim ii As Integer, ss As String
  
  Dim sCP As String 'CLIENTE PRINCIPAL
  Dim sSC As String 'SUB-CLIENTE
  Dim Fila As Integer
  Dim sCI As String
  
  
  If Label7.Caption = "" Then Label7.Caption = "-"
  
  sCP = Label6.Caption
  sSC = Label7.Caption
  
  If Label6.Caption = "-" Then
    MsgBox "Debe Seleccionar/Fijar el Cliente...", vbCritical, "Información"
    Exit Sub
  End If
  
  
  
  
    
'  cp = Trim(Mid(cCP.Text, 1, 6))
'  sC = Trim(Mid(cSC.Text, 1, 6))
'
'  If Trim(sC) = "" Then sC = "0"
'
'  If lCampos.ListCount <= 0 Then
'    MsgBox "Debe Indicar los campos de la Tabla para poder Crearla...", vbCritical, "Información"
'    Exit Sub
'  End If
  
  'iNumeroTablaNueva = Numero_Siguiente_Tabla_Personas()
  
  'sTablaNueva = cp & "-" & IIf(sc <> "0", sc & "-", "") & CStr(iNumeroTablaNueva)
  ''If MsgBox("Recuerde que para poder copiar el contenido de la carpeta del cliente debe estar ubicado en la raíz de la misma.¿Continuar?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

  s = ""
  s = Modulo.La_Tabla_Actual_Personas(sCP, sSC)
  If s <> "" Then
     Select Case MsgBox("El Cliente actualmente tiene Tabla de Personas [" & s & "]..." & vbCrLf & _
              "Haga click en 'SI' si Desea crearla nuevamente y Reemplazar todos los Datos" & Chr(10) _
              & "O Haga Click en No para seguir agregando datos a la tabla" & Chr(10) _
               , vbYesNoCancel + vbDefaultButton2, "Confirme")
       Case vbYes
        If MsgBox("¿Está Seguro?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
               
          sComando = "DROP TABLE [" & s & "]"
          Modulo.ExecSQL sComando
         
          sComando = "DROP TABLE [H" & s & "]"
          Modulo.ExecSQL sComando
        
        
          sComando = "DELETE FROM Personas WHERE Tabla = '" & s & "'"
          Modulo.ExecSQL sComando
        End If
       Case vbCancel
          Exit Sub
       Case vbNo
          Adodc3.ConnectionString = Modulo.DBConexionSQL.ConnectionString
          Adodc3.RecordSource = "select * from [" & s & "]"
          Adodc3.Refresh
          fMensaje2.ProgressBar1.value = 0
          fMensaje2.ProgressBar1.Max = FG1.Rows - 1
          For Fila = 1 To FG1.Rows - 1
          
            fMensaje2.Caption = "Copiando Registro " & CStr(Fila) & " / " & CStr(FG1.Rows)
            fMensaje2.Show
             
            fMensaje2.ProgressBar1.value = fMensaje2.ProgressBar1.value + 1
            
            DoEvents
          
            Adodc3.Recordset.AddNew
            
            sCI = ""
                      
            For i = 0 To FG1.Cols - 1
              
              s = FG1.TextMatrix(0, i) 'El Campo segun Cabecera de FlexGrid
              
              Adodc3.Recordset.Fields(s).value = IIf((Trim(FG1.TextMatrix(Fila, i)) = ""), Null, Trim(FG1.TextMatrix(Fila, i)))
              
              If s = "CEDULA" Then sCI = Trim(FG1.TextMatrix(Fila, i))
              
            Next i
          
            '-Asignar los campos predeterminados ocultos
            
            
            Adodc3.Recordset.Fields("FOTO").value = Trim(IIf(sCI = "", "", Trim(DepurarStr(DepurarStr(DepurarStr(sCI, ","), "."), " "))))
            Adodc3.Recordset.Fields("TIENE_FOTO").value = "N"
            Adodc3.Recordset.Fields("MARCA").value = ""
            
            If Not Modulo.EXISTE_CAMPO_EN_FLEXGRID(FG1, "FECHA") Then
              Adodc3.Recordset.Fields("FECHA").value = Null
            End If
              
            If Not Modulo.EXISTE_CAMPO_EN_FLEXGRID(FG1, "CONTADOR") Then
              Adodc3.Recordset.Fields("CONTADOR").value = 0
            End If
            
            Adodc3.Recordset.Fields("CREACION").value = Date
            
            Adodc3.Recordset.Update
            
          Next Fila
          
          Adodc3.Recordset.Close
          Adodc3.ConnectionString = ""
          
          fMensaje2.Caption = "Auditando FOTOS, Espere..."
          DoEvents
          
          Modulo.Auditar_Fotos sTablaNueva
         
          Unload fMensaje2
          
          MsgBox "La data se copió correctamente", vbInformation
          Exit Sub
     End Select
      
  Else
    
     ''Exit Sub
      
  End If

  
            
  sCP = Label6.Caption
  sSC = Label7.Caption
            
  
  iNumeroTablaNueva = Numero_Siguiente_Tabla_Personas()
  
  sTablaNueva = sCP & "-" & IIf(sSC <> "-", sSC & "-", "") & CStr(iNumeroTablaNueva)
  
  If sSC = "-" Or sSC = "" Then
     sSC = "0"
  End If
  If lTablaVacia = False Then

     If MsgBox("¿Está Seguro de CREAR la Tabla Nueva [" & sTablaNueva & "]?", vbQuestion + vbYesNo, "Confirme") = vbNo Then
       Exit Sub
     End If
  End If
  sComando = "CREATE TABLE [" & sTablaNueva & "] " & _
             "(ID INT NOT NULL IDENTITY(1,1), "

  'sComando = "CREATE TABLE [" & sTablaNueva & "] ("
             
  'CargarCamposPre
  '--Lleva <H>istorico
  sComando2 = "CREATE TABLE [H" & sTablaNueva & "] " & _
              "(ID INT NOT NULL, "
              
  'sComando2 = "CREATE TABLE [H" & sTablaNueva & "] ("
  
  For i = 0 To FG1.Cols - 1
    If UCase(Trim(FG1.TextMatrix(0, i))) <> "ID" Then
  
      sCampo = FG1.TextMatrix(0, i)
      
      sTipo = ELTIPOPRE(sCampo) 'lTipos.List(i)
      If sTipo = "" Then sTipo = "CHAR"
            
      sAncho = CStr(ELANCHOPRE(sCampo))  'lAnchos.List(i)
      If sAncho = "" Or sAncho = "0" Then sAncho = "50"
      
      'buscar ancho en los personalizados:
      For ii = 0 To List2.ListCount - 1
        ss = List2.List(ii)
        If ss = sCampo Then sAncho = List3.List(ii)
      Next ii
         
    
      If sTipo = "CHAR" Then  '--Lleva ancho
        s = sCampo & " " & sTipo & "(" & sAncho & ") "
      Else
        s = sCampo & " " & sTipo & " "
      End If
    
      'If i <> FG1.Cols - 1 Then s = s & ","
      s = s & ","
      If (sCampo <> "FECHA") And (sCampo <> "CONTADOR") And (UCase(sCampo) <> "NROFOTO") Then
         sComando = sComando & s
         sComando2 = sComando2 & s
      End If
      
    End If
  Next i
  
  
    ''Clipboard.Clear
    ''Clipboard.SetText sComando

  'ahora anexarle los campos predeterminados pero ocultos:
  For i = 0 To MAXCAMPOSPRE - 1
    If Trim(aCamposOcul(i)) <> "" Then
      sCampo = aCamposOcul(i)
      sTipo = ELTIPOPRE(sCampo) 'lTipos.List(i)
      sAncho = CStr(ELANCHOPRE(sCampo))  'lAnchos.List(i)
    
      If sTipo = "CHAR" Then  '--Lleva ancho
        s = sCampo & " " & sTipo & "(" & sAncho & ") "
      Else
        s = sCampo & " " & sTipo & " "
             End If
    
      'If i <> MAXCAMPOSPRE - 1 Then s = s & ","
      s = s & ","
    
      sComando = sComando & s
      sComando2 = sComando2 & s
    End If
    'Clipboard.Clear
    'Clipboard.SetText sComando
  Next i
    
    
    'If Not Modulo.EXISTE_CAMPO_EN_FLEXGRID(FG1, "FECHA") Then
    'If Mid(Trim(sComando), Len(Trim(sComando)), 1) <> "," Then
    '  sComando = sComando & ", FECHA DATETIME "
    '  sComando2 = sComando2 & ", FECHA DATETIME "
    '  Clipboard.Clear
    '  Clipboard.SetText sComando
    'Else
    '  sComando = sComando & " FECHA DATETIME "
    '  sComando2 = sComando2 & " FECHA DATETIME "
    'End If
  'End If
  
  'If Not Modulo.EXISTE_CAMPO_EN_FLEXGRID(FG1, "CONTADOR") Then
  '  If Mid(Trim(sComando), Len(Trim(sComando)), 1) <> "," Then
  '    sComando = sComando & ", CONTADOR INTEGER, "
  '    sComando2 = sComando2 & ", CONTADOR INTEGER, "
  '  Else
  '    sComando = sComando & " CONTADOR INTEGER, "
  '    sComando2 = sComando2 & " CONTADOR INTEGER, "
  '  End If
  'End If

  sComando = sComando & " PRIMARY KEY(ID))"
  'sComando = sComando & " )"
  
  'sComando2 = sComando2 & " PRIMARY KEY(ID))"
  sComando2 = Mid(sComando2, 1, Len(sComando2) - 1)
  sComando2 = sComando2 & " )"
  
  On Error Resume Next
  
  '-- Crear Tabla de Datos <Personas> que conectará con Card-5:
  Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
  Modulo.DBComandoSQL.CommandText = sComando
  Clipboard.Clear
  Clipboard.SetText sComando
  If lTablaVacia = False Then Modulo.DBComandoSQL.Execute
  
  If Err.Number <> 0 Then
    MsgBox "Error: No se pudo Crear la Tabla " & sTablaNueva & vbCrLf & Err.Description, vbCritical, "Información"
  Else
    '-- Crear Tabla de Datos <Personas> tipo Historico de Auditoria que
    '-- actualizará el Card-5 y mediante un Trigger desde SQL-Server:
    Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
    Modulo.DBComandoSQL.CommandText = sComando2
    ''Clipboard.Clear
    ''Clipboard.SetText sComando2
    If lTablaVacia = False Then Modulo.DBComandoSQL.Execute
    If Err.Number <> 0 Then
      MsgBox "Error: No se pudo Crear la Tabla H" & sTablaNueva & vbCrLf & Err.Description, vbCritical, "Información"
    Else
      '--Agregar TRIGGER en tabla nueva SQL Server:
      'sComando = "CREATE TRIGGER [TRG_" & sTablaNueva & "] ON [" & sTablaNueva & "] " & vbCrLf & _
                 "FOR UPDATE AS " & vbCrLf & _
                 "SET IDENTITY_INSERT [TRG_" & sTablaNueva & "] ON " & _
                 "INSERT INTO [H" & sTablaNueva & "] SELECT * FROM DELETED"

      'sComando = "CREATE TRIGGER [TRG_" & sTablaNueva & "] ON [" & sTablaNueva & "] " & vbCrLf & _
                 "FOR UPDATE AS " & vbCrLf & _
                 "INSERT INTO [H" & sTablaNueva & "] SELECT * FROM DELETED"
                 
      'sComando = "CREATE TRIGGER [TRG_" & sTablaNueva & "] ON [" & sTablaNueva & "] " & vbCrLf & _
                 "FOR UPDATE AS " & vbCrLf & _
                 "BEGIN " & vbCrLf & _
                 "  declare @CodProducto as char(20) " & vbCrLf & _
                 "  declare @PrecioProducto as float " & vbCrLf & _
                 "  declare @IDPersona as integer " & vbCrLf & _
                 "  --INSERT INTO [H" & sTablaNueva & "] SELECT * FROM DELETED " & vbCrLf & _
                 "  if update(CONTADOR) " & vbCrLf & _
                 "  begin " & vbCrLf & _
                 "    INSERT INTO [H" & sTablaNueva & "] SELECT * FROM DELETED " & vbCrLf & _
                 "    set @IDPersona = (SELECT ID FROM DELETED) " & vbCrLf & _
                 "    set @CodProducto = '' " & vbCrLf & _
                 "    set @PrecioProducto = 0.00 " & vbCrLf & _
                 "    set @CodProducto = (select codigoproductopvc from opciones) " & vbCrLf & _
                 "    if rtrim(ltrim(@CodProducto)) <> '' " & vbCrLf & _
                 "    begin " & vbCrLf & _
                 "      Set @PrecioProducto = (Select Precio From PreciosEspeciales Where Cliente = " & sCP & " And SubCliente = " & sSC & " And CodigoProducto = @CodProducto)" & vbCrLf & _
                 "      if @PrecioProducto is null " & vbCrLf & _
                 "      begin " & vbCrLf & _
                 "        Set @PrecioProducto = (Select Precio From Productos Where Codigo = @CodProducto) " & vbCrLf & _
                 "      end " & vbCrLf & _
                 "      update Productos set existencia = existencia - 1 where Codigo = @CodProducto " & vbCrLf
      'If SC = "0" Then 'Es SOLO cliente:
      'sComando = sComando & _
                 "      update clientes set deuda = deuda + @PrecioProducto where codigo = " & sCP & " " & vbCrLf & _
                 "      update clientes set saldo = deuda - pagos where codigo = " & sCP & " " & vbCrLf
      'Else             'Es sub-cliente:
      'sComando = sComando & _
                 "      update subclientes set deuda = deuda + @PrecioProducto where cliente = " & sCP & " and id = " & sSC & " " & vbCrLf & _
                 "      update subclientes set saldo = deuda - pagos           where cliente = " & sCP & " and id = " & sSC & " " & vbCrLf
      'End If
      'sComando = sComando & _
                 "      insert into [EventosC5] (procesado,idtabla,tabla) values ('N',@IDPersona,'" & sTablaNueva & "')" & vbCrLf & _
                 "    end " & vbCrLf & _
                 "  end " & vbCrLf & _
                 "END" & vbCrLf
                 
      'Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
      'Modulo.DBComandoSQL.CommandText = sComando
      'Modulo.DBComandoSQL.Execute
      ''sComando = "Exec CrearTrigger '" & sTablaNueva & "','" & lCodigoCliente & "','0'"
      sComando = "Exec CrearTrigger '" & sTablaNueva & "','" & Val(Label6.Caption) & "','" & Val(Label7.Caption) & "'"
      Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
      Modulo.DBComandoSQL.CommandText = sComando
      If lTablaVacia = False Then Modulo.DBComandoSQL.Execute

    
      If Err.Number <> 0 Then
        MsgBox "Error en TRIGGER: " & Err.Description, vbCritical, "Información"
      Else
        '--Agregar Tabla en control del cliente:
        sComando = "INSERT INTO Personas (cliente,subcliente,tabla,creacion) VALUES (" & sCP & "," & sSC & ",'" & sTablaNueva & "','" & Format(Now, "yyyyMMdd HH:mm:ss") & "')"
        Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
        Modulo.DBComandoSQL.CommandText = sComando
        ''Clipboard.Clear
        ''Clipboard.SetText sComando
        If lTablaVacia = False Then Modulo.DBComandoSQL.Execute
    
        If Err.Number <> 0 Then
          MsgBox "Error: " & Err.Description, vbCritical, "Información"
        Else
        
          Load fMensaje2
          With fMensaje2
            .Caption = "Copiando los Datos de cada Persona, Espere..."
            .ProgressBar1.Min = 0
            .ProgressBar1.Max = FG1.Rows
            .ProgressBar1.value = 0
          End With
          fMensaje2.Show
          DoEvents
          If lTablaVacia = False Then
          '-- Copiar los registros(filas) de cada persona:
          Adodc3.ConnectionString = Modulo.DBConexionSQL.ConnectionString
          Adodc3.RecordSource = "select * from [" & sTablaNueva & "]"
          Adodc3.Refresh
          
          
          For Fila = 1 To FG1.Rows - 1
          
            fMensaje2.Caption = "Copiando Registro " & CStr(Fila) & " / " & CStr(FG1.Rows)
            fMensaje2.Show
            
            fMensaje2.ProgressBar1.value = fMensaje2.ProgressBar1.value + 1
            
            DoEvents
          
            Adodc3.Recordset.AddNew
            
            sCI = ""
                      
            For i = 0 To FG1.Cols - 1
              
              s = FG1.TextMatrix(0, i) 'El Campo segun Cabecera de FlexGrid
              
              Adodc3.Recordset.Fields(s).value = Trim(FG1.TextMatrix(Fila, i))
              
              If s = "CEDULA" Then sCI = Trim(FG1.TextMatrix(Fila, i))
              
            Next i
          
            '-Asignar los campos predeterminados ocultos
            
            
            Adodc3.Recordset.Fields("FOTO").value = Trim(IIf(sCI = "", "", Trim(DepurarStr(DepurarStr(DepurarStr(sCI, ","), "."), " "))))
            Adodc3.Recordset.Fields("TIENE_FOTO").value = "N"
            Adodc3.Recordset.Fields("MARCA").value = ""
            
            If Not Modulo.EXISTE_CAMPO_EN_FLEXGRID(FG1, "FECHA") Then
              Adodc3.Recordset.Fields("FECHA").value = Null
            End If
              
            If Not Modulo.EXISTE_CAMPO_EN_FLEXGRID(FG1, "CONTADOR") Then
              Adodc3.Recordset.Fields("CONTADOR").value = 0
            End If
            
            Adodc3.Recordset.Fields("CREACION").value = Date
            
            Adodc3.Recordset.Update
            
          Next Fila
          
          Adodc3.Recordset.Close
          Adodc3.ConnectionString = ""
          
          fMensaje2.Caption = "Auditando FOTOS, Espere..."
          DoEvents
          
          Modulo.Auditar_Fotos sTablaNueva
         
          Unload fMensaje2
         End If
         '-----------------------------------------------------------
    '- CREAR LAS CARPETAS Y COPIAR LAS PLANTILLAS
    '-----------------------------------------------------------
    
    'Load fMensaje
    'fMensaje.Label1.Caption = "Copiando Carpetas/Plantillas del Cliente " & eCliente.Text & " , Espere..."
    'fMensaje.Show
    'DoEvents

    'Dim sOri As String
    'Dim sDes As String
    
    'Dim s2 As String
    
    'sOri = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
    'sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")

    
    'If sOri <> "" And sDes <> "" Then
    '   sCopiarCarpetas
    'Else
    '   MsgBox "No se puede determinar el destino a copiar. La carpeta del cliente no se copiará. Vaya a opciones y configure la carpeta de destino", vbExclamation
    'End If
     
    If sOri <> "" And sDes <> "" And 1 = 2 Then
    
      '-- Nombre del Cliente:
      s = sDes & "\" & eCliente.Text
      
      MkDir (s)
      
        '-- Subcarpetas: Carnets - Fotos - Imagenes
        
        s = sDes & "\" & eCliente.Text & "\" & "CARNET"
        MkDir (s)
        
        s = sDes & "\" & eCliente.Text & "\" & "FOTOS"
        MkDir (s)
        
        s = sDes & "\" & eCliente.Text & "\" & "IMAGENES"
        MkDir (s)
        
       ''' '-- Copiar archivos (RAIZ):
        s = sOri & "\FORMATO DE ENTREGA DE CARNETS _nombre cliente.xls"
        s2 = sDes & "\" & eCliente.Text & "\FORMATO DE ENTREGA DE CARNETS " & eCliente.Text & ".xls"
        FileCopy s, s2
        
        's = sOri & "\LISTADO NOMBRE CLIENTE OFICINA.xls"
        's2 = sDes & "\" & tnom.Text & "\LISTADO " & tnom.Text & ".xls"
        'FileCopy s, s2
        
        '-- Copiar archivos (CARNET):
        s = sOri & "\CARNET\BASE CARNET NOMBRE CLIENTE.car"
        s2 = sDes & "\" & eCliente.Text & "\CARNET\BASE CARNET " & eCliente.Text & ".car"
        FileCopy s, s2
        
        s = sOri & "\CARNET\BASE DATOS NOMBRE CLIENTE.mdb"
        s2 = sDes & "\" & eCliente.Text & "\CARNET\BASE DATOS " & eCliente.Text & "_" & Year(Now) & ".mdb"
        FileCopy s, s2
        
        '-- Copiar archivos (IMAGENES):
        ''s = sOri & "\IMAGENES\Autocopia_de_seguridad_deBASE nombre cliente.cdr"
        ''s2 = sDes & "\" & tnom.Text & "\IMAGENES\Autocopia_de_seguridad_de " & tnom.Text & ".cdr"
        ''FileCopy s, s2
        
        s = sOri & "\IMAGENES\BASE nombre cliente.cdr"
        s2 = sDes & "\" & eCliente.Text & "\IMAGENES\BASE " & eCliente.Text & ".cdr"
        FileCopy s, s2
        sCrearArchivoExcel sTablaNueva, eCliente.Text
    End If
    Unload fMensaje
        
          
         Load fFD  'sTablaNueva & "','" & lCodigoCliente
         '''fFD.cCP.Text = lCodigoCliente & " : " & Trim(UCase(eCliente.Text))
         If cEsSubCliente.value = 1 Then
            fFD.cCP.Text = cCP.Text  '''Label6.Caption & " : " & Trim(UCase(eCliente.Text))
            fFD.cSC.Text = Label7.Caption & " : " & Trim(UCase(eCliente.Text))
         Else
            fFD.cCP.Text = lCodigoCliente & " : " & Trim(UCase(eCliente.Text))
            fFD.sCargarDatosCliente
         End If
         fFD.Show vbModal
         Unload fMensaje

          If lTablaVacia = False Then
             MsgBox "Tabla " & sTablaNueva & " y sus carpetas creadas Exitosamente...", vbInformation, "Información"
          Else
             MsgBox "Se copió el contenido de la carpeta " & Dir1 & Chr(13) _
             & " pero no se crearon las tablas porque no se puede determinar la estructura de las mismas. " _
             & Chr(13) & " Para crearlas vaya al menú principal y seleccione la opcion CREAR TABLAS PERSONAS", vbInformation
          End If
          lTablaVacia = True
          Label6.Caption = ""
          Label7.Caption = ""
          Label9.Caption = "-"
          
          FG1.Clear
          FG1.Rows = 2
          FG1.Cols = 1
          
          Adodc1.Recordset.Close
          
          List2.Clear
          List3.Clear
          Unload fMensaje
        End If
      End If
    End If
  End If
  
End Sub

Private Sub bFijarIO_Click()
  Dim s As String, lcod As Long
  Dim s1 As String
  
  If eCliente.Enabled Then
    s = Trim(eCliente.Text)
    
    If Len(s) < MAXLONGCLIENTE Then
      s = Mid(s, 1, MAXLONGCLIENTE)
      eCliente.Text = s
    End If
    
    If Modulo.DBExiste("clientes", "nombre", s) Then
      lcod = CLng(Modulo.DBValorStr("clientes", "nombre", s, "codigo"))
      s1 = Zeros(lcod, 6)
      MsgBox "Ya Existe un Cliente con el Nombre [" & s & "] " & vbCrLf & _
             "tiene Código [" & s1 & "], Revise...", vbCritical, "Información"
             
      If MsgBox("¿Desea Continuar de todas maneras?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
        Label6.Caption = s1
        bCrearCliente.Enabled = False
        eCliente.SetFocus
        lCodigoCliente = s1
        Exit Sub
      End If
      
      'Label6.Caption = s1
    
    Else
      bCrearCliente.Enabled = True
    End If
    
  End If
  
  eCliente.Enabled = Not eCliente.Enabled
  
  If eCliente.Enabled And cEsSubCliente.value = vbUnchecked Then
    Label7.Caption = ""
  End If
  
End Sub

'Private Sub bIzq_Click()
'  Dim c As Integer, i As Integer
'  Dim old1 As String, old2 As String, old3 As String, e As Boolean
'
'  i = 0
'  e = False
'  Do While i < FG2.Cols And Not e
'    FG2.Row = 1
'    FG2.Col = i
'    If FG2.CellBackColor = vbRed Then e = True Else i = i + 1
'  Loop
'
'  If Not e Then Exit Sub
'
'  'Ponerla en Blanco:
'
'  FG2.Row = 0
'  FG2.CellBackColor = vbWhite
'
'  FG2.Col = i
'  FG2.Row = 1
'  FG2.CellBackColor = vbWhite
'
'  FG2.Row = 2
'  FG2.CellBackColor = vbWhite
'
'  c = i
'  'c = FG2.Col
'  If c > 0 Then
'
'    FG2.Col = c
'    FG2.Row = 1
'
'    'lAtributoColor = FG2.CellBackColor
'
'    old1 = FG2.TextMatrix(0, c - 1)
'    old2 = FG2.TextMatrix(1, c - 1)
'    old3 = FG2.TextMatrix(2, c - 1)
'
'    FG2.Row = 0: FG2.Col = c - 1:  FG2.CellBackColor = vbRed 'lAtributoColor
'    FG2.TextMatrix(0, c - 1) = FG2.TextMatrix(0, c)
'
'    FG2.Row = 1: FG2.Col = c - 1:  FG2.CellBackColor = vbRed 'lAtributoColor
'    FG2.TextMatrix(1, c - 1) = FG2.TextMatrix(1, c)
'
'    FG2.Row = 2: FG2.Col = c - 1:  FG2.CellBackColor = vbRed
'    FG2.TextMatrix(2, c - 1) = FG2.TextMatrix(2, c)
'
'
'    FG2.Row = 0: FG2.Col = c:  FG2.CellBackColor = vbWhite
'    FG2.TextMatrix(0, c) = old1
'    FG2.Row = 1: FG2.Col = c:  FG2.CellBackColor = vbWhite
'    FG2.TextMatrix(1, c) = old2
'    FG2.Row = 2: FG2.Col = c:  FG2.CellBackColor = vbWhite
'    FG2.TextMatrix(2, c) = old3
'
'    FG2.Col = FG2.Col - 1
'
'    Modulo.Ajustar_Columnas_FLEXGRID FG2
'  End If
'End Sub


Private Sub bIzq_Click()
  Dim c As Integer, i As Integer
  Dim old1 As String, old2 As String, old3 As String, e As Boolean
  Dim aOLD() As String
  
  i = 0
  e = False
  Do While i < FG1.Cols And Not e
    FG1.Row = 1
    FG1.Col = i
    If FG1.CellBackColor = vbRed Then e = True Else i = i + 1
  Loop
  
  If Not e Then Exit Sub
 
  c = i
  
  'Ponerla en Blanco:
  For i = 0 To FG1.Rows - 1
    FG1.Row = i
    FG1.CellBackColor = vbWhite
  Next i
  
  'c = FG2.Col
  If c > 0 Then
    
    FG1.Col = c
    FG1.Row = 1
    
    
    ReDim aOLD(FG1.Rows)
    
    
    'lAtributoColor = FG2.CellBackColor
    For i = 0 To FG1.Rows - 1
      aOLD(i) = FG1.TextMatrix(i, c - 1)
    Next i
  
    'old1 = FG1.TextMatrix(0, c - 1)
    'old2 = FG1.TextMatrix(1, c - 1)
    'old3 = FG1.TextMatrix(2, c - 1)
    
    For i = 0 To FG1.Rows - 1
      FG1.Row = i: FG1.Col = c - 1: FG1.CellBackColor = vbRed
      FG1.TextMatrix(i, c - 1) = FG1.TextMatrix(i, c)
      
      FG1.Row = i: FG1.Col = c: FG1.CellBackColor = vbWhite
      FG1.TextMatrix(i, c) = aOLD(i)
      
    Next i
    
    
    FG1.Col = FG1.Col - 1
    
    Modulo.Ajustar_Columnas_FLEXGRID_BY_ROWS FG1
    
  End If
End Sub

'Private Sub bDer_Click()
'  Dim c As Integer, i As Integer
'  Dim old1 As String, old2 As String, old3 As String, e As Boolean
'
'  i = 0
'  e = False
'  Do While i < FG2.Cols And Not e
'    FG2.Row = 1
'    FG2.Col = i
'    If FG2.CellBackColor = vbRed Then e = True Else i = i + 1
'  Loop
'
'  If Not e Then Exit Sub
'
'  'Ponerla en Blanco:
'
'  FG2.Row = 0
'  FG2.CellBackColor = vbWhite
'
'  FG2.Col = i
'  FG2.Row = 1
'  FG2.CellBackColor = vbWhite
'
'  FG2.Row = 2
'  FG2.CellBackColor = vbWhite
'
'  c = i
'  'c = FG2.Col
'  If c < FG2.Cols - 1 Then
'
'    FG2.Col = c
'    FG2.Row = 1
'
'    'lAtributoColor = FG2.CellBackColor
'
'    old1 = FG2.TextMatrix(0, c + 1)
'    old2 = FG2.TextMatrix(1, c + 1)
'    old3 = FG2.TextMatrix(2, c + 1)
'
'    FG2.Row = 0: FG2.Col = c + 1:  FG2.CellBackColor = vbRed 'lAtributoColor
'    FG2.TextMatrix(0, c + 1) = FG2.TextMatrix(0, c)
'
'    FG2.Row = 1: FG2.Col = c + 1:  FG2.CellBackColor = vbRed 'lAtributoColor
'    FG2.TextMatrix(1, c + 1) = FG2.TextMatrix(1, c)
'
'    FG2.Row = 2: FG2.Col = c + 1:  FG2.CellBackColor = vbRed
'    FG2.TextMatrix(2, c + 1) = FG2.TextMatrix(2, c)
'
'    FG2.Row = 0: FG2.Col = c:  FG2.CellBackColor = vbWhite
'    FG2.TextMatrix(0, c) = old1
'    FG2.Row = 1: FG2.Col = c:  FG2.CellBackColor = vbWhite
'    FG2.TextMatrix(1, c) = old2
'    FG2.Row = 2: FG2.Col = c:  FG2.CellBackColor = vbWhite
'    FG2.TextMatrix(2, c) = old3
'
'    FG2.Col = FG2.Col + 1
'
'    Modulo.Ajustar_Columnas_FLEXGRID FG2
'
'
'  End If
'
'End Sub


Private Sub bDer_Click()
  Dim c As Integer, i As Integer
  Dim old1 As String, old2 As String, old3 As String, e As Boolean
  Dim aOLD() As String
  
  i = 0
  e = False
  Do While i < FG1.Cols And Not e
    FG1.Row = 1
    FG1.Col = i
    If FG1.CellBackColor = vbRed Then e = True Else i = i + 1
  Loop
  
  If Not e Then Exit Sub
  
  'Ponerla en Blanco:
  
  FG1.Row = 0
  FG1.CellBackColor = vbWhite
  
  FG1.Col = i
  FG1.Row = 1
  FG1.CellBackColor = vbWhite
  
  FG1.Row = 2
  FG1.CellBackColor = vbWhite
    
  c = i
  'c = FG2.Col
  If c < FG1.Cols - 1 Then
  
    FG1.Col = c
    FG1.Row = 1
    
    'lAtributoColor = FG2.CellBackColor
    
    ReDim aOLD(FG1.Rows)
    
    For i = 0 To FG1.Rows - 1
      aOLD(i) = FG1.TextMatrix(i, c + 1)
    Next i
    
    For i = 0 To FG1.Rows - 1
      FG1.Row = i: FG1.Col = c + 1: FG1.CellBackColor = vbRed
      FG1.TextMatrix(i, c + 1) = FG1.TextMatrix(i, c)
      
      FG1.Row = i: FG1.Col = c: FG1.CellBackColor = vbWhite
      FG1.TextMatrix(i, c) = aOLD(i)
    Next i
          
    
    FG1.Col = FG1.Col + 1
    
    Modulo.Ajustar_Columnas_FLEXGRID_BY_ROWS FG1
        
    
  End If

End Sub


'Private Sub ChequearCampos()
'  Dim i As Integer, j As Integer, k As Integer, m As Integer
'  Dim s As String, s1 As String
'  Dim e As Boolean
'  Dim aNuevas() As String
'  Dim KNuevas As Integer
'  Dim l As Long
'
'
'  KNuevas = 0
'
'  For i = 0 To DataGrid1.Columns.Count - 1
'    s = DataGrid1.Columns(i).Caption
'    j = 0
'    e = False
'    Do While j < FG2.Cols And Not e
'      s1 = FG2.TextMatrix(0, j)
'      If s = s1 Then e = True
'      j = j + 1
'    Loop
'
'    If Not e Then
'      ReDim Preserve aNuevas(KNuevas + 1)
'      aNuevas(KNuevas) = s
'      KNuevas = KNuevas + 1
'
'      'FG2.Cols = FG2.Cols + 1
'      'FG2.TextMatrix(0, FG2.Cols - 1) = s
'    End If
'
'  Next i
'
'  If KNuevas = 0 Then Exit Sub
'
'  'buscar su posicion en la 1era.
'  For i = 0 To KNuevas
'    k = Modulo.Nro_Columna_DataGrid(DataGrid1, aNuevas(i))
'    If k > 0 Then
'      s = DataGrid1.Columns(k - 1).Caption 'su lado izquierdo
'      'buscar esta etiqueta en el flexgrid:
'      j = 0
'      m = -1
'      Do While j < FG2.Cols And m = -1
'        If FG2.TextMatrix(0, j) = Modulo.Str_Replace(s, " ", "_") Then m = j Else j = j + 1
'      Loop
'      If m >= 0 Then
'        FG2.Cols = FG2.Cols + 1
'        For j = FG2.Cols - 1 To m + 1 Step -1
'          'old = FG.TextMatrix(0, j)
'          FG2.TextMatrix(0, j) = FG2.TextMatrix(0, j - 1) 'Titulo
'
'          FG2.Row = 2: FG2.Col = j: FG2.CellAlignment = flexAlignCenterCenter
'
'          FG2.Row = 1: FG2.Col = j: FG2.CellAlignment = flexAlignCenterCenter
'          FG2.TextMatrix(1, j) = FG2.TextMatrix(1, j - 1) 'Tipo
'
'          FG2.Row = 2: FG2.Col = j: FG2.CellAlignment = flexAlignCenterCenter
'          FG2.TextMatrix(2, j) = FG2.TextMatrix(2, j - 1) 'Ancho
'
'        Next j
'
'        'por defecto el tipo es "CHAR" y el tamaño es 50
'        FG2.TextMatrix(0, m + 1) = Modulo.Str_Replace(aNuevas(i), " ", "_")
'
'        FlexXY FG2, 1, m + 1, "CHAR", flexAlignCenterCenter
'        FlexXY FG2, 2, m + 1, "50", flexAlignCenterCenter
'
'      End If
'    End If
'  Next i
'
'  Modulo.Ajustar_Columnas_FLEXGRID FG2
'
'End Sub

Private Sub bLeerMDB_Click()
  Dim s As String, s1 As String
  Dim i As Integer, j As Integer
  Dim LaTabla As String
  Dim r As New ADODB.Recordset
  
  Dim KBlancos As Integer
  On Error GoTo falla
  List2.Clear
  
  If File1.ListCount > 0 Then
    ''If File1.ListIndex >= 0 Then
    If ListViewArchivo.ListItems.Count >= 0 Then
      s = ListViewArchivo.SelectedItem.Text
      '''s = File1.List(File1.ListIndex)
      s = Dir1.Path & "\" & s
      If UCase(Right(s, 3)) <> "MDB" Then
        MsgBox "No es un Archivo Access .MDB", vbCritical, "Información"
      Else
      
        Cargar_Campos_Predeterminados
           
        ConexionAccess = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                         "Data Source=" & s & ";" & _
                         "Persist Security Info=False"
                 
        Tablas_De_BD_Access List1, ConexionAccess
                         
        If List1.ListCount > 1 Then
          MsgBox "Se ha detectado que existe más de una Tabla de Datos en el .MDB", vbCritical, "Información"
          Load fAbrirTablaMDB
          
          fAbrirTablaMDB.List1.Clear
          For i = 0 To List1.ListCount - 1
            fAbrirTablaMDB.List1.AddItem List1.List(i)
          Next i
          fAbrirTablaMDB.List1.ListIndex = 0
          Modulo.fModalResult = Modulo.fModalResultCANCEL
          Modulo.vTemporal1 = ""
          fAbrirTablaMDB.Show vbModal
          
          If Modulo.fModalResult = Modulo.fModalResultOK Then
            LaTabla = Modulo.vTemporal1
          End If
          
        Else
          If List1.ListCount = 1 Then
            LaTabla = List1.List(0)
            lTablaVacia = False
          Else
            MsgBox "La Base de Datos No Contiene DATOS de Personas...", vbCritical, "Información"
            lTablaVacia = True
            Exit Sub
          End If
        End If
                                 
        Adodc1.ConnectionString = ConexionAccess
        Adodc1.RecordSource = "SELECT * FROM [" & LaTabla & "]"
        Adodc1.Refresh
        
        Label9.Caption = Zeros(Adodc1.Recordset.RecordCount, 6)
        
        
        'Cargar datos en el FlexGrid 1:
        FG1.Clear
        FG1.Rows = 2
        FG1.Cols = 1 'Adodc1.Recordset.Fields.Count
        
        'Carga la Cabecera: Los Titulos de las Columnas.
        
        If Adodc1.Recordset.RecordCount > 0 Then
           lTablaVacia = False
        Else
           lTablaVacia = True
        End If
        For i = 0 To Adodc1.Recordset.Fields.Count - 1
          's = Modulo.DepurarStr(UCase(Adodc1.Recordset.Fields(i).Name), " ")
          s = UCase(Adodc1.Recordset.Fields(i).Name)
          
          'If s = "FECHA" Or s = "CONTADOR" Then
          '  If FG1.TextMatrix(0, 0) <> "" Then FG1.Cols = FG1.Cols + 1
          '  FG1.TextMatrix(0, FG1.Cols - 1) = s
          'End If
          If s <> "FECHA" And s <> "CONTADOR" Then
             If Not EsCampoOcul(s) And (s <> "ID") Then
               If FG1.TextMatrix(0, 0) <> "" Then FG1.Cols = FG1.Cols + 1
               FG1.TextMatrix(0, FG1.Cols - 1) = s
             End If
          Else
             If FG1.TextMatrix(0, 0) <> "" Then FG1.Cols = FG1.Cols + 1
             FG1.TextMatrix(0, FG1.Cols - 1) = s
          End If
        Next i
        
        'Carga las filas por cada una de los titulos de columnas:
        
        i = 1
        Do While Not Adodc1.Recordset.EOF
          For j = 0 To FG1.Cols - 1
            s = FG1.TextMatrix(0, j)
            If IsNull(Adodc1.Recordset.Fields(s).value) Then
              FG1.TextMatrix(i, j) = ""
            Else
              FG1.Row = i: FG1.Col = j: FG1.CellAlignment = flexAlignLeftCenter
              FG1.TextMatrix(i, j) = Trim(Adodc1.Recordset.Fields(s).value)
            End If
          Next j
          Adodc1.Recordset.MoveNext
          If Not Adodc1.Recordset.EOF Then
            FG1.Rows = FG1.Rows + 1
            i = i + 1
          End If
        Loop
        
        '-- Depurar el caracter extraño en la columna CEDULA
        
        Dim la_col_ci As Integer
        
        la_col_ci = Modulo.Nro_Columna_FlexGrid(FG1, "CEDULA")
        
        If la_col_ci >= 0 Then
          
          For i = 1 To FG1.Rows - 1
        
            FG1.TextMatrix(i, la_col_ci) = Modulo.StrDepurarChar9(FG1.TextMatrix(i, la_col_ci)) 'DepurarValorCEDULA(FG1.TextMatrix(i, la_col_ci))
            
          Next i
          
        End If
        
        '-- Depurar si hay lineas "en blanco":
        Dim e As Boolean 'Existe valor
        e = False
        For i = 1 To FG1.Rows - 1
          e = False
          j = 0
          Do While j < FG1.Cols And Not e
            If Trim(FG1.TextMatrix(i, j)) <> "" Then e = True Else j = j + 1
          Loop
          If Not e Then FG1.TextMatrix(i, 0) = "*"
        Next i
        
        Dim LC As Integer
        
        LC = La_Fila_Existe_Valor_En_Columna(FG1, 0, "*")
        
        Do While LC >= 0
        
          'Existe_Valor_En_Columna(FG1, 0, "*")
          FG1.RemoveItem (LC)
                    
          LC = La_Fila_Existe_Valor_En_Columna(FG1, 0, "*")
        Loop
                      
        
        Modulo.DepurarTitutlosFlexGrid FG1
        
        Modulo.Ajustar_Columnas_FLEXGRID_BY_ROWS FG1
        
'        For i = 0 To DataGrid1.Columns.Count - 1
'          DataGrid1.Columns(i).Caption = UCase(DataGrid1.Columns(i).Caption)
'        Next i
'
'        'datagrid1.RowHeight =
'
'        'ChequearCampos
'
'
'
'        KBlancos = 0
'        For i = 0 To FG2.Cols - 1
'          If Trim(FG2.TextMatrix(0, i)) = "" Then
'            KBlancos = KBlancos + 1
'          End If
'        Next i
'        If KBlancos > 0 Then FG2.Cols = FG2.Cols - KBlancos
'
'
'        For i = 0 To DataGrid1.Columns.Count - 1
'          DataGrid1.Columns(i).Caption = "-" & Zeros(i + 1, 2) & "- " & UCase(DataGrid1.Columns(i).Caption)
'        Next i
'
'        'Realiza la correspondencia entre las columnas por el Nº
'        For i = 0 To DataGrid1.Columns.Count - 1
'          s1 = Trim(Mid(DataGrid1.Columns(i).Caption, 1, 4))
'          s = Trim(Mid(DataGrid1.Columns(i).Caption, 5))
'          j = Modulo.Nro_Columna_FlexGrid(FG2, s)
'          If j >= 0 Then Modulo.FlexXY FG2, 3, j, s1, flexAlignCenterCenter
'        Next i
            
          
          
        

        
        
        
      End If
    Else
      MsgBox "Debe Seleccionar el archivo con extensión [.MDB]", vbCritical, "Información"
    End If
  End If
falla:
  If Err.Number <> 0 Then
     MsgBox Err.Number & "::" & Err.Description, vbCritical
  End If
End Sub

Private Sub bNuevaCol_Click()
  Dim s As String
  s = ""
  s = InputBox("Nombre de la Nueva Columna:", "Indique el Nombre")
  If s <> "" Then
    FG1.Cols = FG1.Cols + 1
    FG1.TextMatrix(0, FG1.Cols - 1) = UCase(s)
  End If
End Sub



Private Sub cCP_Click()
  Label6.Caption = Mid(cCP.Text, 1, 6)
End Sub

Private Sub cEsSubCliente_Click()
  If cEsSubCliente.value = vbChecked Then
    Label3.Visible = True
    cCP.Visible = True
    bBuscarCP.Visible = True
    
    Cargar_Clientes
    
    If cCP.ListCount > 0 Then cCP.ListIndex = 0
    
  Else
    Label3.Visible = False
    cCP.Visible = False
    bBuscarCP.Visible = False
  End If
End Sub

Public Sub sCopiarCarpetas(ByVal Carpeta As DirListBox, ByVal Archivo As FileListBox, ByVal SubDirectorio As String, ByVal Dir11 As DirListBox, ByVal File22 As FileListBox)
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim sOri As String
   Dim sDes As String
   Dim lCarpeta As String
   Dim ArchivoaBorrar As String
  ' Dim Dir11 As DirListBox
  ' Dim File22 As FileListBox
   'Set Dir11 = Controls.Add("VB.DirListbox", "Dir11")
   'If Err.Number <> 0 Then MsgBox "error"
   'Set File22 = Controls.Add("VB.filelistbox", "File22")
   'With Dir11
    '    .Visible = True
    '    .Top = 100
    '    .Left = 100
    '    .Height = 255
    'End With
   
    Dim s2 As String
    On Error Resume Next
    sOri = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
    sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
    If fCargaClientes.cEsSubCliente.value = vbChecked Then
       s = sDes & "\" & Mid(fCargaClientes.cCP.Text, 10, Len(fCargaClientes.cCP.Text) - 9) & "\" & eCliente.Text
       If SubDirectorio <> "" Then
          s = s & "\" & SubDirectorio
       Else
         
         MkDir (s)
      
        '-- Subcarpetas: Carnets - Fotos - Imagenes
        
         s = sDes & "\" & Mid(fCargaClientes.cCP.Text, 10, Len(fCargaClientes.cCP.Text) - 9) & "\" & eCliente.Text & "\" & "CARNET"
         MkDir (s)
        '
         s = sDes & "\" & Mid(fCargaClientes.cCP.Text, 10, Len(fCargaClientes.cCP.Text) - 9) & "\" & eCliente.Text & "\" & "FOTOS"
         MkDir (s)
        
         s = sDes & "\" & Mid(fCargaClientes.cCP.Text, 10, Len(fCargaClientes.cCP.Text) - 9) & "\" & eCliente.Text & "\" & "IMAGENES"
         MkDir (s)
         s = sDes & "\" & Mid(fCargaClientes.cCP.Text, 10, Len(fCargaClientes.cCP.Text) - 9) & "\" & eCliente.Text
       End If
    Else
       s = sDes & "\" & eCliente.Text
      If SubDirectorio <> "" Then
         s = s & "\" & SubDirectorio
      Else
         
         MkDir (s)
      
        '-- Subcarpetas: Carnets - Fotos - Imagenes
        
         s = sDes & "\" & eCliente.Text & "\" & "CARNET"
         MkDir (s)
        '
         s = sDes & "\" & eCliente.Text & "\" & "FOTOS"
         MkDir (s)
        
         s = sDes & "\" & eCliente.Text & "\" & "IMAGENES"
         MkDir (s)
         s = sDes & "\" & eCliente.Text
      End If
       
    End If
      
      
   Load fMensaje2
   
   ''PARA COPIAR LOS ARCHIVOS QUE SE ENCUENTRAN EN LA RAIZ DE LA CARPTE PRINCIPAL
      'File2.Path = Dir1
      Archivo.Path = Carpeta
      With fMensaje2
         .ProgressBar1.Min = 0
         .ProgressBar1.value = 0
         .ProgressBar1.Max = Archivo.ListCount
      End With
      fMensaje2.Show
      DoEvents
      'lCarpeta = Trim(sDeteCarpeta(Dir1))
       
      For j = 0 To Archivo.ListCount - 1
           fMensaje2.ProgressBar1.value = fMensaje2.ProgressBar1.value + 1
           fMensaje2.Label1.Caption = "Copiando Archivo " & Archivo.List(j) & " en " & Carpeta & " ..."
           fMensaje2.Label1.Refresh
           CopyFile Carpeta & "\" & Archivo.List(j), s & "\" & Archivo.List(j), True
           If Err.Number <> 0 Then
              fMensaje2.Label1.Caption = Err.Number & "::" & Err.Description & " Archivo " & sDes & "\" & eCliente.Text & "\" & lCarpeta & "\" & Archivo.List(j)
              fMensaje2.Label1.Refresh
           End If
      Next j
   
   '***************************************************************************************
   
   
   With frmCopiarCarpetas
      With fMensaje2
         .ProgressBar1.Min = 0
         .ProgressBar1.value = 0
         .ProgressBar1.Max = Carpeta.ListCount
      End With
      fMensaje2.Show
      DoEvents
       fMensaje2.ProgressBar1.value = fMensaje2.ProgressBar1.value + 1
       fMensaje2.Label1.Caption = "Copiando carpeta " & .txtCarpetaCarnet.Text & " ..."
       fMensaje2.Label1.Refresh
          Archivo.Path = .txtCarpetaCarnet
          i = 0
          For j = 0 To Archivo.ListCount - 1
              'fMensaje2.ProgressBar1.value = fMensaje2.ProgressBar1.value + 1
              fMensaje2.Label1.Caption = "Verificando Archivo " & Archivo.List(j) & " en " & Carpeta & " ..."
              fMensaje2.Label1.Refresh
              'CopyFile Carpeta & "\" & Archivo.List(j), s & "\" & Archivo.List(j), True
              If UCase(Mid(Archivo.List(j), Len(Archivo.List(j)) - 3, 4)) = ".CAR" Then
            
                i = i + 1
              End If
              
              If Err.Number <> 0 Then
                 fMensaje2.Label1.Caption = Err.Number & "::" & Err.Description & " Archivo " & sDes & "\" & eCliente.Text & "\" & lCarpeta & "\" & Archivo.List(j)
                 fMensaje2.Label1.Refresh
              End If
          Next j
          If i = 1 Then
             For j = 0 To Archivo.ListCount - 1
                If UCase(Mid(Archivo.List(j), Len(Archivo.List(j)) - 3, 4)) = ".CAR" Then
                   CopyFile .txtCarpetaCarnet & "\" & Archivo.List(j), s & "\CARNET\BASE CARNET " & fCargaClientes.eCliente.Text & ".CAR", True
                   If UCase(Archivo.List(j)) <> UCase("BASE CARNET " & fCargaClientes.eCliente.Text & ".CAR") Then
                      ArchivoaBorrar = Archivo.List(j)
                   Else
                      ArchivoaBorrar = ""
                   End If
                   Exit For
                End If
             Next j
          End If
          sCopiarCarpeta1 .txtCarpetaCarnet.Text, s & "\" & "CARNET"
          If ArchivoaBorrar <> "" Then Kill s & "\CARNET\" & ArchivoaBorrar
       fMensaje2.ProgressBar1.value = fMensaje2.ProgressBar1.value + 1
       fMensaje2.Label1.Caption = "Copiando carpeta " & .txtCarpetaFotos.Text & " ..."
       fMensaje2.Label1.Refresh
          sCopiarCarpeta1 .txtCarpetaFotos.Text, s & "\" & "FOTOS"
       
       fMensaje2.ProgressBar1.value = fMensaje2.ProgressBar1.value + 1
       fMensaje2.Label1.Caption = "Copiando carpeta " & .txtCarpetaImagenes.Text & " ..."
       fMensaje2.Label1.Refresh
          sCopiarCarpeta1 .txtCarpetaImagenes.Text, s & "\" & "IMAGENES"
          MkDir s & "\HISTORICO"
          MkDir s & "\HISTORICO\" & .txtNuevaCarpeta.Text
       For i = 0 To .ListHistorico.ListCount - 1
          fMensaje2.ProgressBar1.value = fMensaje2.ProgressBar1.value + 1
          fMensaje2.Label1.Caption = "Copiando carpeta " & .ListHistorico.List(i) & " ..."
          fMensaje2.Label1.Refresh
             lCarpeta = Trim(sDeteCarpeta(.ListHistorico.List(i)))
             sCopiarCarpeta1 .ListHistorico.List(i), s & "\HISTORICO\" & .txtNuevaCarpeta & "\" & lCarpeta
       Next i
       
   End With
   'For i = 0 To Carpeta.ListCount - 1
   '   'File2.Path = Dir1.List(i)
   '   Archivo.Path = Carpeta.List(i)
   '   With fMensaje2
   '      .ProgressBar1.Min = 0
   '      .ProgressBar1.value = 0
   '      .ProgressBar1.Max = Carpeta.ListCount
   '   End With
   '   fMensaje2.Show
   '   DoEvents
   '   lCarpeta = Trim(sDeteCarpeta(Carpeta.List(i)))
   '
   '       If Dir(s & "\" & lCarpeta, vbDirectory) = "" Then
   '          'No Existe la carpeta
   '          MkDir s & "\" & lCarpeta
   '       End If
   '
   '    ''For j = 0 To Archivo.ListCount - 1
   '        fMensaje2.ProgressBar1.value = fMensaje2.ProgressBar1.value + 1
   '        fMensaje2.Label1.Caption = "Copiando carpeta " & Carpeta.List(i) & " ..."
   '        fMensaje2.Label1.Refresh
   '        Select Case Carpeta.List(i)
   '           Case lDirCarnet
   '              'CopyFile Carpeta.List(i) & "\" & Archivo.List(j), s & "\" & "CARNET" & "\" & Archivo.List(j), True
   '              sCopiarCarpeta1 Carpeta.List(i), s & "\" & "CARNET"
   '           Case lDirFotos
   '              '''CopyFile Dir1.List(i) & "\" & File2.List(j), sDes & "\" & eCliente.Text & "\" & "FOTOS" & "\" & File2.List(j), True
   '              'CopyFile Carpeta.List(i) & "\" & Archivo.List(j), s & "\" & "FOTOS" & "\" & Archivo.List(j), True
   '              sCopiarCarpeta1 Carpeta.List(i), s & "\" & "FOTOS"
   '           Case lDirImagenes
   '              'CopyFile Carpeta.List(i) & "\" & Archivo.List(j), s & "\" & "IMAGENES" & "\" & Archivo.List(j), True
   '              ''''CopyFile Dir1.List(i) & "\" & File2.List(j), sDes & "\" & eCliente.Text & "\" & "IMAGENES" & "\" & File2.List(j), True
   '              sCopiarCarpeta1 Carpeta.List(i), s & "\" & "IMAGENES"
   '           Case Else
   '              'CopyFile Dir1.List(i) & "\" & File2.List(j), sDes & "\" & eCliente.Text & "\" & lCarpeta & "\" & File2.List(j), True
   '              sCopiarCarpeta1 Carpeta.List(i), s & "\" & lCarpeta
   '        End Select
   '        If Err.Number <> 0 Then
   '           fMensaje2.Label1.Caption = Err.Number & "::" & Err.Description & " Archivo " & sDes & "\" & eCliente.Text & "\" & lCarpeta
   '           fMensaje2.Label1.Refresh
   '        End If
   '    ''Next j
   'Next i
   Unload fMensaje2
   On Error GoTo falla
   Err.Clear
falla:
   If Err.Number <> 0 Then
      MsgBox Err.Number & "::" & Err.Description, vbCritical
      Err.Clear
   End If
End Sub
Private Sub sCopiarCarpeta1(argOrigen As String, argDestino As String)
   Dim FS As Object

   Set FS = CreateObject("Scripting.FileSystemObject")
   FS.CopyFolder argOrigen, argDestino

End Sub

Private Function ExisteCarpeta(argDir As String) As Boolean
 
End Function


Private Function sDeteCarpeta(argDir As String) As String
   Dim aux As String
   Dim i As Integer
   aux = ""
   For i = Len(argDir) To 1 Step -1
      If Mid(argDir, i, 1) <> "\" Then
       aux = Mid(argDir, i, 1) & aux
      Else
        Exit For
      End If
   Next i
   sDeteCarpeta = aux
End Function


Private Sub Command1_Click()
'sCopiarCarpetas
End Sub

Private Sub cmdCopiarFotosEnSitio_Click()
    Dim i As Integer
   Dim s As String
   On Error GoTo falla
   If eCliente.Text = "" Then
      MsgBox "Debe seleccionar el cliente y el archivo .MDB para continuar", vbExclamation
      Exit Sub
   End If
   
   If FG1.Rows <= 3 And FG1.Cols <= 2 Then
      MsgBox "Debe seleccionar el cliente y el archivo .MDB para continuar", vbExclamation
      Exit Sub
   End If
   Dialog1.ShowOpen
   
   File3.Path = Mid(Dialog1.FileName, 1, Len(Dialog1.FileName) - Len(Dialog1.FileTitle))
   
   '- CREAR LAS CARPETAS Y COPIAR LAS PLANTILLAS
    '-----------------------------------------------------------
    
   

    Dim sOri As String
    Dim sDes As String
      
    sOri = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
    sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")

    
    
    
    If sOri <> "" And sDes <> "" Then
    
      '-- Nombre del Cliente:
      s = sDes & "\" & eCliente.Text
      
     'MkDir (s)
      
        '-- Subcarpetas: Carnets - Fotos - Imagenes
        
        's = sDes & "\" & eCliente.Text & "\" & "CARNET"
        'MkDir (s)
        
        's = sDes & "\" & eCliente.Text & "\" & "FOTOS"
        'MkDir (s)
        
        's = sDes & "\" & eCliente.Text & "\" & "IMAGENES"
        'MkDir (s)
        
       ''' '-- Copiar archivos (RAIZ):
        ''s = sOri & "\FORMATO DE ENTREGA DE CARNETS _nombre cliente.xls"
        ''s2 = sDes & "\" & eCliente.Text & "\FORMATO DE ENTREGA DE CARNETS " & eCliente.Text & ".xls"
        ''FileCopy s, s2
        
        's = sOri & "\LISTADO NOMBRE CLIENTE OFICINA.xls"
        's2 = sDes & "\" & tnom.Text & "\LISTADO " & tnom.Text & ".xls"
        'FileCopy s, s2
        
        '-- Copiar archivos (CARNET):
        ''s = sOri & "\CARNET\BASE CARNET NOMBRE CLIENTE.car"
        ''s2 = sDes & "\" & eCliente.Text & "\CARNET\BASE CARNET " & eCliente.Text & ".car"
        ''FileCopy s, s2
        
        ''s = sOri & "\CARNET\BASE DATOS NOMBRE CLIENTE.mdb"
        ''s2 = sDes & "\" & eCliente.Text & "\CARNET\BASE DATOS " & eCliente.Text & "_" & Year(Now) & ".mdb"
        ''FileCopy s, s2
        
        '-- Copiar archivos (IMAGENES):
        ''s = sOri & "\IMAGENES\Autocopia_de_seguridad_deBASE nombre cliente.cdr"
        ''s2 = sDes & "\" & tnom.Text & "\IMAGENES\Autocopia_de_seguridad_de " & tnom.Text & ".cdr"
        ''FileCopy s, s2
        
        ''s = sOri & "\IMAGENES\BASE nombre cliente.cdr"
        ''s2 = sDes & "\" & eCliente.Text & "\IMAGENES\BASE " & eCliente.Text & ".cdr"
        ''FileCopy s, s2
        '
    End If
       Adodc1.Recordset.MoveFirst
       Do While Adodc1.Recordset.EOF = False
          If Trim(Adodc1.Recordset!NROFOTO) <> "DSC00000.JPG" Then
             FileCopy File3.Path & "\" & Trim(Adodc1.Recordset!NROFOTO), sDes & "\" & eCliente.Text & "\" & "FOTOS" & "\" & Trim(Replace(Adodc1.Recordset!cedula, ".", "")) & ".JPG"
          End If
          Adodc1.Recordset.MoveNext
       Loop
       MsgBox "Las Fotos fueron copiadas según su numero de cédula en la carpeta fotos del cliente", vbInformation
       s = ""
       s = Modulo.La_Tabla_Actual_Personas(Label6.Caption, Label7.Caption)
       DBConexionSQL.Execute "ALTER TABLE [" & s & "] DROP COLUMN NROFOTO"
       DBConexionSQL.Execute "ALTER TABLE [H" & s & "] DROP COLUMN NROFOTO"
       Adodc1.Recordset.MoveFirst
       
falla:
   If Err.Number <> 0 Then
   End If
End Sub

Private Sub Command2_Click()
   If eCliente.Text = "" Then
      MsgBox "Debe seleccionar una carpeta a copiar para el cliente", vbExclamation
      Exit Sub
   End If
    Load frmCopiarCarpetas
    frmCopiarCarpetas.lbltitulo.Caption = "LISTA DE CARPETAS PARA """ & eCliente.Text & """"
    frmCopiarCarpetas.ListCarpetas.Path = Dir1.Path
    frmCopiarCarpetas.ListCarpetas.Refresh
    frmCopiarCarpetas.sCargarListCarpetas
    frmCopiarCarpetas.Show
    

End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
  sCargarlistViewArchivo
End Sub

Private Sub sCargarlistViewArchivo()
  Dim Archivo As File
  Dim i As Integer
  Dim lItem As ListItem
  'If argFiltro = "" Or argFiltro = "*.*" Then
     ListViewArchivo.ListItems.Clear
     For i = 0 To File1.ListCount - 1
        Set Archivo = FSO.GetFile(File1.Path & "\" & File1.List(i))
        Set lItem = ListViewArchivo.ListItems.Add(, , File1.List(i))
           lItem.SubItems(1) = Archivo.Size
           lItem.SubItems(2) = Archivo.DateLastModified
     Next i
  'Else
  '   ListViewArchivo.ListItems.Clear
  '   For i = 0 To File1.ListCount - 1
  '
  '      Set Archivo = FSO.GetFile(File1.Path & "\" & File1.List(i))
  '      If Mid(Archivo.Name, Len(Archivo.Name) - 3, 3) = argFiltro Then
  '         Set lItem = ListViewArchivo.ListItems.Add(, , File1.List(i))
  '            lItem.SubItems(1) = Archivo.Size
  '            lItem.SubItems(2) = Archivo.DateLastModified
  '      End If
  '   Next i
  'End If
  Set FSO = Nothing
End Sub



Private Sub Dir1_Click()
  Dim s As String
  If Dir1.ListIndex >= -1 Then
    If eCliente.Enabled Then
      s = Dir1.List(Dir1.ListIndex)
      eCliente.Text = ExtraerArchivo(s)
      RutaAccess = Dir1.List(Dir1.ListIndex)
      'lCliente.Caption = eCliente.Text
      'eCliente.Text = Dir1
    End If
  End If
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

Private Sub FG1_Click()
  Dim i As Integer, j As Integer, k As Integer
  
  If FG1.Row >= 0 Then
    j = FG1.Col
  
    If FG1.CellBackColor = vbRed Then
      FG1.CellBackColor = vbWhite
    Else
    
      For k = 0 To FG1.Rows - 1
        FG1.Row = k
        For i = 0 To FG1.Cols - 1
          FG1.Col = i
          FG1.CellBackColor = vbWhite
        Next i
      Next k
      
      
      For k = 0 To FG1.Rows - 1
        FG1.Row = k
        FG1.Col = j
        FG1.CellBackColor = vbRed
      Next k
            
    
    End If
    'FG2.TextMatrix(FG2.Row, FG2.Col) = " "
  End If



End Sub

Private Sub FG1_DblClick()
  Dim c As Integer, k As Integer
  Dim s As String, s1 As String, e As Boolean
  Dim i As Integer
  
  If FG1.Row >= 0 And FG1.Row < FG1.Rows Then
    If FG1.Col >= 0 And FG1.Col < FG1.Cols Then
      s = FG1.TextMatrix(0, FG1.Col)
      s1 = s
      Load fCamposPersonas
      With fCamposPersonas
        .eCampo.Text = s
        '.cTipos.Clear
        '.cTipos.AddItem "CHAR"
        '.cTipos.AddItem "INTEGER"
        '.cTipos.AddItem "DATETIME"
        '.cTipos.ListIndex = 0
        '.cTipos.ListIndex = Modulo.Buscar_ComboLen(.cTipos, FG2.TextMatrix(1, FG2.Col), 3)
        .eAncho.Text = "50" 'FG2.TextMatrix(2, FG2.Col)
        
        i = -1
        Do While i < List2.ListCount And .eAncho.Text = "50"
          If List2.List(i) = s Then
            .eAncho.Text = List3.List(i)
          Else
            i = i + 1
          End If
        Loop
         
        
        
        'Campos_DataGrid_En_Combo DataGrid1, .Combo1
        '.Combo1.ListIndex = 0
        
        '.Combo1.Clear
        '.Combo1.AddItem FG2.TextMatrix(0, FG2.Col)
      End With
      Modulo.vTemporal1 = ""
      Modulo.vTemporal2 = ""

      fCamposPersonas.Show vbModal
           
      's = Trim(InputBox("Indique Nombre del Campo:", "Confirme", s))
      s = Trim(Modulo.vTemporal1)
      
      If s = s1 Then
        k = POS_LIST(List2, Modulo.vTemporal1)
        If k >= 0 Then
          List3.List(k) = Modulo.vTemporal2
        Else
          List2.AddItem s
          List3.AddItem Modulo.vTemporal2
        End If
        FG1.TextMatrix(0, FG1.Col) = s
      Else
        If s <> s1 And s <> "" Then
          If Modulo.EXISTE_CAMPO_EN_FLEXGRID(FG1, s) Then
            MsgBox "El Titulo [" & s & "] Ya Existe en las Columnas, Revise...", vbCritical, "Información"
          Else
            k = POS_LIST(List2, Modulo.vTemporal1)
          
            If k >= 0 Then
              List3.List(k) = Modulo.vTemporal2
            Else
              List2.AddItem s
              List3.AddItem Modulo.vTemporal2
            End If
          
            FG1.TextMatrix(0, FG1.Col) = s

            'FG1.TextMatrix(1, FG2.Col) = Modulo.vTemporal2
          End If
        End If
      End If
    End If
  End If
   
End Sub

'Private Sub eCed_GotFocus()
'  SeleccionarCampo eCed
'End Sub





Private Sub FG2_Click()
  Dim i As Integer, j As Integer, k As Integer
  
  If FG2.Row = 1 Or FG2.Row = 2 Then
    j = FG2.Col
  
    If FG2.CellBackColor = vbRed Then
      FG2.CellBackColor = vbWhite
    Else
    
      For k = 0 To 2
        FG2.Row = k
        For i = 0 To FG2.Cols - 1
          FG2.Col = i
          FG2.CellBackColor = vbWhite
        Next i
      Next k
      
      FG2.Row = 0
      FG2.Col = j
      FG2.CellBackColor = vbRed
      
      FG2.Row = 1
      FG2.Col = j
      FG2.CellBackColor = vbRed
      
      FG2.Row = 2
      FG2.Col = j
      FG2.CellBackColor = vbRed
      
    
    End If
    'FG2.TextMatrix(FG2.Row, FG2.Col) = " "
  End If
  
  If FG2.Row = 3 Then
    'Combo1.Visible = True
    'Combo1.Left = FG2.CellLeft
    'Combo1.Top = FG2.CellTop
  End If
    
  
End Sub

Private Sub FG2_DblClick()
  Dim c As Integer
  Dim s As String, s1 As String
  If FG2.Row >= 0 And FG2.Row < FG2.Rows Then
    If FG2.Col >= 0 And FG2.Col < FG2.Cols Then
      s = FG2.TextMatrix(0, FG2.Col)
      s1 = s
      Load fCamposPersonas
      With fCamposPersonas
        .eCampo.Text = s
        '.cTipos.Clear
        '.cTipos.AddItem "CHAR"
        '.cTipos.AddItem "INTEGER"
        '.cTipos.AddItem "DATETIME"
        '.cTipos.ListIndex = 0
        '.cTipos.ListIndex = Modulo.Buscar_ComboLen(.cTipos, FG2.TextMatrix(1, FG2.Col), 3)
        .eAncho.Text = FG2.TextMatrix(2, FG2.Col)
        
        'Campos_DataGrid_En_Combo DataGrid1, .Combo1
        '.Combo1.ListIndex = 0
        
        '.Combo1.Clear
        '.Combo1.AddItem FG2.TextMatrix(0, FG2.Col)
      End With
      Modulo.vTemporal1 = ""
      Modulo.vTemporal2 = ""
      Modulo.vTemporal3 = ""
      fCamposPersonas.Show vbModal
           
      's = Trim(InputBox("Indique Nombre del Campo:", "Confirme", s))
      s = Trim(Modulo.vTemporal1)
      If s <> s1 And s <> "" Then
        If Modulo.EXISTE_CAMPO_EN_FLEXGRID(FG2, s) Then
          MsgBox "El Titulo [" & s & "] Ya Existe en las Columnas, Revise...", vbCritical, "Información"
        Else
          FG2.TextMatrix(0, FG2.Col) = s
          FG2.TextMatrix(1, FG2.Col) = Modulo.vTemporal2
          FG2.TextMatrix(2, FG2.Col) = Modulo.vTemporal3
        End If
      End If
    End If
  End If
End Sub

Private Sub File1_PathChange()
  'CargarArchivosDeDisco
  'lRutaActual.Caption = File1.Path
  
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

Private Sub Cargar_Campos_Predeterminados()
  Dim s As String
  's = "ID|CEDULA|NOMBRE|CARGO|VENCE|FOTO|TIENE_FOTO|MARCA|FECHA|CONTADOR|CREACION"
  s = "CEDULA|NOMBRE|CARGO|VENCE|FECHA|CONTADOR"
  FG2.Clear
  FG2.Rows = FGROWS
  FG2.FormatString = s
  '-- CEDULA
  Modulo.FlexXY FG2, 1, 0, "CHAR", flexAlignCenterCenter
  Modulo.FlexXY FG2, 2, 0, "20", flexAlignCenterCenter
  '-- NOMBRE
  Modulo.FlexXY FG2, 1, 1, "CHAR", flexAlignCenterCenter
  Modulo.FlexXY FG2, 2, 1, "50", flexAlignCenterCenter
  '-- CARGO:
  Modulo.FlexXY FG2, 1, 2, "CHAR", flexAlignCenterCenter
  Modulo.FlexXY FG2, 2, 2, "50", flexAlignCenterCenter
  '-- VENCE:
  Modulo.FlexXY FG2, 1, 3, "CHAR", flexAlignCenterCenter
  Modulo.FlexXY FG2, 2, 3, "20", flexAlignCenterCenter
  '-- FOTO:
  Modulo.FlexXY FG2, 1, 4, "DATETIME", flexAlignCenterCenter
  Modulo.FlexXY FG2, 2, 4, "", flexAlignCenterCenter
  '-- TIENE FOTO?:
  Modulo.FlexXY FG2, 1, 5, "INTEGER", flexAlignCenterCenter
  Modulo.FlexXY FG2, 2, 5, "", flexAlignCenterCenter
  '-- TIENE MARCA:
'  Modulo.FlexXY FG2, 1, 6, "CHAR", flexAlignCenterCenter
'  Modulo.FlexXY FG2, 2, 6, "1", flexAlignCenterCenter
'  '-- FECHA:
'  Modulo.FlexXY FG2, 1, 7, "DATE", flexAlignCenterCenter
'  Modulo.FlexXY FG2, 2, 7, "", flexAlignCenterCenter
'  '-- CONTADOR:
'  Modulo.FlexXY FG2, 1, 8, "INT", flexAlignCenterCenter
'  Modulo.FlexXY FG2, 2, 8, "", flexAlignCenterCenter
'  '-- CREACION:
'  Modulo.FlexXY FG2, 1, 9, "DATE", flexAlignCenterCenter
'  Modulo.FlexXY FG2, 2, 9, "", flexAlignCenterCenter
End Sub

Private Sub ListViewArchivo_DblClick()
   bLeerMDB_Click
End Sub

Private Sub Option1_Click()
  If Option1.value Then
    File1.Pattern = "*.*"
  Else
    File1.Pattern = "*.mdb"
  End If
  sCargarlistViewArchivo
End Sub

Private Sub Option2_Click()
  If Option2.value Then
    File1.Pattern = "*.mdb"
  Else
    File1.Pattern = "*.*"
  End If
  sCargarlistViewArchivo
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
    s = Zeros(r.Fields("codigo").value, 6) & " : " & Trim(r.Fields("nombre").value)
    cCP.AddItem s
    r.MoveNext
  Loop
  
  r.Close
End Sub




Private Sub Form_Load()
  Dim UI As String
  lTablaVacia = True
  Label6.Caption = "-"
  Label7.Caption = "-"
  
  Label9.Caption = "-"
  
  List2.Clear
  Label3.Visible = False
  cCP.Visible = False
  bBuscarCP.Visible = False
   
  'lCliente.Visible = False
  
  UI = Mid(App.Path, 1, 2)
    
  eCliente.Text = ""
  
  CargarCamposPre
  
    
  Cargar_Campos_Predeterminados
  
  Dir1.Path = UI
  
  File1.Refresh
  'CargarArchivosDeDisco
  
  RutaAccess = ""
  
  'Combo1.Visible = False
  
  sCargarlistViewArchivo
 
End Sub

Private Sub ListViewARCHIVO_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

With ListViewArchivo
       
        Dim i As Long
        Dim Formato As String
        Dim strData() As String
           
        Dim Columna As Long
           
        ''Call SendMessage(Me.hWnd, WM_SETREDRAW, 0&, 0&)
           
           
        Columna = ColumnHeader.Index - 1
           
        '''''''''''''''''''''''''''''''''''''''''''''
        ' Tipo de dato a ordenar
        ''''''''''''''''''''''''''''''''''''''''''''''
           
        Select Case UCase$(ColumnHeader.Tag)
       
           
        ' Fecha
        '''''''''''''''''''''''''''''''''''''''''''''
        Case "DATE"
           
            Formato = "YYYYMMDDHhNnSs"
           
            ' Ordena alfabéticamente la columna con Fechas _
              ( es la columna que tiene en el tag el valor DATE )
           
            With .ListItems
                If (Columna > 0) Then
                    For i = 1 To .Count
                        With .Item(i).ListSubItems(Columna)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    Formato)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .Item(i)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    Formato)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                End If
            End With
               
            ' Ordena alfabéticamente
               
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
               
            With .ListItems
                If (Columna > 0) Then
                    For i = 1 To .Count
                        With .Item(i).ListSubItems(Columna)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .Item(i)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i
                End If
            End With
               
        ' Datos de numéricos
        '''''''''''''''''''''''''''''''''''''''''''''
        Case "NUMBER"
           
            ' Ordena alfabéticamente la columna con números _
              ( es la columna que tiene en el tag el valor NUMBER )
           
            Formato = String(30, "0") & "." & String(30, "0")
                   
            With .ListItems
                If (Columna > 0) Then
                    For i = 1 To .Count
                        With .Item(i).ListSubItems(Columna)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        Formato)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        Formato))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .Item(i)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        Formato)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        Formato))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                End If
            End With
               
            ' Ordena alfabéticamente
               
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
               
            With .ListItems
                If (Columna > 0) Then
                    For i = 1 To .Count
                        With .Item(i).ListSubItems(Columna)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .Item(i)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i
                End If
            End With
           
        Case Else
                       
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
               
        End Select
       
    End With
       
    ''Call SendMessage(Me.hWnd, WM_SETREDRAW, 1&, 0&)
    ListViewArchivo.Refresh
       
End Sub
  
Private Function InvNumber(ByVal Number As String) As String
    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function





