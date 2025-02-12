VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmInventario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inevntario"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6840
      TabIndex        =   17
      Text            =   "0"
      Top             =   4860
      Width           =   1515
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7680
      TabIndex        =   15
      Text            =   "0"
      Top             =   4500
      Width           =   675
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6840
      TabIndex        =   13
      Text            =   "0"
      Top             =   4140
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   8295
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6180
         TabIndex        =   11
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4920
         TabIndex        =   9
         Top             =   600
         Width           =   915
      End
      Begin VB.TextBox txtPrecio 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   3060
         TabIndex        =   8
         Top             =   600
         Width           =   1275
      End
      Begin MSDataListLib.DataCombo cmbProducto 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   600
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label5 
         Caption         =   "Total:"
         Height          =   195
         Left            =   6840
         TabIndex        =   10
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Precio:"
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   4980
         TabIndex        =   6
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Producto:"
         Height          =   255
         Left            =   300
         TabIndex        =   4
         Top             =   300
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7680
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSDataListLib.DataCombo cmbCliente 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "DataCombo1"
   End
   Begin MSComctlLib.ListView ListVInventario 
      Height          =   1935
      Left            =   180
      TabIndex        =   0
      Top             =   2100
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3413
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CodigoProducto"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Precio"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Cantidad"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "SubTotal"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.Label Label8 
      Caption         =   "TOTAL:"
      Height          =   255
      Left            =   5940
      TabIndex        =   16
      Top             =   4920
      Width           =   675
   End
   Begin VB.Label Label7 
      Caption         =   "Descuento:"
      Height          =   255
      Left            =   5700
      TabIndex        =   14
      Top             =   4560
      Width           =   915
   End
   Begin VB.Label Label6 
      Caption         =   "SubTotal:"
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
