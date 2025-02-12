VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fSistema 
   Caption         =   "Sistema de Carnetización ENTEK, C.A."
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15240
   Icon            =   "fSistema.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "fSistema.frx":0442
   ScaleHeight     =   8130
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   870
      Left            =   -30
      ScaleHeight     =   810
      ScaleWidth      =   15135
      TabIndex        =   13
      Top             =   1740
      Visible         =   0   'False
      Width           =   15195
      Begin VB.CommandButton Botones2 
         Height          =   780
         Index           =   9
         Left            =   9210
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   30
         Width           =   1000
      End
      Begin VB.CommandButton Botones2 
         Height          =   780
         Index           =   8
         Left            =   8190
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Botones2 
         Height          =   780
         Index           =   7
         Left            =   7170
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Botones2 
         Height          =   780
         Index           =   6
         Left            =   6150
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Botones2 
         Height          =   780
         Index           =   5
         Left            =   5130
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Botones2 
         Height          =   780
         Index           =   4
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Botones2 
         Height          =   780
         Index           =   3
         Left            =   3090
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Botones2 
         Height          =   780
         Index           =   2
         Left            =   2070
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Botones2 
         Height          =   780
         Index           =   1
         Left            =   1050
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Botones2 
         Height          =   780
         Index           =   0
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   870
      Left            =   -30
      ScaleHeight     =   810
      ScaleWidth      =   15135
      TabIndex        =   2
      Top             =   870
      Width           =   15195
      Begin VB.CommandButton Botones 
         Height          =   780
         Index           =   9
         Left            =   9210
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   30
         Width           =   1000
      End
      Begin VB.CommandButton Botones 
         Height          =   780
         Index           =   8
         Left            =   8190
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Botones 
         Height          =   780
         Index           =   7
         Left            =   7170
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Botones 
         Height          =   780
         Index           =   6
         Left            =   6150
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Botones 
         Height          =   780
         Index           =   5
         Left            =   5130
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Botones 
         Height          =   780
         Index           =   4
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Botones 
         Height          =   780
         Index           =   3
         Left            =   3090
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Botones 
         Height          =   780
         Index           =   2
         Left            =   2070
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Botones 
         Height          =   780
         Index           =   1
         Left            =   1050
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Botones 
         Height          =   780
         Index           =   0
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   30
         Visible         =   0   'False
         Width           =   1000
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6180
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fSistema.frx":153A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fSistema.frx":156C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fSistema.frx":157D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fSistema.frx":15B6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fSistema.frx":1C3A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fSistema.frx":1FDBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fSistema.frx":2103D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1535
      ButtonWidth     =   1773
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clientes"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sub-Clientes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Por Lotes"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Diario"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   7725
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "ENTEK, C.A."
            TextSave        =   "ENTEK, C.A."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "Versión 1.0 - 2009 (c)"
            TextSave        =   "Versión 1.0 - 2009 (c)"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "30/04/2012"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "08:15 a.m."
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mArc 
      Caption         =   "&Archivos"
      Begin VB.Menu mCli 
         Caption         =   "&Clientes"
      End
      Begin VB.Menu msc 
         Caption         =   "&Sub-Clientes"
      End
      Begin VB.Menu r1 
         Caption         =   "-"
      End
      Begin VB.Menu mCre 
         Caption         =   "&Crear Tabla Personas [Card-5]"
      End
      Begin VB.Menu mPA 
         Caption         =   "&Carnets Por Lotes"
      End
      Begin VB.Menu r2 
         Caption         =   "-"
      End
      Begin VB.Menu mfdi 
         Caption         =   "&Formatos de Diseños"
      End
      Begin VB.Menu mr7 
         Caption         =   "-"
      End
      Begin VB.Menu mPro 
         Caption         =   "&Productos Dpto. Producción"
         Enabled         =   0   'False
      End
      Begin VB.Menu mPV 
         Caption         =   "Productos &Dpto. Ventas"
      End
      Begin VB.Menu r4 
         Caption         =   "-"
      End
      Begin VB.Menu mU 
         Caption         =   "&Usuarios"
      End
      Begin VB.Menu mcu 
         Caption         =   "&Cambio de Usuario"
      End
      Begin VB.Menu mr12 
         Caption         =   "-"
      End
      Begin VB.Menu mSal 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mTran 
      Caption         =   "&Transacciones"
      Begin VB.Menu mEC 
         Caption         =   "&Emisión de Carnets [Mov. Diario]"
      End
      Begin VB.Menu r12 
         Caption         =   "-"
      End
      Begin VB.Menu mrp 
         Caption         =   "&Registro de Pagos"
      End
   End
   Begin VB.Menu mRep 
      Caption         =   "&Reportes"
      Begin VB.Menu mce 
         Caption         =   "&Carnets Entregados"
      End
      Begin VB.Menu mnu_vencimiento 
         Caption         =   "&Análisis de Vencimiento de Clientes"
      End
      Begin VB.Menu mnuLogs 
         Caption         =   "&Logs"
      End
   End
   Begin VB.Menu mSis 
      Caption         =   "&Sistema"
      Begin VB.Menu mCP 
         Caption         =   "&Combos de Productos"
      End
      Begin VB.Menu mOpc 
         Caption         =   "&Opciones Generales"
      End
      Begin VB.Menu mHer 
         Caption         =   "&Herramientas"
         Begin VB.Menu mEF 
            Caption         =   "&Etiquetador de Fotos"
         End
         Begin VB.Menu mGX 
            Caption         =   "&Generar XLS con Listado Fotos"
         End
         Begin VB.Menu mImF 
            Caption         =   "&Importar Fotos de Cámara/Simcard"
         End
         Begin VB.Menu mMCI 
            Caption         =   "&Marcar Carnets (ID) para Imprimir"
         End
         Begin VB.Menu mnuotrasutilidades 
            Caption         =   "&Otras Utilidades"
         End
      End
      Begin VB.Menu mCCCS 
         Caption         =   "&Carga de Clientes y SubClientes"
      End
      Begin VB.Menu mr9 
         Caption         =   "-"
      End
      Begin VB.Menu TransFotoCliente 
         Caption         =   "Transferir fotos entre clientes"
      End
      Begin VB.Menu mr8 
         Caption         =   "-"
      End
      Begin VB.Menu mCB 
         Caption         =   "Configurar &Botones"
      End
      Begin VB.Menu actualizartrigger 
         Caption         =   "ACTUALIZAR TRIGGER"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUsuarios 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu mnuPerfiles 
         Caption         =   "Perfiles de Usuario"
      End
   End
End
Attribute VB_Name = "fSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lCantidad As Integer

Private Sub actualizartrigger_Click()
Dim lReg As New ADODB.Recordset
Dim lCmd As New ADODB.Connection
Dim lSqltxt As String
On Error Resume Next
lCmd.Open Modulo.DBConexionSQL
Set lReg = lCmd.Execute("Select * from Personas order by cliente")
   Do While lReg.EOF = False
      ''lCmd.Execute "Drop Trigger [TRG_" & lReg!tabla & "]"
      lSqltxt = "Exec CrearTrigger '" & Trim(lReg!tabla) & "','" & Trim(lReg!Cliente) & "','" & Trim(lReg!SubCliente) & "'"
      lCmd.Execute lSqltxt
      If Err.Number = 0 Then MsgBox lSqltxt
      lReg.MoveNext
      
   Loop
   MsgBox "LISTO", vbInformation
falla:
   If Err.Number <> 0 Then
      MsgBox Err.Number & "::" & Err.Description, vbCritical
   End If
   
End Sub

Private Sub Botones_Click(Index As Integer)
  Dim i As Integer
  Dim e As Boolean
  Dim s As String
  Dim r As New ADODB.Recordset
  
  'saber si tiene sub boton:
  Picture2.Visible = False
  DoEvents
  i = 0
  j = 0
  e = False
    
  s = "select * from botones2 where botonprincipal = " & CStr(Index + 1) & " order by posicion"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If Not r.EOF Then
    e = True
    j = 0
    Do While Not r.EOF
      Botones2(j).Visible = False
      If Trim(r.Fields("titulo").value) <> "" Then
        Botones2(j).Caption = Trim(r.Fields("titulo").value)
        Botones2(j).ToolTipText = Trim(r.Fields("producto").value)
        Botones2(j).Picture = LoadPicture(Trim(r.Fields("imagen").value))
        Botones2(j).Visible = True
        Picture2.Visible = True
      End If
      j = j + 1
      r.MoveNext
    Loop
  Else
    Picture2.Visible = False
    
    'Disminuir el Producto en INVENTARIO:
    s = Trim(Botones(Index).ToolTipText)
    If s <> "" Then
         Modulo.Actualizar_Existencia_Producto s, (-1)
      
         s = "insert into ProductosMov (codigo,fecha,hora,cantidad,tipo) values ('" & _
          s & "','" & Format(Date, "yyyymmdd") & "','" & _
          Format(Time, "HH:mm") & "'," & (-1) & ",'VTA')"
          
         Modulo.ExecSQL s
         MsgBox "Codigo producto descontado: " & Trim(Botones(Index).ToolTipText) & ". Cantidad:1", vbInformation
      End If
  End If
   
  
    
  r.Close
  Set r = Nothing
  
End Sub


Private Sub Botones2_Click(Index As Integer)
  Dim s As String

  'Disminuir el Producto en INVENTARIO:
  s = Trim(Botones2(Index).ToolTipText)
  If s <> "" Then
       Modulo.Actualizar_Existencia_Producto s, (-1)
    
       s = "insert into ProductosMov (codigo,fecha,hora,cantidad,tipo) values ('" & _
        s & "','" & Format(Date, "yyyymmdd") & "','" & _
        Format(Time, "HH:mm") & "'," & (-1) & ",'VTA')"
          
       Modulo.ExecSQL s
       MsgBox "Codigo producto descontado: " & Trim(Botones2(Index).ToolTipText) & ". Cantidad:1", vbInformation
 End If
  
End Sub

Private Sub Botones_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then

  Dim i As Integer
  Dim e As Boolean
  Dim s As String
  Dim r As New ADODB.Recordset
  
  'saber si tiene sub boton:
  Picture2.Visible = False
  DoEvents
  i = 0
  j = 0
  e = False
    
  s = "select * from botones2 where botonprincipal = " & CStr(Index + 1) & " order by posicion"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If Not r.EOF Then
    e = True
    j = 0
    Do While Not r.EOF
      Botones2(j).Visible = False
      If Trim(r.Fields("titulo").value) <> "" Then
        Botones2(j).Caption = Trim(r.Fields("titulo").value)
        Botones2(j).ToolTipText = Trim(r.Fields("producto").value)
        Botones2(j).Picture = LoadPicture(Trim(r.Fields("imagen").value))
        Botones2(j).Visible = True
        Picture2.Visible = True
      End If
      j = j + 1
      r.MoveNext
    Loop
  Else
    Picture2.Visible = False
    
  'Disminuir el Producto en INVENTARIO:
  s = Trim(Botones(Index).ToolTipText)
  If s <> "" Then
     Load frmCantidad
     frmCantidad.Caption = "Descontar: " & s
     frmCantidad.Left = Botones2(Index).Left
     frmCantidad.Top = Botones2(Index).Top + 2050
       'frmCantidad.txtCantidad.SetFocus
       frmCantidad.Show vbModal
    If lCantidad > 0 Then
       Modulo.Actualizar_Existencia_Producto s, lCantidad * (-1)
    
       s = "insert into ProductosMov (codigo,fecha,hora,cantidad,tipo) values ('" & _
        s & "','" & Format(Date, "yyyymmdd") & "','" & _
        Format(Time, "HH:mm") & "'," & lCantidad * (-1) & ",'VTA')"
          
       Modulo.ExecSQL s
    End If
  End If
    'MsgBox "Cantidad..."
  End If
    
  r.Close
  Set r = Nothing
End If
End Sub

Private Sub Form_Load()
  Dim sOri As String
  Dim sDes As String
  Dim sH As Integer
 On Error GoTo falla
  If App.PrevInstance = True Then
     MsgBox "Ya se está ejecutando un instancia del programa en en este equipo", vbInformation
     End
  End If
  Me.Left = 0
  Me.Top = 0

  sOri = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
  sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
  
  sH = Shell("Net Use Z: \\diseño\d", vbHide)
  
  Picture1.width = Me.width
  Picture2.width = Me.width
  
  Picture1.Visible = True
  Picture2.Visible = False
  
  CargarConfigBotones
  StatusBar1.Panels(2).Text = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
  

  'Shell sOri, vbHide
falla:
   If Err.Number <> 0 Then
    Exit Sub
   End If
End Sub

Public Sub sVerificarAccesos()
  Toolbar1.Buttons(1).Enabled = gPerfil.Clientes
  Me.mCli.Enabled = gPerfil.Clientes
  Toolbar1.Buttons(2).Enabled = gPerfil.SubClientes
  Me.msc.Enabled = gPerfil.SubClientes
  Toolbar1.Buttons(3).Enabled = gPerfil.PorLotes
  Me.mPA.Enabled = gPerfil.PorLotes
  Toolbar1.Buttons(4).Enabled = gPerfil.Diario
  Me.mEC.Enabled = gPerfil.Diario
  Me.mCre.Enabled = gPerfil.TablasPersonas
  Me.mfdi.Enabled = gPerfil.FormatoDiseño
  Me.mPV.Enabled = gPerfil.Inventario
  Me.mU.Enabled = gPerfil.Usuarios
  Me.mrp.Enabled = gPerfil.Pagos
  Me.mRep.Enabled = gPerfil.Reportes
  Me.mCP.Enabled = gPerfil.ComboProductos
  Me.mOpc.Enabled = gPerfil.OpcionesGenerales
  Me.mCCCS.Enabled = gPerfil.CargaClientes
  Me.TransFotoCliente.Enabled = gPerfil.TransferirFotos
  Me.mHer.Enabled = gPerfil.Herramientas
  Me.mCB.Enabled = gPerfil.configurarBotones
  Me.mnuUsuarios.Enabled = gPerfil.Usuarios
  Me.mnuPerfiles.Enabled = gPerfil.PerfilesAcceso
End Sub

Public Sub CargarConfigBotones()
  Dim s As String
  Dim r As New ADODB.Recordset
  Dim i As Integer
  
  s = "select * from Botones order by Posicion"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  i = 1
  
  Picture1.Visible = True

  Do While Not r.EOF
    bMostrar = False
    
    Botones(i - 1).Enabled = False
    
    If Trim(Trim(r.Fields("caption").value)) <> "" Then
      Botones(i - 1).Caption = Trim(r.Fields("caption").value)
      Botones(i - 1).Enabled = True
      Botones(i - 1).Visible = True
    Else
      Botones(i - 1).Caption = ""
    End If
    
    If Trim(r.Fields("producto").value) <> "" Then
      Botones(i - 1).ToolTipText = Trim(r.Fields("producto").value)
    Else
      Botones(i - 1).ToolTipText = ""
    End If
    
    If Trim(r.Fields("imagen").value) <> "" Then
      If Dir(Trim(r.Fields("imagen").value)) <> "" Then
        Botones(i - 1).Picture = LoadPicture(Trim(r.Fields("imagen").value))
      Else
        Botones(i - 1).Picture = Nothing
      End If
    Else
      Botones(i - 1).Picture = Nothing
    End If
    
    i = i + 1
    r.MoveNext
    
  Loop
  r.Close
  Set r = Nothing
  
  If HayBotonActivo() Then
    Picture1.Visible = True
    Apagar_Botones_Inactivos
  End If
  
    
End Sub



Private Function HayBotonActivo() As Boolean
  Dim i As Integer
  Dim e As Boolean
  e = False
  i = 0
  Do While i < Modulo.MAX_BTNS And Not e
    If Botones(i).Enabled Then e = True Else i = i + 1
  Loop
  HayBotonActivo = e
End Function

Private Sub Apagar_Botones_Inactivos()
  For i = 0 To Modulo.MAX_BTNS - 1
    If Botones(i).Enabled = False Then
      Botones(i).Visible = False
    End If
  Next i
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If DBConexionSQL.State <> adStateClosed Then
    DBConexionSQL.Close
    DBConexionSQL.ConnectionString = ""
  End If
  End
End Sub

Private Sub Form_Resize()
  Picture1.width = Me.width
  Picture2.width = Me.width
End Sub

Private Sub mCB_Click()
  Load fConfigurarBotones
  fConfigurarBotones.Show
End Sub

Private Sub mCCCS_Click()
  fCargaClientes.Show
End Sub

Private Sub mce_Click()
  Load fRep01
  fRep01.Show
End Sub

Private Sub mCli_Click()
  Load fClientes
  fClientes.Show
End Sub

Private Sub mCP_Click()
  fCombosProg.Show
End Sub

Private Sub mCre_Click()
  Load fPersonas
  fPersonas.Show
End Sub

Private Sub mcu_Click()
  Load fLogin
  fLogin.Show vbModal
  
  If USUARIO_ACTUAL <> "" Then
  
    'CONEXION_SQL = "Provider=SQLOLEDB.1;" & _
                 "Password=sql123;" & _
                 "Persist Security Info=True;" & _
                 "User ID=sa;" & _
                 "Initial Catalog=santek;" & _
                 "Data Source=" & Modulo.IP_Servidor
  
    If Not Abrir_BD() Then End
    
    fSistema.Caption = TIT_SISTEMA & " - Estación:" & ESTACION & " Conectado " & Modulo.IP_Servidor & " Usuario: " & Modulo.USUARIO_ACTUAL
  
    fSistema.Show
    fSistema.sVerificarAccesos
    AgregarLogs "Cambio/Usuario->Inicia Sesión"
    
  Else
    'End
  End If

End Sub

Private Sub mEC_Click()
 ' Load fMov
  fMov.Show
End Sub

Private Sub mEF_Click()
  fHEtiquetador.Show
End Sub

Private Sub mfdi_Click()
  'Load fFD
  fFD.Show
End Sub

Private Sub mGX_Click()
  'Load fHXLS
  fHXLS.Show
End Sub

Private Sub mImF_Click()
  'Load fHCopiarOD
  fHCopiarOD.Show
End Sub

Private Sub mMCI_Click()
  'Load fHPersonaID
  fHPersonaID.Show
End Sub

Private Sub mnu_vencimiento_Click()
  frmPreviewVencimiento.Show vbModal
End Sub

Private Sub mnuLogs_Click()
   Form1.Show
End Sub

Private Sub mnuotrasutilidades_Click()
  frmUtilFiles.Show vbModal
End Sub

Private Sub mnuPerfiles_Click()
   frmPerfiles.Show vbModal
End Sub

Private Sub mnuUsuarios_Click()
   frmUsuarios.Show vbModal
End Sub

Private Sub mOpc_Click()
  'Load fOpciones
  fOpciones.Show
End Sub

Private Sub mPA_Click()
  'Load fPersonasAct
  fPersonasAct.Show
End Sub

Private Sub mr5_Click()
  'Load fVentas
  fVentas.Show
End Sub

Private Sub mPV_Click()
  Dim s As String
  Load fProductos
  With fProductos
    .Combo1.Clear
    .Combo1.AddItem "CÓDIGO"
    .Combo1.AddItem "DESCRIPCIÓN"
    .Combo1.ListIndex = 0
    s = "select * from Productos order by codigo"
    .Adodc1.ConnectionString = Modulo.DBConexionSQL.ConnectionString
    .Adodc1.RecordSource = s
    .Adodc1.Refresh
    .Productos_Format_DataGrid
  End With
  fProductos.Show

End Sub

Private Sub mrp_Click()
 
  fPagos.Show
  fPagos.Text1.SetFocus
End Sub

Private Sub msc_Click()
  Load fSubClientes
  fSubClientes.Show
End Sub

Private Sub mU_Click()
  Load fUsuarios
  fUsuarios.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Botones(0).Visible = True
  
  Select Case UCase(Button.Caption)
    Case "SALIR": End
    
    Case "CLIENTES":
      Load fClientes
      fClientes.Show
      
    Case "SUB-CLIENTES":
      Load fSubClientes
      fSubClientes.Show
      
    Case "POR LOTES":
      Load fPersonasAct
      fPersonasAct.Show
  
    Case "DIARIO":
      Load fMov
      fMov.Show
  
  End Select
  
End Sub

Private Sub TransFotoCliente_Click()
   frmTransFotoCliente.Show
End Sub
