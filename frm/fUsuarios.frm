VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fUsuarios 
   Caption         =   "USUARIOS"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Usuarios"
      Height          =   6525
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton bAceptar 
         Caption         =   "Guardar"
         Height          =   500
         Left            =   8220
         Picture         =   "fUsuarios.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5730
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton bCancelar 
         Caption         =   "Cancelar"
         Height          =   500
         Left            =   9360
         Picture         =   "fUsuarios.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5730
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.Frame Frame4 
         Caption         =   "Información del Usuario"
         Height          =   5385
         Left            =   6420
         TabIndex        =   22
         Top             =   1050
         Width           =   5505
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   3720
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   4110
            Width           =   1600
         End
         Begin VB.ListBox List1 
            Enabled         =   0   'False
            Height          =   2085
            Left            =   840
            Style           =   1  'Checkbox
            TabIndex        =   3
            Top             =   1890
            Width           =   4485
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1470
            Width           =   4500
         End
         Begin MSComCtl2.DTPicker tinicio 
            Height          =   285
            Left            =   840
            TabIndex        =   4
            Top             =   4140
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            _Version        =   393216
            Format          =   50200577
            CurrentDate     =   39963
         End
         Begin VB.TextBox tdir 
            Height          =   315
            Left            =   840
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   1080
            Width           =   4500
         End
         Begin VB.TextBox tnom 
            Height          =   315
            Left            =   840
            MaxLength       =   100
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   660
            Width           =   4500
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Estatus:"
            Height          =   195
            Left            =   3060
            TabIndex        =   30
            Top             =   4200
            Width           =   570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Permisos:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   90
            TabIndex        =   29
            Top             =   1950
            Width           =   675
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Inicio:"
            Height          =   195
            Left            =   330
            TabIndex        =   27
            Top             =   4200
            Width           =   420
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Nivel:"
            Height          =   195
            Left            =   360
            TabIndex        =   26
            Top             =   1530
            Width           =   405
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Clave:"
            Height          =   195
            Left            =   330
            TabIndex        =   25
            Top             =   1110
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Usuario:"
            Height          =   195
            Left            =   180
            TabIndex        =   24
            Top             =   690
            Width           =   585
         End
         Begin VB.Label lcod 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "000000"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   840
            TabIndex        =   18
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ID:"
            Height          =   195
            Left            =   540
            TabIndex        =   23
            Top             =   300
            Width           =   210
         End
      End
      Begin VB.Frame Frame3 
         Height          =   885
         Left            =   6420
         TabIndex        =   21
         Top             =   150
         Width           =   5505
         Begin VB.CommandButton bBuscar 
            Caption         =   "Buscar"
            Height          =   550
            Left            =   2880
            Picture         =   "fUsuarios.frx":0B14
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Buscar"
            Top             =   165
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.CommandButton bSalir 
            Caption         =   "Salir"
            Height          =   550
            Left            =   4260
            Picture         =   "fUsuarios.frx":109E
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Salir"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton bBorrar 
            Caption         =   "Borrar"
            Height          =   550
            Left            =   1950
            Picture         =   "fUsuarios.frx":1628
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Borrar"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton bEditar 
            Caption         =   "Editar"
            Height          =   550
            Left            =   1050
            Picture         =   "fUsuarios.frx":1BB2
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Editar"
            Top             =   165
            Width           =   900
         End
         Begin VB.CommandButton bNuevo 
            Caption         =   "Nuevo"
            Height          =   550
            Left            =   150
            Picture         =   "fUsuarios.frx":213C
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Nuevo"
            Top             =   165
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Height          =   765
         Left            =   1620
         TabIndex        =   20
         Top             =   5610
         Width           =   2600
         Begin VB.CommandButton bUltimo 
            Height          =   500
            Left            =   1890
            Picture         =   "fUsuarios.frx":26C6
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Ultimo"
            Top             =   160
            Width           =   600
         End
         Begin VB.CommandButton bSiguiente 
            Height          =   500
            Left            =   1290
            Picture         =   "fUsuarios.frx":2C50
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Siguiente"
            Top             =   160
            Width           =   600
         End
         Begin VB.CommandButton bAnterior 
            Height          =   500
            Left            =   690
            Picture         =   "fUsuarios.frx":31DA
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Anterior"
            Top             =   160
            Width           =   600
         End
         Begin VB.CommandButton bPrimer 
            Height          =   500
            Left            =   90
            Picture         =   "fUsuarios.frx":3764
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Inicio"
            Top             =   160
            Width           =   600
         End
      End
      Begin MSFlexGridLib.MSFlexGrid FG 
         Height          =   5415
         Left            =   30
         TabIndex        =   8
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   9551
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "ID    | Usuario                                                        | Clave                         | Nivel "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   270
         TabIndex        =   28
         Top             =   9300
         Width           =   45
      End
   End
End
Attribute VB_Name = "fUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const FCG = "ID    | Usuario                                                        | Clave                         | Nivel "

Dim RU As New ADODB.Recordset

Const OPR_NUEVO = 1
Const OPR_EDITAR = 2

Dim OPR As Integer  'operacion: 1 nuevo  2 editar
Dim Viendo As Boolean


Private Sub Limpiar_FG()
  FG.Clear
  FG.Rows = 2
  FG.FormatString = FCG
End Sub

Private Sub Limpiar_Txts()
  tnom.Text = ""
  tdir.Text = ""
  Combo1.Clear
  Combo1.AddItem "1-TODO"
  Combo1.AddItem "2-MODIFICAR/ELIMINAR"
  Combo1.AddItem "3-AGREGAR"
  Combo1.AddItem "4-CONSULTAR"
  Combo1.ListIndex = 0
  
  List1.Clear
  tinicio.Value = Date
  Combo2.Clear
  Combo2.AddItem "Activo"
  Combo2.AddItem "Suspendido"
  Combo2.ListIndex = 0
  
  
End Sub

Private Sub Activar_Txts(TF As Boolean)
  tnom.Enabled = TF
  tdir.Enabled = TF
  Combo1.Enabled = TF
  List1.Enabled = TF
  tinicio.Enabled = TF
  Combo2.Enabled = TF
End Sub

Private Sub Activar_Btns(Num As Integer, TF As Boolean)
  If Num = 0 Then 'TODOS los botones
  
    bPrimer.Enabled = TF
    bAnterior.Enabled = TF
    bSiguiente.Enabled = TF
    bUltimo.Enabled = TF
    
    bNuevo.Enabled = TF
    bEditar.Enabled = TF
    bBorrar.Enabled = TF
    bBuscar.Enabled = TF
    
    bSalir.Enabled = TF
       
    'bSubClientes.Enabled = TF
    
    bAceptar.Enabled = TF
    bCancelar.Enabled = TF
    
  Else
  
    Select Case Num
      
      Case 1: bPrimer.Enabled = TF
      Case 2: bAnterior.Enabled = TF
      Case 3: bSiguiente.Enabled = TF
      Case 4: bUltimo.Enabled = TF
      
      Case 5: bNuevo.Enabled = TF
      Case 6: bEditar.Enabled = TF
      Case 7: bBorrar.Enabled = TF
      Case 8: bBuscar.Enabled = TF
      
      Case 9: bSalir.Enabled = TF
          
      Case 10: 'bSubClientes.Enabled = TF
      
      Case 11: bAceptar.Enabled = TF
      Case 12: bCancelar.Enabled = TF
      
    End Select
    
  End If
End Sub

Private Sub Cargar_Usuarios()
  'Dim r As New ADODB.Recordset
  Dim l As Integer
  
  Viendo = False
  
  Limpiar_FG
  
  FG.Clear
  FG.FormatString = FCG
  FG.Rows = 2

  
  If RU.State <> adStateClosed Then RU.Close
    
  RU.Open "SELECT * FROM usuarios ORDER BY id", DBConexionSQL, adOpenDynamic, adLockOptimistic
  l = 1
  Do While Not RU.EOF
    FG.TextMatrix(l, 0) = Zeros(RU.Fields("id").Value, 6)
    FG.TextMatrix(l, 1) = Trim(RU.Fields("usuario").Value)
    FG.TextMatrix(l, 2) = Trim(RU.Fields("clave").Value)
    FG.Row = l: FG.Col = 2: FG.CellAlignment = flexAlignLeftCenter
    
    FG.TextMatrix(l, 3) = Trim(RU.Fields("nivel").Value)
    FG.Row = l: FG.Col = 3: FG.CellAlignment = flexAlignLeftCenter
    
    RU.MoveNext
    
    FG.Col = 0
        
    If Not RU.EOF Then
      l = l + 1
      FG.Rows = FG.Rows + 1
    End If
  Loop
  
  FG.Refresh

  If FG.Rows >= 1 Then
    FG.Row = 1
    FG.Col = 0
    Scroll_Usuario
    'FG.SetFocus
  End If
    
  If Not RU.BOF And Not RU.EOF Then RU.MoveFirst
  
  Call Mostrar_Usuario
  
  Label2.Caption = "[" & CStr(Total_Usuarios()) & " Regs.]"
  'RClientes.Close
  'Set r = Nothing
  Viendo = True
End Sub



Private Sub Scroll_Usuario()
  Dim c As String
  Viendo = False
  If FG.Row >= 1 Then
    c = Trim(FG.TextMatrix(FG.Row, 0))
    If c <> "" Then
      RU.MoveFirst
      RU.Find "id = " & c
      If RU.EOF Then
        MsgBox "Debe Seleccionar el Usuario...", vbCritical, "Información"
      Else
        Mostrar_Usuario
      End If
    Else
      Limpiar_Txts
    End If
  End If
  Viendo = True
End Sub

Private Sub Ubicar_Cursor_Usuario(sCod As String)
  Dim i As Integer
  Dim f As Integer
  Dim e As Boolean
  
  i = 1
  If FG.Rows >= 1 Then
    e = False
    Do While i < FG.Rows And Not e
      If FG.TextMatrix(i, 0) = sCod Then
        e = True
      Else
        i = i + 1
      End If
    Loop
    If (i + 1) < FG.Rows Then FG.Row = i + 1
  End If
      
End Sub

Private Sub Mostrar_Usuario()
  Dim s As String, sRutaCliente As String
  
  Combo1.Clear
  Combo1.AddItem "1-TODO"
  Combo1.AddItem "2-MODIFICAR/ELIMINAR"
  Combo1.AddItem "3-AGREGAR"
  Combo1.AddItem "4-CONSULTAR"
  Combo1.ListIndex = 0
  
  Combo2.Clear
  Combo2.AddItem "Activo"
  Combo2.AddItem "Suspendido"
  Combo2.ListIndex = 0
  

  
  If RU.State <> adStateClosed Then
    If Not RU.EOF Then
      lcod.Caption = Zeros(RU.Fields("id").Value, 6)
      tnom.Text = Trim(RU.Fields("usuario").Value)
      tdir.Text = Trim(RU.Fields("clave").Value)
      Combo1.ListIndex = Modulo.Buscar_ComboLen(Combo1, RU.Fields("nivel").Value, 1)
      
      tinicio.Value = RU.Fields("creado").Value
      Combo2.ListIndex = Modulo.Buscar_ComboLen(Combo2, RU.Fields("estatus").Value, 1)
    End If
  End If
End Sub

Function CodigoUsuarioNuevo() As Long
  Dim r As New ADODB.Recordset
  Dim ccn As Long
  ccn = 0
  r.Open "SELECT * FROM usuario ORDER BY id", DBConexionSQL, adOpenKeyset, adLockReadOnly
  If Not r.EOF Then
    r.MoveLast
    ccn = r.Fields("id").Value
  End If
  CodigoUsuarioNuevo = ccn + 1
End Function

Function Total_Usuarios() As Long
  Dim r As New ADODB.Recordset
  Dim ccn As Long
  ccn = 0
  r.Open "SELECT count(*) FROM usuarios", DBConexionSQL, adOpenKeyset, adLockReadOnly
  If Not r.EOF Then
    If Not IsNull(r.Fields(0).Value) Then
      ccn = r.Fields(0).Value
    End If
  End If
  r.Close
  Set r = Nothing
  Total_Usuarios = ccn
End Function


Private Sub bAceptar_Click()
  Dim s As String, nc As String, s2 As String
  Dim c As Long
  Dim r As New ADODB.Recordset
  Dim sNom As String, s3 As String
  Dim xpre As Long
  
  tnom.Text = Trim(tnom.Text)
  
  If tnom.Text = "" Then
    MsgBox "Faltan Datos, Revise...", vbCritical, "Información"
    tnom.SetFocus
    Exit Sub
  End If
  
  If OPR = OPR_NUEVO Then
  
    If Modulo.DBExiste("usuarios", "usuario", tnom.Text) Then
      MsgBox "Usuario [" & tnom.Text & "] ya está Registrado...", vbCritical, "Información"
      Exit Sub
    End If
  
    Load fMensaje
    fMensaje.Label1.Caption = "Añadiendo Nuevo Usuario, Espere..."
    fMensaje.Show
   
    On Error Resume Next
      
    With RU
      .AddNew
      .Fields("usuario").Value = Trim(tnom.Text)
      .Fields("clave").Value = Trim(tdir.Text)
      .Fields("nivel").Value = Mid(Combo1.Text, 1, 1)
      .Fields("permisos").Value = "SS"
      .Fields("estatus").Value = Mid(Combo2.Text, 1, 1)
      .Fields("creado").Value = tinicio.Value
      .Update
    End With
    
    RU.Close
    RU.Open
    If Not RU.EOF Then RU.MoveLast
    c = RU.Fields("id").Value 'id asignado por el xSQL
    
    Unload fMensaje
    
    If Err.Number <> 0 Then
      MsgBox "Ha Ocurrido un Error al Intentar almacenar el Registro..." & vbCrLf & Err.Description, vbCritical, "Información"
      Exit Sub
    End If
    
    AgregarLogs "Agrega usuario [" & Trim(tnom.Text) & "]"
    
    
    Cargar_Usuarios
    Scroll_Usuario
    Limpiar_Txts
    lcod.Caption = Zeros(c + 1, 6)
    
    tnom.SetFocus
      
    
  Else
  
    If OPR = OPR_EDITAR Then
    
      Load fMensaje
      fMensaje.Label1.Caption = "Actualizando Usuario, Espere..."
      fMensaje.Show
   
      On Error Resume Next
      
      nc = lcod.Caption
      
      s = "UPDATE usuarios SET " & _
          "usuario  = '" & Trim(tnom.Text) & "'," & _
          "clave    = '" & Trim(tdir.Text) & "'," & _
          "nivel    = '" & Mid(Combo1.Text, 1, 1) & "'," & _
          "permisos = '" & "SS" & "'," & _
          "estatus  = '" & Mid(Combo2.Text, 1, 1) & "'," & _
          "creado   = '" & Format(tinicio.Value, "yyyymmdd") & "' " & _
          "WHERE " & _
          "id = " & nc
          
      Set DBComandoSQL.ActiveConnection = DBConexionSQL
      DBComandoSQL.CommandText = s
      DBComandoSQL.Execute
      
      Unload fMensaje
      
      If Err.Number <> 0 Then
        MsgBox "Ha Ocurrido un Error al Intentar actualizar el Registro..." & vbCrLf & Err.Description, vbCritical, "Información"
        Exit Sub
      End If
      
      AgregarLogs "Edita usuario [" & Trim(tnom.Text) & "]"
      
      Cargar_Usuarios
            
      bCancelar_Click
      
      Ubicar_Cursor_Usuario nc
    
    End If

  End If

End Sub


Private Sub bBorrar_Click()
  Dim c As String, n As String, s As String
  Dim s2 As String
  Dim r As New ADODB.Recordset
  
  If FG.Row >= 1 Then
    c = Trim(FG.TextMatrix(FG.Row, 0))
    n = Trim(FG.TextMatrix(FG.Row, 1))
    If c <> "" Then
      
      If MsgBox("¿Está Seguro de Borrar el Usuario " & vbCrLf & c & "-" & n & " ?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
                                        
        s = "DELETE FROM usuarios WHERE id = " & c
        Set DBComandoSQL.ActiveConnection = DBConexionSQL
        DBComandoSQL.CommandText = s
        DBComandoSQL.Execute
                                
        MsgBox "USUARIO [" & c & "] FUE BORRADO (" & Modulo.USUARIO_ACTUAL & " " & Format(Now, "dd/mm/yy hh:mm ampm") & ")", vbInformation, "Información"
        
        AgregarLogs "Borra usuario [" & n & "]"
        
        'FALTA LOGs
        Cargar_Usuarios  'Abre el RecordSet de Clientes
        Scroll_Usuario
      End If
    End If
  End If
End Sub

Private Sub Buscar_Cursor_Usuario(sCod As String)
  Dim i As Integer
  Dim f As Integer
  Dim e As Boolean
  
  i = 1
  If FG.Rows >= 1 Then
    e = False
    Do While i < FG.Rows And Not e
      FG.Row = i
      If FG.TextMatrix(i, 0) = sCod Then
        e = True
      Else
        i = i + 1
      End If
    Loop
  End If
      
End Sub


Private Sub bBuscar_Click()
  Dim f As Integer, c As Integer, i As Integer
  Dim e As Boolean
  
  Modulo.vTemporal1 = ""
  Load fBuscarSimple
  
  With fBuscarSimple
    .Adodc1.ConnectionString = Modulo.DBConexionSQL.ConnectionString
    .Adodc1.Caption = "USUARIOS"
    .Adodc1.RecordSource = "select * from usuarios order by id"
    .Adodc1.Refresh
    .DataGrid1.Refresh
    
    .Combo1.Clear
    .Combo1.AddItem "ID"
    .Combo1.AddItem "USUARIO"
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
    Viendo = False
    i = 1
    e = False
    Do While i < FG.Rows And Not e
      If FG.TextMatrix(i, 0) = Modulo.vTemporal1 Then
        e = True
        FG.Row = i
        FG.Col = 0
      Else
        i = i + 1
      End If
    Loop
    FG.SetFocus
    SendKeys "{LEFT}"
    Scroll_Usuario
    
  End If
  
  
End Sub

Private Sub bEditar_Click()
  If FG.Row >= 1 Then
  
    If Trim(FG.TextMatrix(FG.Row, 0)) = "" Then Exit Sub
    
  Else
  
    Exit Sub
    
  End If
  
  OPR = OPR_EDITAR
  
  fClientes.Caption = "EDITAR CLIENTE"
  
  FG.Enabled = False
  Frame2.Enabled = False
  Activar_Btns 5, False
  Activar_Btns 6, False
  Activar_Btns 7, False
  Activar_Btns 8, False
  
  'Limpiar_Txts
  Activar_Txts True
  
  Activar_Btns 10, True
  Activar_Btns 11, True
  Activar_Btns 12, True
  
  'lcod.Caption = Zeros(CodigoClienteNuevo(), 6)
  tnom.Enabled = False
  lcod.Caption = Trim(FG.TextMatrix(FG.Row, 0))
  
  tdir.SetFocus
  
  
End Sub

Private Sub bNuevo_Click()
  Dim sOri As String
  Dim sDes As String
    
  bUltimo_Click
  DoEvents

  OPR = OPR_NUEVO
  
  fUsuarios.Caption = "NUEVO USUARIO"
  
  FG.Enabled = False
  Frame2.Enabled = False
  Activar_Btns 5, False
  Activar_Btns 6, False
  Activar_Btns 7, False
  Activar_Btns 8, False
  
  Limpiar_Txts
  Activar_Txts True
  
  Activar_Btns 10, True
  Activar_Btns 11, True
  Activar_Btns 12, True
  
  lcod.Caption = "Nuevo"
  
  tnom.SetFocus
  
End Sub

Private Sub bCancelar_Click()
  Dim s As String
  
  OPR = 0
  
  fUsuarios.Caption = "USUARIOS"
  
  FG.Enabled = True
  Frame2.Enabled = True
  Activar_Btns 5, True
  Activar_Btns 6, True
  Activar_Btns 7, True
  Activar_Btns 8, True
  
  Limpiar_Txts
  Activar_Txts False
  lcod.Caption = ""
  
  Activar_Btns 10, False
  Activar_Btns 11, False
  Activar_Btns 12, False
  
  Scroll_Usuario
  
  FG.SetFocus
  SendKeys "{LEFT}"

  
  

End Sub

Private Sub bPrimer_Click()
  'Ir al primer registro del listado
  FG.Row = 1
  Scroll_Usuario
  FG.SetFocus
  SendKeys "{UP}"
End Sub

Private Sub bAnterior_Click()
  'Ir al anterior registro del listado
  If FG.Rows > 2 Then
    FG.Row = FG.Row - 1
    Scroll_Usuario
  End If
  FG.SetFocus
  SendKeys "{LEFT}"
End Sub

Private Sub bSiguiente_Click()
  'Ir al siguiente registro del listado
  If FG.Rows > 2 Then
    If FG.Row <= FG.Rows - 1 Then
      FG.Row = FG.Row + 1
      Scroll_Usuario
    End If
  End If
  FG.SetFocus
  SendKeys "{RIGHT}"
End Sub



Private Sub bUltimo_Click()
  'Ir al ultimo registro del listado
  If FG.Rows > 1 Then
    If FG.Row >= 1 Then
      FG.Row = FG.Rows - 1
      Scroll_Usuario
    End If
  End If
  FG.SetFocus
  SendKeys "{DOWN}"
End Sub

Private Sub bSalir_Click()
  Unload Me
End Sub

Private Sub FG_Click()
  Scroll_Usuario
End Sub

Private Sub FG_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyHome Then bPrimer_Click Else
  If KeyCode = vbKeyEnd Then bUltimo_Click
End Sub

Private Sub FG_RowColChange()
  If Viendo Then Scroll_Usuario
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

Private Sub Form_Load()
  Viendo = False
  OPR = 0
  
  Cargar_Usuarios  'Abre el RecordSet de Usuarios
  
  Limpiar_Txts
  Activar_Txts False
  Activar_Btns 0, False
  Activar_Btns 9, True
  lcod.Enabled = False
  
  Mostrar_Usuario
  
  If FG.Rows > 1 Then
    Activar_Btns 1, True
    Activar_Btns 2, True
    Activar_Btns 3, True
    Activar_Btns 4, True
  End If
  
  Activar_Btns 5, True
  Activar_Btns 6, True
  Activar_Btns 7, True
  Activar_Btns 8, True
  
  Viendo = True
  
  SendKeys "{UP}"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If RU.State <> adStateClosed Then
    RU.Close
    Set RU = Nothing
  End If
  Unload Me
End Sub


Private Sub tinicio_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then Combo2.SetFocus
End Sub

