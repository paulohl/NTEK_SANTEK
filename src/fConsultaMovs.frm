VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fConsultaMovs 
   Caption         =   "Consultar Histórico de Movimientos"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10155
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   9765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15250
      Begin MSFlexGridLib.MSFlexGrid FG 
         Height          =   8385
         Left            =   60
         TabIndex        =   18
         Top             =   1320
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   14790
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFC0FF&
         Height          =   675
         Left            =   13920
         ScaleHeight     =   615
         ScaleWidth      =   1125
         TabIndex        =   15
         Top             =   420
         Width           =   1185
         Begin VB.CommandButton bBuscar 
            Caption         =   "Búscar"
            Height          =   500
            Left            =   120
            Picture         =   "fConsultaMovs.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   900
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1035
         Left            =   3180
         TabIndex        =   12
         Top             =   180
         Width           =   4485
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   360
            Width           =   1245
         End
         Begin VB.TextBox eCedNom 
            Height          =   285
            Left            =   2370
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   360
            Width           =   2000
         End
         Begin VB.CheckBox CheckCedNom 
            Caption         =   "Filtrar =>"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   420
            Width           =   915
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1035
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   2985
         Begin MSComCtl2.DTPicker eDesde 
            Height          =   315
            Left            =   900
            TabIndex        =   9
            Top             =   210
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62128129
            CurrentDate     =   40008
         End
         Begin MSComCtl2.DTPicker eHasta 
            Height          =   315
            Left            =   900
            TabIndex        =   11
            Top             =   570
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62128129
            CurrentDate     =   40008
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            Height          =   195
            Left            =   210
            TabIndex        =   10
            Top             =   630
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Left            =   210
            TabIndex        =   8
            Top             =   270
            Width           =   510
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1035
         Left            =   7680
         TabIndex        =   1
         Top             =   180
         Width           =   6200
         Begin VB.CommandButton bBuscarCP 
            Height          =   345
            Left            =   5700
            Picture         =   "fConsultaMovs.frx":0A02
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   420
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CheckBox CheckCli 
            Caption         =   "Filtrar =>"
            Height          =   195
            Left            =   90
            TabIndex        =   6
            Top             =   420
            Width           =   915
         End
         Begin VB.ComboBox cCP 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   210
            Width           =   3800
         End
         Begin VB.ComboBox cSC 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   630
            Width           =   3800
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   1200
            TabIndex        =   5
            Top             =   270
            Width           =   525
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Cliente:"
            Height          =   195
            Left            =   870
            TabIndex        =   4
            Top             =   690
            Width           =   855
         End
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      Height          =   255
      Left            =   330
      TabIndex        =   22
      Top             =   9840
      Width           =   3660
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9.999.999,99"
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
      Left            =   13590
      TabIndex        =   21
      Top             =   9840
      Width           =   1350
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Monto Bs"
      Height          =   255
      Left            =   12240
      TabIndex        =   20
      Top             =   9840
      Width           =   1350
   End
End
Attribute VB_Name = "fConsultaMovs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EsFaltaPago As Boolean
Private Sub FormatearDG()
  Dim i As Integer
  For i = 0 To DataGrid1.Columns.Count - 1
    DataGrid1.Columns(i).Caption = UCase(DataGrid1.Columns(i).Caption)
  Next i
  
  DataGrid1.Columns(0).Visible = False
  
  
End Sub

Public Sub bBuscar_Click()
  Dim s As String, fd As String, fh As String
  
  Dim rM As New ADODB.Recordset 'Movimientos
  Dim rc As New ADODB.Recordset 'Clientes
  Dim rS As New ADODB.Recordset 'Subclientes
  Dim rP As New ADODB.Recordset 'Personas
  
  Dim l1 As Long, l2 As Long
  Dim s1 As String, s2 As String
  
  Dim TM As Double
  
  If eDesde.value > eHasta.value Then
    MsgBox "Las Fechas son inválidas, Revise...", vbCritical, "Información"
    Exit Sub
  End If
  
  rS.Open "select * from subclientes order by id", Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
   
  fd = Format(eDesde.value, "yyyy/mm/dd")
  fh = Format(eHasta.value, "yyyy/mm/dd")
  
  's = "select D.id, D.Fecha, D.Hora, D.cliente, D.subcliente from Diario as D " & _
      "inner join clientes as c on D.cliente = C.cliente where fecha >= '" & fd & "' and D.fecha <= '" & fh & "' order by D.id"
      
  s = "select * from Diario where Fecha >= '" & fd & "' and Fecha <= '" & fh & "' "
      
  If CheckCedNom.value = vbChecked Then
    s1 = eCedNom.Text
    If InStr(s1, "*") > 0 Then Mid(s1, InStr(s1, "*"), 1) = "%"
    If InStr(s1, "*") > 0 Then Mid(s1, InStr(s1, "*"), 1) = "%"
       
    If Combo1.Text = "CEDULA" Then
      s = "select * from Diario where Fecha >= '" & fd & "' and Fecha <= '" & fh & "' and cedula = '" & Trim(s1) & "' "
    Else
      s = "select * from Diario where Fecha >= '" & fd & "' and Fecha <= '" & fh & "' and nombre like '" & Trim(s1) & "' "
    End If
  End If
  
  If CheckCli.value = vbChecked Then
    s2 = Mid(cCP.Text, 1, 6)
    s3 = Trim(Mid(cSC.Text, 1, 6))
    
    If s3 = "" Then s3 = "0"
    
    s = s & " and cliente = " & s2 & " and subcliente = " & s3 & " "
  End If
 
  s = s & " order by id"
    
  rM.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  
  FG.Clear
  FG.Rows = 2
  FG.Cols = 9 'rM.Fields.Count
  
  FG.TextMatrix(0, 0) = "ID"
  FG.TextMatrix(0, 1) = "FECHA"
  FG.TextMatrix(0, 2) = "CLIENTE"
  FG.TextMatrix(0, 3) = "SUBCLIENTE"
  FG.TextMatrix(0, 4) = "CEDULA"
  FG.TextMatrix(0, 5) = "NOMBRE"
  FG.TextMatrix(0, 6) = "OBSERVACIONES"
  FG.TextMatrix(0, 7) = "MONTO Bs."
  FG.TextMatrix(0, 8) = "Cancelado"
  
  FG.ColWidth(0) = 400 'ID
  FG.ColWidth(1) = 1000 'Fecha y Hora
  FG.ColWidth(2) = 2000 'Cliente
  FG.ColWidth(3) = 2000 'Subcliente
  FG.ColWidth(4) = 900 'Cedula
  FG.ColWidth(5) = 2100 'Nombre
  FG.ColWidth(6) = 5400 'Observaciones
  FG.ColWidth(7) = 1000 'Observaciones
  FG.ColWidth(8) = 500 'Cancelado (S/N)
  
  TM = 0#
  
  f = 1
  Do While Not rM.EOF
    FG.TextMatrix(f, 0) = rM.Fields("ID").value
    FG.TextMatrix(f, 1) = Trim(Modulo.FechaNormal(rM.Fields("Fecha").value) & " " & rM.Fields("Hora").value): FG.Row = f: FG.Col = 1: FG.CellAlignment = flexAlignLeftCenter
    FG.TextMatrix(f, 2) = Zeros(rM.Fields("Cliente").value, 6): FG.Row = f: FG.Col = 2: FG.CellAlignment = flexAlignLeftCenter
    
      s = Modulo.DBValorStr("clientes", "codigo", CStr(rM.Fields("Cliente").value), "nombre")
      If s <> "" Then
        FG.TextMatrix(f, 2) = Zeros(rM.Fields("Cliente").value, 6) & " " & s
      End If
    
    FG.TextMatrix(f, 3) = Zeros(rM.Fields("Subcliente").value, 6)
    
      l1 = rM.Fields("CLIENTE").value
      l2 = rM.Fields("SUBCLIENTE").value
    
      If l2 <> 0 Then
        If rS.RecordCount > 0 Then
          rS.MoveFirst
          bExiste = False
          Do While Not rS.EOF And Not bExiste
            If rS.Fields("cliente").value = l1 And _
               rS.Fields("id").value = l2 Then
               FG.TextMatrix(f, 3) = Zeros(l2, 6) & " " & Trim(rS.Fields("NOMBRE").value)
               FG.Row = f: FG.Col = 3: FG.CellAlignment = flexAlignLeftCenter
               bExiste = True
            End If
            rS.MoveNext
          Loop
        End If
      Else
        FG.TextMatrix(f, 3) = "-": FG.Row = f: FG.Col = 3: FG.CellAlignment = flexAlignLeftCenter
      End If
    
    
    FG.TextMatrix(f, 4) = Trim(rM.Fields("Cedula").value): FG.Row = f: FG.Col = 4: FG.CellAlignment = flexAlignLeftCenter
    FG.TextMatrix(f, 5) = ""
    
    s = Modulo.DBValorStr(Trim(rM.Fields("Tabla").value), "cedula", Trim(rM.Fields("Cedula").value), "nombre")
    If s <> "" Then FG.TextMatrix(f, 5) = s
    
    FG.TextMatrix(f, 6) = rM.Fields("Observaciones").value
    
    FG.TextMatrix(f, 7) = Format(rM.Fields("Monto").value, "#,0.00")
    Select Case UCase(rM.Fields("Pago").value)
       Case "N" ' Falta Pago
        EsFaltaPago = True
        FG.TextMatrix(f, 8) = "NO"
       Case "S" 'Pago
        EsFaltaPago = False
        FG.TextMatrix(f, 8) = "SI"
    End Select
    
    
    TM = TM + IIf(IsNull(rM.Fields("Monto").value), 0, rM.Fields("Monto").value)
    
    rM.MoveNext
    If Not rM.EOF Then
      f = f + 1
      FG.Rows = FG.Rows + 1
      
      Label5.Caption = "Items en este Reporte: " & CStr(FG.Rows - 1)
    End If
  Loop
  
  Label4.Caption = Format(TM, "#,0.00")
    
  rS.Close
  Set rS = Nothing
  
  rM.Close
  Set rM = Nothing
  
  
  'FormatearDG
  
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

Private Sub cCP_Change()
  Cargar_SubClientes
End Sub

Private Sub cCP_Click()
  Cargar_SubClientes
End Sub

Private Sub CheckCedNom_Click()
  If CheckCedNom.value = vbUnchecked Then
    Combo1.Enabled = False
    eCedNom.Enabled = False
    eCedNom.Text = ""
  Else
    Combo1.Enabled = True
    eCedNom.Enabled = True
    eCedNom.Text = ""
    eCedNom.SetFocus
  End If
End Sub

Private Sub CheckCli_Click()
  If CheckCli.value = vbChecked Then
    cCP.Enabled = True
    cSC.Enabled = True
    bBuscarCP.Enabled = True
  Else
    cCP.Enabled = False
    cSC.Enabled = False
    bBuscarCP.Enabled = False
    'cCP.SetFocus
  End If
End Sub

Private Sub FormatearCedula()
  Dim s As String
  s = Trim(eCedNom.Text)
  If s <> "" Then
    Formatear_Cedula s
    eCedNom.Text = s
  End If
End Sub

Private Sub eCedNom_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    FormatearCedula
  End If
End Sub

Private Sub eCednom_LostFocus()
  If Combo1.Text = "CEDULA" Then FormatearCedula Else eCedNom.Text = UCase(eCedNom.Text)
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
  Do While Not RSubClientes.EOF
    s = Zeros(RSubClientes.Fields("id").value, 6) & " : " & Trim(RSubClientes.Fields("nombre").value)
    cSC.AddItem s
    RSubClientes.MoveNext
  Loop
  RSubClientes.Close
  Set RSubClientes = Nothing
  
  If cSC.ListCount > 0 Then cSC.ListIndex = 0
  
End Sub

Private Sub FG_DblClick()
  Dim i As Integer
  Dim r As New ADODB.Recordset
  Dim s As String, sID As String, s2 As String
  Dim dT As Double, sLoc As String
  
  i = FG.Row
  If i >= 1 Then
  
    Load fMov2
    fMov2.List1.Clear
  
    dT = 0#
    sID = FG.TextMatrix(i, 0)
    sLoc = Modulo.Localizador_Por_ID("Diario", sID)
       
    s = "select * from DiarioDetalle where Localizador = '" & sLoc & "' order by id"
    r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    
    Do While Not r.EOF
      
      s2 = Modulo.Producto_DESCRIPCION(Trim(r.Fields("codigoproducto").value))
      If Len(s2) > 30 Then s2 = Mid(s2, 1, 30) Else s2 = Modulo.BlancosDER(s2, 30)
      
      s = r.Fields("codigoproducto").value & " " & _
          s2 & " " & _
          Modulo.BlancosIZQ(CStr(r.Fields("cantidad").value), 12) & " " & _
          Modulo.BlancosIZQ(Format(r.Fields("precio").value, "#,0.00"), 11) & " " & _
          Modulo.BlancosIZQ(Format(r.Fields("subtotal").value, "#,0.00"), 12) & "       " & _
          IIf(r.Fields("Entregado").value = "S", "SI", "NO")
          
      dT = dT + r.Fields("subtotal").value

      fMov2.List1.AddItem s
      fMov2.List2.AddItem r!ID
      r.MoveNext
    Loop
    
    r.Close
    Set r = Nothing
    fMov2.Localizador = sLoc
    fMov2.lTotal.Caption = Format(dT, "#,0.00")
    If FG.TextMatrix(i, 8) = "NO" Then fMov2.cmdPagar.Visible = True
    fMov2.txtObservaciones.Text = FG.TextMatrix(i, 6)
    fMov2.Show vbModal
    
  End If
    
End Sub

Private Sub Form_Load()
  eDesde.value = Date
  eHasta.value = Date
  
  eCedNom.Text = ""
  CheckCedNom.value = vbUnchecked
  eCedNom.Text = ""
  eCedNom.Enabled = False
  Combo1.Enabled = False
  
  
  Combo1.Clear
  Combo1.AddItem "CEDULA"
  'Combo1.AddItem "NOMBRE"
  Combo1.ListIndex = 0
  
  cCP.Enabled = False
  cSC.Enabled = False
  bBuscarCP.Enabled = False
  CheckCli.value = vbuncheked
  
  Label4.Caption = "0,00"
  
  Cargar_Clientes
  
  
End Sub

Private Sub Form_Resize()
  Frame1.width = Me.width - 130
  FG.width = Me.width - 250
End Sub
