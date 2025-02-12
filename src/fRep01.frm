VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form fRep01 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Carnets Entregados"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8340
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   500
      Left            =   3720
      Picture         =   "fRep01.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Height          =   2715
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8325
      Begin VB.ComboBox cmbOpciones 
         Height          =   315
         ItemData        =   "fRep01.frx":058A
         Left            =   1710
         List            =   "fRep01.frx":0597
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1020
         Width           =   2535
      End
      Begin VB.CommandButton bAceptar 
         Caption         =   "Generar"
         Height          =   500
         Left            =   6510
         Picture         =   "fRep01.frx":05B9
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1410
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin MSComCtl2.DTPicker eDesde 
         Height          =   345
         Left            =   1710
         TabIndex        =   3
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   62324737
         CurrentDate     =   40053
      End
      Begin VB.CommandButton bBuscarCP 
         Height          =   345
         Left            =   7590
         Picture         =   "fRep01.frx":0B43
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   5715
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   5715
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Cliente Principal posee Tabla de Personas"
         Height          =   225
         Left            =   1740
         TabIndex        =   2
         Top             =   1500
         Visible         =   0   'False
         Width           =   3405
      End
      Begin MSComCtl2.DTPicker eHasta 
         Height          =   345
         Left            =   4470
         TabIndex        =   4
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   62324737
         CurrentDate     =   40053
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
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
         Left            =   3750
         TabIndex        =   12
         Top             =   1860
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
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
         Left            =   990
         TabIndex        =   11
         Top             =   1860
         Width           =   660
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
         Left            =   180
         TabIndex        =   10
         Top             =   240
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
         Left            =   600
         TabIndex        =   9
         Top             =   660
         Width           =   1080
      End
   End
End
Attribute VB_Name = "fRep01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RClientes As New ADODB.Recordset
Dim RSubClientes As New ADODB.Recordset

Private Sub Cargar_Clientes()
  'Dim r As New ADODB.Recordset
  Dim s As String
  Dim l As Integer
  Dim sCod As String
  
  If RClientes.State <> adStateClosed Then RClientes.Close
  
  s = "SELECT * FROM Clientes ORDER BY Codigo"
  
  RClientes.Open s, DBConexionSQL, adOpenKeyset, adLockReadOnly
  
  cCP.Clear
  
  Do While Not RClientes.EOF
    s = Zeros(RClientes.Fields("codigo").value, 6) & " : " & Trim(RClientes.Fields("nombre").value)
    cCP.AddItem s
    RClientes.MoveNext
  Loop
  
  RClientes.Close
End Sub


Private Sub Cargar_SubClientes()
  'Dim r As New ADODB.Recordset
  Dim s As String
  Dim l As Integer
  Dim sCod As String
  
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
  
  If cSC.ListCount > 0 Then cSC.ListIndex = 0
  cSC.ListIndex = -1
End Sub

Private Sub bAceptar_Click()
  Dim sCliente As String, sSubCliente As String, sDesde As String, sHasta As String
  
  Dim s As String
  Dim s2 As String
  Dim sT1 As String, sT2 As String
  Dim r As New ADODB.Recordset
  Dim k As Long
  
  Dim R_I As Long
  Dim R_CI As String
  Dim R_Nombre As String
  Dim R_Cargo As String
  Dim R_Impreso As String
  
  Dim SqlTxt As String
On Error GoTo falla
  If eDesde.value > eHasta.value Then
    MsgBox "Las Fechas son inválidas, Revise...", vbCritical, "Información"
    Exit Sub
  End If
  
  If cCP.Text = "" Then
     MsgBox "Debe seleccionar un Cliente", vbExclamation
     Exit Sub
  End If
  Load fMensaje
  fMensaje.Label1.Caption = "Generando Reporte, Espere..."
  fMensaje.Show
  DoEvents
  
  'Preparar la Tabla Temporal del Reporte:
  s = "delete from Reporte01 where estacion = " & Val(Modulo.ESTACION)
  Modulo.ExecSQL s
    
  sCliente = cCP.Text
  sSubCliente = cSC.Text
  sDesde = Format(eDesde.value, "yyyymmdd")
  sHasta = Format(eHasta.value, "yyyymmdd")
    
  'If sSubCliente = "" Or sSubCliente = "-" Then sSubCliente = ""
  
  sT1 = Modulo.La_Tabla_Actual_Personas(sCliente, sSubCliente)
  sT2 = "H" & sT1
  
  '-- I : Buscar en la Historica de resguardo:
  Select Case cmbOpciones.Text
     Case "IMPRESOS"
          s = "select * from [" & sT2 & "] where " & _
          "Fecha between ('" & sDesde & " 00:00:00') and (" & _
          "'" & sHasta & " 23:59:59') order by fecha"
     'Case "NO IMPRESOS"
     '     s = "select * from [" & sT2 & "] where " & _
     '     "Fecha ='' "
          ''s = ""
     'Case "TODOS"
     '     s = "select * from [" & sT2 & "] WHERE FECHA IS NOT NULL Order by Fecha"
     
 
  
           SqlTxt = s
            ''"Fecha <= '" & sHasta & "' and " & _
           ''"Contador > 0 order by Cedula, Fecha"
           ''r.Close
           r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
           k = 0
  
           R_I = 0
  
     

          Do While Not r.EOF
          ''If Not IsNull(r.Fields("CONTADOR").value) And _
             Not IsNull(r.Fields("FECHA").value) Then
       
             k = k + Val(r.Fields("CONTADOR").value)
      
             R_I = R_I + 1
             R_CI = Trim(r.Fields("CEDULA").value)
             R_Nombre = Trim(r.Fields("NOMBRE").value)
             R_Cargo = Trim(r.Fields("CARGO").value)
             R_Impreso = Format(r.Fields("Fecha").value, "dd/mm/yyyy")
      
             If Len(R_Nombre) > 50 Then R_Nombre = Mid(R_Nombre, 1, 50)
             If Len(R_Cargo) > 50 Then R_Cargo = Mid(R_Cargo, 1, 50)
      
             s = "insert into Reporte01 (estacion,numero,cedula,nombre,cargo,impreso) values (" & _
                 Val(Modulo.ESTACION) & ",'" & Zeros(R_I, 4) & "','" & _
                 R_CI & "','" & R_Nombre & "','" & R_Cargo & "','" & R_Impreso & "')"
           
             Modulo.ExecSQL s
         
            ''End If
             r.MoveNext
          Loop
  
          r.Close
  End Select
  '-- II : Buscar en la actual de personas:
If cmbOpciones.Text <> "IMPRESOS" Then
  Select Case cmbOpciones.Text
     'Case "IMPRESOS"
     '   s = "select * from [" & sT1 & "] where " & _
     '    "Fecha >= '" & sDesde & "' and " & _
     '    "Fecha <= '" & sHasta & "' and " & _
     '    "Contador > 0 order by Cedula, Fecha"
     Case "NO IMPRESOS"
          s = "select * from [" & sT1 & "] where " & _
          "Fecha IS NULL"
     Case "TODOS"
        ' no impresos
          s = "select * from [" & sT1 & "] where " & _
          "Fecha IS NULL"
        ' impresos
          s2 = "select * from [" & sT2 & "] where " & _
          "Fecha between ('" & sDesde & " 00:00:00') and (" & _
          "'" & sHasta & " 23:59:59') order by fecha"
  End Select
  
  SqlTxt = s
  
  
  
      
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  Do While Not r.EOF
    'If Not IsNull(r.Fields("CONTADOR").value) And _
       Not IsNull(r.Fields("FECHA").value) Then
       
      If IsNull(r.Fields("CONTADOR").value) = False Then
         k = k + Val(r.Fields("CONTADOR").value)
      Else
         k = k + 0
      End If
      
      R_I = R_I + 1
      R_CI = Trim(r.Fields("CEDULA").value)
      R_Nombre = Trim(r.Fields("NOMBRE").value)
      R_Cargo = Trim(r.Fields("CARGO").value)
      R_Impreso = Format(r.Fields("FECHA").value, "dd/mm/yyyy")
      
      If Len(R_Nombre) > 50 Then R_Nombre = Mid(R_Nombre, 1, 50)
      If Len(R_Cargo) > 50 Then R_Cargo = Mid(R_Cargo, 1, 50)
     
      
      s = "insert into Reporte01 (estacion,numero,cedula,nombre,cargo,impreso) values (" & _
           Val(Modulo.ESTACION) & ",'" & Zeros(R_I, 4) & "','" & _
           R_CI & "','" & R_Nombre & "','" & R_Cargo & "','" & R_Impreso & "')"
           
      Modulo.ExecSQL s
      
    ''End If
    r.MoveNext
  Loop
  
  r.Close
  
  Set r = Nothing
  
 If cmbOpciones.Text = "TODOS" Then
  
  SqlTxt = s2
  
  
  
      
  r.Open s2, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  Do While Not r.EOF
    'If Not IsNull(r.Fields("CONTADOR").value) And _
       Not IsNull(r.Fields("FECHA").value) Then
       
      If IsNull(r.Fields("CONTADOR").value) = False Then
         k = k + Val(r.Fields("CONTADOR").value)
      Else
         k = k + 0
      End If
      
      R_I = R_I + 1
      R_CI = Trim(r.Fields("CEDULA").value)
      R_Nombre = Trim(r.Fields("NOMBRE").value)
      R_Cargo = Trim(r.Fields("CARGO").value)
      R_Impreso = Format(r.Fields("creacion").value, "dd/mm/yyyy")
      
      If Len(R_Nombre) > 50 Then R_Nombre = Mid(R_Nombre, 1, 50)
      If Len(R_Cargo) > 50 Then R_Cargo = Mid(R_Cargo, 1, 50)
     
      
      s = "insert into Reporte01 (estacion,numero,cedula,nombre,cargo,impreso) values (" & _
           Val(Modulo.ESTACION) & ",'" & Zeros(R_I, 4) & "','" & _
           R_CI & "','" & R_Nombre & "','" & R_Cargo & "','" & R_Impreso & "')"
           
      Modulo.ExecSQL s
      
    ''End If
    r.MoveNext
  Loop
  
  r.Close
  
  Set r = Nothing
 End If
  
  
  
  
End If
  Unload fMensaje
  Set r = DBConexionSQL.Execute("Select * from Reporte01 where estacion=" & Modulo.ESTACION)
  If r.EOF = True Then
     MsgBox "No Existe Información con estos Parámetros...", vbCritical, "Información"
     Unload fMensaje
     Exit Sub
  End If
  'If (k <= 0 And cmbOpciones.Text <> "NO IMPRESOS") And cmbOpciones.Text <> "TODOS" Then
  '  MsgBox "No Existe Información con estos Parámetros...", vbCritical, "Información"
  '  Unload fMensaje
  'Else
    
    Load frmPreviewCarnetsImpresos
    frmPreviewCarnetsImpresos.sCargarReporteCarnetsImpresos Trim(Mid(cCP.Text, 9)), Trim(Mid(cSC.Text, 9)), eDesde.value, eHasta.value, SqlTxt
    frmPreviewCarnetsImpresos.Show vbModal
             'DataEnvironment1.Connection1.Close
             'DataEnvironment1.Connection1.ConnectionString = Modulo.DBConexionSQL.ConnectionString
    
   ' If DataEnvironment1.Connection1.State = adStateClosed Then
   '   DataEnvironment1.Connection1.Open
   ' End If
   ' DataReport1.Sections("Sección2").Controls("Etiqueta2").Caption = Format(Date, "dd/mm/yyyy")
   ' DataReport1.Sections("Sección2").Controls("Etiqueta9").Caption = cCP.Text
   ' DataReport1.Sections("Sección2").Controls("Etiqueta10").Caption = cSC.Text
   '
   ' DataReport1.Sections("Sección2").Controls("Etiqueta7").Visible = False
   ' DataReport1.Sections("Sección1").Controls("txtimpreso").Visible = False
   '
   ' If Check2.Value = vbChecked Then
   '   DataReport1.Sections("Sección2").Controls("Etiqueta7").Visible = True
   '   DataReport1.Sections("Sección1").Controls("txtimpreso").Visible = True
   ' End If
   '
   ' DataReport1.Show
  'End If
  
falla:
  If Err.Number <> 0 Then MsgBox Err.Number & "::" & Err.Description, vbCritical
  
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
  Unload Me
End Sub

Private Sub cCP_Click()
  If Trim(cCP.Text) <> "" Then Cargar_SubClientes
End Sub

Private Sub Check1_Click()
  If Check1.value = vbChecked Then
    cSC.Clear
    'Cargar_Tablas_Personas
  Else
    cCP_Click
  End If
End Sub

Private Sub Form_Load()
  Cargar_Clientes
  Cargar_SubClientes
  eDesde.value = Now
  eHasta.value = Now
  cmbOpciones.Text = "IMPRESOS"
End Sub

