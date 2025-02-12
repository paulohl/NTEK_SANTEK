VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fBuscarPersona 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Persona"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   13290
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   5415
      Left            =   0
      TabIndex        =   9
      Top             =   1110
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   9551
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin VB.CommandButton bAceptar 
      Caption         =   "Aceptar"
      Height          =   500
      Left            =   5610
      Picture         =   "fBuscarPersona.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6720
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   500
      Left            =   6780
      Picture         =   "fBuscarPersona.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13245
      Begin VB.CommandButton bBuscar 
         Height          =   345
         Left            =   5910
         Picture         =   "fBuscarPersona.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   510
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   5
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
      Begin VB.Label lBuscando 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Buscando en Base de Datos, Espere..."
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7860
         TabIndex        =   10
         Top             =   540
         Visible         =   0   'False
         Width           =   3195
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
      Left            =   9870
      TabIndex        =   4
      Top             =   6570
      Width           =   45
   End
End
Attribute VB_Name = "fBuscarPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EncabezadoFG As String

Private Sub BuscarValor()
  Dim f As Integer, c As Integer
  Dim Hay As Boolean
  Dim s As String, sTabla As String, sCliente As String
  Dim sSubCliente As String
  Dim sSQL As String
  Dim sBuscar As String
  Dim i As Integer
  Dim aCols() As Long
  Dim sElCliente As String, sElSubCliente As String
  
  
  Dim rc As New ADODB.Recordset
  Dim rS As New ADODB.Recordset
  
  Dim rP As New ADODB.Recordset
  Dim rT As New ADODB.Recordset
      
  rc.Open "select * from clientes", Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  rS.Open "select * from subclientes", Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  
  sBuscar = Trim(eBuscar.Text)
 
  If sBuscar <> "" Then
  
    FG.Clear
    FG.Rows = 2
    FG.Cols = 7
    
    FG.TextMatrix(0, 0) = "ID"
    FG.TextMatrix(0, 1) = "CLIENTE"
    FG.TextMatrix(0, 2) = "SUBCLIENTE"
    FG.TextMatrix(0, 3) = "CEDULA"
    FG.TextMatrix(0, 4) = "NOMBRE"
    FG.TextMatrix(0, 5) = "CARGO"
    
    FG.TextMatrix(0, 6) = "TABLA"
    
    'FG.ColWidth(0) = 2000
    'FG.ColWidth(1) = 1000
    'FG.ColWidth(2) = 300
    'FG.ColWidth(3) = 1000
    'FG.ColWidth(4) = 500
    
    lBuscando.Visible = True
    
    'Load fMensaje
    'fMensaje.Label1.Caption = "Buscando en Base de Datos, Espere..."
    'fMensaje.Show
    DoEvents
              
    s = "select * from personas order by cliente, subcliente"
    rP.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    f = 1
    Do While Not rP.EOF
    
      lBuscando.Caption = "Buscando en Base de Datos [" & CStr(f) & "/" & CStr(rP.RecordCount) & "], Espere..."
          
      'fMensaje.Label1.Caption = "Buscando en Base de Datos [" & CStr(f) & "/" & CStr(rP.RecordCount) & "], Espere..."
      'fMensaje.Show
      DoEvents
    
      sCliente = Zeros(rP.Fields("Cliente").Value, 6)
      sSubCliente = Zeros(rP.Fields("SubCliente").Value, 6)
      
      'FG.TextMatrix(f, 0) = sCliente & " "
      
      sElCliente = sCliente & " "
      
      If rc.RecordCount > 0 Then
        rc.MoveFirst
        rc.Find "codigo = " & sCliente & " "
        If Not rc.EOF Then
          'FG.TextMatrix(f, 0) = FG.TextMatrix(f, 0) & Trim(rC.Fields("nombre").Value)
          sElCliente = sElCliente & Trim(rc.Fields("nombre").Value)
        End If
      End If
      
      'FG.TextMatrix(f, 1) = sSubCliente & " "
      
      If rP.Fields("SubCliente").Value <= 0 Then
        sElSubCliente = "-"
      Else
      
        sElSubCliente = sSubCliente & " "
      
        If rS.RecordCount > 0 Then
          rS.MoveFirst
          rS.Filter = "cliente = " & sCliente & " and id = " & sSubCliente & " "
          If Not rS.EOF Then
            'FG.TextMatrix(f, 1) = FG.TextMatrix(f, 1) & Trim(rS.Fields("nombre").Value)
            sElSubCliente = sElSubCliente & Trim(rS.Fields("nombre").Value)
          End If
          rS.Filter = adFilterNone
        End If
        
      End If
      
      
      sTabla = Trim(rP.Fields("Tabla").Value)
      If InStr(sBuscar, "*") > 0 Then Mid(sBuscar, InStr(sBuscar, "*"), 1) = "%"
      If InStr(sBuscar, "*") > 0 Then Mid(sBuscar, InStr(sBuscar, "*"), 1) = "%"
            
      If Combo1.Text = "CEDULA" Then
        s = "select * from [" & sTabla & "] where cedula = '" & sBuscar & "'"
      Else
        s = "select * from [" & sTabla & "] where nombre like '" & sBuscar & "'"
      End If
      
      rT.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
      If Not rT.EOF Then
      
        'Mostrar en FlexGrid:
        FG.TextMatrix(f, 0) = CStr(rT.Fields("ID").Value): FG.CellAlignment = flexAlignLeftCenter
        FG.TextMatrix(f, 1) = sElCliente: FG.Row = f: FG.Col = 0: FG.CellAlignment = flexAlignLeftCenter
        FG.TextMatrix(f, 2) = sElSubCliente: FG.Row = f: FG.Col = 1: FG.CellAlignment = flexAlignLeftCenter
      
        FG.TextMatrix(f, 3) = IIf(IsNull(Trim(rT.Fields("CEDULA").Value)), "", Trim(rT.Fields("CEDULA").Value)): FG.Row = f: FG.Col = 2: FG.CellAlignment = flexAlignLeftCenter
        FG.TextMatrix(f, 4) = IIf(IsNull(Trim(rT.Fields("NOMBRE").Value)), "", Trim(rT.Fields("NOMBRE").Value)): FG.Row = f: FG.Col = 3: FG.CellAlignment = flexAlignLeftCenter
        FG.TextMatrix(f, 5) = IIf(IsNull(Trim(rT.Fields("CARGO").Value)), "", Trim(rT.Fields("CARGO").Value)): FG.Row = f: FG.Col = 4: FG.CellAlignment = flexAlignLeftCenter
        
        FG.TextMatrix(f, 6) = sTabla
        
        f = f + 1
        FG.Rows = FG.Rows + 1
        
      End If
      rT.Close
      
      rP.MoveNext
      
    Loop
    
    'Unload fMensaje
    
    rP.Close
    
    rc.Close
    rS.Close
    
    lBuscando.Visible = False
    
    'rT.Close
  
  End If
  
  Set rP = Nothing
  Set rc = Nothing
  Set rS = Nothing
  Set rT = Nothing
  
  
  
End Sub


Private Sub bAceptar_Click()
  Dim sC As String, sSC As String
  Dim sTablaActiva As String
  Dim sTablaBusqueda As String
  
  If FG.Row >= 1 Then
    
    If Trim(fMov.GRID.TextMatrix(fMov.GRID.Row, 1)) = "" Then
    
      sC = FG.TextMatrix(FG.Row, 1)
      sSC = FG.TextMatrix(FG.Row, 2)
      
      sTablaActiva = Modulo.La_Tabla_Actual_Personas(sC, sSC)
      sTablaBusqueda = FG.TextMatrix(FG.Row, 6)
      
      If sTablaActiva <> sTablaBusqueda Then
        
        MsgBox "La Persona que ha Seleccionado está Registrada " & vbCrLf & _
               "en una Tabla Histórica del Cliente...  ", vbCritical, "Información"
               
        Exit Sub
               
      Else
           
        fMov.GRID.TextMatrix(fMov.GRID.Row, 3) = FG.TextMatrix(FG.Row, 1)
        fMov.GRID.TextMatrix(fMov.GRID.Row, 4) = FG.TextMatrix(FG.Row, 2)
        fMov.GRID.TextMatrix(fMov.GRID.Row, 5) = FG.TextMatrix(FG.Row, 3)
      
      End If
      
      'Modulo.vTemporal1 = FG.TextMatrix(FG.Row, 2)
      Modulo.fModalResult = Modulo.fModalResultOK
      Unload Me
    End If
  End If
  
  'If Not Adodc1.Recordset.EOF Then
  '  Modulo.vTemporal1 = Zeros(Adodc1.Recordset.Fields("codigo").Value, 6)
  '  Unload Me
  'End If
End Sub

Private Sub bBuscar_Click()
  FormatearCedula
  BuscarValor
End Sub

Private Sub bCancelar_Click()
  Modulo.fModalResult = Modulo.fModalResultCANCEL
  Unload Me
End Sub

Private Sub DataGrid1_DblClick()
  If Not Adodc1.Recordset.EOF Then
    Modulo.vTemporal1 = Zeros(Adodc1.Recordset.Fields("codigo").Value, 6)
    Unload Me
  End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    If Not Adodc1.Recordset.EOF Then
      Modulo.vTemporal1 = Zeros(Adodc1.Recordset.Fields("codigo").Value, 6)
      Unload Me
    End If
  End If
End Sub

Private Sub eBuscar_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnMayusculas eBuscar
    FormatearCedula
    BuscarValor
  End If
End Sub


Private Sub FormatearCedula()
  Dim s As String
  s = Trim(eBuscar.Text)
  If s <> "" Then
    Formatear_Cedula s
    eBuscar.Text = s
  End If
End Sub



Private Sub FG_DblClick()
  Dim c As Integer
  
'  If FG.Rows > 1 Then
'    If FG.Row >= 1 Then
'      c = 0
'      Do While c < FG.Cols
'        FG.Col = c
'        FG.CellBackColor = vbBlue
'        FG.CellForeColor = vbYellow
'        c = c + 1
'      Loop
          
'      bAceptar_Click
'    End If
'  End If
End Sub

Private Sub eBuscar_LostFocus()
  EnMayusculas eBuscar
  FormatearCedula
End Sub

Private Sub Form_Load()
  lBuscando.Visible = False
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


