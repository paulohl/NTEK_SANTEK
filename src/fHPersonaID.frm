VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fHPersonaID 
   Caption         =   "Marcar Carnets (ID) para Imprimir"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9100
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFC0FF&
         Height          =   645
         Left            =   4950
         ScaleHeight     =   585
         ScaleWidth      =   1035
         TabIndex        =   13
         Top             =   8460
         Width           =   1100
         Begin VB.CommandButton bCard5 
            Caption         =   "CARD-5"
            Height          =   500
            Left            =   60
            Picture         =   "fHPersonaID.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   900
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Inicializar « Marca » para Imprimir Despues."
         Height          =   285
         Left            =   510
         TabIndex        =   11
         Top             =   8520
         Value           =   1  'Checked
         Width           =   3675
      End
      Begin VB.Frame Frame2 
         Height          =   2205
         Left            =   90
         TabIndex        =   3
         Top             =   150
         Width           =   8925
         Begin VB.CheckBox Check2 
            Caption         =   "Cliente Principal posee Tabla de Personas"
            Height          =   225
            Left            =   2430
            TabIndex        =   12
            Top             =   1110
            Width           =   3825
         End
         Begin VB.TextBox eCed 
            Height          =   345
            Left            =   3390
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   1590
            Width           =   1935
         End
         Begin VB.ComboBox cSC 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   720
            Width           =   4200
         End
         Begin VB.ComboBox cCP 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   300
            Width           =   4200
         End
         Begin VB.CommandButton bBuscarCP 
            Height          =   345
            Left            =   6810
            Picture         =   "fHPersonaID.frx":058A
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Indique el Número de Cédula:"
            Height          =   195
            Left            =   1200
            TabIndex        =   10
            Top             =   1650
            Width           =   2100
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Cliente:"
            Height          =   195
            Left            =   1530
            TabIndex        =   8
            Top             =   780
            Width           =   855
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Cliente Principal:"
            Height          =   195
            Left            =   1230
            TabIndex        =   7
            Top             =   360
            Width           =   1170
         End
      End
      Begin MSFlexGridLib.MSFlexGrid FG 
         Height          =   6105
         Left            =   90
         TabIndex        =   2
         Top             =   2400
         Width           =   8990
         _ExtentX        =   15849
         _ExtentY        =   10769
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         FormatString    =   "Nº    | CÉDULA                 |  ID              | OBSERVACION                                                    "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton bCancelar 
         Caption         =   "Cancelar"
         Height          =   500
         Left            =   7530
         Picture         =   "fHPersonaID.frx":068C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   8520
         UseMaskColor    =   -1  'True
         Width           =   900
      End
   End
End
Attribute VB_Name = "fHPersonaID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cFG = "Nº    | CÉDULA                 |  ID              | OBSERVACION                                                    "

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



Private Sub SeleccionarCampo(eTexto As TextBox)
  eTexto.SelStart = 0
  eTexto.SelLength = Len(eTexto.Text)
End Sub

Private Sub bCard5_Click()
  Dim SC As String, sSC As String
  Dim sRuta As String, sCard5 As String
  Dim s As String
  Dim i As Integer
  
  Load fMensaje
  fMensaje.Caption = "Preparando para Ejecutar «CARD-5», Espere..."
  fMensaje.Show
  DoEvents
  
  SC = Mid(cCP.Text, 1, 6)
  sSC = Mid(cSC.Text, 1, 6)
    
  If SC <> "" Then
    If sSC = "" Or sSC = "-" Then sSC = "0" Else sSC = Mid(sSC, 1, 6)
    Auditar_Fotos Modulo.La_Tabla_Actual_Personas(SC, sSC)
  End If
        
  Dim sOri As String
  Dim sDes As String

  sOri = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
  sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
  
  SC = Trim(Mid(cCP.Text, 9))
  sSC = Trim(Mid(cSC.Text, 9))
       
  If Trim(sSC) = "" Or Trim(sSC) = "-" Then
    sRuta = sDes & "\" & SC & "\CARNET" & "\" & SC & ".car"
  Else
    'sSC = Trim(Mid(sSC, 7))
    sRuta = sDes & "\" & SC & "\" & sSC & "\CARNET" & "\" & sSC & ".car"
  End If
  
  If Dir(sRuta) = "" Then
    'probar con otro nombre antiguo: "BASE"
    If Trim(sSC) = "" Or Trim(sSC) = "-" Then
      sRuta = sDes & "\" & SC & "\CARNET" & "\BASE " & SC & ".car"
    Else
      'sSC = Trim(Mid(sSC, 7))
      sRuta = sDes & "\" & SC & "\" & sSC & "\CARNET" & "\BASE " & sSC & ".car"
    End If
  End If
  
    
  

  s = GetSetting(APPNAME, "Opciones", "RutaCard5", "")

  sCard5 = s & " " & Chr(34) & sRuta & Chr(34)

  If Shell(sCard5, vbMaximizedFocus) = 0# Then
    MsgBox "Error: No se pudo Iniciar Card-5" & vbCrLf & CStr(Err.Number) & ":" & Err.Description, "Información"
  End If
  
  Unload fMensaje

End Sub

Private Sub cCP_Click()
  If Trim(cCP.Text) <> "" Then Cargar_SubClientes
  'Cargar_Tablas_Personas
  FG.Clear
  FG.Rows = 2
  FG.FormatString = cFG
  
End Sub

Private Sub Check2_Click()
  If Check2.Value = vbChecked Then
    cSC.Clear
    'Cargar_Tablas_Personas
  Else
    cCP_Click
  End If
End Sub

Private Sub cSC_Click()
  'Cargar_Tablas_Personas
End Sub

Private Sub eCed_GotFocus()
  Color_Fila_Disponible
  'SeleccionarCampo eCed
  eCed.Text = ""

End Sub

Private Sub eCed_KeyPress(KeyAscii As Integer)
  Dim i As Integer, j As Integer
  Dim r As New ADODB.Recordset, sOb As String
  Dim s As String
  
  i = 1
  j = -1
  If KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
    KeyAscii = 0
    Exit Sub
  End If
    
  If KeyAscii = vbKeyReturn Then
  
    KeyAscii = 0
  
    If Existe_Valor_FG(Trim(eCed.Text), 2) Then
      MsgBox "Cédula ya Existe en el Listado...", vbCritical, "Información"
      Exit Sub
    End If
  
    Do While (i < FG.Rows) And (j = -1)
      If Trim(FG.TextMatrix(i, 2)) = "" Then j = i
      i = i + 1
    Loop
    If j <> -1 Then
      
      Dim s1 As String, s2 As String
      Dim st As String, sCed As String
      
      sCed = eCed.Text
      
      If IsNumeric(eCed.Text) Then
        sCed = Format(CLng(eCed.Text), "#,0")
      End If
        
  
      s1 = Trim(cCP.Text)
      s2 = Trim(cSC.Text)
      If s1 <> "" Then
        st = Modulo.La_Tabla_Actual_Personas(s1, s2)
        If st <> "" Then
          s = "select id from [" & st & "] where cedula = '" & Trim(sCed) & "'"
          r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
          If r.EOF Then
            MsgBox "Cedula [" & eCed.Text & "] No Existe...", vbCritical, "Información"
          Else
          
            If Modulo.Existe_Valor_En_Columna(FG, 1, sCed) Then
            
              MsgBox "Ya Existe la Cédula " & sCed & " en el Listado...", vbCritical, "Información"
              
            Else
            
              sOb = ""
              Do While sOb = ""
                sOb = InputBox("Indique Observaciones:", "Reimpresión", sOb)
              Loop
              
              FG.Row = j: FG.Col = 2: FG.CellAlignment = flexAlignLeftCenter
              FG.TextMatrix(j, 0) = CStr(FG.Rows - 1)
           
              FG.Row = j: FG.Col = 1: FG.CellAlignment = flexAlignLeftCenter
              FG.TextMatrix(j, 1) = sCed 'Trim(eCed.Text)
            
              FG.Row = j: FG.Col = 2: FG.CellAlignment = flexAlignLeftCenter
              FG.TextMatrix(j, 2) = CStr(r.Fields("id").Value)
              
              FG.Row = j: FG.Col = 3: FG.CellAlignment = flexAlignLeftCenter
              FG.TextMatrix(j, 3) = sOb
                          
              If Check1.Value = vbChecked Then
                s = "update [" & st & "] set marca = 'I' where id = " & CStr(r.Fields("id").Value) & " "
                Modulo.ExecSQL s
              End If
              
              If s2 = "" Or s2 = "-" Then s2 = "0"
              
              s = "insert into Reimpresiones " & _
                  "(fecha,hora,cliente,subcliente,tabla,cedula,idpersona,observaciones) values " & _
                  "('" & _
                  Format(Date, "yyyymmdd") & "','" & _
                  Format(Time, "HH:mm") & "'," & CStr(Mid(s1, 1, 6) & "," & _
                  CStr(Mid(s2, 1, 6)) & ",'" & st & "','" & sCed & "'," & CStr(r.Fields("id").Value)) & ",'" & _
                  sOb & "')"
                  
              Modulo.ExecSQL s
                       
              FG.Rows = FG.Rows + 1
            End If
            
          End If
          r.Close
          Set r = Nothing
        End If
      End If
      
      'SeleccionarCampo eCed
      eCed.Text = ""
      eCed.SetFocus
      
      Color_Fila_Disponible
      
    Else
      MsgBox "No hay Fila Vacía donde copiar la cédula...", vbCritical, "Información"
    End If
  End If
  
End Sub

Private Sub FG_DblClick()
  Dim s As String
  If FG.Col = 2 Then 'está en posición de cédula
    s = FG.TextMatrix(FG.Row, FG.Col)
    s = InputBox("Indique Cedula:", "Editar", s)
    FG.TextMatrix(FG.Row, FG.Col) = s
    eCed.SetFocus
  End If
End Sub

Private Function Existe_Valor_FG(sValor As String, iColumna As Integer) As Boolean
  Dim i As Integer
  Dim e As Boolean
  i = 0
  e = False
  Do While i < FG.Rows And Not e
    If FG.TextMatrix(i, iColumna) = sValor Then
      e = True
    Else
      i = i + 1
    End If
  Loop
  Existe_Valor_FG = e
End Function

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

Private Sub Color_Fila_Disponible()
  Dim i As Integer, j As Integer
  
  For i = 1 To FG.Rows - 1
    For j = 0 To FG.Cols - 1
      FG.Row = i
      FG.Col = j
      FG.CellBackColor = vbWhite
    Next j
  Next i
  
  For i = 1 To FG.Rows - 1
    If Trim(FG.TextMatrix(i, 2)) = "" Then
      For j = 0 To FG.Cols - 1
        FG.Row = i
        FG.Col = j
        FG.CellBackColor = vbGreen
      Next j
      Exit Sub
    End If
  Next i
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
    s = Zeros(r.Fields("codigo").Value, 6) & " : " & Trim(r.Fields("nombre").Value)
    cCP.AddItem s
    r.MoveNext
  Loop
  
  r.Close
  Set r = Nothing
End Sub


Private Sub Cargar_SubClientes()
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim l As Integer
  Dim sCod As String
  
  cSC.Clear
  
  'Viendo = False
  
  sCod = "000000"
  If Trim(cCP.Text) <> "" Then sCod = Mid(cCP.Text, 1, 6)
      
  s = "SELECT * FROM SubClientes WHERE Cliente = " & sCod & " ORDER BY Id"
  
  r.Open s, DBConexionSQL, adOpenDynamic, adLockOptimistic
  l = 1
  Do While Not r.EOF
    s = Zeros(r.Fields("id").Value, 6) & " : " & Trim(r.Fields("nombre").Value)
    cSC.AddItem s
    r.MoveNext
  Loop
  r.Close
  
  If cSC.ListCount > 0 Then cSC.ListIndex = 0
  
End Sub


Private Sub Form_Load()
  Dim UI As String
  
  UI = Mid(App.Path, 1, 2)
    
  eCed.Text = ""
    
  FG.Clear
  FG.Rows = 2
  FG.FormatString = cFG
  
  'Dir1.Path = UI
  
  'File1.Refresh
  'CargarArchivosDeDisco
  Cargar_Clientes
  
  Check1.Value = vbChecked
  
End Sub
