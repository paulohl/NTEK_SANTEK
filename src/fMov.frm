VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Begin VB.Form fMov 
   Caption         =   "Movimiento Diario - Emisión de Carnets"
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Refrescar"
      Height          =   345
      Left            =   7740
      TabIndex        =   19
      Top             =   60
      Width           =   1455
   End
   Begin VB.CommandButton bLD 
      Caption         =   "Listar Datos"
      Height          =   345
      Left            =   9540
      TabIndex        =   16
      Top             =   60
      Width           =   1275
   End
   Begin VB.Frame Frame6 
      Caption         =   "Movimientos"
      Height          =   9830
      Left            =   0
      TabIndex        =   1
      Top             =   390
      Width           =   15195
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFC0FF&
         Height          =   675
         Left            =   4230
         ScaleHeight     =   615
         ScaleWidth      =   2505
         TabIndex        =   14
         Top             =   9090
         Width           =   2565
         Begin VB.CommandButton bTomarFoto 
            Caption         =   "Transferir Foto"
            Height          =   500
            Left            =   1380
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton bCard5 
            Caption         =   "CARD-5"
            Height          =   500
            Left            =   210
            Picture         =   "fMov.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   900
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFC0FF&
         Height          =   675
         Left            =   7170
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   12
         Top             =   9090
         Width           =   1395
         Begin VB.CommandButton bConsultarItem 
            Caption         =   "Ver Detalles"
            Height          =   500
            Left            =   150
            Picture         =   "fMov.frx":058A
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   1000
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFC0FF&
         Height          =   675
         Left            =   8850
         ScaleHeight     =   615
         ScaleWidth      =   4635
         TabIndex        =   10
         Top             =   9090
         Width           =   4695
         Begin VB.CommandButton cmdActualizarPersona 
            Caption         =   "Actualizar Persona"
            Height          =   500
            Left            =   1500
            Picture         =   "fMov.frx":0F8C
            TabIndex        =   18
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
         Begin VB.CommandButton bHistorico 
            Caption         =   "Histórico Movimientos"
            Height          =   500
            Left            =   2880
            Picture         =   "fMov.frx":198E
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   1700
         End
         Begin VB.CommandButton bBusqueda 
            Caption         =   "Búscar Persona"
            Height          =   500
            Left            =   120
            Picture         =   "fMov.frx":2390
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFC0FF&
         Height          =   675
         Left            =   1110
         ScaleHeight     =   615
         ScaleWidth      =   2685
         TabIndex        =   7
         Top             =   9090
         Width           =   2745
         Begin VB.CommandButton bNuevo 
            Caption         =   "Nuevo"
            Height          =   500
            Left            =   210
            Picture         =   "fMov.frx":2D92
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton bOtro 
            Caption         =   "Otro"
            Height          =   500
            Left            =   1410
            Picture         =   "fMov.frx":3794
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   900
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   -60
         Top             =   5400
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
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
         Caption         =   "Personas"
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
      Begin VB.CommandButton bCancelar 
         Caption         =   "Cancelar"
         Height          =   500
         Left            =   13710
         Picture         =   "fMov.frx":4196
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   9180
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CommandButton bGuardar 
         Caption         =   "Guardar"
         Enabled         =   0   'False
         Height          =   500
         Left            =   12570
         Picture         =   "fMov.frx":4720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   9180
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   900
      End
      Begin ubGridControl.ubGrid GRID 
         CausesValidation=   0   'False
         Height          =   8775
         Left            =   60
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   300
         Width           =   15045
         _ExtentX        =   26538
         _ExtentY        =   15478
         Rows            =   1
         Cols            =   5
         Redraw          =   -1  'True
         ShowGrid        =   -1  'True
         GridSolid       =   -1  'True
         GridLineColor   =   12632256
         BackColorAlt    =   16777152
         UseBackColorAlt =   -1  'True
         BackColorFixed  =   12632256
         RowHeader       =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontEdit {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorBkg    =   12632256
      End
   End
   Begin VB.Label lFecha 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   2
      Top             =   60
      Width           =   3855
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   450
      TabIndex        =   0
      Top             =   90
      Width           =   600
   End
End
Attribute VB_Name = "fMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFechaVencimientoFD As String ' para que traiga la fecha de vencimiento del formato de diseño


Dim ComboVenceListo As Boolean
Dim ComboCargoListo As Boolean
'Const FGC = "ID          | Nombre                                                    | RIF Nº             | Telefonos                                    "

Const FGC = "Código    | Nombre                                                          | RIF Nº                | Telefonos                                "

Dim RClientes As New ADODB.Recordset
Dim RSubClientes As New ADODB.Recordset

Dim Viendo As Boolean

Dim OPR As Integer  'operacion: 1 nuevo  2 editar
Public lTipoPago As Integer '0=Nopago; 1=pago; 0=Pagado por Adelantado;
Const OPR_Agregando = 1
Const OPR_Editando = 2

Public vRUTAINICIAL As String

Const COL_FECHAMOV = 2
Const COL_CLIENTE = 3
Const COL_SUBCLIENTE = 4
Const COL_CEDULA = 5

Const COL_ANCHO_MIN_CLIENTE = 150
Const COL_ANCHO_MAX_CLIENTE = 300

Const COL_ANCHO_MIN_SUB = 32
Const COL_ANCHO_MAX_SUB = 200

Const COL_TITULO_CEDULA = "CEDULA"
Const COL_TITULO_OBSERVACIONES = "OBSERVACIONES"

Const COL_ANCHO_1 = 80
Const COL_ANCHO_NOMBRE = 150
Const COL_ANCHO_OBSERVACIONES = 300





'Private Sub Formatear_Cedula(xTexto As TextBox)
'  Dim i As Integer
'  Dim d As Double
'  Dim s As String
'  If IsNumeric(xTexto.Text) Then
'    d = CDbl(xTexto.Text)
'    s = Format(d, "#,0")
'    xTexto.Text = s
'  End If
'End Sub

Private Sub Formatear_Str_Cedula(ByRef sCedula As String)
  Dim i As Integer
  Dim d As Double
  Dim s As String
  If IsNumeric(sCedula) Then
    d = CDbl(sCedula)
    s = Format(d, "#,0")
    sCedula = s
  End If
End Sub

Private Function LeerNuevoIDMov() As String
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim sID As String
  sID = ""
  s = "select top 1 id from diario order by id desc"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If Not r.EOF Then sID = CStr(r.Fields("id").value)
  r.Close
  Set r = Nothing
  LeerNuevoIDMov = sID
End Function

Private Function QuitarPuntos_CEDULA(sCed As String) As String
  Dim i As Integer, j As Integer, k As Integer
  Dim s As String
  s = ""
  For i = 1 To Len(sCed)
    If Mid(sCed, i, 1) <> "." And Mid(sCed, i, 1) <> "," Then
      s = s & Mid(sCed, i, 1)
    End If
  Next i
  QuitarPuntos_CEDULA = s
End Function


Private Function ColocarPuntos_CEDULA(sCed As String) As String
  Dim i As Integer, j As Integer, k As Integer
  Dim c As String
  Dim s As String
  Dim d As Double
  s = sCed
  If IsNumeric(s) Then
    d = CDbl(s)
    s = Format(d, "#,0")
  End If
  ColocarPuntos_CEDULA = s
End Function


'Almacena los items (movimientos) con ID en blanco (nuevos)
Private Sub GuardarMovimientos(Fila As Integer)
  Dim c As Integer
  Dim s As String, sSQL As String, sValor As String
  
  Dim sID As String         'ID
  Dim sFH As String         'FechaHora
  Dim sF As String          'Fecha
  Dim sH As String          'Hora
  Dim scte As String        'Cliente
  Dim sSCte As String       'Sub Cliente
  Dim sTabla As String      'Tabla
  Dim sCed As String        'Cedula
  Dim sObs As String        'Observaciones
  
  Dim dFecha As Date
  
  Dim rP As New ADODB.Recordset
  Dim bPersonaExiste As Boolean
  Dim i As Integer, iPC As Integer, iOBS As Integer
  Dim bActualizar As Boolean, sFoto As String
  Dim ssID As String, p As Integer
  
  Dim sDeudaT As String
  
  'MsgBox "Registro Guardado"
  'Exit Sub
  
  Load fMensaje
   
  'For Fila = 1 To GRID.Rows
  
  fMensaje.Label1.Caption = "Guardando Movimiento, Espere..."
  fMensaje.Show
  DoEvents
    
  sID = Trim(GRID.TextMatrix(Fila, 1))
    
  If sID = "" Then  'como no tiene ID en el GRID se asume q es Mov.Nuevo.
      
    sFH = GRID.TextMatrix(Fila, 2)
    sF = Mid(sFH, 1, 10)
    sH = Mid(sFH, 11)
    
    sF = Modulo.FechaInvertida(sF)
    
    scte = Mid(GRID.TextMatrix(Fila, 3), 1, 6)
    sSCte = GRID.TextMatrix(Fila, 4)
    sCed = Trim(GRID.TextMatrix(Fila, 5))
      
    If sSCte = "" Or sSCte = "-" Then sSCte = "0" Else sSCte = Mid(GRID.TextMatrix(Fila, 4), 1, 6)
      
    If scte = "" Then Exit Sub 'Falta Cliente
      
    If sCed = "" Then 'Falta Cedula
      MsgBox "Debe introducir la Cédula de la Persona, Revise...", vbCritical, "Información"
      Exit Sub
    End If
      
    sTabla = Modulo.La_Tabla_Actual_Personas(scte, sSCte)
    c = Modulo.La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES)
    If c > 0 Then
      sObs = GRID.TextMatrix(Fila, c)
    Else
      sObs = ""
      MsgBox "Debe Indicar la Observación del Movimiento para poder Guardarlo, Revise...", vbCritical, "Información"
      Exit Sub
    End If
      
    'dFecha = CDate(sFH)
    'sFH = Format(dFecha, "yyyymmdd HH:mm")
    
    'sF = Format(dFecha, "yyyy/mm/dd")
    'sH = Format(dFecha, "hh:mm ampm")
                        
    'sSQL = "insert into diario (fechahora,cliente,subcliente,tabla,cedula,observaciones) values ('" & _
           sFH & "','" & _
           scte & "','" & _
           sSCte & "','" & _
           sTabla & "','" & _
           sCed & "','" & _
           sObs & "')"
           
        
           
    sDeudaT = Format(Modulo.vMOVPAGO, "#0.00")  'CDbl(Modulo.vMOVPAGO)
    p = InStr(sDeudaT, ",")
    If p > 0 Then
      Mid(sDeudaT, p, 1) = "."
    End If
    Dim lReg As New ADODB.Recordset
    Dim lReg2 As New ADODB.Recordset
    Dim lCn As New ADODB.Connection
    Dim lSumaID As Integer
    lCn.ConnectionString = Modulo.DBConexionSQL
    lCn.Open
    lSumaID = 0
    Set lReg = lCn.Execute("Select max(id) as id from [" & sTabla & "]")
    Set lReg2 = lCn.Execute("Select [id] as id from [" & sTabla & "] where cedula='" & sCed & "'")
    If lReg2.EOF = True Then
      lSumaID = IIf(IsNull(lReg!ID), 0, lReg!ID) + 1
    Else
      lSumaID = lReg2!ID
    End If
    sSQL = "insert into diario (localizador,fecha,hora,cliente,subcliente,tabla,cedula,observaciones,pago,monto,idCarnet,TipoPago) values ('" & _
           Modulo.vLocalizador & "','" & _
           sF & "','" & _
           sH & "'," & _
           scte & "," & _
           sSCte & ",'" & _
           sTabla & "','" & _
           sCed & "','" & _
           sObs & "','" & IIf(Modulo.vMOVPAGO > 0#, "S", "N") & "'," & sDeudaT & "," & lSumaID & "," & lTipoPago & ")"
                                                                                      '(IIf(IsNull(lReg!ID), 0, lSumaID))                                                                                                        'antes  lReg!ID+1
             '(GRID.TextMatrix(GRID.Rows, 1) + 1) = para el id del carnet y poder ubicar el registro
             
    On Error Resume Next
    Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
    Modulo.DBComandoSQL.CommandText = sSQL
    Modulo.DBComandoSQL.Execute
      
    ''Clipboard.Clear
    ''Clipboard.SetText sSQL
    If Err.Number <> 0 Then
      MsgBox "Ha ocurrido un Error al intentar GUARDAR registro del Diario..." & vbCrLf & _
             Err.Description, vbCritical, "Información"
      Exit Sub
    Else
      
      s = LeerNuevoIDMov()
      ssID = s
      If s <> "" Then GRID.TextMatrix(Fila, 1) = s
      
      AgregarLogs "Agrega Movimiento al DIARIO [" & s & "]"
        
      '********************************************
      'Actualizar el record de Persona en Linea:
      '********************************************
      If sCed <> "" Then
        s = "select * from [" & sTabla & "] where cedula = '" & sCed & "'"
        rP.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
        bPersonaExiste = Not rP.EOF
          
        If bPersonaExiste Then
            
          '*************************************************************
          '* Los campos a ACTUALIZAR si la persona existe serian los que
          '* estan a la derecha de CEDULA hasta llegar a OBSERVACIONES.
          '*************************************************************
          iPC = Modulo.La_Columna_GRID(GRID, COL_TITULO_CEDULA)
          iOBS = Modulo.La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES)
            
          If iOBS > iPC Then
              
            sSQL = "update [" & sTabla & "] set "
              
            bActualizar = False
              
            For i = iPC + 1 To iOBS - 1
              s = GRID.TextMatrix(0, i)
              
              If Modulo.EXISTE_CAMPO(rP, s) Then 'Es columna de la Tabla
                           
                sValor = Trim(GRID.TextMatrix(Fila, i))
                If sValor <> "" Then
                  'Actualizar el campo:
                  sSQL = sSQL & s & " = '" & sValor & "',"
                  bActualizar = True
                End If
              End If
              
            Next i
              
            If Mid(sSQL, Len(sSQL), 1) = "," Then Mid(sSQL, Len(sSQL), 1) = " "
              
            sSQL = sSQL & " where cedula = '" & sCed & "'"
              
            If bActualizar Then
              On Error Resume Next
              Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
              Modulo.DBComandoSQL.CommandText = sSQL
              Modulo.DBComandoSQL.Execute
              If Err.Number <> 0 Then
                MsgBox "Ha ocurrido un Error al intentar ACTUALIZAR registro de Persona..." & vbCrLf & _
                Err.Description, vbCritical, "Información"
                Exit Sub
              End If
            End If
          End If
            
        Else
          
          '*************************************************************
          '* AGREGAR NUEVA PERSONA pq la CEDULA NO existe serian los que
          '* estan a la derecha de CEDULA hasta llegar a OBSERVACIONES.
          '*************************************************************
            
          Dim sValores As String, db_campo As String
                   
          sSQL = "insert into [" & sTabla & "] ("
          sValores = ""
              
          For i = 0 To rP.Fields.Count - 1
            
            db_campo = UCase(rP.Fields(i).Name)
            If db_campo <> "ID" Then
              sSQL = sSQL & db_campo & ","
                  
              If Modulo.Existe_Columna_GRID(GRID, db_campo) Then
                  
                  'Select Case fPersonasAct.Adodc1.Recordset.Fields(i).Type
                  '  Case adChar: s = "CHAR"
                  '  Case adInteger: s = "INTEGER"
                  '  Case adDBTimeStamp: s = "DATETIME"
                  '  Case adDouble: s = "FLOAT"
                  'End Select
                  
                If rP.Fields(i).Type = adChar Or _
                   rP.Fields(i).Type = adDBDate Then
                     
                  sValores = sValores & "'" & GRID.TextMatrix(Fila, Modulo.La_Columna_GRID(GRID, db_campo)) & "',"
                Else
                  If rP.Fields(i).Type = adInteger Or _
                     rP.Fields(i).Type = adDouble Then
                    sValores = sValores & " " & GRID.TextMatrix(Fila, Modulo.La_Columna_GRID(GRID, db_campo)) & " ,"
                  End If
                End If
              End If
            End If
          Next i
              
          '--Completar los Predeterminados:
            
          If Modulo.EXISTE_CAMPO(rP, "FOTO") Then
            s = sCed
            s = QuitarPuntos_CEDULA(sCed)
            sValores = sValores & "'" & s & "',"
          End If
            
          If Modulo.EXISTE_CAMPO(rP, "TIENE_FOTO") Then
            sValores = sValores & "'N',"
          End If
            
          If Modulo.EXISTE_CAMPO(rP, "MARCA") Then
            sValores = sValores & "' ',"
          End If
           
          'If Modulo.EXISTE_CAMPO(rP, "FECHA") Then
          '  sValores = sValores & "'" & Format(Date, "yyyymmdd") & " 00:00:00',"
          'End If
              
          If Modulo.EXISTE_CAMPO(rP, "FECHA") Then
            sValores = sValores & "NULL,"
          End If
              
          If Modulo.EXISTE_CAMPO(rP, "CONTADOR") Then
            sValores = sValores & "0,"
          End If
             
          If Modulo.EXISTE_CAMPO(rP, "CREACION") Then
            sValores = sValores & "'" & Format(Date, "yyyymmdd") & " 00:00:00',"
          End If
              
          '--Finalizar el comando SQL:
          If Mid(sSQL, Len(sSQL), 1) = "," Then Mid(sSQL, Len(sSQL), 1) = " "
          If Mid(sValores, Len(sValores), 1) = "," Then Mid(sValores, Len(sValores), 1) = " "
          sSQL = sSQL & ") VALUES (" & sValores & ")"
            
          On Error Resume Next
          Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
          Modulo.DBComandoSQL.CommandText = sSQL
          ''Clipboard.Clear
          ''Clipboard.SetText sSQL
          Modulo.DBComandoSQL.Execute
          If Err.Number <> 0 Then
            MsgBox "Ha ocurrido un Error al intentar AGREGAR registro de Persona..." & vbCrLf & _
                   Err.Description, vbCritical, "Información"
            Exit Sub
          End If
        End If
      End If
      
      '**************************************************************
      ' Asociar los Detalles del Movimiento previamente almacenados:
      '**************************************************************
      'ssID = LeerNuevoIDMov()
      's = "update DiarioDetalle set idDiario = " & CStr(ssID) & " where " & _
          "idDiario = -1 and " & _
          "estacion = '" & Modulo.ESTACION & "'"
          
      'Modulo.ExecSQL s
      
      'Realizar chequeo de las Fotos en su carpeta para mantener
      'actualizado el campo Tiene_Foto:
      Auditar_Fotos Modulo.La_Tabla_Actual_Personas(scte, sSCte)
           
      
      
    End If
  End If
  
  
  MostrarObservaciones
     
  Unload fMensaje
   
End Sub

Private Sub bBusqueda_Click2()
  Dim f As Integer, c As Integer, i As Integer
  Dim e As Boolean
  
  Modulo.vTemporal1 = ""
  Load fBuscarPersona
  
  With fBuscarPersona
    
    .Combo1.Clear
    .Combo1.AddItem "CEDULA"
    .Combo1.AddItem "NOMBRE"
    .Combo1.ListIndex = 0
    
    
    .FG.Cols = 6
    .FG.TextMatrix(0, 0) = "ID"
    .FG.TextMatrix(0, 1) = "CLIENTE"
    .FG.TextMatrix(0, 2) = "SUBCLIENTE"
    .FG.TextMatrix(0, 3) = "CEDULA"
    .FG.TextMatrix(0, 4) = "NOMBRE"
    .FG.TextMatrix(0, 5) = "CARGO"
    
    .FG.ColWidth(0) = 500
    .FG.ColWidth(1) = 4000
    .FG.ColWidth(2) = 2800
    .FG.ColWidth(3) = 1000
    .FG.ColWidth(4) = 3000
    .FG.ColWidth(5) = 1500
 
           
  End With
   
  'fBuscarSimple.BuscarSimple fClientes.FG
  Modulo.vTemporal1 = ""
  'fBuscarSimple.Option2.Value = True 'Buscar por nombre por defecto
  Modulo.fModalResult = Modulo.fModalResultCANCEL
  
  fBuscarPersona.Show vbModal
  
'  If Modulo.vTemporal1 <> "" Then
'    If Len(Modulo.vTemporal1) < 6 Then Modulo.vTemporal1 = Zeros(CLng(Modulo.vTemporal1), 6)
'    cCP.ListIndex = Modulo.Buscar_ComboLen(cCP, Modulo.vTemporal1, 6)
'  End If
  
  If Modulo.fModalResult = Modulo.fModalResultOK Then
  
    'GRID_AfterEdit GRID.Row, GRID.Col, ""
    
  End If
    
  
  
  
  
  
  
  
  


End Sub

Private Sub bBusqueda_Click()
  frmBuscar.lTabla = "Diario"
  frmBuscar.Show vbModal
End Sub

Private Sub bCancelar_Click()
  Unload Me
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
  
  For i = 1 To GRID.Rows
    SC = Trim(GRID.TextMatrix(i, 3))
    sSC = Trim(GRID.TextMatrix(i, 4))
    
    If SC <> "" Then
      SC = Mid(SC, 1, 6)
      If sSC = "" Or sSC = "-" Then sSC = "0" Else sSC = Mid(sSC, 1, 6)
      
      Auditar_Fotos Modulo.La_Tabla_Actual_Personas(SC, sSC)
      
    End If
  Next i
      
  Dim sOri As String
  Dim sDes As String
  Dim lResp As String
  Dim lR2 As String
  sOri = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
  sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
  
  If GRID.Row >= 1 And GRID.Row <= GRID.Rows Then i = GRID.Row Else Exit Sub
      
  SC = Trim(GRID.TextMatrix(i, 3))
  sSC = Trim(GRID.TextMatrix(i, 4))
      
  SC = Trim(Mid(SC, 7))
       
  If Trim(sSC) = "" Or Trim(sSC) = "-" Then
    sRuta = sDes & "\" & SC & "\CARNET" & "\BASE CARNET " & SC & ".car"
    lR2 = sDes & "\" & SC & "\CARNET"
  Else
    sSC = Trim(Mid(sSC, 7))
    sRuta = sDes & "\" & SC & "\" & sSC & "\CARNET" & "\BASE CARNET " & sSC & ".car"
    lR2 = sDes & "\" & SC & "\" & sSC & "\CARNET"
  End If

  s = GetSetting(APPNAME, "Opciones", "RutaCard5", "")

  sCard5 = s & " " & Chr(34) & sRuta & Chr(34)
  ''verificar si existe mas de una archiv .Car, de sera asi permitirle al usuario seleccionar el .car
  Load frmSeleccionarCard5
  lResp = frmSeleccionarCard5.sCargarArchivosCard5(lR2)
  If lResp <> lR2 Then
     frmSeleccionarCard5.Show vbModal
  Else
     If Shell(sCard5, vbMaximizedFocus) = 0# Then
       MsgBox "Error: No se pudo Iniciar Card-5" & vbCrLf & CStr(Err.Number) & ":" & Err.Description, "Información"
     End If
  End If
  Unload fMensaje
    
End Sub


Private Sub Ver_Indicaciones_Cliente(bMostrarMensaje As Boolean)
'  Dim s As String, sind As String
'  Dim r As New ADODB.Recordset
'
'  If Trim(cCP.Text) <> "" Then
'    sind = ""
'    s = "select indicaciones from clientes where codigo = " & Mid(cCP.Text, 1, 6)
'    r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
'    If Not r.EOF Then
'      If Not IsNull(r.Fields("indicaciones").Value) Then
'        sind = Trim(r.Fields("indicaciones").Value)
'      End If
'    End If
'    r.Close
'    Set r = Nothing
'    If sind <> "" Then
'      Load fIndicaciones
'      fIndicaciones.Text1.Text = sind
'      fIndicaciones.Show vbModal
'    Else
'      If bMostrarMensaje Then
'        MsgBox "No hay Indicaciones Especiales de este Cliente...", vbCritical, "Información"
'      End If
'    End If
'  End If
End Sub

Private Sub cCP_LostFocus()
  Ver_Indicaciones_Cliente False
End Sub



Function Numero_Siguiente_Tabla_Personas() As Integer
'  Dim r As New ADODB.Recordset
'  Dim s As String
'  Dim sC As String, sSC As String
'  Dim k As Integer, n As Integer
'
'  sC = Trim(Mid(cCP.Text, 1, 6))
'  sSC = Trim(Mid(cSC.Text, 1, 6))
'  If sSC = "" Then sSC = "0"
'
'  s = "select * from Personas where " & _
'      "cliente    = " & sC & " and " & _
'      "subcliente = " & sSC & " order by id "
'
'  k = 0
'
'  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
'  Do While Not r.EOF
'    s = Trim(r.Fields("tabla").Value)
'    s = Trim(Mid(s, InStrRev(s, "-") + 1))
'    k = CInt(s)
'    r.MoveNext
'  Loop
'  r.Close
'  Set r = Nothing
'  Numero_Siguiente_Tabla_Personas = k + 1
End Function


Private Sub bConsultarItem_Click()
  Dim i As Integer
  Dim r As New ADODB.Recordset
  Dim s As String, sID As String, s2 As String
  Dim dT As Double, xLoc As String
  
  i = GRID.Row
  If i >= 1 Then
  
    If Trim(GRID.TextMatrix(i, 1)) = "" Then Exit Sub
  
    Load fMov2
    fMov2.List1.Clear
  
    dT = 0#
    sID = GRID.TextMatrix(i, 1) 'ID
    
    xLoc = Modulo.Localizador_Por_ID("Diario", sID)
    If xLoc = "" Then
      MsgBox "No Tiene Detalles este Movimiento...", vbCritical, "Información"
      Exit Sub
    End If
    
    fMov2.List2.Clear
    
    s = "select * from DiarioDetalle where Localizador = '" & xLoc & "' order by id"
    r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    
    Do While Not r.EOF
      
      s2 = Modulo.Producto_DESCRIPCION(Trim(r.Fields("codigoproducto").value))
      If Len(s2) > 30 Then s2 = Mid(s2, 1, 30) Else s2 = Modulo.BlancosDER(s2, 30)
      
      's = r.Fields("codigoproducto").Value & " " & _
          s2 & " " & _
          Modulo.BlancosIZQ(CStr(r.Fields("cantidad").Value), 12) & " " & _
          Modulo.BlancosIZQ(Format(r.Fields("precio").Value, "#,0.00"), 11) & " " & _
          Modulo.BlancosIZQ(Format(r.Fields("subtotal").Value, "#,0.00"), 12)
          
          
      s = r.Fields("codigoproducto").value & " " & _
          s2 & " " & _
          Modulo.BlancosIZQ(CStr(r.Fields("cantidad").value), 12) & " " & _
          Modulo.BlancosIZQ(Format(r.Fields("precio").value, "#,0.00"), 11) & " " & _
          Modulo.BlancosIZQ(Format(r.Fields("subtotal").value, "#,0.00"), 12) & "       " & _
          IIf(r.Fields("Entregado").value = "S", "SI", "NO")
          
      dT = dT + r.Fields("subtotal").value
      
      fMov2.List2.AddItem r.Fields("id").value
      
      fMov2.List1.AddItem s
      r.MoveNext
    Loop
    
    r.Close
    Set r = Nothing
    
    fMov2.lTotal.Caption = Format(dT, "#,0.00")
    
    fMov2.Caption = "Detalle del Movimiento " & GRID.TextMatrix(GRID.Row, 1)
    
    fMov2.Show vbModal
    
    'fMov2.Localizador = xLoc
    
   ' If FG.TextMatrix(i, 8) = "NO" Then fMov2.cmdPagar.Visible = True
   ' fMov2.txtObservaciones.Text = FG.TextMatrix(i, 6)
   ' fMov2.Show vbModal

    
  End If

End Sub

Private Sub bGuardar_Click()
  'GuardarMovimientos
  'bGuardar.Enabled = False
End Sub

Function Hay_Mov_Pendiente() As Boolean
  Dim i As Integer
  Dim pendiente As Boolean
  Dim lCn As New ADODB.Connection
  Dim lReg As New ADODB.Recordset
  lCn.ConnectionString = Modulo.DBConexionSQL
  lCn.Open
  Set lReg = lCn.Execute("Select ID from diario where fecha='" & Format(Now, "yyyyMMdd") & "'")
  If lReg.EOF = True Then
     Hay_Mov_Pendiente = False
     Exit Function
  End If
  pendiente = False
  i = 1
  
  Do While i <= GRID.Rows And Not pendiente
    If Trim(GRID.TextMatrix(i, 1)) = "" Then ' ID
      pendiente = True
    End If
    i = i + 1
  Loop
  Hay_Mov_Pendiente = pendiente
End Function

Private Sub bHistorico_Click()
  fConsultaMovs.Show vbModal
End Sub

Private Sub bLD_Click()
  VentanaMostrarCampos
End Sub

Private Sub bNuevo_Click()
  Dim co As Integer, i As Integer
  co = La_Columna_GRID(GRID, "OBSERVACIONES")
  If Hay_Mov_Pendiente() Then
    MsgBox "Existe un Movimiento que aun no se ha Guardado," & vbCrLf & _
           "Diríjase a Observaciones, introduzaca un valor y pulse <ENTER>", vbCritical, "Información"
    If co > 0 Then
      GRID.Col = co
      SendKeys "F2"
      Exit Sub
    End If
  Else
    'If GRID.Rows >= 1 And (GRID.TextMatrix(GRID.Rows, 1)) = "" Then
      GRID.Rows = GRID.Rows + 1
    'End If
      GRID.TextMatrix(GRID.Rows, COL_FECHAMOV) = Format(Now, "dd/mm/yyyy hh:mm ampm")
      'GRID.Row = GRID.Rows
      GRID.Col = COL_CLIENTE
      GRID.SetFocus
      SendKeys "{HOME}"
      For i = 1 To GRID.Rows
        SendKeys "{DOWN}"
      Next i
      'GRID.Row = GRID.Rows
    
  End If
  
  'bGuardar.Enabled = False
  
End Sub

Private Sub bOtro_Click()
  Dim co As Integer, i As Integer, CE As Integer
  Dim Fila As Integer
    
  co = La_Columna_GRID(GRID, "OBSERVACIONES")
  If Hay_Mov_Pendiente() Then
    MsgBox "Existe un Movimiento que aun no se ha Guardado," & vbCrLf & _
           "Complete los datos y pulse la tecla <ENTER> en la " & vbCrLf & _
           "Columna <Observaciones> con sus respectivos datos...", vbCritical, "Información"
    Exit Sub
  Else
    GRID.Rows = GRID.Rows + 1
    GRID.TextMatrix(GRID.Rows, COL_FECHAMOV) = Format(Now, "dd/mm/yyyy hh:mm ampm")
    GRID.TextMatrix(GRID.Rows, COL_CLIENTE) = GRID.TextMatrix(GRID.Rows - 1, COL_CLIENTE)
    GRID.TextMatrix(GRID.Rows, COL_SUBCLIENTE) = GRID.TextMatrix(GRID.Rows - 1, COL_SUBCLIENTE)
    'sFechaVencimientoFD Mid(GRID.TextMatrix(GRID.Rows - 1, COL_CLIENTE), 1, 6)
    CE = La_Columna_GRID(GRID, "CEDULA")
    If CE > 0 Then GRID.Col = CE Else GRID.Col = COL_CLIENTE
    GRID.SetFocus
    SendKeys "{HOME}"
    For i = 1 To GRID.Rows
      SendKeys "{DOWN}"
    Next i
    'GRID.Row = GRID.Rows
  End If
  
  'bGuardar.Enabled = False

End Sub

Private Sub bTomarFoto_Click()
  Dim SC As String, sSC As String
  Dim sRuta As String, sF As String
  Dim s As String
  Dim i As Integer
  If GRID.Rows <= 1 And GRID.TextMatrix(1, 1) = "" Then Exit Sub
  Load fMensaje
  fMensaje.Caption = "Preparando para Ejecutar «TOMAR FOTO», Espere..."
  fMensaje.Show
  DoEvents
  
  Dim sOri As String
  Dim sDes As String

  sOri = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
  sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")

  s = GetSetting(APPNAME, "Opciones", "RutaFotos", "")

  'sF = s & " " & Chr(34) & sRuta & Chr(34)
  
  sF = s
  
  Unload fMensaje
    
  
  FrmSelCarnets.Show vbModal
  'frmTransferirFotos.Show vbModal
  'If Shell(sF, vbMaximizedFocus) = 0# Then
  '  MsgBox "Error: No se pudo Iniciar TOMAR-FOTOS" & vbCrLf & CStr(Err.Number) & ":" & Err.Description, "Información"
  'End If
  
 'GRID.Cols = GRID.Cols + 1
 'GRID.ColMask(GRID.Cols) = checkmark

End Sub

Private Sub Command1_Click()
 Form_Load
End Sub

Private Sub cmdActualizarPersona_Click()
   'Load fPersonasAct
   fPersonasAct.Show
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
                 'si presiona F5 refrescar
  If KeyCode = 116 Then
     Form_Load
  End If
End Sub

Private Sub Form_Load()
  NCol = 0
  OPR = 0
  'Cargar_Tablas_Personas
  'Cargar_Clientes
  'Cargar_SubClientes
  
  'Cargar_Ordenar_Por
  'bRefrescar_Click
  lFecha.Caption = Format(Now, "Long Date")
  'eValor.Text = ""
  'eObs.Text = ""
  
  'CargarFGDiario
  
  'lInicializadas.Clear
  
  CargarGridEditorDiario
  
  If Trim(GRID.TextMatrix(GRID.Rows, 1)) = "" Then
    GRID.TextMatrix(GRID.Rows, 2) = Format(Now, "dd/mm/yyyy hh:mm ampm")
  End If
  
  GRID.ColAllowEdit(1) = False  'No Editar ID
  'If UCase(GRID.TextMatrix(0, 8)) = "VENCE" Then
  ''GRID.AddButton 8 ' boton para fecha de vencimiento
  '   GRID.AddLookup 8, "2 AÑOS" ''UCase(MonthName(Month(Date))) & " " & Year(Date) + 1
  '   GRID.AddLookup 8, "1 AÑO"  ''UCase(MonthName(Month(Date) + 5)) & " " & Year(Date)
  '   GRID.AddLookup 8, "6 MESES"  ''UCase(MonthName(Month(Date) + 5)) & " " & Year(Date)
  '   GRID.AddLookup 8, "3 MESES"  ''UCase(MonthName(Month(Date) + 5)) & " " & Year(Date)
  ''GRID.ColMask(1) = checkmark
  'End If
  'Picture5.Visible = False
  '2 AÑOS / UN AÑO / 6 MESES / 3 MESES
  ComboVenceListo = False
  ComboCargoListo = False
End Sub

Private Sub CargarComboSubCliente(sCliente As String)
  Dim r As New ADODB.Recordset
  Dim s As String, k As Integer
  sCliente = Mid(sCliente, 1, 6)
  s = "select * from subclientes where cliente = " & sCliente & " order by id"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  GRID.RemoveLookup COL_SUBCLIENTE
  k = 0
  Do While Not r.EOF
    GRID.AddLookup COL_SUBCLIENTE, Zeros(r.Fields("id").value, 6) & " " & r.Fields("Nombre").value
    If k = 0 Then GRID.TextMatrix(GRID.Row, COL_SUBCLIENTE) = Zeros(r.Fields("id").value, 6) & " " & r.Fields("Nombre").value
    'Llenar la celda siempre que quede la ultima tabla => La Actual:
    'GRID.TextMatrix(GRID.Row, COLSUBCLIENTE) = Zeros(r.Fields("id").Value, 6) & " " & r.Fields("Nombre").Value
    
    r.MoveNext
    k = k + 1
  Loop
  r.Close
  Set r = Nothing
  GRID.TextMatrix(GRID.Row, COL_SUBCLIENTE) = ""
  If k = 0 Then GRID.TextMatrix(GRID.Row, COL_SUBCLIENTE) = "-"
End Sub

Private Sub CargarComboTablasCliente(iCliente As Integer, iSubCliente As Integer)
'  Dim r As New ADODB.Recordset
'  Dim s As String, k As Integer
'  If iCliente > 0 And iSubCliente > 0 Then
'    s = "select * from personas where cliente = " & CStr(iCliente) & " and subcliente = " & CStr(iSubCliente)
'  Else
'    s = "select * from personas where cliente = " & CStr(iCliente) & " and subcliente = 0"
'  End If
'
'  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
'  GRID.RemoveLookup COLTABLA
'  k = 0
'  Do While Not r.EOF
'    GRID.AddLookup COLTABLA, Trim(r.Fields("tabla").Value)
'    'If k = 0 Then GRID.TextMatrix(GRID.Row, COLTABLA) = Trim(r.Fields("tabla").Value)
'    GRID.TextMatrix(GRID.Row, COLTABLA) = Trim(r.Fields("tabla").Value)
'    r.MoveNext
'    k = k + 1
'  Loop
'  r.Close
'  Set r = Nothing
'  If k = 0 Then GRID.TextMatrix(GRID.Row, COLTABLA) = ""
  
End Sub

Private Sub CargarComboCedulasCliente(sTabla As String)
  Dim r As New ADODB.Recordset
  Dim s As String, k As Integer
  Dim cc As Integer
  s = "select * from [" & sTabla & "] order by cedula"
  cc = Modulo.La_Columna_GRID(GRID, "CEDULA")
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  GRID.RemoveLookup cc
  k = 0
  Do While Not r.EOF
    GRID.AddLookup cc, Trim(r.Fields("cedula").value)
    If k = 0 Then GRID.TextMatrix(GRID.Row, cc) = Trim(r.Fields("cedula").value)
    r.MoveNext
    k = k + 1
  Loop
  r.Close
  Set r = Nothing
  If k = 0 Then GRID.TextMatrix(GRID.Row, cc) = ""
End Sub

Private Sub CargarComboObservaciones()
  Dim r As New ADODB.Recordset
  Dim s As String, k As Integer
  k = Modulo.La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES)
  If k >= 0 Then
    s = "select * from observaciones order by observacion"
    r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    GRID.ColAllowEdit(k) = True
    GRID.AddButton k
    'GRID.RemoveLookup k
    Do While Not r.EOF
      GRID.AddLookup k, Trim(r.Fields("observacion").value)
      'If k = 0 Then GRID.TextMatrix(GRID.Row, COLTABLA) = Trim(r.Fields("tabla").Value)
      'GRID.TextMatrix(GRID.Row, k) = Trim(r.Fields("observacion").Value)
      r.MoveNext
    Loop
    r.Close
  End If
  Set r = Nothing
End Sub



Function ExisteCedulaEnPersonas(sTabla As String, sCedula As String) As Long
  Dim r As New ADODB.Recordset
  Dim s As String
  
  ExisteCedulaEnPersonas = -1
  
  Formatear_Str_Cedula sCedula
  
  s = "select ID from [" & sTabla & "] where cedula = '" & sCedula & "'"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If Not r.EOF Then
    If Not IsNull(r.Fields("ID").value) Then
      ExisteCedulaEnPersonas = r.Fields("ID").value
    End If
  End If
  
  r.Close
  Set r = Nothing
End Function

Private Sub AnexarColumnasDePersonas(sTabla As String, lID As Long, GRID As ubGrid, iFila As Long, bValorEnBlanco As Boolean)
  Dim r As New ADODB.Recordset
  Dim s As String, ss As String
  Dim i As Integer, p As Integer, po As Integer, pc As Integer
  Dim bExisteReg As Boolean
  
  'If lID <= 0 Then Exit Sub
  
    
  s = "select * from [" & sTabla & "] where id = " & CStr(lID)
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  
  bExisteReg = Not r.EOF
    
  'If Not r.EOF Then
    For i = 0 To r.Fields.Count - 1
      s = UCase(r.Fields(i).Name)
      ss = ""
      'If Not IsNull(r.Fields(s).Value) Then ss = Trim(r.Fields(s).Value)
      
      'If bExisteReg Then ss = Trim(r.Fields(s).Value)
      'If bValorEnBlanco Then ss = ""
      'p = Modulo.La_Columna_GRID(GRID, s)
      If EsColumnaMostrable(s) And (s <> COL_TITULO_CEDULA) Then
      
        If bExisteReg Then ss = IIf(IsNull(Trim(r.Fields(s).value)), "", r.Fields(s).value)
        If bValorEnBlanco Then ss = ""
        p = Modulo.La_Columna_GRID(GRID, s)
        'If s = "VENCE" Then
        '   GRID.TextMatrix(GRID.Row, p) = lFechaVencimientoFD
        '   ss = lFechaVencimientoFD
        'End If
   
        If p < 1 Then 'la columna no existe en el grid :: agregarla
          po = Modulo.La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES)
          
          If po >= 0 Then 'Existe la columna de Observaciones
          
            'If GRID.ColHasButton(po) Then GRID.RemoveButton po
          
            'Hay que insertar la nueva antes de observaciones
            GRID.Cols = GRID.Cols + 1
            Intercambiar_Columnas_Grid GRID, po, GRID.Cols
            
            GRID.TextMatrix(0, po) = s
            GRID.TextMatrix(iFila, po) = ss
            GRID.ColWidth(po) = COL_ANCHO_1
            If s = "NOMBRE" Then GRID.ColWidth(GRID.Cols) = COL_ANCHO_NOMBRE
            If s = COL_TITULO_OBSERVACIONES Then GRID.ColWidth(GRID.Cols) = COL_ANCHO_OBSERVACIONES
            
            GRID.ColEditWidth(GRID.Cols) = r.Fields(i).DefinedSize
                        
          Else
            GRID.Cols = GRID.Cols + 1
            GRID.TextMatrix(0, GRID.Cols) = s
            GRID.TextMatrix(iFila, GRID.Cols) = ss
                                   
            GRID.ColWidth(GRID.Cols) = COL_ANCHO_1
            If s = "NOMBRE" Then GRID.ColWidth(GRID.Cols) = COL_ANCHO_NOMBRE
            If s = COL_TITULO_OBSERVACIONES Then GRID.ColWidth(GRID.Cols) = COL_ANCHO_OBSERVACIONES
            
            GRID.ColEditWidth(GRID.Cols) = r.Fields(i).DefinedSize
                       
          End If
          
        Else
          GRID.TextMatrix(iFila, p) = ss
        End If
      End If
    Next i
        
    po = Modulo.La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES)
    If po < 0 Then 'No Existe la columna de Observaciones => Crearla
      GRID.Cols = GRID.Cols + 1
      GRID.TextMatrix(0, GRID.Cols) = COL_TITULO_OBSERVACIONES
      GRID.ColWidth(GRID.Cols) = COL_ANCHO_OBSERVACIONES
      'CargarComboObservaciones
      
    End If
    
    pc = Modulo.La_Columna_GRID(GRID, COL_TITULO_CEDULA)
    po = Modulo.La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES)
    
    
    GRID.AddButton po
    'Ajustar_Columna_GRID GRID, pc + 1, po
    
    
  'End If
  r.Close
  Set r = Nothing
End Sub

Private Sub Intercambiar_Columnas_Grid(xGrid As ubGrid, iCOL1 As Integer, iCOL2 As Integer)
  Dim aux As String
  Dim Filas As Long, i As Long
  Filas = xGrid.Rows
  For i = 0 To Filas
    aux = xGrid.TextMatrix(i, iCOL1)
    xGrid.TextMatrix(i, iCOL1) = xGrid.TextMatrix(i, iCOL2)
    xGrid.TextMatrix(i, iCOL2) = aux
  Next i
End Sub


Private Sub LimpiarCeldasPersona(iFila As Long)
  Dim c As Integer, o As Integer
  Dim i As Integer
  c = Modulo.La_Columna_GRID(GRID, COL_TITULO_CEDULA)
  o = Modulo.La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES)
  If o > c Then
    'Desde la cedula hasta Observaciones::Limpiar
    For i = c + 1 To o
      GRID.TextMatrix(iFila, i) = ""
    Next i
  End If
End Sub


Private Sub Form_Resize()
  Frame6.width = Me.width - 130
  GRID.width = Me.width - 280
  
End Sub

Private Sub GRID_AfterEdit(ByVal Row As Long, ByVal Col As Long, ByVal NewValue As String)
  Dim s As String, C1 As String, C2 As String, C3 As String, sIndicaciones As String
  Dim cced As Integer, idced As Long, co As Integer, sOb As String
  Dim tc As String, SC As String, sSC As String, st As String
  Dim lMes As Integer
  Dim lAño As Integer

  'If fMov.ActiveControl Is bBusqueda Then Exit Sub
        
  NewValue = UCase(NewValue)
  GRID.TextMatrix(GRID.Row, GRID.Col) = NewValue
    
  s = GRID.TextMatrix(Row, Col)
  s = NewValue
  C1 = ""
  C2 = ""
  
  cced = Modulo.La_Columna_GRID(GRID, COL_TITULO_CEDULA)
  co = Modulo.La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES)
  
  SC = Trim(GRID.TextMatrix(GRID.Row, 3))
  sSC = Trim(GRID.TextMatrix(GRID.Row, 4))
  
   
  If UCase(GRID.TextMatrix(0, Col)) = "VENCE" Then
  'GRID.AddButton 8 ' boton para fecha de vencimiento
     If ComboVenceListo = False Then
        ComboVenceListo = True
        GRID.AddLookup Col, "2 AÑOS" ''UCase(MonthName(Month(Date))) & " " & Year(Date) + 1
        GRID.AddLookup Col, "1 AÑO"  ''UCase(MonthName(Month(Date) + 5)) & " " & Year(Date)
        GRID.AddLookup Col, "6 MESES"  ''UCase(MonthName(Month(Date) + 5)) & " " & Year(Date)
        GRID.AddLookup Col, "3 MESES"  ''UCase(MonthName(Month(Date) + 5)) & " " & Year(Date)
        GRID.AddLookup Col, "PERSONALIZAR"
        sFechaVencimientoFD Mid(GRID.TextMatrix(GRID.Rows - 1, COL_CLIENTE), 1, 6)
        GRID.AddLookup Col, MonthName(Month(lFechaVencimientoFD)) & " " & Year(lFechaVencimientoFD)
        GRID.TextMatrix(GRID.Row, Col) = MonthName(Month(lFechaVencimientoFD)) & " " & Year(lFechaVencimientoFD)
        GRID.Col = GRID.Col + 1
     End If
  'GRID.ColMask(1) = checkmark
     Select Case GRID.TextMatrix(GRID.Row, Col)
        Case "2 AÑOS"
           GRID.TextMatrix(GRID.Row, Col) = UCase(MonthName(Month(Date))) & " " & Year(Date) + 2
        Case "1 AÑO"
           GRID.TextMatrix(GRID.Row, Col) = UCase(MonthName(Month(Date))) & " " & Year(Date) + 1
        Case "6 MESES"
           If Month(Date) <= 6 Then
              GRID.TextMatrix(GRID.Row, Col) = UCase(MonthName(Month(Date) + 5)) & " " & Year(Date)
           Else
              lMes = Month(Date)
              lMes = lMes + 6
              lAño = 0
              If lMes > 12 Then
                 lMes = lMes - 12
                 lAño = 1
              End If
              GRID.TextMatrix(GRID.Row, Col) = UCase(MonthName(lMes)) & " " & Year(Date) + lAño
           End If
        Case "3 MESES"
           If Month(Date) <= 9 Then
              GRID.TextMatrix(GRID.Row, Col) = UCase(MonthName(Month(Date) + 2)) & " " & Year(Date)
           Else
              lMes = Month(Date)
              lMes = lMes + 3
              If lMes > 12 Then
                 lMes = lMes - 12
                 lAño = 1
              End If
              GRID.TextMatrix(GRID.Row, Col) = UCase(MonthName(lMes)) & " " & Year(Date) + lAño
           End If
        Case "PERSONALIZAR"
           frmVencimiento.lColumna = Col
           frmVencimiento.Show vbModal
           
      End Select
      
      ComboVenceListo = True
   End If
   
    'If UCase(GRID.TextMatrix(0, Col)) = "CARGO" And ComboCargoListo = False Then
    ''GRID.AddButton 8 ' boton para fecha de vencimiento
    '    Dim lRegCargo As New ADODB.Recordset
    '    Dim lSC As String
    '    Dim lsSC As String
    '    Dim lst As String
    '    Dim ls As String
    '    ComboCargoListo = True
    '    lSC = Mid(Trim(GRID.TextMatrix(GRID.Row, 3)), 1, 6)
    '    lsSC = Mid(Trim(GRID.TextMatrix(GRID.Row, 4)), 1, 6)
   '
   '     lst = Modulo.La_Tabla_Actual_Personas(lSC, lsSC)
   '     ls = GRID.TextMatrix(0, GRID.Col)
   '     lRegCargo.Open "select DISTINCT(" & ls & ") from [" & lst & "] order by " & ls, Modulo.DBConexionSQL, adOpenStatic
   '     Do While lRegCargo.EOF = False
   '        If lRegCargo.Fields(ls).Value <> "" Then GRID.AddLookup Col, UCase(lRegCargo.Fields(ls).Value)
   '        lRegCargo.MoveNext
   '     Loop
   ' End If
 
   
   
  Select Case Col
    Case COL_SUBCLIENTE:
      '-Antes de introducir la cedula, buscar las indicaciones especiales
      '-del cliente/subcliente y mostrarlo al operador:
      s = Trim(LAS_INDICACIONES_CLIENTE(GRID.TextMatrix(Row, 3), GRID.TextMatrix(Row, 4)))
      If s <> "" Then
        Load fIndicaciones
        fIndicaciones.Text1.Text = s
        fIndicaciones.Show vbModal
      End If
      
      cced = Modulo.La_Columna_GRID(GRID, COL_TITULO_CEDULA)
      
      If cced > 1 Then
        GRID.Col = cced - 1
        GRID.ColAllowEdit(cced - 1) = True
        GRID.SetFocus
        SendKeys "{ENTER}"
      End If


    Case COL_CLIENTE:
    
      'Busqueda CLIENTE
      Modulo.vTemporal1 = ""
      Load fBuscarSimple
  
      With fBuscarSimple
        .Combo1.Clear
        .Combo1.AddItem "CODIGO"
        .Combo1.AddItem "NOMBRE"
        .Combo1.ListIndex = 1
    
        .Adodc1.ConnectionString = Modulo.DBConexionSQL.ConnectionString
        .Adodc1.RecordSource = "select codigo, nombre, direccion, telefonos from clientes order by codigo"
        .Adodc1.Caption = "CLIENTES"
        .Adodc1.Refresh
        .DataGrid1.Refresh
        AjustaColumnaDataGrid .DataGrid1, "NOMBRE", 4000
      End With
      DoEvents
  
      Modulo.vTemporal1 = ""
      fBuscarSimple.Show vbModal
   
      If Modulo.vTemporal1 <> "" Then
        'sFechaVencimientoFD Modulo.vTemporal1
      If Len(Modulo.vTemporal1) < 6 Then Modulo.vTemporal1 = Zeros(CLng(Modulo.vTemporal1), 6)
        GRID.TextMatrix(Row, Col) = Modulo.vTemporal1 & " " & Trim(Modulo.DBValorStr("clientes", "codigo", Modulo.vTemporal1, "nombre"))
        
        'sIndicaciones = DBValorVariant("clientes", "codigo", CLng(Modulo.vTemporal1), "indicaciones")
        
        'sIndicaciones = LAS_INDICACIONES_CLIENTE(sC, sSC)
        
        'If sIndicaciones <> "" Then
        '  Load fIndicaciones
        '  fIndicaciones.Text1.Text = sIndicaciones
        '  fIndicaciones.Show vbModal
        'End If
                
        If DBValorLng("subclientes", "cliente", CLng(Modulo.vTemporal1), "id") = -1 Then
          GRID.TextMatrix(Row, COL_SUBCLIENTE) = "-"
          
          If cced > 1 Then
            GRID.Col = cced
            GRID.ColAllowEdit(cced) = True
            GRID.SetFocus
          End If
        Else
          'Por defecto si existen SubClientes cargarlos en un Listbox:
          C1 = Mid(GRID.TextMatrix(Row, COL_CLIENTE), 1, 6)
          If C1 <> "" Then
            CargarComboSubCliente C1
            If La_Columna_GRID(GRID, "CEDULA") > 0 Then
              GRID.Col = La_Columna_GRID(GRID, "CEDULA")
              GRID.SetFocus
              SendKeys "{ENTER}"
            End If
            If GRID.ColWidth(COL_SUBCLIENTE) < COL_ANCHO_MAX_SUB Then
              GRID.ColWidth(COL_SUBCLIENTE) = COL_ANCHO_MAX_SUB
            End If
          End If
        End If
        
        '-Chequear si el Cliente Principal TIENE tabla propia de Personas:
        st = La_Tabla_Actual_Personas(Mid(GRID.TextMatrix(Row, COL_CLIENTE), 1, 6), "-")
        
        If st <> "" Then
          GRID.AddLookup COL_SUBCLIENTE, "-"
        End If
       
      End If
      
      
      '-Antes de introducir la cedula, buscar las indicaciones especiales
      '-del cliente/subcliente y mostrarlo al operador:
      s = Trim(LAS_INDICACIONES_CLIENTE(GRID.TextMatrix(Row, 3), GRID.TextMatrix(Row, 4)))
      If s <> "" Then
        Load fIndicaciones
        fIndicaciones.Text1.Text = s
        fIndicaciones.Show vbModal
      End If

        
                           
    Case cced:
      If Len(Trim(s)) > 8 Then
         MsgBox "El número de cédula no debe exceder de los 8 dígitos", vbExclamation
         GRID.SetFocus
         Exit Sub
      End If
      s = Trim(GRID.TextMatrix(Row, cced))
      C1 = Trim(GRID.TextMatrix(Row, COL_CLIENTE))
      C2 = Trim(GRID.TextMatrix(Row, COL_SUBCLIENTE))
      C3 = Modulo.La_Tabla_Actual_Personas(C1, C2)
      If C3 = "" Then
        MsgBox "El Cliente No Tiene Tabla de Personas Creada, Revise...", vbCritical, "Información"
      Else
      
        Formatear_Cedula s
        GRID.TextMatrix(Row, cced) = s
        
        idced = ExisteCedulaEnPersonas(C3, s)
        LimpiarCeldasPersona Row
                
        '''''Modulo.Inicializar_Personas_Marca C3, " "
        
        If idced > 0 Then
          AnexarColumnasDePersonas C3, idced, GRID, Row, False
        Else
          'MsgBox "Cédula [" & s & "] NO Existe...", vbCritical, "Información"
          AnexarColumnasDePersonas C3, idced, GRID, Row, True
        End If
        
        Chequear_COLUMNAS_para_Edicion GRID.Row
               
        'Ajustar_Columna_GRID grid
        
        
      End If
      
    Case co:
      
      'sOb = Trim(GRID.TextMatrix(Row, Col))
      'If sOb = "" Then
      '  MsgBox "Debe Indicar la Observación para poder Guardar el Movimiento...", vbCritical, "Información"
      '  GRID.SetFocus
      '  SendKeys "{F2}"
      'End If
      
      'bGuardar.Enabled = Se_Puede_GUARDAR_Movimiento()
      
  End Select
  
  If GRID.Col < GRID.Cols Then GRID.Col = GRID.Col + 1
    
End Sub

Function Se_Puede_GUARDAR_Movimiento() As Boolean
  Dim Fila As Integer
  Dim i As Integer, HayValor As Boolean
  HayValor = False
  If GRID.Rows >= 1 Then
    Fila = GRID.Row
    If Fila >= 1 And Fila <= GRID.Rows Then
      'Revision de todas las celdas deben contener VALORES
      i = 2  'la 1 es el ID x defecto es ""
      HayValor = True
      Do While i < GRID.Cols And HayValor
        HayValor = IIf(Trim(GRID.TextMatrix(Fila, i)) = "", False, True)
        i = i + 1
      Loop
    End If
  End If
  Se_Puede_GUARDAR_Movimiento = HayValor
End Function

Private Function Fue_Inicializada(sTablaCliente As String) As Boolean
'  Dim i As Integer
'  Dim FI As Boolean
'  i = 0
'  FI = False
'  Do While i < lInicializadas.ListCount And Not FI
'    If UCase(lInicializadas.List(i)) = UCase(sTablaCliente) Then
'      FI = True
'    End If
'    i = i + 1
'  Loop
'  Fue_Inicializada = FI
End Function


Private Sub Chequear_COLUMNAS_para_Edicion(Fila As Integer)
  Dim cc As Integer, co As Integer, Col As Integer
  Dim C1 As String, C2 As String, tc As String
  Dim ct As String
  cc = La_Columna_GRID(GRID, COL_TITULO_CEDULA)
  co = La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES)
  
  'Si ya Esta Grabado... No puede Modificar!
  If Trim(GRID.TextMatrix(Fila, 1)) = "" Then
  
    'Si el campo a editar es la cedula, determinar por el cliente y
    'sub-cliente la tabla de personas a usar para inicializar la MARCA interna
    
    C1 = GRID.TextMatrix(Fila, COL_CLIENTE)
    C2 = GRID.TextMatrix(Fila, COL_SUBCLIENTE)
       
    For Col = cc + 1 To co - 1
    
      GRID.ColAllowEdit(Col) = True 'Por defecto es cada columna es EDITABLE
    
      'Si es columna editable DESPUES de la CEDULA - Chequear q sea CAMPO de la TABLA INTERNA:
      ct = GRID.TextMatrix(0, Col)
      If Not Campo_de_la_Tabla(La_Tabla_Actual_Personas(C1, C2), ct) Then
        GRID.ColAllowEdit(Col) = False
        GRID.TextMatrix(Fila, Col) = "*NO*"
      Else
        GRID.ColEditWidth(Col) = Modulo.Size_Str_Campo(La_Tabla_Actual_Personas(C1, C2), GRID.TextMatrix(0, Col))
      End If
      
    Next Col
    
  End If
  
End Sub


Private Sub GRID_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  Dim cc As Integer, co As Integer
  Dim C1 As String, C2 As String, tc As String
  Dim ct As String
  Dim s As String
  Dim lMes As Integer
  Dim lAño As Integer
  cc = La_Columna_GRID(GRID, COL_TITULO_CEDULA)
  co = La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES)
  
  'Si ya Esta Grabado... No puede Modificar!
  If Trim(GRID.TextMatrix(Row, 1)) = "" Then
  
    'Si el campo a editar es la cedula, determinar por el cliente y
    'sub-cliente la tabla de personas a usar para inicializar la MARCA interna
    
    If cc = Col Then
    
      If Trim(GRID.TextMatrix(Row, COL_CLIENTE)) = "" Then
        MsgBox "Debe Seleccionar el CLIENTE del Movimiento...", vbCritical, "Información"
        GRID.ColAllowEdit(Col) = False
        Exit Sub
      End If
      
      
'      'Determinar cual es la tabla personas del cliente (pq no tiene subcliente)
'      c1 = Mid(GRID.TextMatrix(Row, COL_CLIENTE), 1, 6)
'      c2 = GRID.TextMatrix(Row, COL_SUBCLIENTE)
'      tc = Modulo.La_Tabla_Actual_Personas(c1, c2)
'      If Not Fue_Inicializada(tc) Then
'        Modulo.Inicializar_Personas_Marca tc, " "
'        lInicializadas.AddItem UCase(tc)
'      End If

    End If
    
    GRID.ColAllowEdit(Col) = True 'Por defecto es EDITABLE
    
    'Si es columna editable DESPUES de la CEDULA - Chequear q sea CAMPO de la TABLA INTERNA:
    If Col > cc And Col < co Then
      C1 = GRID.TextMatrix(Row, COL_CLIENTE)
      C2 = GRID.TextMatrix(Row, COL_SUBCLIENTE)
      ct = GRID.TextMatrix(0, Col)
      If Not Campo_de_la_Tabla(La_Tabla_Actual_Personas(C1, C2), ct) Then
        GRID.ColAllowEdit(Col) = False
        GRID.TextMatrix(Row, Col) = "*NO*"
        If GRID.Col < GRID.Cols Then GRID.Col = GRID.Col + 1
        Exit Sub
      End If
      
      If GRID.ColAllowEdit(Col) Then
        'GRID.ColEditWidth(Col) = Modulo.Size_Str_Campo(La_Tabla_Actual_Personas(C1, C2), GRID.TextMatrix(0, Col))
      End If
      
    End If
    
    If GRID.ColAllowEdit(Col) Then
      GRID.SelectText = True
    Else
      Cancel = True
      'If GRID.Col < GRID.Cols Then GRID.Col = GRID.Col + 1
    End If
    
  If UCase(GRID.TextMatrix(0, Col)) = "VENCE" Then
     If ComboVenceListo = False Then
        ComboVenceListo = True
        'GRID.AddButton 8 ' boton para fecha de vencimiento
        GRID.AddLookup Col, "2 AÑOS" ''UCase(MonthName(Month(Date))) & " " & Year(Date) + 1
        GRID.AddLookup Col, "1 AÑO"  ''UCase(MonthName(Month(Date) + 5)) & " " & Year(Date)
        GRID.AddLookup Col, "6 MESES"  ''UCase(MonthName(Month(Date) + 5)) & " " & Year(Date)
        GRID.AddLookup Col, "3 MESES"  ''UCase(MonthName(Month(Date) + 5)) & " " & Year(Date)
        GRID.AddLookup Col, "PERSONALIZAR"
        sFechaVencimientoFD Mid(GRID.TextMatrix(GRID.Row, COL_CLIENTE), 1, 6)
        GRID.AddLookup Col, lFechaVencimientoFD
        GRID.TextMatrix(GRID.Row, Col) = lFechaVencimientoFD
        GRID.Col = GRID.Col + 1
        'GRID.ColMask(1) = checkmark
     
     End If
     If UCase(GRID.TextMatrix(0, Col)) = "VENCE" Then
      Select Case GRID.TextMatrix(GRID.Row, Col)
         Case "2 AÑOS"
             GRID.TextMatrix(GRID.Row, Col) = UCase(MonthName(Month(Date))) & " " & Year(Date) + 2
         Case "1 AÑO"
             GRID.TextMatrix(GRID.Row, Col) = UCase(MonthName(Month(Date))) & " " & Year(Date) + 1
        Case "6 MESES"
           If Month(Date) <= 6 Then
              GRID.TextMatrix(GRID.Row, Col) = UCase(MonthName(Month(Date) + 5)) & " " & Year(Date)
           Else
              lMes = Month(Date)
              lMes = lMes + 6
              lAño = 0
              If lMes > 12 Then
                 lMes = lMes - 12
                 lAño = 1
              End If
              GRID.TextMatrix(GRID.Row, Col) = UCase(MonthName(lMes)) & " " & Year(Date) + lAño
           End If
        Case "3 MESES"
           If Month(Date) <= 9 Then
              GRID.TextMatrix(GRID.Row, Col) = UCase(MonthName(Month(Date) + 2)) & " " & Year(Date)
           Else
              lMes = Month(Date)
              lMes = lMes + 3
              If lMes > 12 Then
                 lMes = lMes - 12
                 lAño = 1
              End If
              GRID.TextMatrix(GRID.Row, Col) = UCase(MonthName(lMes)) & " " & Year(Date) + lAño
           End If
        Case "PERSONALIZAR"
           frmVencimiento.lColumna = Col
           frmVencimiento.Show vbModal
      End Select
      
      End If
  End If
    
 ' If UCase(GRID.TextMatrix(0, Col)) = "CARGO" And ComboCargoListo = False Then
 '   'GRID.AddButton 8 ' boton para fecha de vencimiento
 '       Dim lRegCargo As New ADODB.Recordset
 '       Dim lSC As String
 '       Dim lsSC As String
 '       Dim lst As String
 '       Dim ls As String
 '       ComboCargoListo = True
 '       lSC = Mid(Trim(GRID.TextMatrix(GRID.Row, 3)), 1, 6)
 '       lsSC = Mid(Trim(GRID.TextMatrix(GRID.Row, 4)), 1, 6)
 '
 '       lst = Modulo.La_Tabla_Actual_Personas(lSC, lsSC)
 '       ls = GRID.TextMatrix(0, GRID.Col)
 '       lRegCargo.Open "select DISTINCT(" & ls & ") from [" & lst & "] order by " & ls, Modulo.DBConexionSQL, adOpenStatic
 '       Do While lRegCargo.EOF = False
 '          If lRegCargo.Fields(ls).Value <> "" Then GRID.AddLookup Col, UCase(lRegCargo.Fields(ls).Value)
 '          lRegCargo.MoveNext
 '       Loop
 ' End If
 
    
    
    If GRID.TextMatrix(0, Col) = "VENCE" Then
      Dim ProximaFecha As Date
      ProximaFecha = Date + 365 'al año siguiente...
      'GRID.TextMatrix(Row, Col) = Format(ProximaFecha, "dd/mm/yyyy")
      ''Clipboard.Clear
      ''Clipboard.SetText Format(ProximaFecha, "dd/mm/yyyy")
      ''GRID.ColWidth
       GRID.ColWidth(8) = 130
    End If
      
    
        
  Else
    Cancel = True
  End If
End Sub

Private Sub GRID_BtnClick(ByVal Row As Long, ByVal Col As Long)
  Dim cc As Integer, co As Integer
  Dim C1 As String, C2 As String, C3 As String, sOb As String
  Dim sIndicaciones As String, s As String
  Dim SC As String, sSC As String
  Dim dTC As Double, dTA As Double, dTS As Double
  ''Dim lFechaVencimientoFD As String ' para que traiga la fecha de vencimiento del formato de diseño
  
  'Si ya Esta Grabado... No puede Modificar!
  If Trim(GRID.TextMatrix(Row, 1)) <> "" Then Exit Sub
  
  
  cc = Modulo.La_Columna_GRID(GRID, "CEDULA")
  co = Modulo.La_Columna_GRID(GRID, "OBSERVACIONES")
  
  SC = Trim(GRID.TextMatrix(GRID.Row, 3))
  sSC = Trim(GRID.TextMatrix(GRID.Row, 4))

  
  
  Select Case Col
    Case COL_SUBCLIENTE:
        cc = Modulo.La_Columna_GRID(GRID, "CEDULA")
        If cc > 1 Then
          GRID.Col = cc
          GRID.ColAllowEdit(cc) = True
          GRID.SetFocus
          'SendKeys "{ENTER}"
        End If

    
    Case COL_CLIENTE:
    
      'Busqueda CLIENTE
      Modulo.vTemporal1 = ""
      Load fBuscarSimple
  
      With fBuscarSimple
        .Combo1.Clear
        .Combo1.AddItem "CODIGO"
        .Combo1.AddItem "NOMBRE"
        .Combo1.ListIndex = 1
    
        .Adodc1.ConnectionString = Modulo.DBConexionSQL.ConnectionString
        .Adodc1.RecordSource = "select codigo, nombre, direccion, telefonos from clientes order by codigo"
        .Adodc1.Caption = "CLIENTES"
        .Adodc1.Refresh
        .DataGrid1.Refresh
        AjustaColumnaDataGrid .DataGrid1, "NOMBRE", 4000
      End With
      DoEvents
  
      Modulo.vTemporal1 = ""
      fBuscarSimple.Show vbModal
   
      If Modulo.vTemporal1 <> "" Then
        'sFechaVencimientoFD Modulo.vTemporal1
        If Len(Modulo.vTemporal1) < 6 Then Modulo.vTemporal1 = Zeros(CLng(Modulo.vTemporal1), 6)
        GRID.TextMatrix(Row, Col) = Modulo.vTemporal1 & " " & Modulo.DBValorStr("clientes", "codigo", Modulo.vTemporal1, "nombre")
        
        If DBValorLng("subclientes", "cliente", CLng(Modulo.vTemporal1), "id") = -1 Then
          GRID.TextMatrix(Row, COL_SUBCLIENTE) = "-"
        Else
          'Por defecto colocar en la celda, el primer sub-cliente:
          GRID.TextMatrix(Row, COL_SUBCLIENTE) = Modulo.Zeros(DBValorLng("subclientes", "cliente", CLng(Modulo.vTemporal1), "id"), 6)
        End If
      
        'sIndicaciones = DBValorVariant("clientes", "codigo", CLng(Modulo.vTemporal1), "indicaciones")
      
        SC = Trim(GRID.TextMatrix(GRID.Row, 3))
        sSC = Trim(GRID.TextMatrix(GRID.Row, 4))
      
        
        sIndicaciones = LAS_INDICACIONES_CLIENTE(SC, sSC)
        
        If sIndicaciones <> "" Then
          Load fIndicaciones
          fIndicaciones.Text1.Text = sIndicaciones
          fIndicaciones.Show vbModal
        End If
        
        If DBValorLng("subclientes", "cliente", CLng(Modulo.vTemporal1), "id") = -1 Then
          GRID.TextMatrix(Row, COL_SUBCLIENTE) = "-"
          GRID.Col = cc 'automatico pasa al campo cedula
          GRID.SetFocus
          
          

          
          
        Else
          'Por defecto si existen SubClientes cargarlos en un Listbox:
          C1 = Trim(GRID.TextMatrix(Row, COL_CLIENTE))
          If C1 <> "" Then
            CargarComboSubCliente C1
            If La_Columna_GRID(GRID, "CEDULA") > 0 Then
              GRID.Col = La_Columna_GRID(GRID, "CEDULA")
            End If
            If GRID.ColWidth(COL_SUBCLIENTE) < COL_ANCHO_MAX_SUB Then
              GRID.ColWidth(COL_SUBCLIENTE) = COL_ANCHO_MAX_SUB
            End If
          End If
        End If
        
        
        '-Chequear si el Cliente Principal TIENE tabla propia de Personas:
        Dim st As String
        st = La_Tabla_Actual_Personas(Mid(GRID.TextMatrix(Row, COL_CLIENTE), 1, 6), "-")
        
        If st <> "" Then
          GRID.AddLookup COL_SUBCLIENTE, "-"
        End If

        
        
        
        
        
        
        
        
        
        
        
        cc = Modulo.La_Columna_GRID(GRID, "CEDULA")
        If cc > 1 Then
          GRID.Col = cc
          GRID.ColAllowEdit(cc) = True
          GRID.SetFocus
          'SendKeys "{ENTER}"
        End If
        
        
        
        
        

      End If

      
      
    Case co:
    
      '----------------------------------
      'Chequear si está en observaciones:
      '----------------------------------
      If GRID.Col = La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES) Then
        s = Trim(GRID.TextMatrix(GRID.Row, COL_CLIENTE))
        If s = "" Then
          MsgBox "Debe Indicar el Cliente a Procesar, Revise...", vbCritical, "Información"
        Else
          cc = La_Columna_GRID(GRID, "CEDULA")
          If cc > 0 Then
            s = Trim(GRID.TextMatrix(GRID.Row, cc))
            If s = "" Then
              MsgBox "Debe Indicar la Cédula de la Persona, Revise...", vbCritical, "Información"
              GRID.SetFocus
            Else
              If Not Se_Puede_GUARDAR_Movimiento() Then
                If MsgBox("Faltan Datos Requeridos para Guardar el Movimiento. ¿Desea Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
              End If
                C1 = Trim(GRID.TextMatrix(GRID.Row, COL_CLIENTE))
                C2 = Trim(GRID.TextMatrix(GRID.Row, COL_SUBCLIENTE))
                C3 = Modulo.La_Tabla_Actual_Personas(C1, C2)
                If C3 <> "" Then
                  If Trim(GRID.TextMatrix(GRID.Row, 1)) = "" Then   'ID por guardar
                  
                    'Load fObservaciones
                    
                    Modulo.vTemporal1 = ""
                    
                    Modulo.vTemporal2 = GRID.TextMatrix(Row, 3)
                    Modulo.vTemporal3 = GRID.TextMatrix(Row, 4)
                    
                    Modulo.fModalResult = Modulo.fModalResultCANCEL
                    Load fObservaciones
        
                    fObservaciones.lCliente.Caption = GRID.TextMatrix(Row, 3)
                    fObservaciones.lSubCliente.Caption = GRID.TextMatrix(Row, 4)
                    
                    dTC = 0#
                    dTA = 0#
                    dTS = 0#
                    
                    Modulo.Resumen_Cuenta_Cliente GRID.TextMatrix(Row, 3), _
                                                  GRID.TextMatrix(Row, 4), _
                                                  dTC, dTA, dTS

                    fObservaciones.lTC.Caption = Format(dTC, "#,0.00")
                    fObservaciones.lTA.Caption = Format(dTA, "#,0.00")
                    fObservaciones.lTS.Caption = Format(dTS, "#,0.00")
                    fObservaciones.EsActualizar = False
                    fObservaciones.Show vbModal
        
                    If Modulo.fModalResult = Modulo.fModalResultOK Then
                      If Modulo.vTemporal1 <> "" Then GRID.TextMatrix(Row, Col) = Trim(Modulo.vTemporal1)
                       
                      GuardarMovimientos GRID.Row
                          
                      Marcar_Personas_CEDULA C3, s
                      MsgBox "Registro de Movimiento Guardado...", vbInformation, "Información"
                    End If
                  End If
                Else
                  MsgBox "No hay Tabla Creada para Personas del Cliente...", vbCritical, "Información"
                End If
              '''End If
            End If
          End If
        End If
      End If
  End Select
End Sub

Private Sub sFechaVencimientoFD(argCliente As String)
   Dim lCn As New ADODB.Connection
   Dim lReg As New ADODB.Recordset
   lCn.ConnectionString = Modulo.DBConexionSQL
   lCn.Open

   Set lReg = lCn.Execute("Select vencimiento from formatodisenodetalle where cliente=" & argCliente)
   If lReg.EOF = False Then
      If IsNull(lReg!vencimiento) = False Then
         lFechaVencimientoFD = UCase(MonthName(Month(lReg!vencimiento)) & " " & Year(lReg!vencimiento))
      Else
         lFechaVencimientoFD = ""
      End If
   Else
      lFechaVencimientoFD = ""
   End If

End Sub

Private Sub lAnchos_Click()
'  If lAnchos.ListIndex >= 0 Then
'    lCampos.ListIndex = lAnchos.ListIndex
'    lTipos.ListIndex = lAnchos.ListIndex
'  End If
End Sub

Private Sub lCampos_Click()
'  If lCampos.ListIndex >= 0 Then
'    lTipos.ListIndex = lCampos.ListIndex
'    lAnchos.ListIndex = lCampos.ListIndex
'  End If
End Sub


Private Sub LeerTablaPersonas()
'  If lTablas.ListCount <= 0 Then Exit Sub
'  If lTablas.ListIndex >= 0 Then
'    lCreadas.ListIndex = lTablas.ListIndex
'    Mostrar_Estructura
'    Mostrar_Personas_Registradas
'    'RefrescarPersonas
'  End If
End Sub




Private Sub Mostrar_Personas_Registradas()
'  Dim i As Integer
'  Dim s As String
'  Dim c As String
'
'  i = -1
'  If lTablas.ListIndex >= 0 Then i = lTablas.ListIndex
'  If lCreadas.ListIndex >= 0 Then i = lCreadas.ListIndex
'  If i = -1 Then
'    MsgBox "Debe Seleccionar Tabla de Personas primero.", vbCritical, "Información"
'  Else
'    '--Ordenar por defecto: CEDULA (campo del sistema):
'    c = ""
'    If lCampos.ListCount > 0 Then c = lCampos.List(0)
'    s = "SELECT * FROM [" & lTablas.List(i) & "]" & IIf(c <> "", " ORDER BY " & c, "")
'    s = "SELECT * FROM [" & lTablas.List(i) & "] ORDER BY cedula"
'
'    'If Adodc1.Recordset.State <> adStateClosed Then Adodc1.Recordset.Close
'
'    'Load fMensaje
'    'fMensaje.Label1.Caption = "Accesando Listado del Cliente, Espere..."
'    'fMensaje.Show
'    'DoEvents
'
'    Adodc1.ConnectionString = Modulo.DBConexionSQL.ConnectionString
'    Adodc1.RecordSource = s
'    Adodc1.Refresh
'    'DataGrid1.Refresh
'    'DoEvents
'
'    'CargarAdoEnFG
'    HFG.Clear
'    HFG.Rows = 2
'
'    With HFG
'      .FixedCols = 0
'      '.FixedRows = 1
'      .AllowBigSelection = False
'      .FillStyle = flexFillSingle
'      .Redraw = True
'    End With
'
'    CargarAdoEnFG HFG
'
'    'Set HFG.DataSource = Adodc1.Recordset
'    'AlinearHFG
'
'
'    'Unload fMensaje
'  End If
  
End Sub

Private Sub CargarAdoEnFG(FG As MSHFlexGrid)
  Dim f As Integer, c As Integer
  Dim s As String
  
  Dim width As Integer
  Dim aWidths() As Single
  
  Dim maxWidth As Single, celdaText As String
  Dim saveFont As StdFont, oldScaleMode As Integer
 
  FG.Clear
  If Adodc1.Recordset.RecordCount <= 0 Then Exit Sub
  
  FG.Rows = Adodc1.Recordset.RecordCount + 1
  FG.FixedRows = 1
  FG.Cols = Adodc1.Recordset.Fields.Count
  FG.FixedCols = 0
  For c = 0 To Adodc1.Recordset.Fields.Count - 1
    FG.TextMatrix(0, c) = Adodc1.Recordset.Fields(c).Name
  Next c
  
  ReDim aWidths(Adodc1.Recordset.Fields.Count)
  
  ' Guardamos la fuente del DataGrid para luego reestablecerla
  Set saveFont = FG.Parent.Font
  Set FG.Parent.Font = FG.Font
    
  ' Ajustar el ScaleMode en vbTwips para el formulario
  oldScaleMode = FG.Parent.ScaleMode
  FG.Parent.ScaleMode = vbTwips
  
  For f = 1 To Adodc1.Recordset.RecordCount
    For c = 0 To Adodc1.Recordset.Fields.Count - 1
      s = ""
      If Not IsNull(Adodc1.Recordset.Fields(c).value) Then
        s = Adodc1.Recordset.Fields(c).value
      End If
        
      FG.TextMatrix(f, c) = Trim(s)
      
      celdaText = FG.TextMatrix(f, c)
            
      'Almacena el Ancho del texto de la celda del Datagrid
      width = FG.Parent.TextWidth(celdaText) + 150
                
      'Si el ancho de la celda es mayor se actualiza la variable maxWidth _
       y se establece el ancho de la columna
      If width > aWidths(c) Then
        aWidths(c) = width
        FG.ColWidth(c) = width
      End If
      
    Next c
    Adodc1.Recordset.MoveNext
  Next f
  
  'For c = 0 To Adodc1.Recordset.Fields.Count - 1
  '  FG.ColWidth(c) = aWidths(c)
  'Next
    
  'restablecemos la fuente del DataGrid y el scaleMode
  Set FG.Parent.Font = saveFont
  FG.Parent.ScaleMode = oldScaleMode
  
  Adodc1.Recordset.MoveFirst
End Sub

'Private Sub AlinearHFG()
'  Dim f As Integer, c As Integer
'  Dim s As String
'
'  Dim width As Integer
'  Dim aWidths() As Single
'
'  Dim maxWidth As Single, celdaText As String, encabezadotext As String
'  Dim saveFont As StdFont, oldScaleMode As Integer
'
'  If Adodc1.Recordset.RecordCount <= 0 Then Exit Sub
'
'  If HFG.Rows - 1 <> Adodc1.Recordset.RecordCount Then Exit Sub 'no lleno!
'
'
'  ReDim aWidths(Adodc1.Recordset.Fields.Count)
'
'  ' Guardamos la fuente del DataGrid para luego reestablecerla
'  Set saveFont = HFG.Parent.Font
'  Set HFG.Parent.Font = HFG.Font
'
'  ' Ajustar el ScaleMode en vbTwips para el formulario
'  oldScaleMode = HFG.Parent.ScaleMode
'  HFG.Parent.ScaleMode = vbTwips
'
'  If Adodc1.Recordset.RecordCount > 0 Then Adodc1.Recordset.MoveFirst Else Exit Sub
'
'
'  For f = 1 To Adodc1.Recordset.RecordCount
'    For c = 0 To Adodc1.Recordset.Fields.Count - 1
'      s = ""
'      If Not IsNull(Adodc1.Recordset.Fields(c).Value) Then
'        s = Adodc1.Recordset.Fields(c).Value
'      End If
'
'      HFG.TextMatrix(f, c) = Trim(s)
'
'      celdaText = HFG.TextMatrix(f, c)
'
'      'Almacena el Ancho del texto de la celda del Datagrid
'      width = HFG.Parent.TextWidth(celdaText) + 200
'
'      'ahora si es menor al Titulo del encabezado... tomar el encabezado!
'      encabezadotext = HFG.TextMatrix(0, c)
'      If HFG.Parent.TextWidth(encabezadotext) + 200 > width Then
'        width = HFG.Parent.TextWidth(encabezadotext) + 200
'      End If
'
'
'      'Si el ancho de la celda es mayor se actualiza la variable maxWidth _
'       y se establece el ancho de la columna
'      If width > aWidths(c) Then
'        aWidths(c) = width
'        HFG.ColWidth(c) = width
'      End If
'
'    Next c
'    Adodc1.Recordset.MoveNext
'  Next f
'
'  For c = 0 To Adodc1.Recordset.Fields.Count - 1
'    HFG.ColWidth(c) = aWidths(c)
'  Next
'
'  'restablecemos la fuente del DataGrid y el scaleMode
'  'Set HFG.Parent.Font = saveFont
'  'HFG.Parent.ScaleMode = oldScaleMode
'
'  Adodc1.Recordset.MoveFirst
'End Sub



Private Sub Mostrar_Estructura()
'  Dim r As New ADODB.Recordset
'  Dim s As String, s1 As String
'  Dim i As Integer
'
'  If lTablas.ListIndex >= 0 Then
'
'    s = lTablas.List(lTablas.ListIndex)
'    s = "select count(*) from [" & s & "]"
'    r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
'    Label26.Caption = "-"
'    If Not r.EOF Then Label26.Caption = CStr(r.Fields(0).Value)
'    r.Close
'
'    s = lTablas.List(lTablas.ListIndex)
'    s = "select * from [" & s & "]"
'    r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
'    '--Extraer toda la info de la tabla estructura:
'    lCampos.Clear
'    lTipos.Clear
'    lAnchos.Clear
'
'    cCampos.Clear
'
'
'    For i = 0 To r.Fields.Count - 1
'      lCampos.AddItem r.Fields(i).Name
'      s = ""
'      Select Case r.Fields(i).Type
'        Case adChar: s = "CHAR"
'        Case adInteger: s = "INTEGER"
'        Case adDBTimeStamp: s = "DATETIME"
'        Case adDouble: s = "FLOAT"
'      End Select
'      lTipos.AddItem s
'      s1 = ""
'
'      If s = "CHAR" Then s1 = CStr(r.Fields(i).DefinedSize)
'
'      lAnchos.AddItem s1
'
'      cCampos.AddItem r.Fields(i).Name
'    Next i
'
'    If cCampos.ListCount > 0 Then
'      If Modulo.EXISTE_CAMPO(r, "CEDULA") = True Then
'        cCampos.ListIndex = Modulo.Buscar_Combo(cCampos, "CEDULA")
'      End If
'    End If
'
'    r.Close
'    Set r = Nothing
'  End If
  
    
End Sub

Private Sub lTipos_Click()
'  If lTipos.ListIndex >= 0 Then
'    lCampos.ListIndex = lTipos.ListIndex
'    lAnchos.ListIndex = lTipos.ListIndex
'  End If
End Sub

Private Sub Cargar_Clientes_COMBO(xCombo As ComboBox)
  'Dim r As New ADODB.Recordset
  Dim s As String
  Dim l As Integer
  Dim sCod As String
  
  If RClientes.State <> adStateClosed Then RClientes.Close
  
  s = "SELECT * FROM Clientes ORDER BY Codigo"
  
  RClientes.Open s, DBConexionSQL, adOpenKeyset, adLockReadOnly
  
  xCombo.Clear
    
  Do While Not RClientes.EOF
    s = Zeros(RClientes.Fields("codigo").value, 6) & " : " & Trim(RClientes.Fields("nombre").value)
    xCombo.AddItem s
    RClientes.MoveNext
  Loop
  
  RClientes.Close
End Sub


Private Sub CargarFGDiario()
'  Dim rDiario As New ADODB.Recordset
'  Dim rCliente As New ADODB.Recordset
'  Dim rTablas As New ADODB.Recordset
'  Dim rPersona As New ADODB.Recordset
'
'  Dim s As String
'  Dim sFecha As String
'  Dim i As Integer, l As Integer
'
'  Dim sTablaCliente As String
'
'  sFecha = Format(Date, "dd/mm/yyyy")
'  s = "select * from Diario where fechahora >= '" & sFecha & "' and fechahora <= '" & sFecha & "' order by fechahora"
'
'  rDiario.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
'
'  FG1.Clear
'  FG1.Rows = 2
'  FG1.Cols = 5
'
'
'  FG1.TextMatrix(0, 0) = "FECHA"
'  FG1.TextMatrix(0, 1) = "CLIENTE"
'  FG1.TextMatrix(0, 2) = "SUBCLIENTE"
'  FG1.TextMatrix(0, 3) = "TABLA"
'  FG1.TextMatrix(0, 4) = "CEDULA"
'
'  l = 1
'
'  Do While Not rDiario.EOF
'
'    FG1.TextMatrix(l, 0) = Format(rDiario.Fields("FECHAHORA").Value, "dd/mm/yyyy hh:mm am/pm")
'    FG1.TextMatrix(l, 1) = Trim(rDiario.Fields("CLIENTE").Value)
'    FG1.TextMatrix(l, 2) = Trim(rDiario.Fields("SUBCLIENTE").Value)
'    FG1.TextMatrix(l, 3) = Trim(rDiario.Fields("TABLA").Value)
'    FG1.TextMatrix(l, 4) = Trim(rDiario.Fields("CEDULA").Value)
'
'    sTablaCliente = Trim(rDiario.Fields("TABLA").Value)
'
'    s = "select * from [" & sTablaCliente & "]"
'    rPersona.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
'    For i = 0 To rPersona.Fields.Count - 1
'      If EsColumnaMostrable(UCase(rPersona.Fields(i).Name)) = True Then
'        FG1.Cols = FG1.Cols + 1
'        FG1.TextMatrix(0, FG1.Cols) = rPersona.Fields(i).Name
'        FG1.TextMatrix(l, FG1.Cols) = rPersona.Fields(i).Value
'      End If
'    Next i
'    rPersona.Close
'
'
'    rDiario.MoveNext
'  Loop
'
'  rDiario.Close
'
'  'rCliente.Close
'  'rTablas.Close
'  'rPersona.Close
'
'  Set rDiario = Nothing
'  Set rCliente = Nothing
'  Set rTablas = Nothing
'  Set rPersona = Nothing
End Sub

Private Sub CargarGridEditorDiario()
  Dim rDiario As New ADODB.Recordset
  Dim RClientes As New ADODB.Recordset
  Dim RSubClientes As New ADODB.Recordset
  Dim rTablas As New ADODB.Recordset
  Dim rPersona As New ADODB.Recordset
  
  Dim s As String, s1 As String, s2 As String
  Dim sFecha As String, sCampo As String
  Dim i As Integer, l As Integer, l1 As Long, l2 As Long
  Dim bExiste As Boolean
  
  Dim sTablaCliente As String
  
  Load fMensaje
  fMensaje.Label1.Caption = "Leyendo Movimientos del Día " & Format(Date, "dd/mm/yyyy") & ", Espere..."
  fMensaje.Show
  
  s = "select * from clientes order by codigo"
  RClientes.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  
  s = "select * from subclientes order by id"
  RSubClientes.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  
  
     '' AND OBSERVACIONES <> 'PORLOTES'
  sFecha = Format(Date, "yyyy/mm/dd")
  s = "select * from Diario where fecha = '" & sFecha & "' AND OBSERVACIONES <> 'PORLOTES' order by id"
  
  rDiario.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  
  GRID.Clear
  GRID.Rows = 1
  GRID.Cols = 5 + NCol
  
  l = 1
  
  If Not rDiario.EOF Then
    fMensaje.Label1.Caption = "Leyendo Movimientos del Día " & Format(Date, "dd/mm/yyyy") & ", Espere... [" & CStr(l) & "/" & CStr(rDiario.RecordCount) & "]"
    fMensaje.Show
  End If
  DoEvents
  
  
  With GRID
    .AutoNewRow = False
    .RowHeader = False
    .TextMatrix(0, 0 + NCol) = "Check"
    .TextMatrix(0, 1 + NCol) = "ID"
    .TextMatrix(0, 2 + NCol) = "FECHA-MOV"
    .TextMatrix(0, 3 + NCol) = "CLIENTE"
    .TextMatrix(0, 4 + NCol) = "SUBCLIENTE"
    '.TextMatrix(0, 5) = "TABLA"
    .TextMatrix(0, 5 + NCol) = "CEDULA"
    .ColWidth(0 + NCol) = 20
    .ColWidth(1 + NCol) = 30  'ID
    .ColWidth(2 + NCol) = 70  'Fecha
    .ColWidth(3 + NCol) = COL_ANCHO_MIN_CLIENTE 'Cliente
    .ColWidth(4 + NCol) = COL_ANCHO_MIN_SUB    'Subcliente
    '.ColWidth(5) = 120  'Tabla
    .ColWidth(5 + NCol) = 100 'Cedula
    
    'Columna COMBOBOX de CLIENTES:
'    rCliente.Open "select * from clientes order by codigo", Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
'    i = 1
'    Do While Not rCliente.EOF
'      .AddLookup 2, Zeros(rCliente.Fields("Codigo").Value, 6) & " " & rCliente.Fields("Nombre").Value
'      rCliente.MoveNext
'      i = i + 1
'    Loop
'    rCliente.Close

     .AddButton COL_CLIENTE
     .ColAllowEdit(COL_CLIENTE) = False
     
     '.AddLookup COLSUBCLIENTE
     .ColAllowEdit(COL_SUBCLIENTE) = True
     
         
    
  End With
  
  l = 1
  
  Do While Not rDiario.EOF
  
    GRID.TextMatrix(l, 1 + NCol) = CStr(rDiario.Fields("ID").value)
    GRID.TextMatrix(l, 2 + NCol) = Modulo.FechaNormal(rDiario.Fields("FECHA").value) & " " & _
                            Trim(rDiario.Fields("HORA").value)
    GRID.TextMatrix(l, 3 + NCol) = Zeros(rDiario.Fields("CLIENTE").value, 6)
    
      l1 = rDiario.Fields("CLIENTE").value
      If RClientes.RecordCount > 0 Then
        RClientes.MoveFirst
        RClientes.Find "codigo = '" & CStr(l1) & "'"
        If Not RClientes.EOF Then
          s = Zeros(l1, 6) & " " & Trim(RClientes.Fields("NOMBRE").value)
          GRID.TextMatrix(l, 3 + NCol) = s
        End If
      End If
      
    GRID.TextMatrix(l, 4 + NCol) = Trim(rDiario.Fields("SUBCLIENTE").value)
    
      l1 = rDiario.Fields("CLIENTE").value
      l2 = rDiario.Fields("SUBCLIENTE").value
    
      If l2 <> 0 Then
        If RSubClientes.RecordCount > 0 Then
          RSubClientes.MoveFirst
          bExiste = False
          Do While Not RSubClientes.EOF And Not bExiste
            If RSubClientes.Fields("cliente").value = l1 And _
               RSubClientes.Fields("id").value = l2 Then
               GRID.TextMatrix(l, 4 + NCol) = Zeros(l2, 6) & " " & Trim(RSubClientes.Fields("NOMBRE").value)
               bExiste = True
            End If
            RSubClientes.MoveNext
          Loop
        End If
      Else
        GRID.TextMatrix(l, 4 + NCol) = "-"
      End If
    
    'GRID.TextMatrix(l, 5) = Trim(rDiario.Fields("TABLA").Value)
    GRID.TextMatrix(l, 5 + NCol) = Trim(rDiario.Fields("CEDULA").value)
       
    sTablaCliente = Trim(rDiario.Fields("TABLA").value)
    
    s = "select * from [" & sTablaCliente & "] where cedula = '" & Trim(rDiario.Fields("CEDULA").value) & "' order by id"
    rPersona.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    For i = 0 To rPersona.Fields.Count - 1
      sCampo = UCase(rPersona.Fields(i).Name)
      If EsColumnaMostrable(sCampo) = True Then
        If Modulo.Existe_Columna_GRID(GRID, sCampo) = True Then
          If Not rPersona.EOF Then
            GRID.TextMatrix(l, Modulo.La_Columna_GRID(GRID, sCampo)) = IIf(IsNull(rPersona.Fields(sCampo).value), "", Trim(rPersona.Fields(sCampo).value))
          End If
        Else
          GRID.Cols = GRID.Cols + 1
          GRID.TextMatrix(0, GRID.Cols) = sCampo
          If sCampo = "CARGO" Or sCampo = "VENCE" Then
             GRID.ColWidth(GRID.Cols) = 100
          End If
          If Not rPersona.EOF Then
            GRID.TextMatrix(l, GRID.Cols) = Trim(rPersona.Fields(sCampo).value)
          End If
        End If
      End If
    Next i
        
    rPersona.Close
        
    'Por ultimo mostrar las Observaciones:
'    If EsColumnaMostrable(COL_TITULO_OBSERVACIONES) = True Then
'      If Not Modulo.Existe_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES) Then
'        GRID.Cols = GRID.Cols + 1
'        GRID.TextMatrix(0, GRID.Cols) = COL_TITULO_OBSERVACIONES
'      End If
'
'      If Modulo.Existe_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES) = True Then
'        If Not rDiario.EOF Then
'          GRID.TextMatrix(l, Modulo.La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES)) = IIf(IsNull(rDiario.Fields(COL_TITULO_OBSERVACIONES).Value), "", Trim(rDiario.Fields(COL_TITULO_OBSERVACIONES).Value))
'          'GRID.TextMatrix(l, GRID.Cols) = IIf(IsNull(rDiario.Fields(COL_TITULO_OBSERVACIONES).Value), "", Trim(rDiario.Fields(COL_TITULO_OBSERVACIONES).Value))
'        End If
'      End If
'    End If
    
    rDiario.MoveNext
    
    If Not rDiario.EOF Then
      GRID.Rows = GRID.Rows + 1
      l = l + 1
      fMensaje.Label1.Caption = "Leyendo Movimientos del Día " & Format(Date, "dd/mm/yyyy") & ", Espere... [" & CStr(l) & "/" & CStr(rDiario.RecordCount) & "]"
      fMensaje.Show
      DoEvents
    End If
  Loop
  
  GRID.Row = GRID.Rows
   
  rDiario.Close
  RClientes.Close
  RSubClientes.Close
  'rTablas.Close
  'rPersona.Close
  Unload fMensaje
  GRID.Row = 1 'para colocar el cursor en la primera posición y que se puedan ver todos los registros del dia desde el principio
  Set rDiario = Nothing
  Set RClientes = Nothing
  Set RSubClientes = Nothing
  Set rTablas = Nothing
  Set rPersona = Nothing
  
  
  'Modulo.La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES)
  
  
End Sub


Private Function EsColumnaMostrable(sColumna As String) As Boolean
  EsColumnaMostrable = False
  If sColumna <> "ID" And _
     sColumna <> "FOTO" And _
     sColumna <> "FECHA" And _
     sColumna <> "CONTADOR" And _
     sColumna <> "TIENE_FOTO" And _
     sColumna <> "CREACION" And _
     sColumna <> "MARCA" Then
     EsColumnaMostrable = True
  End If
End Function

Private Sub GRID_DblClick()
  Dim f As Integer, c As Integer
  Dim sCampo As String, sValor As String, sTabla As String
  Dim sCedula As String
  Dim s As String
  
  f = GRID.Row 'fila
  c = GRID.Col 'columna
  
  If f >= 1 And f <= GRID.Rows Then
  
    sValor = GRID.TextMatrix(f, c)
    
    If Trim(GRID.TextMatrix(f, 1)) <> "" Then 'ya es registro guardado
    
      sCampo = GRID.TextMatrix(0, c)
      
      If sCampo = COL_TITULO_CEDULA Then
      
        MsgBox "No se puede Modificar el Campo clave CEDULA...", vbCritical, "Información"
        
        Exit Sub
        
      End If
      
      sCedula = GRID.TextMatrix(f, COL_CEDULA)
      
      sTabla = Modulo.La_Tabla_Actual_Personas(GRID.TextMatrix(f, 3), GRID.TextMatrix(f, 4))
      
      If Modulo.Campo_de_la_Tabla(sTabla, sCampo) Then
                
        sValor = InputBox("Indique " & sCampo & ":", "Editar Datos", sValor)
      
        If sValor <> "" Then
      
          GRID.TextMatrix(f, c) = sValor
        
          s = "update [" & sTabla & "] " & _
              "set " & sCampo & " = '" & sValor & "' " & _
              "where cedula = '" & sCedula & "'"
            
          Modulo.ExecSQL s
          
        End If
          
      End If
      
    End If
    
  End If
    
End Sub

Private Sub GRID_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
  Dim C1 As String, C2 As String, C3 As String, s As String
  Dim cc As Integer
  
  Dim dTC As Double, dTA As Double, dTS As Double

  
'  If KeyCode = vbKeyF2 Then
'    If GRID.ColAllowEdit(GRID.Col) Then
'      'KeyCode = 0
'      SendKeys "{ENTER}"
'    End If
'
'  Else
  
    
    If KeyCode = vbKeyReturn Then
    
      KeyCode = 0
    
    
      '----------------------------------
      'Chequear si está en observaciones:
      '----------------------------------
      If GRID.Col = La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES) Then
        s = Trim(GRID.TextMatrix(GRID.Row, COL_CLIENTE))
        If s = "" Then
          MsgBox "Debe Indicar el Cliente a Procesar, Revise...", vbCritical, "Información"
        Else
          cc = La_Columna_GRID(GRID, "CEDULA")
          If cc > 0 Then
            s = Trim(GRID.TextMatrix(GRID.Row, cc))
            If s = "" Then
              MsgBox "Debe Indicar la Cédula de la Persona, Revise...", vbCritical, "Información"
              GRID.SetFocus
            Else
              If Not Se_Puede_GUARDAR_Movimiento() Then
                MsgBox "Faltan Datos Requeridos para Guardar el Movimiento, Revise...", vbCritical, "Información"
              Else
                C1 = Trim(GRID.TextMatrix(GRID.Row, COL_CLIENTE))
                C2 = Trim(GRID.TextMatrix(GRID.Row, COL_SUBCLIENTE))
                C3 = Modulo.La_Tabla_Actual_Personas(C1, C2)
                If C3 <> "" Then
                  If Trim(GRID.TextMatrix(GRID.Row, 1)) = "" Then   'ID por guardar
                  
                    'Load fObservaciones
                    Modulo.vTemporal1 = ""
                    
                    Modulo.vTemporal2 = GRID.TextMatrix(GRID.Row, 3)
                    Modulo.vTemporal3 = GRID.TextMatrix(GRID.Row, 4)

                    
                    Modulo.fModalResult = Modulo.fModalResultCANCEL
                    Load fObservaciones
        
                    fObservaciones.lCliente.Caption = GRID.TextMatrix(GRID.Row, 3)
                    fObservaciones.lSubCliente.Caption = GRID.TextMatrix(GRID.Row, 4)
                    
                    
                    dTC = 0#
                    dTA = 0#
                    dTS = 0#
                    
                    Modulo.Resumen_Cuenta_Cliente GRID.TextMatrix(GRID.Row, 3), _
                                                  GRID.TextMatrix(GRID.Row, 4), _
                                                  dTC, dTA, dTS

                    fObservaciones.lTC.Caption = Format(dTC, "#,0.00")
                    fObservaciones.lTA.Caption = Format(dTA, "#,0.00")
                    fObservaciones.lTS.Caption = Format(dTS, "#,0.00")
                    fObservaciones.EsActualizar = False
                    fObservaciones.Show vbModal
        
                    If Modulo.fModalResult = Modulo.fModalResultOK Then
                      If Modulo.vTemporal1 <> "" Then GRID.TextMatrix(GRID.Row, GRID.Col) = Trim(Modulo.vTemporal1)
                       
                      GuardarMovimientos GRID.Row
                          
                      Marcar_Personas_CEDULA C3, s
                      MsgBox "Registro de Movimiento Guardado...", vbInformation, "Información"
                    End If
                  End If
                Else
                  MsgBox "No hay Tabla Creada para Personas del Cliente...", vbCritical, "Información"
                End If
              End If
            End If
          End If
        End If
      End If
      
    End If
End Sub

Private Sub VentanaMostrarCampos()
  Dim s As String, s1 As String, st As String
  Dim SC As String, sSC As String
  Dim r As New ADODB.Recordset
  
  If GRID.Row >= 1 Then
    SC = Mid(Trim(GRID.TextMatrix(GRID.Row, 3)), 1, 6)
    sSC = Mid(Trim(GRID.TextMatrix(GRID.Row, 4)), 1, 6)
    
    st = Modulo.La_Tabla_Actual_Personas(SC, sSC)
        
    If Trim(st) <> "" Then
    
      Load fMostrarCampos
      With fMostrarCampos
        s = GRID.TextMatrix(0, GRID.Col)
        s1 = "select DISTINCT(" & s & ") from [" & st & "] order by " & s
        On Error Resume Next
        r.Open s1, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
        If Err.Number <> 0 Then
          MsgBox "Columna [" & s & "] No Disponible para el Cliente...", vbCritical, "Información"
          Unload fMostrarCampos
          r.Close
          Set r = Nothing
          Exit Sub
        End If
        .Caption = s
        .List1.Clear
        Do While Not r.EOF
          If r.Fields(s).value <> "" Then .List1.AddItem UCase(r.Fields(s).value)
          r.MoveNext
        Loop
        r.Close
        Set r = Nothing
        Modulo.vTemporal1 = ""
        .Show vbModal
        If Modulo.vTemporal1 <> "" Then
          GRID.TextMatrix(GRID.Row, GRID.Col) = Modulo.vTemporal1
        End If
      End With
      
    End If
  End If
    
End Sub

Private Sub GRID_KeyPress(ByVal KeyAscii As Integer)
  Dim co As Integer, cc As Integer
  Dim sOb As String, s As String, C1 As String, C2 As String, C3 As String
  
  Dim dTC As Double, dTA As Double, dTS As Double
  
  If KeyAscii = vbKeyEscape Then
    If Trim(GRID.TextMatrix(GRID.Rows, 1)) = "" Then
      CargarGridEditorDiario
      GRID.TextMatrix(GRID.Rows, 2) = Format(Now, "dd/mm/yyyy hh:mm ampm")
      GRID.ColAllowEdit(1) = False  'No Editar ID
    End If
    Exit Sub
  End If

    

  
  If Not fObservaciones Is Nothing Then Exit Sub
  
  co = La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES)
  
  If KeyAscii = vbKeyReturn Then
  
    KeyAscii = 0
    
    If co = GRID.Col Then 'Está en la columna de observaciones
  
      If GRID.ColAllowEdit(co) = True Then
      
        '----------------------------------
        'Chequear si está en observaciones:
        '----------------------------------
        If GRID.Col = La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES) Then
          s = Trim(GRID.TextMatrix(GRID.Row, COL_CLIENTE))
          If s = "" Then
            MsgBox "Debe Indicar el Cliente a Procesar, Revise...", vbCritical, "Información"
          Else
            cc = La_Columna_GRID(GRID, "CEDULA")
            If cc > 0 Then
              s = Trim(GRID.TextMatrix(GRID.Row, cc))
              If s = "" Then
                MsgBox "Debe Indicar la Cédula de la Persona, Revise...", vbCritical, "Información"
                GRID.SetFocus
              Else
                If Not Se_Puede_GUARDAR_Movimiento() Then
                  MsgBox "Faltan Datos Requeridos para Guardar el Movimiento, Revise...", vbCritical, "Información"
                Else
                  C1 = Trim(GRID.TextMatrix(GRID.Row, COL_CLIENTE))
                  C2 = Trim(GRID.TextMatrix(GRID.Row, COL_SUBCLIENTE))
                  C3 = Modulo.La_Tabla_Actual_Personas(C1, C2)
                  If C3 <> "" Then
                    If Trim(GRID.TextMatrix(GRID.Row, 1)) = "" Then   'ID por guardar
                    
                      'Load fObservaciones
                      Modulo.vTemporal1 = ""
                      
                      Modulo.vTemporal2 = GRID.TextMatrix(GRID.Row, 3)
                      Modulo.vTemporal3 = GRID.TextMatrix(GRID.Row, 4)

                      Modulo.fModalResult = Modulo.fModalResultCANCEL
                      Load fObservaciones
                      
                      dTC = 0#
                      dTA = 0#
                      dTS = 0#
                    
                      Modulo.Resumen_Cuenta_Cliente GRID.TextMatrix(GRID.Row, 3), _
                                                    GRID.TextMatrix(GRID.Row, 4), _
                                                    dTC, dTA, dTS

                      fObservaciones.lTC.Caption = Format(dTC, "#,0.00")
                      fObservaciones.lTA.Caption = Format(dTA, "#,0.00")
                      fObservaciones.lTS.Caption = Format(dTS, "#,0.00")
                      
                      
                      
          
                      fObservaciones.lCliente.Caption = GRID.TextMatrix(GRID.Row, 3)
                      fObservaciones.lSubCliente.Caption = GRID.TextMatrix(GRID.Row, 4)
                      fObservaciones.EsActualizar = False
                      fObservaciones.Show vbModal
          
                      If Modulo.fModalResult = Modulo.fModalResultOK Then
                        If Modulo.vTemporal1 <> "" Then GRID.TextMatrix(GRID.Row, GRID.Col) = Trim(Modulo.vTemporal1)
                         
                        GuardarMovimientos GRID.Row
                            
                        Marcar_Personas_CEDULA C3, s
                        MsgBox "Registro de Movimiento Guardado...", vbInformation, "Información"
                      End If
                    End If
                  Else
                    MsgBox "No hay Tabla Creada para Personas del Cliente...", vbCritical, "Información"
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    End If
    KeyAscii = 0
  End If
        

End Sub

Private Sub MostrarObservaciones()
  Dim i As Integer
  Dim r As New ADODB.Recordset
  Dim sF As String, s As String, sID As String
  Dim lID As Long, iOBS As Integer
    
  iOBS = Modulo.La_Columna_GRID(GRID, COL_TITULO_OBSERVACIONES)
  
  If iOBS <= 0 Then Exit Sub
  
  
  sF = Format(Date, "yyyy/mm/dd")
  
  s = "select * from diario where fecha = '" & sF & "' order by id"
  
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  
  If Not r.EOF Then
    
    For i = 1 To GRID.Rows
    
      If Trim(GRID.TextMatrix(i, iOBS)) = "" Then 'Buscar la obs. y mostrarla
          
        sID = Trim(GRID.TextMatrix(i, 1))
        If sID <> "" Then
          r.MoveFirst
          r.Find "ID = " & sID & " "
          If Not r.EOF Then
            GRID.TextMatrix(i, iOBS) = Trim(r.Fields("observaciones").value)
          End If
        End If
      End If
    Next i
    
  End If
            
  r.Close
  Set r = Nothing
    
  
End Sub
