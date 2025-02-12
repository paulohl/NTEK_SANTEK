Attribute VB_Name = "Modulo"
Declare Function apiCopyFile Lib "kernel32" Alias "CopyFileA" _
(ByVal lpExistingFileName As String, _
ByVal lpNewFileName As String, _
ByVal bFailIfExists As Long) As Long

Option Explicit
           
Type Productos
   CodigoProducto As String
   Cantidad As Integer
End Type

Type TPerfiles
Clientes As Boolean
ClientesAgregar As Boolean
ClientesEditar  As Boolean
ClientesEliminar    As Boolean
ClientesReportes    As Boolean
SubClientes As Boolean
SubClientesAgregar  As Boolean
SubClientesEliminar As Boolean
SubClientesEditar   As Boolean
SubClientesReportes As Boolean
PorLotes    As Boolean
PorLotesAgregar As Boolean
PorLotesEliminar    As Boolean
PorLotesEditar  As Boolean
PorLotesReportes    As Boolean
TablasPersonas  As Boolean
TablasPersonasAgregar   As Boolean
TablasPersonasEliminar  As Boolean
TablasPersonasEditar    As Boolean
TablasPersonasReportes  As Boolean
Diario  As Boolean
FormatoDiseño   As Boolean
FormatoDiseñoAgregar    As Boolean
FormatoDiseñoEliminar   As Boolean
FormatodiseñoEditar As Boolean
FormatoDiseñoReportes   As Boolean
Inventario  As Boolean
InventarioAgregar   As Boolean
InventarioEliminar  As Boolean
InventarioEditar    As Boolean
InventarioReportes  As Boolean
Usuarios    As Boolean
UsuariosAgregar As Boolean
UsuariosEliminar    As Boolean
UsuariosEditar  As Boolean
UsuariosReportes    As Boolean
PerfilesAcceso As Boolean
PerfilesAgregar As Boolean
PerfilesEliminar As Boolean
PerfilesEditar As Boolean
PerfilesReportes As Boolean
Pagos   As Boolean
PagosAgregar    As Boolean
PagosEliminar   As Boolean
PagosEditar As Boolean
PagosReportes   As Boolean
Reportes    As Boolean
ComboProductos  As Boolean
OpcionesGenerales   As Boolean
CargaClientes   As Boolean
TransferirFotos As Boolean
configurarBotones   As Boolean
Herramientas As Boolean
End Type

Public gPerfil As TPerfiles

Dim CONEXION_SQL As String
Dim ODBC As String

Public IP_Servidor As String
Public DBConexionSQL As New ADODB.Connection
Public DBComandoSQL As New ADODB.Command

Public USUARIO_ACTUAL As String
Public NIVEL_ACTUAL As String
Public PERMISOS_ACTUAL As String

Public ESTACION As String

Public ClientePPAL As String

Public Const APPNAME = "SANTEK"

Public vTemporal1 As String
Public vTemporal2 As String
Public vTemporal3 As String

Public fModalResult As String
Public Const fModalResultOK = "OK"
Public Const fModalResultCANCEL = "CANCEL"

Public vIndiceFoto As Integer
Public vID As Long
Public vRUTAINICIAL As String
Public vRUTA_CARD5 As String
Public vMOVPAGO As Double

Public Const MARCA_IMPRESION_CARD5 = "I"
Public Const TIT_SISTEMA = "Sistema Administrativo de Carnetización ENTEK, C.A."

Public vLocalizador As String

Public vArrayCampos() As String
Public vMaxArrayCampos As Integer
Public vMontoAnterior As Double

Public Const EXTJPG = ".JPG"
Public NCol As Integer ' para incrementar el numero de columnas y poder agregar la columna "Check" en la forma fmov

Public Const MAX_BTNS = 10
Type Observaciones
   Codigo As Integer
   Observacion As String
End Type
Public Function CopyFile(ArchivoOrigen As String, ArchivoDestino As String, Pisar As Boolean) As Long
   CopyFile = apiCopyFile(ArchivoOrigen, ArchivoDestino, Pisar)
End Function


Public Sub sCargarPerfilUsuario(argCodigoPerfil As Integer)
   Dim lReg As New ADODB.Recordset
   Dim lCn As New ADODB.Connection
   lCn.Open Modulo.DBConexionSQL
   Set lReg = lCn.Execute("Select * From Perfiles Where codigo=" & argCodigoPerfil & " And Activo = 1")
   If lReg.EOF = False Then
     With gPerfil
      .CargaClientes = lReg!CargaClientes
      .Clientes = lReg!Clientes
      .ClientesAgregar = lReg!ClientesAgregar
      .ClientesEditar = lReg!ClientesEditar
      .ClientesEliminar = lReg!ClientesEliminar
      .ClientesReportes = lReg!ClientesReportes
      .ComboProductos = lReg!ComboProductos
      .configurarBotones = lReg!configurarBotones
      .Diario = lReg!Diario
      .FormatoDiseño = lReg!FormatoDiseño
      .FormatoDiseñoAgregar = lReg!FormatoDiseñoAgregar
      .FormatodiseñoEditar = lReg!FormatodiseñoEditar
      .FormatoDiseñoEliminar = lReg!FormatoDiseñoEliminar
      .FormatoDiseñoReportes = lReg!FormatoDiseñoReportes
      .Inventario = lReg!Inventario
      .InventarioAgregar = lReg!InventarioAgregar
      .InventarioEditar = lReg!InventarioEditar
      .InventarioEliminar = lReg!InventarioEliminar
      .InventarioReportes = lReg!InventarioReportes
      .OpcionesGenerales = lReg!OpcionesGenerales
      .Pagos = lReg!Pagos
      .PagosAgregar = lReg!PagosAgregar
      .PagosEditar = lReg!PagosEditar
      .PagosEliminar = lReg!PagosEliminar
      .PagosReportes = lReg!PagosReportes
      .PorLotes = lReg!PorLotes
      .PorLotesAgregar = lReg!PorLotesAgregar
      .PorLotesEditar = lReg!PorLotesEditar
      .PorLotesEliminar = lReg!PorLotesEliminar
      .PorLotesReportes = lReg!PorLotesReportes
      .Reportes = lReg!Reportes
      .SubClientes = lReg!SubClientes
      .SubClientesAgregar = lReg!SubClientesAgregar
      .SubClientesEditar = lReg!SubClientesEditar
      .SubClientesEliminar = lReg!SubClientesEliminar
      .SubClientesReportes = lReg!SubClientesReportes
      .TablasPersonas = lReg!TablasPersonas
      .TablasPersonasAgregar = lReg!TablasPersonasAgregar
      .TablasPersonasEditar = lReg!TablasPersonasEditar
      .TablasPersonasEliminar = lReg!TablasPersonasEliminar
      .TablasPersonasReportes = lReg!TablasPersonasReportes
      .TransferirFotos = lReg!TransferirFotos
      .Usuarios = lReg!Usuarios
      .UsuariosAgregar = lReg!UsuariosAgregar
      .UsuariosEditar = lReg!UsuariosEditar
      .UsuariosEliminar = lReg!UsuariosEliminar
      .UsuariosReportes = lReg!UsuariosReportes
      .Herramientas = lReg!Herramientas
      .PerfilesAcceso = lReg!PerfilesAcceso
      .PerfilesAgregar = lReg!PerfilesAgregar
      .PerfilesEditar = lReg!PerfilesEditar
      .PerfilesEliminar = lReg!PerfilesEliminar
      .PerfilesReportes = lReg!PerfilesReportes
     End With
   Else
      MsgBox "No existe perfil de usuario", vbCritical
      End
   End If
   
   
End Sub




Sub Main()

  IP_Servidor = "" '"192.168.10.150"
  
  ODBC = "SANTEK"
  
  'CONEXION_SQL = "Provider=MSDASQL.1;" & _
                 "Persist Security Info=False;" & _
                 "Data Source=" & ODBC & ";" & _
                 "Initial Catalog=santek"
                 
  'CONEXION_SQL = "driver={SQL Server};server=" & IP_Servidor & ";" & _
                 "database=santek;Username=sa;PWD=sql123;"
                 
  CONEXION_SQL = "Provider=SQLOLEDB.1;" & _
                 "Password=sa;" & _
                 "Persist Security Info=True;" & _
                 "User ID=sa;" & _
                 "Initial Catalog=santek;" & _
                 "Data Source=" & IP_Servidor
                 
                 
  DBConexionSQL.ConnectionString = CONEXION_SQL
  
  'If Not Abrir_BD() Then End
  
  USUARIO_ACTUAL = ""
  NIVEL_ACTUAL = ""
  PERMISOS_ACTUAL = ""
  If App.PrevInstance = True Then
     MsgBox "ya se está ejecutando una instancia del programa en este equipo", vbExclamation
     End
  End If
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
    
    ClientePPAL = ""
  
    ESTACION = GetSetting(APPNAME, "Opciones", "Estacion", "")
  
    Load fSistema
    fSistema.sVerificarAccesos
    fSistema.Caption = TIT_SISTEMA & " - Estación:" & ESTACION & " Conectado " & Modulo.IP_Servidor & " Usuario: " & Modulo.USUARIO_ACTUAL
  
    fSistema.Show
        
    fSistema.Enabled = True
    
    AgregarLogs "Inicia Sesión"
    
    
  Else
    End
  End If
    

End Sub

Function Abrir_BD() As Boolean
  On Error Resume Next
  Dim lNombreNuevo As String
  'Load fMensaje
  'fMensaje.Caption = "Conectando, Espere..."
  'fMensaje.Show
  
  If DBConexionSQL.State = adStateClosed Then
    CONEXION_SQL = "Provider=SQLOLEDB.1;" & _
                 "Password=;" & _
                 "Persist Security Info=True;" & _
                 "User ID=sa;PWD=ntekca;" & _
                 "Initial Catalog=santek;" & _
                 "Data Source=" & Modulo.IP_Servidor
                 
    DBConexionSQL.ConnectionString = CONEXION_SQL
  End If
  
  If DBConexionSQL.State = adStateClosed Then DBConexionSQL.Open
  
  'Unload fMensaje
  
  If Err.Number <> 0 Then
    MsgBox "Imposible Conectar con Servidor..." & vbCrLf & Err.Description, vbCritical, "Información"
    Abrir_BD = False
    'lNombreNuevo = InputBox("Verifique el nombre del servidor", Modulo.IP_Servidor)
    'If Modulo.IP_Servidor <> lNombreNuevo Then
    'End If
  Else
    Abrir_BD = True
  End If

End Function

Function Zeros(iValor As Long, iCantidad As Integer) As String
  Dim s As String
  s = Trim(CStr(iValor))
  If Len(s) < iCantidad Then
    s = String(iCantidad - Len(s), "0") & s
  End If
  Zeros = s
End Function

Function BlancosIZQ(sValor As String, iCantidad As Integer) As String
  Dim s As String
  s = Trim(sValor)
  If Len(s) < iCantidad Then
    s = String(iCantidad - Len(s), " ") & s
  End If
  BlancosIZQ = s
End Function

Function BlancosDER(sValor As String, iCantidad As Integer) As String
  Dim s As String
  s = Trim(sValor)
  If Len(s) < iCantidad Then
    s = s & String(iCantidad - Len(s), " ")
  End If
  BlancosDER = s
End Function

Function DepurarStr(sValor As String, CarDep As String) As String
  Dim i As Integer
  Dim s As String
  s = ""
  For i = 1 To Len(sValor)
    If Mid(sValor, i, 1) <> CarDep Then s = s & Mid(sValor, i, 1)
  Next i
  DepurarStr = s
End Function


' \\ -- Extraer solo la ruta sin el archivo mediante la función Left e InstrRev
Function GetPath(sPath As String, Caracter As String) As String
    If sPath <> "" And Caracter <> "" Then
       GetPath = Left(sPath, InStrRev(sPath, Caracter))
    End If
End Function

' \\ -- Extraer solo la extension, "."
' \\ -- Extraer solo el archivo, "\"
Function ExtraerFilePath(Path As String, Caracter As String) As String
    Dim ret As String
    If Caracter = "." And InStr(Path, Caracter) = 0 Then Exit Function
    ret = Right(Path, Len(Path) - InStrRev(Path, Caracter))
    ' -- Retorna el valor
    ExtraerFilePath = ret
End Function

Function ExtraerArchivo(Path As String) As String
  Dim i As Integer
  Dim s As String
  Dim e As Boolean
  s = ""
  e = False
  i = Len(Path)
  Do While i >= 1 And Not e
    If Mid(Path, i, 1) = "/" Or Mid(Path, i, 1) = "\" Then
      e = True
    Else
      s = Mid(Path, i, 1) & s
      i = i - 1
    End If
  Loop
  ExtraerArchivo = s
End Function

Function QuitarExtension(Path_Archivo As String) As String
  Dim s As String
  Dim i As Integer, j As Integer
  Dim HP As Boolean
  
  j = -1
  For i = Len(Path_Archivo) To 1 Step -1
    If Mid(Path_Archivo, i, 1) = "." Then
      j = i
      HP = True
    End If
  Next i
  
  QuitarExtension = ""
  
  If HP Then
    s = Mid(Path_Archivo, 1, j - 1)
    QuitarExtension = s
  End If

End Function

Sub EnMayusculas(ByRef oTexto As TextBox)
  Dim i As Integer
  oTexto.Text = UCase(oTexto.Text)
  i = Len(oTexto.Text)
  oTexto.SelStart = i
End Sub

Function Nombre_Directorio_Valido(sNom As String) As Boolean
  On Error Resume Next
  Dim s As String
  s = Modulo.GetPath(App.Path & "\", "\") & sNom
  MkDir s
  If Err.Number <> 0 Then
    MsgBox "El Texto [" & sNom & "] No es Válido para crear una Carpeta," & vbCrLf & _
           "Revise que no contenga los caracteres \ / : * ? '' < > |", vbCritical, "Información"
    Nombre_Directorio_Valido = False
  Else
    RmDir s
    Nombre_Directorio_Valido = True
  End If
End Function

Function Nombre_Directorio_Repetido(sNom As String, bClientePpal As Boolean) As Boolean
  Dim r As New ADODB.Recordset
  Dim s As String, s1 As String
  If bClientePpal = True Then
    s = "SELECT * FROM Clientes WHERE Nombre = '" & sNom & "'"
  Else
    s = "SELECT * FROM SubClientes WHERE Nombre = '" & sNom & "'"
  End If
  r.Open s, Modulo.DBConexionSQL, adOpenKeyset, adLockReadOnly
  If Not r.EOF Then
    If bClientePpal Then
      s1 = "Existe un Registro con Nombre [" & Trim(r.Fields("nombre").value) & "]" & vbCrLf & _
           "Código Cliente = " & Zeros(r.Fields("codigo").value, 6) & " - RIF: " & Trim(r.Fields("rif").value)
    Else
      s1 = "Existe un Registro con Nombre [" & Trim(r.Fields("nombre").value) & "]" & vbCrLf & _
           "Código Cliente = " & Zeros(r.Fields("cliente").value, 6) & " - RIF: " & Trim(r.Fields("rif").value)
    End If
           
    MsgBox s1, vbCritical, "Información"
    Nombre_Directorio_Repetido = True
  Else
    Nombre_Directorio_Repetido = False
  End If
  r.Close
  Set r = Nothing
End Function

Function Ejecutar_DOS(Comando As String) As String
  Dim oShell As WshShell
  Dim oExec As WshExec
  Dim ret As String

  Set oShell = New WshShell
  DoEvents

  ' ejecutar el comando
  Set oExec = oShell.Exec("%comspec% /c " & Comando)
  ret = oExec.StdOut.ReadAll()

  ' retornar la salida y devolverla a la función
  Ejecutar_DOS = ret ' Replace(ret, Chr(10), vbNewLine)

  DoEvents
  'Me.SetFocus
End Function

Function Buscar_Combo(ByRef xCombo As ComboBox, sValor As String) As Integer
  Dim s As String
  Dim i As Integer
  Dim e As Boolean
  i = 0
  e = False
  Do While (i < xCombo.ListCount) And Not e
    If xCombo.List(i) = sValor Then e = True Else i = i + 1
  Loop
  If e Then Buscar_Combo = i Else Buscar_Combo = 0
End Function

Function Buscar_ComboLen(ByRef xCombo As ComboBox, sValor As String, iCantidad As Integer) As Integer
  Dim s As String
  Dim i As Integer
  Dim e As Boolean
  i = 0
  e = False
  Do While (i < xCombo.ListCount) And Not e
    s = Mid(xCombo.List(i), 1, iCantidad)
    If sValor = s Then e = True Else i = i + 1
  Loop
  If e Then Buscar_ComboLen = i Else Buscar_ComboLen = 0
End Function

Function Str_Replace(sCadena As String, sCaracterOLD As String, sCaracterNEW As String) As String
  Dim sr As String
  Dim i As Integer
  sr = ""
  For i = 1 To Len(sCadena)
    If Mid(sCadena, i, 1) = sCaracterOLD Then
      sr = sr & sCaracterNEW
    Else
      sr = sr & Mid(sCadena, i, 1)
    End If
  Next i
  Str_Replace = sr
End Function

Sub FlexXY(FG As MSFlexGrid, Fila As Integer, Columna As Integer, sTexto As String, Alineacion As Long)
  If Fila >= 0 And Fila < FG.Rows And Columna >= 0 And Columna < FG.Cols Then
    FG.Row = Fila
    FG.Col = Columna
    FG.CellAlignment = Alineacion
    FG.TextMatrix(Fila, Columna) = sTexto
  End If
End Sub

Function Existe_Tabla(sTabla As String) As Boolean
  Dim s As String
  Dim r As New ADODB.Recordset
  Dim ET As Boolean
  ET = False
  s = "select * from sysobjects where name='" & sTabla & "' and type='U'"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If Not r.EOF Then ET = True
  r.Close
  Set r = Nothing
  Existe_Tabla = ET
End Function

Public Sub ExecSQL(sSQL As String)
  On Error Resume Next
  Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
  Modulo.DBComandoSQL.CommandText = sSQL
  Modulo.DBComandoSQL.Execute
  If Err.Number <> 0 Then
    MsgBox "Error al Ejecutar Comando..." & vbCrLf & _
           "[" & Err.Number & "] " & Err.Description, vbCritical, "Información"
  End If
End Sub

Public Function Total_Registros(sTabla As String) As Long
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim TR As Long
  TR = 0
  s = "select count(*) from [" & sTabla & "]"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If Not r.EOF Then
    If Not IsNull(r.Fields(0).value) Then TR = r.Fields(0).value
  End If
  r.Close
  Set r = Nothing
  Total_Registros = TR
End Function

Public Function Nro_Columna_FlexGrid(ByRef FlexGrid As MSFlexGrid, sColumna As String) As Integer
  Dim cc As Integer
  Dim i As Integer
  i = 0
  cc = -1
  Do While i < FlexGrid.Cols And cc = -1
    If UCase(FlexGrid.TextMatrix(0, i)) = sColumna Then cc = i
    i = i + 1
  Loop
  Nro_Columna_FlexGrid = cc
End Function

Public Function Nro_Columna_DataGrid(ByRef xDataGrid As DataGrid, sColumna As String) As Integer
  Dim cc As Integer
  Dim i As Integer
  i = 0
  cc = -1
  Do While i < xDataGrid.Columns.Count And cc = -1
    If UCase(xDataGrid.Columns(i).Caption) = UCase(sColumna) Then cc = i
    i = i + 1
  Loop
  Nro_Columna_DataGrid = cc
End Function

Public Function EXISTE_CAMPO(ByRef oRecordset As ADODB.Recordset, sNombre As String) As Boolean
  Dim i As Integer
  Dim e As Boolean
  i = 0
  e = False
  Do While i < oRecordset.Fields.Count And Not e
    If UCase(oRecordset.Fields(i).Name) = UCase(sNombre) Then
      e = True
    Else
      i = i + 1
    End If
  Loop
  EXISTE_CAMPO = e
End Function

Public Function EXISTE_CAMPO_EN_FLEXGRID(ByRef FlexGrid As MSFlexGrid, sNombre As String) As Boolean
  Dim i As Integer
  Dim e As Boolean
  i = 0
  e = False
  Do While i < FlexGrid.Cols And Not e
    If UCase(FlexGrid.TextMatrix(0, i)) = UCase(sNombre) Then
      e = True
    Else
      i = i + 1
    End If
  Loop
  EXISTE_CAMPO_EN_FLEXGRID = e
End Function

Function Correo_E(sCliente As String, sSubCliente As String) As String
  Dim CE As String
  Dim s As String
  Dim r As New ADODB.Recordset
  CE = ""
  If sCliente <> "" Then
    If sSubCliente = "" Then
      s = "select email from clientes where codigo = " & sCliente & " "
    Else
      s = "select email from subclientes where cliente = " & sCliente & " and id = " & sSubCliente
    End If
    r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    If Not r.EOF Then
      If Not IsNull(r.Fields("email").value) Then
        CE = Trim(r.Fields("email").value)
      End If
    End If
  End If
  Correo_E = CE
End Function

Public Function Fila_En_Blanco_En_FlexGrid(ByRef FlexGrid As MSFlexGrid, iFila As Integer) As Boolean
  Dim i As Integer
  Dim bEnBlanco As Boolean
  i = 0
  bEnBlanco = True
  Do While i < FlexGrid.Cols And bEnBlanco
    If Trim(FlexGrid.TextMatrix(iFila, i)) <> "" Then bEnBlanco = False Else i = i + 1
  Loop
  Fila_En_Blanco_En_FlexGrid = bEnBlanco
End Function

Public Sub Actualizar_Tiene_Foto(sRutaFOTO As String, sTablaCS As String)
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim sFoto As String
  Dim sTieneFoto As String
  Dim bTieneChar9 As Boolean
  
  s = "select * from [" & sTablaCS & "]"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  Do While Not r.EOF
    sFoto = Trim(r.Fields("foto").value)
    bTieneChar9 = StrTieneChar9(sFoto)
    If bTieneChar9 Then sFoto = StrDepurarChar9(Trim(r.Fields("foto").value))
    s = sRutaFOTO & "\" & sFoto & ".JPG"
    If Dir(s) <> "" Then sTieneFoto = "S" Else sTieneFoto = "N"
    s = "update [" & sTablaCS & "] set tiene_foto = '" & sTieneFoto & "' where id = " & CStr(r.Fields("id").value) & " "
    Modulo.ExecSQL s
    
    If bTieneChar9 Then
      s = "update [" & sTablaCS & "] set Foto = '" & sFoto & "' where id = " & CStr(r.Fields("id").value)
      Modulo.ExecSQL s
    End If
    
    
    r.MoveNext
  Loop
  r.Close
  Set r = Nothing
End Sub

Public Function DepurarValorCEDULA(sValor As String) As String
  Dim s As String, c As String
  Dim i As Integer
  s = ""
  For i = 1 To Len(sValor)
    c = Mid(sValor, i, 1)
    If Asc(c) >= Asc("0") And Asc(c) <= Asc("9") Then
      s = s & c
    Else
      If Asc(c) >= Asc("A") And Asc(c) <= Asc("Z") Then
        s = s & c
      Else
        If Asc(c) >= Asc("a") And Asc(c) <= Asc("z") Then
          s = s & c
        Else
          If Asc(c) = Asc(".") Or Asc(c) = Asc(",") Or Asc(c) = Asc("-") Then
            s = s & c
          End If
        End If
      End If
    End If
  Next i
  DepurarValorCEDULA = s
End Function

Public Function Existe_Valor_En_Columna(FlexGrid As MSFlexGrid, iColumna As Integer, sValor As String) As Boolean
  Dim e As Boolean
  Dim i As Integer
  i = 1
  e = False
  Do While i < FlexGrid.Rows - 1 And Not e
    If Trim(FlexGrid.TextMatrix(i, iColumna)) = sValor Then
      e = True
    Else
      i = i + 1
    End If
  Loop
  Existe_Valor_En_Columna = e
End Function

Public Function La_Fila_Existe_Valor_En_Columna(FlexGrid As MSFlexGrid, iColumna As Integer, sValor As String) As Integer
  Dim e As Boolean
  Dim i As Integer
  i = 1
  e = False
  Do While i < FlexGrid.Rows - 1 And Not e
    If Trim(FlexGrid.TextMatrix(i, iColumna)) = sValor Then
      e = True
    Else
      i = i + 1
    End If
  Loop
  If e Then La_Fila_Existe_Valor_En_Columna = i Else La_Fila_Existe_Valor_En_Columna = -1
End Function


Public Sub Auditar_Fotos(sTablaC5 As String, Optional argCedula As String)    'Utilizar "" en parametro para todas las tablas!
  Dim r As New ADODB.Recordset
  
  Dim rc As New ADODB.Recordset
  Dim rsc As New ADODB.Recordset
  
  Dim rC5 As New ADODB.Recordset
    
  Dim s As String
  Dim Cliente As Long
  Dim SubCliente As Long
  
  Dim sNomCliente As String, sNomSubCliente As String
  
  Dim sRuta As String, i As Integer
  
  If sTablaC5 = "" Then
    s = "select * from personas"
  Else
    s = "select * from personas where tabla = '" & sTablaC5 & "'"
  End If
  
  Load fMensaje
  fMensaje.Label1.Caption = "Auditando FOTOS, Espere..."
  'fMensaje.Show
  DoEvents
    
  i = 1
  
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  
  Do While Not r.EOF
  
    '''''''''fMensaje.Label1.Caption = "Auditando FOTOS (" & CStr(i) & "/" & CStr(r.RecordCount) & "), Espere..."
    'fMensaje.Show
    DoEvents
  
    Cliente = r.Fields("cliente").value
    SubCliente = r.Fields("subcliente").value
    
    sNomCliente = ""
    sNomSubCliente = ""
    
    '-Buscar nombre del cliente:
    s = "select nombre from clientes where codigo = " & CStr(Cliente) & " "
    rc.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    If Not rc.EOF Then
      sNomCliente = Trim(rc.Fields("nombre").value)
    End If
    rc.Close
    Set rc = Nothing
    
    '-Buscar nombre del Sub-cliente:
    If SubCliente > 0 Then
      s = "select nombre from subclientes where cliente = " & CStr(Cliente) & " and id = " & CStr(SubCliente) & " "
      rc.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
      If Not rc.EOF Then
        sNomSubCliente = Trim(rc.Fields("nombre").value)
      End If
      rc.Close
      Set rc = Nothing
    Else
      sNomSubCliente = ""
    End If
    
    '-Armar la Ruta:
    Dim sOri As String
    Dim sDes As String
      
    sOri = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
    sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
    
    If sOri <> "" And sDes <> "" Then
      
      sRuta = sDes & "\" & sNomCliente & "\" & IIf(sNomSubCliente <> "", sNomSubCliente & "\", "") & "FOTOS"
      If argCedula = "" Then
         s = "select * from [" & Trim(r.Fields("Tabla").value) & "]"
      Else
         s = "select * from [" & Trim(r.Fields("Tabla").value) & "] where Cedula='" & argCedula & "'"
      End If
      rC5.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
      Do While Not rC5.EOF
        If Modulo.EXISTE_CAMPO(rC5, "TIENE_FOTO") = True Then
          If StrDepurarChar9(IIf(IsNull(Trim(rC5.Fields("foto").value)), "", Trim(rC5.Fields("foto").value))) <> "" Then
            s = sRuta & "\" & StrDepurarChar9(Trim(rC5.Fields("foto").value)) & ".JPG"
            If Dir(s) <> "" Then 'Existe la FOTO en su carpeta...
              'If UCase(Trim(rC5.Fields("tiene_foto").Value)) = "N" Then
                 If UCase(Trim(rC5.Fields("contador").value)) = 0 Then
                     s = "update [" & Trim(r.Fields("Tabla").value) & "] set " & _
                     "tiene_foto = 'S', MARCA='I', foto = '" & StrDepurarChar9(Trim(rC5.Fields("foto").value)) & "' where id = " & CStr(rC5.Fields("id").value) & " "
                 Else
                     s = "update [" & Trim(r.Fields("Tabla").value) & "] set " & _
                     "tiene_foto = 'S', foto = '" & StrDepurarChar9(Trim(rC5.Fields("foto").value)) & "' where id = " & CStr(rC5.Fields("id").value) & " "
                 End If
                Modulo.ExecSQL s
              'End If
            Else 'No Existe la FOTO en su carpeta...
              'If UCase(Trim(rC5.Fields("tiene_foto").Value)) = "S" Then
                s = "update [" & Trim(r.Fields("Tabla").value) & "] set " & _
                    "tiene_foto = 'N',MARCA='' where id = " & CStr(rC5.Fields("id").value) & " "
                Modulo.ExecSQL s
              'End If
            End If
          End If
        End If
        rC5.MoveNext
      Loop
      rC5.Close
      Set rC5 = Nothing
    End If
    r.MoveNext
    i = i + 1
  Loop
  r.Close
  Set r = Nothing
  Unload fMensaje
End Sub

Public Function StrSoloDigitos(sCadena As String) As String
  Dim s As String, c As String
  Dim i As Integer
  s = ""
  For i = 1 To Len(sCadena)
    c = Mid(sCadena, i, 1)
    If Asc(c) >= Asc("0") And Asc(c) <= Asc("9") Then s = s & c
  Next i
  StrSoloDigitos = s
End Function

Public Function StrDepurarChar9(sCadena As String) As String
  Dim s As String, c As String
  Dim i As Integer
  s = ""
  For i = 1 To Len(sCadena)
    c = Mid(sCadena, i, 1)
    If Asc(c) <> 9 Then s = s & c
  Next i
  StrDepurarChar9 = s
End Function

Public Function StrTieneChar9(sCadena As String) As String
  Dim s As String, c As String
  Dim i As Integer, e As Boolean
  i = 1
  e = False
  Do While i < Len(sCadena) And Not e
    c = Mid(sCadena, i, 1)
    If Asc(c) = 9 Then e = True Else i = i + 1
  Loop
  StrTieneChar9 = e
End Function


Public Sub CargarOpcionesCorreo(ByRef s1 As String, ByRef s2 As String, ByRef s3 As String, s4 As String, s5 As String, s6 As String)
  Dim s As String
  Dim r As New ADODB.Recordset
  s1 = ""
  s2 = ""
  s3 = ""
  s4 = ""
  s5 = ""
  s6 = ""
  's = "select * from opciones"
  'r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  'If Not r.EOF Then
  '  If Not IsNull(r.Fields("titulomensajecorreo").Value) Then s1 = Trim(r.Fields("titulomensajecorreo").Value)
  '  If Not IsNull(r.Fields("cuerpomensajecorreo").Value) Then s2 = Trim(r.Fields("cuerpomensajecorreo").Value)
    
  '  If Not IsNull(r.Fields("codigoproductopvc").Value) Then s3 = Trim(r.Fields("codigoproductopvc").Value)
  'End If
  'r.Close
  'Set r = Nothing
  
  
  
  
  s1 = GetSetting(APPNAME, "Opciones", "TituloMensajeCorreo", "")
  s2 = GetSetting(APPNAME, "Opciones", "CuerpoMensajeCorreo", "")
  s3 = GetSetting(APPNAME, "Opciones", "CodigoCarnet", "")
  s4 = GetSetting(APPNAME, "Opciones", "UsuarioEmail", "")
  s5 = GetSetting(APPNAME, "Opciones", "ContraseñaEmail", "")
  s6 = GetSetting(APPNAME, "Opciones", "ConCopiaEmail", "")
  
End Sub


Public Function MandaMail(para As String, _
                          cc As String, _
                          BCC As String, _
                          Titulo As String, _
                          Cuerpo As String, _
                          ArchivoAdj As ListBox) As Boolean

  Dim correo As Outlook.Application
  Dim Item As Outlook.MailItem
  Dim sClase As String
  Dim i As Integer
  Load fMensaje
  'fMensaje.Label1.Caption = "Generando Correo MS-OUTLOOK, Espere..."
  'fMensaje.Show
  'DoEvents
  
  sClase = "Outlook.Application" '"Microsoft.Office.Interop.Outlook.Application"  '"Outlook.Application"

  Set correo = New Outlook.Application  'GetObject("", sClase)
  
  'If correo Is Nothing Then
  '  Set correo = CreateObject("", sClase)
  'End If

  Set Item = correo.CreateItem(olMailItem)

  On Error GoTo ErrorMAPI

  DoEvents

  If Trim(para) <> "" Then Item.To = para
  If Trim(BCC) <> "" Then Item.BCC = BCC
  If Trim(cc) <> "" Then Item.cc = cc
  
  Item.Subject = Titulo
  Item.Body = Cuerpo
  
  For i = 0 To ArchivoAdj.ListCount - 1
     Item.Attachments.Add ArchivoAdj.List(i)
  Next i
  'Unload fMensaje
  'item.Display True
  
  Item.Send
  
  'correo.Session.SendAndReceive True
  
    
  Set Item = Nothing
  Set correo = Nothing
  MandaMail = True
  MsgBox "Correo electrónico enviado con éxito", vbInformation
Salir:
  Exit Function
ErrorMAPI:
  MandaMail = False
  Screen.MousePointer = vbDefault
  MsgBox "Hubo un error al enviar el correo electrónico, el error fue debido a: " + Err.Description
  Resume Salir
End Function


Public Sub Ajustar_Columna_DataGrid( _
                            un_DataGrid As DataGrid, _
                            Ado As Adodc, _
                            Optional AccForHeaders As Boolean)

    Dim TempCol() As Integer
    Dim Nregistros As Integer, NCampos As Integer
    Dim Fila As Long, Col As Long, width As Single
    Dim maxWidth As Single, celdaText As String
    Dim saveFont As StdFont, oldScaleMode As Integer
    
    'Variables para la cantidad de registros y columnas
    Nregistros = Ado.Recordset.RecordCount
    NCampos = Ado.Recordset.Fields.Count
    
    'Array para almacenar el ancho de cada columna
    ReDim TempCol(NCampos)
    
    
    ' Si el número de registros es igual a 0 salimos
    If Nregistros = 0 Then Exit Sub
    
    ' Guardamos la fuente del DataGrid para luego reestablecerla
    Set saveFont = un_DataGrid.Parent.Font
    Set un_DataGrid.Parent.Font = un_DataGrid.Font
    
    ' Ajustar el ScaleMode en vbTwips para el formulario
    oldScaleMode = un_DataGrid.Parent.ScaleMode
    un_DataGrid.Parent.ScaleMode = vbTwips
    
    'Movemos al Primer registro
    If Ado.Recordset.RecordCount > 0 Then
      Ado.Recordset.MoveFirst
    End If
    maxWidth = 0
    
    'Recorremos las columnas
    For Col = 0 To NCampos - 1
        Ado.Recordset.MoveFirst
        

        If AccForHeaders Then
            'Almacenamos el Ancho del texto de la columna
            maxWidth = un_DataGrid.Parent.TextWidth _
                      (un_DataGrid.Columns(Col).Text) + 200
        End If
        
        
        Ado.Recordset.MoveFirst
        'Recorremos los registros de esta columna
        For Fila = 0 To Nregistros - 1
            
            
            If NCampos = 1 Then
            Else
                celdaText = un_DataGrid.Columns(Col).Text
            End If
            
            'Almacena el Ancho del texto de la celda del Datagrid
            width = un_DataGrid.Parent.TextWidth(celdaText) + 200
            
            'Si el ancho de la celda es mayor se actualiza la variable maxWidth _
            y se establece el ancho de la columna
            If width > maxWidth Then
               maxWidth = width
               un_DataGrid.Columns(Col).width = maxWidth
            End If
            
            ' Movemos el Ado al Siguiente registro
            Ado.Recordset.MoveNext
        Next Fila
        'Almacenamos el ancho de la columna
        TempCol(Col) = maxWidth
        
      'Ado.Recordset.MoveNext
    Next Col
    
    'Recorremos cada columna y le asignamos el ancho
    For Col = 0 To NCampos - 1
        un_DataGrid.Columns(Col).width = TempCol(Col)
    Next
    
    'restablecemos la fuente del DataGrid y el scaleMode
    Set un_DataGrid.Parent.Font = saveFont
    un_DataGrid.Parent.ScaleMode = oldScaleMode
    
    Ado.Recordset.MoveFirst
    
    Erase TempCol
End Sub

Public Sub Ajustar_Columna_DataGrid2( _
                            un_DataGrid As DataGrid, _
                            Ado As Adodc, _
                            Optional AccForHeaders As Boolean)

    Dim TempCol() As Integer
    Dim TempMaxWidth() As Single
    Dim Nregistros As Integer, NCampos As Integer
    Dim Fila As Long, Col As Long, width As Single
    Dim maxWidth As Single, celdaText As String
    Dim saveFont As StdFont, oldScaleMode As Integer
    Dim r As New ADODB.Recordset
    
    'Variables para la cantidad de registros y columnas
    Nregistros = Ado.Recordset.RecordCount
    NCampos = Ado.Recordset.Fields.Count
    
    'Array para almacenar el ancho de cada columna
    ReDim TempCol(NCampos)
    ReDim TempMaxWidth(NCampos)
    
    ' Si el número de registros es igual a 0 salimos
    If Nregistros = 0 Then Exit Sub
    
    ' Guardamos la fuente del DataGrid para luego reestablecerla
    Set saveFont = un_DataGrid.Parent.Font
    Set un_DataGrid.Parent.Font = un_DataGrid.Font
    
    ' Ajustar el ScaleMode en vbTwips para el formulario
    oldScaleMode = un_DataGrid.Parent.ScaleMode
    un_DataGrid.Parent.ScaleMode = vbTwips
    
    'Movemos al Primer registro
    If Ado.Recordset.RecordCount > 0 Then
      Ado.Recordset.MoveFirst
    End If
    maxWidth = 0
    
    'recorrer las filas:
    For Fila = 0 To Nregistros - 1
    
        For Col = 0 To NCampos - 1
            
          celdaText = Trim(un_DataGrid.Columns(Col).Text)
            
          'Almacena el Ancho del texto de la celda del Datagrid
          width = un_DataGrid.Parent.TextWidth(celdaText) + 150
                
          'Si el ancho de la celda es mayor se actualiza la variable maxWidth _
           y se establece el ancho de la columna
          If width > TempMaxWidth(Col) Then
            TempMaxWidth(Col) = width
            un_DataGrid.Columns(Col).width = width
          End If
            
        Next Col
        
      Ado.Recordset.MoveNext
    Next
      
    'Recorremos cada columna y le asignamos el ancho
    For Col = 0 To NCampos - 1
        un_DataGrid.Columns(Col).width = TempMaxWidth(Col)
    Next
    
    'restablecemos la fuente del DataGrid y el scaleMode
    Set un_DataGrid.Parent.Font = saveFont
    un_DataGrid.Parent.ScaleMode = oldScaleMode
    
    Ado.Recordset.MoveFirst
    
    Erase TempCol
End Sub

Public Sub Ajustar_Columna_GRID(xGrid As ubGrid, _
                                 iColDesde As Integer, _
                                 iColHasta As Integer)
    Dim TempCol() As Integer
    Dim TempMaxWidth() As Single
    Dim Nregistros As Integer, NCampos As Integer
    Dim Fila As Long, Col As Long, width As Single
    Dim maxWidth As Single, celdaText As String
    Dim saveFont As StdFont, oldScaleMode As Integer
    Dim r As New ADODB.Recordset
    
    'Variables para la cantidad de registros y columnas
    Nregistros = xGrid.Rows  'Ado.Recordset.RecordCount
    NCampos = xGrid.Cols  'Ado.Recordset.Fields.Count
    
    'Array para almacenar el ancho de cada columna
    ReDim TempCol(NCampos)
    ReDim TempMaxWidth(NCampos)
    
    ' Si el número de registros es igual a 0 salimos
    If Nregistros = 0 Then Exit Sub
    
    ' Guardamos la fuente del DataGrid para luego reestablecerla
    Set saveFont = xGrid.Parent.Font
    Set xGrid.Parent.Font = xGrid.Font
    
    ' Ajustar el ScaleMode en vbTwips para el formulario
    oldScaleMode = xGrid.Parent.ScaleMode
    xGrid.Parent.ScaleMode = vbTwips
    
    maxWidth = 0
    
    'recorrer las filas:
    For Fila = 0 To Nregistros
    
        For Col = iColDesde To iColHasta
            
          celdaText = Trim(xGrid.TextMatrix(Fila, Col))
            
          'Almacena el Ancho del texto de la celda del Datagrid
          width = xGrid.Parent.TextWidth(celdaText) + 150
                
          'Si el ancho de la celda es mayor se actualiza la variable maxWidth _
           y se establece el ancho de la columna
          If width > TempMaxWidth(Col) Then
            TempMaxWidth(Col) = width
            xGrid.ColWidth(Col) = width
          End If
            
        Next Col
    Next
      
    'Recorremos cada columna y le asignamos el ancho
    'For Col = 0 To NCampos - 1
    '    un_DataGrid.Columns(Col).width = TempMaxWidth(Col)
    'Next
    
    'restablecemos la fuente del DataGrid y el scaleMode
    Set xGrid.Parent.Font = saveFont
    xGrid.Parent.ScaleMode = oldScaleMode
    
    Erase TempCol
End Sub


Public Sub Ajustar_Columna_FLEXGRID(xGrid As MSFlexGrid, _
                                    iColDesde As Integer, _
                                    iColHasta As Integer)
    Dim Fila As Long, Col As Long, width As Single
    Dim maxWidth As Single, celdaText As String
    Dim saveFont As StdFont, oldScaleMode As Integer
    Dim NCampos As Integer
    'Variables para la cantidad de registros y columnas
    
    'Array para almacenar el ancho de cada columna
    NCampos = iColHasta - iColDesde + 1
    ReDim TempCol(NCampos)
    ReDim TempMaxWidth(NCampos)

    
    ' Si el número de registros es igual a 0 salimos
    
    ' Guardamos la fuente del DataGrid para luego reestablecerla
    Set saveFont = xGrid.Parent.Font
    Set xGrid.Parent.Font = xGrid.Font
    
    ' Ajustar el ScaleMode en vbTwips para el formulario
    oldScaleMode = xGrid.Parent.ScaleMode
    xGrid.Parent.ScaleMode = vbTwips
    
    maxWidth = 0
    
    'recorrer las filas:
    For Fila = 0 To 0
    
        For Col = iColDesde To iColHasta 'xGrid.Cols - 1
                    
          celdaText = Trim(xGrid.TextMatrix(Fila, Col))
            
          'Almacena el Ancho del texto de la celda del Datagrid
          width = xGrid.Parent.TextWidth(celdaText) + 150
                
          'Si el ancho de la celda es mayor se actualiza la variable maxWidth _
           y se establece el ancho de la columna
          If width > TempMaxWidth(Col) Then
            TempMaxWidth(Col) = width
            xGrid.ColWidth(Col) = width
          End If
            
        Next Col
    Next
      
    'Recorremos cada columna y le asignamos el ancho
    'For Col = 0 To NCampos - 1
    '    un_DataGrid.Columns(Col).width = TempMaxWidth(Col)
    'Next
    
    'restablecemos la fuente del DataGrid y el scaleMode
    Set xGrid.Parent.Font = saveFont
    xGrid.Parent.ScaleMode = oldScaleMode
    
    Erase TempCol
End Sub

Public Sub Ajustar_Columnas_FLEXGRID(xGrid As MSFlexGrid)

    Dim Fila As Long, Col As Long, width As Single
    Dim maxWidth As Single, celdaText As String
    Dim saveFont As StdFont, oldScaleMode As Integer
    Dim NCampos As Integer
    'Variables para la cantidad de registros y columnas
    
    'Array para almacenar el ancho de cada columna
    NCampos = xGrid.Cols
    ReDim TempCol(NCampos)
    ReDim TempMaxWidth(NCampos)

    
    ' Si el número de registros es igual a 0 salimos
    
    ' Guardamos la fuente del DataGrid para luego reestablecerla
    Set saveFont = xGrid.Parent.Font
    Set xGrid.Parent.Font = xGrid.Font
    
    ' Ajustar el ScaleMode en vbTwips para el formulario
    oldScaleMode = xGrid.Parent.ScaleMode
    xGrid.Parent.ScaleMode = vbTwips
    
    maxWidth = 0
    
    'recorrer las filas:
    For Fila = 0 To 0
    
        For Col = 0 To xGrid.Cols - 1
                    
          celdaText = Trim(xGrid.TextMatrix(Fila, Col))
            
          'Almacena el Ancho del texto de la celda del Datagrid
          width = xGrid.Parent.TextWidth(celdaText) + 150
                
          'Si el ancho de la celda es mayor se actualiza la variable maxWidth _
           y se establece el ancho de la columna
          If width > TempMaxWidth(Col) Then
            TempMaxWidth(Col) = width
            xGrid.ColWidth(Col) = width
          End If
            
        Next Col
    Next
      
    'Recorremos cada columna y le asignamos el ancho
    'For Col = 0 To NCampos - 1
    '    un_DataGrid.Columns(Col).width = TempMaxWidth(Col)
    'Next
    
    'restablecemos la fuente del DataGrid y el scaleMode
    Set xGrid.Parent.Font = saveFont
    xGrid.Parent.ScaleMode = oldScaleMode
    
    Erase TempCol
End Sub

Public Sub Ajustar_Columnas_FLEXGRID_BY_ROWS(xGrid As MSFlexGrid)

    Dim Fila As Long, Col As Long, width As Single
    Dim maxWidth As Single, celdaText As String
    Dim saveFont As StdFont, oldScaleMode As Integer
    Dim NCampos As Integer
    'Variables para la cantidad de registros y columnas
    
    'Array para almacenar el ancho de cada columna
    NCampos = xGrid.Cols
    ReDim TempCol(NCampos)
    ReDim TempMaxWidth(NCampos)

    
    ' Si el número de registros es igual a 0 salimos
    
    ' Guardamos la fuente del DataGrid para luego reestablecerla
    Set saveFont = xGrid.Parent.Font
    Set xGrid.Parent.Font = xGrid.Font
    
    ' Ajustar el ScaleMode en vbTwips para el formulario
    oldScaleMode = xGrid.Parent.ScaleMode
    xGrid.Parent.ScaleMode = vbTwips
    
    maxWidth = 0
    
    'recorrer las filas:
    For Fila = 0 To xGrid.Rows - 1
    
        For Col = 0 To xGrid.Cols - 1
                    
          celdaText = Trim(xGrid.TextMatrix(Fila, Col))
            
          'Almacena el Ancho del texto de la celda del Datagrid
          width = xGrid.Parent.TextWidth(celdaText) + 150
                
          'Si el ancho de la celda es mayor se actualiza la variable maxWidth _
           y se establece el ancho de la columna
          If width > TempMaxWidth(Col) Then
            TempMaxWidth(Col) = width
            xGrid.ColWidth(Col) = width
          End If
            
        Next Col
    Next Fila
      
    'Recorremos cada columna y le asignamos el ancho
    For Col = 0 To NCampos - 1
      xGrid.ColWidth(Col) = TempMaxWidth(Col)
    Next
    
    'restablecemos la fuente del DataGrid y el scaleMode
    Set xGrid.Parent.Font = saveFont
    xGrid.Parent.ScaleMode = oldScaleMode
    
    Erase TempCol
End Sub



Function Hay_Seleccion(xListBox As ListBox) As Boolean
  Dim i As Integer
  Dim Hay As Boolean
  Hay = False
  i = 0
  Do While i < xListBox.ListCount And Not Hay
    If xListBox.Selected(i) = True Then
       Hay = True
       Exit Do
    End If
    i = i + 1
  Loop
  Hay_Seleccion = Hay
End Function

Function No_Hay_Seleccion(xListBox As ListBox) As Boolean
  Dim i As Integer
  Dim Hay As Boolean
  Hay = False
  i = 0
  Do While i < xListBox.ListCount And Not Hay
    If xListBox.Selected(i) = False Then Hay = True
    i = i + 1
  Loop
  No_Hay_Seleccion = Hay
End Function


Sub AjustaColumnaDataGrid(xDataGrid As DataGrid, sColumna As String, iAncho As Integer)
  Dim i As Integer
  Dim e As Boolean
  For i = 0 To xDataGrid.Columns.Count - 1
    xDataGrid.Columns(i).Caption = UCase(xDataGrid.Columns(i).Caption)
  Next i
  
  i = 0
  e = False
  Do While i < xDataGrid.Columns.Count And Not e
    If UCase(sColumna) = xDataGrid.Columns(i).Caption Then
      e = True
    Else
      i = i + 1
    End If
  Loop
  
  If e Then xDataGrid.Columns(i).width = iAncho
End Sub

Function La_Columna_GRID(GRID As ubGrid, sCampo As String) As Integer
  Dim i As Integer
  Dim e As Boolean
  Dim CG As Integer
  i = 1
  e = False
  CG = -1
  Do While i <= GRID.Cols And Not e
    If UCase(GRID.TextMatrix(0, i)) = UCase(sCampo) Then
      e = True
    Else
      i = i + 1
    End If
  Loop
  If e Then CG = i
  La_Columna_GRID = CG
End Function

Function DBExiste(sTabla As String, sCampoBuscar As String, sBuscar As String) As Boolean
  Dim r As New ADODB.Recordset
  Dim s As String
  s = "select " & sCampoBuscar & " from [" & sTabla & "] where " & sCampoBuscar & " = '" & sBuscar & "'"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  DBExiste = Not r.EOF
  r.Close
  Set r = Nothing
End Function


Function DBValorStr(sTabla As String, sCampoBuscar As String, sBuscar As String, sCampoRetorno As String) As String
  Dim r As New ADODB.Recordset
  Dim s As String
  s = "select " & sCampoRetorno & " from [" & sTabla & "] where " & sCampoBuscar & " = '" & sBuscar & "'"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  s = ""
  If Not r.EOF Then
    If Not IsNull(r.Fields(sCampoRetorno).value) Then
      s = r.Fields(sCampoRetorno).value
    End If
  End If
  DBValorStr = s
End Function

Function DBValorLng(sTabla As String, sCampoBuscar As String, sBuscar As Long, sCampoRetorno As String) As Long
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim X As Long
  's = "select " & sCampoRetorno & " from [" & sTabla & "] where " & sCampoBuscar & " = " & CStr(sBuscar) & " "
  s = "select " & sCampoRetorno & " from [" & sTabla & "] where " & sCampoBuscar & " = " & CStr(sBuscar) & " order by " & sCampoRetorno
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  X = -1
  If Not r.EOF Then
    If Not IsNull(r.Fields(sCampoRetorno).value) Then
      X = r.Fields(sCampoRetorno).value
    End If
  End If
  DBValorLng = X
End Function

Function DBValorDouble(sTabla As String, sCampoBuscar As String, sBuscar As String, sCampoRetorno As String) As Double
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim X As Double
  's = "select " & sCampoRetorno & " from [" & sTabla & "] where " & sCampoBuscar & " = " & CStr(sBuscar) & " "
  s = "select " & sCampoRetorno & " from [" & sTabla & "] where " & sCampoBuscar & " = '" & CStr(sBuscar) & "' order by " & sCampoRetorno
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  X = -1
  If Not r.EOF Then
    If Not IsNull(r.Fields(sCampoRetorno).value) Then
      X = r.Fields(sCampoRetorno).value
    End If
  End If
  DBValorDouble = X
End Function


Function DBValorVariant(sTabla As String, sCampoBuscar As String, sBuscar As Long, sCampoRetorno As String) As Variant
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim X As Variant
  s = "select " & sCampoRetorno & " from [" & sTabla & "] where " & sCampoBuscar & " = " & CStr(sBuscar) & " "
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  X = -1
  If Not r.EOF Then
    If Not IsNull(r.Fields(sCampoRetorno).value) Then
      X = r.Fields(sCampoRetorno).value
    End If
  End If
  DBValorVariant = X
End Function

Function Existe_Columna_GRID(ByRef xGrid As ubGrid, sCampo As String) As Boolean
  Dim i As Integer
  Dim e As Boolean
  e = False
  i = 1
  Do While i <= xGrid.Cols And Not e
    If UCase(xGrid.TextMatrix(0, i)) = UCase(sCampo) Then
      e = True
    Else
      i = i + 1
    End If
  Loop
  Existe_Columna_GRID = e
End Function

Function La_Tabla_Actual_Personas(sCliente As String, sSubCliente As String) As String
  Dim s As String
  Dim r As New ADODB.Recordset
  Dim sTabla As String
  
  If sCliente = "" Then
    La_Tabla_Actual_Personas = ""
    Exit Function
  End If
  
  sCliente = Mid(sCliente, 1, 6)
  If Trim(sSubCliente) = "" Or Trim(sSubCliente) = "-" Then
    sSubCliente = "0"
  Else
    sSubCliente = Mid(sSubCliente, 1, 6)
  End If
  
  sTabla = ""
  s = "select tabla from personas where cliente = " & sCliente & " and subcliente = " & sSubCliente & " order by id"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  Do While Not r.EOF
    sTabla = Trim(r.Fields("tabla").value)
    r.MoveNext
  Loop
  r.Close
  Set r = Nothing
  
  La_Tabla_Actual_Personas = sTabla
End Function

Function Campo_de_la_Tabla(sTabla As String, sCampo As String) As Boolean
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim i As Integer, e As Boolean
  s = "select * from [" & sTabla & "]"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  e = False
  i = 0
  Do While i < r.Fields.Count And Not e
    If UCase(r.Fields(i).Name) = UCase(sCampo) Then e = True
    i = i + 1
  Loop
  r.Close
  Set r = Nothing
  Campo_de_la_Tabla = e
End Function

Public Sub Inicializar_Personas_Marca(sTabla As String, sCaracter As String)
  Dim s As String
  
  'alter table <table_name> disable trigger {<trigger_name> | all}
  
  s = "alter table [" & sTabla & "] disable trigger [TRG_" & sTabla & "]"
  Modulo.ExecSQL s
  s = "update [" & sTabla & "] set Marca = '" & sCaracter & "'"
  Modulo.ExecSQL s
  s = "alter table [" & sTabla & "] enable trigger [TRG_" & sTabla & "]"
  Modulo.ExecSQL s
End Sub

Public Sub Marcar_Personas_ID(sTabla As String, sID As String)
  Dim s As String
  s = "update [" & sTabla & "] set " & _
      "Marca = '" & MARCA_IMPRESION_CARD5 & "' where " & _
      "ID    = " & sID & " "
  Modulo.ExecSQL s
End Sub

Public Sub Marcar_Personas_CEDULA(sTabla As String, sCedula As String)
  Dim s As String
  s = "update [" & sTabla & "] set " & _
      "Marca  = '" & MARCA_IMPRESION_CARD5 & "' where " & _
      "Cedula = '" & sCedula & "'"
  Modulo.ExecSQL s
End Sub

Function FechaInvertida(sFechaDDMMYYYY As String) As String   'YYYY_MM_DD
  Dim s As String
  s = ""
  If Len(sFechaDDMMYYYY) = 10 Then 'Tiene separadores ... / -
    '- Fecha: DD_MM_YYYY
    '-        1234567890
    s = Mid(sFechaDDMMYYYY, 7, 4) & "/" & Mid(sFechaDDMMYYYY, 4, 2) & "/" & Mid(sFechaDDMMYYYY, 1, 2)
  End If
  FechaInvertida = s
End Function

Function FechaNormal(sFechaInvertidaYYYYMMDD As String) As String   'DD_MM_YYYY
  Dim s As String
  s = ""
  If Len(sFechaInvertidaYYYYMMDD) = 10 Then 'Tiene separadores ... / -
    '- Fecha: YYYY_MM_DD
    '-        1234567890
    s = Mid(sFechaInvertidaYYYYMMDD, 9, 2) & "/" & Mid(sFechaInvertidaYYYYMMDD, 6, 2) & "/" & Mid(sFechaInvertidaYYYYMMDD, 1, 4)
  End If
  FechaNormal = s
End Function

Public Sub Formatear_Cedula(ByRef sCed As String)
  Dim i As Integer
  Dim d As Double
  Dim s As String
  If IsNumeric(sCed) Then
    If Len(sCed) > 3 Then
       d = CDbl(sCed)
       s = Format(d, "#,0")
       sCed = s
    Else
       s = Mid(sCed, 1, 2) & "." & Mid(sCed, 3, 1)
       sCed = s
    End If
    
  End If
End Sub

Public Sub Fields_DataGrid_En_Mayusculas(xDataGrid As DataGrid)
  Dim i As Integer
  For i = 0 To xDataGrid.Columns.Count - 1
    xDataGrid.Columns(i).Caption = UCase(xDataGrid.Columns(i).Caption)
  Next i
End Sub

Public Sub Fields_FlexGrid_En_Mayusculas(xFlexGrid As MSFlexGrid)
  Dim i As Integer
  For i = 0 To xFlexGrid.Cols - 1 - 1
    xFlexGrid.TextMatrix(0, i) = UCase(xFlexGrid.TextMatrix(0, i))
  Next i
End Sub

Public Sub Actualizar_Pago_Cliente(lCliente As Long, lSubCliente As Long, dMontoPAGO As Double)
  Dim s As String, sM As String, s1 As String
  Dim sSigno As String
  
  If dMontoPAGO > 0 Then sSigno = "+" Else sSigno = "-"
  
  sM = Trim(Str(Abs(dMontoPAGO)))
  If lSubCliente = 0 Then
      's = "update clientes set Deuda = Deuda + " & sM & " where codigo = " & CStr(lCliente)
      s1 = "update clientes set Pagos = Pagos " & sSigno & " " & sM & " where codigo = " & CStr(lCliente)
    's = "update clientes set pagos = pagos " & sSigno & " " & sM & " where codigo = " & CStr(lCliente)
    's1 = "update clientes set saldo = deuda - pagos where codigo = " & CStr(lCliente)
  Else
      's = "update subclientes set Deuda = Deuda + " & sM & " where cliente = " & CStr(lCliente) & " and id = " & CStr(lSubCliente) & " "
      s1 = "update subclientes set Pagos = Pagos " & sSigno & " " & sM & " where cliente = " & CStr(lCliente) & " and id = " & CStr(lSubCliente) & " "

    's = "update subclientes set pagos = pagos " & sSigno & " " & sM & " where cliente = " & CStr(lCliente) & " and id = " & CStr(lSubCliente) & " "
    's1 = "update subclientes set saldo = deuda - pagos where cliente = " & CStr(lCliente) & " and id = " & CStr(lSubCliente) & " "
  End If
  
  'Modulo.ExecSQL s
  Modulo.ExecSQL s1

End Sub

Public Sub Registrar_Pago_Cliente(lCliente As Long, lSubCliente As Long, dMontoPAGO As Double)
  Dim s As String, sM As String, s1 As String
  Dim sSigno As String
  
  If dMontoPAGO > 0 Then sSigno = "+" Else sSigno = "-"
  
  sM = Trim(Str(Abs(dMontoPAGO)))
  If lSubCliente = 0 Then
      's = "update clientes set Deuda = Deuda + " & sM & " where codigo = " & CStr(lCliente)
      s1 = "update clientes set Pagos = Pagos + " & sM & " where codigo = " & CStr(lCliente)
    's = "update clientes set pagos = pagos " & sSigno & " " & sM & " where codigo = " & CStr(lCliente)
    's1 = "update clientes set saldo = deuda - pagos where codigo = " & CStr(lCliente)
  Else
      's = "update subclientes set Deuda = Deuda + " & sM & " where cliente = " & CStr(lCliente) & " and id = " & CStr(lSubCliente) & " "
      s1 = "update subclientes set Pagos = Pagos + " & sM & " where cliente = " & CStr(lCliente) & " and id = " & CStr(lSubCliente) & " "

    's = "update subclientes set pagos = pagos " & sSigno & " " & sM & " where cliente = " & CStr(lCliente) & " and id = " & CStr(lSubCliente) & " "
    's1 = "update subclientes set saldo = deuda - pagos where cliente = " & CStr(lCliente) & " and id = " & CStr(lSubCliente) & " "
  End If
  
  'Modulo.ExecSQL s
  Modulo.ExecSQL s1

End Sub




Public Sub Actualizar_Deuda_Cliente(lCliente As Long, lSubCliente As Long, dMontoDEUDA As Double)
  Dim s As String, sM As String, s1 As String
  Dim sSigno As String
  
  If dMontoDEUDA > 0 Then sSigno = "+" Else sSigno = "-"
  
  sM = Trim(Str(Abs(dMontoDEUDA)))
  If lSubCliente = 0 Then
    s = "update clientes set deuda = deuda " & sSigno & " " & sM & " where codigo = " & CStr(lCliente)
    's1 = "update clientes set saldo = deuda - pagos where codigo = " & CStr(lCliente)
  Else
    s = "update subclientes set deuda = deuda " & sSigno & " " & sM & " where cliente = " & CStr(lCliente) & " and id = " & CStr(lSubCliente) & " "
    's1 = "update subclientes set saldo = deuda - pagos where cliente = " & CStr(lCliente) & " and id = " & CStr(lSubCliente) & " "
  End If
  
  Modulo.ExecSQL s
  'Modulo.ExecSQL s1
End Sub

Public Function TIENE_Subcliente(lCliente As Long) As Boolean
  Dim s As String
  Dim r As New ADODB.Recordset
  Dim bT As Boolean
  s = "select * from subclientes where cliente = " & CStr(lCliente) & " "
  bT = False
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  bT = Not r.EOF
  r.Close
  Set r = Nothing
  TIENE_Subcliente = bT
End Function

Public Function El_Primer_Subcliente(lCliente As Long) As Long
  Dim s As String
  Dim r As New ADODB.Recordset
  Dim lPS As Long
  s = "select * from subclientes where cliente = " & CStr(lCliente) & " order by id"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  lPS = 0
  If Not r.EOF Then lPS = r.Fields("id").value
  r.Close
  Set r = Nothing
  El_Primer_Subcliente = lPS
End Function

Public Sub Actualizar_Existencia_Producto(sCodigo As String, iExistencia As Long)
  Dim s As String, sM As String, s1 As String
  Dim sSigno As String
  Dim iEx As Long
  
  If iExistencia > 0 Then sSigno = "+" Else sSigno = "-"
  
  iEx = Abs(iExistencia)
  
  s = "update productos set Existencia = Existencia " & sSigno & " " & CStr(iEx) & " where codigo = '" & sCodigo & "'"
  
  Modulo.ExecSQL s
End Sub

Public Function Producto_DESCRIPCION(sCodigo As String) As String
  Dim s As String
  Dim r As New ADODB.Recordset
  s = "select descripcion from productos where codigo = '" & sCodigo & "'"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  s = ""
  If Not r.EOF Then
    If Not IsNull(r.Fields("descripcion").value) Then
      s = Trim(r.Fields("descripcion").value)
    End If
  End If
  r.Close
  Set r = Nothing
  Producto_DESCRIPCION = s
End Function

Public Function DBPrecioEspecial(lCliente As Long, lSubCliente As Long, sCodigoProducto As String) As Double
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim dP As Double
  
  dP = -1#
  s = "SELECT MAX(Iddiseño) AS IdDiseño From PreciosEspeciales WHERE " & _
      "cliente        = " & CStr(lCliente) & " and " & _
      "subcliente     = " & CStr(lSubCliente) & " and " & _
      "codigoproducto = '" & sCodigoProducto & "'"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If r.EOF = False And IsNull(r!IDDiseño) = False Then
      s = "select Precio from PreciosEspeciales where " & _
      "cliente        = " & CStr(lCliente) & " and " & _
      "subcliente     = " & CStr(lSubCliente) & " and " & _
      "codigoproducto = '" & sCodigoProducto & "' And " & _
      "IDDiseño       =  " & r!IDDiseño
      r.Close
      r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
      If Not r.EOF Then
        If Not IsNull(r.Fields("precio").value) Then
          dP = r.Fields("precio").value
        End If
      End If
      r.Close
      Set r = Nothing
  End If
  DBPrecioEspecial = dP
End Function

Public Function EXISTE_LIST(xList As ListBox, sValor As String) As Boolean
  Dim i As Integer
  Dim e As Boolean
  i = 0
  e = False
  Do While i < xList.ListCount And Not e
    If Trim(xList.List(i)) = Trim(sValor) Then e = True
    i = i + 1
  Loop
  EXISTE_LIST = e
End Function

Public Function POS_LIST(xList As ListBox, sValor As String) As Integer
  Dim i As Integer
  Dim e As Boolean
  i = 0
  e = False
  Do While i < xList.ListCount And Not e
    If Trim(xList.List(i)) = Trim(sValor) Then
      e = True
    Else
      i = i + 1
    End If
  Loop
  If e Then POS_LIST = i Else POS_LIST = -1
End Function

Public Function VALOR_LIST_SECOND(xList1 As ListBox, xList2 As ListBox, sValor As String) As String
  Dim i As Integer
  Dim s As String
  Dim e As Boolean
  i = 0
  e = False
  Do While i < xList1.ListCount And Not e
    If xList1.List(i) = sValor Then
      e = True
    Else
      i = i + 1
    End If
  Loop
  s = ""
  If e Then
    s = xList2.List(i)
    VALOR_LIST_SECOND = s
  End If
End Function


Public Function LA_RUTA_DEL_CLIENTE(sCliente As String, sSCliente As String) As String
  Dim s As String, sRutaCliente As String, sRutaSubCliente As String
  Dim sNomCliente As String, sNomSubCliente As String
  Dim r As New ADODB.Recordset, sF As String
  Dim sNomArchivoInicial As String, sNomArchivoNuevo As String
  
  Dim iCliente As Long, iSCliente As Long
  
  LA_RUTA_DEL_CLIENTE = ""
  
  sRutaCliente = ""
  
  If UCase(sCliente) = "NUEVO" Then Exit Function
    
  's = "select RutaDestinoDatosCliente from Opciones"
  'r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  'If Not r.EOF Then
  '  If Not IsNull(r.Fields("RutaDestinoDatosCliente").Value) Then
  '
  '    sRutaCliente = Trim(r.Fields("RutaDestinoDatosCliente").Value)
  '
  '  End If
  'End If
  'r.Close
  'Set r = Nothing
  
  sRutaCliente = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
  
  
  
  If sRutaCliente = "" Then
    MsgBox "Falta Configurar Ruta de Datos del Cliente, Revise...", vbCritical, "Información"
    Exit Function
  End If
  
  
  sCliente = Mid(sCliente, 1, 6)
  If Trim(sSCliente) = "" Or Trim(sSCliente) = "-" Then sSCliente = "0" Else sSCliente = Mid(sSCliente, 1, 6)
  
  iCliente = CLng(sCliente)
  iSCliente = CLng(sSCliente)
    
  
  sNomCliente = ""
  sNomSubCliente = ""
  
  '1.Obtener Nombre del Cliente:
  s = "select nombre from clientes where codigo = " & CStr(iCliente) & " "
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If Not r.EOF Then
    If Not IsNull(r.Fields("nombre").value) Then
      sNomCliente = Trim(r.Fields("nombre").value)
    End If
  End If
  r.Close
     
  If sNomCliente <> "" Then
  
    '2.Obtener Nombre del Sub-cliente:
    sNomSubCliente = ""
    If iSCliente > 0 Then
      s = "select nombre from subclientes where cliente = " & CStr(iCliente) & " and id = " & iSCliente & " "
      r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
      If Not r.EOF Then
        If Not IsNull(r.Fields("nombre").value) Then
          sNomSubCliente = Trim(r.Fields("nombre").value)
        End If
      End If
      r.Close
  
      If sNomSubCliente = "" Then Exit Function
    End If
    
    sRutaCliente = sRutaCliente & "\" & sNomCliente & IIf(sNomSubCliente <> "", "\" & sNomSubCliente, "") & "\IMAGENES"
    
    LA_RUTA_DEL_CLIENTE = sRutaCliente
    
  End If
  
  Set r = Nothing
  
End Function


Public Function LA_RUTA_FOTO_DEL_CLIENTE(sCliente As String, sSCliente As String) As String
  Dim s As String, sRutaCliente As String, sRutaSubCliente As String
  Dim sNomCliente As String, sNomSubCliente As String
  Dim r As New ADODB.Recordset, sF As String
  Dim sNomArchivoInicial As String, sNomArchivoNuevo As String
  
  Dim iCliente As Long, iSCliente As Long
  
  LA_RUTA_FOTO_DEL_CLIENTE = ""
  
  If sCliente = "Nuevo" Then Exit Function
  
  sRutaCliente = ""
  
  sRutaCliente = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
 
  's = "select RutaDestinoDatosCliente from Opciones"
  'r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  'If Not r.EOF Then
  '  If Not IsNull(r.Fields("RutaDestinoDatosCliente").Value) Then
  '
  '    sRutaCliente = Trim(r.Fields("RutaDestinoDatosCliente").Value)
  '
  '  End If
  'End If
  'r.Close
  'Set r = Nothing
  
  If sRutaCliente = "" Then
    MsgBox "Falta Configurar Ruta de Datos del Cliente, Revise...", vbCritical, "Información"
    Exit Function
  End If
  
  
  sCliente = Mid(sCliente, 1, 6)
  If Trim(sSCliente) = "" Or Trim(sSCliente) = "-" Then sSCliente = "0" Else sSCliente = Mid(sSCliente, 1, 6)
  
  iCliente = CLng(sCliente)
  iSCliente = CLng(sSCliente)
    
  
  sNomCliente = ""
  sNomSubCliente = ""
  
  '1.Obtener Nombre del Cliente:
  s = "select nombre from clientes where codigo = " & CStr(iCliente) & " "
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If Not r.EOF Then
    If Not IsNull(r.Fields("nombre").value) Then
      sNomCliente = Trim(r.Fields("nombre").value)
    End If
  End If
  r.Close
     
  If sNomCliente <> "" Then
  
    '2.Obtener Nombre del Sub-cliente:
    sNomSubCliente = ""
    If iSCliente > 0 Then
      s = "select nombre from subclientes where cliente = " & CStr(iCliente) & " and id = " & iSCliente & " "
      r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
      If Not r.EOF Then
        If Not IsNull(r.Fields("nombre").value) Then
          sNomSubCliente = Trim(r.Fields("nombre").value)
        End If
      End If
      r.Close
  
      If sNomSubCliente = "" Then Exit Function
    End If
    
    sRutaCliente = sRutaCliente & "\" & sNomCliente & IIf(sNomSubCliente <> "", "\" & sNomSubCliente, "") & "\FOTOS"
    
    LA_RUTA_FOTO_DEL_CLIENTE = sRutaCliente
    
  End If
  
  Set r = Nothing
  
End Function


Public Function LAS_INDICACIONES_CLIENTE(sCliente As String, sSCliente As String) As String
  Dim s As String
  Dim r As New ADODB.Recordset, sF As String
  Dim iCliente As Long, iSCliente As Long
  Dim sIndicaciones As String
  
  LAS_INDICACIONES_CLIENTE = ""
  
  sCliente = Mid(sCliente, 1, 6)
  If Trim(sSCliente) = "" Then sSCliente = "0" Else sSCliente = Mid(sSCliente, 1, 6)
  
  If sCliente = "" Then
    LAS_INDICACIONES_CLIENTE = ""
    Exit Function
  End If
  
  iCliente = CLng(sCliente)
  If Trim(sSCliente) = "-" Then sSCliente = "0"
  
  iSCliente = CLng(sSCliente)
  
  s = "select especificaciones from FormatoDiseno where " & _
      "cliente    = " & CStr(iCliente) & " and " & _
      "subcliente = " & CStr(iSCliente) & " "
      
  sIndicaciones = ""
  
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  
  If Not r.EOF Then
    If Not IsNull(r.Fields("especificaciones").value) Then
      sIndicaciones = Trim(r.Fields("especificaciones").value)
    End If
  End If
  
  r.Close
  Set r = Nothing
  
  LAS_INDICACIONES_CLIENTE = sIndicaciones
  
End Function

Public Function HAY_SELECCION_FG(Flex As MSFlexGrid, iColumna As Integer, sCadenaHAY As String) As Boolean
  Dim i As Integer
  Dim Hay As Boolean
  i = 1
  Hay = False
  Do While i < Flex.Rows And Not Hay
    If Flex.TextMatrix(i, iColumna) = sCadenaHAY Then Hay = True
    i = i + 1
  Loop
  HAY_SELECCION_FG = Hay
End Function

Public Function GenerarLocalizador() As String
  Dim i As Integer
  Dim s As String
  Dim Letra As Integer  'Desde 65..90  'A'..'Z'
  Dim Digito As Integer 'Desde 48..57  '0'..'9'
  Dim LD As Integer      'Letra o Digito
  
  Randomize
  
  s = ""
  ' Generate random value between 1 and 6.
  'Dim value As Integer = CInt(Int((6 * Rnd()) + 1))
  
  For i = 1 To 10
  
    'Generate random value between 1 and 10 : 50%Letra / 50%Digito
    LD = CInt(Int((10 * Rnd()) + 1))
    If LD <= 5 Then '50% es Letra
      'Generate random value between 65 and 90.
      Letra = CInt(Int((90 - 65 + 1) * Rnd() + 65))
              'CLng((Minimo - Maximo) * Rnd + Maximo)
      s = s & Chr(Letra)
    Else
      'Generate random value between 48 and 57.
      Digito = CInt(Int((57 - 48 + 1) * Rnd() + 48))
      s = s & Chr(Digito)
    End If
    
  Next i
  
  GenerarLocalizador = s
  
End Function

Public Function Localizador_Por_ID(sTabla As String, sID As String) As String
  Dim s As String
  Dim r As New ADODB.Recordset
  Dim sL As String
  s = "select Localizador from [" & sTabla & "] where ID = " & sID & " "
  sL = ""
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If Not r.EOF Then
    If Not IsNull(r.Fields("localizador").value) Then
      sL = r.Fields("localizador").value
    End If
  End If
  r.Close
  Set r = Nothing
  Localizador_Por_ID = sL
End Function

Public Function Size_Str_Campo(sTabla As String, sCampo As String) As Integer
  Dim s As String
  Dim r As New ADODB.Recordset
  Dim iSize As Integer
  iSize = 0
  s = "select " & sCampo & " from [" & sTabla & "]"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  iSize = IIf(IsNull(r.Fields(sCampo).DefinedSize), 0, r.Fields(sCampo).DefinedSize)
  r.Close
  Set r = Nothing
  Size_Str_Campo = iSize
End Function

Public Sub Resumen_Cuenta_Cliente(sCliente As String, sSubCliente As String, ByRef dCargos As Double, ByRef dAbonos As Double, ByRef dSaldo As Double)
  Dim s As String
  Dim r As New ADODB.Recordset
  Dim LC As Long, lSC As Long
  
  dCargos = 0#
  dAbonos = 0#
  dSaldo = 0#
  
  If Trim(sCliente) = "" Or Trim(sCliente) = "-" Then Exit Sub
    
  LC = CLng(Trim(Mid(sCliente, 1, 6)))
  
  If Trim(sSubCliente) = "" Or Trim(sSubCliente) = "-" Then
    sSubCliente = "0"
  Else
    sSubCliente = Mid(sSubCliente, 1, 6)
  End If
  
  lSC = CLng(sSubCliente)
  
  If lSC = 0 Then
    s = "select deuda, pagos, saldo from clientes where codigo = " & CStr(LC)
  Else
    s = "select deuda, pagos, saldo from subclientes where cliente = " & CStr(LC) & " and id = " & CStr(lSC)
  End If
  
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  
  If Not r.EOF Then
    dCargos = IIf(IsNull(r.Fields("deuda").value), 0, r.Fields("deuda").value)
    'dCargos = r.Fields("deuda").value
    dAbonos = IIf(IsNull(r.Fields("pagos").value), 0, r.Fields("pagos").value)
    'dAbonos = r.Fields("pagos").value
    dSaldo = IIf(IsNull(r.Fields("saldo").value), 0, r.Fields("saldo").value)
  End If
  
  r.Close
  Set r = Nothing

End Sub

Public Sub Resumen_Cuenta_Cliente_Codigo(sCliente As String, ByRef dCargos As Double, ByRef dAbonos As Double, ByRef dSaldo As Double)
  Dim s As String
  Dim r As New ADODB.Recordset
  Dim LC As Long, lSC As Long
  
  dCargos = 0#
  dAbonos = 0#
  dSaldo = 0#
  
  If Trim(sCliente) = "" Or Trim(sCliente) = "-" Then Exit Sub
    
  LC = CLng(Trim(Mid(sCliente, 1, 6)))
  
  s = "select deuda, pagos, saldo from clientes where codigo = " & CStr(LC)
  
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  
  If Not r.EOF Then
    dCargos = r.Fields("deuda").value
    dAbonos = r.Fields("pagos").value
    dSaldo = r.Fields("saldo").value
  End If
  
  r.Close
  Set r = Nothing

End Sub

Public Sub Campos_DataGrid_En_Combo(xDataGrid As DataGrid, cCombo As ComboBox)
  Dim i As Integer
  cCombo.Clear
  For i = 0 To xDataGrid.Columns.Count - 1
    cCombo.AddItem xDataGrid.Columns(i).Caption
  Next i
End Sub

Public Sub DepurarTitutlosFlexGrid(xFG As MSFlexGrid)
  Dim i As Integer
  Dim s As String
  For i = 0 To xFG.Cols - 1
    s = Modulo.DepurarStr(xFG.TextMatrix(0, i), " ")
    xFG.TextMatrix(0, i) = UCase(s)
  Next i
End Sub


Public Sub Tablas_De_BD_Access(xList As ListBox, ConnectString As String)
   'Dim ConnectString As String
   Dim ADOXConnection As Object
   Dim ADODBConnection As Object
   Dim Table As Variant

   'ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;data source=" & Database
   
   Set ADOXConnection = CreateObject("ADOX.Catalog")
   Set ADODBConnection = CreateObject("ADODB.Connection")
   ADODBConnection.Open ConnectString
   ADOXConnection.ActiveConnection = ADODBConnection
   xList.Clear
   For Each Table In ADOXConnection.Tables
     If LCase(Mid(Table.Name, 1, 4)) <> "msys" Then
       xList.AddItem LCase(Table.Name)
     End If
      'If LCase(Table.Name) = LCase(TableName) Then
      '   IsExistingTable = True
      '   Exit For
      'End If
   Next
   ADODBConnection.Close

End Sub

Public Sub AgregarLogs(sOperacion As String)
  Dim s As String
  s = "insert into Logs (fecha,hora,usuario,estacion,operacion) values ('" & _
       Format(Date, "yyyymmdd") & "','" & _
       Format(Time, "HH:mm") & "','" & _
       Modulo.USUARIO_ACTUAL & "','" & _
       Modulo.ESTACION & "','" & _
       sOperacion & "')"
  Modulo.ExecSQL s
End Sub

Public Function Usuario_VALIDO(sUsuario As String, sPass As String) As Boolean
  Dim s As String
  Dim r As New ADODB.Recordset
  
  If DBConexionSQL.State = adStateClosed Then Abrir_BD
  
  s = "select * from Usuarios where " & _
      "usuario = '" & sUsuario & "' and " & _
      "clave   = '" & sPass & "'"
      
  r.Open s, Modulo.DBConexionSQL, adOpenStatic
  
  If Not r.EOF Then
    Usuario_VALIDO = True
    
    If r.Fields("estatus").value = "S" Then Usuario_VALIDO = False
    
  Else
    Usuario_VALIDO = False
  End If
  
  r.Close
  Set r = Nothing
End Function

Public Function Usuario_VALIDO2(sUsuario As String, sPass As String, ByRef sNivel As String, ByRef sPermisos As String) As Boolean
  Dim s As String
  Dim r As New ADODB.Recordset
  On Error GoTo falla
  If DBConexionSQL.State = adStateClosed Then Abrir_BD
   
  
  s = "select * from Usuarios where " & _
      "usuario = '" & sUsuario & "' and " & _
      "clave   = '" & sPass & "'"
      
  r.Open s, Modulo.DBConexionSQL, adOpenStatic
  
  If Not r.EOF Then
    sNivel = r.Fields("nivel").value
    sPermisos = r.Fields("permisos").value
    Usuario_VALIDO2 = True
    
    If r.Fields("estatus").value = "S" Then Usuario_VALIDO2 = False
  Else
    sNivel = ""
    sPermisos = ""
    Usuario_VALIDO2 = False
  End If
  
  r.Close
  Set r = Nothing
falla:
  If Err.Number <> 0 Then
      MsgBox Err.Number & "::" & Err.Description, vbCritical
  End If
End Function

Public Function CarnetsEntregados(sCliente As String, sSubCliente As String, sDesde As String, sHasta As String) As Long
  Dim s As String
  Dim sT1 As String, sT2 As String
  Dim r As New ADODB.Recordset
  Dim k As Long
  
  'If sSubCliente = "" Or sSubCliente = "-" Then sSubCliente = ""
  On Error GoTo falla
  sT1 = Modulo.La_Tabla_Actual_Personas(sCliente, sSubCliente)
  sT2 = "H" & sT1
  
  '-- I : Buscar en la Historica de resguardo:
  
  s = "select * from [" & sT2 & "] where " & _
      "Fecha >= '" & sDesde & "' and " & _
      "Fecha <= '" & sHasta & "' and " & _
      "Contador > 0 order by Fecha"
      
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  k = 0
  Do While Not r.EOF
    If Not IsNull(r.Fields("CONTADOR").value) Then
      k = k + r.Fields("CONTADOR").value
    End If
    r.MoveNext
  Loop
  
  r.Close
  
  '-- II : Buscar en la actual de personas:
  
  s = "select * from [" & sT1 & "] where " & _
      "Fecha >= '" & sDesde & "' and " & _
      "Fecha <= '" & sHasta & "' and " & _
      "Contador > 0 order by Fecha"
      
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  Do While Not r.EOF
    If Not IsNull(r.Fields("CONTADOR").value) Then
      k = k + r.Fields("CONTADOR").value
    End If
    r.MoveNext
  Loop
  
  r.Close
  
  Set r = Nothing
  
  CarnetsEntregados = k
falla:
  If Err.Number <> 0 Then
  
  End If
End Function

Public Sub sCrearArchivoExcel(argTabla As String, argNombreCliente As String)
 On Error GoTo fallo
 Dim xlApp As Excel.Application
 Dim xlWB As Excel.Workbook
 Dim xlWS As Excel.Worksheet
 Dim lReg As New ADODB.Recordset
 Dim lRutaActual As String
 Dim argDestino As String
 Dim i As Integer
 Dim j ' campo en el archiv excel
 
  argDestino = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
  argDestino = argDestino & "\" & argNombreCliente & "\LISTADO " & argNombreCliente & " OFICINA.xlsx"

 
 Set xlApp = New Excel.Application
 Set xlWB = xlApp.Workbooks.Add
 Set xlWS = xlWB.Worksheets.Add
 
 lReg.Open "select * from [" & argTabla & "]", Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
 j = 1
 For i = 0 To lReg.Fields.Count - 1
    If UCase(lReg.Fields(i).Name) <> "ID" And _
       UCase(lReg.Fields(i).Name) <> "TIENE_FOTO" And _
       UCase(lReg.Fields(i).Name) <> "MARCA" And _
       UCase(lReg.Fields(i).Name) <> "FECHA" And _
       UCase(lReg.Fields(i).Name) <> "CONTADOR" And _
       UCase(lReg.Fields(i).Name) <> "CREACION" And _
       UCase(lReg.Fields(i).Name) <> "FOTO" Then
          xlWS.Cells(1, j).value = lReg.Fields(i).Name
          j = j + 1
    End If
    
 Next i
 

 'xlWS.Cells(1, 3).Value = "Mundo"
 ' Nos guardara esto en el fichero excel llamdo prueba.xls
 'Clipboard.Clear
 'Clipboard.SetText argDestino
 
 xlWS.SaveAs argDestino
 xlApp.Quit
 Set xlWS = Nothing
 Set xlWB = Nothing
 Set xlApp = Nothing
 ''FileCopy App.Path & "\" & argNombreCliente & ".xls", argDestino
 'MsgBox ("Archivo creado en C:\"), vbInformation
 Exit Sub
fallo:
 MsgBox ("Error al crear el Archivo Excel" & Err.Number & "::" & Err.Description), vbInformation
 xlApp.Quit
 Set xlWS = Nothing
 Set xlWB = Nothing
 Set xlApp = Nothing

End Sub

