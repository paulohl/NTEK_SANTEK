Attribute VB_Name = "Modulo"
Option Explicit



Dim CONEXION_SQL As String
Dim ODBC As String

Public IP_Servidor As String
Public DBConexionSQL As New ADODB.Connection
Public DBComandoSQL As New ADODB.Command
Public DBUsuario As String
Public ESTACION As String

Public Const APPNAME = "SANTEK-SERVER"

Public fModalResult As String
Public Const fModalResultOK = "OK"
Public Const fModalResultCANCEL = "CANCEL"

Public Const MARCA_IMPRESION_CARD5 = "I"

'Declaración de las funciones API's para escribir y leer archivos INI.
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Sub Main()
  Dim sIni As String
  IP_Servidor = "" '"192.168.10.150"
  
  'ODBC = "SANTEK"
  
  'CONEXION_SQL = "Provider=MSDASQL.1;" & _
                 "Persist Security Info=False;" & _
                 "Data Source=" & ODBC & ";" & _
                 "Initial Catalog=santek"
                 
  'CONEXION_SQL = "driver={SQL Server};server=" & IP_Servidor & ";" & _
                 "database=santek;Username=sa;PWD=sql123;"
                 
  sIni = App.Path & "\" & "Config.Ini"
  If Dir(sIni) = "" Then
    MsgBox "Falta Configuración en [Config.Ini]...", vbCritical, "Información"
    End
  End If
  
  IP_Servidor = Modulo.INI_Read(sIni, "CONEXION", "IP_SERVIDOR", "")
  
  If IP_Servidor = "" Then
    MsgBox "Falta IP de Servidor...", vbCritical, "Información"
    End
  End If
    
  CONEXION_SQL = "Provider=SQLOLEDB.1;" & _
                 "Password=sql123;" & _
                 "Persist Security Info=True;" & _
                 "User ID=sa;" & _
                 "Initial Catalog=santek;" & _
                 "Data Source=" & IP_Servidor
                 
  DBConexionSQL.ConnectionString = CONEXION_SQL
  
  If Not Abrir_BD() Then End
  
  Load fPpal
  
  fPpal.Show
End Sub

Function Abrir_BD() As Boolean
  On Error Resume Next
  
  Load fMensaje
  fMensaje.Caption = "Conectando, Espere..."
  fMensaje.Show
  
  DBConexionSQL.Open
  
  Unload fMensaje
  
  If Err.Number <> 0 Then
    MsgBox "Imposible Conectar con Servidor [ODBC=" & ODBC & "]" & vbCrLf & Err.Description, vbCritical, "Información"
    Abrir_BD = False
  Else
    Abrir_BD = True
  End If

End Function



'Función para leer los datos en archivos INI:
Public Function INI_Read(Filename As String, Key_Value As String, Key_Name As String, Optional ByVal Default As String) As String
  'On Error GoTo ErrOut
  Dim Size As Integer
  Dim value As String

  'Comprobamos que el archivo existe.
  'If Not SYS_FileExists(Filename) Then Err.Raise 53
  If Dir(Filename) = "" Then Err.Raise 53

  'Se define el tamaño maximo de caracteres
  'que podra tener la variable Value
  value = Space(200)
  'Se utiliza la función para obtener
  'el valor de la clave
  Size = GetPrivateProfileString(Key_Value, Key_Name, "", value, Len(value), Filename)
  'Si el tamaño es mayor a -1 entonces
  'se ha encontrado el valor de la clave
  If Size > 0 Then
  value = Left$(value, Size)
  Else
  INI_Read = Default
  End If

  'Devolver el dato...
  'Verificar que el dato no sea nulo,
  'en caso de ser nulo de se devuelve
  'el valor por defecto (Default)
  If Len(value) Then
  INI_Read = value
  Else
  INI_Read = Default
  End If
  Exit Function

ErrOut:
  INI_Read = Default
End Function

'Función para escrbir datos en archivos INI.
Public Function INI_Write(Filename As String, Key_Value As String, Key_Name As String, value As String) As Long
  'On Error GoTo ErrOut
  Dim Size As Integer

  'Escribimos el valor de la clave en el INI
  Size = WritePrivateProfileString(Key_Value, Key_Name, value, Filename)
  INI_Write = 1
  Exit Function

ErrOut:
  INI_Write = 0
End Function
