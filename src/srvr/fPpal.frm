VERSION 5.00
Begin VB.Form fPpal 
   Caption         =   "Server Monitor"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   3060
      Top             =   180
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFC0&
      Height          =   6300
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5085
   End
End
Attribute VB_Name = "fPpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const TBL = "EventosC5"

Private Sub Revisar_Eventos()
  Dim s As String, sTabla As String, lIDTabla As Long
  Dim lIDEvento As Long
  Dim r As New ADODB.Recordset
  
  Timer1.Enabled = False
  
  s = "select * from [" & TBL & "] where Procesado = 'N'"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  Do While Not r.EOF
    lIDEvento = r.Fields("id").Value
    lIDTabla = r.Fields("idTabla").Value
    sTabla = Trim(r.Fields("Tabla").Value)
    If sTabla <> "" Then
      '1. actualizar la Marca en la Tabla de Personas a "blancos":
      s = "update [" & sTabla & "] set Marca = '' where ID = " & CStr(lIDTabla)
      Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
      Modulo.DBComandoSQL.CommandText = s
      Modulo.DBComandoSQL.Execute
      
      '2. actualizar el registro
      s = "update [" & TBL & "] set Procesado = 'S' where id = " & CStr(lIDEvento)
      Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
      Modulo.DBComandoSQL.CommandText = s
      Modulo.DBComandoSQL.Execute
      
      s = Format(Now, "dd/mm/yy HH:mm") & " " & "Procesado ID " & CStr(lIDTabla) & " de " & sTabla
      List1.AddItem s
      GrabarLineaTXT s
    End If
    r.MoveNext
  Loop
  r.Close
  Set r = Nothing
  
  Timer1.Enabled = True
End Sub

Private Sub Form_Load()
  Dim s As String
  
  fPpal.Top = 0
  fPpal.Left = Screen.Width - fPpal.Width
  List1.Clear
  Timer1.Interval = 10000  'cada 10 segundos (1 seg = 1000 milisegundos)
  Timer1.Enabled = True
  fPpal.Caption = "Server Monitor - Esperando..."
  
  s = Format(Now, "dd/mm/yy HH:mm") & " Monitoreo Activado, chequeo cada " & CStr(Timer1.Interval / 1000) & " segundo(s)..."
  List1.AddItem s
  
  GrabarLineaTXT s
  
End Sub

Private Sub Timer1_Timer()
  fPpal.Caption = "Server Monitor - Procesando..."
  Revisar_Eventos
  fPpal.Caption = "Server Monitor - Esperando..."
End Sub

Private Sub GrabarLineaTXT(sLinea As String)
  Dim sArchivoTXT As String

  sArchivoTXT = App.Path & "\" & "EventosC5.TXT"
  
  If Dir(sArchivoTXT) = "" Then  'No existe...crearlo!
    
    Open sArchivoTXT For Output As #1
    
  Else
  
    Open sArchivoTXT For Append As #1 'Abrirlo para agregar!
    
  End If
  
  Print #1, sLinea
  
  Close #1
End Sub
