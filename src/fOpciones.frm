VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form fOpciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13740
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   13740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      Caption         =   "Producto Servicio de Fotografía (In Situ)"
      Height          =   705
      Left            =   6840
      TabIndex        =   31
      Top             =   2190
      Width           =   6800
      Begin VB.TextBox eSF 
         Height          =   285
         Left            =   870
         TabIndex        =   32
         Text            =   "Text6"
         Top             =   270
         Width           =   2085
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SERVICIO FOTO"
         Height          =   315
         Left            =   3030
         TabIndex        =   34
         Top             =   270
         Width           =   3645
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   300
         Width           =   540
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Producto Diseño de Carnet (Arte)"
      Height          =   705
      Left            =   6840
      TabIndex        =   27
      Top             =   1470
      Width           =   6800
      Begin VB.TextBox eCA 
         Height          =   285
         Left            =   870
         TabIndex        =   28
         Text            =   "Text6"
         Top             =   270
         Width           =   2085
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DISEÑO ARTE"
         Height          =   315
         Left            =   3030
         TabIndex        =   29
         Top             =   270
         Width           =   3645
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Software FOTOS"
      Height          =   675
      Left            =   0
      TabIndex        =   26
      Top             =   3120
      Width           =   6800
      Begin VB.CommandButton bFotos 
         Caption         =   "..."
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
         Left            =   6180
         TabIndex        =   18
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   90
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   240
         Width           =   6000
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "PC / Estación de Red"
      Height          =   705
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   6800
      Begin VB.TextBox ePC 
         Height          =   315
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "Text6"
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código de PC:"
         Height          =   195
         Left            =   2250
         TabIndex        =   25
         Top             =   300
         Width           =   1020
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Producto Carnet PVC"
      Height          =   705
      Left            =   6840
      TabIndex        =   21
      Top             =   750
      Width           =   6800
      Begin VB.TextBox eCodigo 
         Height          =   285
         Left            =   870
         TabIndex        =   19
         Text            =   "Text6"
         Top             =   270
         Width           =   2085
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CARNET TIPO PVC"
         Height          =   315
         Left            =   3030
         TabIndex        =   23
         Top             =   270
         Width           =   3645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   300
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Software CARD-5"
      Height          =   675
      Left            =   0
      TabIndex        =   20
      Top             =   2400
      Width           =   6800
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   90
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   240
         Width           =   6000
      End
      Begin VB.CommandButton bCard5 
         Caption         =   "..."
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
         Left            =   6180
         TabIndex        =   16
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Para Envio de Correo (Errores en Listado de Personas)"
      Height          =   3315
      Left            =   6840
      TabIndex        =   10
      Top             =   2940
      Width           =   6800
      Begin VB.CommandButton cmdVerificarEmail 
         Caption         =   "Verificar Email"
         Height          =   435
         Left            =   5580
         TabIndex        =   41
         Top             =   2640
         Width           =   915
      End
      Begin VB.TextBox txtConCopia 
         Height          =   285
         Left            =   4860
         TabIndex        =   40
         Top             =   2280
         Width           =   1635
      End
      Begin VB.TextBox txtContraseña 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   38
         Top             =   2340
         Width           =   1995
      End
      Begin VB.TextBox txtUsuario 
         Height          =   285
         Left            =   2640
         TabIndex        =   36
         Top             =   1980
         Width           =   1995
      End
      Begin VB.TextBox Text4 
         Height          =   855
         Left            =   90
         MaxLength       =   254
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Text            =   "fOpciones.frx":0000
         Top             =   1020
         Width           =   6600
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   90
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   480
         Width           =   6600
      End
      Begin VB.Label Label14 
         Caption         =   "E-mail con copia a:"
         Height          =   195
         Left            =   4920
         TabIndex        =   39
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Contraseña de Usuario de Correo:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2400
         Width           =   2595
      End
      Begin VB.Label Label12 
         Caption         =   "Nombre de Usuario de Correo:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Contenido del Mensaje (Plantilla):"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   810
         Width           =   2340
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Titulo del Mensaje:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   270
         Width           =   1335
      End
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   500
      Left            =   5820
      Picture         =   "fOpciones.frx":0008
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4380
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton bAceptar 
      Caption         =   "Guardar"
      Height          =   500
      Left            =   4620
      Picture         =   "fOpciones.frx":0592
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4380
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4650
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   690
      Width           =   6800
      Begin VB.CommandButton bDestino 
         Caption         =   "..."
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
         Left            =   6150
         TabIndex        =   7
         Top             =   1140
         Width           =   435
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   1140
         Width           =   6000
      End
      Begin VB.CommandButton bOrigen 
         Caption         =   "..."
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
         Left            =   6150
         TabIndex        =   4
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   90
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   480
         Width           =   6000
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Destino Datos de Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   900
         Width           =   2190
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Origen Plantilla de Datos de Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2925
      End
   End
End
Attribute VB_Name = "fOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oMail As clsCDOmail
Attribute oMail.VB_VarHelpID = -1

Private Sub bAceptar_Click()
  Dim s As String
  Dim r As New ADODB.Recordset
  
  Text1.Text = Trim(Text1.Text)
  Text2.Text = Trim(Text2.Text)
  If Text1.Text = "" Or Text2.Text = "" Then
    MsgBox "Faltan Datos, Revise...", vbCritical, "Información"
  Else
    SaveSetting APPNAME, "Opciones", "RutaOrigen", Text1.Text
    SaveSetting APPNAME, "Opciones", "RutaDestino", Text2.Text
    SaveSetting APPNAME, "Opciones", "RutaCard5", Text5.Text
    SaveSetting APPNAME, "Opciones", "RutaFotos", Text6.Text
    SaveSetting APPNAME, "Opciones", "Estacion", ePC.Text
    
    SaveSetting APPNAME, "Opciones", "TituloMensajeCorreo", Text3.Text
    SaveSetting APPNAME, "Opciones", "CuerpoMensajeCorreo", Text4.Text
        
    SaveSetting APPNAME, "Opciones", "CodigoCarnet", eCodigo.Text
    SaveSetting APPNAME, "Opciones", "CodigoDiseno", eCA.Text
    SaveSetting APPNAME, "Opciones", "CodigoServicioFoto", eSF.Text
    
    SaveSetting APPNAME, "Opciones", "UsuarioEmail", txtUsuario.Text
    SaveSetting APPNAME, "Opciones", "ContraseñaEmail", txtContraseña.Text
    SaveSetting APPNAME, "Opciones", "ConCopiaEmail", txtConCopia.Text

    
    s = "update opciones set codigoproductopvc = '" & eCodigo.Text & "'"
    Modulo.ExecSQL s
    
    MsgBox "Datos Almacenados Correctamente...", vbInformation, "Información"
    
    Modulo.ESTACION = GetSetting(APPNAME, "Opciones", "Estacion", "")
    fSistema.Caption = TIT_SISTEMA & " - Estación:" & ESTACION
    
    
  End If
    
End Sub

Private Sub bCancelar_Click()
  vTemporal1 = ""
  Unload Me
End Sub

Private Sub bCard5_Click()
  Load fDir2
  vTemporal1 = ""
  
  If Text5.Text = "" Then
    fDir2.Drive1.Drive = Mid(App.Path, 1, 1)
    fDir2.Dir1.Path = GetPath(App.Path & "\", "\")
  Else
    fDir2.Drive1.Drive = Mid(Text5.Text, 1, 1)
    fDir2.Dir1.Path = GetPath(Text5.Text & "\", "\")
  End If
    
  If Text5.Text <> "" Then
    vTemporal1 = Text5.Text
    fDir2.Text1.Text = Text5.Text
  End If
  fDir2.Show vbModal
  If vTemporal1 <> "" Then Text5.Text = vTemporal1
End Sub

Private Sub bDestino_Click()
On Error GoTo falla
  Load fDir
  vTemporal1 = ""
  If Len(Text2.Text) = 3 Then
     Text2.Text = Mid(Text2.Text, 1, 2)
  End If
  If Text2.Text = "" Then
    fDir.Drive1.Drive = Mid(App.Path, 1, 1)
    fDir.Dir1.Path = GetPath(App.Path & "\", "\")
  Else
    fDir.Drive1.Drive = Mid(Text2.Text, 1, 1)
    fDir.Dir1.Path = GetPath(Text2.Text & "\", "\")
  End If
    
  If Text2.Text <> "" Then
    vTemporal1 = Text2.Text
    fDir.Text1.Text = Text2.Text
  End If
  fDir.Show vbModal
  If vTemporal1 <> "" Then Text2.Text = vTemporal1
falla:
  If Err.Number <> 0 Then
     MsgBox Err.Number & "::" & Err.Description, vbCritical
  End If
End Sub

Private Sub bFotos_Click()
  Load fDir2
  vTemporal1 = ""
  
  If Text6.Text = "" Then
    fDir2.Drive1.Drive = Mid(App.Path, 1, 1)
    fDir2.Dir1.Path = GetPath(App.Path & "\", "\")
  Else
    fDir2.Drive1.Drive = Mid(Text6.Text, 1, 1)
    fDir2.Dir1.Path = GetPath(Text6.Text & "\", "\")
  End If
    
  If Text6.Text <> "" Then
    vTemporal1 = Text6.Text
    fDir2.Text1.Text = Text6.Text
  End If
  fDir2.Show vbModal
  If vTemporal1 <> "" Then Text6.Text = vTemporal1
End Sub

Private Sub bOrigen_Click()
  Load fDir
  vTemporal1 = ""
  
  If Text1.Text = "" Then
    fDir.Drive1.Drive = Mid(App.Path, 1, 1)
    fDir.Dir1.Path = GetPath(App.Path & "\", "\")
  Else
    fDir.Drive1.Drive = Mid(Text1.Text, 1, 1)
    fDir.Dir1.Path = GetPath(Text1.Text & "\", "\")
  End If
    
  If Text1.Text <> "" Then
    vTemporal1 = Text1.Text
    fDir.Text1.Text = Text1.Text
  End If
  fDir.Show vbModal
  If vTemporal1 <> "" Then Text1.Text = vTemporal1
End Sub

Private Function BuscarProducto(sCodigo As String) As String
  Dim r As New ADODB.Recordset
  Dim s As String, s1 As String
  s1 = ""
  s = "select * from productos where codigo = '" & sCodigo & "'"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If Not r.EOF Then s1 = Trim(r.Fields("descripcion").value)
  r.Close
  Set r = Nothing
  BuscarProducto = s1
End Function

Private Sub cmdVerificarEmail_Click()
   
    Load fMensaje
    fMensaje.Label1.Caption = "Ejecutando prueba de envío de correo electrónico..."
    fMensaje.Label1.Refresh
    fMensaje.Show
    DoEvents
    Set oMail = New clsCDOmail
    With oMail
         'datos para enviar
        .servidor = "mail.cantv.net"
        .puerto = 25
        '.UseAuntentificacion = True
        '.ssl = True
        .Usuario = txtUsuario.Text
        .PassWord = txtContraseña.Text
        
        .Asunto = "Prueba de correo. " & Text3.Text
        .Adjunto = ""
        .de = txtConCopia.Text
        .para = txtConCopia.Text
        .Mensaje = "PRUEBA DE CORREO. " & Text4.Text
        
        .Enviar_Backup ' manda el mail
    
    End With
    
    Set oMail = Nothing
  
  'If MandaMail(Trim(eEmail.Text), "", "", Trim(eTitulo.Text), Trim(eMensaje.Text), Trim(eAdjunto.Text)) = True Then
  '  Unload Me
  'End If
   Unload fMensaje



End Sub

' envio completo
Private Sub oMail_EnvioCompleto()
    MsgBox "Mensaje enviado", vbInformation
End Sub
' error al enviar
Private Sub oMail_Error(Descripcion As String, Numero As Variant)
    MsgBox Descripcion, vbCritical, Numero
End Sub

Private Sub eCA_Change()
  eCA.Text = UCase(eCA.Text)
  Label8.Caption = BuscarProducto(eCA.Text)
  SendKeys "{END}"
End Sub

Private Sub eCodigo_Change()
  eCodigo.Text = UCase(eCodigo.Text)
  Label6.Caption = BuscarProducto(eCodigo.Text)
  SendKeys "{END}"
End Sub

Private Sub eSF_Change()
  eSF.Text = UCase(eSF.Text)
  Label11.Caption = BuscarProducto(eSF.Text)
  SendKeys "{END}"
End Sub


'Public Sub CargarOpcionesCorreo(ByRef s1 As String, ByRef s2 As String)
'  Dim s As String
'  Dim r As New ADODB.Recordset
'  s1 = ""
'  s2 = ""
'  s = "select * from opciones"
'  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
'  If Not r.EOF Then
'    If Not IsNull(r.Fields("titulomensajecorreo").Value) Then s1 = Trim(r.Fields("titulomensajecorreo").Value)
'    If Not IsNull(r.Fields("cuerpomensajecorreo").Value) Then s2 = Trim(r.Fields("cuerpomensajecorreo").Value)
'  End If
'  r.Close
'  Set r = Nothing
'End Sub

Private Sub Form_Load()
  Dim s1 As String, s2 As String, s3 As String, s4 As String
  Dim s5 As String, s6 As String, s7 As String
  Dim s8 As String, s9 As String, s10 As String
  Dim s11 As String
  Text1.Text = ""
  Text2.Text = ""
  
  s1 = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
  s2 = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
  s4 = GetSetting(APPNAME, "Opciones", "Estacion", "")
  
  s5 = GetSetting(APPNAME, "Opciones", "TituloMensajeCorreo", "")
  s6 = GetSetting(APPNAME, "Opciones", "CuerpoMensajeCorreo", "")
  s7 = GetSetting(APPNAME, "Opciones", "RutaCard5", "")
  s8 = GetSetting(APPNAME, "Opciones", "RutaFotos", "")
  s9 = GetSetting(APPNAME, "Opciones", "CodigoCarnet", "")
  s10 = GetSetting(APPNAME, "Opciones", "CodigoDiseno", "")
  s11 = GetSetting(APPNAME, "Opciones", "CodigoServicioFoto", "")
  txtUsuario.Text = GetSetting(APPNAME, "Opciones", "UsuarioEmail", "")
  txtContraseña.Text = GetSetting(APPNAME, "Opciones", "ContraseñaEmail", "")
  txtConCopia.Text = GetSetting(APPNAME, "Opciones", "ConCopiaEmail", "")
  
  ePC.Text = s4
  
  Text1.Text = s1
  Text2.Text = s2
  
  Text3.Text = ""
  Text4.Text = ""
  eCodigo.Text = ""
  s1 = "": s2 = ""
  'CargarOpcionesCorreo s1, s2, s3
  
  Text3.Text = s1
  Text4.Text = s2
  eCodigo.Text = s3
  Text5.Text = s7
  Text6.Text = s8
  
  eCodigo.Text = s9
  eCA.Text = s10
  
  Text3.Text = s5
  Text4.Text = s6
  
  eSF.Text = s11
  
End Sub



