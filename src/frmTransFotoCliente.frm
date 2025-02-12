VERSION 5.00
Begin VB.Form frmTransFotoCliente 
   Caption         =   "frmTransFotoCliente"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   255
      Left            =   5460
      TabIndex        =   15
      Top             =   2880
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   255
      Left            =   5460
      TabIndex        =   14
      Top             =   1320
      Width           =   795
   End
   Begin VB.ListBox ListClientes 
      Height          =   2595
      Left            =   240
      OLEDragMode     =   1  'Automatic
      TabIndex        =   10
      Top             =   1620
      Width           =   2175
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   4500
      Width           =   1275
   End
   Begin VB.CommandButton cmdTransferir 
      Caption         =   "Transferir"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   4500
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "Imagen para Selección"
      Height          =   2475
      Left            =   3060
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
      Begin VB.Image Image1 
         Height          =   2175
         Left            =   180
         OLEDragMode     =   1  'Automatic
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3315
      Left            =   6480
      TabIndex        =   5
      Top             =   840
      Width           =   2175
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   11
         Top             =   420
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Transferir al Cliente"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   180
         Width           =   1755
      End
      Begin VB.Label Label3 
         Caption         =   "Imagen a Transferir"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   1695
      End
      Begin VB.Image Image2 
         Height          =   2175
         Left            =   180
         OLEDropMode     =   1  'Manual
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.TextBox txtNombre 
      Height          =   255
      Left            =   1020
      TabIndex        =   4
      Top             =   780
      Width           =   2595
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "OK"
      Height          =   435
      Left            =   2820
      TabIndex        =   2
      Top             =   180
      Width           =   495
   End
   Begin VB.TextBox txtCedula 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   300
      Width           =   1755
   End
   Begin VB.Label Label2 
      Caption         =   "Transferir desde:"
      Height          =   255
      Left            =   300
      TabIndex        =   9
      Top             =   1260
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   300
      TabIndex        =   3
      Top             =   780
      Width           =   675
   End
   Begin VB.Label lblCedula 
      Caption         =   "Cédula:"
      Height          =   255
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "frmTransFotoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lOrigen As String
Dim lDestino As String
Dim lCedula As String
Private Sub cmdBuscar_Click()
   Dim lReg As New ADODB.Recordset
   Dim lReg2 As New ADODB.Recordset
   Dim lCn As New ADODB.Connection
   Dim TablaActual As String
   Dim Encontrado As Boolean
   Encontrado = False
   lReg.Open "Select p.Id,p.cliente,p.Subcliente,p.tabla,c.Nombre as NombreCliente,s.Nombre as nombreSubCliente from Personas p inner join clientes c on c.Codigo=p.cliente inner join Subclientes s on s.id=p.subcliente or p.cliente>=0", Modulo.DBConexionSQL, adOpenStatic
   lCedula = txtCedula.Text
   txtCedula.Text = fFormatearCedula(txtCedula.Text)
   ListClientes.Clear
   TablaActual = lReg!tabla
   Do While lReg.EOF = False
      
      lReg2.Open "Select * from [" & Trim(lReg!tabla) & "] Where Cedula='" & txtCedula.Text & "'", Modulo.DBConexionSQL, adOpenKeyset
      If lReg2.EOF = False Then
         txtNombre.Text = Trim(lReg2!Nombre)
         Encontrado = True
         If lReg!SubCliente = 0 Then
            ListClientes.AddItem Trim(lReg!NombreCliente)
            'ListClientes.AddItem Trim(lReg!NombreCliente)
         Else
            ListClientes.AddItem Trim(lReg!NombreSubCliente)
            'ListClientes.AddItem Trim(lReg!NombreSubCliente)
         End If
      End If
      lReg2.Close
      lReg.MoveNext
      If lReg.EOF = True Then Exit Do
      If TablaActual = lReg!tabla Then
        Do While lReg.EOF = False
           If TablaActual <> lReg!tabla Then Exit Do
           lReg.MoveNext
        Loop
        If lReg.EOF = False Then TablaActual = lReg!tabla
      Else
        TablaActual = lReg!tabla
      End If
   Loop
   
End Sub


Private Function fFormatearCedula(argCedula) As String
  Dim s As String
  s = Trim(argCedula)
  If s <> "" Then
    Formatear_Cedula s
    fFormatearCedula = s
  End If
End Function

Private Sub cmdCerrar_Click()
  Unload Me
End Sub

Private Sub cmdTransferir_Click()
  Dim lDst As String
  Dim lOrig As String
  If MsgBox("¿Esta Seguro?", vbYesNo + vbQuestion) = vbYes Then
        'lOrig = lDestino & "\" & Trim(ListClientes.List(ListClientes.ListIndex)) & "\" & "FOTOS"
        lOrig = Image2.Tag
        lDst = lDestino & "\" & Trim(txtCliente.Text) & "\" & "FOTOS\" & Replace(lCedula, ".", "") & ".jpg"
        FileCopy lOrig, lDst
        MsgBox "Transferencia realizada", vbInformation
  End If
End Sub

Private Sub Command1_Click()
   txtCliente.Text = ListClientes.List(ListClientes.ListIndex)
End Sub

Private Sub Command2_Click()
 Image2.Picture = Image1.Picture
End Sub

Private Sub Form_Load()
   
   lOrigen = GetSetting(APPNAME, "Opciones", "RutaFotos", "")
   lDestino = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
   If Len(lDestino) = 3 Then lDestino = Mid(lDestino, 1, 2)
End Sub

Private Sub Image2_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
  'MsgBox "dragover2"
  Image2.Picture = Image1.Picture
  Image2.Tag = Image1.Tag
End Sub

Private Sub ListClientes_Click()
   Dim lDst As String
   On Error GoTo falla
   
   lDst = lDestino & "\" & Trim(ListClientes.List(ListClientes.ListIndex)) & "\" & "FOTOS\" & Replace(lCedula, ".", "") & ".jpg"
   Image1.Picture = LoadPicture(lDst)
   Image1.Tag = lDst
falla:
   If Err.Number <> 0 Then Image1.Picture = LoadPicture()
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 
End Sub

Private Sub List1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
 
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
   If KeyAscii = 10 Or KeyAscii = 13 Then
      cmdBuscar_Click
   End If
End Sub

Private Sub txtCliente_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
   txtCliente.Text = ListClientes.List(ListClientes.ListIndex)
End Sub
