VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCopiarCarpetas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar Carpetas Principales"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13590
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   13590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd2Historico 
      Caption         =   ">>"
      Height          =   255
      Left            =   6960
      TabIndex        =   24
      Top             =   3780
      Width           =   435
   End
   Begin VB.CommandButton cmdAtrasHistorico 
      Caption         =   "<<"
      Height          =   255
      Left            =   6300
      TabIndex        =   23
      Top             =   3780
      Width           =   435
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   7620
      TabIndex        =   19
      Top             =   2700
      Width           =   5835
      Begin VB.TextBox txtNuevaCarpeta 
         Height          =   315
         Left            =   2400
         TabIndex        =   21
         Top             =   240
         Width           =   3255
      End
      Begin VB.ListBox ListHistorico 
         Height          =   1230
         Left            =   180
         OLEDropMode     =   1  'Manual
         TabIndex        =   20
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label Label4 
         Caption         =   "Carpeta Histórico:"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   300
         Width           =   1635
      End
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   4140
      TabIndex        =   18
      Top             =   5520
      Width           =   1515
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1395
      Left            =   840
      TabIndex        =   17
      Top             =   5520
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2461
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   555
      Left            =   7080
      TabIndex        =   16
      Top             =   4860
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   555
      Left            =   4980
      TabIndex        =   15
      Top             =   4860
      Width           =   1575
   End
   Begin VB.CommandButton cmdAtrasImagenes 
      Caption         =   "<<"
      Height          =   255
      Left            =   6300
      TabIndex        =   13
      Top             =   2400
      Width           =   435
   End
   Begin VB.CommandButton cmdAtrasFotos 
      Caption         =   "<<"
      Height          =   255
      Left            =   6300
      TabIndex        =   12
      Top             =   1560
      Width           =   435
   End
   Begin VB.CommandButton cmdAtrasCarnet 
      Caption         =   "<<"
      Height          =   255
      Left            =   6300
      TabIndex        =   11
      Top             =   720
      Width           =   435
   End
   Begin VB.CommandButton cmd2Imagenes 
      Caption         =   ">>"
      Height          =   255
      Left            =   6960
      TabIndex        =   10
      Top             =   2400
      Width           =   435
   End
   Begin VB.CommandButton cmd2Fotos 
      Caption         =   ">>"
      Height          =   255
      Left            =   6960
      TabIndex        =   9
      Top             =   1560
      Width           =   435
   End
   Begin VB.CommandButton cmd2Carnet 
      Caption         =   ">>"
      Height          =   255
      Left            =   6960
      TabIndex        =   8
      Top             =   720
      Width           =   435
   End
   Begin VB.DirListBox ListCarpetas 
      Height          =   4140
      Left            =   240
      TabIndex        =   7
      Top             =   540
      Width           =   5775
   End
   Begin VB.TextBox txtCarpetaImagenes 
      Height          =   315
      Left            =   7620
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   2340
      Width           =   5715
   End
   Begin VB.TextBox txtCarpetaFotos 
      Height          =   315
      Left            =   7620
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   1500
      Width           =   5715
   End
   Begin VB.TextBox txtCarpetaCarnet 
      Height          =   315
      Left            =   7620
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   660
      Width           =   5715
   End
   Begin VB.ListBox ListCarpetas1 
      Height          =   840
      Left            =   240
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   540
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Label lbltitulo 
      Caption         =   "LISTA DE CARPETAS"
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
      Left            =   300
      TabIndex        =   14
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label Label3 
      Caption         =   "CARPETA IMAGENES"
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
      Left            =   7620
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "CARPETA FOTOS"
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
      Left            =   7620
      TabIndex        =   5
      Top             =   1200
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "CARPETA CARNET"
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
      Left            =   7620
      TabIndex        =   4
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmCopiarCarpetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd2Carnet_Click()
Dim i As Integer
   If ListCarpetas.ListIndex < 0 Then Exit Sub
   If txtCarpetaCarnet.Text <> "" Then Exit Sub
   For i = 0 To ListHistorico.ListCount - 1
       If ListHistorico.List(i) = ListCarpetas.List(ListCarpetas.ListIndex) Then Exit Sub
       If ListCarpetas.List(ListCarpetas.ListIndex) = txtCarpetaImagenes.Text Then Exit Sub
       If ListCarpetas.List(ListCarpetas.ListIndex) = txtCarpetaFotos.Text Then Exit Sub
   Next i
   
   txtCarpetaCarnet.Text = ListCarpetas.List(ListCarpetas.ListIndex)
   txtCarpetaCarnet.Tag = ListCarpetas.ListIndex
   'ListCarpetas.RemoveItem ListCarpetas.ListIndex
End Sub

Private Sub cmd2Fotos_Click()
Dim i As Integer
   If ListCarpetas.ListIndex < 0 Then Exit Sub
   If txtCarpetaFotos.Text <> "" Then Exit Sub
   For i = 0 To ListHistorico.ListCount - 1
       If ListHistorico.List(i) = ListCarpetas.List(ListCarpetas.ListIndex) Then Exit Sub
       If ListCarpetas.List(ListCarpetas.ListIndex) = txtCarpetaImagenes.Text Then Exit Sub
       If ListCarpetas.List(ListCarpetas.ListIndex) = txtCarpetaCarnet.Text Then Exit Sub
   Next i
   txtCarpetaFotos.Text = ListCarpetas.List(ListCarpetas.ListIndex)
   txtCarpetaFotos.Tag = ListCarpetas.ListIndex
   'ListCarpetas.RemoveItem ListCarpetas.ListIndex
End Sub

Private Sub cmd2Historico_Click()
   Dim i As Integer
   If ListCarpetas.ListIndex < 0 Then Exit Sub
   For i = 0 To ListHistorico.ListCount - 1
       If ListHistorico.List(i) = ListCarpetas.List(ListCarpetas.ListIndex) Then Exit Sub
       If ListCarpetas.List(ListCarpetas.ListIndex) = txtCarpetaCarnet.Text Then Exit Sub
       If ListCarpetas.List(ListCarpetas.ListIndex) = txtCarpetaFotos.Text Then Exit Sub
       If ListCarpetas.List(ListCarpetas.ListIndex) = txtCarpetaImagenes.Text Then Exit Sub
   Next i
   If ListCarpetas.List(ListCarpetas.ListIndex) = txtCarpetaCarnet.Text Then Exit Sub
   If ListCarpetas.List(ListCarpetas.ListIndex) = txtCarpetaFotos.Text Then Exit Sub
   If ListCarpetas.List(ListCarpetas.ListIndex) = txtCarpetaImagenes.Text Then Exit Sub

   ListHistorico.AddItem ListCarpetas.List(ListCarpetas.ListIndex)
End Sub

Private Sub cmd2Imagenes_Click()
 Dim i As Integer
   If ListCarpetas.ListIndex < 0 Then Exit Sub
   If txtCarpetaImagenes.Text <> "" Then Exit Sub
   For i = 0 To ListHistorico.ListCount - 1
       If ListHistorico.List(i) = ListCarpetas.List(ListCarpetas.ListIndex) Then Exit Sub
       If ListCarpetas.List(ListCarpetas.ListIndex) = txtCarpetaFotos.Text Then Exit Sub
       If ListCarpetas.List(ListCarpetas.ListIndex) = txtCarpetaCarnet.Text Then Exit Sub
   Next i
   txtCarpetaImagenes.Text = ListCarpetas.List(ListCarpetas.ListIndex)
   txtCarpetaImagenes.Tag = ListCarpetas.ListIndex
   'ListCarpetas.RemoveItem ListCarpetas.ListIndex
End Sub

Private Sub cmdAceptar_Click()
 If txtCarpetaCarnet.Text = "" Or txtCarpetaFotos.Text = "" Or txtCarpetaImagenes.Text = "" Then
    MsgBox "No ha completado la selección de las carpetas Principales", vbExclamation
    Exit Sub
 End If
 If ListHistorico.ListCount <= 0 Then
    If MsgBox("La lista de carpetas de histórico esta vacía.¿Desea Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
 End If
 
   With fCargaClientes
      .lDirCarnet = txtCarpetaCarnet.Text
      .lDirFotos = txtCarpetaFotos.Text
      .lDirImagenes = txtCarpetaImagenes.Text
      
      Load fMensaje
      fMensaje.Label1.Caption = "Copiando Carpetas/Plantillas del Cliente " & fCargaClientes.eCliente.Text & " , Espere..."
      fMensaje.Show
      DoEvents

      Dim sOri As String
      Dim sDes As String
    
      Dim s2 As String
    
      sOri = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
      sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")

    
      If sOri <> "" And sDes <> "" Then
        .sCopiarCarpetas ListCarpetas, File1, "", .Dir2, .File2
        Unload fMensaje
        MsgBox "Copia de archivos realizada", vbInformation
      Else
        MsgBox "No se puede determinar el destino a copiar. La carpeta del cliente no se copiará. Vaya a opciones y configure la carpeta de destino", vbExclamation
      End If
      Unload Me
   End With
   
End Sub

Private Sub cmdAtrasCarnet_Click()

  If txtCarpetaCarnet.Text = "" Then Exit Sub
'  ListCarpetas.AddItem txtCarpetaCarnet.Text, txtCarpetaCarnet.Tag
  txtCarpetaCarnet.Text = ""
  txtCarpetaCarnet.Tag = ""
End Sub

Private Sub cmdAtrasFotos_Click()
If txtCarpetaFotos.Text = "" Then Exit Sub
  'ListCarpetas.AddItem txtCarpetaFotos.Text, txtCarpetaFotos.Tag
  txtCarpetaFotos.Text = ""
  txtCarpetaFotos.Tag = ""

End Sub

Private Sub cmdAtrasHistorico_Click()
If ListHistorico.ListIndex < 0 Then Exit Sub
ListHistorico.RemoveItem ListHistorico.ListIndex
End Sub

Private Sub cmdAtrasImagenes_Click()
  If txtCarpetaImagenes.Text = "" Then Exit Sub
  'ListCarpetas.AddItem txtCarpetaImagenes.Text, txtCarpetaImagenes.Tag
  txtCarpetaImagenes.Text = ""
  txtCarpetaImagenes.Tag = ""

End Sub

Public Sub sCargarListCarpetas()
   Dim i As Integer
   ListCarpetas1.Clear
   For i = 0 To ListCarpetas.ListCount
      ListCarpetas1.AddItem ListCarpetas.List(i)
   Next i
End Sub


Private Sub Command1_Click()
  Unload Me
End Sub


Private Sub Form_Load()
   txtNuevaCarpeta.Text = Year(Date)
End Sub

Private Sub ListHistorico_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
   cmd2Historico_Click
End Sub

Private Sub txtCarpetaCarnet_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
   cmd2Carnet_Click
End Sub

Private Sub txtCarpetaFotos_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
   cmd2Fotos_Click
End Sub

Private Sub txtCarpetaImagenes_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
   cmd2Imagenes_Click
End Sub

Private Sub txtNuevaCarpeta_GotFocus()
   txtNuevaCarpeta.SelStart = 0
   txtNuevaCarpeta.SelLength = Len(txtNuevaCarpeta.Text)
   txtNuevaCarpeta.SetFocus
End Sub
