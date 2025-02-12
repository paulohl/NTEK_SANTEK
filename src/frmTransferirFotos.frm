VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTransferirFotos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferencia de Fotos"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   14640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBajar 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   8
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton cmdSubir 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   7
      Top             =   780
      Width           =   375
   End
   Begin VB.ListBox ListFotos 
      Height          =   2205
      Left            =   10320
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.FileListBox ExisteArchivo 
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   5940
      TabIndex        =   3
      Top             =   3120
      Width           =   1635
   End
   Begin VB.CommandButton cmdTransferir 
      Caption         =   "Transferir"
      Height          =   495
      Left            =   3300
      TabIndex        =   2
      Top             =   3120
      Width           =   1635
   End
   Begin VB.ListBox ListCedulas 
      Height          =   2010
      Left            =   1440
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
   Begin VB.FileListBox Fotos 
      Height          =   2625
      Left            =   7020
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   4080
      Width           =   1695
   End
   Begin MSComctlLib.ListView ListvCarnets 
      Height          =   2535
      Left            =   180
      TabIndex        =   5
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cliente"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "SubCliente"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cedula"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nombre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Apellido"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tabla"
         Object.Width           =   2
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   12660
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmTransferirFotos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lOrigen As String
Dim lDestino As String
Dim lNombreCliente As String
Dim lNombreSubCliente As String
Dim lSubCliente As String
Dim lCedula As String

Sub sCargarListaCLientesSel()
   Dim i As Integer
   With fMov
           
    'MsgBox Mid(.GRID.TextMatrix(.GRID.Row, 3), 7, Len(.GRID.TextMatrix(.GRID.Row, 3)) - 6)
    lNombreCliente = Mid(.GRID.TextMatrix(.GRID.Row, 3), 7, Len(.GRID.TextMatrix(.GRID.Row, 3)) - 6)
    If Len(.GRID.TextMatrix(.GRID.Row, 4)) > 1 Then lSubCliente = Mid(.GRID.TextMatrix(.GRID.Row, 4), 7, Len(.GRID.TextMatrix(.GRID.Row, 4)) - 6)
    lCedula = .GRID.TextMatrix(.GRID.Row, 5)
    lCedula = Replace(lCedula, ".", "")
       
       If lSubCliente = "" Then
             
       Else
          
       End If
   End With
   
End Sub


Private Sub cmdBajar_Click()
   Dim lAux As String
   If ListFotos.ListIndex < ListFotos.ListCount - 1 Then
      lAux = ListFotos.List(ListFotos.ListIndex + 1)
      ListFotos.List(ListFotos.ListIndex + 1) = ListFotos.List(ListFotos.ListIndex)
      ListFotos.List(ListFotos.ListIndex) = lAux
      ListFotos.ListIndex = ListFotos.ListIndex + 1
   End If
   'MsgBox ListFotos.ListIndex & ":::" & ListFotos.ListCount
End Sub

Private Sub cmdSubir_Click()
   Dim lAux As String
   If ListFotos.ListIndex > 0 Then
      lAux = ListFotos.List(ListFotos.ListIndex - 1)
      ListFotos.List(ListFotos.ListIndex - 1) = ListFotos.List(ListFotos.ListIndex)
      ListFotos.List(ListFotos.ListIndex) = lAux
      ListFotos.ListIndex = ListFotos.ListIndex - 1
   End If
   
End Sub

Private Sub Command1_Click()
   'If MsgBox("¿Esta seguro de hacer la transferencia?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
   Dim i As Integer
   For i = 1 To ExisteArchivo.ListCount
      If UCase(ExisteArchivo.List(i)) = lCedula & ".JPG" Then
        If MsgBox("Ya Existe una foto para esta cedula. Desea Sobreescribirla?", vbYesNo + vbQuestion) = vbNo Then
           Exit Sub
        Else
           Exit For
        End If
      End If
   Next i
   FileCopy lOrigen & "\" & Fotos.List(Fotos.ListIndex), lDestino & "\" & lCedula & ".jpg"
   Kill lOrigen & "\" & Fotos.List(Fotos.ListIndex)
   Fotos.Refresh
   Unload Me
End Sub

Private Sub cmdTransferir_Click()
  Dim i As Integer
  Dim lDst As String 'destino
  Dim lContinuar As Boolean
  On Error GoTo falla
  For i = 1 To ListvCarnets.ListItems.Count
     lNombreCliente = ListvCarnets.ListItems(i).SubItems(1)
     lNombreCliente = Trim(Mid(lNombreCliente, 7, Len(lNombreCliente) - 6))
     lNombreSubCliente = ListvCarnets.ListItems(i).SubItems(2)
     If Len(lNombreSubCliente) > 2 Then lNombreSubCliente = Trim(Mid(lNombreSubCliente, 7, Len(lNombreSubCliente) - 6))
     If Len(lDestino) = 3 Then
        lDestino = Mid(lDestino, 1, 2)
     End If
     If Len(lNombreSubCliente) > 2 Then
        lDst = lDestino & "\" & Trim(lNombreCliente) & "\" & lNombreSubCliente & "\" & "FOTOS"
     Else
        lDst = lDestino & "\" & Trim(lNombreCliente) & "\" & "FOTOS"
     End If
     'MsgBox lDst
     ExisteArchivo.Path = lDst
     ExisteArchivo.Refresh
     lCedula = ListvCarnets.ListItems(i).SubItems(3)
     lCedula = Replace(lCedula, ".", "")
     lContinuar = True
     For j = 0 To ExisteArchivo.ListCount - 1
      If UCase(ExisteArchivo.List(j)) = lCedula & ".JPG" Then
        If MsgBox("Ya Existe una foto para esta cedula(" & lCedula & " " & ListvCarnets.ListItems(i).SubItems(4) & "). Desea Sobreescribirla?", vbYesNo + vbQuestion) = vbYes Then
           lContinuar = True
           Exit For
        Else
           lContinuar = False
           Exit For
        End If
      End If
     Next j
     If lContinuar = True Then
        FileCopy lOrigen & "\" & ListFotos.List(i - 1), lDst & "\" & lCedula & ".jpg"
        'Kill lOrigen & "\" & Fotos.List(i - 1)
        'Fotos.Refresh
        
     End If
     Marcar_Personas_CEDULA ListvCarnets.ListItems(i).SubItems(6), lCedula
     Modulo.ExecSQL "Update Diario set transferido=1 where ID=" & ListvCarnets.ListItems(i).Text
  Next i
  MsgBox "Transferencia de Fotos Realizada", vbInformation

  For i = 1 To ListvCarnets.ListItems.Count
     Kill lOrigen & "\" & Fotos.List(i - 1)
  Next i
  
  ListvCarnets.ListItems.Clear
  ListFotos.Clear
  Unload Me
falla:
  If Err.Number <> 0 Then
     MsgBox Err.Number & "::" & Err.Description, vbCritical
  End If
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Public Sub sCargarLista()
   Dim i As Integer
   Dim lItem As ListItem
   
   lOrigen = GetSetting(APPNAME, "Opciones", "RutaFotos", "")
   lDestino = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
   sCargarListaCLientesSel
   'lDestino = lDestino & "\" & Trim(lNombreCliente) & "\" & "FOTOS"
   Fotos.Path = lOrigen
   Fotos.Refresh
   ExisteArchivo.Path = lDestino
   ListCedulas.AddItem lCedula
   
   ListvCarnets.ListItems.Clear
   With FrmSelCarnets.ListvCarnets
      For i = 1 To .ListItems.Count
         If .ListItems(i).Checked = True Then
            Set lItem = ListvCarnets.ListItems.Add(, , .ListItems(i).Text)
               lItem.SubItems(1) = .ListItems(i).SubItems(1)
               lItem.SubItems(2) = .ListItems(i).SubItems(2)
               lItem.SubItems(3) = .ListItems(i).SubItems(3)
               lItem.SubItems(4) = .ListItems(i).SubItems(4)
               lItem.SubItems(5) = .ListItems(i).SubItems(5)
               lItem.SubItems(6) = .ListItems(i).SubItems(6)
         End If
      Next i
   End With
   ListFotos.Clear
   For i = 0 To Fotos.ListCount - 1
      ListFotos.AddItem Fotos.List(i)
   Next i
End Sub

Private Sub Form_Load()
  
  sCargarLista
  's = GetSetting(APPNAME, "Opciones", "RutaFotos", "")
  
End Sub

Private Sub Fotos_Click()
  If Fotos.ListCount < 0 Then Exit Sub
  Image1.Picture = LoadPicture(lOrigen & "\" & Fotos.List(Fotos.ListIndex))
End Sub

Private Sub Label2_Click()
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   List1.AddItem ListCedulas.List(ListCedulas.ListIndex)
End Sub

Private Sub ListCedulas_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim aux As String
'   aux = ListCedulas.List(ListCedulas.ListIndex)
'   ListCedulas.List (ListCedulas.ListIndex)
 
End Sub

Private Sub ListCedulas_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
 'MsgBox ListCedulas.List(ListCedulas.ListIndex)
 
End Sub

Private Sub ListFotos_Click()
  If ListFotos.ListCount < 0 Then Exit Sub
  Image1.Picture = LoadPicture(lOrigen & "\" & ListFotos.List(ListFotos.ListIndex))

End Sub

Private Sub ListFotos_KeyDown(KeyCode As Integer, Shift As Integer)
   ListFotos_Click
End Sub

Private Sub ListFotos_KeyUp(KeyCode As Integer, Shift As Integer)
   ListFotos_Click
End Sub

Private Sub ListvCarnets_Click()
   ListFotos.ListIndex = ListvCarnets.SelectedItem.Index - 1
   'Fotos.Selected = True
End Sub

Private Sub ListvCarnets_KeyDown(KeyCode As Integer, Shift As Integer)
   ListFotos.ListIndex = ListvCarnets.SelectedItem.Index - 1

End Sub

Private Sub ListvCarnets_KeyUp(KeyCode As Integer, Shift As Integer)
   ListFotos.ListIndex = ListvCarnets.SelectedItem.Index - 1

End Sub
