VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmSelCarnets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccione Los Carnets"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Seleccionar todos"
      Height          =   195
      Left            =   420
      TabIndex        =   4
      Top             =   120
      Width           =   1755
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Marcar Transferido"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      ToolTipText     =   "Limpiar lista para fotos ya existentes en el cliente"
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2580
      TabIndex        =   1
      Top             =   3060
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListvCarnets 
      Height          =   2415
      Left            =   300
      TabIndex        =   0
      Top             =   360
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
End
Attribute VB_Name = "FrmSelCarnets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
  Dim i As Integer
 For i = 1 To ListvCarnets.ListItems.Count
    ListvCarnets.ListItems(i).Checked = Check1.value
 Next i
  
End Sub

Private Sub Command1_Click()
 Dim i As Integer
 Dim lSeleccion As Boolean
 lSeleccion = False
 Load frmTransferirFotos
 frmTransferirFotos.sCargarLista
 For i = 1 To ListvCarnets.ListItems.Count
    If ListvCarnets.ListItems(i).Checked = True Then
       lSeleccion = True
       Exit For
    End If
 Next
 If lSeleccion = False Then
    MsgBox "Debe Seleccionar al menos un Registro", vbExclamation
    Exit Sub
 End If
 
 If frmTransferirFotos.ListvCarnets.ListItems.Count <> frmTransferirFotos.Fotos.ListCount Then
    MsgBox "El Número de Registros seleccionados y el número de fotos Tomadas debe ser igual(" & frmTransferirFotos.ListvCarnets.ListItems.Count & " <> " & frmTransferirFotos.Fotos.ListCount & ")", vbExclamation
    Exit Sub
 End If
 frmTransferirFotos.Show vbModal
 Unload Me
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub
Private Sub sCargarDiario()
   Dim i As Integer
   Dim lItem As ListItem
   Dim lReg As New ADODB.Recordset
   ''lReg.Close
   lReg.Open "Select id,Cedula,transferido,tabla from diario where fecha='" & Format(Now, "yyyy/MM/dd") & "' order by ID", Modulo.DBConexionSQL, adOpenKeyset
     If lReg.EOF = True Then Exit Sub
      With fMov.GRID
    
       For i = 1 To .Rows
          lReg.MoveFirst
          lReg.Find "ID=" & .TextMatrix(i, 1)
          If lReg.EOF = False Then
            If Val(lReg!Transferido) = 0 And Trim(lReg!cedula) = .TextMatrix(i, 5) Then
             Set lItem = ListvCarnets.ListItems.Add(, , .TextMatrix(i, 1))
              lItem.SubItems(1) = .TextMatrix(i, 3)
              lItem.SubItems(2) = .TextMatrix(i, 4)
              lItem.SubItems(3) = .TextMatrix(i, 5)
              lItem.SubItems(4) = .TextMatrix(i, 6)
              lItem.SubItems(5) = .TextMatrix(i, 7)
              lItem.SubItems(6) = Trim(lReg!tabla)
           End If
          End If
       Next i
      End With
   If ListvCarnets.ListItems.Count <= 0 Then
      MsgBox "Todos los carnets elaborados hasta el momento ya fueron transferidos", vbInformation

   End If
End Sub

Private Sub sCargarDiario1()
   Dim lReg As New ADODB.Recordset
   Dim lReg2 As New ADODB.Recordset
   Dim lCn As New ADODB.Connection
   Dim lItem As ListItem
   lCn.ConnectionString = Modulo.DBConexionSQL
   lCn.Open
   ListvCarnets.ListItems.Clear
   Set lReg = lCn.Execute("Select ID,Cedula,Tabla from Diario where Fecha=getdate()")
   Do While lReg.EOF = False
      Set lReg2 = lCn.Execute("Select * From [" & lReg!tabla & "] Where Cedula='" & lReg!cedula & "'")
         If lReg2.EOF = False Then
            Set lItem = ListvCarnets.ListItems.Add(, , lReg!ID)
               lItem.SubItems(1) = lReg!cedula
               lItem.SubItems(2) = lReg!Nombre
               lItem.SubItems(3) = lReg!cargo
               
         End If
      lReg.MoveNext
   Loop

End Sub

Private Sub Command3_Click()
Dim lReg As New ADODB.Recordset
Dim lCn As New ADODB.Connection
 If MsgBox("¿Seguro que desea limpiar el listado de fotos por transferir?", vbYesNo + vbQuestion) = vbYes Then
   lCn.ConnectionString = Modulo.DBConexionSQL
   lCn.Open
   For i = 1 To ListvCarnets.ListItems.Count
      If ListvCarnets.ListItems(i).Checked = True Then
         lCn.Execute ("update Diario set transferido =1 where cedula='" & ListvCarnets.ListItems(i).ListSubItems(3) & "' and ID=" & ListvCarnets.ListItems(i).Text)
         ListvCarnets.ListItems.Remove i
      End If
   Next i
   
 End If
End Sub

Private Sub Form_Load()
   sCargarDiario
End Sub
