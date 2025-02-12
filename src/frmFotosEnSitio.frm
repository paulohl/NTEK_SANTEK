VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmFotosEnSitio 
   Caption         =   "Fotos en Sitio"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   LinkTopic       =   "Form2"
   ScaleHeight     =   8325
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Programacion\SANTEK\FotosEnSitio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   7740
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   540
      Visible         =   0   'False
      Width           =   1515
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   1320
      Top             =   3780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar Archivo Access"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fotos Numeradas Cantidad 10 / 20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   3540
      TabIndex        =   2
      Top             =   1740
      Width           =   5835
      Begin MSComctlLib.ListView ListVFotosNumeradas 
         Height          =   2295
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Doble click si desea quitarlo de la lista"
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   4048
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "NroFoto"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cédula"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listado. Cantidad 20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   9255
      Begin MSComctlLib.ListView ListvResultados 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   5530
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cédula"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cargo"
            Object.Width           =   4410
         EndProperty
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Subcliente:"
      Height          =   315
      Left            =   660
      TabIndex        =   9
      Top             =   900
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   660
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblSubcliente 
      Caption         =   "lblSubCliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      TabIndex        =   7
      Top             =   1260
      Width           =   6315
   End
   Begin VB.Label lblCliente 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   420
      Width           =   6315
   End
End
Attribute VB_Name = "frmFotosEnSitio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lUltimoNumero As Integer
Private Sub Command2_Click()
   Dim txtCriterio As String
   Dim lCodiClie As String
   Dim lCodiSubClie As String
   Dim SqlTxt As String
   Dim lReg As New ADODB.Recordset
   lCodiClie = fPersonasAct.lTablas.List(fPersonasAct.lTablas.ListIndex)
   lCodiSubClie = Mid(lblSubcliente.Caption, 1, 6)
   txtCriterio = InputBox("Introduzca el número de cédula:")
   Formatear_Cedula txtCriterio
   If txtCriterio <> "" Then
      SqlTxt = "Select * from [" & lCodiClie & "] where cedula='" & txtCriterio & "'"
      Set lReg = DBConexionSQL.Execute(SqlTxt)
      If lReg.EOF = False Then
         frmNroFotos.Show
         frmNroFotos.lblNombre.Caption = Trim(lReg!Nombre)
         'frmNroFotos.lblApellido.Caption = Trim(lReg!apellidos)
         'frmNroFotos.txtNroFoto.Text =
      End If
   End If
End Sub


Private Sub sCargarListado()
   Dim lReg As New ADODB.Recordset
   Dim lTabla As String
   Dim lItem As ListItem
   Dim i As Integer
   lTabla = fPersonasAct.lTablas.List(fPersonasAct.lTablas.ListIndex)
   Set lReg = DBConexionSQL.Execute("Select * from [" & lTabla & "] order by nombre")
   ListvResultados.ListItems.Clear
   Do While lReg.EOF = False
      Set lItem = ListvResultados.ListItems.Add(, , lReg!ID)
      lItem.SubItems(1) = Trim(lReg!cedula)
      lItem.SubItems(2) = Trim(lReg!Nombre)
      For i = 0 To lReg.Fields.Count - 1
         If UCase(lReg.Fields(i).Name) = "CARGO" Then
            lItem.SubItems(3) = Trim(lReg!cargo)
         End If
      Next i
      lReg.MoveNext
   Loop
   sUltimoNumero
   Frame1.Caption = "Listado. Cantidad " & ListvResultados.ListItems.Count
   Frame2.Caption = "Fotos Numeradas Cantidad " & ListVFotosNumeradas.ListItems.Count

End Sub



Private Sub cmdBuscar_Click()
   frmBuscarEnSitio.Show
End Sub

Private Sub sCargarFotosNumeradas()
   Dim lReg As New ADODB.Recordset
   Dim lTabla As String
   Dim lItem As ListItem
   Dim i As Integer
   lTabla = fPersonasAct.lTablas.List(fPersonasAct.lTablas.ListIndex)
   Set lReg = DBConexionSQL.Execute("Select * from [" & lTabla & "] order by NroFoto")
   ListVFotosNumeradas.ListItems.Clear
   Do While lReg.EOF = False
      If IsNull(lReg!NROFOTO) = False And lReg!NROFOTO <> "0" Then
         Set lItem = ListVFotosNumeradas.ListItems.Add(, , IIf(IsNull(lReg!NROFOTO), "", lReg!NROFOTO))
         lItem.SubItems(1) = lReg!cedula
         lItem.SubItems(2) = lReg!Nombre
         For i = 1 To ListvResultados.ListItems.Count
            If Trim(lReg!cedula) = ListvResultados.ListItems(i).SubItems(1) Then
               ListvResultados.ListItems.Remove i
               Exit For
            End If
         Next i
      End If
      lReg.MoveNext
   Loop
   Frame1.Caption = "Listado. Cantidad " & ListvResultados.ListItems.Count
   Frame2.Caption = "Fotos Numeradas Cantidad " & ListVFotosNumeradas.ListItems.Count

End Sub



Public Sub sAgregarEnFotosNumeradas(argNum As String, argCedula As String, argNombre As String)
   Dim lItem As ListItem
   Dim i As Integer
      Set lItem = ListVFotosNumeradas.ListItems.Add(, , argNum)
      lItem.SubItems(1) = argCedula
      lItem.SubItems(2) = argNombre
'quitar persona del listado general
   For i = 1 To ListvResultados.ListItems.Count
      If Trim(ListvResultados.ListItems(i).SubItems(1)) = argCedula Then
         ListvResultados.ListItems.Remove i
         Exit For
      End If
   Next i
   Frame1.Caption = "Listado. Cantidad " & ListvResultados.ListItems.Count
   Frame2.Caption = "Fotos Numeradas Cantidad " & ListVFotosNumeradas.ListItems.Count
End Sub

Private Sub Command1_Click()
 Dim NombreTabla As String
 Dim SqlTxt As String
 Dim lTabla As String
 Dim lReg As New ADODB.Recordset
 Dim lCampo As String
 Dim i As Integer
 On Error GoTo falla
 If MsgBox("¿Esta seguro que desea realizar esta operación?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
 If lblSubcliente.Caption = "" Then
    Dialog1.FileName = UCase(Replace(lblCliente.Caption, ":", "_")) & ".MDB"
    NombreTabla = UCase(Replace(lblCliente.Caption, ":", "_"))
    NombreTabla = Replace(NombreTabla, ".", " ")
 Else
    Dialog1.FileName = UCase(Replace(lblCliente.Caption, ":", "_")) & "_" & UCase(Replace(lblSubcliente.Caption, ":", "_")) & ".MDB"
    NombreTabla = UCase(Replace(lblCliente.Caption, ":", "_")) & "_" & UCase(Replace(lblSubcliente.Caption, ":", "_"))
    NombreTabla = Replace(NombreTabla, ".", " ")
 End If
 'Dialog1.ShowSave
 'tabla de la base de datos SQL del cliente
 lTabla = fPersonasAct.lTablas.List(fPersonasAct.lTablas.ListIndex)
 
 
 SqlTxt = "CREATE TABLE [" & NombreTabla & "] ("
 
 Set lReg = DBConexionSQL.Execute("Select * from [" & lTabla & "] order by NroFoto")
 If lReg.EOF = False Then
    For i = 0 To lReg.Fields.Count - 1
     If UCase(lReg.Fields(i).Name) <> "FECHA" And UCase(lReg.Fields(i).Name) <> "MARCA" And UCase(lReg.Fields(i).Name) <> "CONTADOR" Then
       If UCase(lReg.Fields(i).Name) = "ID" Then
          lCampo = lReg.Fields(i).Name & " int Primary Key, "
       Else
          lCampo = lReg.Fields(i).Name & " char(50), "
       End If
       SqlTxt = SqlTxt & lCampo
     End If
    Next i
    SqlTxt = Mid(SqlTxt, 1, Len(SqlTxt) - 2) & ")"
    '''Data1.DatabaseName
    'Clipboard.Clear
    'Clipboard.SetText SqlTxt
    Data1.Database.Execute SqlTxt
   Do While lReg.EOF = False
      lCampo = ""
      SqlTxt = "Insert into [" & NombreTabla & "] values('"
      For i = 0 To lReg.Fields.Count - 1
         If UCase(lReg.Fields(i).Name) = "NROFOTO" Then
            lCampo = lCampo & "DSC" & Right("00000" & Trim(lReg.Fields(i).value), 5) & ".JPG','"
         Else
            If UCase(lReg.Fields(i).Name) <> "FECHA" And UCase(lReg.Fields(i).Name) <> "MARCA" And UCase(lReg.Fields(i).Name) <> "CONTADOR" Then
               lCampo = lCampo & Trim(lReg.Fields(i).value) & "','"
            End If
         End If
      Next i
      SqlTxt = SqlTxt & lCampo
      SqlTxt = Mid(SqlTxt, 1, Len(SqlTxt) - 2) & ")"
      Clipboard.Clear
      Clipboard.SetText SqlTxt

      Data1.Database.Execute SqlTxt
      lReg.MoveNext
      
   Loop
     MsgBox "Tabla " & NombreTabla & " Generada en " & Data1.DatabaseName, vbInformation
  
 End If
 


falla:
  If Err.Number <> 0 Then
     MsgBox Err.Number & "::" & Err.Description & " Vuelva a intentarlo", vbCritical
     Data1.Database.Execute "DROP TABLE [" & NombreTabla & "]"
  End If
End Sub

Private Sub Form_Load()
   sCargarListado
   sCargarFotosNumeradas
   Data1.DatabaseName = App.Path & "\FotosEnSitio.mdb"
End Sub
 
Public Sub sUltimoNumero()
   Dim lReg As New ADODB.Recordset
   Set lReg = DBConexionSQL.Execute("select max(nrofoto) as NroFoto from [" & fPersonasAct.lTablas.List(fPersonasAct.lTablas.ListIndex) & "]")
   If IsNull(lReg!NROFOTO) = False Then
      lUltimoNumero = lReg!NROFOTO
   Else
      lUltimoNumero = 0
   End If
End Sub

Private Sub ListVFotosNumeradas_DblClick()
   If ListVFotosNumeradas.ListItems.Count <= 0 Then Exit Sub
   If MsgBox("¿Seguro que desea quitar de la lista a " & Trim(ListVFotosNumeradas.SelectedItem.SubItems(2)) & " ?", vbQuestion + vbYesNo) = vbYes Then
      DBConexionSQL.Execute "Update [" & fPersonasAct.lTablas.List(fPersonasAct.lTablas.ListIndex) & "] set NroFoto=0 where CEdula='" & ListVFotosNumeradas.SelectedItem.SubItems(1) & "'"
      sCargarListado
      sCargarFotosNumeradas

   End If

End Sub

Private Sub ListvResultados_DblClick()
  If ListvResultados.ListItems.Count <= 0 Then Exit Sub
  frmNroFotos.Show
  frmNroFotos.lblNombre.Caption = Trim(ListvResultados.SelectedItem.SubItems(2))
  frmNroFotos.txtNroFoto.Text = lUltimoNumero + 1
  frmNroFotos.txtCedula.Text = Trim(ListvResultados.SelectedItem.SubItems(1))
End Sub

Private Sub ListvResultados_KeyPress(KeyAscii As Integer)
  If KeyAscii = 10 Or KeyAscii = 13 Then
     ListvResultados_DblClick
  End If
End Sub
