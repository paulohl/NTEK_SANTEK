VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form fPagos2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Pagos"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInventario 
      Caption         =   "Inventario"
      Height          =   615
      Left            =   4320
      Picture         =   "fPagos2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2640
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton bAceptar 
      Caption         =   "Aceptar"
      Height          =   500
      Left            =   1830
      Picture         =   "fPagos2.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2670
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   500
      Left            =   2970
      Picture         =   "fPagos2.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2670
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Height          =   2565
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   5685
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   540
         Width           =   3975
      End
      Begin VB.TextBox eBanco 
         Height          =   315
         Left            =   1290
         MaxLength       =   20
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2130
         Width           =   3945
      End
      Begin VB.TextBox eNumero 
         Height          =   315
         Left            =   1290
         MaxLength       =   20
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1740
         Width           =   3945
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Pago"
         Height          =   675
         Left            =   2850
         TabIndex        =   16
         Top             =   930
         Width           =   2685
         Begin VB.OptionButton Option3 
            Caption         =   "TDB"
            Height          =   195
            Left            =   1920
            TabIndex        =   20
            Top             =   300
            Width           =   675
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Cheque"
            Height          =   195
            Left            =   150
            TabIndex        =   8
            Top             =   315
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Efectivo"
            Height          =   195
            Left            =   1020
            TabIndex        =   9
            Top             =   300
            Width           =   915
         End
      End
      Begin MSComCtl2.DTPicker eFecha 
         Height          =   285
         Left            =   1290
         TabIndex        =   2
         Top             =   930
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         Format          =   55508993
         CurrentDate     =   40014
      End
      Begin VB.TextBox ePre 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1290
         TabIndex        =   3
         Text            =   "Text3"
         Top             =   1320
         Width           =   1065
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   660
         TabIndex        =   17
         Top             =   2190
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Subcliente:"
         Height          =   195
         Left            =   390
         TabIndex        =   15
         Top             =   600
         Width           =   795
      End
      Begin VB.Label lCliente 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         Height          =   255
         Left            =   1290
         TabIndex        =   0
         Top             =   240
         Width           =   4000
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   600
         TabIndex        =   14
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto Bs:"
         Height          =   195
         Left            =   480
         TabIndex        =   13
         Top             =   1350
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   720
         TabIndex        =   12
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   660
         TabIndex        =   11
         Top             =   270
         Width           =   525
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   465
      Left            =   180
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   1065
   End
End
Attribute VB_Name = "fPagos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bCancelar_Click()
  Modulo.vMontoAnterior = 0#
  Unload Me
End Sub

Private Sub cmdInventario_Click()
  frmInventario.Show vbModal
End Sub

Private Sub eFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then ePre.SetFocus
End Sub

Private Sub eNumero_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then eBanco.SetFocus
End Sub

Private Sub eBanco_KeyPress(KeyAscii As Integer)
  'If KeyAscii = vbKeyReturn Then eObservaciones.SetFocus
End Sub

Private Sub ePre_GotFocus()
  ePre.SelStart = 0
  ePre.SelLength = Len(ePre.Text)
End Sub

Private Sub ePre_KeyPress(KeyAscii As Integer)
  If KeyAscii <> vbKeyDelete And _
     KeyAscii <> vbKeyBack And _
     KeyAscii <> vbKeyReturn And _
     KeyAscii <> Asc(".") Then
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0
      Beep
    End If
  Else
    If KeyAscii = vbKeyReturn Then eNumero.SetFocus
    If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
  End If
End Sub

Private Sub ePre_LostFocus()
  Dim d As Double
  Dim s As String
  s = Trim(ePre.Text)
  If s = "" Then s = "0,00"
  On Error Resume Next
  d = CDbl(s)
  If Err.Number <> 0 Then
    MsgBox "El monto No es válido, Revise...", vbCritical, "Información"
    ePre.SetFocus
  Else
    ePre.Text = Format(d, "#,0.00")
  End If
End Sub

Private Sub bAceptar_Click()
  Dim s As String, sSQL As String
  Dim dNE As Double, dAD As Double
  Dim dFecha As Date
  Dim s1 As String, s2 As String, s3 As String, s4 As String
  Dim bEditaExistencia As Boolean
  Dim dExNueva As Double
  Dim sTipo As String
  Dim dMonto As Double
  
  If IsNumeric(ePre.Text) Then
    dMonto = CDbl(ePre.Text)
    If dMonto = 0# Then
      MsgBox "Falta Monto del Pago a Registrar, Revise...", vbCritical, "Información"
      ePre.SetFocus
      Exit Sub
    End If
  End If
  If Option2.value = True Then
     If eNumero.Text = "" Then
        MsgBox "Debe introducir el número de Cheque", vbExclamation
        Exit Sub
     End If
     
     If eBanco.Text = "" Then
        MsgBox "Debe Introducir el Banco", vbExclamation
        Exit Sub
     End If
  End If
  
  If Option3.value = True Then
     If eNumero.Text = "" Then
        MsgBox "Debe introducir el número de tarjeta", vbExclamation
        Exit Sub
     End If
     
     If eBanco.Text = "" Then
        MsgBox "Debe Introducir el Banco", vbExclamation
        Exit Sub
     End If
  End If
  
  
  
  If Me.Caption = "AGREGAR PAGO DE CLIENTE" Then
  
    '-Para los formatos numericos del SQL server con PUNTO (USA-Convention)
    s1 = ePre.Text
    s1 = Str(CDbl(s1))
    
    s2 = Trim(Combo1.Text)
    If Trim(s2) = "" Or Trim(s2) = "-" Then s2 = "0" Else s2 = Mid(Combo1.Text, 1, 6)
        
    s3 = Format(eFecha.value, "yyyymmdd")
    
    sTipo = ""
    If Option1.value Then
       sTipo = "E"
       eNumero.Text = ""
       eBanco.Text = ""
    End If
    If Option2.value Then sTipo = "C"
        
    sSQL = "insert into pagos (cliente,subcliente,fecha,monto,tipo,numero,banco) values (" & _
           Mid(lCliente.Caption, 1, 6) & "," & _
           s2 & ",'" & s3 & "'," & s1 & ",'" & sTipo & "','" & Trim(eNumero.Text) & "','" & Trim(eBanco.Text) & "')"
           
    Modulo.ExecSQL sSQL
    
    If Err.Number = 0 Then
    
      'Actualizar el monto PAGO del cliente:
      Modulo.Actualizar_Pago_Cliente CLng(Mid(lCliente.Caption, 1, 6)), CLng(s2), Val(s1)
    
      
      fPagos.Adodc2.Refresh
      fPagos.DataGrid2.Refresh
      fPagos.Pagos_Format_DataGrid
      
      AgregarLogs "Agrega Pago Cliente [" & Mid(lCliente.Caption, 1, 20) & "]"
           
      
      ePre.Text = "0,00"
      eNumero.Text = ""
      eBanco.Text = ""
      'eObservaciones.Text = ""
    
      eFecha.SetFocus
    End If
    
  Else
  
    'EDITAR PAGO DE CLIENTE
    
    
    s1 = ePre.Text
    s1 = Str(CDbl(s1))
    
    s2 = Trim(Combo1.Text)
    If Trim(s2) = "" Then s2 = "0"
    
    If s2 <> "0" Then s2 = Mid(s2, 1, 6)
        
    s3 = Format(eFecha.value, "yyyymmdd HH:mm")
    
    sTipo = ""
    If Option1.value Then sTipo = "E"
    If Option2.value Then sTipo = "C"
    
    dMonto = Val(s1)
            
    If Modulo.vMontoAnterior <> dMonto Then 'Hubo Cambio de Monto!
    
      Modulo.Actualizar_Pago_Cliente Trim(Mid(lCliente.Caption, 1, 6)), CLng(s2), (Modulo.vMontoAnterior * -1#)
      
      Modulo.Actualizar_Pago_Cliente Trim(Mid(lCliente.Caption, 1, 6)), CLng(s2), dMonto
      
    End If
    
    sSQL = "update pagos set " & _
           "fecha = '" & s3 & "'," & _
           "monto = " & s1 & "," & _
           "tipo  = '" & sTipo & "'," & _
           "numero = '" & Trim(eNumero.Text) & "'," & _
           "banco = '" & Trim(eBanco.Text) & "' " & _
           "where " & _
           "id = " & Label4.Caption & " "
           
'           "cliente = " & Mid(lCliente.Caption, 1, 6) & " and " & _
'           "subcliente = " & s2 & " "
           
    Modulo.ExecSQL sSQL
      
    fPagos.Adodc2.Refresh
    fPagos.DataGrid2.Refresh
    fPagos.Pagos_Format_DataGrid
    
    Modulo.vMontoAnterior = 0#
    
    AgregarLogs "Edita Pago Cliente [" & Mid(lCliente.Caption, 1, 20) & "]"

    Unload Me
  End If
    
  fPagos.Totalizar_Pagos
  fPagos.Combo2_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'fProductos.DataGrid1.Refresh
  fPagos.DataGrid1.Visible = False
  fPagos.DataGrid1.Visible = True
  'fPagos.DataGrid1.SetFocus
End Sub

Private Sub Option1_Click()
   eNumero.Enabled = False
   eBanco.Enabled = False
End Sub

Private Sub Option2_Click()
   eNumero.Enabled = True
   eBanco.Enabled = True
End Sub

Private Sub Option3_Click()
   eNumero.Enabled = True
   eBanco.Enabled = True
End Sub
