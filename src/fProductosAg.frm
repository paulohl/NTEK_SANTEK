VERSION 5.00
Begin VB.Form fProductosAg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar / Editar Producto"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2745
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6405
      Begin VB.Frame Frame2 
         Caption         =   "Modificar Existencia"
         Height          =   1035
         Left            =   2550
         TabIndex        =   14
         Top             =   990
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox eAD 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1740
            TabIndex        =   6
            Text            =   "Text3"
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "En cantidad:"
            Height          =   195
            Left            =   750
            TabIndex        =   15
            Top             =   390
            Width           =   900
         End
      End
      Begin VB.CommandButton bCancelar 
         Caption         =   "Cancelar"
         Height          =   500
         Left            =   5340
         Picture         =   "fProductosAg.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2070
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton bAceptar 
         Caption         =   "Aceptar"
         Height          =   500
         Left            =   4200
         Picture         =   "fProductosAg.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2070
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.TextBox eCre 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         TabIndex        =   5
         Text            =   "Text6"
         Top             =   1800
         Width           =   1065
      End
      Begin VB.TextBox eExi 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1290
         TabIndex        =   4
         Text            =   "Text5"
         Top             =   1410
         Width           =   1065
      End
      Begin VB.TextBox ePre 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1290
         TabIndex        =   3
         Text            =   "Text3"
         Top             =   1020
         Width           =   1065
      End
      Begin VB.TextBox eDes 
         Height          =   315
         Left            =   1290
         MaxLength       =   100
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   630
         Width           =   4995
      End
      Begin VB.TextBox eCod 
         Height          =   315
         Left            =   1290
         MaxLength       =   20
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   210
         Width           =   2025
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Creado el:"
         Height          =   195
         Left            =   450
         TabIndex        =   13
         Top             =   1830
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Existencia:"
         Height          =   195
         Left            =   450
         TabIndex        =   12
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Precio:"
         Height          =   195
         Left            =   720
         TabIndex        =   11
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   330
         TabIndex        =   10
         Top             =   660
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   660
         TabIndex        =   9
         Top             =   270
         Width           =   540
      End
   End
End
Attribute VB_Name = "fProductosAg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bCancelar_Click()
  Unload Me
End Sub

Private Sub eCod_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then eDes.SetFocus
End Sub

Private Sub eCod_LostFocus()
  Dim s As String, sSQL As String
  eCod.Text = UCase(eCod.Text)
  s = Trim(eCod.Text)
  If Me.Caption = "AGREGAR PRODUCTO" Then
    If Modulo.DBExiste("productos", "codigo", s) = True Then
      MsgBox "El Código está Registrado en otro Artículo, Revise...", vbCritical, "Información"
      eCod.SetFocus
    End If
  End If
End Sub

Private Sub eDes_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then ePre.SetFocus
End Sub

Private Sub eDes_LostFocus()
  eDes.Text = UCase(eDes.Text)
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
    'If KeyAscii = vbKeyReturn Then eAD.SetFocus
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

Private Sub eExi_GotFocus()
  eExi.SelStart = 0
  eExi.SelLength = Len(eExi.Text)
End Sub

Private Sub eExi_KeyPress(KeyAscii As Integer)
  If KeyAscii <> vbKeyDelete And _
     KeyAscii <> vbKeyBack And _
     KeyAscii <> vbKeyReturn And _
     KeyAscii <> Asc(".") Then
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0
      Beep
    End If
  Else
    If KeyAscii = vbKeyReturn Then bAceptar.SetFocus
    If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
  End If
End Sub

Private Sub eExi_LostFocus()
  Dim d As Double
  Dim s As String
  s = Trim(eExi.Text)
  If s = "" Then s = "0,00"
  On Error Resume Next
  d = CDbl(s)
  If Err.Number <> 0 Then
    MsgBox "La cantidad No es válida, Revise...", vbCritical, "Información"
    eExi.SetFocus
  Else
    eExi.Text = Format(d, "#,0.00")
  End If
End Sub

Private Sub eAD_GotFocus()
  eAD.SelStart = 0
  eAD.SelLength = Len(eAD.Text)
End Sub

Private Sub eAD_KeyPress(KeyAscii As Integer)
  If KeyAscii <> vbKeyDelete And _
     KeyAscii <> vbKeyBack And _
     KeyAscii <> vbKeyReturn And _
     KeyAscii <> Asc("-") And _
     KeyAscii <> Asc(".") Then
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0
      Beep
    End If
  Else
    If KeyAscii = vbKeyReturn Then bAceptar.SetFocus
    If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
  End If
End Sub

Private Sub eAD_LostFocus()
  Dim d As Double
  Dim s As String
  s = Trim(eAD.Text)
  If s = "" Then s = "0,00"
  On Error Resume Next
  d = CDbl(s)
  If Err.Number <> 0 Then
    MsgBox "El monto No es válido, Revise...", vbCritical, "Información"
    eAD.SetFocus
  Else
    eAD.Text = Format(d, "#,0.00")
  End If
End Sub


Private Sub bAceptar_Click()
  Dim s As String, sSQL As String
  Dim dNE As Double, dAD As Double
  Dim dFecha As Date
  Dim s1 As String, s2 As String, s3 As String
  Dim bEditaExistencia As Boolean
  Dim dExNueva As Double
  
  
  If Trim(eCod.Text) = "" Then
    MsgBox "Falta Código del Artículo, Revise...", vbCritical, "Información"
    eCod.SetFocus
    Exit Sub
  End If
    
  If Trim(eDes.Text) = "" Then
    MsgBox "Falta Descripción del Artículo, Revise...", vbCritical, "Información"
    eDes.SetFocus
    Exit Sub
  End If
    
  If Trim(ePre.Text) = "" Then
    MsgBox "Falta Precio del Artículo, Revise...", vbCritical, "Información"
    ePre.SetFocus
    Exit Sub
  End If
  

  If Trim(eExi.Text) = "" Then
    MsgBox "Falta Existencia del Artículo, Revise...", vbCritical, "Información"
    eExi.SetFocus
    Exit Sub
  End If
  
  If Me.Caption = "AGREGAR PRODUCTO" Then
  
    '-Para los formatos numericos del SQL server con PUNTO (USA-Convention)
    s1 = eExi.Text
    s1 = Str(CDbl(s1))
    
    s2 = ePre.Text
    s2 = Str(CDbl(s2))
    
    s3 = eCre.Text
    dFecha = CDate(s3)
    s3 = Format(dFecha, "yyyymmdd HH:mm")
    
    AgregarLogs "Agrega Producto [" & eCod.Text & "]"
    
    
    sSQL = "insert into productos (codigo,descripcion,existencia,precio,creado) values " & _
           "('" & eCod.Text & "','" & _
                  eDes.Text & "', " & _
                  s1 & " , " & _
                  s2 & " ,'" & _
                  s3 & "')"
    
    Modulo.ExecSQL sSQL
    
    
    
    If Err.Number = 0 Then
    
      fProductos.Adodc1.Refresh
      fProductos.DataGrid1.Refresh
      fProductos.Productos_Format_DataGrid
    
      eCod.Text = ""
      eDes.Text = ""
      ePre.Text = "0,00"
      eExi.Text = "0,00"
      eCod.SetFocus
    End If
    
    
    
  Else
  
    'EDITAR PRODUCTO
    eAD.Text = Trim(eAD.Text)
    If eAD.Text = "" Then eAD.Text = "0,00"
    On Error Resume Next
    dNE = CDbl(eAD.Text)
    
    dExNueva = CDbl(eAD.Text)
    
    If Err.Number <> 0 Then
      MsgBox "El Monto del Aumento/Disminución de Existencia es Inválido, Revise...", vbCritical, "Información"
      eAD.SetFocus
      Exit Sub
    End If
    
    bEditaExistencia = False
    If dNE <> 0 Then bEditaExistencia = True
    
    dNE = CDbl(eExi.Text) + CDbl(dNE)
    
    '-Para los formatos numericos del SQL server con PUNTO (USA-Convention)
    s1 = eExi.Text
    s1 = Str(CDbl(s1))
    
    s2 = ePre.Text
    s2 = Str(CDbl(s2))
    
    s3 = dNE
    s3 = Str(dNE)
    
    AgregarLogs "Edita Producto [" & eCod.Text & "]"
        
    sSQL = "update productos set " & _
           "descripcion = '" & eDes.Text & "'," & _
           "existencia  = " & s3 & "," & _
           "precio      = " & s2 & " " & _
           "where codigo='" & eCod.Text & "'"
           
    Modulo.ExecSQL sSQL
    If Err.Number = 0 Then
    
      If bEditaExistencia Then
        sSQL = "insert into ProductosMov (codigo,fecha,hora,cantidad,tipo) values ('" & _
                eCod.Text & "','" & _
                Format(Date, "yyyymmdd") & "','" & _
                Format(Time, "HH:mm") & "'," & _
                Str(dExNueva) & ",'" & _
                "A/D')"
        Modulo.ExecSQL sSQL
      End If
      
      fProductos.Adodc1.Refresh
      fProductos.DataGrid1.Refresh
      fProductos.Productos_Format_DataGrid

      Unload Me
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'fProductos.DataGrid1.Refresh
  fProductos.DataGrid1.Visible = False
  fProductos.DataGrid1.Visible = True
  fProductos.DataGrid1.SetFocus
End Sub
