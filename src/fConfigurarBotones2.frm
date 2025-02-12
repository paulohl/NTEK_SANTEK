VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fConfigurarBotones2 
   Caption         =   "Configurar Botones de SUB-MENU"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bAceptar 
      Caption         =   "Guardar"
      Height          =   500
      Left            =   4950
      Picture         =   "fConfigurarBotones2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3810
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   500
      Left            =   6120
      Picture         =   "fConfigurarBotones2.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3810
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11925
      Begin VB.Frame Frame2 
         Height          =   3465
         Left            =   60
         TabIndex        =   3
         Top             =   180
         Width           =   11655
         Begin MSFlexGridLib.MSFlexGrid FG2 
            Height          =   3135
            Left            =   90
            TabIndex        =   4
            Top             =   240
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   5530
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            AllowBigSelection=   0   'False
            SelectionMode   =   1
            AllowUserResizing=   3
            FormatString    =   $"fConfigurarBotones2.frx":0B14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6480
            TabIndex        =   6
            Top             =   0
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   " Botón Principal Nº "
            Height          =   195
            Left            =   5100
            TabIndex        =   5
            Top             =   0
            Width           =   1380
         End
      End
   End
End
Attribute VB_Name = "fConfigurarBotones2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const cFG2 = "Nº   | Título                                                   | Producto                                                  | Imagen                                                                                      "


Private Sub Limpiar_PPAL()
  Dim i As Integer
  FG2.Clear
  FG2.Rows = 1
  FG2.FormatString = cFG2
  
  For i = 0 To MAX_BTNS - 1
    FG2.Rows = FG2.Rows + 1
    FG2.TextMatrix(FG2.Rows - 1, 0) = CStr(i + 1)
  Next i
  
  
End Sub

Private Sub bAceptar_Click()
  Dim r As New ADODB.Recordset
  Dim s As String, i As Integer, j As Integer
  
  Load fMensaje
  fMensaje.Caption = "Guardando Configuración, Espere..."
  fMensaje.Show
  DoEvents
  
  If InStr(Me.Caption, "SUB-MENU") <= 0 Then
    
    s = "select * from Botones"
    r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    If r.EOF Then
      'No tiene registros, completar en blando la tabla...
      For i = 0 To MAX_BTNS - 1
        s = "insert into Botones (posicion,caption,producto,imagen) values (" & _
             CStr(i) & ",'','','')"
        Modulo.ExecSQL s
      Next i
    End If
    
    For i = 1 To FG2.Rows - 1
      s = "update Botones set caption = '" & Trim(FG.TextMatrix(i, 1)) & "' where Posicion = " & CStr(i - 1)
      Modulo.ExecSQL s
      
      s = "update Botones set producto = '" & Trim(Mid(FG.TextMatrix(i, 2), 1, 20)) & "' where Posicion = " & CStr(i - 1)
      Modulo.ExecSQL s
      
      s = "update Botones set imagen = '" & Trim(FG.TextMatrix(i, 3)) & "' where Posicion = " & CStr(i - 1)
      Modulo.ExecSQL s
    Next i
    
  Else
  
    s = "delete Botones2 where BotonPrincipal = " & Label2.Caption
    Modulo.ExecSQL s
    
    For i = 0 To MAX_BTNS - 1
    
      s = "insert into Botones2 (BotonPrincipal,Posicion,Titulo,Producto,Imagen) values " & _
              "(" & Label2.Caption & "," & CStr(i) & ",'" & _
              FG2.TextMatrix(i + 1, 1) & "','" & _
              Mid(FG2.TextMatrix(i + 1, 2), 1, 20) & "','" & _
              FG2.TextMatrix(i + 1, 3) & "')"
              
      Modulo.ExecSQL s
    Next i
    
  End If
    
  Unload fMensaje
  
  Unload Me
  
End Sub

Private Sub bCancelar_Click()
  Unload Me
End Sub

Private Sub CargarProductos(xCombo As ComboBox)
  Dim s As String
  Dim r As New ADODB.Recordset
  s = "select * from Productos order by codigo"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  
  xCombo.Clear
  xCombo.AddItem ""
  
  'xList.Clear
  'xList.AddItem ""
  
  Do While Not r.EOF
    'xList.AddItem r.Fields("codigo").Value
    xCombo.AddItem r.Fields("codigo").Value & " : " & r.Fields("descripcion").Value
    r.MoveNext
  Loop
  r.Close
  Set r = Nothing
End Sub

Private Sub CargarSubMenu(iPosicion As Integer)
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim i As Integer, j As Integer
  
  FG2.Clear
  FG2.Rows = 1
  FG2.FormatString = cFG2
   
  
  s = "select * from Botones2 where BotonPrincipal = " & CStr(iPosicion) & " order by posicion"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If r.EOF Then
    
    For i = 0 To Modulo.MAX_BTNS - 1
      FG2.Rows = FG2.Rows + 1
      FG2.TextMatrix(i + 1, 0) = CStr(i + 1)
      FG2.TextMatrix(i + 1, 1) = ""
      FG2.TextMatrix(i + 1, 2) = ""
      FG2.TextMatrix(i + 1, 3) = ""
    Next i
    
  Else
    
    i = 1
    Do While Not r.EOF
      FG2.Rows = FG2.Rows + 1
      FG2.TextMatrix(i, 0) = CStr(r.Fields("posicion").Value)
      FG2.TextMatrix(i, 1) = Trim(r.Fields("titulo").Value)
      FG2.TextMatrix(i, 2) = r.Fields("producto").Value & " : " & Modulo.DBValorStr("productos", "codigo", r.Fields("producto").Value, "descripcion")
      FG2.TextMatrix(i, 3) = Trim(r.Fields("imagen").Value)
      r.MoveNext
    Loop
    
  End If
  
  r.Close
  Set r = Nothing
  
End Sub

Private Sub FG_RowColChange()
  Dim s As String
  Dim i As Integer
  
  s = "delete from Botones2 where BotonPrincipal = " & CStr(CInt(Label2.Caption) - 1)
  Modulo.ExecSQL s
  
  For i = 0 To Modulo.MAX_BTNS - 1
    s = "insert into Botones2 (botonprincipal,posicion,titulo,producto,imagen) values (" & _
         Label2.Caption & "," & CStr(i) & ",'" & FG2.TextMatrix(i + 1, 1) & "','" & Mid(FG2.TextMatrix(i + 1, 2), 1, 20) & "','" & FG2.TextMatrix(i + 1, 3) & "')"
    
    Modulo.ExecSQL s
  Next i
  
End Sub



Private Sub FG_Click()
  Dim sN As String
  
  If FG.Row < FG.Rows And FG.Col < FG.Cols Then
  
    sN = FG.TextMatrix(FG.Row, FG.Col)
    
    Label2.Caption = sN
    
    CargarSubMenu CInt(sN)
    
  End If
    
    
    



End Sub

Private Sub FG_DblClick()
  Dim s As String
  If FG.Row < FG.Rows Then
    If FG.Col < FG.Cols Then
      Load fDatosBoton
      With fDatosBoton
        .lNumero.Caption = FG.TextMatrix(FG.Row, 0)
        .eTitulo.Text = FG.TextMatrix(FG.Row, 1)
        CargarProductos .cProductos
        If .cProductos.ListCount > 0 Then .cProductos.ListIndex = 0
        .eImagen.Text = FG.TextMatrix(FG.Row, 3)
      End With
      fDatosBoton.Show
    End If
  End If
End Sub


Private Sub FG2_DblClick()
  Dim s As String
  If FG2.Row < FG2.Rows Then
    If FG2.Col < FG2.Cols Then
      Load fDatosBoton
      With fDatosBoton
        .Caption = "DATOS DEL SUB-BOTÓN"
        .lNumero.Caption = FG2.TextMatrix(FG2.Row, 0)
        .eTitulo.Text = FG2.TextMatrix(FG2.Row, 1)
        CargarProductos .cProductos
        If .cProductos.ListCount > 0 Then .cProductos.ListIndex = 0
        .eImagen.Text = FG2.TextMatrix(FG2.Row, 3)
      End With
      fDatosBoton.Show
    End If
  End If
End Sub


Private Sub Form_Load()
  Limpiar_PPAL
  
  
End Sub
