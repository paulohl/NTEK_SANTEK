VERSION 5.00
Begin VB.Form fMov2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle del Movimiento"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bCard5 
      Caption         =   "CARD-5"
      Height          =   495
      Left            =   6600
      Picture         =   "fMov2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.TextBox txtObservaciones 
      Height          =   855
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   5280
      Width           =   3435
   End
   Begin VB.CommandButton cmdPagar 
      Caption         =   "Pagar"
      Height          =   495
      Left            =   5400
      TabIndex        =   11
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   2580
      TabIndex        =   10
      Top             =   3480
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CommandButton bAceptar 
      Caption         =   "Aceptar"
      Height          =   500
      Left            =   4140
      Picture         =   "fMov2.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5610
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   30
      TabIndex        =   0
      Top             =   300
      Width           =   10520
   End
   Begin VB.Label Label8 
      Caption         =   "Observaciones:"
      Height          =   315
      Left            =   300
      TabIndex        =   13
      Top             =   4920
      Width           =   1875
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Entreg."
      Height          =   255
      Left            =   9960
      TabIndex        =   9
      Top             =   30
      Width           =   600
   End
   Begin VB.Label lTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999.999,99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7920
      TabIndex        =   8
      Top             =   4920
      Width           =   1995
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Bs."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6690
      TabIndex        =   7
      Top             =   4920
      Width           =   1275
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Subtotal"
      Height          =   255
      Left            =   8190
      TabIndex        =   5
      Top             =   30
      Width           =   1755
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Precio"
      Height          =   255
      Left            =   7020
      TabIndex        =   4
      Top             =   30
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   6270
      TabIndex        =   3
      Top             =   30
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripción"
      Height          =   255
      Left            =   2250
      TabIndex        =   2
      Top             =   30
      Width           =   4000
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Código"
      Height          =   255
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   2200
   End
End
Attribute VB_Name = "fMov2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Localizador As String
Private Sub bAceptar_Click()
'  Dim sN As String
'  Dim s As String
'  Dim sCadena As String
'  Dim i As Integer
'  If List1.ListIndex >= 0 Then
'
'    For i = 0 To List1.ListCount - 1
'
'      sCadena = List1.List(i)
'
'      sN = Trim(Mid(sCadena, Len(sCadena) - 1, 1))  'S - N
'
'      s = "update DiarioDetalle set " & _
'          "Entregado = '" & sN & "' where id = " & List2.List(i)
'
'      Modulo.ExecSQL s
'    Next i
'  End If
'
  Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdPagar_Click()
  Dim SC As String, sSC As String
  Dim dTC As Double, dTA As Double, dTS As Double
   
   Modulo.vTemporal1 = ""
             
   Modulo.vTemporal2 = fConsultaMovs.FG.TextMatrix(fConsultaMovs.FG.Row, 2)
   Modulo.vTemporal3 = fConsultaMovs.FG.TextMatrix(fConsultaMovs.FG.Row, 3)
                    
   Modulo.fModalResult = Modulo.fModalResultCANCEL
   Load fObservaciones
        
   fObservaciones.lCliente.Caption = Modulo.vTemporal2
   fObservaciones.lSubCliente.Caption = Modulo.vTemporal3
                    
   dTC = 0#
   dTA = 0#
   dTS = 0#
                    
   Modulo.Resumen_Cuenta_Cliente Modulo.vTemporal2, _
                   Modulo.vTemporal3, _
                   dTC, dTA, dTS

   fObservaciones.lTC.Caption = Format(dTC, "#,0.00")
   fObservaciones.lTA.Caption = Format(dTA, "#,0.00")
   fObservaciones.lTS.Caption = Format(dTS, "#,0.00")
   With fObservaciones
       SC = "Select * from DiarioDetalle where Localizador ='" & Localizador & "' order By ID"
       Dim lCommand As New ADODB.Command
       Dim i As Integer
       lCommand.ActiveConnection = DBConexionSQL.ConnectionString
       lCommand.CommandType = adCmdText
       lCommand.CommandTimeout = 15
       lCommand.CommandText = ("Select * from DiarioDetalle where Localizador ='" & Localizador & "' order By ID")
       Set .lRegDetalles = lCommand.Execute()
       Do While .lRegDetalles.EOF = False
          For i = 1 To .FG.Rows - 1
             If .FG.TextMatrix(i, 1) = Trim(.lRegDetalles!CodigoProducto) Then
                .FG.Row = i
                .FG.Col = 0
                .FG_Click
                Exit For
             End If
          Next i
             .lRegDetalles.MoveNext
       Loop
       
   End With
   fObservaciones.EsActualizar = True
   fObservaciones.sLocalizador = Localizador
   fObservaciones.Show vbModal
   fConsultaMovs.bBuscar_Click
   Unload Me
   ''If Modulo.fModalResult = Modulo.fModalResultOK Then
   ''   If Modulo.vTemporal1 <> "" Then GRID.TextMatrix(Row, Col) = Trim(Modulo.vTemporal1)
                       
         'GuardarMovimientos GRID.Row
                          
         'Marcar_Personas_CEDULA C3, s
   ''      MsgBox "Registro de Movimiento Guardado...", vbInformation, "Información"
   ''End If
  
End Sub

Private Sub bCard5_Click()
  Dim SC As String, sSC As String
  Dim sRuta As String, sCard5 As String
  Dim s As String
  Dim i As Integer
  
  Load fMensaje
  fMensaje.Caption = "Preparando para Ejecutar «CARD-5», Espere..."
  'fMensaje.Show
  DoEvents
  
  
    SC = Trim(fConsultaMovs.FG.TextMatrix(fConsultaMovs.FG.Row, 2))
    sSC = Trim(fConsultaMovs.FG.TextMatrix(fConsultaMovs.FG.Row, 3))
    
    If SC <> "" Then
      SC = Mid(SC, 1, 6)
      If sSC = "" Or sSC = "-" Then sSC = "0" Else sSC = Mid(sSC, 1, 6)
      
      ''Auditar_Fotos Modulo.La_Tabla_Actual_Personas(SC, sSC)
      
    End If
  
      
  Dim sOri As String
  Dim sDes As String
  Dim lResp As String
  Dim lR2 As String
  sOri = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
  sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
  
 '' If GRID.Row >= 1 And GRID.Row <= GRID.Rows Then i = GRID.Row Else Exit Sub
      
    SC = Trim(fConsultaMovs.FG.TextMatrix(fConsultaMovs.FG.Row, 2))
    sSC = Trim(fConsultaMovs.FG.TextMatrix(fConsultaMovs.FG.Row, 3))
      
    SC = Trim(Mid(SC, 7))
       
  If Trim(sSC) = "" Or Trim(sSC) = "-" Then
    sRuta = sDes & "\" & SC & "\CARNET" & "\BASE CARNET " & SC & ".car"
    lR2 = sDes & "\" & SC & "\CARNET"
  Else
    sSC = Trim(Mid(sSC, 7))
    sRuta = sDes & "\" & SC & "\" & sSC & "\CARNET" & "\BASE CARNET " & sSC & ".car"
    lR2 = sDes & "\" & SC & "\" & sSC & "\CARNET"
  End If

  s = GetSetting(APPNAME, "Opciones", "RutaCard5", "")

  sCard5 = s & " " & Chr(34) & sRuta & Chr(34)
  ''verificar si existe mas de una archiv .Car, de sera asi permitirle al usuario seleccionar el .car
  Load frmSeleccionarCard5
  lResp = frmSeleccionarCard5.sCargarArchivosCard5(lR2)
  If lResp <> lR2 Then
     frmSeleccionarCard5.Show vbModal
  Else
     If Shell(sCard5, vbMaximizedFocus) = 0# Then
       MsgBox "Error: No se pudo Iniciar Card-5" & vbCrLf & CStr(Err.Number) & ":" & Err.Description, "Información"
     End If
  End If
  Unload fMensaje
    
End Sub



Private Sub List1_DblClick()
 ' Dim sN As String
 ' Dim s As String
 ' Dim sCadena As String
  
 ' If List1.ListIndex >= 0 Then
 '   sCadena = List1.List(List1.ListIndex)
 '   sN = Trim(Mid(sCadena, Len(sCadena) - 1, 2))  'S - N
 '   If sN = "NO" Then
 '     s = Mid(sCadena, 1, Len(sCadena) - 2) & "SI"
 '   Else
 '     s = Mid(sCadena, 1, Len(sCadena) - 2) & "NO"
 '   End If
 '   List1.List(List1.ListIndex) = s
 ' End If
    
End Sub

