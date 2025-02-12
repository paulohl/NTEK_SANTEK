VERSION 5.00
Begin VB.Form fPersonas 
   Caption         =   "Creación Tabla de Personas"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Cliente Principal posee Tabla de Personas"
      Height          =   225
      Left            =   7590
      TabIndex        =   45
      Top             =   540
      Width           =   3825
   End
   Begin VB.Frame Frame6 
      Height          =   4995
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   10515
      Begin VB.Frame Frame1 
         Caption         =   "En Caso de RENOVACIÓN"
         Height          =   1245
         Left            =   60
         TabIndex        =   36
         Top             =   3690
         Width           =   6795
         Begin VB.CommandButton bCrear 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Crear"
            Height          =   345
            Left            =   4770
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   750
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Crear la Tabla y copiar todos los registros de Personas"
            Height          =   225
            Left            =   270
            TabIndex        =   40
            Top             =   840
            Value           =   -1  'True
            Width           =   4305
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Crear la Tabla Vacía. (Sin registros de Personas)"
            Height          =   225
            Left            =   270
            TabIndex        =   39
            Top             =   600
            Width           =   4065
         End
         Begin VB.ComboBox cNuevaTbl 
            Height          =   315
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   240
            Width           =   2715
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "> Crear Nueva Tabla con la misma Estructura de:"
            Height          =   195
            Left            =   150
            TabIndex        =   37
            Top             =   300
            Width           =   3480
         End
      End
      Begin VB.CommandButton bCancelar 
         Caption         =   "Salir"
         Height          =   500
         Left            =   8190
         Picture         =   "fPersonas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4140
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.Frame Frame7 
         Caption         =   "Estructura de Datos Tabla de Personas"
         Height          =   3525
         Left            =   4590
         TabIndex        =   10
         Top             =   150
         Width           =   5745
         Begin VB.ListBox List3 
            Height          =   1035
            Left            =   4560
            TabIndex        =   44
            Top             =   1800
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.ListBox List2 
            Height          =   840
            Left            =   2940
            TabIndex        =   43
            Top             =   1890
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.ListBox List1 
            Height          =   1035
            Left            =   840
            TabIndex        =   42
            Top             =   1860
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.CommandButton bBajar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "â"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   9.75
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1650
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   660
            UseMaskColor    =   -1  'True
            Width           =   345
         End
         Begin VB.CommandButton bSubir 
            BackColor       =   &H00E0E0E0&
            Caption         =   "á"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   9.75
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   660
            UseMaskColor    =   -1  'True
            Width           =   345
         End
         Begin VB.CommandButton bActualizarTabla 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Actualizar Tabla"
            Height          =   345
            Left            =   4140
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   3060
            Width           =   1500
         End
         Begin VB.CommandButton bCamposDefecto 
            Caption         =   "Campos Predeterminados"
            Height          =   345
            Left            =   180
            TabIndex        =   26
            Top             =   3060
            Width           =   2040
         End
         Begin VB.CommandButton bCrearTabla 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Crear Nueva Tabla"
            Height          =   345
            Left            =   2370
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   3060
            Width           =   1650
         End
         Begin VB.CommandButton bMenos 
            BackColor       =   &H00C0C0FF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5190
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   990
            UseMaskColor    =   -1  'True
            Width           =   345
         End
         Begin VB.ListBox lAnchos 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1860
            Left            =   4110
            TabIndex        =   22
            Top             =   960
            Width           =   900
         End
         Begin VB.ListBox lTipos 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1860
            Left            =   2340
            TabIndex        =   20
            Top             =   960
            Width           =   1725
         End
         Begin VB.ListBox lCampos 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1860
            Left            =   180
            TabIndex        =   18
            Top             =   960
            Width           =   2085
         End
         Begin VB.CommandButton bMas 
            BackColor       =   &H00FFFFC0&
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5190
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   345
         End
         Begin VB.TextBox eAncho 
            Height          =   315
            Left            =   4590
            MaxLength       =   3
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   240
            Width           =   400
         End
         Begin VB.ComboBox cTipos 
            Height          =   315
            Left            =   2700
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   240
            Width           =   1300
         End
         Begin VB.TextBox eCampo 
            Height          =   315
            Left            =   750
            MaxLength       =   12
            TabIndex        =   12
            Text            =   "eCampo"
            Top             =   240
            Width           =   1500
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Anchos"
            Height          =   255
            Left            =   4110
            TabIndex        =   23
            Top             =   720
            Width           =   700
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tipos"
            Height          =   255
            Left            =   2340
            TabIndex        =   21
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Campos"
            Height          =   255
            Left            =   180
            TabIndex        =   19
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Ancho:"
            Height          =   195
            Left            =   4050
            TabIndex        =   15
            Top             =   300
            Width           =   510
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   2280
            TabIndex        =   13
            Top             =   300
            Width           =   360
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Campo:"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   300
            Width           =   540
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tablas Personas Creadas"
         Height          =   3525
         Left            =   60
         TabIndex        =   6
         Top             =   150
         Width           =   4455
         Begin VB.CommandButton bVerContenido 
            Caption         =   "Ver"
            Height          =   345
            Left            =   3810
            TabIndex        =   32
            Top             =   3000
            Width           =   465
         End
         Begin VB.ListBox lCreadas 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2310
            Left            =   2760
            TabIndex        =   28
            Top             =   480
            Width           =   1600
         End
         Begin VB.CommandButton bBorrarTabla 
            Caption         =   "Borrar Tabla"
            Height          =   345
            Left            =   90
            TabIndex        =   27
            Top             =   3000
            Width           =   1080
         End
         Begin VB.ListBox lTablas 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2310
            Left            =   90
            TabIndex        =   7
            Top             =   480
            Width           =   2390
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "-"
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
            Left            =   3390
            TabIndex        =   30
            Top             =   3090
            Width           =   75
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "REGS:"
            Height          =   195
            Left            =   2790
            TabIndex        =   29
            Top             =   3090
            Width           =   495
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ID"
            Height          =   255
            Left            =   90
            TabIndex        =   9
            Top             =   240
            Width           =   2390
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Creada"
            Height          =   255
            Left            =   2760
            TabIndex        =   8
            Top             =   240
            Width           =   1600
         End
      End
   End
   Begin VB.ComboBox cSC 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   510
      Width           =   5715
   End
   Begin VB.ComboBox cCP 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   5715
   End
   Begin VB.CommandButton bBuscarCP 
      Height          =   345
      Left            =   7560
      Picture         =   "fPersonas.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Sub-Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   570
      TabIndex        =   4
      Top             =   570
      Width           =   1080
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Cliente Principal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   1485
   End
End
Attribute VB_Name = "fPersonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const FGC = "ID          | Nombre                                                    | RIF Nº             | Telefonos                                    "

Dim RClientes As New ADODB.Recordset
Dim RSubClientes As New ADODB.Recordset

Const OPR_NUEVO = 1
Const OPR_EDITAR = 2

Dim OPR As Integer  'operacion: 1 nuevo  2 editar
Dim Viendo As Boolean

Const MAXCAMPOSPRE = 9
Dim aCamposPre(9) As String





Private Sub bActualizarTabla_Click()
  Dim r As New ADODB.Recordset
  Dim cp As String, SC As String
  Dim s As String, sComando As String, sTabla As String
  Dim i As Integer, j As Integer, iNumeroTablaNueva As Integer
  Dim sCampo As String, sTipo As String, sAncho As String
  Dim TR As Long, e As Boolean, sComando2 As String
  
  Dim rnueva As New ADODB.Recordset
  Dim rvieja As New ADODB.Recordset
  Dim EstadoAgregando As Boolean
  
  cp = Trim(Mid(cCP.Text, 1, 6))
  SC = Trim(Mid(cSC.Text, 1, 6))
  
  If Trim(SC) = "" Then SC = "0"
  
  
  i = -1
  If lTablas.ListIndex >= 0 Then i = lTablas.ListIndex
  If lCreadas.ListIndex >= 0 Then i = lCreadas.ListIndex
  
  If i = -1 Then
    MsgBox "Debe Seleccionar la Tabla a Modificar/Actualizar los campos...", vbCritical, "Información"
    Exit Sub
  End If
  
  sTabla = lTablas.List(i)
  
  If MsgBox("¿Está Seguro de ACTUALIZAR los campos de la Tabla [" & sTabla & "]?" & vbCrLf & _
            "Se inicializará el Histórico de Modificaciones.", vbQuestion + vbYesNo, "Confirme") = vbNo Then
    Exit Sub
  End If
  
  '0.Borrar si existe "Temporal" antes de iniciar las operaciones:
  If Existe_Tabla("Temporal") Then ExecSQL "DROP TABLE [Temporal]"
  If Existe_Tabla("H" & sTabla) Then ExecSQL "DROP TABLE [H" & sTabla & "]"
                   
  '1.Crear una Tabla "Nueva" con nombre "Temporal" con los campos
  '  actuales:
  'sComando = "CREATE TABLE [Temporal] ("
  sComando = "CREATE TABLE [Temporal] " & _
             "(ID INT NOT NULL IDENTITY(1,1), "
  
  'sComando2 = "CREATE TABLE [H" & sTabla & "] ("
  sComando2 = "CREATE TABLE [H" & sTabla & "] " & _
              "(ID INT NOT NULL, "
  
  s = ""
  For i = 0 To lCampos.ListCount - 1
    sCampo = lCampos.List(i)
    sTipo = lTipos.List(i)
    sAncho = lAnchos.List(i)

    If sTipo = "CHAR" Then  '--Lleva ancho
      s = sCampo & " " & sTipo & "(" & sAncho & ") "
    Else
      s = sCampo & " " & sTipo & " "
    End If

    If i <> lCampos.ListCount - 1 Then s = s & ","

    sComando = sComando & s
    sComando2 = sComando2 & s
  Next i
  
  sComando = sComando & " PRIMARY KEY(ID))"
  sComando2 = sComando2 & " )"
  
      
  On Error Resume Next
      
  ExecSQL (sComando)
  ExecSQL (sComando2)
  If Err.Number <> 0 Then
     MsgBox "Error: No se pudo Crear la Tabla Temporal en el Sistema..." & vbCrLf & Err.Description, vbCritical, "Información"
  Else
     '2.Abrir la tabla "vieja" y "nueva"... si hay registros => copiarlos
     rnueva.Open "select * from temporal", Modulo.DBConexionSQL, adOpenDynamic, adLockOptimistic
     rvieja.Open "select * from [" & sTabla & "]", Modulo.DBConexionSQL, adOpenKeyset, adLockReadOnly
     EstadoAgregando = False
        
     Do While Not rvieja.EOF
          
       For i = 0 To lCampos.ListCount - 1
            'El campo de la data nueva:
            sCampo = lCampos.List(i)
            
            'Buscarlo en la data vieja:
            j = 0
            e = False
            Do While (j < rvieja.Fields.Count) And Not e
              If sCampo = rvieja.Fields(j).Name Then
                e = True
              Else
                j = j + 1
              End If
            Loop
            
            If e Then
              If Not EstadoAgregando Then
                rnueva.AddNew
                EstadoAgregando = True
              End If
            
              'el campo esta definido, asignarlo para guardarlo:
              If EstadoAgregando Then
                rnueva.Fields(sCampo).value = rvieja.Fields(sCampo).value
              End If
            End If
            
       Next i
          
       If EstadoAgregando Then
           rnueva.Update
           EstadoAgregando = False
       End If
          
       rvieja.MoveNext
     Loop
        
     rvieja.Close
     rnueva.Close
        
     Set rvieja = Nothing
     Set rnueva = Nothing

  End If
  
  '3.Borrar la tabla "vieja" y renombrar "nueva" por la tabla
  '  de datos actual.
  If Existe_Tabla(sTabla) Then
    s = "DROP TABLE [" & sTabla & "]"
    ExecSQL s
  End If
  
  'Exec sp_rename 'Trabajadores', 'Personal'
        
  s = "EXEC sp_rename 'Temporal', '" & sTabla & "'"
  ExecSQL s
  
  '4.Incluir el TRIGGER otra vez pq es una tabla nueva:
  
  's = "CREATE TRIGGER [TRG_" & sTabla & "] ON [" & sTabla & "] " & vbCrLf & _
      "FOR UPDATE AS " & vbCrLf & _
      "IF UPDATE(fecha) OR UPDATE(contador) " & vbCrLf & _
      "BEGIN " & vbCrLf & _
      "  INSERT INTO [H" & sTabla & "] SELECT * FROM DELETED " & vbCrLf & _
      "END"
  
  's = "CREATE TRIGGER [TRG_" & sTabla & "] ON [" & sTabla & "] " & vbCrLf & _
      "FOR UPDATE AS " & vbCrLf & _
      "INSERT INTO [H" & sTabla & "] SELECT * FROM DELETED"
      
      
      
  'sComando = "CREATE TRIGGER [TRG_" & sTabla & "] ON [" & sTabla & "] " & vbCrLf & _
             "FOR UPDATE AS " & vbCrLf & _
             "BEGIN " & vbCrLf & _
             "  declare @CodProducto as char(20) " & vbCrLf & _
             "  declare @PrecioProducto as float " & vbCrLf & _
             "  declare @IDPersona as integer " & vbCrLf & _
             "  --INSERT INTO [H" & sTabla & "] SELECT * FROM DELETED " & vbCrLf & _
             "  if update(CONTADOR) " & vbCrLf & _
             "  begin " & vbCrLf & _
             "    INSERT INTO [H" & sTabla & "] SELECT * FROM DELETED " & vbCrLf & _
             "    set @IDPersona = (SELECT ID FROM DELETED) " & vbCrLf & _
             "    set @CodProducto = '' " & vbCrLf & _
             "    set @PrecioProducto = 0.00 " & vbCrLf & _
             "    set @CodProducto = (select codigoproductopvc from opciones) " & vbCrLf & _
             "    if rtrim(ltrim(@CodProducto)) <> '' " & vbCrLf & _
             "    begin " & vbCrLf & _
             "      Set @PrecioProducto = (Select Precio From Productos Where Codigo = @CodProducto) " & vbCrLf & _
             "      update Productos set existencia = existencia - 1 where Codigo = @CodProducto " & vbCrLf
  'If SC = "0" Then 'Es SOLO cliente:
  '  sComando = sComando & _
             "      update clientes set deuda = deuda + @PrecioProducto where codigo = " & cp & " " & vbCrLf & _
             "      update clientes set saldo = deuda - pagos where codigo = " & cp & " " & vbCrLf
  'Else             'Es sub-cliente:
  '  sComando = sComando & _
             "      update subclientes set deuda = deuda + @PrecioProducto where cliente = " & cp & " and id = " & SC & " " & vbCrLf & _
             "      update subclientes set saldo = deuda - pagos           where cliente = " & cp & " and id = " & SC & " " & vbCrLf
  'End If
  
   ' sComando = sComando & _
             "      insert into [EventosC5] (procesado,idtabla,tabla) values ('N',@IDPersona,'" & sTabla & "') " & vbCrLf & _
             "    end " & vbCrLf & _
             "  end " & vbCrLf & _
             "END" & vbCrLf
      
  'Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
  '    Modulo.DBComandoSQL.CommandText = sComando
  '    Modulo.DBComandoSQL.Execute
       sComando = "Exec CrearTrigger '" & sTabla & "','" & Mid(cCP.Text, 1, 6) & "','0'"
      Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
      Modulo.DBComandoSQL.CommandText = sComando
      Modulo.DBComandoSQL.Execute
  
  sCrearArchivoExcel sTabla, Mid(cCP.Text, 10, Len(cCP.Text))
   
  If Err.Number <> 0 Then
    MsgBox "Error en TRIGGER: " & Err.Description, vbCritical, "Información"
  Else
    MsgBox "Actualización Efectuada Correctamente...", vbInformation, "Información"
  End If

End Sub




Private Sub bBorrarTabla_Click()
  Dim i As Integer
  Dim s As String, s1 As String
  On Error Resume Next
  
  i = -1
  If lTablas.ListIndex >= 0 Then i = lTablas.ListIndex
  If lCreadas.ListIndex >= 0 Then i = lCreadas.ListIndex
  
  If i = -1 Then Exit Sub
  
  s = lTablas.List(i)
  If MsgBox("¿Está Seguro de BORRAR la Tabla Personas [" & s & "] y su Contenido?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
    If Modulo.Existe_Tabla(s) Then
      s1 = "DROP TABLE [" & s & "]"
      ExecSQL s1
    End If
    
    If Modulo.Existe_Tabla("H" & s) Then
      s1 = "DROP TABLE [H" & s & "]"
      ExecSQL s1
    End If
    
    If Err.Number <> 0 Then
      MsgBox "Error al Borrar la Tabla..." & vbCrLf & _
             Err.Description, vbCritical, "Información"
    Else
      
      s1 = "DELETE FROM Personas WHERE Tabla = '" & s & "'"
      ExecSQL s1
      
      If Err.Number <> 0 Then
        MsgBox "Error al Borrar la Tabla..." & vbCrLf & _
                Err.Description, vbCritical, "Información"
      Else
        MsgBox "Tabla Borrada Correctamente...", vbInformation, "Información"
        Cargar_Tablas_Personas
      End If
    End If
  End If
   
End Sub



Private Sub bBuscarCP_Click()
  Dim f As Integer, c As Integer, i As Integer
  Dim e As Boolean
  
  Modulo.vTemporal1 = ""
  Load fBuscarSimple
  
  With fBuscarSimple
    .Adodc1.ConnectionString = Modulo.DBConexionSQL.ConnectionString
    .Adodc1.Caption = "CLIENTES"
    .Adodc1.RecordSource = "select codigo, nombre, direccion, telefonos from clientes order by codigo"
    .Adodc1.Refresh
    .DataGrid1.Refresh
    
    .Combo1.Clear
    .Combo1.AddItem "CODIGO"
    .Combo1.AddItem "NOMBRE"
    .Combo1.ListIndex = 1
        
    Fields_DataGrid_En_Mayusculas .DataGrid1
    
    .DataGrid1.Columns(0).width = 800 'codigo
    .DataGrid1.Columns(1).width = 4000 'nombre
    .DataGrid1.Columns(2).width = 3000 'direccion
    .DataGrid1.Columns(3).width = 2000 'telefonos
  End With
   
  'fBuscarSimple.BuscarSimple fClientes.FG
  Modulo.vTemporal1 = ""
  'fBuscarSimple.Option2.Value = True 'Buscar por nombre por defecto
  fBuscarSimple.Show vbModal
  
  If Modulo.vTemporal1 <> "" Then
    cCP.ListIndex = Modulo.Buscar_ComboLen(cCP, Mid(Modulo.vTemporal1, 1, 6), 6)
  End If

End Sub

Private Sub bCamposDefecto_Click()
  Dim cte As String
  Dim scte As String
  
  cte = Trim(Mid(cCP.Text, 1, 6))
  scte = Trim(Mid(cSC.Text, 1, 6))
  
  If cte = "" Then
    MsgBox "Debe Seleccionar primero el Cliente...", vbCritical, "Información"
    Exit Sub
  End If
  
  If MsgBox("¿Está Seguro que desea cargar la estructura con los campos predeterminados?", vbQuestion + vbYesNo, "Confirme") = vbNo Then
    Exit Sub
  End If
  
  lCampos.Clear
  lTipos.Clear
  lAnchos.Clear
  
  '-- CEDULA:
  lCampos.AddItem "CEDULA"
  lTipos.AddItem "CHAR"
  lAnchos.AddItem "20"
  
  '-- NOMBRE:
  lCampos.AddItem "NOMBRE"
  lTipos.AddItem "CHAR"
  lAnchos.AddItem "50"
  
  '-- CARGO:
  lCampos.AddItem "CARGO"
  lTipos.AddItem "CHAR"
  lAnchos.AddItem "50"
  
  '-- VENCE:
  lCampos.AddItem "VENCE"
  lTipos.AddItem "CHAR"
  lAnchos.AddItem "20"
  
  '-- FOTO:
  lCampos.AddItem "FOTO"
  lTipos.AddItem "CHAR"
  lAnchos.AddItem "20"
  
  '-- TIENE FOTO?:
  lCampos.AddItem "TIENE_FOTO"
  lTipos.AddItem "CHAR"
  lAnchos.AddItem "1"
  
    '-- TIENE MARCA:
  lCampos.AddItem "MARCA"
  lTipos.AddItem "CHAR"
  lAnchos.AddItem "1"
  
  '-- FECHA:
  lCampos.AddItem "FECHA"
  lTipos.AddItem "DATETIME"
  lAnchos.AddItem " "
  
  '-- CONTADOR:
  lCampos.AddItem "CONTADOR"
  lTipos.AddItem "INTEGER"
  lAnchos.AddItem ""
 
  '-- CREACION:
  lCampos.AddItem "CREACION"
  lTipos.AddItem "DATETIME"
  lAnchos.AddItem " "
  
  
  
End Sub

Private Sub CargarCamposPre()  '--Predeterminados
  aCamposPre(0) = "CEDULA"
  aCamposPre(1) = "NOMBRE"
  aCamposPre(2) = "CARGO"
  aCamposPre(3) = "VENCE"
  aCamposPre(4) = "FOTO"
  aCamposPre(5) = "TIENE_FOTO"
  aCamposPre(6) = "MARCA"
  aCamposPre(7) = "FECHA"
  aCamposPre(8) = "CONTADOR"
End Sub

Private Function EsCampoPre(sCualCampo As String) As Boolean
  Dim i As Integer
  Dim ECP As Boolean
  ECP = False
  i = 0
  Do While (i < MAXCAMPOSPRE) And Not ECP
    If aCamposPre(i) = sCualCampo Then ECP = True Else i = i + 1
  Loop
  EsCampoPre = ECP
End Function


Private Sub bCrearTabla_Click()
  Dim r As New ADODB.Recordset
  Dim cp As String, SC As String
  Dim s As String, sComando As String, sTablaNueva As String
  Dim i As Integer, iNumeroTablaNueva As Integer
  Dim sCampo As String, sTipo As String, sAncho As String
  Dim sComando2 As String
  
  cp = Trim(Mid(cCP.Text, 1, 6))
  SC = Trim(Mid(cSC.Text, 1, 6))
  
  If Trim(SC) = "" Then SC = "0"
  
  If lCampos.ListCount <= 0 Then
    MsgBox "Debe Indicar los campos de la Tabla para poder Crearla...", vbCritical, "Información"
    Exit Sub
  End If
  
  'iNumeroTablaNueva = Numero_Siguiente_Tabla_Personas()
  
  'sTablaNueva = cp & "-" & IIf(sc <> "0", sc & "-", "") & CStr(iNumeroTablaNueva)
  
  iNumeroTablaNueva = Numero_Siguiente_Tabla_Personas()
  
  sTablaNueva = cp & "-" & IIf(SC <> "0", SC & "-", "") & CStr(iNumeroTablaNueva)
  
  If MsgBox("¿Está Seguro de CREAR la Tabla Nueva [" & sTablaNueva & "]?", vbQuestion + vbYesNo, "Confirme") = vbNo Then
    Exit Sub
  End If
  
  
  
  sComando = "CREATE TABLE [" & sTablaNueva & "] " & _
             "(ID INT NOT NULL IDENTITY(1,1), "

  'sComando = "CREATE TABLE [" & sTablaNueva & "] ("
             

  '--Lleva <H>istorico
  sComando2 = "CREATE TABLE [H" & sTablaNueva & "] " & _
              "(ID INT NOT NULL, "
              
  'sComando2 = "CREATE TABLE [H" & sTablaNueva & "] ("
              
              
  For i = 0 To lCampos.ListCount - 1
  
    sCampo = lCampos.List(i)
    sTipo = lTipos.List(i)
    sAncho = lAnchos.List(i)
    
    If sTipo = "CHAR" Then  '--Lleva ancho
      s = sCampo & " " & sTipo & "(" & sAncho & ") "
    Else
      s = sCampo & " " & sTipo & " "
    End If
    
    If i <> lCampos.ListCount - 1 Then s = s & ","
    
    sComando = sComando & s
    sComando2 = sComando2 & s
    
  Next i
  
  sComando = sComando & ", PRIMARY KEY(ID))"
  'sComando = sComando & " )"
  
  'sComando2 = sComando2 & " PRIMARY KEY(ID))"
  sComando2 = sComando2 & " )"
  
  On Error Resume Next
  
  '-- Crear Tabla de Datos <Personas> que conectará con Card-5:
  Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
  Modulo.DBComandoSQL.CommandText = sComando
  Modulo.DBComandoSQL.Execute
  
  If Err.Number <> 0 Then
    MsgBox "Error: No se pudo Crear la Tabla " & sTablaNueva & vbCrLf & Err.Description, vbCritical, "Información"
  Else
    '-- Crear Tabla de Datos <Personas> tipo Historico de Auditoria que
    '-- actualizará el Card-5 y mediante un Trigger desde SQL-Server:
    Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
    Modulo.DBComandoSQL.CommandText = sComando2
    Modulo.DBComandoSQL.Execute
    If Err.Number <> 0 Then
      MsgBox "Error: No se pudo Crear la Tabla H" & sTablaNueva & vbCrLf & Err.Description, vbCritical, "Información"
    Else
      '--Agregar TRIGGER en tabla nueva SQL Server:
      'sComando = "CREATE TRIGGER [TRG_" & sTablaNueva & "] ON [" & sTablaNueva & "] " & vbCrLf & _
                 "FOR UPDATE AS " & vbCrLf & _
                 "SET IDENTITY_INSERT [TRG_" & sTablaNueva & "] ON " & _
                 "INSERT INTO [H" & sTablaNueva & "] SELECT * FROM DELETED"

      'sComando = "CREATE TRIGGER [TRG_" & sTablaNueva & "] ON [" & sTablaNueva & "] " & vbCrLf & _
                 "FOR UPDATE AS " & vbCrLf & _
                 "INSERT INTO [H" & sTablaNueva & "] SELECT * FROM DELETED"
                 
      ''sComando = "CREATE TRIGGER [TRG_" & sTablaNueva & "] ON [" & sTablaNueva & "] " & vbCrLf & _
                 "FOR UPDATE AS " & vbCrLf & _
                 "BEGIN " & vbCrLf & _
                 "  declare @CodProducto as char(20) " & vbCrLf & _
                 "  declare @PrecioProducto as float " & vbCrLf & _
                 "  declare @IDPersona as integer " & vbCrLf & _
                 "  --INSERT INTO [H" & sTablaNueva & "] SELECT * FROM DELETED " & vbCrLf & _
                 "  if update(CONTADOR) " & vbCrLf & _
                 "  begin " & vbCrLf & _
                 "    INSERT INTO [H" & sTablaNueva & "] SELECT * FROM DELETED " & vbCrLf & _
                 "    set @IDPersona = (SELECT ID FROM DELETED) " & vbCrLf & _
                 "    set @CodProducto = '' " & vbCrLf & _
                 "    set @PrecioProducto = 0.00 " & vbCrLf & _
                 "    set @CodProducto = (select codigoproductopvc from opciones) " & vbCrLf & _
                 "    if rtrim(ltrim(@CodProducto)) <> '' " & vbCrLf & _
                 "    begin " & vbCrLf & _
                 "      Set @PrecioProducto = (Select Precio From PreciosEspeciales Where Cliente = " & cp & " And SubCliente = " & SC & " And CodigoProducto = @CodProducto)" & vbCrLf & _
                 "      if @PrecioProducto is null " & vbCrLf & _
                 "      begin " & vbCrLf & _
                 "        Set @PrecioProducto = (Select Precio From Productos Where Codigo = @CodProducto) " & vbCrLf & _
                 "      end " & vbCrLf & _
                 "      update Productos set existencia = existencia - 1 where Codigo = @CodProducto " & vbCrLf
      ''If SC = "0" Then 'Es SOLO cliente:
      ''sComando = sComando & _
                 "      update clientes set deuda = deuda + @PrecioProducto where codigo = " & cp & " " & vbCrLf & _
                 "      update clientes set saldo = deuda - pagos where codigo = " & cp & " " & vbCrLf
      ''Else             'Es sub-cliente:
      ''sComando = sComando & _
                 "      update subclientes set deuda = deuda + @PrecioProducto where cliente = " & cp & " and id = " & SC & " " & vbCrLf & _
                 "      update subclientes set saldo = deuda - pagos           where cliente = " & cp & " and id = " & SC & " " & vbCrLf
      ''End If
      ''sComando = sComando & _
                 "      insert into [EventosC5] (procesado,idtabla,tabla) values ('N',@IDPersona,'" & sTablaNueva & "')" & vbCrLf & _
                 "    end " & vbCrLf & _
                 "  end " & vbCrLf & _
                 "END" & vbCrLf
                 
      ''Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
      ''Modulo.DBComandoSQL.CommandText = sComando
      ''Modulo.DBComandoSQL.Execute
      sComando = "Exec CrearTrigger '" & sTablaNueva & "','" & Mid(cCP.Text, 1, 6) & "','" & SC & "'"
      Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
      Modulo.DBComandoSQL.CommandText = sComando
      Modulo.DBComandoSQL.Execute
 
    
      If Err.Number <> 0 Then
        MsgBox "Error en TRIGGER: " & Err.Description, vbCritical, "Información"
      Else
        '--Agregar Tabla en control del cliente:
        sComando = "INSERT INTO Personas (cliente,subcliente,tabla,creacion) VALUES (" & cp & "," & SC & ",'" & sTablaNueva & "','" & Format(Now, "yyyyMMdd HH:mm:ss") & "')"
        Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
        Modulo.DBComandoSQL.CommandText = sComando
              Clipboard.Clear
      Clipboard.SetText sComando

        Modulo.DBComandoSQL.Execute
    
        If Err.Number <> 0 Then
          MsgBox "Error: " & Err.Description, vbCritical, "Información"
        Else
          MsgBox "Tabla " & sTablaNueva & " creada Exitosamente...", vbInformation, "Información"
        End If
        
        AgregarLogs "Crear Tabla Personas [" & sTablaNueva & "]"
        
        
      End If
    End If
  End If
  
  Cargar_Tablas_Personas

End Sub

Private Sub bCrear_Click()
  If Trim(cNuevaTbl.Text) <> "" Then
    CrearTabla_Nueva cNuevaTbl.Text
  End If
End Sub


Private Sub CrearTabla_Nueva(sTablaOrigen As String)
  Dim r As New ADODB.Recordset
  Dim cp As String, SC As String
  Dim s As String, sComando As String, sTablaNueva As String
  Dim i As Integer, iNumeroTablaNueva As Integer
  Dim sCampo As String, sTipo As String, sAncho As String
  Dim sComando2 As String
  
  cp = Trim(Mid(cCP.Text, 1, 6))
  SC = Trim(Mid(cSC.Text, 1, 6))
  
  If Trim(SC) = "" Then SC = "0"
  
  If lCampos.ListCount <= 0 Then
    MsgBox "Debe Indicar los campos de la Tabla para poder Crearla...", vbCritical, "Información"
    Exit Sub
  End If
  
  iNumeroTablaNueva = Numero_Siguiente_Tabla_Personas()
  
  sTablaNueva = cp & "-" & IIf(SC <> "0", SC & "-", "") & CStr(iNumeroTablaNueva)
  
  If MsgBox("¿Está Seguro de CREAR la Tabla Nueva [" & sTablaNueva & "]?", vbQuestion + vbYesNo, "Confirme") = vbNo Then
    Exit Sub
  End If
  
  Load fMensaje
  fMensaje.Label1.Caption = "Creando Tablas, Espere..."
  fMensaje.Show
  DoEvents


  
  sComando = "CREATE TABLE [" & sTablaNueva & "] " & _
             "(ID INT NOT NULL IDENTITY(1,1), "

  'sComando = "CREATE TABLE [" & sTablaNueva & "] ("
             

  '--Lleva <H>istorico
  sComando2 = "CREATE TABLE [H" & sTablaNueva & "] " & _
              "(ID INT NOT NULL, "
              
  'sComando2 = "CREATE TABLE [H" & sTablaNueva & "] ("
              
              
  For i = 0 To lCampos.ListCount - 1
  
    sCampo = List1.List(i)
    sTipo = List2.List(i)
    sAncho = List3.List(i)
    
    If sTipo = "CHAR" Then  '--Lleva ancho
      s = sCampo & " " & sTipo & "(" & sAncho & ") "
    Else
      s = sCampo & " " & sTipo & " "
    End If
    
    If i <> List1.ListCount - 1 Then s = s & ","
    
    sComando = sComando & s
    sComando2 = sComando2 & s
    
  Next i
  
  sComando = sComando & " PRIMARY KEY(ID))"
  'sComando = sComando & " )"
  
  'sComando2 = sComando2 & " PRIMARY KEY(ID))"
  sComando2 = sComando2 & " )"
  On Error Resume Next
  
  '-- Crear Tabla de Datos <Personas> que conectará con Card-5:
  Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
  Modulo.DBComandoSQL.CommandText = sComando
  Modulo.DBComandoSQL.Execute
  sCrearArchivoExcel sTablaNueva, Mid(cCP.Text, 10, Len(cCP.Text))
  If Err.Number <> 0 Then
    MsgBox "Error: No se pudo Crear la Tabla " & sTablaNueva & vbCrLf & Err.Description, vbCritical, "Información"
  Else
    '-- Crear Tabla de Datos <Personas> tipo Historico de Auditoria que
    '-- actualizará el Card-5 y mediante un Trigger desde SQL-Server:
    Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
    Modulo.DBComandoSQL.CommandText = sComando2
    Modulo.DBComandoSQL.Execute
    If Err.Number <> 0 Then
      MsgBox "Error: No se pudo Crear la Tabla H" & sTablaNueva & vbCrLf & Err.Description, vbCritical, "Información"
    Else
      '--Agregar TRIGGER en tabla nueva SQL Server:
      'sComando = "CREATE TRIGGER [TRG_" & sTablaNueva & "] ON [" & sTablaNueva & "] " & vbCrLf & _
                 "FOR UPDATE AS " & vbCrLf & _
                 "SET IDENTITY_INSERT [TRG_" & sTablaNueva & "] ON " & _
                 "INSERT INTO [H" & sTablaNueva & "] SELECT * FROM DELETED"

      'sComando = "CREATE TRIGGER [TRG_" & sTablaNueva & "] ON [" & sTablaNueva & "] " & vbCrLf & _
                 "FOR UPDATE AS " & vbCrLf & _
                 "INSERT INTO [H" & sTablaNueva & "] SELECT * FROM DELETED"
                 
      '"IF UPDATE(fecha) OR UPDATE(contador) " & vbCrLf & _

      
      sComando = "Exec CrearTrigger '" & sTablaNueva & "','" & Mid(cCP.Text, 1, 6) & "','0'"
      Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
      Modulo.DBComandoSQL.CommandText = sComando
      Modulo.DBComandoSQL.Execute
    
      If Err.Number <> 0 Then
        MsgBox "Error en TRIGGER: " & Err.Description, vbCritical, "Información"
      Else
        '--Agregar Tabla en control del cliente:
        sComando = "INSERT INTO Personas (cliente,subcliente,tabla,creacion) VALUES (" & cp & "," & SC & ",'" & sTablaNueva & "','" & Format(Now, "yyyyMMdd HH:mm:ss") & "')"
        Clipboard.Clear
        Clipboard.SetText sComando

        Set Modulo.DBComandoSQL.ActiveConnection = Modulo.DBConexionSQL
        Modulo.DBComandoSQL.CommandText = sComando
        Modulo.DBComandoSQL.Execute
         
        If Err.Number <> 0 Then
          MsgBox "Error: " & Err.Description, vbCritical, "Información"
        Else
        
          's = "insert into [" & sTablaNueva & "] select * from [" & sTablaOrigen & "]"
          'Modulo.ExecSQL s
          Load fMensaje
          fMensaje.Label1.Caption = "Agregando personas a la tabla, Espere..."
          fMensaje.Show
          DoEvents

          Agregar_Todas_Personas sTablaOrigen, sTablaNueva
          Unload fMensaje
          MsgBox "Tabla " & sTablaNueva & " creada Exitosamente...", vbInformation, "Información"
        End If
      End If
    End If
  End If
  
  Cargar_Tablas_Personas

End Sub

Private Sub Agregar_Todas_Personas(sTablaOrigen As String, sTablaDestino As String)
  Dim rO As New ADODB.Recordset, rD As New ADODB.Recordset
  Dim s As String, i As Integer
  
  s = "select * from [" & sTablaOrigen & "]"
  rO.Open s, Modulo.DBConexionSQL, adOpenKeyset, adLockReadOnly
    
  s = "select * from [" & sTablaDestino & "]"
  rD.Open s, Modulo.DBConexionSQL, adOpenDynamic, adLockOptimistic
  
  Do While Not rO.EOF
  
    rD.AddNew
    For i = 0 To rO.Fields.Count - 1
      If rO.Fields(i).Name <> "ID" Then
        rD.Fields(i).value = rO.Fields(i).value
      End If
    Next i
    rD.Update
    
    rO.MoveNext
  Loop
  
  rD.Close
  rO.Close
  Set rD = Nothing
  Set rO = Nothing
End Sub

Private Sub bMas_Click()
  If Trim(eCampo.Text) <> "" Then
    If EsCampoPre(Trim(eCampo.Text)) = True Or UCase(Trim(eCampo.Text)) = "ID" Then
      MsgBox "Utilice otro nombre para el campo [" & Trim(eCampo.Text) & "], ya está siendo usado internamente por Sistema.", vbCritical, "Información"
      Exit Sub
    End If
  
    lCampos.AddItem eCampo.Text
    lTipos.AddItem cTipos.Text
    lAnchos.AddItem eAncho.Text
    
    lCampos.ListIndex = lCampos.ListCount - 1
    eCampo.Text = ""
    cTipos.ListIndex = 0
    eAncho.Text = "20"
    eCampo.SetFocus
  End If
    
End Sub

Private Sub bMenos_Click()
  Dim i As Integer
  i = lCampos.ListIndex
  If i >= 0 Then
    If EsCampoPre(lCampos.List(i)) Then
      MsgBox "El Campo [" & lCampos.List(i) & "] es del Sistema, No se puede Suprimir...", vbCritical, "Información"
    Else
      If MsgBox("¿Está Seguro de Remover el Campo [" & lCampos.List(i) & "]?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
        lCampos.RemoveItem i
        lTipos.RemoveItem i
        lAnchos.RemoveItem i
      End If
    End If
  End If
End Sub

Private Sub bCancelar_Click()
  Unload Me
End Sub

Private Sub bSubir_Click()
  Dim i As Integer
  Dim sAux As String
  i = lCampos.ListIndex
  If i >= 0 Then
    If i = 0 Then 'Ya no Sube Mas...
      Beep
    Else
      sAux = lCampos.List(i - 1)
      lCampos.List(i - 1) = lCampos.List(i)
      lCampos.List(i) = sAux
      lCampos.ListIndex = i - 1
      
      sAux = lTipos.List(i - 1)
      lTipos.List(i - 1) = lTipos.List(i)
      lTipos.List(i) = sAux
      lTipos.ListIndex = i - 1
      
      sAux = lAnchos.List(i - 1)
      lAnchos.List(i - 1) = lAnchos.List(i)
      lAnchos.List(i) = sAux
      lAnchos.ListIndex = i - 1
      
    End If
  End If
End Sub

Private Sub bBajar_Click()
  Dim i As Integer
  Dim sAux As String
  i = lCampos.ListIndex
  If i >= 0 Then
    If i = lCampos.ListCount - 1 Then 'Ya no Baja Mas...
      Beep
    Else
      sAux = lCampos.List(i + 1)
      lCampos.List(i + 1) = lCampos.List(i)
      lCampos.List(i) = sAux
      lCampos.ListIndex = i + 1
      
      sAux = lTipos.List(i + 1)
      lTipos.List(i + 1) = lTipos.List(i)
      lTipos.List(i) = sAux
      lTipos.ListIndex = i + 1
      
      sAux = lAnchos.List(i + 1)
      lAnchos.List(i + 1) = lAnchos.List(i)
      lAnchos.List(i) = sAux
      lAnchos.ListIndex = i + 1
      
      
    End If
  End If
End Sub


Private Sub bVerContenido_Click()
  Dim i As Integer
  Dim s As String
  
  i = -1
  If lTablas.ListIndex >= 0 Then i = lTablas.ListIndex
  If lCreadas.ListIndex >= 0 Then i = lCreadas.ListIndex
  If i < 0 Then
    MsgBox "¿Debe Seleccionar una Tabla para visualizar las Personas...", vbCritical, "Información"
  Else
    s = lTablas.List(i)
    If Modulo.Total_Registros(s) > 0 Then
      Load fVerTablaPersonas
      With fVerTablaPersonas
        .Caption = "Contenido de Tabla:" & s
        .Adodc1.ConnectionString = Modulo.DBConexionSQL.ConnectionString
        .Adodc1.RecordSource = "select * from [" & s & "]"
        .Adodc1.Refresh
        .DataGrid1.Refresh
        .Label2.Caption = CStr(Modulo.Total_Registros(s))
      End With
      fVerTablaPersonas.Show vbModal
    End If
  End If
End Sub

Private Sub cCP_Click()
  cSC.BackColor = &HC0FFFF
  If Trim(cCP.Text) <> "" Then Cargar_SubClientes
  Cargar_Tablas_Personas
End Sub


Private Sub Check1_Click()
  If Check1.value = vbChecked Then
    cSC.Clear
    Cargar_Tablas_Personas
  Else
    cCP_Click
  End If
End Sub

Private Sub cSC_Change()
  Cargar_Tablas_Personas
End Sub

Private Sub cSC_Click()
  Cargar_Tablas_Personas
End Sub

Private Sub cTipos_Click()
  eAncho.Text = ""
  eAncho.Enabled = False
  If cTipos.Text = "CHAR" Then
    eAncho.Text = "20"
    eAncho.Enabled = True
  End If
End Sub

Private Sub eAncho_GotFocus()
  eAncho.SelStart = 0
  eAncho.SelLength = Len(eAncho.Text)
End Sub

Private Sub eAncho_KeyPress(KeyAscii As Integer)
  If KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0
    End If
  End If
End Sub

Private Sub eAncho_LostFocus()
  If Trim(eAncho.Text) = "" Then eAncho.Text = "1"
  If Not IsNumeric(eAncho.Text) Then
    MsgBox "Debe Introducir Ancho del Campo en Números (Ejm. 25), Revise...", vbCritical, "Información"
    eAncho.SetFocus
  End If
End Sub

Private Sub eCampo_Change()
  EnMayusculas eCampo
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

Sub Cargar_Tablas_Personas()
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim SC As String, sSC As String
  
  lTablas.Clear
  lCreadas.Clear
  
  cNuevaTbl.Clear
  
  SC = Trim(Mid(cCP.Text, 1, 6))
  sSC = Trim(Mid(cSC.Text, 1, 6))
  If sSC = "" Then sSC = "0"
  
  If SC = "" Then Exit Sub
  
  s = "select * from Personas where " & _
      "cliente    = " & SC & " and " & _
      "subcliente = " & sSC & " order by id "

  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  Do While Not r.EOF
    lTablas.AddItem Trim(r.Fields("Tabla").value)
    lCreadas.AddItem Format(r.Fields("creacion").value, "dd/mm/yyyy")
    
    cNuevaTbl.AddItem Trim(r.Fields("Tabla").value)
    r.MoveNext
  Loop
  r.Close
  Set r = Nothing
  
  If cNuevaTbl.ListCount > 0 Then cNuevaTbl.ListIndex = 0
End Sub

Function Numero_Siguiente_Tabla_Personas() As Integer
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim SC As String, sSC As String
  Dim k As Integer, n As Integer
    
  SC = Trim(Mid(cCP.Text, 1, 6))
  sSC = Trim(Mid(cSC.Text, 1, 6))
  If sSC = "" Then sSC = "0"
  
  s = "select * from Personas where " & _
      "cliente    = " & SC & " and " & _
      "subcliente = " & sSC & " order by id "
      
  k = 0

  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  Do While Not r.EOF
    s = Trim(r.Fields("tabla").value)
    s = Trim(Mid(s, InStrRev(s, "-") + 1))
    k = CInt(s)
    r.MoveNext
  Loop
  r.Close
  Set r = Nothing
  Numero_Siguiente_Tabla_Personas = k + 1
End Function

Private Sub Cargar_Clientes()
  'Dim r As New ADODB.Recordset
  Dim s As String
  Dim l As Integer
  Dim sCod As String
  
  If RClientes.State <> adStateClosed Then RClientes.Close
  
  s = "SELECT * FROM Clientes ORDER BY Codigo"
  
  RClientes.Open s, DBConexionSQL, adOpenKeyset, adLockReadOnly
  
  cCP.Clear
  
  Do While Not RClientes.EOF
    s = Zeros(RClientes.Fields("codigo").value, 6) & " : " & Trim(RClientes.Fields("nombre").value)
    cCP.AddItem s
    RClientes.MoveNext
  Loop
  
  RClientes.Close
End Sub


Private Sub Cargar_SubClientes()
  'Dim r As New ADODB.Recordset
  Dim s As String
  Dim l As Integer
  Dim sCod As String
  
  cSC.Clear
  
  Viendo = False
  
  If RSubClientes.State <> adStateClosed Then RSubClientes.Close
  
  sCod = "000000"
  If Trim(cCP.Text) <> "" Then sCod = Mid(cCP.Text, 1, 6)
      
  s = "SELECT * FROM SubClientes WHERE Cliente = " & sCod & " ORDER BY Id"
  
  RSubClientes.Open s, DBConexionSQL, adOpenDynamic, adLockOptimistic
  l = 1
  If RSubClientes.EOF = False Then
    cSC.BackColor = &H80FFFF
    Do While Not RSubClientes.EOF
      s = Zeros(RSubClientes.Fields("id").value, 6) & " : " & Trim(RSubClientes.Fields("nombre").value)
      cSC.AddItem s
      RSubClientes.MoveNext
    Loop
  Else
     cSC.BackColor = &HC0FFFF
  End If
  RSubClientes.Close
  ''cSC.BackColor = &HC0FFFF  &H00C0FFFF&  &H0080FFFF&
  
  cSC.ListIndex = -1
  ''If cSC.ListCount > 0 Then cSC.ListIndex = 0
  
End Sub


Private Sub Form_Load()
  CargarCamposPre
  
  eCampo.Text = ""
  eAncho.Text = ""
  
  cTipos.Clear
  cTipos.AddItem "CHAR"
  cTipos.AddItem "INTEGER"
  cTipos.AddItem "DATETIME"
  cTipos.AddItem "FLOAT"
  cTipos.ListIndex = 0

  
  
  Cargar_Tablas_Personas
  Cargar_Clientes
  Cargar_SubClientes
  
  cNuevaTbl.Clear
  
End Sub

Private Sub lAnchos_Click()
  If lAnchos.ListIndex >= 0 Then
    lCampos.ListIndex = lAnchos.ListIndex
    lTipos.ListIndex = lAnchos.ListIndex
  End If
End Sub

Private Sub lCampos_Click()
  If lCampos.ListIndex >= 0 Then
    lTipos.ListIndex = lCampos.ListIndex
    lAnchos.ListIndex = lCampos.ListIndex
  End If
End Sub

Private Sub lCreadas_Click()
  If lCreadas.ListIndex >= 0 Then
    lTablas.ListIndex = lCreadas.ListIndex
    Mostrar_Estructura
  End If
End Sub

Private Sub lTablas_Click()
  If lTablas.ListIndex >= 0 Then
    lCreadas.ListIndex = lTablas.ListIndex
    Mostrar_Estructura
  End If
    
End Sub

Private Sub Mostrar_Estructura()
  Dim r As New ADODB.Recordset
  Dim s As String, s1 As String
  Dim i As Integer
  
  If lTablas.ListIndex >= 0 Then
  
    s = lTablas.List(lTablas.ListIndex)
    s = "select count(*) from [" & s & "]"
    r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    Label26.Caption = "-"
    If Not r.EOF Then Label26.Caption = CStr(r.Fields(0).value)
    r.Close
  
    s = lTablas.List(lTablas.ListIndex)
    s = "select * from [" & s & "]"
    r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    '--Extraer toda la info de la tabla estructura:
    lCampos.Clear: List1.Clear
    lTipos.Clear: List2.Clear
    lAnchos.Clear: List3.Clear
    
    For i = 0 To r.Fields.Count - 1
      If UCase(r.Fields(i).Name) <> "ID" Then
        lCampos.AddItem r.Fields(i).Name
        List1.AddItem r.Fields(i).Name
        
        s = ""
        Select Case r.Fields(i).Type
          Case adChar: s = "CHAR"
          Case adInteger: s = "INTEGER"
          Case adDBTimeStamp: s = "DATETIME"
          Case adDouble: s = "FLOAT"
        End Select
        lTipos.AddItem s
        List2.AddItem s
        
        s1 = ""
           
        If s = "CHAR" Then s1 = CStr(r.Fields(i).DefinedSize)
        
        lAnchos.AddItem s1
        List3.AddItem s1
      End If
    Next i
    r.Close
    Set r = Nothing
  End If
    
End Sub

Private Sub Cargar_Estructura(sTabla As String)
  Dim r As New ADODB.Recordset
  Dim s As String, s1 As String
  Dim i As Integer
  
  If sTabla <> "" Then
  
    s = "select count(*) from [" & sTabla & "]"
    r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    Label26.Caption = "-"
    If Not r.EOF Then Label26.Caption = CStr(r.Fields(0).value)
    r.Close
  
    s = "select * from [" & sTabla & "]"
    r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
    '--Extraer toda la info de la tabla estructura:
    List1.Clear
    List2.Clear
    List3.Clear
    For i = 0 To r.Fields.Count - 1
      If UCase(r.Fields(i).Name) <> "ID" Then
        List1.AddItem r.Fields(i).Name
        s = ""
        Select Case r.Fields(i).Type
          Case adChar: s = "CHAR"
          Case adInteger: s = "INTEGER"
          Case adDBTimeStamp: s = "DATETIME"
          Case adDouble: s = "FLOAT"
        End Select
        List2.AddItem s
        s1 = ""
           
        If s = "CHAR" Then s1 = CStr(r.Fields(i).DefinedSize)
        
        List3.AddItem s1
      End If
    Next i
    r.Close
    Set r = Nothing
  End If
    
End Sub




Private Sub lTipos_Click()
  If lTipos.ListIndex >= 0 Then
    lCampos.ListIndex = lTipos.ListIndex
    lAnchos.ListIndex = lTipos.ListIndex
  End If
End Sub
