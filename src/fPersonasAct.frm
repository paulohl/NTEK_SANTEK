VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form fPersonasAct 
   Caption         =   "Introducir / Actualizar Personas"
   ClientHeight    =   10200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10200
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   11880
      TabIndex        =   39
      Top             =   120
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.ListBox ListArchivos 
      Height          =   645
      ItemData        =   "fPersonasAct.frx":0000
      Left            =   13680
      List            =   "fPersonasAct.frx":0002
      TabIndex        =   38
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Cliente Principal posee Tabla de Personas"
      Height          =   225
      Left            =   7590
      TabIndex        =   35
      Top             =   540
      Width           =   3825
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   525
      Left            =   10890
      TabIndex        =   34
      Top             =   210
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   926
      _Version        =   393216
   End
   Begin VB.Frame Frame6 
      Height          =   9255
      Left            =   0
      TabIndex        =   5
      Top             =   930
      Width           =   15195
      Begin VB.CommandButton Command2 
         Caption         =   "Toma de Fotos en Sitio"
         Height          =   495
         Left            =   10560
         TabIndex        =   40
         Top             =   540
         Width           =   1395
      End
      Begin VB.CommandButton bCard5 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Card-5"
         Height          =   465
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   690
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Frame Frame1 
         Caption         =   "Personas Registradas"
         Height          =   7455
         Left            =   90
         TabIndex        =   20
         Top             =   1740
         Width           =   15045
         Begin VB.CommandButton cmdExportar 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Exportar"
            Height          =   375
            Left            =   9780
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Exportar Datos a Excel"
            Top             =   6960
            Width           =   1000
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Archivos Huérfanos"
            Height          =   495
            Left            =   10920
            TabIndex        =   37
            ToolTipText     =   "Buscar Archivos huérfanos en la carpeta de fotos del cliente"
            Top             =   6840
            Width           =   1035
         End
         Begin VB.CommandButton cmdProcesarCarnets 
            Caption         =   "Procesar Carnets"
            Height          =   495
            Left            =   13020
            TabIndex        =   36
            Top             =   6840
            Width           =   915
         End
         Begin VB.CommandButton bAuditarFotos 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Comprobar Fotos"
            Height          =   500
            Left            =   12000
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   6840
            Width           =   915
         End
         Begin VB.CommandButton bCancelar 
            Caption         =   "Salir"
            Height          =   500
            Left            =   14040
            Picture         =   "fPersonasAct.frx":0004
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   6840
            UseMaskColor    =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton bBorrarTodo 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Borrar TODO"
            Height          =   345
            Left            =   8640
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   6480
            Width           =   1065
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "fPersonasAct.frx":058E
            Height          =   6165
            Left            =   60
            TabIndex        =   29
            Top             =   180
            Width           =   14865
            _ExtentX        =   26220
            _ExtentY        =   10874
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton bRefrescar 
            Height          =   345
            Left            =   4260
            Picture         =   "fPersonasAct.frx":05A3
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Refrescar"
            Top             =   6510
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.ComboBox cOrden 
            Height          =   315
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   6540
            Width           =   1215
         End
         Begin VB.CommandButton bBuscarPersona 
            Height          =   345
            Left            =   4740
            Picture         =   "fPersonasAct.frx":0FA5
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   6510
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton bImportar 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Importar"
            Height          =   375
            Left            =   9780
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   6480
            Width           =   1000
         End
         Begin VB.CommandButton bBorrar 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Borrar"
            Height          =   345
            Left            =   7560
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   6480
            Width           =   945
         End
         Begin VB.CommandButton bEditar 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Editar"
            Height          =   345
            Left            =   6420
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   6480
            Width           =   1000
         End
         Begin VB.CommandButton bAgregar 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Agregar"
            Height          =   345
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   6480
            Width           =   1000
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   90
            Top             =   6510
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   582
            ConnectMode     =   3
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   2
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Personas"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ordenar por:"
            Height          =   195
            Left            =   2070
            TabIndex        =   26
            Top             =   6570
            Width           =   885
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Estructura Tabla de Personas"
         Height          =   1605
         Left            =   4590
         TabIndex        =   10
         Top             =   150
         Width           =   5205
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
            Height          =   735
            Left            =   4020
            TabIndex        =   15
            Top             =   480
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
            Height          =   735
            Left            =   2250
            TabIndex        =   13
            Top             =   480
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
            Height          =   735
            Left            =   90
            TabIndex        =   11
            Top             =   480
            Width           =   2085
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Anchos"
            Height          =   255
            Left            =   4020
            TabIndex        =   16
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tipos"
            Height          =   255
            Left            =   2250
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Campos"
            Height          =   255
            Left            =   90
            TabIndex        =   12
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tablas Personas Creadas"
         Height          =   1605
         Left            =   60
         TabIndex        =   6
         Top             =   150
         Width           =   4455
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
            Height          =   735
            Left            =   2760
            TabIndex        =   17
            Top             =   480
            Width           =   1600
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
            Height          =   735
            Left            =   90
            TabIndex        =   7
            Top             =   480
            Width           =   2385
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
            TabIndex        =   19
            Top             =   1320
            Width           =   75
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "REGS:"
            Height          =   195
            Left            =   2790
            TabIndex        =   18
            Top             =   1320
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
      Top             =   90
      Width           =   5715
   End
   Begin VB.CommandButton bBuscarCP 
      Height          =   345
      Left            =   7560
      Picture         =   "fPersonasAct.frx":10A7
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
Attribute VB_Name = "fPersonasAct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const FGC = "Código    | Nombre                                                          | RIF Nº                | Telefonos                                "

Dim RClientes As New ADODB.Recordset
Dim RSubClientes As New ADODB.Recordset

Dim Viendo As Boolean

Dim OPR As Integer  'operacion: 1 nuevo  2 editar

Public vRUTAINICIAL As String

Private Sub bAgregar_Click()
  Dim i As Integer
  Dim s As String
  
  vID = -1
    
  Load fPersonasReg
  
  Modulo.vIndiceFoto = -1
  
  With fPersonasReg
    For i = 0 To Me.lCampos.ListCount - 1
      'If UCase(Me.lCampos.List(i)) <> "ID" Then
        .Labels(i).Caption = Me.lCampos.List(i)
            
        Select Case Me.lTipos.List(i)
          Case "CHAR"
            .Texts(i).Enabled = True
            .Texts(i).MaxLength = CInt(Me.lAnchos.List(i))
            Select Case Me.lCampos.List(i)
              Case "CEDULA"
                '--Tratamiento especial: alineacion, teclas, formateo, puntos
                .Texts(i).Alignment = vbRightJustify
                '.Texts(i).MaxLength = CInt(Me.lAnchos.List(i))
              Case "CREACION"
                .Texts(i).Text = Format(Now, "dd/mm/yyyy hh:mm:ss ampm")
              Case "FOTO"
                Modulo.vIndiceFoto = i
              Case "TIENE_FOTO"
                .Texts(i).Text = "N"
            End Select
          
          Case "DATETIME"
            Select Case Me.lCampos.List(i)
              Case "CREACION"
                .Texts(i).Text = Format(Now, "dd/mm/yyyy hh:mm:ss ampm")
            End Select
        End Select
      'End If
    Next i

    Modulo.fModalResult = ""
    .Show vbModal
    
    '-- Proceso de guardado del registro:
    If Modulo.fModalResult = fModalResultOK Then
    
      Load fMensaje
      fMensaje.Show
      DoEvents
      
      Adodc1.Recordset.AddNew
      s = ""
    
      With fPersonasReg
        For i = 0 To Me.lCampos.ListCount - 1
          If UCase(Me.lCampos.List(i)) <> "ID" Then
            If Me.lTipos.List(i) = "CHAR" Then
              s = Trim(.Texts(i).Text)
              Adodc1.Recordset.Fields(.Labels(i).Caption).value = s
            Else
              If Me.lTipos.List(i) = "DATETIME" Then
                If UCase(Me.lCampos.List(i)) = "CREACION" Then
                  Adodc1.Recordset.Fields(.Labels(i).Caption).value = Now
                End If
              End If
            End If
          End If
        Next i
      End With
      
      Adodc1.Recordset.Update
      
      If lTablas.ListIndex >= 0 Then Label26.Caption = CStr(Modulo.Total_Registros(lTablas.List(lTablas.ListIndex)))
      
      
      
      AgregarLogs "Agrega Persona en [" & IIf(lTablas.ListIndex >= 0, lTablas.List(lTablas.ListIndex), "-") & "]"
            
      Unload fMensaje
      
    End If
  
  End With
  
  Unload fPersonasReg
  
  If lTablas.ListIndex >= 0 Then
    Auditar_Fotos lTablas.List(lTablas.ListIndex)
  End If
  
End Sub

Public Sub bAuditarFotos_Click()
  If lTablas.ListIndex >= 0 Then
    Auditar_Fotos lTablas.List(lTablas.ListIndex)
    Mostrar_Personas_Registradas
    bRefrescar_Click
  End If
End Sub

Private Sub bBorrar_Click()
  Dim i As Integer
  Dim s As String
  Dim s1 As String
  Dim sID As Long
  
  If lTablas.ListIndex >= 0 Then
    s = lTablas.List(lTablas.ListIndex)
    If Modulo.Total_Registros(s) > 0 Then
      sID = Adodc1.Recordset.Fields("ID").value
      s = Trim(Adodc1.Recordset.Fields("CEDULA").value) & " / " & Trim(Adodc1.Recordset.Fields("NOMBRE").value)
      
      If MsgBox("¿Está Seguro de Borrar el Registro de la Persona [" & s & "]?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
        s1 = lTablas.List(lTablas.ListIndex)
        s = "DELETE FROM [" & s1 & "] WHERE ID = " & sID & " "
        Modulo.ExecSQL s
        Adodc1.Refresh
        DataGrid1.Refresh
        If lTablas.ListIndex >= 0 Then Label26.Caption = CStr(Modulo.Total_Registros(lTablas.List(lTablas.ListIndex)))
        
        AgregarLogs "Borra Persona en [" & s1 & "]"
        
      End If
    End If
  End If

End Sub

Private Sub bBorrarTodo_Click()
  Dim s As String
  If MsgBox("¿Está Seguro de Borrar TODOS los registros de este Cliente?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
    If lTablas.ListIndex >= 0 Then
      s = lTablas.List(lTablas.ListIndex)
      If s <> "" Then
        s = "delete from [" & s & "]"
        Modulo.ExecSQL (s)
        MsgBox "Registros Borrados Exitosamente...", vbInformation, "Información"
        Adodc1.Refresh
        DataGrid1.Refresh
        
        AgregarLogs "Borra TODO Persona en [" & s & "]"
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

Private Sub bBuscarPersona_Click()
   frmBuscar.lTabla = "[" & lTablas.List(lTablas.ListIndex) & "]"
   frmBuscar.Show vbModal
End Sub

Private Sub bCancelar_Click()
  Unload Me
End Sub

Private Sub bEditar_Click()
  Dim i As Integer
  Dim s As String
 
  Load fPersonasReg
  
  Modulo.vIndiceFoto = -1
  
  Modulo.vID = Adodc1.Recordset.Fields("id").value
  
  With fPersonasReg
    .Texts(0).Text = Adodc1.Recordset.Fields("id").value
        
    For i = 0 To Me.lCampos.ListCount - 1
      'If UCase(Me.lCampos.List(i)) <> "ID" Then
        .Labels(i).Caption = Me.lCampos.List(i)
                   
        Select Case Me.lTipos.List(i)
          Case "CHAR"
            'If Me.lCampos.List(i) = "CEDULA" Then .Texts(i).Enabled = False Else .Texts(i).Enabled = True
            If IsNull(Adodc1.Recordset.Fields(.Labels(i).Caption).value) Then
              .Texts(i).Text = ""
            Else
              .Texts(i).Text = Trim(Adodc1.Recordset.Fields(.Labels(i).Caption).value)
            End If
            .Texts(i).Enabled = True
            .Texts(i).MaxLength = CInt(Me.lAnchos.List(i))
            Select Case Me.lCampos.List(i)
              Case "CEDULA"
                '--Tratamiento especial: alineacion, teclas, formateo, puntos
                .Texts(i).Alignment = vbRightJustify
                '.Texts(i).Enabled = False
                
                '.Texts(i).MaxLength = CInt(Me.lAnchos.List(i))
              Case "CREACION"
                .Texts(i).Text = Format(Adodc1.Recordset.Fields("CREACION").value, "dd/mm/yyyy hh:mm:ss ampm")
              Case "FOTO"
                Modulo.vIndiceFoto = i
              Case "TIENE_FOTO"
                .Texts(i).Text = IIf(IsNull(Adodc1.Recordset.Fields(.Labels(i).Caption).value), "N", Adodc1.Recordset.Fields(.Labels(i).Caption).value)
            End Select
          
          Case "DATETIME"
            Select Case Me.lCampos.List(i)
              Case "CREACION"
                .Texts(i).Text = Format(Adodc1.Recordset.Fields("CREACION").value, "dd/mm/yyyy hh:mm:ss ampm")
            End Select
          
        .Texts(i).Visible = True
        End Select
      'End If
    Next i

    Modulo.fModalResult = ""
    
    .Show vbModal
    
    '-- Proceso de guardado del registro en edicion:
    If Modulo.fModalResult = fModalResultOK Then
    
      Load fMensaje
      fMensaje.Show
      DoEvents
      
      'Adodc1.Recordset.AddNew
      s = ""
    
      With fPersonasReg
        For i = 0 To Me.lCampos.ListCount - 1
          If UCase(Me.lCampos.List(i)) <> "ID" Then
            If Me.lTipos.List(i) = "CHAR" Then
              s = UCase(Trim(.Texts(i).Text))
              Adodc1.Recordset.Fields(.Labels(i).Caption).value = s
              
            Else
              If Me.lTipos.List(i) = "DATETIME" Then
                If UCase(Me.lCampos.List(i)) = "CREACION" Then
                  Adodc1.Recordset.Fields(.Labels(i).Caption).value = Now
                End If
              End If
            End If
          End If
        Next i
      End With
      
      Adodc1.Recordset.Update
      
      AgregarLogs "Edita Persona en [" & IIf(lTablas.ListIndex >= 0, lTablas.List(lTablas.ListIndex), "-") & "]"
      
      Unload fMensaje
      
    End If
  
  End With
  
  Unload fPersonasReg
  
  DataGrid1.Refresh
  
  
  
  
  
  
  
End Sub

Private Sub bImportar_Click()
  Dim a As String, s1 As String, s2 As String, s3 As String
  
  If lTablas.ListIndex >= 0 Then
  
    s1 = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
    s2 = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
    
    a = "LISTADO " & Trim(Mid(fPersonasAct.cCP.Text, 9))
  
    If Trim(fPersonasAct.cSC.Text) <> "" Then
      a = a & "_" & Trim(Mid(fPersonasAct.cSC.Text, 9))
    End If
  
    a = a & "_" & Format(Date, "dd") & "_" & _
                  Format(Date, "mm") & "_" & _
                  Format(Date, "yy") & "_OFICINA.XLS"
  
    s3 = s2 & "\" & Trim(Mid(fPersonasAct.cCP.Text, 9)) & "\" & _
                    IIf(Trim(fPersonasAct.cSC.Text) <> "", Trim(Mid(fPersonasAct.cSC.Text, 9)) & "\", "")
                    
    Modulo.vRUTAINICIAL = s3
    
    Load fImportar
    
    fImportar.Text1 = ""
    fImportar.Show
    
    AgregarLogs "Inicia Importacion de Personas"
    
  Else
    MsgBox "Debe Seleccionar la Tabla de Personas que desea Gestionar.", vbCritical, "Información"
  End If
End Sub

Sub RefrescarPersonas()
  Dim s As String, c As String
  Dim i As Integer
  i = lTablas.ListIndex
  If i >= 0 Then
    c = cOrden.Text
    If Trim(c) <> "" Then
      s = "SELECT * FROM [" & lTablas.List(i) & "]" & IIf(c <> "", " ORDER BY " & c, "")
      Adodc1.ConnectionString = Modulo.DBConexionSQL.ConnectionString
      Adodc1.RecordSource = s
      Adodc1.Refresh
      DataGrid1.Refresh
    End If
  Else
    MsgBox "Debe Seleccionar la Tabla de Personas...", vbCritical, "Información"
  End If
End Sub

Private Sub bRefrescar_Click()
  RefrescarPersonas
End Sub

Private Sub cCP_Click()
  cSC.BackColor = &HC0FFFF
  If Trim(cCP.Text) <> "" Then Cargar_SubClientes
  Cargar_Tablas_Personas
  
  
'  If Trim(cCP.Text) <> "" Then Cargar_SubClientes
 ' Cargar_Tablas_Personas
End Sub


Private Sub Check1_Click()
  If Check1.value = vbChecked Then
    cSC.Clear
    Cargar_Tablas_Personas
  Else
    cCP_Click
  End If
End Sub

Private Sub cmdExportar_Click()
   Dim lItem As ListItem
   Dim i As Integer
   On Error GoTo falla
   If fPersonasAct.Adodc1.Recordset.RecordCount <> 0 Then
      DataGrid1.Row = 0
     
      For i = 0 To DataGrid1.Columns.Count - 1
         Set lItem = frmColumnasExportar.ListVExportar.ListItems.Add(, , DataGrid1.Columns(i).Caption)
            lItem.SubItems(1) = i
      Next i
      frmColumnasExportar.Show vbModal
   Else
      MsgBox "No Existen datos para exportar", vbExclamation
   End If
falla:
   If Err.Number <> 0 Then
     Select Case Err.Number
       Case 91
           MsgBox "No Existen datos para exportar", vbExclamation
       Case Else
          MsgBox Err.Number & "::" & Err.Description, vbCritical
     End Select
   End If
End Sub

Private Sub cmdProcesarCarnets_Click()
Dim dTC As Double, dTA As Double, dTS As Double
  On Error GoTo falla
If Adodc1.Recordset.RecordCount <= 0 Then
   MsgBox "No existe ningún Registro que procesar", vbExclamation
   Exit Sub
End If
  'frmVencimiento.Show vbModal
  'If Modulo.fModalResult = Modulo.fModalResultOK Then
   Modulo.vTemporal1 = ""
             
   Modulo.vTemporal2 = cCP.Text
   Modulo.vTemporal3 = cSC.Text
                    
   Modulo.fModalResult = Modulo.fModalResultCANCEL
   Load fObservaciones
        
   fObservaciones.lCliente.Caption = Modulo.vTemporal2
   fObservaciones.lSubCliente.Caption = Modulo.vTemporal3
    
     Modulo.Resumen_Cuenta_Cliente cCP.Text, cSC.Text, dTC, dTA, dTS

     fObservaciones.lTC.Caption = Format(dTC, "#,0.00")
     fObservaciones.lTA.Caption = Format(dTA, "#,0.00")
     fObservaciones.lTS.Caption = Format(dTS, "#,0.00")
     fObservaciones.lCliente.Caption = cCP.Text
     fObservaciones.lSubCliente.Caption = cSC.Text
     fObservaciones.EsActualizar = False
     
     fObservaciones.EsGuardarPorLotes = True
     Load fObservaciones
     fObservaciones.Caption = "Procesar Carnets Por Lotes"
     fObservaciones.Show
     'fObservaciones.CargarProductos
falla:
     If Err.Number <> 0 Then
        Select Case Err.Number
           Case 91
              MsgBox "No existen registros que procesar", vbExclamation
              Exit Sub
           Case Else
              MsgBox Err.Number & "::" & Err.Description, vbCritical
        End Select
     End If
     'fObservaciones.FG_Click
  'End If
  'MsgBox "Listado de Carnets Procesados", vbInformation
End Sub

Private Sub cmdProcesarCarnets_Click_()
  Dim lCmd As New ADODB.Command
  Dim lCn As New ADODB.Connection
  Dim lCodigoPVC As String
  Dim lReg As New ADODB.Recordset
  Dim SqlTxt As String
  Dim lCliente As String
  Dim lSubCliente As String
  Dim lMonto As Currency
  Dim lObservaciones As String
  lCliente = Mid(cCP.Text, 1, 6)
  If cSC.Text = "" Then
     lSubCliente = "0"
  Else
     lSubCliente = Mid(cSC.Text, 1, 6)
  End If
  lObservaciones = "PORLOTES"
  ''SaveSetting APPNAME, "Opciones", "CodigoCarnet", eCodigo.Text
  lCodigoPVC = GetSetting(APPNAME, "Opciones", "CodigoCarnet", "")
  If lCodigoPVC = "" Then
     MsgBox "Faltan datos en la configuración. Seleccione en el menú 'Opciones' y agregue el código del PVC ", vbExclamation
     Exit Sub
  End If
  SqlTxt = "Select * from PreciosEspeciales where Cliente=" & lCliente & " And SubCliente=" & lSubCliente & " and CodigoProducto='" & lCodigoPVC & "'"
  
  lReg.Open SqlTxt, Modulo.DBConexionSQL, adOpenKeyset
  If lReg.EOF = True Then
     lReg.Close
     lReg.Open "Select * from Productos Where codigo='" & lCodigoPVC & "'"
     If lReg.EOF = True Then
        MsgBox "No se puede continuar. No es posible encontrar el codigo del PVC", vbCritical
        Exit Sub
     End If
  End If
  Adodc1.Recordset.MoveFirst
  Do While Adodc1.Recordset.EOF = False
     If UCase(Adodc1.Recordset.Fields("Marca")) = "I" Then
        lMonto = lReg!precio * 1
        lCmd.ActiveConnection = Modulo.DBConexionSQL
        lCmd.CommandTimeout = 15
        lCmd.CommandText = "GuardarPorLotes"
        lCmd.Parameters.Append lCmd.CreateParameter("@Fecha", adDate, adParamInput, , Now)
        lCmd.Parameters.Append lCmd.CreateParameter("@Cliente", adInteger, adParamInput, , Mid(cCP.Text, 1, 6))
        lCmd.Parameters.Append lCmd.CreateParameter("@SubCliente", adInteger, adParamInput, , lSubCliente)
        lCmd.Parameters.Append lCmd.CreateParameter("@Tabla", adChar, adParamInput, 20, lTablas.List(lTablas.ListIndex))
        lCmd.Parameters.Append lCmd.CreateParameter("@Cedula", adChar, adParamInput, 20, Adodc1.Recordset.Fields("Cedula"))
        lCmd.Parameters.Append lCmd.CreateParameter("@Observaciones", adVarChar, adParamInput, 500, lObservaciones)
        lCmd.Parameters.Append lCmd.CreateParameter("@Pago", adChar, adParamInput, 1, "S")
        lCmd.Parameters.Append lCmd.CreateParameter("@Monto", adCurrency, adParamInput, , lMonto)
        lCmd.Parameters.Append lCmd.CreateParameter("@IdCarnet", adInteger, adParamInput, , Adodc1.Recordset.Fields("ID"))
        lCmd.Parameters.Append lCmd.CreateParameter("@Impreso", adSmallInt, adParamInput, , 0)
        lCmd.Parameters.Append lCmd.CreateParameter("@TipoPago", adSmallInt, adParamInput, , 1)
        lCmd.CommandType = adCmdStoredProc
        lCmd.Execute
        Set lCmd = Nothing
     End If
     Adodc1.Recordset.MoveNext
  Loop
End Sub

Private Sub Command1_Click()
  Dim sOri As String
  Dim sDes As String
  Dim sRuta As String
  Dim i As Integer
  Dim lReg As New ADODB.Recordset
  Dim lCedula As String
  sOri = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
  sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
    
  If sOri <> "" And sDes <> "" Then
     sRuta = sDes & "\" & Trim(Mid(cCP.Text, 9)) & "\" & IIf(cSC.Text <> "", Trim(Mid(cSC.Text, 9)) & "\", "") & "FOTOS"
     File1.Path = sRuta
     ListArchivos.Clear
     For i = 0 To File1.ListCount - 1
        ListArchivos.AddItem File1.List(i)
     Next i
  End If
  
     lReg.Open "Select cedula from [" & Trim(lTablas.List(lTablas.ListIndex)) & "] order by cedula", Modulo.DBConexionSQL, adOpenKeyset
  
  Load fMensaje
  fMensaje.Label1.Caption = "Buscando Archivos huérfanos..."
  fMensaje.Label1.Refresh
  fMensaje.Show
  DoEvents
  i = 1
  
     For i = 0 To ListArchivos.ListCount
        lReg.MoveFirst
        Do While lReg.EOF = False
           lCedula = Trim(lReg!cedula)
           lCedula = Replace(lCedula, ".", "")
           If ListArchivos.List(i) = "" Then Exit For
           If UCase(ListArchivos.List(i)) = lCedula & ".JPG" Then
              ListArchivos.RemoveItem i
              i = i - 1
              Exit Do
           End If
           lReg.MoveNext
        Loop
     Next i
     Modulo.ExecSQL "Delete From ArchivosHuerfanos"
     For i = 0 To ListArchivos.ListCount
       If ListArchivos.List(i) <> "" Then
           Modulo.ExecSQL "Insert into ArchivosHuerfanos (Archivo,CodigoCliente,NombreCliente,CodigoSubCliente,NombreSubCliente) " _
                       & " Values('" & ListArchivos.List(i) & "','" & Mid(cCP.Text, 1, 6) & "','" & Trim(Mid(cCP.Text, 9)) & "','" _
                       & Mid(cSC.Text, 1, 6) & "','" & Trim(Mid(cSC.Text, 9)) & "')"
                       
       End If
     Next i
     Unload fMensaje
     If ListArchivos.ListCount > 0 Then
        frmPreviewHuerfanos.sPrepararEmail Mid(cCP.Text, 1, 6), IIf(Len(cSC.Text) >= 6, Mid(cSC.Text, 1, 6), "")
        frmPreviewHuerfanos.ListAdjunto.Clear
        For i = 0 To ListArchivos.ListCount - 1
           frmPreviewHuerfanos.ListAdjunto.AddItem File1.Path & "\" & ListArchivos.List(i)
        Next i
        sRuta = sDes & "\" & Trim(Mid(cCP.Text, 9)) & "\" & IIf(cSC.Text <> "", Trim(Mid(cSC.Text, 9)) & "\", "")
        frmPreviewHuerfanos.lRuta = sRuta

        frmPreviewHuerfanos.Show vbModal
     Else
        MsgBox "No existen archivos huerfanos en la carpeta de fotos de este cliente", vbInformation
     End If
End Sub

Private Sub Command2_Click()
   If lTablas.Text = "" Then
      MsgBox "Debe seleccionar un cliente y/o subcliente y luego una tabla para continuar", vbInformation
      Exit Sub
   End If
   sAgregarCampoNroFoto
   frmFotosEnSitio.lblCliente.Caption = cCP.Text
   
   If cSC.Text <> "" Then frmFotosEnSitio.Label2.Visible = True
   frmFotosEnSitio.lblSubcliente.Caption = cSC.Text
   frmFotosEnSitio.Show
End Sub

Private Sub sAgregarCampoNroFoto()
   Dim lReg As New ADODB.Recordset
   Dim i As Integer
   Dim lExisteColumna As Boolean
   lExisteColumna = False
   Set lReg = DBConexionSQL.Execute("Select * from [" & lTablas.List(lTablas.ListIndex) & "]")
   For i = 0 To lReg.Fields.Count - 1
      If UCase(lReg.Fields(i).Name) = "NROFOTO" Then
         lExisteColumna = True
         Exit Sub
      End If
   Next i
   If lExisteColumna = False Then  ' si no existe la columna Nrofoto, crerarla
      DBConexionSQL.Execute "ALTER TABLE [" & lTablas.List(lTablas.ListIndex) & "] ADD NroFoto INT"
      DBConexionSQL.Execute "ALTER TABLE [H" & lTablas.List(lTablas.ListIndex) & "] ADD NroFoto INT"
   End If
End Sub

Private Sub cSC_Change()
  Cargar_Tablas_Personas
End Sub

Private Sub cSC_Click()
  Cargar_Tablas_Personas
End Sub

Private Sub DataGrid1_DblClick()
  Call bEditar_Click
  DataGrid1.Refresh
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

Private Sub Cargar_Tablas_Personas()
  Dim r As New ADODB.Recordset
  Dim s As String
  Dim SC As String, sSC As String
  
  lTablas.Clear
  lCreadas.Clear
  
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
    r.MoveNext
  Loop
  r.Close
  Set r = Nothing
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
  If cSC.ListCount > 0 Then cSC.ListIndex = 0
  cSC.ListIndex = -1
  
End Sub

Private Sub Cargar_Ordenar_Por()
  cOrden.Clear
  cOrden.AddItem "ID"
  cOrden.AddItem "CEDULA"
  cOrden.AddItem "NOMBRE"
  'cOrden.AddItem "CARGO"
  'cOrden.AddItem "VENCE"
  'cOrden.AddItem "FOTO"
  'cOrden.AddItem "FECHA"
  'Orden.AddItem "CONTADOR"
  'cOrden.AddItem "TIENE_FOTO"
  'cOrden.AddItem "CREACION"
  cOrden.ListIndex = 1

End Sub


Private Sub Form_Load()
  'cTipos.Clear
  'cTipos.AddItem "CHAR"
  'cTipos.AddItem "INTEGER"
  'cTipos.AddItem "DATETIME"
  'cTipos.AddItem "FLOAT"
  'cTipos.ListIndex = 0
  OPR = 0
  Cargar_Tablas_Personas
  Cargar_Clientes
  Cargar_SubClientes
  
  Cargar_Ordenar_Por
  cOrden.ListIndex = 0
  'bRefrescar_Click
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
    Mostrar_Personas_Registradas
  End If
End Sub

Private Sub lTablas_Click()
  If lTablas.ListIndex >= 0 Then
    lCreadas.ListIndex = lTablas.ListIndex
    Mostrar_Estructura
    Mostrar_Personas_Registradas
    bRefrescar_Click
    bAuditarFotos_Click
  End If
End Sub

Private Sub Mostrar_Personas_Registradas()
  Dim i As Integer
  Dim s As String
  Dim c As String
  
  i = -1
  If lTablas.ListIndex >= 0 Then i = lTablas.ListIndex
  If lCreadas.ListIndex >= 0 Then i = lCreadas.ListIndex
  If i = -1 Then
    MsgBox "Debe Seleccionar Tabla de Personas primero.", vbCritical, "Información"
  Else
    '--Ordenar por defecto: CEDULA (campo del sistema):
    c = ""
    If lCampos.ListCount > 0 Then c = lCampos.List(0)
    s = "SELECT * FROM [" & lTablas.List(i) & "]" & IIf(c <> "", " ORDER BY " & c, "")
    s = "SELECT * FROM [" & lTablas.List(i) & "] ORDER BY cedula"
    
    'If Adodc1.Recordset.State <> adStateClosed Then Adodc1.Recordset.Close
    
    Adodc1.ConnectionString = Modulo.DBConexionSQL.ConnectionString
    Adodc1.RecordSource = s
    Adodc1.Refresh
    DataGrid1.Refresh
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
    lCampos.Clear
    lTipos.Clear
    lAnchos.Clear
    For i = 0 To r.Fields.Count - 1
      lCampos.AddItem r.Fields(i).Name
      s = ""
      Select Case r.Fields(i).Type
        Case adChar: s = "CHAR"
        Case adInteger: s = "INTEGER"
        Case adDBTimeStamp: s = "DATETIME"
        Case adDouble: s = "FLOAT"
      End Select
      lTipos.AddItem s
      s1 = ""
      
      If s = "CHAR" Then s1 = CStr(r.Fields(i).DefinedSize)
            
      lAnchos.AddItem s1
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
