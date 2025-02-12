VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form fSubClientes 
   Caption         =   "SUB-CLIENTES"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   47
      Top             =   90
      Width           =   5715
   End
   Begin VB.CommandButton bBuscarCP 
      Height          =   345
      Left            =   7560
      Picture         =   "fSubClientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   90
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Sub-Clientes"
      Height          =   7545
      Left            =   0
      TabIndex        =   26
      Top             =   510
      Width           =   15225
      Begin VB.Frame Frame8 
         Caption         =   "Resumen de Cuenta"
         Height          =   765
         Left            =   2730
         TabIndex        =   48
         Top             =   5970
         Width           =   5205
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00808000&
            Height          =   525
            Left            =   60
            ScaleHeight     =   465
            ScaleWidth      =   5040
            TabIndex        =   49
            Top             =   180
            Width           =   5100
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Carnets Entregados a la Fecha:"
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   1260
               TabIndex        =   96
               Top             =   210
               Width           =   2235
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   3540
               TabIndex        =   95
               Top             =   210
               Width           =   120
            End
            Begin VB.Label lTD 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0,00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   780
               TabIndex        =   55
               Top             =   0
               Width           =   390
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Deuda Bs:"
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   30
               TabIndex        =   54
               Top             =   0
               Width           =   750
            End
            Begin VB.Label lTP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0,00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   2550
               TabIndex        =   53
               Top             =   0
               Width           =   390
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pagos Bs:"
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   1830
               TabIndex        =   52
               Top             =   0
               Width           =   720
            End
            Begin VB.Label lTS 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0,00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   4080
               TabIndex        =   51
               Top             =   0
               Width           =   390
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Saldo Bs:"
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   3390
               TabIndex        =   50
               Top             =   0
               Width           =   675
            End
         End
      End
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   30
         TabIndex        =   45
         Top             =   720
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CommandButton bAceptar 
         Caption         =   "Aceptar"
         Height          =   500
         Left            =   6630
         Picture         =   "fSubClientes.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   6870
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton bCancelar 
         Caption         =   "Cancelar"
         Height          =   500
         Left            =   7740
         Picture         =   "fSubClientes.frx":068C
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   6870
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.Frame Frame4 
         Caption         =   "Información del Sub-Cliente"
         Height          =   2475
         Left            =   8010
         TabIndex        =   29
         Top             =   990
         Width           =   7200
         Begin VB.OptionButton oSuspendido 
            Caption         =   "Suspendido"
            Height          =   225
            Left            =   5820
            TabIndex        =   23
            Top             =   2190
            Width           =   1245
         End
         Begin VB.OptionButton oActivo 
            Caption         =   "Activo"
            Height          =   225
            Left            =   4620
            TabIndex        =   22
            Top             =   2190
            Width           =   915
         End
         Begin MSComCtl2.DTPicker tinicio 
            Height          =   285
            Left            =   870
            TabIndex        =   21
            Top             =   2130
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            _Version        =   393216
            Format          =   62455809
            CurrentDate     =   39963
         End
         Begin VB.TextBox tcontelf 
            Height          =   315
            Left            =   4560
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   1830
            Width           =   2400
         End
         Begin VB.TextBox tcon 
            Height          =   315
            Left            =   870
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   1830
            Width           =   3150
         End
         Begin VB.TextBox tweb 
            Height          =   315
            Left            =   4560
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   1470
            Width           =   2400
         End
         Begin VB.TextBox temail 
            Height          =   315
            Left            =   870
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   1470
            Width           =   3150
         End
         Begin VB.TextBox tfax 
            Height          =   315
            Left            =   4560
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   1140
            Width           =   2400
         End
         Begin VB.TextBox ttel 
            Height          =   315
            Left            =   870
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   1140
            Width           =   3150
         End
         Begin VB.TextBox tdir 
            Height          =   315
            Left            =   870
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   810
            Width           =   6100
         End
         Begin VB.TextBox tnom 
            Height          =   315
            Left            =   870
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   480
            Width           =   6100
         End
         Begin VB.TextBox tnit 
            Height          =   315
            Left            =   5370
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   150
            Width           =   1600
         End
         Begin VB.TextBox trif 
            Height          =   315
            Left            =   2790
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   150
            Width           =   1600
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Estatus:"
            Height          =   195
            Left            =   3930
            TabIndex        =   42
            Top             =   2190
            Width           =   570
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Inicio:"
            Height          =   195
            Left            =   390
            TabIndex        =   41
            Top             =   2190
            Width           =   420
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Telf:"
            Height          =   195
            Left            =   4170
            TabIndex        =   40
            Top             =   1860
            Width           =   315
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Contacto:"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   1860
            Width           =   690
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Web:"
            Height          =   195
            Left            =   4110
            TabIndex        =   38
            Top             =   1500
            Width           =   390
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "EMail:"
            Height          =   195
            Left            =   330
            TabIndex        =   37
            Top             =   1500
            Width           =   435
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
            Height          =   195
            Left            =   4200
            TabIndex        =   36
            Top             =   1170
            Width           =   300
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Teléfonos:"
            Height          =   195
            Left            =   30
            TabIndex        =   35
            Top             =   1170
            Width           =   750
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   90
            TabIndex        =   34
            Top             =   870
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   210
            TabIndex        =   33
            Top             =   510
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "NIT Nº:"
            Height          =   195
            Left            =   4740
            TabIndex        =   32
            Top             =   210
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "RIF Nº:"
            Height          =   195
            Left            =   2190
            TabIndex        =   31
            Top             =   210
            Width           =   525
         End
         Begin VB.Label lcod 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "000000"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   870
            TabIndex        =   10
            Top             =   210
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   240
            TabIndex        =   30
            Top             =   210
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   765
         Left            =   8010
         TabIndex        =   28
         Top             =   240
         Width           =   5270
         Begin VB.CommandButton bBuscar 
            Caption         =   "Buscar"
            Height          =   500
            Left            =   2850
            Picture         =   "fSubClientes.frx":0C16
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   160
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.CommandButton bSalir 
            Caption         =   "Salir"
            Height          =   500
            Left            =   4140
            Picture         =   "fSubClientes.frx":11A0
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   160
            UseMaskColor    =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton bBorrar 
            Caption         =   "Borrar"
            Height          =   500
            Left            =   1950
            Picture         =   "fSubClientes.frx":172A
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   160
            UseMaskColor    =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton bEditar 
            Caption         =   "Editar"
            Height          =   500
            Left            =   1050
            Picture         =   "fSubClientes.frx":1CB4
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   160
            Width           =   900
         End
         Begin VB.CommandButton bNuevo 
            Caption         =   "Nuevo"
            Height          =   500
            Left            =   150
            Picture         =   "fSubClientes.frx":223E
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   160
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Height          =   765
         Left            =   120
         TabIndex        =   27
         Top             =   5970
         Width           =   2600
         Begin VB.CommandButton bUltimo 
            Height          =   500
            Left            =   1890
            Picture         =   "fSubClientes.frx":27C8
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Ultimo"
            Top             =   160
            Width           =   600
         End
         Begin VB.CommandButton bSiguiente 
            Height          =   500
            Left            =   1290
            Picture         =   "fSubClientes.frx":2D52
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Siguiente"
            Top             =   160
            Width           =   600
         End
         Begin VB.CommandButton bAnterior 
            Height          =   500
            Left            =   690
            Picture         =   "fSubClientes.frx":32DC
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Anterior"
            Top             =   160
            Width           =   600
         End
         Begin VB.CommandButton bPrimer 
            Height          =   500
            Left            =   90
            Picture         =   "fSubClientes.frx":3866
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Inicio"
            Top             =   160
            Width           =   600
         End
      End
      Begin MSFlexGridLib.MSFlexGrid FG 
         Height          =   5625
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   9922
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"fSubClientes.frx":3DF0
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3225
         Left            =   8010
         TabIndex        =   56
         Top             =   3510
         Width           =   7100
         _ExtentX        =   12515
         _ExtentY        =   5689
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "P. Autorizada Nº 01"
         TabPicture(0)   =   "fSubClientes.frx":3E80
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label18"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label19"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label20"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label32"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label33"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "eCedAuto"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "eNomAuto"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "eCarAuto"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "eTLFAuto"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Frame6"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "eFotoAuto"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "bExaminar"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).ControlCount=   12
         TabCaption(1)   =   "P. Autorizada Nº 02"
         TabPicture(1)   =   "fSubClientes.frx":3E9C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Command1"
         Tab(1).Control(1)=   "eFotoAuto2"
         Tab(1).Control(2)=   "Frame5"
         Tab(1).Control(3)=   "eTLFAuto2"
         Tab(1).Control(4)=   "eCarAuto2"
         Tab(1).Control(5)=   "eNomAuto2"
         Tab(1).Control(6)=   "eCedAuto2"
         Tab(1).Control(7)=   "Label35"
         Tab(1).Control(8)=   "Label34"
         Tab(1).Control(9)=   "Label24"
         Tab(1).Control(10)=   "Label25"
         Tab(1).Control(11)=   "Label26"
         Tab(1).ControlCount=   12
         TabCaption(2)   =   "P. Autorizada Nº 03"
         TabPicture(2)   =   "fSubClientes.frx":3EB8
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Command2"
         Tab(2).Control(1)=   "eFotoAuto3"
         Tab(2).Control(2)=   "Frame10"
         Tab(2).Control(3)=   "eTLFAuto3"
         Tab(2).Control(4)=   "eCarAuto3"
         Tab(2).Control(5)=   "eNomAuto3"
         Tab(2).Control(6)=   "eCedAuto3"
         Tab(2).Control(7)=   "Label27"
         Tab(2).Control(8)=   "Label28"
         Tab(2).Control(9)=   "Label29"
         Tab(2).Control(10)=   "Label30"
         Tab(2).Control(11)=   "Label31"
         Tab(2).ControlCount=   12
         Begin VB.CommandButton bExaminar 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3510
            TabIndex        =   77
            Top             =   2610
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox eFotoAuto 
            Height          =   315
            Left            =   930
            TabIndex        =   76
            Text            =   "Text1"
            Top             =   2250
            Width           =   3150
         End
         Begin VB.Frame Frame6 
            Caption         =   "Foto"
            Height          =   2775
            Left            =   4530
            TabIndex        =   75
            Top             =   360
            Width           =   2475
            Begin VB.Image Image1 
               Height          =   2535
               Left            =   60
               Stretch         =   -1  'True
               Top             =   180
               Width           =   2355
            End
         End
         Begin VB.TextBox eTLFAuto 
            Height          =   315
            Left            =   930
            TabIndex        =   74
            Text            =   "Text1"
            Top             =   1860
            Width           =   3150
         End
         Begin VB.TextBox eCarAuto 
            Height          =   315
            Left            =   930
            TabIndex        =   73
            Text            =   "Text1"
            Top             =   1470
            Width           =   3150
         End
         Begin VB.TextBox eNomAuto 
            Height          =   315
            Left            =   930
            TabIndex        =   72
            Text            =   "Text1"
            Top             =   1110
            Width           =   3150
         End
         Begin VB.TextBox eCedAuto 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   930
            TabIndex        =   71
            Text            =   "Text1"
            Top             =   690
            Width           =   1300
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -71490
            TabIndex        =   70
            Top             =   2610
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox eFotoAuto2 
            Height          =   315
            Left            =   -74070
            TabIndex        =   69
            Text            =   "Text1"
            Top             =   2250
            Width           =   3150
         End
         Begin VB.Frame Frame5 
            Caption         =   "Foto"
            Height          =   2775
            Left            =   -70470
            TabIndex        =   68
            Top             =   360
            Width           =   2475
            Begin VB.Image Image2 
               Height          =   2535
               Left            =   60
               Stretch         =   -1  'True
               Top             =   180
               Width           =   2355
            End
         End
         Begin VB.TextBox eTLFAuto2 
            Height          =   315
            Left            =   -74070
            TabIndex        =   67
            Text            =   "Text1"
            Top             =   1860
            Width           =   3150
         End
         Begin VB.TextBox eCarAuto2 
            Height          =   315
            Left            =   -74070
            TabIndex        =   66
            Text            =   "Text1"
            Top             =   1470
            Width           =   3150
         End
         Begin VB.TextBox eNomAuto2 
            Height          =   315
            Left            =   -74070
            TabIndex        =   65
            Text            =   "Text1"
            Top             =   1110
            Width           =   3150
         End
         Begin VB.TextBox eCedAuto2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   -74070
            TabIndex        =   64
            Text            =   "Text1"
            Top             =   690
            Width           =   1300
         End
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -71490
            TabIndex        =   63
            Top             =   2610
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox eFotoAuto3 
            Height          =   315
            Left            =   -74070
            TabIndex        =   62
            Text            =   "Text1"
            Top             =   2250
            Width           =   3150
         End
         Begin VB.Frame Frame10 
            Caption         =   "Foto"
            Height          =   2775
            Left            =   -70470
            TabIndex        =   61
            Top             =   360
            Width           =   2475
            Begin VB.Image Image3 
               Height          =   2535
               Left            =   60
               Stretch         =   -1  'True
               Top             =   180
               Width           =   2355
            End
         End
         Begin VB.TextBox eTLFAuto3 
            Height          =   315
            Left            =   -74070
            TabIndex        =   60
            Text            =   "Text1"
            Top             =   1860
            Width           =   3150
         End
         Begin VB.TextBox eCarAuto3 
            Height          =   315
            Left            =   -74070
            TabIndex        =   59
            Text            =   "Text1"
            Top             =   1470
            Width           =   3150
         End
         Begin VB.TextBox eNomAuto3 
            Height          =   315
            Left            =   -74070
            TabIndex        =   58
            Text            =   "Text1"
            Top             =   1140
            Width           =   3150
         End
         Begin VB.TextBox eCedAuto3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   -74070
            TabIndex        =   57
            Text            =   "Text1"
            Top             =   690
            Width           =   1300
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Foto:"
            Height          =   195
            Left            =   -74490
            TabIndex        =   94
            Top             =   2310
            Width           =   360
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   -74790
            TabIndex        =   93
            Top             =   1920
            Width           =   675
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Foto:"
            Height          =   195
            Left            =   510
            TabIndex        =   92
            Top             =   2310
            Width           =   360
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   210
            TabIndex        =   91
            Top             =   1920
            Width           =   675
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Cargo:"
            Height          =   195
            Left            =   390
            TabIndex        =   90
            Top             =   1530
            Width           =   465
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Cédula:"
            Height          =   195
            Left            =   330
            TabIndex        =   89
            Top             =   750
            Width           =   540
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   270
            TabIndex        =   88
            Top             =   1140
            Width           =   600
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Foto:"
            Height          =   195
            Left            =   -74640
            TabIndex        =   87
            Top             =   2310
            Width           =   360
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   -74940
            TabIndex        =   86
            Top             =   1920
            Width           =   675
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Cargo:"
            Height          =   195
            Left            =   -74610
            TabIndex        =   85
            Top             =   1530
            Width           =   465
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Cédula:"
            Height          =   195
            Left            =   -74670
            TabIndex        =   84
            Top             =   750
            Width           =   540
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   -74730
            TabIndex        =   83
            Top             =   1140
            Width           =   600
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Foto:"
            Height          =   195
            Left            =   -74490
            TabIndex        =   82
            Top             =   2310
            Width           =   360
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   -74790
            TabIndex        =   81
            Top             =   1920
            Width           =   675
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Cargo:"
            Height          =   195
            Left            =   -74610
            TabIndex        =   80
            Top             =   1530
            Width           =   465
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Cédula:"
            Height          =   195
            Left            =   -74670
            TabIndex        =   79
            Top             =   750
            Width           =   540
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   -74730
            TabIndex        =   78
            Top             =   1140
            Width           =   600
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   6930
         TabIndex        =   43
         Top             =   105
         Width           =   45
      End
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
      TabIndex        =   44
      Top             =   150
      Width           =   1485
   End
End
Attribute VB_Name = "fSubClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const FGC = "ID          | Nombre                                                    | RIF Nº             | Telefonos                                    "

Dim RSubClientes As New ADODB.Recordset

Const OPR_NUEVO = 1
Const OPR_EDITAR = 2

Dim OPR As Integer  'operacion: 1 nuevo  2 editar
Dim Viendo As Boolean


Private Sub Limpiar_FG()
  FG.Clear
  FG.Rows = 2
  FG.FormatString = FGC
End Sub

Private Sub Limpiar_Txts()
  trif.Text = ""
  tnit.Text = ""
  tnom.Text = ""
  tdir.Text = ""
  ttel.Text = ""
  tfax.Text = ""
  temail.Text = ""
  tweb.Text = ""
  tcon.Text = ""
  tcontelf.Text = ""
  tinicio.value = Date
  oActivo.value = True
  oSuspendido.value = False
  
  eCedAuto.Text = ""
  eNomAuto.Text = ""
  eCarAuto.Text = ""
  eTLFAuto.Text = ""
  eFotoAuto.Text = ""
  
  Set Image1.Picture = Nothing
  
  eCedAuto2.Text = ""
  eNomAuto2.Text = ""
  eCarAuto2.Text = ""
  eTLFAuto2.Text = ""
  eFotoAuto2.Text = ""
  Set Image2.Picture = Nothing
  
  eCedAuto3.Text = ""
  eNomAuto3.Text = ""
  eCarAuto3.Text = ""
  eTLFAuto3.Text = ""
  eFotoAuto3.Text = ""
  Set Image3.Picture = Nothing
  

  
End Sub

Private Sub Activar_Txts(TF As Boolean)
  'TF (True / False)
  trif.Enabled = TF
  tnit.Enabled = TF
  tnom.Enabled = TF
  tdir.Enabled = TF
  ttel.Enabled = TF
  tfax.Enabled = TF
  temail.Enabled = TF
  tweb.Enabled = TF
  tcon.Enabled = TF
  tcontelf.Enabled = TF
  tinicio.Enabled = TF
  oActivo.Enabled = TF
  oSuspendido.Enabled = TF
  
  eCedAuto.Enabled = TF
  eNomAuto.Enabled = TF
  eCarAuto.Enabled = TF
  eTLFAuto.Enabled = TF
  eFotoAuto.Enabled = TF
  
  'Set Image1.Picture = Nothing
  
  eCedAuto2.Enabled = TF
  eNomAuto2.Enabled = TF
  eCarAuto2.Enabled = TF
  eTLFAuto2.Enabled = TF
  eFotoAuto2.Enabled = TF
  
  'Set Image2.Picture = Nothing
  
  eCedAuto3.Enabled = TF
  eNomAuto3.Enabled = TF
  eCarAuto3.Enabled = TF
  eTLFAuto3.Enabled = TF
  eFotoAuto3.Enabled = TF
  
  'Set Image3.Picture = Nothing
  
  
End Sub

Private Sub Activar_Btns(Num As Integer, TF As Boolean)
  If Num = 0 Then 'TODOS los botones
  
    bPrimer.Enabled = TF
    bAnterior.Enabled = TF
    bSiguiente.Enabled = TF
    bUltimo.Enabled = TF
    
    bNuevo.Enabled = TF
    bEditar.Enabled = TF
    bBorrar.Enabled = TF
    bBuscar.Enabled = TF
    
    bSalir.Enabled = TF
       
    'bSubClientes.Enabled = TF
    
    bAceptar.Enabled = TF
    bCancelar.Enabled = TF
    
  Else
  
    Select Case Num
      
      Case 1: bPrimer.Enabled = TF
      Case 2: bAnterior.Enabled = TF
      Case 3: bSiguiente.Enabled = TF
      Case 4: bUltimo.Enabled = TF
      
      Case 5: bNuevo.Enabled = TF
      Case 6: bEditar.Enabled = TF
      Case 7: bBorrar.Enabled = TF
      Case 8: bBuscar.Enabled = TF
      
      Case 9: bSalir.Enabled = TF
          
      'Case 10: bSubClientes.Enabled = TF
      
      Case 11: bAceptar.Enabled = TF
      Case 12: bCancelar.Enabled = TF
      
    End Select
    
  End If
End Sub

Private Sub Cargar_Clientes()
  Dim r As New ADODB.Recordset
  r.Open "SELECT * FROM clientes ORDER BY codigo", DBConexionSQL, adOpenDynamic, adLockOptimistic
  cCP.Clear
  Do While Not r.EOF
    cCP.AddItem Trim(Zeros(r.Fields("codigo").value, 6) & " : " & r.Fields("nombre").value)
    r.MoveNext
  Loop
  If cCP.ListCount > 0 Then cCP.ListIndex = 0
End Sub


Public Sub Cargar_SubClientes()
  'Dim r As New ADODB.Recordset
  Dim s As String
  Dim l As Integer
  Dim sCod As String
  
  Viendo = False
  
  Limpiar_FG
  
  If RSubClientes.State <> adStateClosed Then RSubClientes.Close
  
  sCod = "000000"
  If Trim(cCP.Text) <> "" Then sCod = Mid(cCP.Text, 1, 6)
      
  s = "SELECT * FROM SubClientes WHERE Cliente = " & sCod & " ORDER BY Id"
  
  RSubClientes.Open s, DBConexionSQL, adOpenDynamic, adLockOptimistic
  l = 1
  Do While Not RSubClientes.EOF
  
    FG.TextMatrix(l, 0) = Zeros(RSubClientes.Fields("id").value, 6)
    FG.TextMatrix(l, 1) = Trim(RSubClientes.Fields("nombre").value)
    FG.TextMatrix(l, 2) = IIf(IsNull(Trim(RSubClientes.Fields("rif").value)), "", Trim(RSubClientes.Fields("rif").value))
    FG.Row = l: FG.Col = 2: FG.CellAlignment = flexAlignLeftCenter
    FG.TextMatrix(l, 3) = IIf(IsNull(Trim(RSubClientes.Fields("telefonos").value)), "", Trim(RSubClientes.Fields("telefonos").value))
    FG.Row = l: FG.Col = 3: FG.CellAlignment = flexAlignLeftCenter
    
    RSubClientes.MoveNext
    
    FG.Col = 0
    
    If Not RSubClientes.EOF Then
      l = l + 1
      FG.Rows = FG.Rows + 1
    End If
  Loop
  
  FG.Refresh

  If FG.Rows >= 1 Then
    FG.Row = 1
    FG.Col = 0
    Scroll_SubCliente
    'FG.SetFocus
  End If
    
  If Not RSubClientes.BOF And Not RSubClientes.EOF Then RSubClientes.MoveFirst
  
  Call Mostrar_SubCliente
  
  Label2.Caption = "[" & CStr(Total_SubClientes()) & " Regs.]"
  'RClientes.Close
  'Set r = Nothing
  Viendo = True
End Sub

Private Sub Scroll_SubCliente()
  Dim c As String
  Viendo = False
  If FG.Row >= 1 Then
    c = Trim(FG.TextMatrix(FG.Row, 0))
    If c <> "" Then
      RSubClientes.MoveFirst
      RSubClientes.Find "id = " & c
      If RSubClientes.EOF Then
        MsgBox "Debe Seleccionar el Sub-Cliente...", vbCritical, "Información"
      Else
        Mostrar_SubCliente
      End If
    Else
      Limpiar_Txts
    End If
  End If
  Viendo = True
End Sub

Private Sub Ubicar_Cursor_SubCliente(sCod As String)
  Dim i As Integer
  Dim f As Integer
  Dim e As Boolean
  
  i = 1
  If FG.Rows >= 1 Then
    e = False
    Do While i < FG.Rows And Not e
      If FG.TextMatrix(i, 0) = sCod Then
        e = True
      Else
        i = i + 1
      End If
    Loop
    If (i + 1) < FG.Rows Then FG.Row = i + 1
  End If
      
End Sub

Private Sub Mostrar_SubCliente()
  Dim s As String, sRutaCliente As String
  Dim s1 As String
  
  If RSubClientes.State <> adStateClosed Then
    If Not RSubClientes.EOF Then
      If OPR = OPR_NUEVO Then
        lcod.Caption = "Nuevo"
      Else
        lcod.Caption = Zeros(RSubClientes.Fields("id").value, 6)
      End If
      trif.Text = IIf(IsNull(Trim(RSubClientes.Fields("rif").value)), "", Trim(RSubClientes.Fields("rif").value))
      tnit.Text = IIf(IsNull(Trim(RSubClientes.Fields("nit").value)), "", Trim(RSubClientes.Fields("nit").value))
      'tnit.Text = Trim(RSubClientes.Fields("nit").value)
      tnom.Text = Trim(RSubClientes.Fields("nombre").value)
      tdir.Text = IIf(IsNull(Trim(RSubClientes.Fields("direccion").value)), "", Trim(RSubClientes.Fields("direccion").value))
      'tdir.Text = Trim(RSubClientes.Fields("direccion").value)
      ttel.Text = IIf(IsNull(Trim(RSubClientes.Fields("telefonos").value)), "", Trim(RSubClientes.Fields("telefonos").value))
      'ttel.Text = Trim(RSubClientes.Fields("telefonos").value)
      tfax.Text = IIf(IsNull(Trim(RSubClientes.Fields("fax").value)), "", Trim(RSubClientes.Fields("fax").value))
      'tfax.Text = Trim(RSubClientes.Fields("fax").value)
      temail.Text = IIf(IsNull(Trim(RSubClientes.Fields("email").value)), "", Trim(RSubClientes.Fields("email").value))
      'temail.Text = Trim(RSubClientes.Fields("email").value)
      tweb.Text = IIf(IsNull(Trim(RSubClientes.Fields("website").value)), "", Trim(RSubClientes.Fields("website").value))
      'tweb.Text = Trim(RSubClientes.Fields("website").value)
      tcon.Text = IIf(IsNull(Trim(RSubClientes.Fields("contacto").value)), "", Trim(RSubClientes.Fields("contacto").value))
      'tcon.Text = Trim(RSubClientes.Fields("contacto").value)
      tcontelf.Text = IIf(IsNull(Trim(RSubClientes.Fields("contactotlf").value)), "", Trim(RSubClientes.Fields("contactotlf").value))
      'tcontelf.Text = Trim(RSubClientes.Fields("contactotlf").value)
      tinicio.value = Trim(RSubClientes.Fields("fechainicio").value)
      oActivo.value = False
      oSuspendido.value = False
      If RSubClientes.Fields("activo").value = "S" Then
        oActivo.value = True
      Else
        oSuspendido.value = True
      End If
      
      '* P.Autorizada #1
      If Not IsNull(RSubClientes.Fields("cedulaauto1").value) Then eCedAuto.Text = Trim(RSubClientes.Fields("cedulaauto1").value)
      
      If Not IsNull(RSubClientes.Fields("nombreauto1").value) Then eNomAuto.Text = Trim(RSubClientes.Fields("nombreauto1").value)
      
      If Not IsNull(RSubClientes.Fields("cargoauto1").value) Then eCarAuto.Text = Trim(RSubClientes.Fields("cargoauto1").value)
      
      If Not IsNull(RSubClientes.Fields("telefauto1").value) Then eTLFAuto.Text = Trim(RSubClientes.Fields("telefauto1").value)
      
      If Not IsNull(RSubClientes.Fields("fotoauto1").value) Then eFotoAuto.Text = Trim(RSubClientes.Fields("fotoauto1").value)
      
      
      
      
      Set Image1.Picture = Nothing
      
      If Not IsNull(RSubClientes.Fields("fotoauto1").value) Then
      
        s = Trim(RSubClientes.Fields("fotoauto1").value)
        If s <> "" Then
          If InStr(s, "/") > 0 Or InStr(s, "\") > 0 Then 'Buscar completa la cadena de FOTO
            Image1.Picture = LoadPicture(s)
          Else
            s1 = "-"
            If FG.Row >= 0 Then s1 = FG.TextMatrix(FG.Row, 0)
                   
            sRutaCliente = Modulo.LA_RUTA_FOTO_DEL_CLIENTE(cCP.Text, s1)
            If sRutaCliente <> "" Then
              s = sRutaCliente & "\" & s
              If Dir(s) <> "" Then Image1.Picture = LoadPicture(s)
            End If
          End If
        End If
      End If
      
      
      '* P.Autorizada #2
      eCedAuto2.Text = IIf(IsNull(Trim(RSubClientes.Fields("cedulaauto2").value)), "", Trim(RSubClientes.Fields("cedulaauto2").value))
      'eCedAuto2.Text = Trim(RSubClientes.Fields("cedulaauto2").value)
      eNomAuto2.Text = IIf(IsNull(Trim(RSubClientes.Fields("nombreauto2").value)), "", Trim(RSubClientes.Fields("nombreauto2").value))
      'eNomAuto2.Text = Trim(RSubClientes.Fields("nombreauto2").value)
      eCarAuto2.Text = IIf(IsNull(Trim(RSubClientes.Fields("cargoauto2").value)), "", Trim(RSubClientes.Fields("cargoauto2").value))
      'eCarAuto2.Text = Trim(RSubClientes.Fields("cargoauto2").value)
      eTLFAuto2.Text = IIf(IsNull(Trim(RSubClientes.Fields("telefAuto2").value)), "", Trim(RSubClientes.Fields("telefAuto2").value))
      'eTLFAuto2.Text = Trim(RSubClientes.Fields("telefauto2").value)
      eFotoAuto2.Text = IIf(IsNull(Trim(RSubClientes.Fields("fotoauto2").value)), "", Trim(RSubClientes.Fields("fotoauto2").value))
      'eFotoAuto2.Text = Trim(RSubClientes.Fields("fotoauto2").value)
      
      Set Image2.Picture = Nothing
  
      s = IIf(IsNull(Trim(RSubClientes.Fields("fotoauto2").value)), "", Trim(RSubClientes.Fields("fotoauto2").value))
      's = Trim(RSubClientes.Fields("fotoauto2").value)
      If s <> "" Then
        If InStr(s, "/") > 0 Or InStr(s, "\") > 0 Then 'Buscar completa la cadena de FOTO
          Image2.Picture = LoadPicture(s)
        Else
          s1 = "-"
          If FG.Row >= 0 Then s1 = FG.TextMatrix(FG.Row, 0)
                   
          sRutaCliente = Modulo.LA_RUTA_FOTO_DEL_CLIENTE(cCP.Text, s1)
          If sRutaCliente <> "" Then
            s = sRutaCliente & "\" & s
            If Dir(s) <> "" Then Image2.Picture = LoadPicture(s)
          End If
        End If
      End If
      
      '* P.Autorizada #3
      eCedAuto3.Text = IIf(IsNull(Trim(RSubClientes.Fields("cedulaauto3").value)), "", Trim(RSubClientes.Fields("cedulaauto3").value))
      'eCedAuto3.Text = Trim(RSubClientes.Fields("cedulaauto3").value)
      eNomAuto3.Text = IIf(IsNull(Trim(RSubClientes.Fields("nombreauto3").value)), "", Trim(RSubClientes.Fields("nombreauto3").value))
      'eNomAuto3.Text = Trim(RSubClientes.Fields("nombreauto3").value)
      eCarAuto3.Text = IIf(IsNull(Trim(RSubClientes.Fields("cargoauto3").value)), "", Trim(RSubClientes.Fields("cargoauto3").value))
      'eCarAuto3.Text = Trim(RSubClientes.Fields("cargoauto3").value)
      eTLFAuto3.Text = IIf(IsNull(Trim(RSubClientes.Fields("telefauto3").value)), "", Trim(RSubClientes.Fields("telefauto3").value))
      'eTLFAuto3.Text = Trim(RSubClientes.Fields("telefauto3").value)
      eFotoAuto3.Text = IIf(IsNull(Trim(RSubClientes.Fields("fotoauto3").value)), "", Trim(RSubClientes.Fields("fotoauto3").value))
      'eFotoAuto3.Text = Trim(RSubClientes.Fields("fotoauto3").value)
      
      Set Image3.Picture = Nothing
  
      s = IIf(IsNull(Trim(RSubClientes.Fields("fotoauto3").value)), "", Trim(RSubClientes.Fields("fotoauto3").value))
      's = Trim(RSubClientes.Fields("fotoauto3").value)
      If s <> "" Then
        If InStr(s, "/") > 0 Or InStr(s, "\") > 0 Then 'Buscar completa la cadena de FOTO
          Image3.Picture = LoadPicture(s)
        Else
          s1 = "-"
          If FG.Row >= 0 Then s1 = FG.TextMatrix(FG.Row, 0)
                   
          sRutaCliente = Modulo.LA_RUTA_FOTO_DEL_CLIENTE(cCP.Text, s1)
          If sRutaCliente <> "" Then
            s = sRutaCliente & "\" & s
            If Dir(s) <> "" Then Image3.Picture = LoadPicture(s)
          End If
        End If
      End If
      
      
      
      lTD.Caption = Format(RSubClientes.Fields("deuda").value, "#,0.00")
      lTP.Caption = Format(RSubClientes.Fields("pagos").value, "#,0.00")
      lTS.Caption = Format(RSubClientes.Fields("saldo").value, "#,0.00")

      
      
    End If
  End If
End Sub

Function CodigoSubClienteNuevo() As Long
  Dim r As New ADODB.Recordset
  Dim ccn As Long
  ccn = 0
  r.Open "SELECT * FROM subclientes ORDER BY Id", DBConexionSQL, adOpenKeyset, adLockReadOnly
  If Not r.EOF Then
    r.MoveLast
    ccn = r.Fields("id").value
  End If
  CodigoSubClienteNuevo = ccn + 1
End Function

Function Total_SubClientes() As Long
  Dim r As New ADODB.Recordset
  Dim ccn As Long
  ccn = 0
  Modulo.ClientePPAL = Mid(cCP.Text, 1, 6)
  If Trim(Modulo.ClientePPAL) = "" Then Modulo.ClientePPAL = "000000"
  r.Open "SELECT count(*) FROM subclientes WHERE cliente = " & Modulo.ClientePPAL & " ", DBConexionSQL, adOpenKeyset, adLockReadOnly
  If Not r.EOF Then
    If Not IsNull(r.Fields(0).value) Then
      ccn = r.Fields(0).value
    End If
  End If
  r.Close
  Set r = Nothing
  Total_SubClientes = ccn
End Function

Private Function fExisteSubCliente(argRIF As String) As Boolean
  Dim lCn As New ADODB.Connection
  Dim lReg As New ADODB.Recordset
  lCn.ConnectionString = Modulo.DBConexionSQL
  lCn.Open
  Set lReg = lCn.Execute("Select * from subclientes where Rif='" & UCase(argRIF) & "'") '' and codigo <> " & Val(lcod))
  If lReg.EOF = False Then
     fExisteSubCliente = True
  Else
     fExisteSubCliente = False
  End If
End Function

Private Function fExisteCliente(argRIF As String) As Boolean
  Dim lCn As New ADODB.Connection
  Dim lReg As New ADODB.Recordset
  lCn.ConnectionString = Modulo.DBConexionSQL
  lCn.Open
  Set lReg = lCn.Execute("Select * from clientes where Rif='" & UCase(argRIF) & "' and codigo <> " & Val(lcod))
  If lReg.EOF = False Then
     fExisteCliente = True
  Else
     fExisteCliente = False
  End If
End Function


Private Sub bAceptar_Click()
  Dim s As String, nc As String
  Dim c As Long
  Dim i As Integer
  
  Dim sCP As String, snp As String
  
  If Trim(tnom) = "" Then
    MsgBox "Debe Introducir el Nombre del Sub-Cliente, Revise...", vbCritical, "Información"
    tnom.SetFocus
    Exit Sub
  End If
  
  If OPR = OPR_NUEVO Then
      
    If Trim(cCP.Text) = "" Then
      MsgBox "No hay Cliente Principal Seleccionado, Revise...", vbCritical, "Información"
      Exit Sub
    End If
    
    If fExisteCliente(trif.Text) = True Then
       MsgBox "Este RIF ya existe en la base datos del sistema. Es cliente (" & trif.Text & ")", vbExclamation
       trif.SelStart = 0
       trif.SelLength = Len(trif.Text)
       trif.SetFocus
       
       Exit Sub
    End If

    
    If fExisteSubCliente(trif.Text) = True Then
       MsgBox "Este RIF ya existe en la base datos del sistema. Es subcliente (" & trif.Text & ")", vbExclamation
       trif.SelStart = 0
       trif.SelLength = Len(trif.Text)
       trif.SetFocus
       
       Exit Sub
    End If

    On Error Resume Next
    
    sCP = Trim(Mid(cCP.Text, 1, 6))
    snp = Trim(Mid(cCP.Text, 10))
    
    
    '//Agregar el nuevo sub-cliente...
    
    If Not Nombre_Directorio_Valido(tnom.Text) Then Exit Sub
    If Nombre_Directorio_Repetido(tnom.Text, False) Then Exit Sub

      
    Load fMensaje
    fMensaje.Label1.Caption = "Añadiendo Sub-Cliente [" & Trim(tnom.Text) & "], Espere..."
    fMensaje.Show
    DoEvents
      
    With RSubClientes
      .AddNew
      .Fields("cliente").value = CLng(sCP)
      .Fields("rif").value = Trim(trif.Text)
      .Fields("nit").value = Trim(tnit.Text)
      .Fields("nombre").value = Trim(tnom.Text)
      .Fields("direccion").value = Trim(tdir.Text)
      .Fields("telefonos").value = Trim(ttel.Text)
      .Fields("fax").value = Trim(tfax.Text)
      .Fields("email").value = Trim(temail.Text)
      .Fields("website").value = Trim(tweb.Text)
      .Fields("contacto").value = Trim(tcon.Text)
      .Fields("contactotlf").value = Trim(tcontelf.Text)
      .Fields("fechainicio").value = tinicio.value
      If oActivo.value Then .Fields("activo").value = "S" Else .Fields("activo").value = "N"
      
      .Fields("deuda").value = 0#
      .Fields("pagos").value = 0#
      .Fields("saldo").value = 0#
      
      .Fields("cedulaauto1").value = Trim(eCedAuto.Text)
      .Fields("nombreauto1").value = Trim(eNomAuto.Text)
      .Fields("cargoauto1").value = Trim(eCarAuto.Text)
      .Fields("telefauto1").value = Trim(eTLFAuto.Text)
      .Fields("fotoauto1").value = Trim(eFotoAuto.Text)
      
      .Fields("cedulaauto2").value = Trim(eCedAuto2.Text)
      .Fields("nombreauto2").value = Trim(eNomAuto2.Text)
      .Fields("cargoauto2").value = Trim(eCarAuto2.Text)
      .Fields("telefauto2").value = Trim(eTLFAuto2.Text)
      .Fields("fotoauto2").value = Trim(eFotoAuto2.Text)
      
      .Fields("cedulaauto3").value = Trim(eCedAuto3.Text)
      .Fields("nombreauto3").value = Trim(eNomAuto3.Text)
      .Fields("cargoauto3").value = Trim(eCarAuto3.Text)
      .Fields("telefauto3").value = Trim(eTLFAuto3.Text)
      .Fields("fotoauto3").value = Trim(eFotoAuto3.Text)
      
      
      
      .Update
    End With
        
    RSubClientes.Close
    RSubClientes.Open
    'If Not RSubClientes.EOF Then RSubClientes.MoveLast
    'c = RSubClientes.Fields("id").Value 'id asignado por el MYSQL
        
    Unload fMensaje
    
    If Err.Number <> 0 Then
      MsgBox "Ha Ocurrido un Error al Intentar almacenar el Registro..." & vbCrLf & Err.Description, vbCritical, "Información"
      Exit Sub
    End If
    
    AgregarLogs "Agrega SubCliente [" & Mid(Trim(tnom.Text), 1, 20) & "...]"
    
    Crear_Carpetas_SubClientes
    Dim j As Integer
    For j = 0 To (fPersonas.cCP.ListCount) - 1
       If Trim(fPersonas.cCP.List(j)) = Trim(cCP.List(cCP.ListIndex)) Then
           fPersonas.cCP.Text = fPersonas.cCP.List(j)
           fPersonas.cSC.Text = fPersonas.cSC.List(fPersonas.cSC.ListCount - 1)
          Exit For
       End If
    Next j
    
    
    fPersonas.Show vbModal
        
            
    Cargar_SubClientes
    Scroll_SubCliente
    Limpiar_Txts
    
    lcod.Caption = "Nuevo" 'Zeros(CodigoSubClienteNuevo(), 6)
    
    trif.SetFocus
    
  Else
  
    If OPR = OPR_EDITAR Then
    
      Load fMensaje
      fMensaje.Label1.Caption = "Actualizando Cliente, Espere..."
      fMensaje.Show
   
      On Error Resume Next
      
      nc = lcod.Caption
      sCP = Trim(Mid(cCP.Text, 1, 6))
      
      s = "UPDATE subclientes SET " & _
          "rif         = '" & Trim(trif.Text) & "',  nit         = '" & Trim(tnit.Text) & "'," & _
          "nombre      = '" & Trim(tnom.Text) & " ', direccion   = '" & Trim(tdir.Text) & "'," & _
          "telefonos   = '" & Trim(ttel.Text) & "',  fax         = '" & Trim(tfax.Text) & "'," & _
          "email       = '" & Trim(temail.Text) & "',website     = '" & Trim(tweb.Text) & "'," & _
          "contacto    = '" & Trim(tcon.Text) & "',  contactotlf = '" & Trim(tcontelf.Text) & "'," & _
          "fechainicio = '" & Format(tinicio.value, "yyyymmdd") & "'," & _
          "activo      = '" & IIf(oActivo.value = True, "S", "N") & "'," & _
          "cedulaauto1 = '" & Trim(eCedAuto.Text) & "'," & _
          "nombreauto1 = '" & Trim(eNomAuto.Text) & "'," & _
          "cargoauto1  = '" & Trim(eCarAuto.Text) & "'," & _
          "telefauto1  = '" & Trim(eTLFAuto.Text) & "'," & _
          "fotoauto1   = '" & Trim(eFotoAuto.Text) & "'," & _
          "cedulaauto2 = '" & Trim(eCedAuto2.Text) & "'," & _
          "nombreauto2 = '" & Trim(eNomAuto2.Text) & "'," & _
          "cargoauto2  = '" & Trim(eCarAuto2.Text) & "'," & _
          "telefauto2  = '" & Trim(eTLFAuto2.Text) & "'," & _
          "fotoauto2   = '" & Trim(eFotoAuto2.Text) & "'," & _
          "cedulaauto3 = '" & Trim(eCedAuto3.Text) & "'," & _
          "nombreauto3 = '" & Trim(eNomAuto3.Text) & "'," & _
          "cargoauto3  = '" & Trim(eCarAuto3.Text) & "'," & _
          "telefauto3  = '" & Trim(eTLFAuto3.Text) & "'," & _
          "fotoauto3   = '" & Trim(eFotoAuto3.Text) & "' " & _
          "WHERE " & _
          "id = " & nc & " and cliente = " & sCP & " "
          
      Set DBComandoSQL.ActiveConnection = DBConexionSQL
      DBComandoSQL.CommandText = s
      DBComandoSQL.Execute
      
      Unload fMensaje
      
      If Err.Number <> 0 Then
        MsgBox "Ha Ocurrido un Error al Intentar actualizar el Registro..." & vbCrLf & Err.Description, vbCritical, "Información"
        Exit Sub
      End If
      
      AgregarLogs "Actualiza SubCliente [" & Mid(Trim(tnom.Text), 1, 20) & "...]"
      
      SSTab1.Tab = 0
      
      Cargar_SubClientes
            
      bCancelar_Click
      
      Ubicar_Cursor_SubCliente nc
    
    End If

  End If

End Sub


Private Sub Crear_Carpetas_SubClientes()
  Dim c As String, so As String, s As String
  Dim s2 As String
  
  
    '-----------------------------------------------------------
    '- CREAR LAS CARPETAS Y COPIAR LAS PLANTILLAS
    '-----------------------------------------------------------
    
    Load fMensaje
    fMensaje.Label1.Caption = "Creando Carpetas/Plantillas del Sub-Cliente, Espere..."
    fMensaje.Show
    
    Dim sOri As String
    Dim sDes As String
      
    sOri = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
    sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
    
    If sOri <> "" And sDes <> "" Then
    
      '-- Nombre del Cliente:
      '-- 123456789012345
      '-- 000001 : COMERC
      so = sDes & "\" & Trim(Mid(cCP.Text, 10)) & "\" & Trim(tnom.Text)
      
      MkDir (so)
      
      '-- Subcarpetas: Carnets - Fotos - Imagenes
        
        s = so & "\" & "CARNET"
        MkDir (s)
        
        s = so & "\" & "FOTOS"
        MkDir (s)
        
        s = so & "\" & "IMAGENES"
        MkDir (s)
        
        '-- Copiar archivos (RAIZ):
        s = sOri & "\FORMATO DE ENTREGA DE CARNETS _nombre cliente.xls"
        s2 = so & "\FORMATO DE ENTREGA DE CARNETS " & tnom.Text & ".xls"
        FileCopy s, s2
        
        s = sOri & "\LISTADO NOMBRE CLIENTE OFICINA.xls"
        s2 = so & "\LISTADO " & Trim(tnom.Text) & ".xls"
        FileCopy s, s2
        
        '-- Copiar archivos (CARNET):
        s = sOri & "\CARNET\BASE CARNET NOMBRE CLIENTE.car"
        s2 = so & "\CARNET\" & Trim(tnom.Text) & ".car"
        FileCopy s, s2
        
        s = sOri & "\CARNET\BASE DATOS NOMBRE CLIENTE.mdb"
        s2 = so & "\CARNET\" & Trim(tnom.Text) & ".mdb"
        FileCopy s, s2
        
        '-- Copiar archivos (IMAGENES):
        s = sOri & "\IMAGENES\Autocopia_de_seguridad_deBASE nombre cliente.cdr"
        s2 = so & "\IMAGENES\Autocopia_de_seguridad_de " & Trim(tnom.Text) & ".cdr"
        FileCopy s, s2
        
        s = sOri & "\IMAGENES\BASE nombre cliente.cdr"
        s2 = so & "\IMAGENES\" & Trim(tnom.Text) & ".cdr"
        FileCopy s, s2
    End If
    
    Unload fMensaje
End Sub

Function Nombre_SubDirectorio_Repetido(sNom As String) As Boolean
  Dim r As New ADODB.Recordset
  Dim s As String
  s = "SELECT * FROM SubClientes WHERE Nombre = '" & sNom & "'"
  r.Open s, Modulo.DBConexionSQL, adOpenKeyset, adLockReadOnly
  If Not r.EOF Then
    MsgBox "Existe un Sub-Cliente con el Nombre [" & Trim(r.Fields("nombre").value) & "]" & vbCrLf & _
           "Código SubCliente = " & Zeros(r.Fields("id").value, 6) & " - RIF: " & Trim(r.Fields("rif").value), vbCritical, "Información"
    Nombre_SubDirectorio_Repetido = True
  Else
    Nombre_SubDirectorio_Repetido = False
  End If
  r.Close
  Set r = Nothing
End Function







Private Sub bBorrar_Click()
  Dim c As String, n As String, s As String, s2 As String
  Dim sOri As String
  Dim sDes As String
      
  sOri = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
  sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
    
  If sOri = "" Or sDes = "" Then
    MsgBox "Debe Configurar la Ruta de las Carpetas Origen/Destino...", vbCritical, "Información"
    Exit Sub
  End If
  
  If FG.Row >= 1 Then
    c = Trim(FG.TextMatrix(FG.Row, 0))
    n = Trim(FG.TextMatrix(FG.Row, 1))
    If c <> "" Then
      
      If MsgBox("¿Está Seguro de Borrar el Sub-Cliente " & vbCrLf & c & "-" & n & "?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
        s = "DELETE FROM SubClientes WHERE id = " & c & " AND cliente = " & Modulo.ClientePPAL & " "
        Set DBComandoSQL.ActiveConnection = DBConexionSQL
        DBComandoSQL.CommandText = s
        DBComandoSQL.Execute
        
        '-- Borrando las carpetas del sub-cliente:
        
        s = Trim(Mid(cCP.Text, 10))
        s2 = s
        
        s = "RMDIR " & Chr(34) & sDes & "\" & s2 & "\" & tnom.Text & Chr(34) & " /S /Q"
        
        s = Ejecutar_DOS(s)

        
          
'''''        '-- Ir dentro de la carpeta del sub-cliente y borrar todos los
'''''        '-- archivos:
'''''        s = sDes & "\" & s & "\" & tnom.Text
'''''        File1.Path = s
'''''        File1.Refresh
'''''        If File1.ListCount > 0 Then Kill s & "\*.*"
'''''
'''''          s = sDes & "\" & s2 & "\" & tnom.Text & "\CARNET"
'''''          File1.Path = s
'''''          File1.Refresh
'''''          If File1.ListCount > 0 Then Kill s & "\*.*"
'''''
'''''          s = sDes & "\" & s2 & "\" & tnom.Text & "\FOTOS"
'''''          File1.Path = s
'''''          File1.Refresh
'''''          If File1.ListCount > 0 Then Kill s & "\*.*"
'''''
'''''          s = sDes & "\" & s2 & "\" & tnom.Text & "\IMAGENES"
'''''          File1.Path = s
'''''          File1.Refresh
'''''          If File1.ListCount > 0 Then Kill s & "\*.*"
'''''
'''''          '-- borrar las sub-carpetas:
'''''          s = sDes & "\" & s2 & "\" & tnom.Text & "\CARNET"
'''''          RmDir s
'''''
'''''          s = sDes & "\" & s2 & "\" & tnom.Text & "\FOTOS"
'''''          RmDir s
'''''
'''''          s = sDes & "\" & s2 & "\" & tnom.Text & "\IMAGENES"
'''''          RmDir s
'''''
'''''          '-- Borrar la carpeta del SUB-Cliente:
'''''          s = sDes & "\" & s2 & "\" & tnom.Text
'''''          RmDir s
'''''        'End If
                 
        AgregarLogs "Borra SubCliente [" & c & "...]"
        
        MsgBox "SUB-CLIENTE [" & c & "] FUE BORRADO (" & Modulo.USUARIO_ACTUAL & " " & Format(Now, "dd/mm/yy hh:mm ampm") & ")", vbInformation, "Información"
        
        
        
        'FALTA LOGs
        Cargar_SubClientes  'Abre el RecordSet de Clientes
        Scroll_SubCliente
      End If
    End If
  End If
End Sub

Private Sub bBuscar_Click()
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
    Viendo = False
    i = 1
    e = False
    Do While i < FG.Rows And Not e
      If FG.TextMatrix(i, 0) = Modulo.vTemporal1 Then
        e = True
        FG.Row = i
        FG.Col = 0
      Else
        i = i + 1
      End If
    Loop
    FG.SetFocus
    SendKeys "{LEFT}"
    'Scroll_Cliente
    
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

Public Sub bEditar_Click()
  If FG.Row >= 1 Then
  
    If Trim(FG.TextMatrix(FG.Row, 0)) = "" Then Exit Sub
    
  Else
  
    Exit Sub
    
  End If
  
  OPR = OPR_EDITAR
  
  cCP.Enabled = False
  
  fSubClientes.Caption = "EDITAR SUB-CLIENTE"
  
  FG.Enabled = False
  Frame2.Enabled = False
  Activar_Btns 5, False
  Activar_Btns 6, False
  Activar_Btns 7, False
  Activar_Btns 8, False
  
  'Limpiar_Txts
  Activar_Txts True
  
  Activar_Btns 10, True
  Activar_Btns 11, True
  Activar_Btns 12, True
  
  'lcod.Caption = Zeros(CodigoClienteNuevo(), 6)
  tnom.Enabled = False
  
  trif.SetFocus
  
  
End Sub

Private Sub bNuevo_Click()
  OPR = OPR_NUEVO
  
  cCP.Enabled = False
  
  fSubClientes.Caption = "NUEVO SUB-CLIENTE"
  
  FG.Enabled = False
  Frame2.Enabled = False
  Activar_Btns 5, False
  Activar_Btns 6, False
  Activar_Btns 7, False
  Activar_Btns 8, False
  
  Limpiar_Txts
  Activar_Txts True
  
  Activar_Btns 10, True
  Activar_Btns 11, True
  Activar_Btns 12, True
  
  lcod.Caption = Zeros(CodigoSubClienteNuevo(), 6)
  
  trif.SetFocus
  
     
  
  
End Sub

Private Sub bCancelar_Click()
  OPR = 0
  
  cCP.Enabled = True
  
  fSubClientes.Caption = "SUB-CLIENTES"
  
  FG.Enabled = True
  Frame2.Enabled = True
  Activar_Btns 5, True
  Activar_Btns 6, True
  Activar_Btns 7, True
  Activar_Btns 8, True
  
  Limpiar_Txts
  Activar_Txts False
  lcod.Caption = ""
  
  Activar_Btns 10, False
  Activar_Btns 11, False
  Activar_Btns 12, False
  
  Scroll_SubCliente
  
  SSTab1.Tab = 0
  
  FG.SetFocus
  SendKeys "{UP}"
  
  

End Sub

Private Sub bPrimer_Click()
  'Ir al primer registro del listado
  FG.Row = 1
  Scroll_SubCliente
  FG.SetFocus
  SendKeys "{UP}"
End Sub

Private Sub bAnterior_Click()
  'Ir al anterior registro del listado
  If FG.Rows > 2 Then
    FG.Row = FG.Row - 1
    Scroll_SubCliente
  End If
  FG.SetFocus
  SendKeys "{LEFT}"
End Sub

Private Sub bSiguiente_Click()
  'Ir al siguiente registro del listado
  If FG.Rows > 2 Then
    If FG.Row <= FG.Rows - 1 Then
      FG.Row = FG.Row + 1
      Scroll_SubCliente
    End If
  End If
  FG.SetFocus
  SendKeys "{RIGHT}"
End Sub



Private Sub bUltimo_Click()
  'Ir al ultimo registro del listado
  If FG.Rows > 1 Then
    If FG.Row >= 1 Then
      FG.Row = FG.Rows - 1
      Scroll_SubCliente
    End If
  End If
  FG.SetFocus
  SendKeys "{DOWN}"
End Sub

Private Sub bSalir_Click()
  Unload Me
End Sub

Private Sub cCP_Click()
  If Trim(cCP.Text) <> "" Then Cargar_SubClientes
End Sub

Private Sub eCarAuto2_Change()
  EnMayusculas eCarAuto2
End Sub

Private Sub eCarAuto3_Change()
  EnMayusculas eCarAuto3
End Sub

Private Sub FG_Click()
  Scroll_SubCliente
End Sub

Private Sub FG_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyHome Then bPrimer_Click Else
  If KeyCode = vbKeyEnd Then bUltimo_Click
End Sub

Private Sub FG_RowColChange()
  If Viendo Then Scroll_SubCliente
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

Private Sub Form_Load()
  'lcodppal.Caption = Modulo.ClientePPAL
  SCargarfSubClientes
  
End Sub

Public Sub SCargarfSubClientes()
Cargar_Clientes
  
  Viendo = False
  OPR = 0
  
  Cargar_SubClientes  'Abre el RecordSet de Clientes
  
  Limpiar_Txts
  Activar_Txts False
  Activar_Btns 0, False
  Activar_Btns 9, True
  lcod.Enabled = False
  
  Mostrar_SubCliente
  
  If FG.Rows > 1 Then
    Activar_Btns 1, True
    Activar_Btns 2, True
    Activar_Btns 3, True
    Activar_Btns 4, True
  End If
  
  Activar_Btns 5, True
  Activar_Btns 6, True
  Activar_Btns 7, True
  Activar_Btns 8, True
  
  Viendo = True
  
  SendKeys "{UP}"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If RSubClientes.State <> adStateClosed Then
    RSubClientes.Close
    Set RSubClientes = Nothing
  End If
  Unload Me
End Sub

Private Sub tcon_Change()
  EnMayusculas tcon
End Sub

Private Sub tcontelf_Change()
    EnMayusculas tcontelf
End Sub

Private Sub tdir_Change()
    EnMayusculas tdir
End Sub

Private Sub temail_Change()
    EnMayusculas temail
End Sub

Private Sub tfax_Change()
    EnMayusculas tfax
End Sub

Private Sub tinicio_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then oActivo.SetFocus
End Sub

Private Sub tnit_Change()
    EnMayusculas tnit
End Sub

Private Sub tnom_Change()
    EnMayusculas tnom
End Sub

Private Sub trif_Change()
  EnMayusculas trif
End Sub

Private Sub ttel_Change()
    EnMayusculas ttel
End Sub

Private Sub tweb_Change()
    EnMayusculas tweb
End Sub

Private Sub eCarAuto_Change()
  EnMayusculas eCarAuto
End Sub

Private Sub eCedAuto_Change()
  Dim s As String, d As Double
  eCedAuto.Text = UCase(eCedAuto.Text)
  If IsNumeric(eCedAuto.Text) Then
    d = CDbl(eCedAuto.Text)
    s = Format(d, "#,0")
    eCedAuto.Text = s
    If eCedAuto.Enabled Then eCedAuto.SelStart = Len(eCedAuto.Text)
            
    eFotoAuto.Text = CStr(d) & EXTJPG
    If eFotoAuto.Text = EXTJPG Then eFotoAuto.Text = ""
  Else
    eFotoAuto.Text = eCedAuto.Text & ".JPG"
    If eFotoAuto.Text = EXTJPG Then eFotoAuto.Text = ""
  End If
End Sub

Private Sub eCedAuto_KeyPress(KeyAscii As Integer)
  Dim s As String, d As Double
  If KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
    KeyAscii = 0
  Else
    If KeyAscii = vbKeyReturn Then
      If IsNumeric(eCedAuto.Text) Then
        d = CDbl(eCedAuto.Text)
        s = Format(d, "#,0")
        eCedAuto.Text = s
        eFotoAuto.Text = CStr(d) & ".JPG"
      Else
        eFotoAuto.Text = eCedAuto.Text & ".JPG"
      End If
    End If
  End If
End Sub

Private Sub eCedAuto_LostFocus()
  Dim s As String, d As Double
  If IsNumeric(eCedAuto.Text) Then
    d = CDbl(eCedAuto.Text)
    s = Format(d, "#,0")
    eCedAuto.Text = s
  End If
End Sub

Private Sub eFotoAuto_Change()
  Dim s As String, sRutaCliente As String, s1 As String
  s = Trim(eFotoAuto.Text)
  If s <> "" Then
    If InStr(s, "/") > 0 Or InStr(s, "\") > 0 Then 'Buscar completa la cadena de FOTO
      Image1.Picture = LoadPicture(s)
    Else
      s1 = "-"
      If FG.Row >= 1 And FG.Row < FG.Rows Then s1 = FG.TextMatrix(FG.Row, 0)
      sRutaCliente = Modulo.LA_RUTA_FOTO_DEL_CLIENTE(cCP.Text, s1)
      If sRutaCliente <> "" Then
        s = sRutaCliente & "\" & s '& EXTJPG
        If Dir(s) <> "" Then Image1.Picture = LoadPicture(s) Else Set Image1.Picture = Nothing
      End If
    End If
  End If
End Sub

Private Sub eNomAuto_Change()
  EnMayusculas eNomAuto
End Sub

'************************************
Private Sub eCedAuto2_Change()
  Dim s As String, d As Double
  eCedAuto2.Text = UCase(eCedAuto2.Text)
  If IsNumeric(eCedAuto2.Text) Then
    d = CDbl(eCedAuto2.Text)
    s = Format(d, "#,0")
    eCedAuto2.Text = s
    If eCedAuto2.Enabled Then eCedAuto2.SelStart = Len(eCedAuto2.Text)
            
    eFotoAuto2.Text = CStr(d) & EXTJPG
  Else
    eFotoAuto2.Text = eCedAuto2.Text & EXTJPG
  End If
  If eFotoAuto2.Text = EXTJPG Then eFotoAuto2.Text = ""
End Sub

Private Sub eCedAuto2_KeyPress(KeyAscii As Integer)
  Dim s As String, d As Double
  If KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
    KeyAscii = 0
  Else
    If KeyAscii = vbKeyReturn Then
      If IsNumeric(eCedAuto2.Text) Then
        d = CDbl(eCedAuto2.Text)
        s = Format(d, "#,0")
        eCedAuto2.Text = s
        eFotoAuto2.Text = CStr(d) & EXTJPG
      Else
        eFotoAuto2.Text = eCedAuto2.Text & EXTJPG
      End If
    End If
  End If
End Sub

Private Sub eCedAuto2_LostFocus()
  Dim s As String, d As Double
  If IsNumeric(eCedAuto2.Text) Then
    d = CDbl(eCedAuto2.Text)
    s = Format(d, "#,0")
    eCedAuto2.Text = s
  End If
End Sub

Private Sub eFotoAuto2_Change()
  Dim s As String, sRutaCliente As String, s1 As String
  s = Trim(eFotoAuto2.Text)
  If s <> "" Then
    If InStr(s, "/") > 0 Or InStr(s, "\") > 0 Then 'Buscar completa la cadena de FOTO
      Image2.Picture = LoadPicture(s)
    Else
      s1 = "-"
      If FG.Row >= 1 And FG.Row < FG.Rows Then s1 = FG.TextMatrix(FG.Row, 0)
      sRutaCliente = Modulo.LA_RUTA_FOTO_DEL_CLIENTE(cCP.Text, s1)
      If sRutaCliente <> "" Then
        s = sRutaCliente & "\" & s '& EXTJPG
        If Dir(s) <> "" Then Image2.Picture = LoadPicture(s) Else Set Image2.Picture = Nothing
      End If
    End If
  End If
End Sub

Private Sub eNomAuto2_Change()
  EnMayusculas eNomAuto2
End Sub

'************************************
Private Sub eCedAuto3_Change()
  Dim s As String, d As Double
  eCedAuto3.Text = UCase(eCedAuto3.Text)
  If IsNumeric(eCedAuto3.Text) Then
    d = CDbl(eCedAuto3.Text)
    s = Format(d, "#,0")
    eCedAuto3.Text = s
    If eCedAuto3.Enabled Then eCedAuto3.SelStart = Len(eCedAuto3.Text)
            
    eFotoAuto3.Text = CStr(d) & EXTJPG
  Else
    eFotoAuto3.Text = eCedAuto3.Text & EXTJPG
  End If
  If eFotoAuto3.Text = EXTJPG Then eFotoAuto3.Text = ""
End Sub

Private Sub eCedAuto3_KeyPress(KeyAscii As Integer)
  Dim s As String, d As Double
  If KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
    KeyAscii = 0
  Else
    If KeyAscii = vbKeyReturn Then
      If IsNumeric(eCedAuto3.Text) Then
        d = CDbl(eCedAuto3.Text)
        s = Format(d, "#,0")
        eCedAuto3.Text = s
        eFotoAuto3.Text = CStr(d) & EXTJPG
      Else
        eFotoAuto3.Text = eCedAuto3.Text & EXTJPG
      End If
    End If
  End If
End Sub

Private Sub eCedAuto3_LostFocus()
  Dim s As String, d As Double
  If IsNumeric(eCedAuto3.Text) Then
    d = CDbl(eCedAuto3.Text)
    s = Format(d, "#,0")
    eCedAuto3.Text = s
  End If
End Sub

Private Sub eFotoAuto3_Change()
  Dim s As String, sRutaCliente As String, s1 As String
  s = Trim(eFotoAuto3.Text)
  If s <> "" Then
    If InStr(s, "/") > 0 Or InStr(s, "\") > 0 Then 'Buscar completa la cadena de FOTO
      Image3.Picture = LoadPicture(s)
    Else
      s1 = "-"
      If FG.Row >= 1 And FG.Row < FG.Rows Then s1 = FG.TextMatrix(FG.Row, 0)
      sRutaCliente = Modulo.LA_RUTA_FOTO_DEL_CLIENTE(cCP.Text, s1)
      If sRutaCliente <> "" Then
        s = sRutaCliente & "\" & s '& EXTJPG
        If Dir(s) <> "" Then Image3.Picture = LoadPicture(s) Else Set Image3.Picture = Nothing
      End If
    End If
  End If
End Sub

Private Sub eNomAuto3_Change()
  EnMayusculas eNomAuto3
End Sub


