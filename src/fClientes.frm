VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form fClientes 
   Caption         =   "CLIENTES"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Clientes"
      Height          =   10965
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   15240
      Begin TabDlg.SSTab SSTab1 
         Height          =   3225
         Left            =   8730
         TabIndex        =   71
         Top             =   5580
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   5689
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         TabCaption(0)   =   "P. Autorizada Nº 01"
         TabPicture(0)   =   "fClientes.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label14"
         Tab(0).Control(1)=   "Label18"
         Tab(0).Control(2)=   "Label17"
         Tab(0).Control(3)=   "Label16"
         Tab(0).Control(4)=   "Label15"
         Tab(0).Control(5)=   "bExaminar"
         Tab(0).Control(6)=   "eFotoAuto"
         Tab(0).Control(7)=   "Frame6"
         Tab(0).Control(8)=   "eTLFAuto"
         Tab(0).Control(9)=   "eCarAuto"
         Tab(0).Control(10)=   "eNomAuto"
         Tab(0).Control(11)=   "eCedAuto"
         Tab(0).ControlCount=   12
         TabCaption(1)   =   "P. Autorizada Nº 02"
         TabPicture(1)   =   "fClientes.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label22"
         Tab(1).Control(1)=   "Label23"
         Tab(1).Control(2)=   "Label24"
         Tab(1).Control(3)=   "Label25"
         Tab(1).Control(4)=   "Label26"
         Tab(1).Control(5)=   "Command1"
         Tab(1).Control(6)=   "eFotoAuto2"
         Tab(1).Control(7)=   "Frame5"
         Tab(1).Control(8)=   "eTLFAuto2"
         Tab(1).Control(9)=   "eCarAuto2"
         Tab(1).Control(10)=   "eNomAuto2"
         Tab(1).Control(11)=   "eCedAuto2"
         Tab(1).ControlCount=   12
         TabCaption(2)   =   "P. Autorizada Nº 03"
         TabPicture(2)   =   "fClientes.frx":0038
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Label27"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Label28"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Label29"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "Label30"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Label31"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "Command2"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "eFotoAuto3"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "Frame10"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "eTLFAuto3"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "eCarAuto3"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "eNomAuto3"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "eCedAuto3"
         Tab(2).Control(11).Enabled=   0   'False
         Tab(2).ControlCount=   12
         Begin VB.TextBox eCedAuto3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   780
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   690
            Width           =   1300
         End
         Begin VB.TextBox eNomAuto3 
            Height          =   315
            Left            =   780
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   1110
            Width           =   3150
         End
         Begin VB.TextBox eCarAuto3 
            Height          =   315
            Left            =   780
            TabIndex        =   38
            Text            =   "Text1"
            Top             =   1500
            Width           =   3150
         End
         Begin VB.TextBox eTLFAuto3 
            Height          =   315
            Left            =   780
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   1860
            Width           =   3150
         End
         Begin VB.Frame Frame10 
            Caption         =   "Foto"
            Height          =   2775
            Left            =   3930
            TabIndex        =   84
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
         Begin VB.TextBox eFotoAuto3 
            Height          =   315
            Left            =   780
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   2250
            Width           =   3150
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
            Left            =   3510
            TabIndex        =   41
            Top             =   2610
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox eCedAuto2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   -74220
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   690
            Width           =   1300
         End
         Begin VB.TextBox eNomAuto2 
            Height          =   315
            Left            =   -74220
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   1110
            Width           =   3150
         End
         Begin VB.TextBox eCarAuto2 
            Height          =   315
            Left            =   -74220
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   1470
            Width           =   3150
         End
         Begin VB.TextBox eTLFAuto2 
            Height          =   315
            Left            =   -74220
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   1860
            Width           =   3150
         End
         Begin VB.Frame Frame5 
            Caption         =   "Foto"
            Height          =   2775
            Left            =   -71070
            TabIndex        =   78
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
         Begin VB.TextBox eFotoAuto2 
            Height          =   315
            Left            =   -74220
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   2250
            Width           =   3150
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
            TabIndex        =   35
            Top             =   2610
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox eCedAuto 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   -74220
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   690
            Width           =   1300
         End
         Begin VB.TextBox eNomAuto 
            Height          =   315
            Left            =   -74220
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   1110
            Width           =   3150
         End
         Begin VB.TextBox eCarAuto 
            Height          =   315
            Left            =   -74220
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   1470
            Width           =   3150
         End
         Begin VB.TextBox eTLFAuto 
            Height          =   315
            Left            =   -74220
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   1860
            Width           =   3150
         End
         Begin VB.Frame Frame6 
            Caption         =   "Foto"
            Height          =   2775
            Left            =   -71070
            TabIndex        =   72
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
         Begin VB.TextBox eFotoAuto 
            Height          =   315
            Left            =   -74220
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   2250
            Width           =   3150
         End
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
            Left            =   -71490
            TabIndex        =   29
            Top             =   2610
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   120
            TabIndex        =   89
            Top             =   1110
            Width           =   600
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Cédula:"
            Height          =   195
            Left            =   180
            TabIndex        =   88
            Top             =   750
            Width           =   540
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Cargo:"
            Height          =   195
            Left            =   240
            TabIndex        =   87
            Top             =   1530
            Width           =   465
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   60
            TabIndex        =   86
            Top             =   1920
            Width           =   675
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Foto:"
            Height          =   195
            Left            =   360
            TabIndex        =   85
            Top             =   2310
            Width           =   360
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   -74880
            TabIndex        =   83
            Top             =   1110
            Width           =   600
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Cédula:"
            Height          =   195
            Left            =   -74820
            TabIndex        =   82
            Top             =   750
            Width           =   540
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Cargo:"
            Height          =   195
            Left            =   -74760
            TabIndex        =   81
            Top             =   1530
            Width           =   465
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   -74940
            TabIndex        =   80
            Top             =   1920
            Width           =   675
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Foto:"
            Height          =   195
            Left            =   -74640
            TabIndex        =   79
            Top             =   2310
            Width           =   360
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   -74880
            TabIndex        =   77
            Top             =   1140
            Width           =   600
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Cédula:"
            Height          =   195
            Left            =   -74820
            TabIndex        =   76
            Top             =   750
            Width           =   540
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Cargo:"
            Height          =   195
            Left            =   -74760
            TabIndex        =   75
            Top             =   1530
            Width           =   465
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   -74940
            TabIndex        =   74
            Top             =   1920
            Width           =   675
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Foto:"
            Height          =   195
            Left            =   -74640
            TabIndex        =   73
            Top             =   2310
            Width           =   360
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Resumen de Cuenta"
         Height          =   1035
         Left            =   8730
         TabIndex        =   63
         Top             =   8820
         Width           =   6405
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00808000&
            Height          =   705
            Left            =   120
            ScaleHeight     =   645
            ScaleWidth      =   6135
            TabIndex        =   64
            Top             =   210
            Width           =   6195
            Begin VB.Label Label33 
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
               Left            =   2310
               TabIndex        =   91
               Top             =   360
               Width           =   630
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Carnets Entregados a la Fecha:"
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   30
               TabIndex        =   90
               Top             =   360
               Width           =   2235
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Saldo Bs:"
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   3900
               TabIndex        =   70
               Top             =   90
               Width           =   675
            End
            Begin VB.Label lTS 
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
               Left            =   4650
               TabIndex        =   69
               Top             =   90
               Width           =   1230
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Créditos Bs:"
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   1980
               TabIndex        =   68
               Top             =   90
               Width           =   840
            End
            Begin VB.Label lTP 
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
               Left            =   2790
               TabIndex        =   67
               Top             =   90
               Width           =   1050
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Débitos Bs:"
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   30
               TabIndex        =   66
               Top             =   90
               Width           =   810
            End
            Begin VB.Label lTD 
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
               Left            =   840
               TabIndex        =   65
               Top             =   90
               Width           =   1110
            End
         End
      End
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   2490
         TabIndex        =   61
         Top             =   750
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CommandButton bAceptar 
         Caption         =   "Aceptar"
         Height          =   500
         Left            =   5460
         Picture         =   "fClientes.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   9360
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton bCancelar 
         Caption         =   "Cancelar"
         Height          =   500
         Left            =   6600
         Picture         =   "fClientes.frx":05DE
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   9390
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.Frame Frame4 
         Caption         =   "Información del Cliente"
         Height          =   4275
         Left            =   8790
         TabIndex        =   47
         Top             =   1260
         Width           =   6400
         Begin VB.CheckBox Check1 
            Caption         =   "Contacto es Autorizado"
            Height          =   375
            Left            =   4560
            TabIndex        =   92
            Top             =   3300
            Width           =   1575
         End
         Begin VB.Frame Frame7 
            Caption         =   "Estatus"
            Height          =   1215
            Left            =   4590
            TabIndex        =   62
            Top             =   2010
            Width           =   1515
            Begin VB.OptionButton oSuspendido 
               Caption         =   "Suspendido"
               Height          =   225
               Left            =   210
               TabIndex        =   23
               Top             =   780
               Width           =   1215
            End
            Begin VB.OptionButton oActivo 
               Caption         =   "Activo"
               Height          =   225
               Left            =   210
               TabIndex        =   22
               Top             =   360
               Width           =   915
            End
         End
         Begin MSComCtl2.DTPicker tinicio 
            Height          =   285
            Left            =   750
            TabIndex        =   21
            Top             =   3900
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            _Version        =   393216
            Format          =   55377921
            CurrentDate     =   39963
         End
         Begin VB.TextBox tcontelf 
            Height          =   315
            Left            =   750
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   3510
            Width           =   3500
         End
         Begin VB.TextBox tcon 
            Height          =   315
            Left            =   750
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   3120
            Width           =   3500
         End
         Begin VB.TextBox tweb 
            Height          =   315
            Left            =   750
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   2730
            Width           =   3500
         End
         Begin VB.TextBox temail 
            Height          =   315
            Left            =   750
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   2340
            Width           =   3500
         End
         Begin VB.TextBox tfax 
            Height          =   315
            Left            =   750
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   1950
            Width           =   3500
         End
         Begin VB.TextBox ttel 
            Height          =   315
            Left            =   750
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   1560
            Width           =   3500
         End
         Begin VB.TextBox tdir 
            Height          =   315
            Left            =   750
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   1170
            Width           =   5400
         End
         Begin VB.TextBox tnom 
            Height          =   315
            Left            =   750
            MaxLength       =   100
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   750
            Width           =   5400
         End
         Begin VB.TextBox tnit 
            Height          =   315
            Left            =   4800
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   240
            Width           =   1300
         End
         Begin VB.TextBox trif 
            Height          =   315
            Left            =   2610
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   240
            Width           =   1300
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Inicio:"
            Height          =   195
            Left            =   300
            TabIndex        =   59
            Top             =   3960
            Width           =   420
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Telf:"
            Height          =   195
            Left            =   390
            TabIndex        =   58
            Top             =   3540
            Width           =   315
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Contacto:"
            Height          =   195
            Left            =   30
            TabIndex        =   57
            Top             =   3150
            Width           =   690
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Web:"
            Height          =   195
            Left            =   330
            TabIndex        =   56
            Top             =   2760
            Width           =   390
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "EMail:"
            Height          =   195
            Left            =   270
            TabIndex        =   55
            Top             =   2400
            Width           =   435
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
            Height          =   195
            Left            =   420
            TabIndex        =   54
            Top             =   1980
            Width           =   300
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   60
            TabIndex        =   53
            Top             =   1590
            Width           =   675
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   30
            TabIndex        =   52
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   90
            TabIndex        =   51
            Top             =   780
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "NIT Nº:"
            Height          =   195
            Left            =   4170
            TabIndex        =   50
            Top             =   300
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "RIF Nº:"
            Height          =   195
            Left            =   2010
            TabIndex        =   49
            Top             =   300
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
            Left            =   720
            TabIndex        =   10
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   90
            TabIndex        =   48
            Top             =   300
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   885
         Left            =   8790
         TabIndex        =   46
         Top             =   210
         Width           =   6285
         Begin VB.CommandButton bBuscar 
            Caption         =   "Buscar"
            Height          =   550
            Left            =   2880
            Picture         =   "fClientes.frx":0B68
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Buscar"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton bSalir 
            Caption         =   "Salir"
            Height          =   550
            Left            =   4260
            Picture         =   "fClientes.frx":10F2
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Salir"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton bBorrar 
            Caption         =   "Borrar"
            Height          =   550
            Left            =   1950
            Picture         =   "fClientes.frx":167C
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Borrar"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton bEditar 
            Caption         =   "Editar"
            Height          =   550
            Left            =   1050
            Picture         =   "fClientes.frx":1C06
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Editar"
            Top             =   165
            Width           =   900
         End
         Begin VB.CommandButton bNuevo 
            Caption         =   "Nuevo"
            Height          =   550
            Left            =   150
            Picture         =   "fClientes.frx":2190
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Nuevo"
            Top             =   165
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Height          =   765
         Left            =   1650
         TabIndex        =   45
         Top             =   9210
         Width           =   2600
         Begin VB.CommandButton bUltimo 
            Height          =   500
            Left            =   1890
            Picture         =   "fClientes.frx":271A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Ultimo"
            Top             =   160
            Width           =   600
         End
         Begin VB.CommandButton bSiguiente 
            Height          =   500
            Left            =   1290
            Picture         =   "fClientes.frx":2CA4
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Siguiente"
            Top             =   160
            Width           =   600
         End
         Begin VB.CommandButton bAnterior 
            Height          =   500
            Left            =   690
            Picture         =   "fClientes.frx":322E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Anterior"
            Top             =   160
            Width           =   600
         End
         Begin VB.CommandButton bPrimer 
            Height          =   500
            Left            =   90
            Picture         =   "fClientes.frx":37B8
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Inicio"
            Top             =   160
            Width           =   600
         End
      End
      Begin MSFlexGridLib.MSFlexGrid FG 
         Height          =   9015
         Left            =   30
         TabIndex        =   0
         Top             =   240
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   15901
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"fClientes.frx":3D42
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   270
         TabIndex        =   60
         Top             =   9300
         Width           =   45
      End
   End
End
Attribute VB_Name = "fClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const FGC = "Código    | Nombre                                                          | RIF Nº                | Telefonos                                "

Dim RClientes As New ADODB.Recordset

Const OPR_NUEVO = 1
Const OPR_EDITAR = 2

Dim OPR As Integer  'operacion: 1 nuevo  2 editar
Dim Viendo As Boolean
Dim TablasCreadas As Boolean


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


  eCedAuto2.Enabled = TF
  eNomAuto2.Enabled = TF
  eCarAuto2.Enabled = TF
  eTLFAuto2.Enabled = TF
  eFotoAuto2.Enabled = TF
  
  eCedAuto3.Enabled = TF
  eNomAuto3.Enabled = TF
  eCarAuto3.Enabled = TF
  eTLFAuto3.Enabled = TF
  eFotoAuto3.Enabled = TF


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
          
      Case 10: 'bSubClientes.Enabled = TF
      
      Case 11: bAceptar.Enabled = TF
      Case 12: bCancelar.Enabled = TF
      
    End Select
    
  End If
End Sub

Private Sub Cargar_Clientes()
  'Dim r As New ADODB.Recordset
  Dim l As Integer
  
  Viendo = False
  
  Limpiar_FG
  
  FG.Clear
  FG.FormatString = FGC
  FG.Rows = 2

  
  If RClientes.State <> adStateClosed Then RClientes.Close
 'RClientes.Close
  RClientes.Open "SELECT * FROM clientes ORDER BY codigo", DBConexionSQL, adOpenDynamic, adLockOptimistic
  l = 1
  Do While Not RClientes.EOF
    FG.TextMatrix(l, 0) = Zeros(RClientes.Fields("codigo").value, 6)
    FG.TextMatrix(l, 1) = Trim(RClientes.Fields("nombre").value)
    FG.TextMatrix(l, 2) = Trim(RClientes.Fields("rif").value)
    FG.Row = l: FG.Col = 2: FG.CellAlignment = flexAlignLeftCenter
    
    FG.TextMatrix(l, 3) = Trim(RClientes.Fields("telefonos").value)
    FG.Row = l: FG.Col = 3: FG.CellAlignment = flexAlignLeftCenter
    
    RClientes.MoveNext
    
    FG.Col = 0
        
    If Not RClientes.EOF Then
      l = l + 1
      FG.Rows = FG.Rows + 1
    End If
  Loop
  
  FG.Refresh

  If FG.Rows >= 1 Then
    FG.Row = 1
    FG.Col = 0
    Scroll_Cliente
    'FG.SetFocus
  End If
    
  If Not RClientes.BOF And Not RClientes.EOF Then RClientes.MoveFirst
  
  Call Mostrar_Cliente
  
  Label2.Caption = "[" & CStr(Total_Clientes()) & " Regs.]"
  'RClientes.Close
  'Set r = Nothing
  Viendo = True
End Sub



Private Sub Scroll_Cliente()
  Dim c As String
  Viendo = False
  ''MsgBox FG.TextMatrix(FG.Rows - 1, 0)
  If FG.Row >= 1 Then
    c = Trim(FG.TextMatrix(FG.Row, 0))
    If c <> "" Then
      RClientes.MoveFirst
      RClientes.Find "codigo = " & c
      If RClientes.EOF Then
        MsgBox "Debe Seleccionar el Cliente...", vbCritical, "Información"
      Else
        Mostrar_Cliente
      End If
    Else
      Limpiar_Txts
    End If
  End If
  Viendo = True
End Sub

Private Sub Ubicar_Cursor_Cliente(sCod As String)
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

Private Sub Mostrar_Cliente()
  Dim s As String, sRutaCliente As String
  
  If RClientes.State <> adStateClosed Then
    If Not RClientes.EOF Then
      lcod.Caption = Zeros(RClientes.Fields("codigo").value, 6)
      trif.Text = Trim(RClientes.Fields("rif").value)
      tnit.Text = Trim(RClientes.Fields("nit").value)
      tnom.Text = Trim(RClientes.Fields("nombre").value)
      tdir.Text = Trim(RClientes.Fields("direccion").value)
      ttel.Text = Trim(RClientes.Fields("telefonos").value)
      tfax.Text = Trim(RClientes.Fields("fax").value)
      temail.Text = Trim(RClientes.Fields("email").value)
      tweb.Text = Trim(RClientes.Fields("website").value)
      tcon.Text = Trim(RClientes.Fields("contacto").value)
      tcontelf.Text = Trim(RClientes.Fields("contactotlf").value)
      tinicio.value = Trim(RClientes.Fields("fechainicio").value)
      oActivo.value = False
      oSuspendido.value = False
      If RClientes.Fields("activo").value = "S" Then
        oActivo.value = True
      Else
        oSuspendido.value = True
      End If
      'eInd.Text = Trim(RClientes.Fields("indicaciones").Value)
      
      eCedAuto.Text = Trim(RClientes.Fields("cedulaauto1").value)
      eNomAuto.Text = Trim(RClientes.Fields("nombreauto1").value)
      eCarAuto.Text = Trim(RClientes.Fields("cargoauto1").value)
      eTLFAuto.Text = Trim(RClientes.Fields("telefauto1").value)
      eFotoAuto.Text = Trim(RClientes.Fields("fotoauto1").value)
      
      Set Image1.Picture = Nothing
  
      s = Trim(RClientes.Fields("fotoauto1").value)
      If s <> "" Then
        If InStr(s, "/") > 0 Or InStr(s, "\") > 0 Then 'Buscar completa la cadena de FOTO
          Image1.Picture = LoadPicture(s)
        Else
          sRutaCliente = Modulo.LA_RUTA_FOTO_DEL_CLIENTE(lcod.Caption, "-")
          If sRutaCliente <> "" Then
            s = sRutaCliente & "\" & s
            If Dir(s) <> "" Then Image1.Picture = LoadPicture(s)
          End If
        End If
      End If
      
      
      
      eCedAuto2.Text = Trim(RClientes.Fields("cedulaauto2").value)
      eNomAuto2.Text = Trim(RClientes.Fields("nombreauto2").value)
      eCarAuto2.Text = Trim(RClientes.Fields("cargoauto2").value)
      eTLFAuto2.Text = Trim(RClientes.Fields("telefauto2").value)
      eFotoAuto2.Text = Trim(RClientes.Fields("fotoauto2").value)
      
      Set Image2.Picture = Nothing
  
      s = Trim(RClientes.Fields("fotoauto2").value)
      If s <> "" Then
        If InStr(s, "/") > 0 Or InStr(s, "\") > 0 Then 'Buscar completa la cadena de FOTO
          Image2.Picture = LoadPicture(s)
        Else
          sRutaCliente = Modulo.LA_RUTA_FOTO_DEL_CLIENTE(lcod.Caption, "-")
          If sRutaCliente <> "" Then
            s = sRutaCliente & "\" & s
            If Dir(s) <> "" Then Image2.Picture = LoadPicture(s)
          End If
        End If
      End If


      eCedAuto3.Text = Trim(RClientes.Fields("cedulaauto3").value)
      eNomAuto3.Text = Trim(RClientes.Fields("nombreauto3").value)
      eCarAuto3.Text = Trim(RClientes.Fields("cargoauto3").value)
      eTLFAuto3.Text = Trim(RClientes.Fields("telefauto3").value)
      eFotoAuto3.Text = Trim(RClientes.Fields("fotoauto3").value)
      
      Set Image3.Picture = Nothing
  
      s = Trim(RClientes.Fields("fotoauto3").value)
      If s <> "" Then
        If InStr(s, "/") > 0 Or InStr(s, "\") > 0 Then 'Buscar completa la cadena de FOTO
          Image3.Picture = LoadPicture(s)
        Else
          sRutaCliente = Modulo.LA_RUTA_FOTO_DEL_CLIENTE(lcod.Caption, "-")
          If sRutaCliente <> "" Then
            s = sRutaCliente & "\" & s
            If Dir(s) <> "" Then Image3.Picture = LoadPicture(s)
          End If
        End If
      End If

      
      
      lTD.Caption = Format(RClientes.Fields("deuda").value, "#,0.00")
      lTP.Caption = Format(RClientes.Fields("pagos").value, "#,0.00")
      lTS.Caption = Format(RClientes.Fields("saldo").value, "#,0.00")
      
      Label33.Caption = Modulo.CarnetsEntregados(lcod.Caption, "-", "19000101", Format(Date, "yyyymmdd"))
      
      
      
    End If
  End If
End Sub

Function CodigoClienteNuevo() As Long
  Dim r As New ADODB.Recordset
  Dim ccn As Long
  ccn = 0
  r.Open "SELECT * FROM clientes ORDER BY codigo", DBConexionSQL, adOpenKeyset, adLockReadOnly
  If Not r.EOF Then
    r.MoveLast
    ccn = r.Fields("codigo").value
  End If
  CodigoClienteNuevo = ccn + 1
End Function

Function Total_Clientes() As Long
  Dim r As New ADODB.Recordset
  Dim ccn As Long
  ccn = 0
  r.Open "SELECT count(*) FROM clientes", DBConexionSQL, adOpenKeyset, adLockReadOnly
  If Not r.EOF Then
    If Not IsNull(r.Fields(0).value) Then
      ccn = r.Fields(0).value
    End If
  End If
  r.Close
  Set r = Nothing
  Total_Clientes = ccn
End Function


Private Sub bAceptar_Click()
  Dim s As String, nc As String, s2 As String
  Dim c As Long
  Dim r As New ADODB.Recordset
  Dim sNom As String, s3 As String
  Dim xpre As Long
  Dim lCliente As String
  tnom.Text = Trim(tnom.Text)
  
  If tnom.Text = "" Or trif.Text = "" Then
    MsgBox "Faltan Datos, Revise...", vbCritical, "Información"
    'tnom.SetFocus
    Exit Sub
  End If
  
  If OPR = OPR_NUEVO Then
    If fExisteCliente(trif.Text) = True Then
       MsgBox "Este RIF ya existe en la base datos del sistema (" & trif.Text & ")", vbExclamation
       trif.SelStart = 0
       trif.SelLength = Len(trif.Text)
       trif.SetFocus
       
       Exit Sub
    End If
    If Not Nombre_Directorio_Valido(tnom.Text) Then Exit Sub
    If Nombre_Directorio_Repetido(tnom.Text, True) Then Exit Sub
  
    Load fMensaje
    fMensaje.Label1.Caption = "Añadiendo Nuevo Cliente, Espere..."
    fMensaje.Show
    DoEvents
    On Error Resume Next
      
    With RClientes
      .AddNew
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
      '.Fields("indicaciones").Value = Trim(eInd.Text)
      
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
    
    RClientes.Close
    RClientes.Open
    If Not RClientes.EOF Then RClientes.MoveLast
    c = RClientes.Fields("codigo").value 'Codigo asignado por el xSQL
    
    Unload fMensaje
    
    If Err.Number <> 0 Then
      MsgBox "Ha Ocurrido un Error al Intentar almacenar el Registro..." & vbCrLf & Err.Description, vbCritical, "Información"
      Exit Sub
    End If
    
    
    AgregarLogs "Agrega Nuevo Cliente [" & Mid(Trim(tnom.Text), 1, 20) & "...]"
    
    
    '-----------------------------------------------------------
    '- CREAR LAS CARPETAS Y COPIAR LAS PLANTILLAS
    '-----------------------------------------------------------
    
    Load fMensaje
    fMensaje.Label1.Caption = "Creando Carpetas/Plantillas del Cliente y Sub-Clientes de [" & Zeros(c, 6) & "] , Espere..."
    fMensaje.Show
    DoEvents

    Dim sOri As String
    Dim sDes As String
      
    sOri = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
    sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")

    
    
    
    If sOri <> "" And sDes <> "" Then
    
      '-- Nombre del Cliente:
      s = sDes & "\" & tnom.Text
      
      MkDir (s)
      
        '-- Subcarpetas: Carnets - Fotos - Imagenes
        
        s = sDes & "\" & tnom.Text & "\" & "CARNET"
        MkDir (s)
        
        s = sDes & "\" & tnom.Text & "\" & "FOTOS"
        MkDir (s)
        
        s = sDes & "\" & tnom.Text & "\" & "IMAGENES"
        MkDir (s)
        
       ''' '-- Copiar archivos (RAIZ):
        s = sOri & "\FORMATO DE ENTREGA DE CARNETS _nombre cliente.xls"
        s2 = sDes & "\" & tnom.Text & "\FORMATO DE ENTREGA DE CARNETS " & tnom.Text & ".xls"
        FileCopy s, s2
        
        's = sOri & "\LISTADO NOMBRE CLIENTE OFICINA.xls"
        's2 = sDes & "\" & tnom.Text & "\LISTADO " & tnom.Text & ".xls"
        'FileCopy s, s2
        
        '-- Copiar archivos (CARNET):
        s = sOri & "\CARNET\BASE CARNET NOMBRE CLIENTE.car"
        s2 = sDes & "\" & tnom.Text & "\CARNET\BASE CARNET " & tnom.Text & ".car"
        FileCopy s, s2
        
        s = sOri & "\CARNET\BASE DATOS NOMBRE CLIENTE.mdb"
        s2 = sDes & "\" & tnom.Text & "\CARNET\BASE DATOS " & tnom.Text & "_" & Year(Now) & ".mdb"
        FileCopy s, s2
        
        '-- Copiar archivos (IMAGENES):
        ''s = sOri & "\IMAGENES\Autocopia_de_seguridad_deBASE nombre cliente.cdr"
        ''s2 = sDes & "\" & tnom.Text & "\IMAGENES\Autocopia_de_seguridad_de " & tnom.Text & ".cdr"
        ''FileCopy s, s2
        
        s = sOri & "\IMAGENES\BASE nombre cliente.cdr"
        s2 = sDes & "\" & tnom.Text & "\IMAGENES\BASE " & tnom.Text & ".cdr"
        FileCopy s, s2
               
    End If
    Unload fMensaje
    TablasCreadas = False
    If MsgBox("¿Desea Crear las Tablas del Cliente en la base de datos ahora?", vbQuestion + vbYesNo) = vbYes Then
       TablasCreadas = True
       Load fMensaje
       fMensaje.Label1.Caption = "Creando Tablas del Cliente de [" & Zeros(c, 6) & "] , Espere..."
       fMensaje.Show
       DoEvents
       Dim lCommand As New ADODB.Command
       lCommand.ActiveConnection = DBConexionSQL.ConnectionString
       lCommand.CommandType = adCmdText
       lCommand.CommandTimeout = 15
       lCommand.CommandText = "exec CrearTablasCliente"
       lCommand.Execute
       sCrearArchivoExcel CStr(Right("000000" & c, 6)) & "-1", tnom.Text
    End If
    '-----------------------------------------------------------------------
    '- Crear la cuenta del Cliente, si tiene algun monto previo para poderlo
    '- cargar.
    '-----------------------------------------------------------------------
       
    'Load fCantidad
    '
    ' With fCantidad
    '  .ePre.Text = "0,0"
    '
    'End With
    'fCantidad.Show vbModal
    
    'Modulo.Crear_Cuenta_Cliente
    
    
    
    
    
    Unload fMensaje
            
    Cargar_Clientes
    Scroll_Cliente
    Limpiar_Txts
    lcod.Caption = Zeros(c + 1, 6)
    
    trif.SetFocus
    lCliente = FG.TextMatrix(FG.Rows - 1, 0)
    If TablasCreadas = True Then
       If MsgBox("¿Desea Editar la Tabla del Cliente?", vbYesNo + vbQuestion) = vbYes Then
          fPersonas.cCP.Text = fPersonas.cCP.List(fPersonas.cCP.ListCount - 1)  '"000068-1 :4TA PERUEWBA"
          fPersonas.Show vbModal
          'lCliente = fPersonas.cCP.List(fPersonas.cCP.ListCount - 1)
       End If
    End If
    'Load fPagos
    ''fPagos.Show
    'lCliente = Mid(lCliente, 1, 6)
    Load fMensaje
    fMensaje.Label1.Caption = "Cargando Módulo de Pagos Para el Cliente Nuevo [" & Zeros(c, 6) & "] , Espere..."
    fMensaje.Show
    DoEvents
    fPagos.Adodc1.Refresh
    fPagos.Adodc1.Recordset.MoveFirst
    fPagos.Adodc1.Recordset.Find "Codigo=" & lCliente, 1, adSearchForward
    If fPagos.Adodc1.Recordset.EOF = False Then
       fPagos.DataGrid1_Click
       fPagos.Show vbModal
    End If
    Unload fMensaje
    'LLAMAR A LA FORMA "DISEÑO"
    fMensaje.Label1.Caption = "Cargando Módulo de Formato de Diseño Para el Cliente Nuevo [" & Zeros(c, 6) & "] , Espere..."
    fMensaje.Show
    DoEvents
    
    Load fFD
    fFD.cCP.Text = lCliente & " : " & Trim(UCase(FG.TextMatrix(FG.Rows - 1, 1)))
    fFD.sCargarDatosCliente
    fFD.Show vbModal
    Unload fMensaje
  Else
  
    If OPR = OPR_EDITAR Then
      
      If fExisteCliente(trif.Text) = True Then
         MsgBox "Este RIF ya existe en la base datos del sistema (" & trif.Text & ")", vbExclamation
         trif.SelStart = 0
         trif.SelLength = Len(trif.Text)
         trif.SetFocus
         bEditar_Click
         Exit Sub
      End If

      Load fMensaje
      fMensaje.Label1.Caption = "Actualizando Cliente, Espere..."
      fMensaje.Show
      '.Unload fMensaje
      DoEvents
      On Error Resume Next
      
      nc = lcod.Caption
      
      s = "UPDATE clientes SET " & _
          "rif         = '" & Trim(trif.Text) & "'," & "nit        = '" & Trim(tnit.Text) & "'," & _
          "nombre      = '" & Trim(tnom.Text) & " '," & "direccion = '" & Trim(tdir.Text) & "'," & _
          "telefonos   = '" & Trim(ttel.Text) & "'," & "fax        = '" & Trim(tfax.Text) & "'," & _
          "email       = '" & Trim(temail.Text) & "'," & "website  = '" & Trim(tweb.Text) & "'," & _
          "contacto    = '" & Trim(tcon.Text) & "'," & "contactotlf = '" & Trim(tcontelf.Text) & "'," & _
          "fechainicio = '" & Format(tinicio.value, "yyyyMMdd") & "'," & _
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
          "codigo = " & nc
          Clipboard.Clear
          Clipboard.SetText s
      Set DBComandoSQL.ActiveConnection = DBConexionSQL
      DBComandoSQL.CommandText = s
      DBComandoSQL.Execute
      
      Unload fMensaje
      
      If Err.Number <> 0 Then
        MsgBox "Ha Ocurrido un Error al Intentar actualizar el Registro..." & vbCrLf & Err.Description, vbCritical, "Información"
        Exit Sub
      End If
      
      AgregarLogs "Actualiza Cliente [" & Mid(Trim(tnom.Text), 1, 20) & "...]"
      
      
      Cargar_Clientes
            
      bCancelar_Click
      
      Ubicar_Cursor_Cliente nc
    
    End If

  End If
 '**************************************************************
 ''Restaurar Botonoes
   OPR = 0
  
  fClientes.Caption = "CLIENTES"
  
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
  
  Scroll_Cliente
  
  FG.SetFocus
  SendKeys "{LEFT}"
''************************************************************

End Sub



Private Sub bBorrar_Click()
  Dim c As String, n As String, s As String
  Dim s2 As String
  Dim r As New ADODB.Recordset
  
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
      
      If MsgBox("¿Está Seguro de Borrar el Cliente " & vbCrLf & c & "-" & n & vbCrLf & _
                "y toda la información asociada (sub-clientes, movimientos," & vbCrLf & _
                "personas, cuentas y carpetas ?", vbQuestion + vbYesNo, "Confirme") = vbYes Then
                
        'MsgBox "Indique Clave ADMINISTRADOR...", vbExclamation, "Confirme"
        
        
        
        '-- Borrar TODO: carpetas Fisicas del cliente:
        
        s = "RMDIR " & Chr(34) & sDes & "\" & n & Chr(34) & " /S /Q"
        
        s = Ejecutar_DOS(s)
               
        Load fMensaje
        fMensaje.Label1.Caption = "Borrando Sub-clientes, Espere..."
        fMensaje.Show
                
'''''        s = "SELECT * FROM Subclientes WHERE cliente = " & c & " ORDER BY id"
'''''        r.Open s, Modulo.DBConexionSQL, adOpenKeyset, adLockReadOnly
'''''        Do While Not r.EOF
'''''          s = Trim(r.Fields("nombre").Value)
'''''          s2 = s
'''''
'''''          '-- Ir dentro de la carpeta del sub-cliente y borrar todos los
'''''          '-- archivos:
'''''          s = sDes & "\" & n & "\" & s2
'''''          File1.Path = s
'''''          File1.Refresh
'''''          If File1.ListCount > 0 Then Kill s & "\*.*"
'''''
'''''          s = sDes & "\" & n & "\" & s2 & "\CARNET"
'''''          File1.Path = s
'''''          File1.Refresh
'''''          If File1.ListCount > 0 Then Kill s & "\*.*"
'''''
'''''          s = sDes & "\" & n & "\" & s2 & "\FOTOS"
'''''          File1.Path = s
'''''          File1.Refresh
'''''          If File1.ListCount > 0 Then Kill s & "\*.*"
'''''
'''''          s = sDes & "\" & n & "\" & s2 & "\IMAGENES"
'''''          File1.Path = s
'''''          File1.Refresh
'''''          If File1.ListCount > 0 Then Kill s & "\*.*"
'''''
'''''          '-- Ir dentro de la carpeta del sub-cliente y borrar las
'''''          '-- sub-carpetas:
'''''          s = sDes & "\" & n & "\" & s2 & "\CARNET"
'''''          RmDir s
'''''
'''''          s = sDes & "\" & n & "\" & s2 & "\FOTOS"
'''''          RmDir s
'''''
'''''          s = sDes & "\" & n & "\" & s2 & "\IMAGENES"
'''''          RmDir s
'''''
'''''          '-- Borrar la carpeta del Cliente:
'''''          s = sDes & "\" & n & "\" & s2
'''''          RmDir s
'''''
'''''          '-- Borrar la carpeta del Cliente:
'''''          's = sDes & "\" & n & "\" & s2
'''''          'RmDir s
'''''
'''''          r.MoveNext
'''''        Loop
'''''        r.Close
'''''        Set r = Nothing
'''''
'''''        '-------------------------------------------------
'''''        '-- Borrar Carpetas del Cliente Principal
'''''        '-------------------------------------------------
'''''
'''''        s = sDes & "\" & n
'''''        File1.Path = s
'''''        File1.Refresh
'''''        If File1.ListCount > 0 Then Kill s & "\*.*"
'''''
'''''        s = sDes & "\" & n & "\CARNET"
'''''        File1.Path = s
'''''        File1.Refresh
'''''        If File1.ListCount > 0 Then Kill s & "\*.*"
'''''
'''''        s = sDes & "\" & n & "\FOTOS"
'''''        File1.Path = s
'''''        File1.Refresh
'''''        If File1.ListCount > 0 Then Kill s & "\*.*"
'''''
'''''        s = sDes & "\" & n & "\IMAGENES"
'''''        File1.Path = s
'''''        File1.Refresh
'''''        If File1.ListCount > 0 Then Kill s & "\*.*"
'''''
'''''        s = sDes & "\" & n & "\CARNET"
'''''        RmDir s
'''''        s = sDes & "\" & n & "\FOTOS"
'''''        RmDir s
'''''        s = sDes & "\" & n & "\IMAGENES"
'''''        RmDir s
'''''        s = sDes & "\" & n
'''''        RmDir s
          
        '------------------------------------------------------
        '-- Borrar Sub-Clientes y Cliente Principal de la B.D.
        '------------------------------------------------------
       
        s = "DELETE FROM subclientes WHERE cliente = " & c
        Set DBComandoSQL.ActiveConnection = DBConexionSQL
        DBComandoSQL.CommandText = s
        DBComandoSQL.Execute

        fMensaje.Label1.Caption = "Borrando Cliente, Espere..."
        fMensaje.Show
                                        
        s = "DELETE FROM clientes WHERE codigo = " & c
        Set DBComandoSQL.ActiveConnection = DBConexionSQL
        DBComandoSQL.CommandText = s
        DBComandoSQL.Execute
        
        Unload fMensaje
        
        AgregarLogs "Borra Cliente [" & Mid(n, 1, 20) & "...]"
                
        MsgBox "CLIENTE [" & c & "] FUE BORRADO (" & Modulo.USUARIO_ACTUAL & " " & Format(Now, "dd/mm/yy hh:mm ampm") & ")", vbInformation, "Información"
        
        'FALTA LOGs
        Cargar_Clientes  'Abre el RecordSet de Clientes
        Scroll_Cliente
      End If
    End If
  End If
End Sub

Private Sub Buscar_Cursor_Cliente(sCod As String)
  Dim i As Integer
  Dim f As Integer
  Dim e As Boolean
  
  i = 1
  If FG.Rows >= 1 Then
    e = False
    Do While i < FG.Rows And Not e
      FG.Row = i
      If FG.TextMatrix(i, 0) = sCod Then
        e = True
      Else
        i = i + 1
      End If
    Loop
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
    Scroll_Cliente
    
  End If
  
  
End Sub

Public Sub bEditar_Click()
  If FG.Row >= 1 Then
  
    If Trim(FG.TextMatrix(FG.Row, 0)) = "" Then Exit Sub
    
  Else
  
    Exit Sub
    
  End If
  
  OPR = OPR_EDITAR
  
  fClientes.Caption = "EDITAR CLIENTE"
  
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
  lcod.Caption = Trim(FG.TextMatrix(FG.Row, 0))
  
  'trif.SetFocus
  
  
End Sub

Private Sub bNuevo_Click()
  Dim sOri As String
  Dim sDes As String
    
  sOri = GetSetting(APPNAME, "Opciones", "RutaOrigen", "")
  sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
  
  If sOri = "" Or sDes = "" Then
    MsgBox "Debe Especificar primero las Carpetas donde se encuentran las " & vbCrLf & _
           "Plantillas y donde se almacenarán los datos de los clientes." & vbCrLf & _
           "Puede ir a Menú Principal -> Sistema -> Opciones", vbCritical, "Información"
    Exit Sub
  End If
  
  bUltimo_Click
  DoEvents

  OPR = OPR_NUEVO
  
  fClientes.Caption = "NUEVO CLIENTE"
  
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
  
  lcod.Caption = "Nuevo"
  
  
  lTD.Caption = "0,00"
  lTP.Caption = "0,00"
  lTS.Caption = "0,00"

  
  trif.SetFocus
  
     
  
  
End Sub

Private Sub bCancelar_Click()
  Dim s As String
  
  If OPR = OPR_NUEVO Then
    '//Borrar los sub-clientes en caso de que hubieran...
    s = "DELETE FROM subclientes WHERE cliente = 0"
    Set DBComandoSQL.ActiveConnection = DBConexionSQL
    DBComandoSQL.CommandText = s
    DBComandoSQL.Execute
  End If

  OPR = 0
  
  fClientes.Caption = "CLIENTES"
  
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
  
  Scroll_Cliente
  
  FG.SetFocus
  SendKeys "{LEFT}"

  
  

End Sub

Private Sub bPrimer_Click()
  'Ir al primer registro del listado
  FG.Row = 1
  Scroll_Cliente
  FG.SetFocus
  SendKeys "{UP}"
End Sub

Private Sub bAnterior_Click()
  'Ir al anterior registro del listado
  If FG.Rows > 2 Then
    FG.Row = FG.Row - 1
    Scroll_Cliente
  End If
  FG.SetFocus
  SendKeys "{LEFT}"
End Sub

Private Sub bSiguiente_Click()
  'Ir al siguiente registro del listado
  If FG.Rows > 2 Then
    If FG.Row <= FG.Rows - 1 Then
      FG.Row = FG.Row + 1
      Scroll_Cliente
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
      Scroll_Cliente
    End If
  End If
  FG.SetFocus
  SendKeys "{DOWN}"
End Sub

Private Sub bSalir_Click()
  Unload Me
End Sub

Private Sub Command3_Click()
'sCrearArchivoExcel
End Sub

Private Sub Check1_Click()
 If Check1.value = 1 Then
    eTLFAuto.Text = tcontelf.Text
    eNomAuto.Text = tcon.Text
 Else
    eTLFAuto.Text = ""
    eNomAuto.Text = ""
 End If
End Sub

Private Sub eCarAuto_Change()
  'eCarAuto.Text = UCase(eCarAuto.Text)
End Sub

Private Sub eCarAuto_LostFocus()
  eCarAuto.Text = UCase(eCarAuto.Text)
End Sub

Private Sub eCarAuto2_LostFocus()
  eCarAuto2.Text = UCase(eCarAuto2.Text)
End Sub

Private Sub eCarAuto3_LostFocus()
  eCarAuto3.Text = UCase(eCarAuto3.Text)
End Sub

Private Sub eCedAuto_Change()
  Dim s As String, d As Double
  
  eCedAuto.Text = UCase(eCedAuto.Text)
  If IsNumeric(eCedAuto.Text) Then
    d = CDbl(eCedAuto.Text)
    s = Format(d, "#,0")
    eCedAuto.Text = s
    If eCedAuto.Enabled Then eCedAuto.SelStart = Len(eCedAuto.Text)
    eFotoAuto.Text = CStr(d) & ".JPG"
  Else
    eFotoAuto.Text = eCedAuto.Text & ".JPG"
  End If
  
  If eFotoAuto.Text = ".JPG" Then eFotoAuto.Text = ""
  
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
        eFotoAuto.Text = CStr(d)
      Else
        eFotoAuto.Text = eCedAuto.Text
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

Private Sub eCedAuto2_Change()
  Dim s As String, d As Double
  eCedAuto2.Text = UCase(eCedAuto2.Text)
  If IsNumeric(eCedAuto2.Text) Then
    d = CDbl(eCedAuto2.Text)
    s = Format(d, "#,0")
    eCedAuto2.Text = s
    If eCedAuto2.Enabled Then SendKeys "{END}"
            
    eFotoAuto2.Text = CStr(d) & ".JPG"
  Else
    eFotoAuto2.Text = eCedAuto2.Text & ".JPG"
  End If
  If eFotoAuto2.Text = ".JPG" Then eFotoAuto2.Text = ""
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
        eFotoAuto2.Text = CStr(d)
      Else
        eFotoAuto2.Text = eCedAuto2.Text
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



Private Sub eCedAuto3_Change()
  Dim s As String, d As Double
  eCedAuto3.Text = UCase(eCedAuto3.Text)
  If IsNumeric(eCedAuto3.Text) Then
    d = CDbl(eCedAuto3.Text)
    s = Format(d, "#,0")
    eCedAuto3.Text = s
    If eCedAuto3.Enabled Then SendKeys "{END}"
    eFotoAuto3.Text = CStr(d) & ".JPG"
  Else
    eFotoAuto3.Text = eCedAuto3.Text & ".JPG"
  End If
  If eFotoAuto3.Text = ".JPG" Then eFotoAuto3.Text = ""
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
        eFotoAuto3.Text = CStr(d)
      Else
        eFotoAuto3.Text = eCedAuto3.Text
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

Private Sub eFotoAuto_Change()
  Dim s As String, sRutaCliente As String
  s = Trim(eFotoAuto.Text)
  If s <> "" Then
    If InStr(s, "/") > 0 Or InStr(s, "\") > 0 Then 'Buscar completa la cadena de FOTO
      Image1.Picture = LoadPicture(s)
    Else
      sRutaCliente = Modulo.LA_RUTA_FOTO_DEL_CLIENTE(lcod.Caption, "-")
      If sRutaCliente <> "" Then
        s = sRutaCliente & "\" & s
        If Dir(s) <> "" Then Image1.Picture = LoadPicture(s) Else Set Image1.Picture = Nothing
      End If
    End If
  End If
End Sub

Private Sub eFotoAuto2_Change()
  Dim s As String, sRutaCliente As String
  s = Trim(eFotoAuto2.Text)
  If s <> "" Then
    If InStr(s, "/") > 0 Or InStr(s, "\") > 0 Then 'Buscar completa la cadena de FOTO
      Image2.Picture = LoadPicture(s)
    Else
      sRutaCliente = Modulo.LA_RUTA_FOTO_DEL_CLIENTE(lcod.Caption, "-")
      If sRutaCliente <> "" Then
        s = sRutaCliente & "\" & s
        If Dir(s) <> "" Then Image2.Picture = LoadPicture(s) Else Set Image2.Picture = Nothing
      End If
    End If
  End If
End Sub

Private Sub eFotoAuto3_Change()
  Dim s As String, sRutaCliente As String
  s = Trim(eFotoAuto3.Text)
  If s <> "" Then
    If InStr(s, "/") > 0 Or InStr(s, "\") > 0 Then 'Buscar completa la cadena de FOTO
      Image3.Picture = LoadPicture(s)
    Else
      sRutaCliente = Modulo.LA_RUTA_FOTO_DEL_CLIENTE(lcod.Caption, "-")
      If sRutaCliente <> "" Then
        s = sRutaCliente & "\" & s
        If Dir(s) <> "" Then Image3.Picture = LoadPicture(s) Else Set Image3.Picture = Nothing
      End If
    End If
  End If
End Sub

Private Sub eNomAuto_Change()
  EnMayusculas eNomAuto
End Sub

Private Sub eNomAuto2_Change()
  EnMayusculas eNomAuto2
End Sub

Private Sub eNomAuto3_Change()
  EnMayusculas eNomAuto3
End Sub

Private Sub eTLFAuto_Change()
  'eTLFAuto.Text = UCase(eTLFAuto.Text)
End Sub

Private Sub eTLFAuto_LostFocus()
  eTLFAuto.Text = UCase(eTLFAuto.Text)
End Sub

Private Sub eTLFAuto2_Change()
  'eTLFAuto2.Text = UCase(eTLFAuto2.Text)
End Sub

Private Sub eTLFAuto3_Change()
  'eTLFAuto3.Text = UCase(eTLFAuto3.Text)
End Sub

Private Sub FG_Click()
  Scroll_Cliente
End Sub

Private Sub FG_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyHome Then bPrimer_Click Else
  If KeyCode = vbKeyEnd Then bUltimo_Click
End Sub

Private Sub FG_RowColChange()
  If Viendo Then Scroll_Cliente
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

Private Sub Form_Load()
  Viendo = False
  OPR = 0
  
  SSTab1.Tab = 0
  
  Cargar_Clientes  'Abre el RecordSet de Clientes
  
  Limpiar_Txts
  Activar_Txts False
  Activar_Btns 0, False
  Activar_Btns 9, True
  lcod.Enabled = False
  
  Mostrar_Cliente
  
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
  
  lTD.Caption = "0,00"
  lTP.Caption = "0,00"
  lTS.Caption = "0,00"
  
  
  
  
  
  SendKeys "{UP}"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If RClientes.State <> adStateClosed Then
    RClientes.Close
    Set RClientes = Nothing
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

Private Sub tnit_Change()
  EnMayusculas tnit
End Sub

Private Sub tnom_Change()
  EnMayusculas tnom
End Sub

Private Sub tnom_KeyPress(KeyAscii As Integer)
  If KeyAscii = Asc("\") Or _
     KeyAscii = Asc("/") Or _
     KeyAscii = Asc(":") Or _
     KeyAscii = Asc("*") Or _
     KeyAscii = Asc("?") Or _
     KeyAscii = Asc("'") Or _
     KeyAscii = Asc("<") Or _
     KeyAscii = Asc(">") Or _
     KeyAscii = Asc("|") Then
    KeyAscii = 0
  End If
  
End Sub

Private Sub trif_Change()
  EnMayusculas trif
End Sub

Private Sub trif_LostFocus()
    'If fExisteCliente(trif.Text) = True Then
    '   MsgBox "Este RIF ya existe en la base datos del sistema (" & trif.Text & ")", vbExclamation
    '   trif.SelStart = 0
    '   trif.SelLength = Len(trif.Text)
    '   trif.SetFocus
    'End If

End Sub

Private Sub ttel_Change()
    EnMayusculas ttel
End Sub

Private Sub tweb_Change()
    EnMayusculas tweb
End Sub
