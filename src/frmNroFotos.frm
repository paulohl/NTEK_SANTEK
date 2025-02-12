VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNroFotos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Numeración de Fotos"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCedula 
      Height          =   285
      Left            =   4680
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   3240
      TabIndex        =   0
      Top             =   1260
      Width           =   1335
   End
   Begin VB.TextBox txtNroFoto 
      Height          =   405
      Left            =   2220
      TabIndex        =   5
      Top             =   1260
      Width           =   390
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   405
      Left            =   2610
      TabIndex        =   4
      Top             =   1260
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   714
      _Version        =   393216
      BuddyControl    =   "txtNroFoto"
      BuddyDispid     =   196611
      OrigLeft        =   3000
      OrigTop         =   1260
      OrigRight       =   3240
      OrigBottom      =   1695
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Número de Foto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblApellido 
      Caption         =   "Appellido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   660
      Width           =   3975
   End
   Begin VB.Label lblNombre 
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   540
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmNroFotos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
   DBConexionSQL.Execute "Update [" & fPersonasAct.lTablas.List(fPersonasAct.lTablas.ListIndex) & "] set NroFoto=" & txtNroFoto.Text & " where cedula='" & txtCedula & "'"
   frmFotosEnSitio.sUltimoNumero
   frmFotosEnSitio.sAgregarEnFotosNumeradas txtNroFoto.Text, txtCedula.Text, lblNombre.Caption
   If frmBuscarEnSitio.Visible = True Then Unload frmBuscarEnSitio
   Unload Me
End Sub

