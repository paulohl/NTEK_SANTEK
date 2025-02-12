VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPreviewCarnetsImpresos 
   Caption         =   "Listado de Carnets Impresos"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12435
   LinkTopic       =   "Form2"
   ScaleHeight     =   5715
   ScaleWidth      =   12435
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   4275
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   4515
      Begin VB.TextBox eEmail 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   270
         Width           =   3375
      End
      Begin VB.TextBox eTitulo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   690
         Width           =   3375
      End
      Begin VB.TextBox eMensaje 
         BackColor       =   &H00FFFFFF&
         Height          =   1545
         Left            =   960
         MaxLength       =   254
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "frmPreviewCarnetsImpresos.frx":0000
         Top             =   1080
         Width           =   3375
      End
      Begin VB.CommandButton cmdEnviar 
         Caption         =   "Enviar"
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   3660
         Width           =   1155
      End
      Begin VB.ListBox ListAdjunto 
         Height          =   645
         Left            =   960
         TabIndex        =   4
         Top             =   2760
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail:"
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
         Left            =   270
         TabIndex        =   12
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titulo:"
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
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje:"
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
         Left            =   45
         TabIndex        =   10
         Top             =   1110
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Adjunto:"
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
         Left            =   180
         TabIndex        =   9
         Top             =   2730
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   540
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   1270
         ButtonWidth     =   1614
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Imprimir"
               Key             =   "kimprimir"
               Description     =   "Imprime el Informe"
               Object.ToolTipText     =   "Imprimir el Informe"
               ImageKey        =   "imprimir"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Seleccionar"
               Key             =   "ksetuprinter"
               Description     =   "Seleccionar Impresora"
               Object.ToolTipText     =   "Seleccionar Impresora"
               ImageKey        =   "printsetup"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cerrar"
               Key             =   "ksalir"
               Description     =   "Cerrar Ventana"
               Object.ToolTipText     =   "Cerrar Ventana"
               ImageKey        =   "salir"
            EndProperty
         EndProperty
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   5385
      Left            =   4560
      TabIndex        =   0
      Top             =   180
      Width           =   7665
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5100
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewCarnetsImpresos.frx":009C
            Key             =   "buscar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewCarnetsImpresos.frx":07AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewCarnetsImpresos.frx":1088
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewCarnetsImpresos.frx":1962
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewCarnetsImpresos.frx":2074
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewCarnetsImpresos.frx":2786
            Key             =   "imprimir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewCarnetsImpresos.frx":2E98
            Key             =   "dollar1"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewCarnetsImpresos.frx":347F
            Key             =   "dollar2"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreviewCarnetsImpresos.frx":3882
            Key             =   "printsetup"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPreviewCarnetsImpresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Report2 As CrystalReport2
Dim s4 As String
Dim s5 As String
Dim s6 As String
Private Sub cmdEnviar_Click()
    Dim lCarpeta As String
    Dim sDes As String
    sDes = GetSetting(APPNAME, "Opciones", "RutaDestino", "")
    sDes = sDes & "\" & Trim(Mid(fRep01.cCP.Text, 9))
     If (fRep01.cSC.Text <> "") Then
        lCarpeta = Trim(Mid(fRep01.cSC.Text, 9) & "-")
        sDes = sDes & "\" & lCarpeta
     Else
        lCarpeta = ""
     End If
     
    ''lCarpeta = IIf(fPersonasAct.cSC.Text <> "", Trim(Mid(fPersonasAct.cSC.Text, 9)))
    Report2.Picture1.Suppress = True
    lCarpeta = sDes & lCarpeta & "\ListaCarnets-" & Trim(Mid(fRep01.cCP.Text, 9) & "-") & lCarpeta & Format(Now, "dd-MM-yyyy HHmmam/pm") & ".xls"
    Report2.ExportOptions.DestinationType = crEDTDiskFile
    Report2.ExportOptions.PDFExportAllPages = True
    Report2.ExportOptions.FormatType = crEFTExcel80    ''crEFTPortableDocFormat
    Report2.ExportOptions.DiskFileName = lCarpeta
    Report2.Export (False)
    ListAdjunto.AddItem lCarpeta
  If MandaMail(Trim(eEmail.Text), "", "", Trim(eTitulo.Text), Trim(eMensaje.Text), ListAdjunto) = True Then
    Unload Me
  End If
   Unload Me


End Sub

Public Sub sPrepararEmail(argCodigoCliente As String, argCodigoSubCliente As String)
  Dim s1 As String
  Dim s2 As String
  Dim s3 As String
  
  eEmail.Text = LCase(Modulo.Correo_E(argCodigoCliente, argCodigoSubCliente))
  
  s1 = ""
  s2 = ""
  s3 = ""
  s4 = ""
  s5 = ""
  s6 = ""
  
  CargarOpcionesCorreo s1, s2, s3, s4, s5, s6
  eTitulo.Text = "Listado de Carnets"
  'eMensaje.Text = "Estimado cliente, anexo le envío el listado de carnets." _
   '& "" _
  ' & "Saludos," _
  ' & "NTEK, C.A."

  '"Estimado cliente, anexo le envío el listado de carnets." & Chr(10) & "Saludos," & Chr(10) & "NTEK, C.A."

End Sub

Private Sub Form_Load()
Dim lReg As New ADODB.Recordset
sPrepararEmail Trim(Mid(fRep01.cCP.Text, 1, 6)), Trim(Mid(fRep01.cSC.Text, 1, 6))
  ''sCargarReporteCarnetsImpresos
End Sub
Public Sub sCargarReporteCarnetsImpresos(argCliente As String, argSubCliente As String, argFechaDesde As String, argFechaHasta As String, argSqltxt As String)
Dim lReg As New ADODB.Recordset
  Set Report2 = New CrystalReport2
 On Error GoTo falla
'lReg.Open argSqltxt, Modulo.DBConexionSQL, adOpenKeyset
'If lReg.EOF = True Then
' MsgBox "Información no encontrada", vbCritical
'Exit Sub
'End If
Set lReg = DBConexionSQL.Execute("Select * from Reporte01 where Estacion=" & Modulo.ESTACION & " Order by Impreso")
If lReg.EOF = True Then Exit Sub

Screen.MousePointer = vbHourglass
Screen.MousePointer = vbDefault
 
Report2.Database.SetDataSource lReg, 3
Report2.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
Report2.PaperSize = crPaperLetterSmall
Report2.PaperOrientation = crPortrait
Report2.txtCliente.SetText argCliente
Report2.txtSubCliente.SetText argSubCliente
Report2.txtFechaDesde.SetText Format(argFechaDesde, "dd/MM/yyyy")
Report2.txtFechaHasta.SetText Format(argFechaHasta, "dd/MM/yyyy")
CRViewer1.ReportSource = Report2
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
CRViewer1.Zoom (100)
CRViewer1.EnablePrintButton = False
sCuadrarForma
CRViewer1.EnableGroupTree = False
falla:
  If Err.Number <> 0 Then MsgBox Err.Number & "::" & Err.Description, vbCritical

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Key
      Case "kimprimir"
         Report2.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
         CRViewer1.PrintReport
      Case "ksetuprinter"
         frmSetupPrinter.Frame_imprimir.Visible = False
         frmSetupPrinter.Frame_Seleccionar.Visible = True
         frmSetupPrinter.Show vbModal
      Case "ksalir"
         Unload Me
   End Select
End Sub


Private Sub Form_Resize()
CRViewer1.Top = 0
'''CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
If ScaleWidth > 0 Then CRViewer1.width = ScaleWidth - 4500

End Sub

Private Sub sCuadrarForma()
    Me.Left = 1
    Me.Top = 1
    Me.width = 11890
    Me.Height = 6690
    '''Report.Text8.SetText ("Hola")
    
End Sub

