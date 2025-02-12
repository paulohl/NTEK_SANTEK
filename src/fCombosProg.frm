VERSION 5.00
Begin VB.Form fCombosProg 
   Caption         =   "Combos de Productos"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton Command2 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10290
         TabIndex        =   15
         Top             =   5370
         Width           =   555
      End
      Begin VB.CommandButton Command1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10290
         TabIndex        =   14
         Top             =   2640
         Width           =   555
      End
      Begin VB.CommandButton bCancelar 
         Caption         =   "Cancelar"
         Height          =   500
         Left            =   5760
         Picture         =   "fCombosProg.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5730
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton bAceptar 
         Caption         =   "Aceptar"
         Height          =   500
         Left            =   4620
         Picture         =   "fCombosProg.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5730
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton bAnexar2 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5190
         TabIndex        =   10
         Top             =   3720
         Width           =   555
      End
      Begin VB.ListBox lC2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   5910
         TabIndex        =   9
         Top             =   3660
         Width           =   5025
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   5910
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3270
         Width           =   1635
      End
      Begin VB.ListBox lC1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   5910
         TabIndex        =   6
         Top             =   960
         Width           =   5025
      End
      Begin VB.CommandButton bAnexar1 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5160
         TabIndex        =   5
         Top             =   1020
         Width           =   555
      End
      Begin VB.ListBox lProductos 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5010
         Left            =   60
         TabIndex        =   3
         Top             =   390
         Width           =   5025
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   5910
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   570
         Width           =   1635
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripción"
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   150
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Productos que Agrupa:"
         Height          =   195
         Left            =   8190
         TabIndex        =   7
         Top             =   3420
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código"
         Height          =   255
         Left            =   90
         TabIndex        =   4
         Top             =   150
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Productos que Agrupa:"
         Height          =   195
         Left            =   8160
         TabIndex        =   1
         Top             =   720
         Width           =   1635
      End
   End
End
Attribute VB_Name = "fCombosProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CargarProductos()
  Dim r As New ADODB.Recordset
  Dim s As String
  s = "select * from Productos order by Codigo"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  lProductos.Clear
  Do While Not r.EOF
    s = r.Fields("codigo").Value & " " & r.Fields("descripcion").Value
    lProductos.AddItem s
    r.MoveNext
  Loop
  If lProductos.ListCount > 0 Then lProductos.ListIndex = 0
End Sub

Private Sub CargarCombo(sNumero As String)
  Dim r As New ADODB.Recordset
  Dim s As String
  s = "select * from CombosProgramados where ComboNumero = '" & sNumero & "' order by id"
  r.Open s, Modulo.DBConexionSQL, adOpenStatic, adLockReadOnly
  If sNumero = "1" Then lC1.Clear Else
  If sNumero = "2" Then lC2.Clear
  
  Do While Not r.EOF
    s = r.Fields("codigoproducto").Value & " " & NomProd(Trim(r.Fields("codigoproducto").Value))
    If sNumero = "1" Then lC1.AddItem s Else
    If sNumero = "2" Then lC2.AddItem s
    
    r.MoveNext
  Loop
  r.Close
  Set r = Nothing
End Sub

Private Function NomProd(sCodigoPro As String) As String
  Dim i As Integer
  Dim s As String, s1 As String
  i = 0
  s = ""
  Do While i < lProductos.ListCount And s = ""
    s1 = Trim(Mid(lProductos.List(i), 1, 20))
    If s1 = sCodigoPro Then s = Trim(Mid(lProductos.List(i), 21))
    i = i + 1
  Loop
  NomProd = s
End Function

Private Function ExisteEnCombo(sNumeroCombo As String, sItem As String) As Boolean
  Dim i As Integer, k As Integer
  Dim s As String, s1 As String
  Dim e As Boolean
  If sNumeroCombo = "1" Then k = lC1.ListCount Else
  If sNumeroCombo = "2" Then k = lC2.ListCount
  i = 0
  s = ""
  e = False
  Do While i < k And Not e
    If sNumeroCombo = "1" Then s1 = lC1.List(i) Else
    If sNumeroCombo = "2" Then s1 = lC2.List(i)
    If s1 = sItem Then e = True
    i = i + 1
  Loop
  ExisteEnCombo = e
End Function

Private Sub bAceptar_Click()
  Dim s As String, s1 As String
  Dim i As Integer
  
  s = "delete from combosprogramados where combonumero = '1'"
  Modulo.ExecSQL s
  
  s = "delete from combosprogramados where combonumero = '2'"
  Modulo.ExecSQL s
  
  For i = 0 To lC1.ListCount - 1
    s1 = Trim(Mid(lC1.List(i), 1, 20))
    s = "insert into combosprogramados (combonumero,codigoproducto) values ('1','" & s1 & "')"
    Modulo.ExecSQL s
  Next i
  
  For i = 0 To lC2.ListCount - 1
    s1 = Trim(Mid(lC2.List(i), 1, 20))
    s = "insert into combosprogramados (combonumero,codigoproducto) values ('2','" & s1 & "')"
    Modulo.ExecSQL s
  Next i
    
  MsgBox "Combos Guardados Correctamente...", vbInformation, "Información"
  Unload Me

End Sub

Private Sub bAnexar1_Click()
  Dim i As Integer
  i = lProductos.ListIndex
  If i >= 0 Then
    If Not ExisteEnCombo("1", lProductos.List(i)) Then
      lC1.AddItem lProductos.List(i)
    Else
      MsgBox "Ya Producto Existe en Combo 1", vbCritical, "Información"
    End If
  End If
End Sub

Private Sub bAnexar2_Click()
  Dim i As Integer
  i = lProductos.ListIndex
  If i >= 0 Then
    If Not ExisteEnCombo("2", lProductos.List(i)) Then
      lC2.AddItem lProductos.List(i)
    Else
      MsgBox "Ya Producto Existe en Combo 2", vbCritical, "Información"
    End If
  End If
End Sub

Private Sub bCancelar_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
  If lC1.ListIndex >= 0 Then lC1.RemoveItem lC1.ListIndex
End Sub

Private Sub Command2_Click()
  If lC2.ListIndex >= 0 Then lC2.RemoveItem lC2.ListIndex
End Sub

Private Sub Form_Load()
  CargarProductos
  Combo1.Clear
  Combo1.AddItem "Combo-1"
  Combo1.ListIndex = 0
  
  Combo2.Clear
  Combo2.AddItem "Combo-2"
  Combo2.ListIndex = 0
  
  CargarCombo "1"
  CargarCombo "2"
End Sub
