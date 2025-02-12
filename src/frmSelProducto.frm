VERSION 5.00
Begin VB.Form frmSelProducto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalles para Acuerdo de Producto"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   2580
      TabIndex        =   17
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   840
      TabIndex        =   16
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3435
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4395
      Begin VB.TextBox txtProducto 
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   300
         Width           =   1455
      End
      Begin VB.ComboBox cmbPagado 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSelProducto.frx":0000
         Left            =   1500
         List            =   "frmSelProducto.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2880
         Width           =   1455
      End
      Begin VB.ComboBox cmbEntregado 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSelProducto.frx":0016
         Left            =   1500
         List            =   "frmSelProducto.frx":0020
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtSubtotal 
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtPrecioAcordado 
         Height          =   315
         Left            =   1500
         TabIndex        =   12
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtPrecioNormal 
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox cmbCantidad 
         Height          =   315
         Left            =   1500
         TabIndex        =   10
         Text            =   "cmbCantidad"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   2595
      End
      Begin VB.Label Label9 
         Caption         =   "Código:"
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Precio Acordado:"
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   1860
         Width           =   1395
      End
      Begin VB.Label Label8 
         Caption         =   "Pagado:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Entregado:"
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Subtotal:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad:"
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Precio Normal:"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción:"
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Producto:"
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Top             =   0
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmSelProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lID As Integer


Private Sub cmbCantidad_Change()
  If IsNumeric(cmbCantidad.Text) = False Then Exit Sub
   
   If Val(txtPrecioAcordado.Text) > 0 Then
      txtSubtotal.Text = Val(Replace(txtPrecioAcordado.Text, ",", ".")) * (cmbCantidad.Text)
   Else
      txtSubtotal.Text = "0"
   End If

End Sub

Private Sub cmbCantidad_Click()
   If Val(txtPrecioAcordado.Text) > 0 Then
      txtSubtotal.Text = Val(Replace(txtPrecioAcordado.Text, ",", ".")) * (cmbCantidad.Text)
   Else
      txtSubtotal.Text = "0"
   End If
   
End Sub

Private Sub cmbCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then ' tecla escape
   Unload Me
End If

End Sub

Private Sub cmbEntregado_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then ' tecla escape
   Unload Me
End If

End Sub

Private Sub cmbPagado_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then ' tecla escape
   Unload Me
End If

End Sub

Private Sub cmdAceptar_Click()
  'VALIDAD CMBCANTIDAD
  If IsNumeric(cmbCantidad.Text) = False Then
     MsgBox "Debe escribir solo numeros en el campo Cantidad", vbCritical
     Exit Sub
  End If
 With fFD.ListvProductos
   .SelectedItem.Text = txtProducto.Text
    .SelectedItem.SubItems(1) = txtDescripcion.Text
    .SelectedItem.SubItems(2) = cmbCantidad.Text
    .SelectedItem.SubItems(3) = txtPrecioNormal.Text
   .SelectedItem.SubItems(4) = txtPrecioAcordado.Text
   .SelectedItem.SubItems(5) = txtSubtotal.Text
   .SelectedItem.SubItems(6) = cmbEntregado.Text
   .SelectedItem.SubItems(7) = cmbPagado.Text
   lID = .SelectedItem.SubItems(8)
   .SelectedItem.Checked = True
   fFD.HayCambios = True
 End With
 Unload Me

End Sub

Private Sub cmdCerrar_Click()
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   'If KeyAscii =  Then
   'End If
   
End Sub

Private Sub Form_Load()
   sCargarCombo
   sCargarDatos
End Sub
Private Sub sCargarDatos()
 With fFD.ListvProductos
   txtProducto.Text = .SelectedItem.Text
   txtDescripcion.Text = UCase(.SelectedItem.SubItems(1))
   cmbCantidad.Text = .SelectedItem.SubItems(2)
   txtPrecioNormal.Text = .SelectedItem.SubItems(3)
   txtPrecioAcordado.Text = IIf(.SelectedItem.SubItems(4) = 0, .SelectedItem.SubItems(3), .SelectedItem.SubItems(4))
   txtSubtotal.Text = .SelectedItem.SubItems(5)
   cmbEntregado.Text = .SelectedItem.SubItems(6)
   cmbPagado.Text = .SelectedItem.SubItems(7)
   'lId = .SelectedItem.SubItems(7)
 End With
End Sub


Private Sub sCargarCombo()
   Dim i As Integer
   For i = 0 To 150
      cmbCantidad.AddItem i
   Next i
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then ' tecla escape
   Unload Me
End If

End Sub

Private Sub txtPrecioAcordado_Change()
   cmbCantidad_Click
End Sub

Private Sub txtPrecioAcordado_Click()
   txtPrecioAcordado.SelStart = 0
   txtPrecioAcordado.SelLength = Len(txtPrecioAcordado.Text)
   txtPrecioAcordado.SetFocus
End Sub

Private Sub txtPrecioAcordado_GotFocus()
   txtPrecioAcordado.SelStart = 0
   txtPrecioAcordado.SelLength = Len(txtPrecioAcordado.Text)
   txtPrecioAcordado.SetFocus
End Sub

Private Sub txtPrecioAcordado_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then ' tecla escape
   Unload Me
End If
End Sub

Private Sub txtPrecioNormal_Change()
  cmbCantidad_Click
End Sub

Private Sub txtPrecioNormal_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then ' tecla escape
   Unload Me
End If

End Sub

Private Sub txtProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then ' tecla escape
   Unload Me
End If

End Sub

Private Sub txtSubtotal_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then ' tecla escape
   Unload Me
End If

End Sub
