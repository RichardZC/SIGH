VERSION 5.00
Begin VB.UserControl ucGestionCajaFact 
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10725
   ScaleHeight     =   3165
   ScaleWidth      =   10725
   Begin VB.Frame Frame 
      Height          =   2565
      Index           =   1
      Left            =   15
      TabIndex        =   3
      Top             =   570
      Width           =   9285
      Begin VB.CheckBox chkCredito 
         Alignment       =   1  'Right Justify
         Caption         =   "Pago a crédito"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5850
         TabIndex        =   9
         Top             =   2205
         Width           =   1515
      End
      Begin VB.TextBox txtIGV 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7395
         TabIndex        =   7
         Top             =   510
         Width           =   975
      End
      Begin VB.CheckBox chkIGV 
         Caption         =   "Genera IGV"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         TabIndex        =   6
         Top             =   2235
         Width           =   1515
      End
      Begin VB.TextBox txtDescripcion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   90
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   165
         Width           =   7260
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7395
         TabIndex        =   4
         Top             =   165
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "IGV"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   8445
         TabIndex        =   8
         Top             =   585
         Width           =   315
      End
   End
   Begin VB.Frame Frame 
      Height          =   495
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   45
      Width           =   9270
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7440
         TabIndex        =   2
         Top             =   180
         Width           =   660
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   165
         Width           =   915
      End
   End
End
Attribute VB_Name = "ucGestionCajaFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event SeIngresoImporte(lnImporte As Double, lnMontoIGV As Double, lbEsCredito As Boolean)
Public Event SeIngresoDescripcion(lcTexto As String)
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_lnIGV As Double

Property Let lnIGV(lValue As Double)
   ml_lnIGV = lValue
End Property



Sub CalculaMontoIGV()
    If chkIGV.Value = 1 Then
       txtIGV.Text = Round(CCur(txtImporte.Text) * ml_lnIGV / 100, 2)
    Else
       txtIGV.Text = 0
    End If
    RaiseEvent SeIngresoImporte(CCur(txtImporte.Text), CCur(txtIGV.Text), IIf(chkCredito.Value = 1, True, False))
End Sub

Private Sub chkCredito_Click()
    CalculaMontoIGV
End Sub

Private Sub chkIGV_Click()
    CalculaMontoIGV
End Sub

Private Sub txtDescripcion_LostFocus()
    If Len(txtDescripcion.Text) > 0 Then
       RaiseEvent SeIngresoDescripcion(txtDescripcion.Text)
    End If
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       'Else
       '    CalculaMontoIGV
       End If
End Sub

Private Sub txtImporte_LostFocus()
    CalculaMontoIGV
    'RaiseEvent SeIngresoImporte(CCur(txtImporte.Text), CCur(txtIGV.Text))
End Sub

Sub LimpiarDatos()
    txtImporte.Text = 0
    txtDescripcion.Text = ""
    txtIGV.Text = 0
    chkCredito.Value = 0
    chkIGV.Value = 0
End Sub
