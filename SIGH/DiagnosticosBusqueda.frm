VERSION 5.00
Begin VB.Form DiagnosticosBusqueda 
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   Icon            =   "DiagnosticosBusqueda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin Galenhos.ucDiagnosticosLista ucDiagnosticosLista1 
      Height          =   5535
      Left            =   -15
      TabIndex        =   1
      Top             =   -15
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   9763
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   45
      TabIndex        =   0
      Top             =   5505
      Width           =   9780
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "DiagnosticosBusqueda.frx":0CCA
         DownPicture     =   "DiagnosticosBusqueda.frx":118E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   4935
         Picture         =   "DiagnosticosBusqueda.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "DiagnosticosBusqueda.frx":1B66
         DownPicture     =   "DiagnosticosBusqueda.frx":1FC6
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   3390
         Picture         =   "DiagnosticosBusqueda.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1365
      End
   End
End
Attribute VB_Name = "DiagnosticosBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado

Property Set DataSource(oValue As ADODB.Recordset)
    Set ucDiagnosticosLista1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucDiagnosticosLista1.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ucDiagnosticosLista1.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ucDiagnosticosLista1.IdRegistroSeleccionado
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    Me.Visible = False
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    Me.Visible = False
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
    Me.ucDiagnosticosLista1.Titulo = "Búsqueda de diagnósticos"
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucDiagnosticosLista1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

