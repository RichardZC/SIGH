VERSION 5.00
Begin VB.Form EstablecimientosBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
   Icon            =   "EstablecimientosBusqueda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin GalenHos.ucEstablecimientosLista ucEstablecimientosLista1 
      Height          =   5115
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   9022
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   90
      TabIndex        =   3
      Top             =   5160
      Width           =   10425
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "EstablecimientosBusqueda.frx":0CCA
         DownPicture     =   "EstablecimientosBusqueda.frx":118E
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
         Left            =   5475
         Picture         =   "EstablecimientosBusqueda.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "EstablecimientosBusqueda.frx":1B66
         DownPicture     =   "EstablecimientosBusqueda.frx":1FC6
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
         Left            =   3930
         Picture         =   "EstablecimientosBusqueda.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "EstablecimientosBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado

Property Set DataSource(oValue As ADODB.Recordset)
    Set ucEstablecimientosLista1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucEstablecimientosLista1.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ucEstablecimientosLista1.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ucEstablecimientosLista1.IdRegistroSeleccionado
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


Private Sub cmdAgregar_Click()
    
End Sub

Private Sub Form_Initialize()
    ucEstablecimientosLista1.ConfigurarEstablecimientos
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
    Me.ucEstablecimientosLista1.Inicializar
    Me.ucEstablecimientosLista1.Titulo = "Búsqueda de Establecimientos"
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucEstablecimientosLista1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub



Private Sub ucEstablecimientosLista1_SeleccionaRegistro(lnIdRegistroSeleccionado As Long)
    If lnIdRegistroSeleccionado > 0 Then
       btnAceptar_Click
    End If
End Sub
