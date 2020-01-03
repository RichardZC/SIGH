VERSION 5.00
Begin VB.Form ServiciosBusqueda 
   Caption         =   "Busqueda De Servicios"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   Icon            =   "ServiciosBusqueda.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   10155
   StartUpPosition =   2  'CenterScreen
   Begin SIGHProxies.ucServiciosListaBus ucServiciosListaBus1 
      Height          =   5445
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   9604
   End
   Begin VB.Frame Frame2 
      Height          =   1035
      Left            =   90
      TabIndex        =   0
      Top             =   5430
      Width           =   9945
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         Height          =   700
         Left            =   5460
         Picture         =   "ServiciosBusqueda.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         Height          =   700
         Left            =   3450
         Picture         =   "ServiciosBusqueda.frx":0DB6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "ServiciosBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdServicio As Long

Property Set DataSource(oValue As ADODB.Recordset)
    Set ucServiciosListaBus1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucServiciosListaBus1.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ucServiciosListaBus1.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ucServiciosListaBus1.IdRegistroSeleccionado
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property
Property Let IdTipoServicio(lValue As Long)
    ml_IdServicio = lValue
End Property
Property Get IdTipoServicio() As Long
    IdTipoServicio = ml_IdServicio
End Property
Property Let HabilitarTipoServicio(lValue As Boolean)
    ucServiciosListaBus1.HabilitarTipoServicio = lValue
End Property
Property Get HabilitarTipoServicio() As Boolean
   HabilitarTipoServicio = ucServiciosListaBus1.HabilitarTipoServicio
End Property
Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    Me.Visible = False
End Sub
Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    Me.Visible = False
End Sub
Private Sub Form_Load()
    
    Me.ucServiciosListaBus1.Titulo = "Búsqueda de Servicios"
    Me.ucServiciosListaBus1.ConfigurarTiposServicio
    ucServiciosListaBus1.IdTipoServicio = ml_IdServicio
    
End Sub


