VERSION 5.00
Begin VB.Form CatalogoServiciosBusqueda 
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin Galenhos.ucCatalogoServiciosListaBus ucCatalogoServiciosListaBus1 
      Height          =   7095
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   12515
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   0
      TabIndex        =   0
      Top             =   7080
      Width           =   10215
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CatalogoServiciosBusqueda.frx":0000
         DownPicture     =   "CatalogoServiciosBusqueda.frx":0460
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
         Left            =   3495
         Picture         =   "CatalogoServiciosBusqueda.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CatalogoServiciosBusqueda.frx":0D4A
         DownPicture     =   "CatalogoServiciosBusqueda.frx":120E
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
         Left            =   5040
         Picture         =   "CatalogoServiciosBusqueda.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CatalogoServiciosBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdTipoCatalogo As Long
'Dim ml_IdDepartamentoHospital As Long
Dim mb_EjecutarBusquedaOnLoad As Boolean

Property Set DataSource(oValue As ADODB.Recordset)
    Set ucCatalogoServiciosListaBus1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucCatalogoServiciosListaBus1.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ucCatalogoServiciosListaBus1.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ucCatalogoServiciosListaBus1.IdRegistroSeleccionado
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property
Property Let IdTipoCatalogo(lValue As Long)
    ml_IdTipoCatalogo = lValue
End Property
Property Get IdTipoCatalogo() As Long
    IdTipoCatalogo = ml_IdTipoCatalogo
End Property
Property Let HabilitarTipoCatalogo(lValue As Boolean)
    ucCatalogoServiciosListaBus1.HabilitarTipoCatalogo = lValue
End Property
Property Get HabilitarTipoCatalogo() As Boolean
    HabilitarTipoCatalogo = ucCatalogoServiciosListaBus1.HabilitarTipoCatalogo
End Property
'Property Let IdDepartamentoHospital(lValue As Long)
'    ml_IdDepartamentoHospital = lValue
'End Property
'Property Get IdDepartamentoHospital() As Long
'    IdDepartamentoHospital = ml_IdDepartamentoHospital
'End Property
Property Let EjecutarBusquedaOnLoad(bValue As Boolean)
    mb_EjecutarBusquedaOnLoad = bValue
End Property
Property Get EjecutarBusquedaOnLoad() As Boolean
    EjecutarBusquedaOnLoad = mb_EjecutarBusquedaOnLoad
End Property

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    Me.Visible = False
End Sub
Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    IdRegistroSeleccionado = 0
    Me.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
    
    
    ucCatalogoServiciosListaBus1.Inicializar
    Me.ucCatalogoServiciosListaBus1.Titulo = "Búsqueda de Catalogo de Servicios"
    Me.ucCatalogoServiciosListaBus1.ConfigurarTiposCatalogos
    ucCatalogoServiciosListaBus1.IdTipoCatalogo = ml_IdTipoCatalogo
    ucCatalogoServiciosListaBus1.SeleccionaTipoCatalogo
    'ucCatalogoServiciosListaBus1.iºº.IdDepartamentoHospital = ml_IdDepartamentoHospital
    If mb_EjecutarBusquedaOnLoad Then
        ucCatalogoServiciosListaBus1.RealizarBusqueda
    End If
    
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucCatalogoServiciosListaBus1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub





