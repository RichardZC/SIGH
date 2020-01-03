VERSION 5.00
Begin VB.Form BuscarCatalogoServiciosHosp 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7560
   ClientLeft      =   8940
   ClientTop       =   4170
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11070
   Begin SIGHNegocios.ucCatServicioHospLista ucCatalogoServicioHospLista1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      _ExtentX        =   17806
      _ExtentY        =   11245
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   0
      TabIndex        =   3
      Top             =   6480
      Width           =   10920
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "BuscarCodigoSunat.frx":0000
         DownPicture     =   "BuscarCodigoSunat.frx":0460
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
         Left            =   4080
         Picture         =   "BuscarCodigoSunat.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "BuscarCodigoSunat.frx":0D4A
         DownPicture     =   "BuscarCodigoSunat.frx":120E
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
         Left            =   5625
         Picture         =   "BuscarCodigoSunat.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "BuscarCatalogoServiciosHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca procedimiento CPT
'        Programado por: Garay M
'        Fecha: Octubre 2014
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdTipoCatalogo As Long
'Dim ml_IdDepartamentoHospital As Long
Dim mb_EjecutarBusquedaOnLoad As Boolean
Dim ml_IdPuntoCarga As Long
Dim ml_IdTipoFinanciamiento As Long
Dim ml_TipoServicioOfrecido As Long

Property Let IdPuntoCarga(lValue As Long)
    ml_IdPuntoCarga = lValue
End Property
Property Get IdPuntoCarga() As Long
    IdPuntoCarga = ml_IdPuntoCarga
End Property
Property Let IdTipoFinanciamiento(lValue As Long)
    ml_IdTipoFinanciamiento = lValue
End Property
Property Get IdTipoFinanciamiento() As Long
    IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property
Property Let TipoServicioOfrecido(lValue As Long)
    ml_TipoServicioOfrecido = lValue
End Property
Property Get TipoServicioOfrecido() As Long
    TipoServicioOfrecido = ml_TipoServicioOfrecido
End Property

Property Set DataSource(oValue As ADODB.Recordset)
    Set ucCatalogoServicioHospLista1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucCatalogoServicioHospLista1.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ucCatalogoServicioHospLista1.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ucCatalogoServicioHospLista1.IdRegistroSeleccionado
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
    ucCatalogoServicioHospLista1.HabilitarTipoCatalogo = lValue
End Property
Property Get HabilitarTipoCatalogo() As Boolean
    HabilitarTipoCatalogo = ucCatalogoServicioHospLista1.HabilitarTipoCatalogo
End Property

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
    
    ucCatalogoServicioHospLista1.inicializar
    Me.ucCatalogoServicioHospLista1.Titulo = "Búsqueda de Catalogo de Servicios"
    Me.ucCatalogoServicioHospLista1.ConfigurarTiposCatalogos
    ucCatalogoServicioHospLista1.IdTipoCatalogo = ml_IdTipoCatalogo
    ucCatalogoServicioHospLista1.IdPuntoCarga = ml_IdPuntoCarga
    ucCatalogoServicioHospLista1.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
    ucCatalogoServicioHospLista1.TipoServicioOfrecido = ml_TipoServicioOfrecido
    ucCatalogoServicioHospLista1.SeleccionaTipoCatalogo
    'ucCatalogoServicioHospLista1.iºº.IdDepartamentoHospital = ml_IdDepartamentoHospital
    If mb_EjecutarBusquedaOnLoad Then
        ucCatalogoServicioHospLista1.RealizarBusqueda
    End If
    
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucCatalogoServicioHospLista1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub ucCatalogoServicioHospLista1_SeleccionaRegistro(lnIdRegistroSeleccionado As Long)
    If lnIdRegistroSeleccionado > 0 Then
       btnAceptar_Click
    End If
End Sub

Private Sub ucCatalogoServicioHospLista1_SePresionoTeclaEspecial(KeyCode As Integer)
    AdministrarKeyPreview KeyCode
End Sub
