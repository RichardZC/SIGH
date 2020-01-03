VERSION 5.00
Begin VB.Form CatalogoServicioBusca 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12090
   Icon            =   "CatalogoServicioBusca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SIGHNegocios.ucCatalServicioLista ucCatalServicioLista1 
      Height          =   6975
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   12303
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   0
      TabIndex        =   3
      Top             =   7080
      Width           =   12045
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CatalogoServicioBusca.frx":000C
         DownPicture     =   "CatalogoServicioBusca.frx":046C
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
         Left            =   4590
         Picture         =   "CatalogoServicioBusca.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CatalogoServicioBusca.frx":0D56
         DownPicture     =   "CatalogoServicioBusca.frx":121A
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
         Left            =   6135
         Picture         =   "CatalogoServicioBusca.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CatalogoServicioBusca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca procedimiento CPT
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdTipoCatalogo As Long
'Dim ml_IdDepartamentoHospital As Long
Dim mb_EjecutarBusquedaOnLoad As Boolean

Property Set DataSource(oValue As ADODB.Recordset)
    Set ucCatalServicioLista1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucCatalServicioLista1.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ucCatalServicioLista1.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ucCatalServicioLista1.IdRegistroSeleccionado
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
    ucCatalServicioLista1.HabilitarTipoCatalogo = lValue
End Property
Property Get HabilitarTipoCatalogo() As Boolean
    HabilitarTipoCatalogo = ucCatalServicioLista1.HabilitarTipoCatalogo
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
    
    
    ucCatalServicioLista1.inicializar
    Me.ucCatalServicioLista1.Titulo = "Búsqueda de Catalogo de Servicios"
    Me.ucCatalServicioLista1.ConfigurarTiposCatalogos
    ucCatalServicioLista1.IdTipoCatalogo = ml_IdTipoCatalogo
    ucCatalServicioLista1.SeleccionaTipoCatalogo
    'ucCatalServicioLista1.iºº.IdDepartamentoHospital = ml_IdDepartamentoHospital
    If mb_EjecutarBusquedaOnLoad Then
        ucCatalServicioLista1.RealizarBusqueda
    End If
    
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucCatalServicioLista1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub





Private Sub ucCatalServicioLista1_SeleccionaRegistro(lnIdRegistroSeleccionado As Long)
    If lnIdRegistroSeleccionado > 0 Then
       btnAceptar_Click
    End If

End Sub
