VERSION 5.00
Begin VB.Form ServiciosBusqueda 
   Caption         =   "Busqueda De Servicios"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   Icon            =   "ServiciosBusqueda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   75
      TabIndex        =   1
      Top             =   5370
      Width           =   9990
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ServiciosBusqueda.frx":08CA
         DownPicture     =   "ServiciosBusqueda.frx":0D8E
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
         Picture         =   "ServiciosBusqueda.frx":127A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ServiciosBusqueda.frx":1766
         DownPicture     =   "ServiciosBusqueda.frx":1BC6
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
         Picture         =   "ServiciosBusqueda.frx":203B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
   End
   Begin Galenhos.ucServiciosListaBus ucServiciosListaBus1 
      Height          =   5385
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   9499
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
Dim ml_IdDepartamentoHospital As Long
Dim mb_EjecutarBusquedaOnLoad As Boolean

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
Property Let IdDepartamentoHospital(lValue As Long)
    ml_IdDepartamentoHospital = lValue
End Property
Property Get IdDepartamentoHospital() As Long
    IdDepartamentoHospital = ml_IdDepartamentoHospital
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
    Me.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
    
    ucServiciosListaBus1.Inicializar
    Me.ucServiciosListaBus1.Titulo = "Búsqueda de Servicios"
    Me.ucServiciosListaBus1.ConfigurarTiposServicio
    ucServiciosListaBus1.IdTipoServicio = ml_IdServicio
    ucServiciosListaBus1.IdDepartamentoHospital = ml_IdDepartamentoHospital
    If mb_EjecutarBusquedaOnLoad Then
        ucServiciosListaBus1.RealizarBusqueda
    End If
    
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucServiciosListaBus1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub



