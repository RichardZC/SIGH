VERSION 5.00
Begin VB.Form ServiciosBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda de Servicios"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   Icon            =   "ServiciosBusqueda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1035
      Left            =   60
      TabIndex        =   3
      Top             =   5280
      Width           =   9945
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         Height          =   700
         Left            =   3450
         Picture         =   "ServiciosBusqueda.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         Height          =   700
         Left            =   5460
         Picture         =   "ServiciosBusqueda.frx":113F
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
   End
   Begin SIGHNegocios.ucServiciosListaBus ucServiciosListaBus1 
      Height          =   5145
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   9075
   End
End
Attribute VB_Name = "ServiciosBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca Servicio del Establecimiento
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdServicio As Long
Dim mc_NombreServicio As String


Property Let SoloIdTipoServicio(lValue As Long)
    Me.ucServiciosListaBus1.SoloIdTipoServicio = lValue
End Property

Property Let NombreServicio(lValue As String)
    ucServiciosListaBus1.NombreServicio = lValue
    mc_NombreServicio = lValue
End Property

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
Property Let idTipoServicio(lValue As Long)
    ml_IdServicio = lValue
End Property
Property Get idTipoServicio() As Long
    idTipoServicio = ml_IdServicio
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

Private Sub Form_Activate()
'    If mc_NombreServicio <> "" Then
        ucServiciosListaBus1.ColocarFocoEnGrillaServicio
'    End If
'    Me.ucServiciosListaBus1.FocusEnDescripcion
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
    
    Me.ucServiciosListaBus1.Titulo = "Búsqueda de Servicios"
    Me.ucServiciosListaBus1.ConfigurarTiposServicio
    ucServiciosListaBus1.idTipoServicio = ml_IdServicio
    
End Sub



Private Sub ucServiciosListaBus1_SeleccionaRegistro(lnIdRegistroSeleccionado As Long)
    If lnIdRegistroSeleccionado > 0 Then
       btnAceptar_Click
    End If

End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub
