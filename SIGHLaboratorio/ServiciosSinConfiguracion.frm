VERSION 5.00
Begin VB.Form ServiciosSinConfiguracion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14235
   Icon            =   "ServiciosSinConfiguracion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   14235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SIGHLaboratorio.ucServiciosXConfigurar ucServiciosXConfigurar 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   15266
   End
End
Attribute VB_Name = "ServiciosSinConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: busca procedimientos sin configuración de resultados
'        Programado por: Madrid S
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdServicio As Long
Property Let NombreServicio(lValue As String)
    ucServiciosXConfigurar.NombreServicio = lValue
End Property

Property Set DataSource(oValue As ADODB.Recordset)
    Set ucServiciosXConfigurar.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucServiciosXConfigurar.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ucServiciosXConfigurar.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ucServiciosXConfigurar.IdRegistroSeleccionado
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
'Property Let HabilitarTipoServicio(lValue As Boolean)
'    ucServiciosXConfigurar.HabilitarTipoServicio = lValue
'End Property
'Property Get HabilitarTipoServicio() As Boolean
'   HabilitarTipoServicio = ucServiciosXConfigurar.HabilitarTipoServicio
'End Property

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    Me.Visible = False
End Sub
Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    Me.Visible = False
End Sub
Private Sub Form_Load()
    
    ucServiciosXConfigurar.Titulo = "Búsqueda de Servicios a Configurar"
    'Me.ucServiciosListaBus1.ConfigurarTiposServicio
'    ucServiciosListaBus1.idTipoServicio = ml_IdServicio
    
End Sub



Private Sub ucServiciosXConfigurar_SeleccionaRegistro(lnIdRegistroSeleccionado As Long)
    If lnIdRegistroSeleccionado > 0 Then
       btnAceptar_Click
    End If

End Sub

