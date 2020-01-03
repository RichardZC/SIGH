VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuscaServicioHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para buscar Servicio del Establecimiento
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ServicioBusqueda As New ServiciosBusqueda
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdRegistroSeleccionado As Long
Dim ml_HabilitarTipoServicio As Boolean
Dim ml_IdTipoServicio As Long
Dim ml_NombreServicio As String
Dim ml_SoloIdTipoServicio As Long

Property Let SoloIdTipoServicio(lValue As Long)
    ml_SoloIdTipoServicio = lValue
End Property

Property Let NombreServicio(lValue As String)
    ml_NombreServicio = lValue
End Property

Sub MostrarFormulario()
    mo_ServicioBusqueda.SoloIdTipoServicio = ml_IdTipoServicio
    mo_ServicioBusqueda.HabilitarTipoServicio = ml_HabilitarTipoServicio
    mo_ServicioBusqueda.NombreServicio = ml_NombreServicio
    mo_ServicioBusqueda.Show 1
    ml_IdRegistroSeleccionado = mo_ServicioBusqueda.IdRegistroSeleccionado
    mi_BotonPresionado = mo_ServicioBusqueda.BotonPresionado
    Set mo_ServicioBusqueda = Nothing
End Sub

Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Property Let idTipoServicio(lValue As Long)
    ml_IdTipoServicio = lValue
End Property
Property Get idTipoServicio() As Long
    idTipoServicio = ml_IdTipoServicio
End Property
Property Let HabilitarTipoServicio(lValue As Boolean)
    ml_HabilitarTipoServicio = lValue
End Property
Property Get HabilitarTipoServicio() As Boolean
   HabilitarTipoServicio = ml_HabilitarTipoServicio
End Property



