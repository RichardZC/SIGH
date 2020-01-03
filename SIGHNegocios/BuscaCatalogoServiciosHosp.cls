VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuscaCatalogoServiciosHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para buscar Diagnóstico
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ServiciosBusqueda As New BuscarCatalogoServiciosHosp
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdRegistroSeleccionado As Long

Dim ml_IdPuntoCarga As Long
Dim ml_IdTipoFinanciamiento As Long
Dim ml_TipoServicioOfrecido As Long

Property Let IdPuntoCarga(lValue As Long)
    ml_IdPuntoCarga = lValue
End Property

Property Let IdTipoFinanciamiento(lValue As Long)
    ml_IdTipoFinanciamiento = lValue
End Property

Property Let TipoServicioOfrecido(lValue As Long)
    ml_TipoServicioOfrecido = lValue
End Property

Private Sub Class_Initialize()
    ml_TipoServicioOfrecido = 2 'insumos y cpt
End Sub

Sub MostrarFormulario()
    mo_ServiciosBusqueda.IdPuntoCarga = ml_IdPuntoCarga
    mo_ServiciosBusqueda.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
    mo_ServiciosBusqueda.TipoServicioOfrecido = ml_TipoServicioOfrecido
    mo_ServiciosBusqueda.Show 1
    ml_IdRegistroSeleccionado = mo_ServiciosBusqueda.IdRegistroSeleccionado
    mi_BotonPresionado = mo_ServiciosBusqueda.BotonPresionado
    Set mo_ServiciosBusqueda = Nothing
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


