VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuscaEstablecimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para buscar Establecimiento MINSA
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim oFormulario As New EstablecimientosBusqueda
Dim ml_IdRegistroSeleccionado As Long
Dim mi_BotonPresionado As sghBotonDetallePresionado
'JGV
Dim ml_NivelEstablecimiento As Long
Dim ms_DescripcionEstablecimiento As String

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
End Property
'JVG
Property Let NivelMaximoEstablecimiento(lValue As Long)
    ml_NivelEstablecimiento = lValue
End Property
'JVG
Property Let DescripcionEstablecimiento(sValue As String)
    ms_DescripcionEstablecimiento = sValue
End Property

Sub MostrarFormulario()
    'JVG
    oFormulario.DescripcionEstablecimiento = ms_DescripcionEstablecimiento
    oFormulario.NivelMaximoEstablecimiento = ml_NivelEstablecimiento
    
    If ms_DescripcionEstablecimiento <> "" Then
        oFormulario.ucEstablecimientosLista1.LimpiarFiltro
        oFormulario.DescripcionEstablecimiento = ms_DescripcionEstablecimiento
        oFormulario.ucEstablecimientosLista1.RealizarBusqueda
    End If
    
    oFormulario.Show 1
    ml_IdRegistroSeleccionado = oFormulario.IdRegistroSeleccionado
    mi_BotonPresionado = oFormulario.BotonPresionado
    Set oFormulario = Nothing
End Sub


