VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuscaCodigoSunat"
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
Dim mo_ServiciosBusqueda As New BuscarCodigosSunat
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdRegistroSeleccionado As Long
Dim ms_codigoSunat As String

Property Let codigoSunat(lValue As String)
    ms_codigoSunat = lValue
End Property
Property Get codigoSunat() As String
    codigoSunat = ms_codigoSunat
End Property


Sub MostrarFormulario()
    mo_ServiciosBusqueda.Show 1
    ms_codigoSunat = mo_ServiciosBusqueda.codigoSunat
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



