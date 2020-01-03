VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuscaMadre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para buscar madre por DNI
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_BusquedaDeMadre As New BusquedaDeMadre
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdAtencionMadreSeleccionado As Long
Dim ml_IdNacimientoSeleccionado As Long

Sub MostrarFormulario()
    mo_BusquedaDeMadre.Show 1
    ml_IdAtencionMadreSeleccionado = mo_BusquedaDeMadre.IdAtencionMadreSeleccionado
    ml_IdNacimientoSeleccionado = mo_BusquedaDeMadre.IdNacimientoSeleccionado
    mi_BotonPresionado = mo_BusquedaDeMadre.BotonPresionado
    Set mo_BusquedaDeMadre = Nothing
End Sub

Property Let IdAtencionMadreSeleccionado(lValue As Long)
    ml_IdAtencionMadreSeleccionado = lValue
End Property
Property Get IdAtencionMadreSeleccionado() As Long
    IdAtencionMadreSeleccionado = ml_IdAtencionMadreSeleccionado
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Property Let IdNacimientoSeleccionado(lValue As Long)
    ml_IdNacimientoSeleccionado = lValue
End Property
Property Get IdNacimientoSeleccionado() As Long
    IdNacimientoSeleccionado = ml_IdNacimientoSeleccionado
End Property

