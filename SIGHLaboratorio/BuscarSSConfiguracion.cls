VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuscarSSConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Buscar configuración de resultados
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ServicioBusqueda As New ServiciosSinConfiguracion
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdRegistroSeleccionado As Long

Sub MostrarFormulario()
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
