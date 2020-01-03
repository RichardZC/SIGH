VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuscarDistrito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para buscar distrito
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Distrito As New frmBusquedaDistrito
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdRegistroSeleccionado As Long
Dim ms_DescripcionDistritoSeleccionado As String

Dim ms_IdDepartamentoBusqueda As Long
Dim ms_IdProvinciaBusqueda As Long
Dim ms_DescripcionDistrito As String

Property Let IdDepartamentoBusqueda(lValue As Long)
    ms_IdDepartamentoBusqueda = lValue
End Property

Property Let IdProvinciaBusqueda(lValue As Long)
    ms_IdProvinciaBusqueda = lValue
End Property

Property Let DescripcionDistrito(sValue As String)
    ms_DescripcionDistrito = sValue
End Property

Sub MostrarFormulario()
    mo_Distrito.IdDepartamentoPorDefecto = ms_IdDepartamentoBusqueda
    mo_Distrito.IdProvinciaPorDefecto = ms_IdProvinciaBusqueda
    mo_Distrito.DescripcionDistrito = ms_DescripcionDistrito
    mo_Distrito.Show 1
    ml_IdRegistroSeleccionado = mo_Distrito.IdDistritoSeleccionado
    ms_DescripcionDistritoSeleccionado = mo_Distrito.NombreDistritoSeleccionado
    mi_BotonPresionado = mo_Distrito.BotonPresionado
    Set mo_Distrito = Nothing
End Sub

Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
End Property

Property Let DescripcionRegistroSeleccionado(sValue As String)
    ms_DescripcionDistritoSeleccionado = sValue
End Property
Property Get DescripcionRegistroSeleccionado() As String
    DescripcionRegistroSeleccionado = ms_DescripcionDistritoSeleccionado
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

