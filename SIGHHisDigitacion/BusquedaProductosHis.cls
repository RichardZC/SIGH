VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BusquedaProductosHis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para buscar productos HIS
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ProductosHisBusqueda As New frmDetalleProductosHis
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdRegistroSeleccionado As Long
Dim ml_CodigoDx As String
Dim ml_descripciondiagcpt As String
Dim ml_MasDeUnDiagnosticos As Integer
Dim lbSoloMuestraDxGalenHos As Boolean
Dim ml_USAcodigoCIEsinPto As Boolean

Property Let CodigoDx(lValue As String)
    ml_CodigoDx = lValue
End Property
Property Get CodigoDx() As String
    CodigoDx = ml_CodigoDx
End Property

Sub MostrarFormulario()
    mo_ProductosHisBusqueda.CodigoDx = ml_CodigoDx
    mo_ProductosHisBusqueda.Show 1
    ml_IdRegistroSeleccionado = mo_ProductosHisBusqueda.IdDiagCpt
    ml_CodigoDx = mo_ProductosHisBusqueda.CodigoDx
    ml_descripciondiagcpt = mo_ProductosHisBusqueda.descripciondiagcpt
    ml_MasDeUnDiagnosticos = mo_ProductosHisBusqueda.MasDeUnDiagnosticos
    mi_BotonPresionado = mo_ProductosHisBusqueda.BotonPresionado
    Set mo_ProductosHisBusqueda = Nothing
End Sub

Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
End Property
Property Let descripciondiagcpt(lValue As String)
    ml_descripciondiagcpt = lValue
End Property
Property Get descripciondiagcpt() As String
    descripciondiagcpt = ml_descripciondiagcpt
End Property
Property Let MasDeUnDiagnosticos(sValue As Integer)
   ml_MasDeUnDiagnosticos = sValue
End Property
Property Get MasDeUnDiagnosticos() As Integer
   MasDeUnDiagnosticos = ml_MasDeUnDiagnosticos
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

