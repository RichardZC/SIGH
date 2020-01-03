VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuscaDiagnosticos"
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
Dim mo_DiagnosticosBusqueda As New DiagnosticosBusqueda
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdRegistroSeleccionado As Long
Dim ml_CodigoDx As String
Dim lbSoloMuestraDxGalenHos As Boolean
Dim ml_USAcodigoCIEsinPto As Boolean
'mgaray20141022
Dim ml_MostrarSoloActivos As Boolean

Property Let USAcodigoCIEsinPto(lValue As Boolean)
    ml_USAcodigoCIEsinPto = lValue
End Property

Property Let SoloMuestraDxGalenHos(lValue As Boolean)
    lbSoloMuestraDxGalenHos = lValue
End Property

Property Let CodigoDx(lValue As String)
    ml_CodigoDx = lValue
End Property
'mgaray20141022
Property Let MostrarSoloActivos(lValue As Boolean)
    ml_MostrarSoloActivos = lValue
End Property

Private Sub Class_Initialize()
    ml_MostrarSoloActivos = True
End Sub

Sub MostrarFormulario()
    mo_DiagnosticosBusqueda.USAcodigoCIEsinPto = ml_USAcodigoCIEsinPto
    mo_DiagnosticosBusqueda.SoloMuestraDxGalenHos = lbSoloMuestraDxGalenHos
    'mgaray20141022
    mo_DiagnosticosBusqueda.ucDiagnosticosLista1.MostrarSoloActivos = ml_MostrarSoloActivos
'    mo_DiagnosticosBusqueda.ucDiagnosticosLista1.MostrarSoloActivos = True
    mo_DiagnosticosBusqueda.CodigoDx = ml_CodigoDx
    mo_DiagnosticosBusqueda.Show 1
    ml_IdRegistroSeleccionado = mo_DiagnosticosBusqueda.IdRegistroSeleccionado
    mi_BotonPresionado = mo_DiagnosticosBusqueda.BotonPresionado
    Set mo_DiagnosticosBusqueda = Nothing
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

