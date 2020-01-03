VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuscaEstablecNoMinsa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para buscar Establecimiento NO MINSA
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim oFormulario As New EstablecimientosNoMinsaBusq
Dim ml_IdRegistroSeleccionado As Long
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim mo_lcNombrePc As String
Dim ml_idUsuario As Long

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Sub MostrarFormulario()
    oFormulario.lcNombrePc = mo_lcNombrePc
    oFormulario.idUsuario = ml_idUsuario
    oFormulario.Show 1
    ml_IdRegistroSeleccionado = oFormulario.IdRegistroSeleccionado
    mi_BotonPresionado = oFormulario.BotonPresionado
    Set oFormulario = Nothing
End Sub



