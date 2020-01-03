VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clFuenteFinancDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para el mantenimiento de Fuente de Financiamiento
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idUsuario As Long
Dim mi_Opcion As sghOpciones
Dim ml_IdFuenteFinanciamiento As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim oFormulario As New FuenteFinanciamientoDetalle

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let IdFuenteFinanciamiento(lValue As Long)
   ml_IdFuenteFinanciamiento = lValue
End Property
Property Get IdFuenteFinanciamiento() As Long
   IdFuenteFinanciamiento = ml_IdFuenteFinanciamiento
End Property



Sub MostrarFormulario()
    oFormulario.Opcion = mi_Opcion
    oFormulario.idUsuario = ml_idUsuario
    oFormulario.IdFuenteFinanciamiento = ml_IdFuenteFinanciamiento
    oFormulario.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    oFormulario.lcNombrePc = mo_lcNombrePc
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub




