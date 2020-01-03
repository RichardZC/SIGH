VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DependenciaExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para el Mantenimiento de Dependencias Externas
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_DependenciaExt As New FarmDependenciaExtDetalle
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_IdDependenciaExt As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Sub MostrarFormulario()
    mo_DependenciaExt.Opcion = mi_Opcion
    mo_DependenciaExt.idUsuario = ml_idUsuario
    mo_DependenciaExt.IdDependenciaExt = ml_IdDependenciaExt
    mo_DependenciaExt.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    mo_DependenciaExt.lcNombrePc = mo_lcNombrePc
    mo_DependenciaExt.Show 1
    Set mo_DependenciaExt = Nothing
End Sub


Property Get Opcion() As sghOpciones
    Opcion = mi_Opcion
End Property


Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property
Property Let IdDependenciaExt(lValue As Long)
   ml_IdDependenciaExt = lValue
End Property
