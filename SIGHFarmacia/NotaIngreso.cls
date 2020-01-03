VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotaIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para Mantenimiento de Nota de Ingreso
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_NotaIngreso As New FarmNotaIngreso
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_movNumero As String
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Sub MostrarFormularioNotaIngreso()
    mo_NotaIngreso.Opcion = mi_Opcion
    mo_NotaIngreso.idUsuario = ml_idUsuario
    mo_NotaIngreso.movNumero = ml_movNumero
    mo_NotaIngreso.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    mo_NotaIngreso.lcNombrePc = mo_lcNombrePc
    mo_NotaIngreso.Show 1
    Set mo_NotaIngreso = Nothing
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
Property Let movNumero(lValue As String)
   ml_movNumero = lValue
End Property




