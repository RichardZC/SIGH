VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtenIntePlanProcedPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtenIntePlanProcedimientoPaciente
'        Programado por: Garay M
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long

Dim ml_IdPlanProcedimientoPaciente As Long
Dim ml_IdPlanIntegralPaciente As Long
Dim ml_IdProducto As Long
Dim ml_IdPlanAtencion As Long
Dim ml_IdAtenInteItemPlan As Long
Dim mda_FechaProgramada As Date
Dim mda_FechaEjecucion As Date
Dim mi_NumeroDosis As Integer
Dim ms_CodigoHIS As String
Dim ml_IdAtencion As Long
Dim ml_IdEstablecimiento As Long

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdPlanProcedimientoPaciente(iValue As Long)
   ml_IdPlanProcedimientoPaciente = iValue
End Property
Property Get IdPlanProcedimientoPaciente() As Long
   IdPlanProcedimientoPaciente = ml_IdPlanProcedimientoPaciente
End Property
Property Let IdPlanIntegralPaciente(lValue As Long)
   ml_IdPlanIntegralPaciente = lValue
End Property
Property Get IdPlanIntegralPaciente() As Long
   IdPlanIntegralPaciente = ml_IdPlanIntegralPaciente
End Property
Property Let IdProducto(iValue As Long)
   ml_IdProducto = iValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let IdPlanAtencion(iValue As Long)
   ml_IdPlanAtencion = iValue
End Property
Property Get IdPlanAtencion() As Long
   IdPlanAtencion = ml_IdPlanAtencion
End Property
Property Let IdAtenInteItemPlan(iValue As Long)
   ml_IdAtenInteItemPlan = iValue
End Property
Property Get IdAtenInteItemPlan() As Long
   IdAtenInteItemPlan = ml_IdAtenInteItemPlan
End Property
Property Let FechaProgramada(daValue As Date)
   mda_FechaProgramada = daValue
End Property
Property Get FechaProgramada() As Date
   FechaProgramada = mda_FechaProgramada
End Property
Property Let FechaEjecucion(daValue As Date)
   mda_FechaEjecucion = daValue
End Property
Property Get FechaEjecucion() As Date
   FechaEjecucion = mda_FechaEjecucion
End Property
Property Let NumeroDosis(iValue As Integer)
   mi_NumeroDosis = iValue
End Property
Property Get NumeroDosis() As Integer
   NumeroDosis = mi_NumeroDosis
End Property
Property Let CodigoHIS(sValue As String)
   ms_CodigoHIS = sValue
End Property
Property Get CodigoHIS() As String
   CodigoHIS = ms_CodigoHIS
End Property
Property Let IdAtencion(iValue As Long)
   ml_IdAtencion = iValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property

Property Let IdEstablecimiento(iValue As Long)
   ml_IdEstablecimiento = iValue
End Property
Property Get IdEstablecimiento() As Long
   IdEstablecimiento = ml_IdEstablecimiento
End Property


