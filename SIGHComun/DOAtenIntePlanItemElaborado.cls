VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtenIntePlanItemElaborado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtenIntePlanItemElaborado
'        Programado por: Garay M
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdAtenInteItemPlan As Long
Dim ml_IdPlanIntegralPaciente As Long
Dim mb_EsElaborado As Boolean
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdAtenInteItemPlan(iValue As Long)
   ml_IdAtenInteItemPlan = iValue
End Property
Property Get IdAtenInteItemPlan() As Long
   IdAtenInteItemPlan = ml_IdAtenInteItemPlan
End Property
Property Let IdPlanIntegralPaciente(lValue As Long)
   ml_IdPlanIntegralPaciente = lValue
End Property
Property Get IdPlanIntegralPaciente() As Long
   IdPlanIntegralPaciente = ml_IdPlanIntegralPaciente
End Property
Property Let EsElaborado(bValue As Boolean)
   mb_EsElaborado = bValue
End Property
Property Get EsElaborado() As Boolean
   EsElaborado = mb_EsElaborado
End Property
