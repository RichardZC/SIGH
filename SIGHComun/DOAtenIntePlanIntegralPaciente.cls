VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtenIntePlanIntePaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtenIntePlanIntegralPaciente
'        Programado por: Garay M
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdPlanIntegralPaciente As Long
Dim mi_IdAtenInteGrupo As Long
Dim ml_IdPaciente As Long
Dim mda_FechaElaboracion As Date
'No persistir en la entidad de dominio
Dim ml_IdAtenInteItemPlan As Long
Dim ml_IdAtencion As Long

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdPlanIntegralPaciente(lValue As Long)
   ml_IdPlanIntegralPaciente = lValue
End Property
Property Get IdPlanIntegralPaciente() As Long
   IdPlanIntegralPaciente = ml_IdPlanIntegralPaciente
End Property
Property Let IdAtenInteGrupo(iValue As Long)
   mi_IdAtenInteGrupo = iValue
End Property
Property Get IdAtenInteGrupo() As Long
   IdAtenInteGrupo = mi_IdAtenInteGrupo
End Property
Property Let IdPaciente(iValue As Long)
   ml_IdPaciente = iValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let FechaElaboracion(daValue As Date)
   mda_FechaElaboracion = daValue
End Property
Property Get FechaElaboracion() As Date
   FechaElaboracion = mda_FechaElaboracion
End Property

Property Let IdAtenInteItemPlan(iValue As Long)
   ml_IdAtenInteItemPlan = iValue
End Property
Property Get IdAtenInteItemPlan() As Long
   IdAtenInteItemPlan = ml_IdAtenInteItemPlan
End Property


Property Let IdAtencion(iValue As Long)
   ml_IdAtencion = iValue
End Property

Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property
