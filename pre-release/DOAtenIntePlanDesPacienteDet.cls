VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtenIntePlanDesPacienteDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtenIntePlanDesPacienteDet
'        Programado por: Garay M
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdPlanDesarrolloPaciente As Long
Dim ml_IdPlanIntegralPaciente As Long
Dim ml_IdItemDesarrollo As Long
Dim ml_OrdenItem As Long
Dim mb_EjecutaAccion As Boolean
Dim mb_RespondioEjecutaAccion As Boolean

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdPlanDesarrolloPaciente(iValue As Long)
   ml_IdPlanDesarrolloPaciente = iValue
End Property
Property Get IdPlanDesarrolloPaciente() As Long
   IdPlanDesarrolloPaciente = ml_IdPlanDesarrolloPaciente
End Property
Property Let IdPlanIntegralPaciente(lValue As Long)
   ml_IdPlanIntegralPaciente = lValue
End Property
Property Get IdPlanIntegralPaciente() As Long
   IdPlanIntegralPaciente = ml_IdPlanIntegralPaciente
End Property
Property Let IdItemDesarrollo(iValue As Long)
   ml_IdItemDesarrollo = iValue
End Property
Property Get IdItemDesarrollo() As Long
   IdItemDesarrollo = ml_IdItemDesarrollo
End Property

Property Let OrdenItem(iValue As Long)
   ml_OrdenItem = iValue
End Property
Property Get OrdenItem() As Long
   OrdenItem = ml_OrdenItem
End Property

Property Let EjecutaAccion(bValue As Boolean)
   mb_EjecutaAccion = bValue
End Property
Property Get EjecutaAccion() As Boolean
   EjecutaAccion = mb_EjecutaAccion
End Property

Property Let RespondioEjecutaAccion(iValue As Boolean)
    mb_RespondioEjecutaAccion = iValue
End Property

Property Get RespondioEjecutaAccion() As Boolean
   RespondioEjecutaAccion = mb_RespondioEjecutaAccion
End Property
