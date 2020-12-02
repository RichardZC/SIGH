VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOPlanFinanciamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOPlanesFinanciamiento
'        Autor: William Castro Grijalva
'        Fecha: 30/08/2004 11:37:57 a.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_IdPlan As Long
Dim ml_IdFuenteFinanciamiento As Long
Dim ml_IdTipoFinanciamiento As Long
Dim ml_IdPlanFinanciamiento As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdPlan(lValue As Long)
   ml_IdPlan = lValue
End Property
Property Get IdPlan() As Long
   IdPlan = ml_IdPlan
End Property
Property Let IdFuenteFinanciamiento(lValue As Long)
   ml_IdFuenteFinanciamiento = lValue
End Property
Property Get IdFuenteFinanciamiento() As Long
   IdFuenteFinanciamiento = ml_IdFuenteFinanciamiento
End Property
Property Let IdTipoFinanciamiento(lValue As Long)
   ml_IdTipoFinanciamiento = lValue
End Property
Property Get IdTipoFinanciamiento() As Long
   IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property
Property Let IdPlanFinanciamiento(lValue As Long)
   ml_IdPlanFinanciamiento = lValue
End Property
Property Get IdPlanFinanciamiento() As Long
   IdPlanFinanciamiento = ml_IdPlanFinanciamiento
End Property


