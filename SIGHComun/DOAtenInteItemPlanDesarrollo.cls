VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtenInteItemPlanDesarrollo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtenInteItemPlanDesarrollo
'        Programado por: Garay M
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdItemPlanDesarrollo As Long
Dim ml_IdPlanAtencion As Long
Dim ml_NumeroSesion As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdItemPlanDesarrollo(lValue As Long)
   ml_IdItemPlanDesarrollo = lValue
End Property
Property Get IdItemPlanDesarrollo() As Long
   IdItemPlanDesarrollo = ml_IdItemPlanDesarrollo
End Property
Property Let IdPlanAtencion(iValue As Long)
   ml_IdPlanAtencion = iValue
End Property
Property Get IdPlanAtencion() As Long
   IdPlanAtencion = ml_IdPlanAtencion
End Property
Property Let NumeroSesion(iValue As Long)
   ml_NumeroSesion = iValue
End Property
Property Get NumeroSesion() As Long
   NumeroSesion = ml_NumeroSesion
End Property
