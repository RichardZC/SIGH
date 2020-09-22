VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtenIntePlanCrecDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtenIntePlanCrecDetalle
'        Programado por: Garay M
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdItemPlanCrecimiento As Long
Dim ml_IdTriajeVariable As Long
Dim ml_OrdenItem As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdItemPlanCrecimiento(lValue As Long)
   ml_IdItemPlanCrecimiento = lValue
End Property
Property Get IdItemPlanCrecimiento() As Long
   IdItemPlanCrecimiento = ml_IdItemPlanCrecimiento
End Property
Property Let IdTriajeVariable(iValue As Long)
   ml_IdTriajeVariable = iValue
End Property
Property Get IdTriajeVariable() As Long
   IdTriajeVariable = ml_IdTriajeVariable
End Property
Property Let OrdenItem(iValue As Long)
   ml_OrdenItem = iValue
End Property
Property Get OrdenItem() As Long
   OrdenItem = ml_OrdenItem
End Property
