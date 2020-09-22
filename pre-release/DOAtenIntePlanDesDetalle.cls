VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtenIntePlanDesDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtenIntePlanDesDetalle
'        Programado por: Garay M
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdItemPlanDesarrollo As Long
Dim ml_IdItemDesarrollo As Long
Dim ml_OrdenItem As Long
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
