VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOPlanesCatalogoBienesInsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOPlanesCatalogoBienesInsumos
'        Autor: William Castro Grijalva
'        Fecha: 09/07/2005 08:28:23 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_IdProducto As Long
Dim mc_PrecioUnitario As Currency
Dim ml_IdPlan As Long
Dim ml_IdPlanCatalogo As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let PrecioUnitario(cValue As Currency)
   mc_PrecioUnitario = cValue
End Property
Property Get PrecioUnitario() As Currency
   PrecioUnitario = mc_PrecioUnitario
End Property
Property Let IdPlan(lValue As Long)
   ml_IdPlan = lValue
End Property
Property Get IdPlan() As Long
   IdPlan = ml_IdPlan
End Property
Property Let IdPlanCatalogo(lValue As Long)
   ml_IdPlanCatalogo = lValue
End Property
Property Get IdPlanCatalogo() As Long
   IdPlanCatalogo = ml_IdPlanCatalogo
End Property


