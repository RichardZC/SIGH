VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOPlanes
'        Autor: William Castro Grijalva
'        Fecha: 04/09/2004 04:04:01 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim mc_Deducible As Currency
Dim mc_Coaseguro As Currency
Dim ms_Descripcion As String
Dim ml_IdPlan As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Deducible(cValue As Currency)
   mc_Deducible = cValue
End Property
Property Get Deducible() As Currency
   Deducible = mc_Deducible
End Property
Property Let Coaseguro(cValue As Currency)
   mc_Coaseguro = cValue
End Property
Property Get Coaseguro() As Currency
   Coaseguro = mc_Coaseguro
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let IdPlan(lValue As Long)
   ml_IdPlan = lValue
End Property
Property Get IdPlan() As Long
   IdPlan = ml_IdPlan
End Property


