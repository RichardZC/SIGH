VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOFacturacionPAgosACuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOFacturacionBienesInsumos
'        Autor: William Castro Grijalva
'        Fecha: 14/04/2006 07:06:56 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_IdAtencion  As Long
Dim ml_IdPagosACuenta As Long
Dim mc_TotalPagado As Currency
Dim mda_FechaPago As Date
Dim ml_IdComprobantePago As Long
Dim ml_IdEmpleadoCajero As Long

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdAtencion(lValue As Long)
   ml_IdAtencion = lValue
End Property

Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property

Property Let IdPagosACuenta(lValue As Long)
   ml_IdPagosACuenta = lValue
End Property
Property Get IdPagosACuenta() As Long
   IdPagosACuenta = ml_IdPagosACuenta
End Property

Property Let TotalPagado(lValue As Currency)
   mc_TotalPagado = lValue
End Property
Property Get TotalPagado() As Currency
   TotalPagado = mc_TotalPagado
End Property

Property Let FechaPago(daValue As Date)
   mda_FechaPago = daValue
End Property
Property Get FechaPago() As Date
   FechaPago = mda_FechaPago
End Property

Property Let IdComprobantePago(lValue As Long)
   ml_IdComprobantePago = lValue
End Property
Property Get IdComprobantePago() As Long
   IdComprobantePago = ml_IdComprobantePago
End Property

Property Let IdEmpleadoCajero(lValue As Long)
   ml_IdEmpleadoCajero = lValue
End Property
Property Get IdEmpleadoCajero() As Long
   IdEmpleadoCajero = ml_IdEmpleadoCajero
End Property
