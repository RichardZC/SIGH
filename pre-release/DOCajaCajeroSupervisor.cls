VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCajaCajeroSupervisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOCajaCajeroSupervisor
'        Autor: William Castro Grijalva
'        Fecha: 16/08/2005 11:06:17
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_IdSupervisor As Long
Dim ml_IdTurno As Long
Dim ml_IdCajero As Long
Dim ml_IdCajeroSupervisor As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdSupervisor(lValue As Long)
   ml_IdSupervisor = lValue
End Property
Property Get IdSupervisor() As Long
   IdSupervisor = ml_IdSupervisor
End Property
Property Let IdTurno(lValue As Long)
   ml_IdTurno = lValue
End Property
Property Get IdTurno() As Long
   IdTurno = ml_IdTurno
End Property
Property Let IdCajero(lValue As Long)
   ml_IdCajero = lValue
End Property
Property Get IdCajero() As Long
   IdCajero = ml_IdCajero
End Property
Property Let IdCajeroSupervisor(lValue As Long)
   ml_IdCajeroSupervisor = lValue
End Property
Property Get IdCajeroSupervisor() As Long
   IdCajeroSupervisor = ml_IdCajeroSupervisor
End Property

