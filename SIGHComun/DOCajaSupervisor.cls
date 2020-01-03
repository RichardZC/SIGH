VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCajaSupervisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'MZD Ini 02/06/2005 [Todo el archivo]

Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOCajaSupervisor
'        Autor: William Castro Grijalva
'        Fecha: 03/06/2005 10:24:53
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_IdEmpleado As Long
Dim ms_EstadoSupervisor As String
Dim ml_IdSupervisor As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdEmpleado(lValue As Long)
   ml_IdEmpleado = lValue
End Property
Property Get IdEmpleado() As Long
   IdEmpleado = ml_IdEmpleado
End Property
Property Let EstadoSupervisor(sValue As String)
   ms_EstadoSupervisor = sValue
End Property
Property Get EstadoSupervisor() As String
   EstadoSupervisor = ms_EstadoSupervisor
End Property
Property Let IdSupervisor(lValue As Long)
   ml_IdSupervisor = lValue
End Property
Property Get IdSupervisor() As Long
   IdSupervisor = ml_IdSupervisor
End Property

