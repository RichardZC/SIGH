VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtencionProcDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOAtencionProcedimientoDetalle
'        Autor: William Castro Grijalva
'        Fecha: 14/06/2005 10:49:27 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_IdAtencionProcedimiento As Long
Dim ml_IdFacturacionServicio As Long
Dim ml_IdMedicoRealiza As Long
Dim ml_IdServicioRealiza As Long
Dim ml_IdProcedimiento As Long
Dim ms_HoraRealizacion As String
Dim mda_FechaRealizacion As Date
Dim ml_IdAtencionProcDetalle As Long
Dim ms_EstadoRegistro As String

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdAtencionProcedimiento(lValue As Long)
   ml_IdAtencionProcedimiento = lValue
End Property
Property Get IdAtencionProcedimiento() As Long
   IdAtencionProcedimiento = ml_IdAtencionProcedimiento
End Property
Property Let IdFacturacionServicio(lValue As Long)
   ml_IdFacturacionServicio = lValue
End Property
Property Get IdFacturacionServicio() As Long
   IdFacturacionServicio = ml_IdFacturacionServicio
End Property
Property Let IdMedicoRealiza(lValue As Long)
   ml_IdMedicoRealiza = lValue
End Property
Property Get IdMedicoRealiza() As Long
   IdMedicoRealiza = ml_IdMedicoRealiza
End Property
Property Let IdServicioRealiza(lValue As Long)
   ml_IdServicioRealiza = lValue
End Property
Property Get IdServicioRealiza() As Long
   IdServicioRealiza = ml_IdServicioRealiza
End Property
Property Let IdProcedimiento(lValue As Long)
   ml_IdProcedimiento = lValue
End Property
Property Get IdProcedimiento() As Long
   IdProcedimiento = ml_IdProcedimiento
End Property
Property Let HoraRealizacion(sValue As String)
   ms_HoraRealizacion = sValue
End Property
Property Get HoraRealizacion() As String
   HoraRealizacion = ms_HoraRealizacion
End Property
Property Let FechaRealizacion(daValue As Date)
   mda_FechaRealizacion = daValue
End Property
Property Get FechaRealizacion() As Date
   FechaRealizacion = mda_FechaRealizacion
End Property
Property Let IdAtencionProcDetalle(lValue As Long)
   ml_IdAtencionProcDetalle = lValue
End Property
Property Get IdAtencionProcDetalle() As Long
   IdAtencionProcDetalle = ml_IdAtencionProcDetalle
End Property

Property Let EstadoRegistro(sValue As String)
   ms_EstadoRegistro = sValue
End Property
Property Get EstadoRegistro() As String
   EstadoRegistro = ms_EstadoRegistro
End Property

