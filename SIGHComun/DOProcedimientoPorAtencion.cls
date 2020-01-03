VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtencionProcedimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOAtencionesProcedimientos
'        Autor: William Castro Grijalva
'        Fecha: 17/03/2005 12:10:57 a.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ms_NroOrden As String
Dim ml_IdMedicoRealiza As Long
Dim ml_IdServicioRealiza As Long
Dim ml_IdDetalleProducto As Long
Dim ml_IdCuentaAtencion As Long
Dim ms_HoraRealizacion As String
Dim mda_FechaRealizacion As Date
Dim ml_IdProcedimiento As Long
Dim ml_IdAtencionProcedimiento As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let NroOrden(sValue As String)
   ms_NroOrden = sValue
End Property
Property Get NroOrden() As String
   NroOrden = ms_NroOrden
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
Property Let IdDetalleProducto(lValue As Long)
   ml_IdDetalleProducto = lValue
End Property
Property Get IdDetalleProducto() As Long
   IdDetalleProducto = ml_IdDetalleProducto
End Property
Property Let IdCuentaAtencion(lValue As Long)
   ml_IdCuentaAtencion = lValue
End Property
Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
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
Property Let IdProcedimiento(lValue As Long)
   ml_IdProcedimiento = lValue
End Property
Property Get IdProcedimiento() As Long
   IdProcedimiento = ml_IdProcedimiento
End Property
Property Let IdAtencionProcedimiento(lValue As Long)
   ml_IdAtencionProcedimiento = lValue
End Property
Property Get IdAtencionProcedimiento() As Long
   IdAtencionProcedimiento = ml_IdAtencionProcedimiento
End Property


