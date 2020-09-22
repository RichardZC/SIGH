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
'        Inicio de código autogenerado para la clase: DOAtencionProcedimientos
'        Autor: William Castro Grijalva
'        Fecha: 17/07/2005 05:19:38 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_IdServicioOrdena As Long
Dim ml_IdMedicoOrdena As Long
Dim ml_IdCuentaAtencion As Long
Dim ms_HoraOrden As String
Dim mda_FechaOrden As Date
Dim ms_NroOrden As String
Dim ml_IdAtencionProcedimiento As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdServicioOrdena(lValue As Long)
   ml_IdServicioOrdena = lValue
End Property
Property Get IdServicioOrdena() As Long
   IdServicioOrdena = ml_IdServicioOrdena
End Property
Property Let IdMedicoOrdena(lValue As Long)
   ml_IdMedicoOrdena = lValue
End Property
Property Get IdMedicoOrdena() As Long
   IdMedicoOrdena = ml_IdMedicoOrdena
End Property
Property Let IdCuentaAtencion(lValue As Long)
   ml_IdCuentaAtencion = lValue
End Property
Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
End Property
Property Let HoraOrden(sValue As String)
   ms_HoraOrden = sValue
End Property
Property Get HoraOrden() As String
   HoraOrden = ms_HoraOrden
End Property
Property Let FechaOrden(daValue As Date)
   mda_FechaOrden = daValue
End Property
Property Get FechaOrden() As Date
   FechaOrden = mda_FechaOrden
End Property
Property Let NroOrden(sValue As String)
   ms_NroOrden = sValue
End Property
Property Get NroOrden() As String
   NroOrden = ms_NroOrden
End Property
Property Let IdAtencionProcedimiento(lValue As Long)
   ml_IdAtencionProcedimiento = lValue
End Property
Property Get IdAtencionProcedimiento() As Long
   IdAtencionProcedimiento = ml_IdAtencionProcedimiento
End Property


