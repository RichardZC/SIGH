VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCuentasEpisodioAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOCuentasEpisodioAtencion
'        Autor: William Castro Grijalva
'        Fecha: 04/04/2005 10:11:52 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim mda_FechaAlta As Date
Dim mda_FechaIngreso As Date
Dim ml_IdEpisodioAtencion As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let FechaAlta(daValue As Date)
   mda_FechaAlta = daValue
End Property
Property Get FechaAlta() As Date
   FechaAlta = mda_FechaAlta
End Property
Property Let FechaIngreso(daValue As Date)
   mda_FechaIngreso = daValue
End Property
Property Get FechaIngreso() As Date
   FechaIngreso = mda_FechaIngreso
End Property
Property Let IdEpisodioAtencion(lValue As Long)
   ml_IdEpisodioAtencion = lValue
End Property
Property Get IdEpisodioAtencion() As Long
   IdEpisodioAtencion = ml_IdEpisodioAtencion
End Property


