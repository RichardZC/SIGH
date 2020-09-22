VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOEnvioHistoriaClinica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOEnviosHistoriaClinica
'        Autor: William Castro Grijalva
'        Fecha: 04/09/2004 12:02:55 a.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ms_HoraPrestamoReal As String
Dim mda_FechaPrestamoReal As Date
Dim ml_IdResponsableRecepcion As Long
Dim ml_IdResponsableEnvio As Long
Dim ml_IdEnvio As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let HoraPrestamoReal(sValue As String)
   ms_HoraPrestamoReal = sValue
End Property
Property Get HoraPrestamoReal() As String
   HoraPrestamoReal = ms_HoraPrestamoReal
End Property
Property Let FechaPrestamoReal(daValue As Date)
   mda_FechaPrestamoReal = daValue
End Property
Property Get FechaPrestamoReal() As Date
   FechaPrestamoReal = mda_FechaPrestamoReal
End Property
Property Let IdResponsableRecepcion(lValue As Long)
   ml_IdResponsableRecepcion = lValue
End Property
Property Get IdResponsableRecepcion() As Long
   IdResponsableRecepcion = ml_IdResponsableRecepcion
End Property
Property Let IdResponsableEnvio(lValue As Long)
   ml_IdResponsableEnvio = lValue
End Property
Property Get IdResponsableEnvio() As Long
   IdResponsableEnvio = ml_IdResponsableEnvio
End Property
Property Let IdEnvio(lValue As Long)
   ml_IdEnvio = lValue
End Property
Property Get IdEnvio() As Long
   IdEnvio = ml_IdEnvio
End Property


