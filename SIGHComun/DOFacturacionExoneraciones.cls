VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOFacturacionExoneraciones"
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
Dim ml_IdExoneracion As Long
Dim mda_FechaExoneracion As Date
Dim mc_TotalExonerado As Currency
Dim ml_IdEmpleadoExonera As Long

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

Property Let IdExoneracion(lValue As Long)
   ml_IdExoneracion = lValue
End Property
Property Get IdExoneracion() As Long
   IdExoneracion = ml_IdExoneracion
End Property
Property Let FechaExoneracion(daValue As Date)
   mda_FechaExoneracion = daValue
End Property
Property Get FechaExoneracion() As Date
   FechaExoneracion = mda_FechaExoneracion
End Property
Property Let IdEmpleadoExonera(lValue As Long)
   ml_IdEmpleadoExonera = lValue
End Property
Property Get IdEmpleadoExonera() As Long
   IdEmpleadoExonera = ml_IdEmpleadoExonera
End Property

Property Let TotalExonerado(lValue As Currency)
   mc_TotalExonerado = lValue
End Property
Property Get TotalExonerado() As Currency
   TotalExonerado = mc_TotalExonerado
End Property

