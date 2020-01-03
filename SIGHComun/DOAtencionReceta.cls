VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtencionReceta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOAtencionesRecetas
'        Autor: William Castro Grijalva
'        Fecha: 27/10/2004 11:51:46 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_IdMedico As Long
Dim ml_IdServicio As Long
Dim ms_FechaReceta As String
Dim ms_NroReceta As String
Dim ml_IdCuentaAtencion As Long
Dim ms_IdAtencionReceta As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdMedico(lValue As Long)
   ml_IdMedico = lValue
End Property
Property Get IdMedico() As Long
   IdMedico = ml_IdMedico
End Property
Property Let IdServicio(lValue As Long)
   ml_IdServicio = lValue
End Property
Property Get IdServicio() As Long
   IdServicio = ml_IdServicio
End Property
Property Let FechaReceta(sValue As String)
   ms_FechaReceta = sValue
End Property
Property Get FechaReceta() As String
   FechaReceta = ms_FechaReceta
End Property
Property Let NroReceta(sValue As String)
   ms_NroReceta = sValue
End Property
Property Get NroReceta() As String
   NroReceta = ms_NroReceta
End Property
Property Let IdCuentaAtencion(lValue As Long)
   ml_IdCuentaAtencion = lValue
End Property
Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
End Property
Property Let IdAtencionReceta(sValue As String)
   ms_IdAtencionReceta = sValue
End Property
Property Get IdAtencionReceta() As String
   IdAtencionReceta = ms_IdAtencionReceta
End Property


