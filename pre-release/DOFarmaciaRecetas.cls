VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOFarmaciaRecetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOFarmaciaRecetas
'        Autor: William Castro Grijalva
'        Fecha: 05/08/2005 08:46:43
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_IdMedicoOrdena As Long
Dim ml_IdServicioOrdena As Long
Dim ms_FechaReceta As String
Dim ms_NroReceta As String
Dim ml_IdCuentaAtencion As Long
Dim ml_IdReceta As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdMedicoOrdena(lValue As Long)
   ml_IdMedicoOrdena = lValue
End Property
Property Get IdMedicoOrdena() As Long
   IdMedicoOrdena = ml_IdMedicoOrdena
End Property
Property Let IdServicioOrdena(lValue As Long)
   ml_IdServicioOrdena = lValue
End Property
Property Get IdServicioOrdena() As Long
   IdServicioOrdena = ml_IdServicioOrdena
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
Property Let IdReceta(lValue As Long)
   ml_IdReceta = lValue
End Property
Property Get IdReceta() As Long
   IdReceta = ml_IdReceta
End Property

