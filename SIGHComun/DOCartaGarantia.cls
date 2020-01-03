VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCartaGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOCartaGarantia
'        Autor: William Castro Grijalva
'        Fecha: 11/09/2005 01:20:47
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_IdCuentaAtencion As Long
Dim mda_FechaVigencia As Date
Dim ms_NroCarta As String
Dim ms_Observacion As String
Dim mc_ValorCobertura As Currency
Dim ml_IdCartaGarantia As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdCuentaAtencion(lValue As Long)
   ml_IdCuentaAtencion = lValue
End Property
Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
End Property
Property Let FechaVigencia(daValue As Date)
   mda_FechaVigencia = daValue
End Property
Property Get FechaVigencia() As Date
   FechaVigencia = mda_FechaVigencia
End Property
Property Let NroCarta(sValue As String)
   ms_NroCarta = sValue
End Property
Property Get NroCarta() As String
   NroCarta = ms_NroCarta
End Property
Property Let Observacion(sValue As String)
   ms_Observacion = sValue
End Property
Property Get Observacion() As String
   Observacion = ms_Observacion
End Property
Property Let ValorCobertura(cValue As Currency)
   mc_ValorCobertura = cValue
End Property
Property Get ValorCobertura() As Currency
   ValorCobertura = mc_ValorCobertura
End Property
Property Let IdCartaGarantia(lValue As Long)
   ml_IdCartaGarantia = lValue
End Property
Property Get IdCartaGarantia() As Long
   IdCartaGarantia = ml_IdCartaGarantia
End Property

