VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOHIS_DatosEstablecimiento"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla HIS_DatosEstablecimiento
'        Programado por: Cachay F
'        Fecha: Octubre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ml_IdDatoEstablec As Long
Dim ml_IdEstablecimiento As Long
Dim ms_Color As String
Dim ml_Turnos As Long
Dim ml_UltimoNroFormatoHIS As Long

Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let IdDatoEstablec(iValue As Long)
   ml_IdDatoEstablec = iValue
End Property
Property Get IdDatoEstablec() As Long
   IdDatoEstablec = ml_IdDatoEstablec
End Property
Property Let IdEstablecimiento(iValue As Long)
   ml_IdEstablecimiento = iValue
End Property
Property Get IdEstablecimiento() As Long
   IdEstablecimiento = ml_IdEstablecimiento
End Property
Property Let Color(sValue As String)
   ms_Color = sValue
End Property
Property Get Color() As String
   Color = ms_Color
End Property
Property Let Turnos(iValue As Long)
   ml_Turnos = iValue
End Property
Property Get Turnos() As Long
   Turnos = ml_Turnos
End Property
Property Let UltimoNroFormatoHIS(iValue As Long)
   ml_UltimoNroFormatoHIS = iValue
End Property
Property Get UltimoNroFormatoHIS() As Long
   UltimoNroFormatoHIS = ml_UltimoNroFormatoHIS
End Property

