VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOHIS_CodigosActividades"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla HIS_CodigosActividades
'        Programado por: Cachay F
'        Fecha: Octubre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ml_IdHisCodActvidad As Long
Dim ml_IdTipoAtencion As Long
Dim ms_CodigoActividad As String
Dim ms_Descripcion As String
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
Property Let IdHisCodActvidad(iValue As Long)
   ml_IdHisCodActvidad = iValue
End Property
Property Get IdHisCodActvidad() As Long
   IdHisCodActvidad = ml_IdHisCodActvidad
End Property
Property Let IdTipoAtencion(iValue As Long)
   ml_IdTipoAtencion = iValue
End Property
Property Get IdTipoAtencion() As Long
   IdTipoAtencion = ml_IdTipoAtencion
End Property
Property Let CodigoActividad(sValue As String)
   ms_CodigoActividad = sValue
End Property
Property Get CodigoActividad() As String
   CodigoActividad = ms_CodigoActividad
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property

