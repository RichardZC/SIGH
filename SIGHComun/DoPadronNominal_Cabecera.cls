VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoPadronNominal_Cabecera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla PadronNominal_Cabecera
'        Programado por: Palomino Y
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ml_IdPaNomCabecera As Long
Dim ml_IdResponsableAtencion As Long
Dim ml_IdCodigoRenaes As Long
Dim ml_Mes As Long
Dim ms_Año As String
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
Property Let IdPaNomCabecera(iValue As Long)
   ml_IdPaNomCabecera = iValue
End Property
Property Get IdPaNomCabecera() As Long
   IdPaNomCabecera = ml_IdPaNomCabecera
End Property
Property Let IdResponsableAtencion(iValue As Long)
   ml_IdResponsableAtencion = iValue
End Property
Property Get IdResponsableAtencion() As Long
   IdResponsableAtencion = ml_IdResponsableAtencion
End Property
Property Let IdCodigoRenaes(iValue As Long)
   ml_IdCodigoRenaes = iValue
End Property
Property Get IdCodigoRenaes() As Long
   IdCodigoRenaes = ml_IdCodigoRenaes
End Property
Property Let Mes(iValue As Long)
   ml_Mes = iValue
End Property
Property Get Mes() As Long
   Mes = ml_Mes
End Property
Property Let Año(sValue As String)
   ms_Año = sValue
End Property
Property Get Año() As String
   Año = ms_Año
End Property

 
