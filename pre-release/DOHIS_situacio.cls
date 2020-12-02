VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOHIS_situacio"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla HIS_situacio
'        Programado por: Cachay F
'        Fecha: Octubre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ml_IdHisSituacio As Long
Dim ms_valores As String
Dim ms_descripcio As String
Dim mv_codigo As Variant
Dim ms_est As String
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
Property Let IdHisSituacio(iValue As Long)
   ml_IdHisSituacio = iValue
End Property
Property Get IdHisSituacio() As Long
   IdHisSituacio = ml_IdHisSituacio
End Property
Property Let valores(sValue As String)
   ms_valores = sValue
End Property
Property Get valores() As String
   valores = ms_valores
End Property
Property Let descripcio(sValue As String)
   ms_descripcio = sValue
End Property
Property Get descripcio() As String
   descripcio = ms_descripcio
End Property
Property Let Codigo(vValue As Variant)
   mv_codigo = vValue
End Property
Property Get Codigo() As Variant
   Codigo = mv_codigo
End Property
Property Let est(sValue As String)
   ms_est = sValue
End Property
Property Get est() As String
   est = ms_est
End Property

