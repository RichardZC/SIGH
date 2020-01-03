VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOHIS_Lotes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla HIS_Lotes
'        Programado por: Cachay F
'        Fecha: Octubre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ml_IdHisLote As Long
Dim ml_IdEstablecimiento As Long
Dim ms_Lote As String
Dim ml_NroHojas As Long
Dim ml_Mes As Long
Dim ml_Anio As Long
Dim ml_IdEstadoLote As Long
Dim ml_Cerrado As Long
Dim ml_DobleDigitacion As Long

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
Property Let IdHisLote(iValue As Long)
   ml_IdHisLote = iValue
End Property
Property Get IdHisLote() As Long
   IdHisLote = ml_IdHisLote
End Property
Property Let IdEstablecimiento(iValue As Long)
   ml_IdEstablecimiento = iValue
End Property
Property Get IdEstablecimiento() As Long
   IdEstablecimiento = ml_IdEstablecimiento
End Property
Property Let Lote(sValue As String)
   ms_Lote = sValue
End Property
Property Get Lote() As String
   Lote = ms_Lote
End Property
Property Let NroHojas(iValue As Long)
   ml_NroHojas = iValue
End Property
Property Get NroHojas() As Long
   NroHojas = ml_NroHojas
End Property
Property Let Mes(iValue As Long)
   ml_Mes = iValue
End Property
Property Get Mes() As Long
   Mes = ml_Mes
End Property
Property Let Anio(iValue As Long)
   ml_Anio = iValue
End Property
Property Get Anio() As Long
   Anio = ml_Anio
End Property
Property Let IdEstadoLote(iValue As Long)
   ml_IdEstadoLote = iValue
End Property
Property Get IdEstadoLote() As Long
   IdEstadoLote = ml_IdEstadoLote
End Property
Property Let DobleDigitacion(iValue As Long)
   ml_DobleDigitacion = iValue
End Property
Property Get DobleDigitacion() As Long
   DobleDigitacion = ml_DobleDigitacion
End Property
Property Let Cerrado(iValue As Long)
   ml_Cerrado = iValue
End Property
Property Get Cerrado() As Long
   Cerrado = ml_Cerrado
End Property
