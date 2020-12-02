VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOHIS_Cabecera"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla HIS_Cabecera
'        Programado por: Cachay F
'        Fecha: Octubre 2013
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ml_IdHisCabecera As Long
Dim ml_IdHisLote As Long
Dim ml_NroHojaHis As Long
Dim ml_NroFormato As Long
Dim ml_idTurno As Long
Dim ml_idUsuario As Long
Dim ml_IdEstadoHis As Long
Dim ml_idMedico As Long
Dim ml_IdEstablecimiento As Long
Dim ml_IdServicio As Long
Dim mda_FechaCreacion As Date

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
Property Let IdHisCabecera(iValue As Long)
   ml_IdHisCabecera = iValue
End Property
Property Get IdHisCabecera() As Long
   IdHisCabecera = ml_IdHisCabecera
End Property
Property Let IdHisLote(iValue As Long)
   ml_IdHisLote = iValue
End Property
Property Get IdHisLote() As Long
   IdHisLote = ml_IdHisLote
End Property
Property Let NroHojaHis(iValue As Long)
   ml_NroHojaHis = iValue
End Property
Property Get NroHojaHis() As Long
   NroHojaHis = ml_NroHojaHis
End Property
Property Let NroFormato(iValue As Long)
   ml_NroFormato = iValue
End Property
Property Get NroFormato() As Long
   NroFormato = ml_NroFormato
End Property
Property Let idTurno(iValue As Long)
   ml_idTurno = iValue
End Property
Property Get idTurno() As Long
   idTurno = ml_idTurno
End Property
Property Let idUsuario(iValue As Long)
   ml_idUsuario = iValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let IdEstadoHis(iValue As Long)
   ml_IdEstadoHis = iValue
End Property
Property Get IdEstadoHis() As Long
   IdEstadoHis = ml_IdEstadoHis
End Property
Property Let idMedico(iValue As Long)
   ml_idMedico = iValue
End Property
Property Get idMedico() As Long
   idMedico = ml_idMedico
End Property
Property Let IdEstablecimiento(iValue As Long)
   ml_IdEstablecimiento = iValue
End Property
Property Get IdEstablecimiento() As Long
   IdEstablecimiento = ml_IdEstablecimiento
End Property
Property Let IdServicio(iValue As Long)
   ml_IdServicio = iValue
End Property
Property Get IdServicio() As Long
   IdServicio = ml_IdServicio
End Property
Property Let FechaCreacion(daValue As Date)
   mda_FechaCreacion = daValue
End Property
Property Get FechaCreacion() As Date
   FechaCreacion = mda_FechaCreacion
End Property


