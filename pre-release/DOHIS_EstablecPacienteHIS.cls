VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOHIS_EstablecPacienteHIS"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla DOHIS_EstablecPacienteHIS
'        Programado por: Cachay F
'        Fecha: Octubre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ml_IdEstablecPacienteHIS As Long
Dim ml_IdEstablecimiento As Long
Dim ml_IdHisPaciente As Long
Dim ms_NroHC_FF As String
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
Property Let IdEstablecPacienteHIS(iValue As Long)
   ml_IdEstablecPacienteHIS = iValue
End Property
Property Get IdEstablecPacienteHIS() As Long
   IdEstablecPacienteHIS = ml_IdEstablecPacienteHIS
End Property
Property Let IdEstablecimiento(iValue As Long)
   ml_IdEstablecimiento = iValue
End Property
Property Get IdEstablecimiento() As Long
   IdEstablecimiento = ml_IdEstablecimiento
End Property
Property Let IdHisPaciente(iValue As Long)
   ml_IdHisPaciente = iValue
End Property
Property Get IdHisPaciente() As Long
   IdHisPaciente = ml_IdHisPaciente
End Property
Property Let NroHC_FF(sValue As String)
   ms_NroHC_FF = sValue
End Property
Property Get NroHC_FF() As String
   NroHC_FF = ms_NroHC_FF
End Property

