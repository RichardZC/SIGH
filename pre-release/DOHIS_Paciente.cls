VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOHIS_Paciente"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla HIS_Paciente
'        Programado por: Cachay F
'        Fecha: Octubre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ml_IdHisPaciente As Long
Dim ms_NroHC_FF As String
'Dim ml_TipoHistoria As Long
Dim ml_Sexo As Long
Dim ml_IdNacionalidad As Long
Dim ms_NroDocIdentidad As String
Dim ms_NroHijo As String
Dim ms_IdEtnia As String
Dim ml_IdPacienteGalenHos As Long
Dim ml_IdTipoDocumento As Long

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

'Property Let TipoHistoria(iValue As Long)
'   ml_TipoHistoria = iValue
'End Property
'Property Get TipoHistoria() As Long
'   TipoHistoria = ml_TipoHistoria
'End Property

Property Let Sexo(iValue As Long)
   ml_Sexo = iValue
End Property
Property Get Sexo() As Long
   Sexo = ml_Sexo
End Property
Property Let IdNacionalidad(iValue As Long)
   ml_IdNacionalidad = iValue
End Property
Property Get IdNacionalidad() As Long
   IdNacionalidad = ml_IdNacionalidad
End Property
Property Let NroDocIdentidad(sValue As String)
   ms_NroDocIdentidad = sValue
End Property
Property Get NroDocIdentidad() As String
   NroDocIdentidad = ms_NroDocIdentidad
End Property
Property Let NroHijo(sValue As String)
   ms_NroHijo = sValue
End Property
Property Get NroHijo() As String
   NroHijo = ms_NroHijo
End Property
Property Let IdEtnia(sValue As String)
   ms_IdEtnia = sValue
End Property
Property Get IdEtnia() As String
   IdEtnia = ms_IdEtnia
End Property
Property Let IdPacienteGalenHos(iValue As Long)
   ml_IdPacienteGalenHos = iValue
End Property
Property Get IdPacienteGalenHos() As Long
   IdPacienteGalenHos = ml_IdPacienteGalenHos
End Property
Property Let IdTipoDocumento(iValue As Long)
   ml_IdTipoDocumento = iValue
End Property
Property Get IdTipoDocumento() As Long
   IdTipoDocumento = ml_IdTipoDocumento
End Property
