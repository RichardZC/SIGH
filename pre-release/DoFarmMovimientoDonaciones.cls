VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFarmMovimientoDonaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FarmMovimientoDonaciones
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ms_movNumero As String
Dim ms_movTipo As String
Dim ml_idCuentaAtencion As Long
Dim ml_idPrescriptorReceta As Long
Dim ml_idCoordinadorServicioSocial As Long
Dim ml_IdDiagnostico As Long

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let movNumero(sValue As String)
   ms_movNumero = sValue
End Property
Property Get movNumero() As String
   movNumero = ms_movNumero
End Property
Property Let movTipo(sValue As String)
   ms_movTipo = sValue
End Property
Property Get movTipo() As String
   movTipo = ms_movTipo
End Property
Property Let idCuentaAtencion(iValue As Long)
   ml_idCuentaAtencion = iValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let idPrescriptorReceta(iValue As Long)
   ml_idPrescriptorReceta = iValue
End Property
Property Get idPrescriptorReceta() As Long
   idPrescriptorReceta = ml_idPrescriptorReceta
End Property
Property Let idCoordinadorServicioSocial(iValue As Long)
   ml_idCoordinadorServicioSocial = iValue
End Property
Property Get idCoordinadorServicioSocial() As Long
   idCoordinadorServicioSocial = ml_idCoordinadorServicioSocial
End Property
Property Let IdDiagnostico(iValue As Long)
   ml_IdDiagnostico = iValue
End Property
Property Get IdDiagnostico() As Long
   IdDiagnostico = ml_IdDiagnostico
End Property


