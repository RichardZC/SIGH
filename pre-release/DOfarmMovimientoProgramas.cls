VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOfarmMovimientoProgramas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla farmMovimientoProgramas
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ms_movNumero As String
Dim ms_movTipo As String
Dim ml_idCoordinador As Long
Dim ms_idPrescriptor As String
Dim ml_IdDiagnostico As Long
Dim ml_IdPaciente As Long
Dim ml_idComponente As Long
Dim ml_idSubComponente As Long
Dim mda_FechaHoraPrescribe As Date
Property Let FechaHoraPrescribe(daValue As Date)
   mda_FechaHoraPrescribe = daValue
End Property
Property Get FechaHoraPrescribe() As Date
   FechaHoraPrescribe = mda_FechaHoraPrescribe
End Property


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
Property Let idCoordinador(iValue As Long)
   ml_idCoordinador = iValue
End Property
Property Get idCoordinador() As Long
   idCoordinador = ml_idCoordinador
End Property
Property Let idPrescriptor(sValue As String)
   ms_idPrescriptor = sValue
End Property
Property Get idPrescriptor() As String
   idPrescriptor = ms_idPrescriptor
End Property
Property Let IdDiagnostico(iValue As Long)
   ml_IdDiagnostico = iValue
End Property
Property Get IdDiagnostico() As Long
   IdDiagnostico = ml_IdDiagnostico
End Property
Property Let IdPaciente(iValue As Long)
   ml_IdPaciente = iValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let idComponente(iValue As Long)
   ml_idComponente = iValue
End Property
Property Get idComponente() As Long
   idComponente = ml_idComponente
End Property
Property Let idSubComponente(iValue As Long)
   ml_idSubComponente = iValue
End Property
Property Get idSubComponente() As Long
   idSubComponente = ml_idSubComponente
End Property
