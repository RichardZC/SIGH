VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFarmMovimientoVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla farmMovimientoVentas
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ms_movNumero As String
Dim ms_movTipo As String
Dim ms_tipoVenta As String
Dim ml_idPreventa As Long
Dim ml_IdTipoFinanciamiento As Long
Dim ml_idPrescriptor As Long
Dim ml_idTipoReceta As Long
Dim ml_IdDiagnostico As Long
Dim ml_idCuentaAtencion As Long
Dim ml_IdPaciente As Long
Dim ml_idFuenteFinanciamiento As Long
Dim ml_IdServicioPaciente As Long
Dim mda_FechaHoraPrescribe As Date
Property Let FechaHoraPrescribe(daValue As Date)
   mda_FechaHoraPrescribe = daValue
End Property
Property Get FechaHoraPrescribe() As Date
   FechaHoraPrescribe = mda_FechaHoraPrescribe
End Property


Property Let IdServicioPaciente(lValue As Long)
   ml_IdServicioPaciente = lValue
End Property
Property Get IdServicioPaciente() As Long
   IdServicioPaciente = ml_IdServicioPaciente
End Property

Property Let idFuenteFinanciamiento(lValue As Long)
   ml_idFuenteFinanciamiento = lValue
End Property
Property Get idFuenteFinanciamiento() As Long
   idFuenteFinanciamiento = ml_idFuenteFinanciamiento
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
Property Let tipoVenta(sValue As String)
   ms_tipoVenta = sValue
End Property
Property Get tipoVenta() As String
   tipoVenta = ms_tipoVenta
End Property
Property Let idPreventa(iValue As Long)
   ml_idPreventa = iValue
End Property
Property Get idPreventa() As Long
   idPreventa = ml_idPreventa
End Property
Property Let IdTipoFinanciamiento(iValue As Long)
   ml_IdTipoFinanciamiento = iValue
End Property
Property Get IdTipoFinanciamiento() As Long
   IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property
Property Let idPrescriptor(iValue As Long)
   ml_idPrescriptor = iValue
End Property
Property Get idPrescriptor() As Long
   idPrescriptor = ml_idPrescriptor
End Property
Property Let idTipoReceta(iValue As Long)
   ml_idTipoReceta = iValue
End Property
Property Get idTipoReceta() As Long
   idTipoReceta = ml_idTipoReceta
End Property
Property Let IdDiagnostico(iValue As Long)
   ml_IdDiagnostico = iValue
End Property
Property Get IdDiagnostico() As Long
   IdDiagnostico = ml_IdDiagnostico
End Property
Property Let idCuentaAtencion(iValue As Long)
   ml_idCuentaAtencion = iValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let IdPaciente(iValue As Long)
   ml_IdPaciente = iValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property


