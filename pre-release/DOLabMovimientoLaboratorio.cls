VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOLabMovimientoLaboratorio"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla LabMovimientoLaboratorio
'        Programado por: Bonilla A
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdMovimiento As Long
Dim ml_IdOrden As Long
Dim ml_CorrelativoAnual As Long
Dim ml_IdCuentaAtencion As Long
Dim ml_IdComprobantePago As Long
Dim ml_idPersonaTomaLab As Long
Dim ml_idPersonaRecoge As Long
Dim ml_IdDiagnostico As Long
Dim ml_EsDiagnosticoDefinitivo As Long
Dim ms_OrdenaPrueba As String
Dim mda_FechaNacimiento As Date
Dim ml_IdTipoSexo As Long
Dim ms_paciente As String
Dim mda_Eo_FUM As Date
Dim ml_Eo_EG As Long
Dim ms_Colegiatura As String
Property Get Colegiatura() As String
   Colegiatura = ms_Colegiatura
End Property
Property Let Colegiatura(sValue As String)
   ms_Colegiatura = sValue
End Property



Property Let Eo_EG(lValue As Long)
   ml_Eo_EG = lValue
End Property
Property Get Eo_EG() As Long
   Eo_EG = ml_Eo_EG
End Property
Property Let Eo_FUM(daValue As Date)
   mda_Eo_FUM = daValue
End Property
Property Get Eo_FUM() As Date
   Eo_FUM = mda_Eo_FUM
End Property
Property Get paciente() As String
   paciente = ms_paciente
End Property
Property Let paciente(sValue As String)
   ms_paciente = sValue
End Property
Property Let IdTipoSexo(lValue As Long)
   ml_IdTipoSexo = lValue
End Property
Property Get IdTipoSexo() As Long
   IdTipoSexo = ml_IdTipoSexo
End Property
Property Let FechaNacimiento(daValue As Date)
   mda_FechaNacimiento = daValue
End Property
Property Get FechaNacimiento() As Date
   FechaNacimiento = mda_FechaNacimiento
End Property



Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property

Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let IdMovimiento(iValue As Long)
   ml_IdMovimiento = iValue
End Property

Property Get IdMovimiento() As Long
   IdMovimiento = ml_IdMovimiento
End Property

Property Let IdOrden(iValue As Long)
   ml_IdOrden = iValue
End Property

Property Get IdOrden() As Long
   IdOrden = ml_IdOrden
End Property

Property Let CorrelativoAnual(iValue As Long)
   ml_CorrelativoAnual = iValue
End Property

Property Get CorrelativoAnual() As Long
   CorrelativoAnual = ml_CorrelativoAnual
End Property

Property Let IdCuentaAtencion(iValue As Long)
   ml_IdCuentaAtencion = iValue
End Property

Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
End Property

Property Let IdComprobantePago(iValue As Long)
   ml_IdComprobantePago = iValue
End Property

Property Get IdComprobantePago() As Long
   IdComprobantePago = ml_IdComprobantePago
End Property

Property Let idPersonaTomaLab(iValue As Long)
   ml_idPersonaTomaLab = iValue
End Property

Property Get idPersonaTomaLab() As Long
   idPersonaTomaLab = ml_idPersonaTomaLab
End Property

Property Let idPersonaRecoge(iValue As Long)
   ml_idPersonaRecoge = iValue
End Property

Property Get idPersonaRecoge() As Long
   idPersonaRecoge = ml_idPersonaRecoge
End Property

Property Let IdDiagnostico(iValue As Long)
   ml_IdDiagnostico = iValue
End Property

Property Get IdDiagnostico() As Long
   IdDiagnostico = ml_IdDiagnostico
End Property

Property Let EsDiagnosticoDefinitivo(iValue As Long)
   ml_EsDiagnosticoDefinitivo = iValue
End Property

Property Get EsDiagnosticoDefinitivo() As Long
   EsDiagnosticoDefinitivo = ml_EsDiagnosticoDefinitivo
End Property

Property Let OrdenaPrueba(sValue As String)
   ms_OrdenaPrueba = sValue
End Property

Property Get OrdenaPrueba() As String
   OrdenaPrueba = ms_OrdenaPrueba
End Property

