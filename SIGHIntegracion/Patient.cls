VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Patient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: clase para integracion de pacientes con otros sistemas
'        Programado por: Garay M
'        Fecha: Octubre 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_paciente As doPaciente
'tipo de sistema : imagenologia, laboratorio
Dim ml_SystemTypeSource As Integer
'Proveedor del sistema
Dim ml_SystemProvider As Integer
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_PatientFactory As New PatientFactory
Dim mo_DOPais As New DOPais
Dim mo_DoDistrito As New DODistrito


Private Sub Class_Initialize()
    'se puede crear parametros que almacenen los tipos de sistemas con los que se puede integrar Sisgalenhos
    'sistema de imagenologia RIS/PACS
'    ml_SystemTypeSource = sghIntegracionTipoSistema.sghRisPacs
    'Proveedor Carestream
    ml_SystemProvider = sghIntegracionProveedorSistema.sghCarestream
End Sub

Property Set Paciente(iValue As doPaciente)
   Set mo_paciente = iValue
End Property

Property Get Paciente() As doPaciente
   Set Paciente = mo_paciente
End Property

Property Set Distrito(iValue As DODistrito)
   Set mo_DoDistrito = iValue
End Property

Property Get Distrito() As DODistrito
   Set Distrito = mo_DoDistrito
End Property

Property Set Pais(iValue As DOPais)
   Set mo_DOPais = iValue
End Property

Property Get Pais() As DOPais
   Set Pais = mo_DOPais
End Property

Property Let SystemTypeSource(iValue As Integer)
   ml_SystemTypeSource = iValue
End Property

Property Get SystemTypeSource() As Integer
   SystemTypeSource = ml_SystemTypeSource
End Property

Property Let SystemProvider(iValue As Integer)
   ml_SystemProvider = iValue
End Property

Property Get SystemProvider() As Integer
   SystemProvider = ml_SystemProvider
End Property

Public Function NuevoPaciente() As doPaciente
    Call setConfigurationPatienteFactory
    Set NuevoPaciente = mo_PatientFactory.NuevoPaciente
End Function

Public Function EditarPaciente() As doPaciente
    Call setConfigurationPatienteFactory
    Set EditarPaciente = mo_PatientFactory.EditarPaciente
End Function

Public Function UnificarPaciente(oDoPacienteAntiguo As doPaciente) As doPaciente
    Call setConfigurationPatienteFactory
    Set UnificarPaciente = mo_PatientFactory.UnificarPaciente(oDoPacienteAntiguo)
End Function

Private Function setConfigurationPatienteFactory() As Boolean
    mo_PatientFactory.SystemProvider = ml_SystemProvider
    mo_PatientFactory.SystemTypeSource = ml_SystemTypeSource
    Set mo_PatientFactory.Paciente = mo_paciente
    
    Set mo_PatientFactory.Pais = mo_DOPais
    Set mo_PatientFactory.Distrito = mo_DoDistrito
End Function
