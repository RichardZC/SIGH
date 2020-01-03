VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Doctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: clase para integracion de Medicos con otros sistemas
'        Programado por: Garay M
'        Fecha: Octubre 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Medico As DOMedico
Dim mo_Empleado As DOEmpleado
'tipo de sistema : imagenologia, laboratorio
Dim ml_SystemTypeSource As Integer
'Proveedor del sistema
Dim ml_SystemProvider As Integer
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_DoctorFactory As New DoctorFactory
Dim mo_DOPais As New DOPais
Dim mo_DoDistrito As New DODistrito
Dim mo_Especialidades As Collection


Private Sub Class_Initialize()
    'se puede crear parametros que almacenen los tipos de sistemas con los que se puede integrar Sisgalenhos
    'sistema de imagenologia RIS/PACS
'    ml_SystemTypeSource = sghIntegracionTipoSistema.sghRisPacs
    'Proveedor Carestream
    ml_SystemProvider = sghIntegracionProveedorSistema.sghCarestream
End Sub

Property Set Medico(iValue As DOMedico)
   Set mo_Medico = iValue
End Property

Property Get Medico() As DOMedico
   Set Medico = mo_Medico
End Property

Property Set Empleado(iValue As DOEmpleado)
   Set mo_Empleado = iValue
End Property

Property Get Empleado() As DOEmpleado
   Set Empleado = mo_Empleado
End Property

Property Set Especialidades(iValue As Collection)
   Set mo_Especialidades = iValue
End Property

Property Get Especialidades() As Collection
   Set Especialidades = mo_Especialidades
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

Public Function NuevoMedico() As DOMedico
    Call setConfigurationPatienteFactory
    Set NuevoMedico = mo_DoctorFactory.NuevoMedico
End Function

Public Function EditarMedico() As DOMedico
    Call setConfigurationPatienteFactory
    Set EditarMedico = mo_DoctorFactory.EditarMedico
End Function

Public Function UnificarMedico(oDoMedicoAntiguo As DOMedico) As DOMedico
    Call setConfigurationPatienteFactory
    Set UnificarMedico = mo_DoctorFactory.UnificarMedico(oDoMedicoAntiguo)
End Function

Private Function setConfigurationPatienteFactory() As Boolean
    mo_DoctorFactory.SystemProvider = ml_SystemProvider
    mo_DoctorFactory.SystemTypeSource = ml_SystemTypeSource
    Set mo_DoctorFactory.Medico = mo_Medico
    
    Set mo_DoctorFactory.Empleado = mo_Empleado
    Set mo_DoctorFactory.Especialidades = mo_Especialidades
    
    Set mo_DoctorFactory.Pais = mo_DOPais
    Set mo_DoctorFactory.Distrito = mo_DoDistrito
End Function

