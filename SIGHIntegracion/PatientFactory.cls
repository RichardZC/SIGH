VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PatientFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: clase para instaciar las clase que implementan las intregraciones con los diferenctes proveedores
'        Programado por: Garay M
'        Fecha: Octubre 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_paciente As doPaciente

Dim ml_SystemTypeSource As Integer
Dim ml_SystemProvider As Integer
Dim mo_DOPais As New DOPais
Dim mo_DoDistrito As New DODistrito

'Dim mo_IntegracionCarestream As New IntegracionCarestream

'==================================================================
'Propiedades de objetos de datos
'==================================================================

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


'==================================================================

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

'Expotar un nuevo paciente
Public Function NuevoPaciente() As doPaciente
    Dim returnValue As doPaciente
    
    Select Case ml_SystemTypeSource
        Case sghIntegracionTipoSistema.sghRisPacs:
            Set returnValue = NuevoPacienteRisPacs(ml_SystemProvider)
        Case Else:
            MsgBox messageErrorIntegrationSystemType(), vbInformation, "Integración"
            Set returnValue = Nothing
    End Select
    Set NuevoPaciente = returnValue
End Function

'Editar un nuevo paciente
Public Function EditarPaciente() As doPaciente
    Dim returnValue As doPaciente
    
    Select Case ml_SystemTypeSource
        Case sghIntegracionTipoSistema.sghRisPacs:
            Set returnValue = EditarPacienteRisPacs(ml_SystemProvider)
        Case Else:
            MsgBox messageErrorIntegrationSystemType(), vbInformation, "Integración"
            Set returnValue = Nothing
    End Select
    Set EditarPaciente = returnValue
End Function

'Unir dos registros que pertezcan a un solo paciente
Public Function UnificarPaciente(oDoPacienteAntiguo As doPaciente) As doPaciente
    Dim returnValue As doPaciente
    
    Select Case ml_SystemTypeSource
        Case sghIntegracionTipoSistema.sghRisPacs:
            Set returnValue = UnificarPacienteRisPacs(ml_SystemProvider, oDoPacienteAntiguo)
        Case Else:
            MsgBox messageErrorIntegrationSystemType(), vbInformation, "Integración"
            Set returnValue = Nothing
    End Select
    Set UnificarPaciente = returnValue
End Function

'
Private Function NuevoPacienteRisPacs(lSystemProvider As Integer) As doPaciente
    Dim returnValue As doPaciente
    Select Case lSystemProvider
        Case sghIntegracionProveedorSistema.sghCarestream:
            Dim mo_IntegracionCarestream As New IntegracionCarestream
            
            Set mo_IntegracionCarestream.Paciente = mo_paciente
            Set mo_IntegracionCarestream.Distrito = mo_DoDistrito
            Set mo_IntegracionCarestream.Pais = mo_DOPais
            
            Set returnValue = mo_IntegracionCarestream.NuevoPaciente
        Case Else:
            MsgBox messageErrorIntegrationSystemProvider(), vbInformation, "Integración"
            Set returnValue = Nothing
    End Select
    Set NuevoPacienteRisPacs = returnValue
End Function

Private Function EditarPacienteRisPacs(lSystemProvider As Integer) As doPaciente
    Dim returnValue As doPaciente
    Select Case lSystemProvider
        Case sghIntegracionProveedorSistema.sghCarestream:
            Dim mo_IntegracionCarestream As New IntegracionCarestream
            
            Set mo_IntegracionCarestream.Paciente = mo_paciente
            Set mo_IntegracionCarestream.Distrito = mo_DoDistrito
            Set mo_IntegracionCarestream.Pais = mo_DOPais
            
            Set returnValue = mo_IntegracionCarestream.EditarPaciente
        Case Else:
            MsgBox messageErrorIntegrationSystemProvider(), vbInformation, "Integración"
            Set returnValue = Nothing
    End Select
    Set EditarPacienteRisPacs = returnValue
End Function

'
Private Function UnificarPacienteRisPacs(lSystemProvider As Integer, oDoPacienteAntiguo As doPaciente) As doPaciente
    Dim returnValue As doPaciente
    Select Case lSystemProvider
        Case sghIntegracionProveedorSistema.sghCarestream:
            Dim mo_IntegracionCarestream As New IntegracionCarestream
            
            Set mo_IntegracionCarestream.Paciente = mo_paciente
            Set mo_IntegracionCarestream.Distrito = mo_DoDistrito
            Set mo_IntegracionCarestream.Pais = mo_DOPais
            
            Set returnValue = mo_IntegracionCarestream.UnificarPaciente(oDoPacienteAntiguo)
        Case Else:
            MsgBox messageErrorIntegrationSystemProvider(), vbInformation, "Integración"
            Set returnValue = Nothing
    End Select
    Set UnificarPacienteRisPacs = returnValue
End Function

'Mensaje que se mostrara en caso no se halla implementado exportacion de datos para un sistema especifico
Private Function messageErrorIntegrationSystemType() As String
    messageErrorIntegrationSystemType = "No se ha Implementado Integracion con Tipo de Sistema : " & getSystemTypeSourceName(ml_SystemTypeSource)
End Function

'Mensaje de error que se mostrara en caso no exista los proveedores
Private Function messageErrorIntegrationSystemProvider() As String
    messageErrorIntegrationSystemProvider = "No se ha Implementado Integracion con Tipo de Sistema : " & getSystemTypeSourceName(ml_SystemTypeSource) _
                                    & ", del proveedor : " & getSystemProviderName(ml_SystemProvider)
End Function

'Obtiene el nombre del sistema al cual se le enviaran los datos del paciente
Private Function getSystemTypeSourceName(lSystemTypeSource As Integer) As String
    Select Case lSystemTypeSource
        Case sghIntegracionTipoSistema.sghRisPacs:
            getSystemTypeSourceName = "RIS/PACS"
        Case Else:
            getSystemTypeSourceName = ""
    End Select
End Function

'Obtiene el nombre del proveedor al cual pertenece el sistema
Private Function getSystemProviderName(lSystemProvider As Integer) As String
    Select Case lSystemProvider
        Case sghIntegracionProveedorSistema.sghCarestream:
            getSystemProviderName = "Carestream"
        Case Else:
            getSystemProviderName = ""
    End Select
End Function
