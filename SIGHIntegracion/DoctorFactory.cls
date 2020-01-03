VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoctorFactory"
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

Dim mo_Medico As DOMedico
Dim mo_Empleado As DOEmpleado
Dim mo_Especialidades As Collection

Dim ml_SystemTypeSource As Integer
Dim ml_SystemProvider As Integer
Dim mo_DOPais As New DOPais
Dim mo_DoDistrito As New DODistrito

'Dim mo_IntegracionCarestream As New IntegracionCarestream

'==================================================================
'Propiedades de objetos de datos
'==================================================================

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

'Expotar un nuevo Medico
Public Function NuevoMedico() As DOMedico
    Dim returnValue As DOMedico
    
    Select Case ml_SystemTypeSource
        Case sghIntegracionTipoSistema.sghRisPacs:
            Set returnValue = NuevoMedicoRisPacs(ml_SystemProvider)
        Case Else:
            MsgBox messageErrorIntegrationSystemType(), vbInformation, "Integración"
            Set returnValue = Nothing
    End Select
    Set NuevoMedico = returnValue
End Function

'Editar un nuevo Medico
Public Function EditarMedico() As DOMedico
    Dim returnValue As DOMedico
    
    Select Case ml_SystemTypeSource
        Case sghIntegracionTipoSistema.sghRisPacs:
            Set returnValue = EditarMedicoRisPacs(ml_SystemProvider)
        Case Else:
            MsgBox messageErrorIntegrationSystemType(), vbInformation, "Integración"
            Set returnValue = Nothing
    End Select
    Set EditarMedico = returnValue
End Function

'Unir dos registros que pertezcan a un solo Medico
Public Function UnificarMedico(oDoMedicoAntiguo As DOMedico) As DOMedico
    Dim returnValue As DOMedico
    
    Select Case ml_SystemTypeSource
        Case sghIntegracionTipoSistema.sghRisPacs:
            Set returnValue = UnificarMedicoRisPacs(ml_SystemProvider, oDoMedicoAntiguo)
        Case Else:
            MsgBox messageErrorIntegrationSystemType(), vbInformation, "Integración"
            Set returnValue = Nothing
    End Select
    Set NuevoMedico = returnValue
End Function

'
Private Function NuevoMedicoRisPacs(lSystemProvider As Integer) As DOMedico
    Dim returnValue As DOMedico
    Select Case lSystemProvider
        Case sghIntegracionProveedorSistema.sghCarestream:
            Dim mo_IntegracionCarestream As New IntegracionCarestream
            
            Set mo_IntegracionCarestream.Medico = mo_Medico
            Set mo_IntegracionCarestream.Distrito = mo_DoDistrito
            Set mo_IntegracionCarestream.Pais = mo_DOPais
            
            Set mo_IntegracionCarestream.Empleado = mo_Empleado
            Set mo_IntegracionCarestream.Especialidades = mo_Especialidades
            
            Set returnValue = mo_IntegracionCarestream.NuevoMedico
        Case Else:
            MsgBox messageErrorIntegrationSystemProvider(), vbInformation, "Integración"
            Set returnValue = Nothing
    End Select
    Set NuevoMedicoRisPacs = returnValue
End Function

Private Function EditarMedicoRisPacs(lSystemProvider As Integer) As DOMedico
    Dim returnValue As DOMedico
    Select Case lSystemProvider
        Case sghIntegracionProveedorSistema.sghCarestream:
            Dim mo_IntegracionCarestream As New IntegracionCarestream
            
            Set mo_IntegracionCarestream.Medico = mo_Medico
            Set mo_IntegracionCarestream.Distrito = mo_DoDistrito
            Set mo_IntegracionCarestream.Pais = mo_DOPais
            
            Set mo_IntegracionCarestream.Empleado = mo_Empleado
            Set mo_IntegracionCarestream.Especialidades = mo_Especialidades
            
            Set returnValue = mo_IntegracionCarestream.EditarMedico
        Case Else:
            MsgBox messageErrorIntegrationSystemProvider(), vbInformation, "Integración"
            Set returnValue = Nothing
    End Select
    Set EditarMedicoRisPacs = returnValue
End Function

'
Private Function UnificarMedicoRisPacs(lSystemProvider As Integer, oDoMedicoAntiguo As DOMedico) As DOMedico
    Dim returnValue As DOMedico
    Select Case lSystemProvider
        Case sghIntegracionProveedorSistema.sghCarestream:
            Dim mo_IntegracionCarestream As New IntegracionCarestream
            
            Set mo_IntegracionCarestream.Medico = mo_Medico
            Set mo_IntegracionCarestream.Distrito = mo_DoDistrito
            Set mo_IntegracionCarestream.Pais = mo_DOPais
            
            Set mo_IntegracionCarestream.Empleado = mo_Empleado
            Set mo_IntegracionCarestream.Especialidades = mo_Especialidades
            
            Set returnValue = mo_IntegracionCarestream.UnificarMedico(oDoMedicoAntiguo)
        Case Else:
            MsgBox messageErrorIntegrationSystemProvider(), vbInformation, "Integración"
            Set returnValue = Nothing
    End Select
    Set UnificarMedicoRisPacs = returnValue
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

'Obtiene el nombre del sistema al cual se le enviaran los datos del Medico
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
