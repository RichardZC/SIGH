VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasIntegracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Reglas Comunes
'        Programado por: Garay M
'        Fecha: Octubre 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_ReglasComunes As New ReglasComunes

Dim lcSql As String

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

Public Function EnviarDatosPacienteRisPacs(oDoPaciente As doPaciente, _
        Optional NuevoPaciente As Boolean = True) As Boolean
        
 Exit Function    'debb
        
    Dim o_IntegrationPatient As SIGHIntegracion.Patient
    Dim oDOPais As New DOPais
    Dim oDODistrito As New DODistrito
    Dim oDOInteoIntegracionSistema As DOInteoIntegracionSistema
    
    Set oDOInteoIntegracionSistema = getProviderIntegration(sghRisPacs)
    
    If Not (oDOInteoIntegracionSistema Is Nothing) Then
        Set o_IntegrationPatient = New SIGHIntegracion.Patient
        o_IntegrationPatient.SystemProvider = oDOInteoIntegracionSistema.IdProveedorSistema  'sghIntegracionProveedorSistema.sghCarestream
        o_IntegrationPatient.SystemTypeSource = sghIntegracionTipoSistema.sghRisPacs
        
        
        Set o_IntegrationPatient.Paciente = oDoPaciente
        
        If oDoPaciente.IdPaisDomicilio > 0 Then
            Set oDOPais = mo_ReglasComunes.PaisSeleccionarPorId(oDoPaciente.IdPaisDomicilio)
        End If
        
        If oDoPaciente.IdDistritoDomicilio > 0 Then
            Set oDODistrito = mo_ReglasComunes.DistritoSeleccionarPorId(oDoPaciente.IdDistritoDomicilio)
        End If
        
        Set o_IntegrationPatient.Pais = oDOPais
        Set o_IntegrationPatient.Distrito = oDODistrito
        
        If NuevoPaciente = True Then
            o_IntegrationPatient.NuevoPaciente
        Else
            o_IntegrationPatient.EditarPaciente
        End If
        
        EnviarDatosPacienteRisPacs = True
    End If
End Function

Public Function UnirDatosPacienteRisPacs(oDoPaciente As doPaciente, oDoPacienteAntiguo As doPaciente) As Boolean
    Dim o_IntegrationPatient As SIGHIntegracion.Patient
    Dim oDOPais As New DOPais
    Dim oDODistrito As New DODistrito
    Dim oDOInteoIntegracionSistema As DOInteoIntegracionSistema
    
    Set oDOInteoIntegracionSistema = getProviderIntegration(sghRisPacs)
    
    If Not (oDOInteoIntegracionSistema Is Nothing) Then
        Set o_IntegrationPatient = New SIGHIntegracion.Patient
        o_IntegrationPatient.SystemProvider = oDOInteoIntegracionSistema.IdProveedorSistema 'sghIntegracionProveedorSistema.sghCarestream
        o_IntegrationPatient.SystemTypeSource = sghIntegracionTipoSistema.sghRisPacs
        
        
        Set o_IntegrationPatient.Paciente = oDoPaciente
        
        If oDoPaciente.IdPaisDomicilio > 0 Then
            Set oDOPais = mo_ReglasComunes.PaisSeleccionarPorId(oDoPaciente.IdPaisDomicilio)
        End If
        
        If oDoPaciente.IdDistritoDomicilio > 0 Then
            Set oDODistrito = mo_ReglasComunes.DistritoSeleccionarPorId(oDoPaciente.IdDistritoDomicilio)
        End If
        
        Set o_IntegrationPatient.Pais = oDOPais
        Set o_IntegrationPatient.Distrito = oDODistrito
        
        o_IntegrationPatient.UnificarPaciente oDoPacienteAntiguo
        
        UnirDatosPacienteRisPacs = True
    End If
End Function


Public Function EnviarDatosMedicoRisPacs(oDoMedico As DOMedico, oDoEmpleado As DOEmpleado, _
            oMedicosEspecialidad As Collection, Optional NuevoMedico As Boolean = True) As Boolean
    Dim o_IntegrationDoctor As SIGHIntegracion.Doctor
    Dim oDOPais As New DOPais
    Dim oDODistrito As New DODistrito
    Dim oDOInteoIntegracionSistema As DOInteoIntegracionSistema
    
    Set oDOInteoIntegracionSistema = getProviderIntegration(sghRisPacs)
    
    If Not (oDOInteoIntegracionSistema Is Nothing) Then
        Set o_IntegrationDoctor = New SIGHIntegracion.Doctor
        o_IntegrationDoctor.SystemProvider = oDOInteoIntegracionSistema.IdProveedorSistema 'sghIntegracionProveedorSistema.sghCarestream
        o_IntegrationDoctor.SystemTypeSource = sghIntegracionTipoSistema.sghRisPacs
        
        
        Set o_IntegrationDoctor.Medico = oDoMedico
        Set o_IntegrationDoctor.Empleado = oDoEmpleado
        Set o_IntegrationDoctor.Especialidades = oMedicosEspecialidad
        
    '    If oDoPaciente.IdPaisDomicilio > 0 Then
    '        Set oDOPais = mo_ReglasComunes.PaisSeleccionarPorId(oDoPaciente.IdPaisDomicilio)
    '    End If
        
    '    If oDoPaciente.IdDistritoDomicilio > 0 Then
    '        Set oDODistrito = mo_ReglasComunes.DistritoSeleccionarPorId(oDoPaciente.IdDistritoDomicilio)
    '    End If
        
    '    Set o_IntegrationPatient.Pais = oDOPais
    '    Set o_IntegrationPatient.Distrito = oDODistrito
    '
        If NuevoMedico = True Then
            o_IntegrationDoctor.NuevoMedico
        Else
            o_IntegrationDoctor.EditarMedico
        End If
        EnviarDatosMedicoRisPacs = True
    End If
End Function


Function ProveedorSistemaSeleccionarTodos() As ADODB.Recordset
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New InteoProveedorSistema
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set ProveedorSistemaSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oConexion = Nothing
miError:
    If Err Then
        MsgBox Err.Description, vbCritical, "Módulo de Intregración"
    End If
End Function


Function TipoSistemaSeleccionarTodos() As ADODB.Recordset
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New InteoTipoSistema
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TipoSistemaSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oConexion = Nothing
miError:
    If Err Then
        MsgBox Err.Description, vbCritical, "Módulo de Intregración"
    End If
End Function

'==============================================================================
'INTEGRACION CON SISTEMAS
'==============================================================================
Function InteoIntegracionSistemaAgregar(oDOInteoIntegracionSistema As DOInteoIntegracionSistema, _
    mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDInteoIntegracionSistema As String) As Boolean
    
    Dim oConexion As New ADODB.Connection
    Dim oInteoIntegracionSistema As New InteoIntegracionSistema
    
    InteoIntegracionSistemaAgregar = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oInteoIntegracionSistema.Conexion = oConexion
    If oInteoIntegracionSistema.Insertar(oDOInteoIntegracionSistema) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOInteoIntegracionSistema.IdUsuarioAuditoria, "A", oDOInteoIntegracionSistema.IdIntegracionSistema, _
                        "InteoIntegracionSistema", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDInteoIntegracionSistema)
        InteoIntegracionSistemaAgregar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oInteoIntegracionSistema.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oInteoIntegracionSistema = Nothing
End Function

Function InteoIntegracionSistemaModificar(oDOInteoIntegracionSistema As DOInteoIntegracionSistema, _
    mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDInteoIntegracionSistema As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oInteoIntegracionSistema As New InteoIntegracionSistema
        
    InteoIntegracionSistemaModificar = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oInteoIntegracionSistema.Conexion = oConexion
    If oInteoIntegracionSistema.modificar(oDOInteoIntegracionSistema) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOInteoIntegracionSistema.IdUsuarioAuditoria, "M", oDOInteoIntegracionSistema.IdIntegracionSistema, _
                            "InteoIntegracionSistema", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDInteoIntegracionSistema)
                            
        InteoIntegracionSistemaModificar = True
        oConexion.CommitTrans
    Else
        oConexion.RollbackTrans
        ms_MensajeError = oInteoIntegracionSistema.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oInteoIntegracionSistema = Nothing
End Function

Function InteoIntegracionSistemaEliminar(oDOInteoIntegracionSistema As DOInteoIntegracionSistema, _
    mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDInteoIntegracionSistema As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oInteoIntegracionSistema As New InteoIntegracionSistema
    
    InteoIntegracionSistemaEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oInteoIntegracionSistema.Conexion = oConexion
    If oInteoIntegracionSistema.Eliminar(oDOInteoIntegracionSistema) Then
        
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOInteoIntegracionSistema.IdUsuarioAuditoria, "E", oDOInteoIntegracionSistema.IdIntegracionSistema, _
                    "InteoIntegracionSistema", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDInteoIntegracionSistema)
                    
        InteoIntegracionSistemaEliminar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oInteoIntegracionSistema.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oInteoIntegracionSistema = Nothing
End Function

Function InteoIntegracionSistemaSeleccionarPorId(IdIntegracionSistema As Integer) As DOInteoIntegracionSistema
    Dim oConexion As New ADODB.Connection
    Dim oInteoIntegracionSistema As New InteoIntegracionSistema
    Dim oDOInteoIntegracionSistema As New DOInteoIntegracionSistema
    
    Set InteoIntegracionSistemaSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oInteoIntegracionSistema.Conexion = oConexion
    oDOInteoIntegracionSistema.IdIntegracionSistema = IdIntegracionSistema
    If oInteoIntegracionSistema.SeleccionarPorId(oDOInteoIntegracionSistema) Then
        Set InteoIntegracionSistemaSeleccionarPorId = oDOInteoIntegracionSistema
    Else
        Set InteoIntegracionSistemaSeleccionarPorId = Nothing
    End If
    oConexion.Close
    ms_MensajeError = oInteoIntegracionSistema.MensajeError
    
    Set oConexion = Nothing
    Set oInteoIntegracionSistema = Nothing
End Function

Function InteoIntegracionSistemaVerificarDuplicado(oDOInteoIntegracionSistema As DOInteoIntegracionSistema) As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oInteoIntegracionSistema As New InteoIntegracionSistema

    Set InteoIntegracionSistemaVerificarDuplicado = Nothing

    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient

    Set oInteoIntegracionSistema.Conexion = oConexion
    Set InteoIntegracionSistemaVerificarDuplicado = oInteoIntegracionSistema.VerificarIntegracionSistema(oDOInteoIntegracionSistema)

    ms_MensajeError = oInteoIntegracionSistema.MensajeError
    oConexion.Close
    Set oConexion = Nothing
    Set oInteoIntegracionSistema = Nothing

End Function

Function InteoIntegracionSistemaFiltrarTodos(oDOInteoIntegracionSistema As DOInteoIntegracionSistema) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oConexion As New ADODB.Connection
    Dim oInteoIntegracionSistema As New InteoIntegracionSistema
    
    Dim ms_MensajeError  As String
    
    Set InteoIntegracionSistemaFiltrarTodos = Nothing
    
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oInteoIntegracionSistema.Conexion = oConexion
    Set InteoIntegracionSistemaFiltrarTodos = oInteoIntegracionSistema.FiltrarTodos(oDOInteoIntegracionSistema)
    
   oConexion.Close
   Set oConexion = Nothing
   Set oInteoIntegracionSistema = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    Exit Function
End Function


Private Function getProviderIntegration(ldTipoSistema As sghIntegracionTipoSistema) As DOInteoIntegracionSistema
On Error GoTo ManejadorDeError
    Dim oConexion As New ADODB.Connection
    Dim oInteoIntegracionSistema As New InteoIntegracionSistema
    Dim oDOInteoIntegracionSistema As New DOInteoIntegracionSistema
    
    Dim ms_MensajeError  As String
    
    oDOInteoIntegracionSistema.IdTipoSistema = ldTipoSistema
    
    Set getProviderIntegration = Nothing
    
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oInteoIntegracionSistema.Conexion = oConexion
    If oInteoIntegracionSistema.SeleccionarProveedorActual(oDOInteoIntegracionSistema) = True Then
        Set getProviderIntegration = oDOInteoIntegracionSistema
    End If
    
    If oInteoIntegracionSistema.MensajeError <> "" Then
        MsgBox oInteoIntegracionSistema.MensajeError, vbInformation, "Integración"
    End If
    
   oConexion.Close
   Set oConexion = Nothing
   Set oInteoIntegracionSistema = Nothing
    Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    Exit Function
End Function
