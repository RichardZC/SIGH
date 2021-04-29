VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasAdmision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Reglas de Admisión
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String
Dim mo_reglasComunes As New ReglasComunes
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim lcBuscaParametro As New SIGHDatos.Parametros



Dim mo_rsAtenInteInmunizaciones As ADODB.Recordset
Dim mo_rsAtenInteDesarrollo As ADODB.Recordset
Dim mo_rsAtenInteCrecimiento As ADODB.Recordset
Dim mo_rsAtenInteSuplemento As ADODB.Recordset
Dim mo_rsAtenInteTamizaje As ADODB.Recordset
Dim mo_rsAtenInteDesarrolloCabecera As ADODB.Recordset
Dim mo_rsAtenInteCrecimientoCabecera As ADODB.Recordset

Dim mo_rsPlanInmunizaciones As ADODB.Recordset
Dim mo_rsPlanDesarrollo As ADODB.Recordset
Dim mo_rsPlanCrecimiento As ADODB.Recordset
Dim mo_rsPlanSuplemento As ADODB.Recordset
Dim mo_rsPlanTamizaje As ADODB.Recordset
Dim mo_idOrdenPago As Long


Property Get idOrdenPago() As Long
   idOrdenPago = mo_idOrdenPago
End Property

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

Property Set setRsAtenInteInmunizaciones(oRs As ADODB.Recordset)
   Set mo_rsAtenInteInmunizaciones = oRs
End Property
Property Set setRsAtenInteDesarrollo(oRs As ADODB.Recordset)
   Set mo_rsAtenInteDesarrollo = oRs
End Property
Property Set setRsAtenInteCrecimiento(oRs As ADODB.Recordset)
   Set mo_rsAtenInteCrecimiento = oRs
End Property
Property Set setRsAtenInteSuplemento(oRs As ADODB.Recordset)
   Set mo_rsAtenInteSuplemento = oRs
End Property
Property Set setRsAtenInteTamizaje(oRs As ADODB.Recordset)
   Set mo_rsAtenInteTamizaje = oRs
End Property

Property Set setRsPlanInmunizaciones(oRs As ADODB.Recordset)
   Set mo_rsPlanInmunizaciones = oRs
End Property
Property Set setRsPlanDesarrollo(oRs As ADODB.Recordset)
   Set mo_rsPlanDesarrollo = oRs
End Property
Property Set setRsPlanCrecimiento(oRs As ADODB.Recordset)
   Set mo_rsPlanCrecimiento = oRs
End Property
Property Set setRsPlanSuplemento(oRs As ADODB.Recordset)
   Set mo_rsPlanSuplemento = oRs
End Property
Property Set setRsPlanTamizaje(oRs As ADODB.Recordset)
   Set mo_rsPlanTamizaje = oRs
End Property

Property Set setRsAtenInteDesarrolloCabecera(oRs As ADODB.Recordset)
   Set mo_rsAtenInteDesarrolloCabecera = oRs
End Property

Property Set setRsAtenInteCrecimientoCabecera(oRs As ADODB.Recordset)
   Set mo_rsAtenInteCrecimientoCabecera = oRs
End Property

'------------------------------------------------------------------------------------------------
'                            REGLAS DE NEGOCIOS DEL MODULO DE PACIENTES
'------------------------------------------------------------------------------------------------

Function PacientesAgregarPacienteEHistoriaClinica(oDoPaciente As DOPaciente, oDoHistoria As DOHistoriaClinica, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNpaciente As String, oDoSunasaPacientesHistoricos As DoSunasaPacientesHistoricos, _
                             Optional oDoPacienteDatosAdd As DoPacienteDatosAdd = Nothing) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oHistoria As New HistoriasClinicas
Dim oDoHistoriaClinica As New DOHistoriaClinica
Dim oSunasaPacientesHistoricos As New SunasaPacientesHistoricos
Dim bProcesoOK As Boolean
    
    PacientesAgregarPacienteEHistoriaClinica = False
    bProcesoOK = True
    
    oConexion.Open sighentidades.CadenaConexion
    Set oPaciente.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oSunasaPacientesHistoricos.Conexion = oConexion
    oConexion.BeginTrans
    
    If oDoHistoria.IdTipoNumeracion <> sghHistoriaDefinitivaManual Then
        oDoHistoria.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoPaciente.IdTipoNumeracion)
'        If lcBuscaParametro.SeleccionaFilaParametro(351) = "S" Then   'Requerimiento INSNSB, Actualizado por FCV 30032015
'            oDoHistoria.NroHistoriaClinica = oDoPaciente.NroDocumento
'        Else
'            oDoHistoria.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoPaciente.IdTipoNumeracion)
'        End If
    End If
    
    If oDoHistoria.NroHistoriaClinica = 0 Then
        If oDoHistoria.IdTipoNumeracion = sghHistoriaDefinitivaReciclada Then
            Trace "La opción de Historias Recicladas aun no esta implementada", "PacientesAgregarPacienteEHistoriaClinica"
            bProcesoOK = False: GoTo Terminar
        End If
    End If
    
    'Agregar paciente
    oDoPaciente.IdTipoNumeracion = oDoHistoria.IdTipoNumeracion
    oDoPaciente.NroHistoriaClinica = oDoHistoria.NroHistoriaClinica
    
    If Not oPaciente.Insertar(oDoPaciente) Then
        bProcesoOK = False: GoTo Terminar
    Else
        If Not (oDoPacienteDatosAdd Is Nothing) Then
            oDoPacienteDatosAdd.idPaciente = oDoPaciente.idPaciente
        End If
        Call PacientesDatosAdicionalesPersonalesAgregar(oDoPacienteDatosAdd, oConexion)
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoPaciente.IdUsuarioAuditoria, "A", oDoPaciente.idPaciente, "Pacientes", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNpaciente)      'ListBarItems.idListItem
    '
    oDoHistoria.idPaciente = oDoPaciente.idPaciente
    oDoHistoria.NroHistoriaClinica = oDoPaciente.NroHistoriaClinica
    Select Case oDoHistoria.IdTipoNumeracion
    Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
        If Not oHistoria.Insertar(oDoHistoria) Then
            bProcesoOK = False: GoTo Terminar
        End If
    End Select
    '
    If oDoSunasaPacientesHistoricos.YaNoTieneSeguro = False Then
        oDoSunasaPacientesHistoricos.idPaciente = oDoPaciente.idPaciente
        If Not oSunasaPacientesHistoricos.Insertar(oDoSunasaPacientesHistoricos) Then
            bProcesoOK = False: GoTo Terminar
        End If
    End If
   
Terminar:
    If bProcesoOK Then
        PacientesAgregarPacienteEHistoriaClinica = True
        oConexion.CommitTrans
    Else

        PacientesAgregarPacienteEHistoriaClinica = False
        oConexion.RollbackTrans
    End If
            
    oConexion.Close
    Set oConexion = Nothing
    Set oPaciente = Nothing
    Set oHistoria = Nothing
    Set oDoHistoriaClinica = Nothing
    Set oSunasaPacientesHistoricos = Nothing

    
End Function

Function PacientesModificarYActualizarHistoriaClinicaDefinitiva(oDoPaciente As DOPaciente, oDoHistoria As DOHistoriaClinica, lIdTipoGenHistoriaClinicaAnterior As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNpaciente As String, oDoSunasaPacientesHistoricos As DoSunasaPacientesHistoricos, _
                             Optional oDoPacienteDatosAdd As DoPacienteDatosAdd = Nothing) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oHistoria As New HistoriasClinicas
Dim oSunasaPacientesHistoricos As New SunasaPacientesHistoricos
Dim mo_ReglasArchivoClinico As New ReglasArchivoClinico                     'debb-18/05/2016
Dim bProcesoOK As Boolean

    PacientesModificarYActualizarHistoriaClinicaDefinitiva = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oPaciente.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oSunasaPacientesHistoricos.Conexion = oConexion
    
    oConexion.BeginTrans
    bProcesoOK = True
    'Agregar historia clinica (Si es tenia historia temporal y ahora tiene historia definitiva)
    Select Case lIdTipoGenHistoriaClinicaAnterior
    Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
        If lIdTipoGenHistoriaClinicaAnterior = sghHistoriaDefinitivaManual Then
            If Not oHistoria.modificar(oDoHistoria) Then
            End If
        End If
    Case Else
        
        'Si era temporal y ahora es definitiva => Genera y agrega la historia
        Select Case oDoHistoria.IdTipoNumeracion
        Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
            
            If oDoHistoria.IdTipoNumeracion <> sghHistoriaDefinitivaManual Then
                oDoHistoria.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoPaciente.IdTipoNumeracion)
            End If
            
            If oDoHistoria.NroHistoriaClinica = 0 Then
                If oDoHistoria.IdTipoNumeracion = sghHistoriaDefinitivaReciclada Then
                    Trace "No se pudo generar el nro de de historia clínica reciclada", "PacientesModificarYActualizarHistoriaClinicaDefinitiva"
                    bProcesoOK = False: GoTo Terminar
                End If
                If oDoHistoria.IdTipoNumeracion = sghHistoriaDefinitivaAutomatica Then
                    Trace "No se pudo generar el nro de de historia clínica automatica, verifique el generador de números ", "PacientesModificarYActualizarHistoriaClinicaDefinitiva"
                    bProcesoOK = False: GoTo Terminar
                End If
            End If
            
            If Not oHistoria.Insertar(oDoHistoria) Then
                ms_MensajeError = oHistoria.MensajeError
                bProcesoOK = False: GoTo Terminar
            End If
            
            oDoPaciente.NroHistoriaClinica = oDoHistoria.NroHistoriaClinica
            oDoPaciente.IdTipoNumeracion = oDoHistoria.IdTipoNumeracion
        
        End Select
    End Select
    
    'Modificar paciente
    If Not oPaciente.modificar(oDoPaciente, True) Then
        ms_MensajeError = oPaciente.MensajeError
        bProcesoOK = False: GoTo Terminar
    Else
        If Not (oDoPacienteDatosAdd Is Nothing) Then
            oDoPacienteDatosAdd.idPaciente = oDoPaciente.idPaciente
        End If
        Call PacientesDatosAdicionalesPersonalesAgregar(oDoPacienteDatosAdd, oConexion)
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoPaciente.IdUsuarioAuditoria, "M", oDoPaciente.idPaciente, "Pacientes", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNpaciente)      'ListBarItems.idListItem
    '
    If oDoSunasaPacientesHistoricos.NuevoSeguro = True Then
        If Not oSunasaPacientesHistoricos.Insertar(oDoSunasaPacientesHistoricos) Then
            bProcesoOK = False: GoTo Terminar
        End If
    Else
        If Not oSunasaPacientesHistoricos.modificar(oDoSunasaPacientesHistoricos) Then
            bProcesoOK = False: GoTo Terminar
        End If
    End If
    
    mo_ReglasArchivoClinico.ActualizaEDAD_en_movimientos oConexion, oDoPaciente.idPaciente           'debb-18/05/2016
    mo_ReglasArchivoClinico.HistoriasClinicasActualizaFmovimiento oConexion, oDoPaciente.idPaciente
    
Terminar:
    If bProcesoOK Then
        PacientesModificarYActualizarHistoriaClinicaDefinitiva = True
        oConexion.CommitTrans
        
    Else
        PacientesModificarYActualizarHistoriaClinicaDefinitiva = False
        oConexion.RollbackTrans
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oPaciente = Nothing
    Set oHistoria = Nothing
    Set oSunasaPacientesHistoricos = Nothing
    Set mo_ReglasArchivoClinico = Nothing               'debb-18/05/2016
End Function
Function PacientesAgregar(oDoPaciente As DOPaciente, _
        Optional oDoPacienteDatosAdd As DoPacienteDatosAdd = Nothing) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
    
    PacientesAgregar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oPaciente.Conexion = oConexion
    If oPaciente.Insertar(oDoPaciente) Then
        If Not (oDoPacienteDatosAdd Is Nothing) Then
            oDoPacienteDatosAdd.idPaciente = oDoPaciente.idPaciente
        End If
        Call PacientesDatosAdicionalesPersonalesAgregar(oDoPacienteDatosAdd, oConexion)
        PacientesAgregar = True
        'mgaray201411f
'        Dim o_ReglasIntegracion As New ReglasIntegracion
'        Call o_ReglasIntegracion.EnviarDatosPacienteRisPacs(oDoPaciente)
    Else
        ms_MensajeError = oPaciente.MensajeError
    End If
    oConexion.Close
    Set oPaciente = Nothing
    Set oConexion = Nothing
End Function

Function PacientesModificar(oDoPaciente As DOPaciente, lbVerificaPacientesRepetidosAntesDeGrabarDatos As Boolean, _
        Optional oDoPacienteDatosAdd As DoPacienteDatosAdd = Nothing) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
    
    PacientesModificar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oPaciente.Conexion = oConexion
    If oPaciente.modificar(oDoPaciente, lbVerificaPacientesRepetidosAntesDeGrabarDatos) Then
        If Not (oDoPacienteDatosAdd Is Nothing) Then
            oDoPacienteDatosAdd.idPaciente = oDoPaciente.idPaciente
        End If
        Call PacientesDatosAdicionalesPersonalesAgregar(oDoPacienteDatosAdd, oConexion)
        PacientesModificar = True
        'mgaray201411f
'        Dim o_ReglasIntegracion As New ReglasIntegracion
'        Call o_ReglasIntegracion.EnviarDatosPacienteRisPacs(oDoPaciente)
    Else
        ms_MensajeError = oPaciente.MensajeError
    End If
    oConexion.Close
    Set oPaciente = Nothing
    Set oConexion = Nothing
End Function

Function PacientesEliminar(oDoPaciente As DOPaciente, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNpaciente As String, oDoSunasaPacientesHistoricos As DoSunasaPacientesHistoricos) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oHistoria As New HistoriasClinicas
Dim oDoHistoria As New DOHistoriaClinica
Dim oSunasaPacientesHistoricos As New SunasaPacientesHistoricos
Dim bProcesoOK As Boolean

    PacientesEliminar = False
    bProcesoOK = True
    
    oConexion.Open sighentidades.CadenaConexion
    Set oPaciente.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oSunasaPacientesHistoricos.Conexion = oConexion
    
    oDoHistoria.NroHistoriaClinica = oDoPaciente.NroHistoriaClinica
    oDoHistoria.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
    
    oConexion.BeginTrans
    
    
    
    'Eliminar historia
    Select Case oDoPaciente.IdTipoNumeracion
    Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
        If Not oHistoria.Eliminar(oDoHistoria) Then
            ms_MensajeError = oHistoria.MensajeError
            bProcesoOK = False: GoTo Terminar
        End If
    End Select
    '
    If Not oSunasaPacientesHistoricos.Eliminar(oDoSunasaPacientesHistoricos) Then
        bProcesoOK = False: GoTo Terminar
    End If
    'Eliminar paciente
    If Not oPaciente.Eliminar(oDoPaciente) Then
        ms_MensajeError = oPaciente.MensajeError
        bProcesoOK = False: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoPaciente.IdUsuarioAuditoria, "E", oDoPaciente.idPaciente, "Pacientes", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNpaciente)      'ListBarItems.idListItem
Terminar:
    If bProcesoOK Then
        PacientesEliminar = True
        oConexion.CommitTrans
    Else
        PacientesEliminar = False
        oConexion.RollbackTrans
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oPaciente = Nothing
    Set oHistoria = Nothing
    Set oSunasaPacientesHistoricos = Nothing
    
End Function

Function PacientesSeleccionarPorId(idPaciente As Long, oConexion As Connection) As DOPaciente
    Dim oPaciente As New Pacientes
    Dim oDoPaciente As New DOPaciente
    Set PacientesSeleccionarPorId = Nothing
    Set oPaciente.Conexion = oConexion
    oDoPaciente.idPaciente = idPaciente
    If oPaciente.SeleccionarPorId(oDoPaciente) Then
        Set PacientesSeleccionarPorId = oDoPaciente
    End If
    ms_MensajeError = oPaciente.MensajeError
    Set oPaciente = Nothing
    Set oDoPaciente = Nothing
End Function

Function RetornaPacientesSeleccionarPorId(idPaciente As Long) As DOPaciente
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    
    oConexion.ConnectionTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    Set RetornaPacientesSeleccionarPorId = PacientesSeleccionarPorId(idPaciente, oConexion)
    
    oConexion.Close
miError:
    If Err Then
        ms_MensajeError = Err.Description
    End If
    Set oConexion = Nothing
End Function


Function PacientesSeleccionarPorHistoriaClinicaDefinitiva(NroHistoriaClinica As Long) As DOPaciente
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oDoPaciente As New DOPaciente
    
    Set PacientesSeleccionarPorHistoriaClinicaDefinitiva = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oPaciente.Conexion = oConexion
    oDoPaciente.NroHistoriaClinica = NroHistoriaClinica
    
    If oPaciente.SeleccionarPorHistoriaClinicaDefinitiva(oDoPaciente) Then
        Set PacientesSeleccionarPorHistoriaClinicaDefinitiva = oDoPaciente
    End If
    oConexion.Close
    ms_MensajeError = oPaciente.MensajeError
    Set oConexion = Nothing
    Set oPaciente = Nothing

End Function
Function PacientesObtenerPacientePorHistoriaClinica(lIdNroHistoriaClinica As Long, lIdTipoNumeracion As Long) As DOPaciente
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oDoPaciente As New DOPaciente
    
    Set PacientesObtenerPacientePorHistoriaClinica = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oPaciente.Conexion = oConexion
    oDoPaciente.idPaciente = oPaciente.ObtenerIdPacientePorHistoriaClinica(lIdNroHistoriaClinica, lIdTipoNumeracion)
    If oPaciente.SeleccionarPorId(oDoPaciente) Then
        Set PacientesObtenerPacientePorHistoriaClinica = oDoPaciente
    End If
    
    oConexion.Close
    ms_MensajeError = oPaciente.MensajeError
    Set oConexion = Nothing
    Set oPaciente = Nothing

End Function

Function PacientesFiltrar(ByVal oPaciente As DOPaciente, lbUsaApellidoMaternoVacio As Boolean, _
                 lbUsaApellidoPaternoVacio As Boolean, lcApellidoVacio As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Pacientes
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    oConexion.CommandTimeout = 150
    Set PacientesFiltrar = oTabla.Filtrar(oPaciente, lbUsaApellidoMaternoVacio, _
                                          lbUsaApellidoPaternoVacio, lcApellidoVacio)
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function
Function PacientesFiltrarHistoriasParaAdmision(ByVal lNroHistoriaClinica As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Pacientes
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set PacientesFiltrarHistoriasParaAdmision = oTabla.FiltrarHistoriasParaAdmision(lNroHistoriaClinica)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


Function PacientesFiltrarConHistoriasTemporales(ByVal oPaciente As DOPaciente) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Pacientes
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set PacientesFiltrarConHistoriasTemporales = oTabla.FiltrarConHistoriasTemporales(oPaciente)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function
Function PacientesFiltrarConHistoriasDefinitivas(ByVal oPaciente As DOPaciente, lcSinApellido As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Pacientes
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set PacientesFiltrarConHistoriasDefinitivas = oTabla.FiltrarConHistoriasDefinitivas(oPaciente, lcSinApellido)
    ms_MensajeError = oTabla.MensajeError
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function PacientesObtenerConElMismoNombre(ByVal oPaciente As DOPaciente) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Pacientes
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set PacientesObtenerConElMismoNombre = oTabla.ObtenerConElMismoNombre(oPaciente)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function PacientesObtenerConElAutogenerado(ByVal oPaciente As DOPaciente) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Pacientes
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set PacientesObtenerConElAutogenerado = oTabla.ObtenerConElMismoAutogenerado(oPaciente)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function PacientesObtenerConElMismoNroHistoriaDefinitiva(ByVal oPaciente As DOPaciente) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Pacientes
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set PacientesObtenerConElMismoNroHistoriaDefinitiva = oTabla.ObtenerConLaMismaHistoriaDefinitiva(oPaciente)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function PacienteTieneCita(FechaIngreso As Date, lIdServicio As Long, lidPaciente As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Pacientes
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set PacienteTieneCita = oTabla.TieneCita(FechaIngreso, lIdServicio, lidPaciente)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing


End Function

Function PacienteCrearNroAutogenerado(oPaciente As DOPaciente)
Dim P1 As String    'Primer digito del apellido paterno
Dim P4 As String    'Cuarto Digito del apellido paterno
Dim M1 As String    'Primer digito del apellido materno
Dim M4 As String    'Cuarto digito del apellido materno
Dim N11 As String   'Primer digito del primer nombre
Dim N41 As String   'Cuarto digito del primer materno
Dim N12 As String   'Primer digito del Ultimo materno
Dim N42 As String   'Cuarto digito del Ultimo materno
Dim D As String     'Digito de verificacion
Dim DD As String
Dim MM As String
Dim AAA As String
Dim sTemp  As String

        DD = Left(oPaciente.FechaNacimiento, 2)
        MM = Mid(oPaciente.FechaNacimiento, 4, 2)
        AAA = Mid(oPaciente.FechaNacimiento, 8, 3)
        DevuelvePrimeryCuartoCaracter oPaciente.ApellidoPaterno, P1, P4
        DevuelvePrimeryCuartoCaracter oPaciente.ApellidoMaterno, M1, M4
        DevuelvePrimeryCuartoCaracter oPaciente.PrimerNombre, N11, N41
        DevuelvePrimeryCuartoCaracter oPaciente.SegundoNombre, N12, N42
        sTemp = AAA + MM + DD & oPaciente.idTipoSexo & P1 + P4 + M1 + M4 + N11 + N41 + N12 + N42
        PacienteCrearNroAutogenerado = sTemp & Modulo10(sTemp) & Format(oPaciente.FechaNacimiento, sighentidades.DevuelveHoraSoloFormato_HM)
        
End Function
Function Modulo10(sValor As String) As Integer
Dim sTemp As String
Dim I As Integer
Dim k As Integer
Dim iTotal As Integer

    sTemp = ""
    
    For I = 1 To Len(sValor)
        If IsNumeric(Mid(sValor, I, 1)) Then
            sTemp = sTemp + Mid(sValor, I, 1)
        Else
            sTemp = sTemp + DevuelveValorEnNumeros(Mid(sValor, I, 1))
        End If
    Next I

    'Acumula total de digitos
    iTotal = 0
    For I = 1 To Len(sTemp)
        If I Mod 2 <> 0 Then
            k = CInt(Mid(sTemp, I, 1)) * 2
            iTotal = iTotal + (k - (k Mod 10)) / 10 + (k Mod 10)
        Else
            iTotal = iTotal + CInt(Mid(sTemp, I, 1))
        End If
    Next I

    If (iTotal Mod 10) = 0 Then
        Modulo10 = 0
    Else
        Modulo10 = 10 - (iTotal Mod 10)
    End If



End Function
Function DevuelveValorEnNumeros(sCaracter As String) As String

    Select Case sCaracter
    Case "A" To "N"
        DevuelveValorEnNumeros = Asc(sCaracter) - 55
    Case "Ñ"
        DevuelveValorEnNumeros = 24
    Case "O" To "Z"
        DevuelveValorEnNumeros = Asc(sCaracter) - 54
    End Select

End Function


Sub DevuelvePrimeryCuartoCaracter(sPalabra As String, C1 As String, C2 As String)
Dim sTemp As String
        If sPalabra <> "" Then
            sTemp = ObtenerUltimaPalabra(EliminarConjunciones(sPalabra))
            C1 = UCase(Left(sTemp, 1))
            C2 = UCase(DevuelveCuartoCaracter(sTemp))
        Else
            C1 = "X"
            C2 = "X"
        End If
End Sub
Function DevuelveCuartoCaracter(sPalabra) As String
    If Len(sPalabra) <= 4 Then
        DevuelveCuartoCaracter = Right(sPalabra, 1)
    Else
        DevuelveCuartoCaracter = Mid(sPalabra, 4, 1)
    End If
End Function
Function ObtenerPrimeraPalabra(sTexto As String) As String
Dim p As String
Dim sTemp As String

    sTemp = Trim(sTexto)
    
    p = InStr(sTemp, " ")
    If p > 0 Then
        ObtenerPrimeraPalabra = Left(sTemp, p - 1)
    Else
        ObtenerPrimeraPalabra = sTemp
    End If
End Function
Function ObtenerUltimaPalabra(sTexto As String) As String
Dim p As String
Dim iUltBlanco As Integer
Dim sTemp As String


    sTemp = Trim(sTexto)

    p = InStr(sTemp, " ")
    iUltBlanco = 0
    Do While p > 0
        iUltBlanco = p
        p = InStr(p + 1, sTemp, " ")
    Loop
    If iUltBlanco > 0 Then
        ObtenerUltimaPalabra = Mid(sTemp, iUltBlanco + 1)
    Else
        ObtenerUltimaPalabra = sTemp
    End If
End Function
Function ObtenerUltimoNombre(sTexto As String) As String
Dim p As String
Dim iUltBlanco As Integer
Dim sTemp As String


    sTemp = Trim(sTexto)

    p = InStr(sTemp, " ")
    iUltBlanco = 0
    Do While p > 0
        iUltBlanco = p
        p = InStr(p + 1, sTemp, " ")
    Loop
    If iUltBlanco > 0 Then
        ObtenerUltimoNombre = Mid(sTemp, iUltBlanco + 1)
    Else
        ObtenerUltimoNombre = ""
    End If
End Function

Function EliminarConjunciones(sPalabra As String)
Dim sTemp As String

        sTemp = sighentidades.ReemplazarCadena(sPalabra, " DE ", " ")
        sTemp = sighentidades.ReemplazarCadena(sTemp, " DEL ", " ")
        sTemp = sighentidades.ReemplazarCadena(sTemp, " EL ", " ")
        sTemp = sighentidades.ReemplazarCadena(sTemp, " LA ", " ")
        sTemp = sighentidades.ReemplazarCadena(sTemp, " LOS ", " ")
        sTemp = sighentidades.ReemplazarCadena(sTemp, " LAS ", " ")

        EliminarConjunciones = sTemp

End Function

'------------------------------------------------------------------------------------------------
'                            REGLAS DE NEGOCIOS DEL MODULO DE CITAS
'------------------------------------------------------------------------------------------------
Sub InicializarPrestamoAutomaticoCE(oDOCita As doCita, oDoPaciente As DOPaciente, oDOHistoriaSolicitada As DOHistoriaSolicitada, bPacienteNuevo As Boolean, lnIdAtencion As Long)
        
        oDOHistoriaSolicitada.idPaciente = oDoPaciente.idPaciente
        oDOHistoriaSolicitada.fechaSolicitud = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY)
        oDOHistoriaSolicitada.horaSolicitud = Format(Now, "hh:mm")
        
        oDOHistoriaSolicitada.fechaRequerida = oDOCita.fecha
        oDOHistoriaSolicitada.horaRequerida = oDOCita.HoraInicio
        
        Select Case oDoPaciente.IdTipoNumeracion
        Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
            If bPacienteNuevo Then
                oDOHistoriaSolicitada.Observacion = "CITA CE: " + "Paciente nuevo - preparar folder - Fecha: " & oDOCita.fecha & " (" & oDOCita.HoraInicio & " - " & oDOCita.horaFin + ")"
            Else
                oDOHistoriaSolicitada.Observacion = "CITA CE: " + "Ubicar historia clinica - Fecha: " & oDOCita.fecha & " (" & oDOCita.HoraInicio & " - " & oDOCita.horaFin + ")"
            End If
        Case Else
            If bPacienteNuevo Then
                oDOHistoriaSolicitada.Observacion = "CITA CE: " + "Paciente nuevo - Crear Nro de Historia y preparar folder - Fecha: " & oDOCita.fecha & " (" & oDOCita.HoraInicio & " - " & oDOCita.horaFin + ")"
            Else
                oDOHistoriaSolicitada.Observacion = "CITA CE: " + "Crear Nro de Historia y preparar folder - Fecha: " & oDOCita.fecha & " (" & oDOCita.HoraInicio & " - " & oDOCita.horaFin + ")"
            End If
        End Select
        
        oDOHistoriaSolicitada.IdMotivo = 1
        oDOHistoriaSolicitada.IdServicio = oDOCita.IdServicio
        oDOHistoriaSolicitada.IdUsuarioAuditoria = oDOCita.IdUsuarioAuditoria
        oDOHistoriaSolicitada.IdMovimiento = 0
        oDOHistoriaSolicitada.idAtencion = lnIdAtencion
        
End Sub
Sub InicializarPrestamoAutomaticoEmerg(oDoAtencion As DOAtencion, oDoPaciente As DOPaciente, oDOHistoriaSolicitada As DOHistoriaSolicitada, bPacienteNuevo As Boolean)
        
        oDOHistoriaSolicitada.idPaciente = oDoPaciente.idPaciente
        oDOHistoriaSolicitada.fechaSolicitud = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY)
        oDOHistoriaSolicitada.horaSolicitud = Format(Now, "hh:mm")
        
        oDOHistoriaSolicitada.fechaRequerida = oDoAtencion.FechaIngreso
        oDOHistoriaSolicitada.horaRequerida = oDoAtencion.HoraIngreso
        
        oDOHistoriaSolicitada.IdMotivo = 1
        oDOHistoriaSolicitada.IdServicio = oDoAtencion.idServicioIngreso
        oDOHistoriaSolicitada.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
           
           oDOHistoriaSolicitada.Observacion = "ADMISION EMERGENCIA: " + IIf(bPacienteNuevo, "PACIENTE NUEVO", "") + " Fecha Ingreso: " & oDoAtencion.FechaIngreso
        
        oDOHistoriaSolicitada.idAtencion = oDoAtencion.idAtencion
        
End Sub
Sub InicializarPrestamoAutomaticoHosp(oDoAtencion As DOAtencion, oDoPaciente As DOPaciente, oDOHistoriaSolicitada As DOHistoriaSolicitada, bPacienteNuevo As Boolean)
        
        oDOHistoriaSolicitada.idPaciente = oDoPaciente.idPaciente
        oDOHistoriaSolicitada.fechaSolicitud = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY)
        oDOHistoriaSolicitada.horaSolicitud = Format(Now, "hh:mm")
        
        oDOHistoriaSolicitada.fechaRequerida = oDoAtencion.FechaIngreso
        oDOHistoriaSolicitada.horaRequerida = oDoAtencion.HoraIngreso
        
        oDOHistoriaSolicitada.IdMotivo = 1
        oDOHistoriaSolicitada.IdServicio = oDoAtencion.idServicioIngreso
        oDOHistoriaSolicitada.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
        
        oDOHistoriaSolicitada.Observacion = "ADMISION HOSP: " + IIf(bPacienteNuevo, "PACIENTE NUEVO", "") + " Fecha Ingreso: " & oDoAtencion.FechaIngreso
        oDOHistoriaSolicitada.idAtencion = oDoAtencion.idAtencion
        
        Dim oBuscaHistoriaSolicitada As New SIGHNegocios.ReglasArchivoClinico
        Dim oRsTmp As New ADODB.Recordset
        Set oRsTmp = oBuscaHistoriaSolicitada.HistoriasSolicitadasSeleccionarPorIdAtencion(oDoAtencion.idAtencion)
        If oRsTmp.RecordCount > 0 Then
            oDOHistoriaSolicitada.IdHistoriaSolicitada = oRsTmp.Fields!IdHistoriaSolicitada
        End If
        Set oRsTmp = Nothing
        Set oBuscaHistoriaSolicitada = Nothing
End Sub

Function CitasModificar(oDOCita As doCita) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCita As New Citas

    CitasModificar = False

    oConexion.Open sighentidades.CadenaConexion
    Set oCita.Conexion = oConexion
    If oCita.modificar(oDOCita) Then
        CitasModificar = True
    Else
        ms_MensajeError = oCita.MensajeError
    End If
    oConexion.Close

End Function


Function CitasSeleccionarPorId(IdCita As Long, oDOCita As doCita, oDoPaciente As DOPaciente, oConexion As Connection) As Boolean
Dim oCita As New Citas
Dim oPaciente As New Pacientes
    CitasSeleccionarPorId = False
    Set oCita.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    oDOCita.IdCita = IdCita
    If oCita.SeleccionarPorId(oDOCita) Then
        oDoPaciente.idPaciente = oDOCita.idPaciente
        If oPaciente.SeleccionarPorId(oDoPaciente) Then
            CitasSeleccionarPorId = True
        End If
    End If
    ms_MensajeError = oCita.MensajeError
    Set oCita = Nothing
    Set oPaciente = Nothing
End Function
Function CitasSeleccionarPorPaciente(lidPaciente As Long, daFechaInicial As Date) As Recordset
Dim oConexion As New ADODB.Connection
Dim oCita As New Citas
Dim oPaciente As New Pacientes

    Set CitasSeleccionarPorPaciente = Nothing

    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient

    Set oCita.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    Set CitasSeleccionarPorPaciente = oCita.SeleccionarPorPacienteYFecha(lidPaciente, daFechaInicial)

    oConexion.Close

    ms_MensajeError = oCita.MensajeError
    Set oConexion = Nothing
    Set oCita = Nothing
    Set oPaciente = Nothing

End Function

'------------------------------------------------------------------------------------------------
'                            AdmisionCEAgregar
'------------------------------------------------------------------------------------------------
Function AdmisionCEAgregar(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, oDoPaciente As DOPaciente, _
                           oDOCita As doCita, oDoHistoria As DOHistoriaClinica, lIdTipoNumeracionAnterior As sghTipoNumeracionDeNroHistoria, _
                           oDiagnosticos As Collection, oProcedimientos As Collection, oExamenes As Collection, _
                           lIdHistoriaAnterior As Long, mo_FacturacionServicios As Collection, DOFacturacionPaquetesUltimo As DOFacturacionPaquetes, _
                           lbYaSeTransfirioHCdeUnServicioAotro As Boolean, lnIdTablaLISTBARITEMS As Long, _
                           lcNombrePc As String, lcObservacionesAuditoria As String, oDoSunasaPacientesHistoricos As DoSunasaPacientesHistoricos, _
                           mo_DoAtencionDatosAdicionales As DoAtencionDatosAdicionales, _
                             Optional oDoPacienteDatosAdd As DoPacienteDatosAdd = Nothing) As Boolean
      Dim oConexion As New ADODB.Connection
      Dim oPaciente As New Pacientes
      Dim oCuentaAtencion As New CuentasAtencion
      Dim oAtencion As New Atenciones
      Dim oDOHistoriaSolicitada As New DOHistoriaSolicitada
      Dim oHistoriaSolicitada As New HistoriasSolicitadas
      Dim oCita As New Citas
      Dim oHistoria As New HistoriasClinicas
      Dim oAtencionesDx As New AtencionesDiagnosticos
      Dim oCatalogoServicios As New CatalogoServicios
      Dim oFacturacionServicios As New FacturacionServicios
      Dim oFactOrdenesServicio As New FactOrdenServicio
      Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
      Dim oFacturacionServicioPagos As New FacturacionServicioPagos
      Dim oFacturacionServicioFinan As New FacturacionServicioFin
      Dim oFactServicioPagos As New FactOrdenServicioPagos
      Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
      Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
      Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
      Dim oDOFacturacionServicio As New DOFacturacionServicios
      Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
      Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
      Dim oDOMovimientoHistoriaClinica As New DOMovimientoHistoriaClinica
      Dim oMovimientosHistoriaClinica As New MovimientosHistoriaClinica
      Dim oSunasaPacientesHistoricos As New SunasaPacientesHistoricos
      Dim oAtencionesDatosAdicionales As New AtencionesDatosAdicionales
      Dim oDOProgramacionMedica As New DOProgramacionMedica, oProgramacionMedica As New ProgramacionMedica
      Dim ml_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
      Dim oBuscaCodigoNombre As New SIGHNegocios.ReglasComunes
      Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
      
     
      
     
      
      Dim lcMinutosCita As String
      Dim lnImporteTotal As Double
      Dim lnPrecioUnitario As Double
      Dim lcCodigoNombre As String
      Dim lcSql As String
      Dim bProcesoOK As Boolean
      Dim bPacienteNuevo As Boolean
      Dim lIdCuentaAtencion As Long
      Dim rs As New Recordset
      Dim oRsTmp1 As New Recordset
      Dim oRsTmp2 As New Recordset
      Dim lnMontoCPTfijoEnsoles As Double
          
10        Set rs = mo_ReglasFacturacion.CuentaAtencionSeleccionarUltimaPorIdPaciente(oDoPaciente.idPaciente)
20        If rs.RecordCount = 0 Then
30            lIdCuentaAtencion = 0
40        Else
50            lIdCuentaAtencion = rs!idCuentaAtencion
60        End If
70        rs.Close
          
80        oConexion.CommandTimeout = 300
90        oConexion.CursorLocation = adUseClient
100       oConexion.Open sighentidades.CadenaConexion
110       oConexion.BeginTrans
          
120       Set oCuentaAtencion.Conexion = oConexion
130       Set oAtencion.Conexion = oConexion
140       Set oPaciente.Conexion = oConexion
150       Set oHistoriaSolicitada.Conexion = oConexion
160       Set oCita.Conexion = oConexion
170       Set oHistoria.Conexion = oConexion
180       Set oAtencionesDx.Conexion = oConexion
190       Set oFactOrdenesServicio.Conexion = oConexion
200       Set oCatalogoServicios.Conexion = oConexion
210       Set oFacturacionServicios.Conexion = oConexion
220       Set oFacturacionServicioPagos.Conexion = oConexion
230       Set oFacturacionServicioFinan.Conexion = oConexion
240       Set oFacturacionServicioDespacho.Conexion = oConexion
250       Set oFactServicioPagos.Conexion = oConexion
260       Set oMovimientosHistoriaClinica.Conexion = oConexion
270       Set oSunasaPacientesHistoricos.Conexion = oConexion
280       Set oAtencionesDatosAdicionales.Conexion = oConexion
290       Set oProgramacionMedica.Conexion = oConexion
          
300       bProcesoOK = True
          
          Set oRsTmp2 = EPS_CPTfijosSolesSeleccionarTodos(oConexion)
          lnMontoCPTfijoEnsoles = 0
          
          
310       bPacienteNuevo = False
320       mo_idOrdenPago = 0
          
          'Paciente nuevo
330       If oDoPaciente.idPaciente = 0 Then
340           bPacienteNuevo = True
              
              'crea la historia clinica, si es difrenete a manual
350           If oDoHistoria.IdTipoNumeracion <> sghHistoriaDefinitivaManual Then
                    oDoHistoria.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoPaciente.IdTipoNumeracion)
'                    If lcBuscaParametro.SeleccionaFilaParametro(351) = "S" Then  'Requerimiento INSNSB, Actualizado por FCV 30032015
'                        oDoHistoria.NroHistoriaClinica = oDoPaciente.NroDocumento
'                    Else
'                        oDoHistoria.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoPaciente.IdTipoNumeracion)
'                    End If
410           End If
              
420           If oDoHistoria.NroHistoriaClinica = 0 Then
430               If oDoHistoria.IdTipoNumeracion = sghHistoriaDefinitivaReciclada Then
440                   Trace "No se pudo generar el nro de de historia clínica reciclada", "AdmisionCEAgregar"
450                   bProcesoOK = False: GoTo Terminar
460               End If
470           End If
              
              
              'Agregar paciente
480           oDoPaciente.IdTipoNumeracion = oDoHistoria.IdTipoNumeracion
490           oDoPaciente.NroHistoriaClinica = oDoHistoria.NroHistoriaClinica
              
500           If Not oPaciente.Insertar(oDoPaciente) Then
510               bProcesoOK = False: ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
520           Else
530               If Not (oDoPacienteDatosAdd Is Nothing) Then
540                   oDoPacienteDatosAdd.idPaciente = oDoPaciente.idPaciente
550               End If
560               Call PacientesDatosAdicionalesPersonalesAgregar(oDoPacienteDatosAdd, oConexion)
570           End If
              
              'Agregar historia clinica (Si es historia definitiva)
580           oDoHistoria.idPaciente = oDoPaciente.idPaciente
590           oDoHistoria.NroHistoriaClinica = oDoPaciente.NroHistoriaClinica
600           Select Case oDoHistoria.IdTipoNumeracion
              Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
                  
610               If Not oHistoria.Insertar(oDoHistoria) Then
620                   bProcesoOK = False: ms_MensajeError = oHistoria.MensajeError: GoTo Terminar
630               End If
640           End Select
          
650       Else
              'Paciente antiguo
              
              'Agregar historia clinica (Si es tenia historia temporal y ahora tiene historia definitiva)
660           Select Case lIdTipoNumeracionAnterior
              Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
                  
670           Case Else
                  
                  'Si era temporal y ahora es definitiva => Genera y agrega la historia
680               Select Case oDoHistoria.IdTipoNumeracion
                  Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
                      
690                   If oDoHistoria.IdTipoNumeracion = sghHistoriaDefinitivaManual Then
700                       oDoHistoria.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoPaciente.IdTipoNumeracion)
710                   End If
                      
720                   If oDoHistoria.NroHistoriaClinica = 0 Then
730                       If oDoHistoria.IdTipoNumeracion = sghHistoriaDefinitivaReciclada Then
740                           Trace "No se pudo generar el nro de de historia clínica reciclada", "AdmisionCEAgregar"
750                           bProcesoOK = False: GoTo Terminar
760                       End If
770                   End If
                      
780                   oDoHistoria.NroHistoriaClinicaAnterior = lIdHistoriaAnterior
790                   oDoHistoria.IdTipoNumeracionAnterior = lIdTipoNumeracionAnterior
                      
800                   If Not oHistoria.Insertar(oDoHistoria) Then
810                       bProcesoOK = False: ms_MensajeError = oHistoria.MensajeError: GoTo Terminar
820                   End If
                      
830                   oDoPaciente.NroHistoriaClinica = oDoHistoria.NroHistoriaClinica

840               End Select
850           End Select
              
860           If Not oPaciente.modificar(oDoPaciente, True) Then
870               bProcesoOK = False: ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
880           Else
890               If Not (oDoPacienteDatosAdd Is Nothing) Then
900                   oDoPacienteDatosAdd.idPaciente = oDoPaciente.idPaciente
910               End If
920               Call PacientesDatosAdicionalesPersonalesAgregar(oDoPacienteDatosAdd, oConexion)
930           End If
                 
940       End If
          '
950       oDOCuentaAtencion.idPaciente = oDoPaciente.idPaciente
960       oDOCuentaAtencion.FechaCreacion = CDate(oDoAtencion.FechaIngreso & " " & oDoAtencion.HoraIngreso)
970       If Not oCuentaAtencion.Insertar(oDOCuentaAtencion) Then
980           bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
990           rs.Close
1000      End If
          'Sunasa
1010      If oDoSunasaPacientesHistoricos.NuevoSeguro = True Then
1020          oDoSunasaPacientesHistoricos.idPaciente = oDoPaciente.idPaciente
1030          If Not oSunasaPacientesHistoricos.Insertar(oDoSunasaPacientesHistoricos) Then
1040              bProcesoOK = False: ms_MensajeError = oSunasaPacientesHistoricos.MensajeError: GoTo Terminar
1050          End If
1060      Else
1070          oDoSunasaPacientesHistoricos.idPaciente = oDoPaciente.idPaciente
1080          If Not oSunasaPacientesHistoricos.modificar(oDoSunasaPacientesHistoricos) Then
1090              bProcesoOK = False: ms_MensajeError = oSunasaPacientesHistoricos.MensajeError: GoTo Terminar
1100          End If
1110      End If
1120      oDoAtencion.idSunasaPacienteHistorico = oDoSunasaPacientesHistoricos.idSunasaPacienteHistorico
          '
          'IMPORTANTE!!!!!:    Aqui se crea ATENCION
1130      oDoAtencion.idPaciente = oDoPaciente.idPaciente
1140      oDoAtencion.idCuentaAtencion = oDOCuentaAtencion.idCuentaAtencion   'Se acocia a su atención
1150      If Not oAtencion.Insertar(oDoAtencion) Then
1160          bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
1170      End If
          '
          
1180      Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "A", oDoAtencion.idAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, lcObservacionesAuditoria)    'ListBarItems.idListItem
          
          'Agregar solicitud de prestamo
          'If mo_ReglasComunes.ParametrosSeleccionarValorIntPorTipoYCodigo("INDICADOR", "IMPRIMIR_RECIBO") = 0 Then
1190          Select Case oDoHistoria.IdTipoNumeracion
              Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
                  'InicializarPrestamoAutomaticoHosp oDOAtencion, oDOPAciente, oDOHistoriaSolicitada, bPacienteNuevo
1200              InicializarPrestamoAutomaticoCE oDOCita, oDoPaciente, oDOHistoriaSolicitada, bPacienteNuevo, oDoAtencion.idAtencion
1210              If Not oHistoriaSolicitada.Insertar(oDOHistoriaSolicitada) Then
1220                  bProcesoOK = False: ms_MensajeError = oHistoriaSolicitada.MensajeError: GoTo Terminar
1230              End If
1240          End Select
          'End If
          
          
          
          'Agregar diagnosticos
1250      If Not oAtencionesDx.ActualizarDiagnosticosAtencion(oDiagnosticos, oDoAtencion.idAtencion) Then
1260          bProcesoOK = False: GoTo Terminar
1270      End If
                  
          'IMPORTANTE AQUI SE AGREGA LA CITA
1280      oDOCita.idAtencion = oDoAtencion.idAtencion
1290      oDOCita.idPaciente = oDoPaciente.idPaciente
1300      If Not oCita.Insertar(oDOCita) Then
1310          bProcesoOK = False: ms_MensajeError = oCita.MensajeError: GoTo Terminar
1320      End If
            

          'Aqui se agrega automaticamente el costo de la CONSULTA, el costo de CARNÉ y FOLDER
          'que fueron almacenados en mo_FacturacionServicios
          'Crear orden de servicio
          Dim oDoFactOrdenServ As New DoFactOrdenServ
1330      With oDoFactOrdenServ
1340          .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL()
1350          .idCuentaAtencion = oDoAtencion.idCuentaAtencion
1360          .IdServicioPaciente = oDoAtencion.idServicioIngreso
1370          .IdEstadoFacturacion = 1               'Registrado
1380          .idPaciente = oDoAtencion.idPaciente
1390          .IdPuntoCarga = 6   'Admision de CE
1400          .IdTipoFinanciamiento = oDoAtencion.IdFormaPago             '(Contado,sis,soat,convenio,exonerado)
1410          .idFuenteFinanciamiento = oDoAtencion.idFuenteFinanciamiento
1420          .idUsuario = oDoAtencion.IdUsuarioAuditoria
1430          .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1440          .fechaDespacho = .FechaCreacion
1450          .IdUsuarioDespacho = oDoAtencion.IdUsuarioAuditoria
              'SCCQ 23-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
               '.CodPres = mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion1 'HRA 09/12/2020 Cambio 46
              'SCCQ 23-04-2021 Cambio 64 Fin (Reversion Cambio 46)
1460      End With
1470      If Not oFactOrdenesServicio.Insertar(oDoFactOrdenServ) Then
1480          bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
1490      End If
1500      If ml_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oDoAtencion.IdFormaPago, oConexion) = True Then
             '***********Contado ***********
1510         With oDOFactOrdenServicioPagos
1520             .FechaCreacion = oDoFactOrdenServ.FechaCreacion
1530             .IdEstadoFacturacion = 1   'Registrado
1540             .IdOrden = oDoFactOrdenServ.IdOrden
1550             .idUsuario = oDoAtencion.IdUsuarioAuditoria
1560             .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1570         End With
1580         If Not oFactServicioPagos.Insertar(oDOFactOrdenServicioPagos) Then
1590             bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
1600         End If
1610      End If
1620      ms_MensajeError = ""
          lnImporteTotal = 0
1630      For Each oDOFacturacionServicio In mo_FacturacionServicios

              If oDoAtencion.EpsPorcentaje > 0 Then
                oRsTmp2.Filter = "idFuenteFinanciamiento=" & oDoAtencion.idFuenteFinanciamiento & _
                              " and idproducto=" & oDOFacturacionServicio.idProducto
                If oRsTmp2.RecordCount > 0 Then
                     lnMontoCPTfijoEnsoles = oRsTmp2!Importe
                End If
              End If

1640          lnPrecioUnitario = 0
1650          oDoCatalogoServicioHosp.PrecioUnitario = 0
              'Despachos    (precio=contado)
1660          lnPrecioUnitario = 0
1670          oDoCatalogoServicioHosp.PrecioUnitario = 0
1680          If oDOFacturacionServicio.idProducto > 0 Then
1690              Set oDoCatalogoServicioHosp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarPorId(oDOFacturacionServicio.idProducto, oDoAtencion.IdFormaPago)
1700              If oDoCatalogoServicioHosp.PrecioUnitario = 0 And oDoCatalogoServicioHosp.SeUsaSinPrecio = False Then
1710                 lcCodigoNombre = oBuscaCodigoNombre.DevuelveNombreMedicamentoOServicioSegunId(oDOFacturacionServicio.idProducto, sghServicio)
1720                 ms_MensajeError = "Tiene problemas con el ID SERVICIO: " & oDOFacturacionServicio.idProducto & " " & lcCodigoNombre & " (parametros)" & Chr(13) & "para el Producto/Plan: " & oDoAtencion.IdFormaPago & Chr(13) & "...consulte al ADMINISTRADOR DEL SISTEMA...."
1730                 bProcesoOK = False: GoTo Terminar
1740              Else
1750                 lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
1760              End If
                 ' If oDoCatalogoServicioHosp.PrecioUnitario > 0 Then
1770                  lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
1780                  With oDoFacturacionServicioDesp
1790                      .cantidad = oDOFacturacionServicio.cantidad
1800                      .IdOrden = oDoFactOrdenServ.IdOrden
1810                      .idProducto = oDOFacturacionServicio.idProducto
1820                      .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1830                      .Precio = lnPrecioUnitario
1840                      .total = oDOFacturacionServicio.cantidad * lnPrecioUnitario
                          'SCCQ 26-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
                          'HRA 09/12/2020 Cambio 46 Inicio
                          '.IdFinan = oDoAtencion.IdFormaPago
                          '.CodPres = mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion1
                          'HRA 09/12/2020 Cambio 46 Fin
                          'SCCQ 26-04-2021 Cambio 64 Fin (Reversion Cambio 46)
                   
1850                  End With
1860                  If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then
1870                         bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
1880                  End If
1890                  Set oDoCatalogoServicioHosp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarPorId(oDOFacturacionServicio.idProducto, oDoAtencion.IdFormaPago)
1900                  If ml_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oDoAtencion.IdFormaPago, oConexion) = True Then
                         '***********Contado ***********
1910                     With oDoFacturacionServicioPagos
1920                         .cantidad = oDOFacturacionServicio.cantidad
1930                         .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
1940                         .idProducto = oDOFacturacionServicio.idProducto
1950                         .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1960                         .Precio = lnPrecioUnitario
1970                         .total = oDOFacturacionServicio.cantidad * lnPrecioUnitario
1980                     End With
1990                     If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
2000                         bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
2010                     End If
2020                     mo_idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
2030                  Else
                         '***********Sis,Soat,exoneracion, convenio ***********
2040                     With oDoFacturacionServicioFinanc
2050                         .CantidadFinanciada = oDOFacturacionServicio.cantidad
2060                         .FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
2070                         .IdOrden = oDoFactOrdenServ.IdOrden
2080                         .idProducto = oDOFacturacionServicio.idProducto
2090                         .IdTipoFinanciamiento = oDoAtencion.IdFormaPago
2100                         .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2110                         .IdUsuarioAutoriza = oDoAtencion.IdUsuarioAuditoria
2120                         .PrecioFinanciado = lnPrecioUnitario
2130                         .TotalFinanciado = oDOFacturacionServicio.cantidad * lnPrecioUnitario
                             .idFuenteFinanciamiento = oDoAtencion.idFuenteFinanciamiento
                             'SCCQ 26-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
                             'HRA 09/12/2020 Cambio 46 Inicio
2140                         '.idFuenteFinanciamiento = oDoAtencion.IdFormaPago
                             '.CodPres = mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion1
                             'HRA 09/12/2020 Cambio 46
                             'SCCQ 26-04-2021 Cambio 64 Fin (Reversion Cambio 46)
                                                
2150                     End With
2160                     If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then
2170                         bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo Terminar
2180                     End If
                         lnImporteTotal = lnImporteTotal + oDoFacturacionServicioFinanc.TotalFinanciado
2190                  End If
                  'End If
2200          End If
2210      Next
          '********************Actualiza Paquete (si fuese el Caso)***************
2220      lbYaSeTransfirioHCdeUnServicioAotro = False
2230      If DOFacturacionPaquetesUltimo.IdComprobantePago > 0 Then
2240          If FacturacionPaquetesActualizaXidComprobIdOrdenIdProducto(DOFacturacionPaquetesUltimo.idProducto, _
                                       oDoAtencion.idCuentaAtencion, DOFacturacionPaquetesUltimo.IdComprobantePago, _
                                       DOFacturacionPaquetesUltimo.idOrdenPago, True, oConexion) = False Then
2250             bProcesoOK = False:  GoTo Terminar
2260          End If
              'actualiza IdPaciente en Boleta
2270          If CajaComprobantesPagoActualizaXpacienteYComprobante(oDoAtencion.idPaciente, _
                                    DOFacturacionPaquetesUltimo.IdComprobantePago, False, oConexion) = False Then
2280             bProcesoOK = False:  GoTo Terminar
2290          End If
              'Agrega MOVIMIENTO DE ARCHIVO si no esta la Historia ubicada en Archivo Clinico}
2300          Set oRsTmp1 = mo_ReglasArchivoClinico.MovimientosHistoriaClinicaSeleccionaUltimoMovimientoPorPaciente(oDoAtencion.idPaciente)
2310          If oRsTmp1.RecordCount > 0 Then
2320              If oRsTmp1.Fields!idServicioDestino <> oBuscaCodigoNombre.ParametrosIdServicioArchivoClinico Then
2330                  oDOMovimientoHistoriaClinica.idPaciente = oDoAtencion.idPaciente
2340                  oDOMovimientoHistoriaClinica.FechaMovimiento = lcBuscaParametro.RetornaFechaHoraServidorSQL()
2350                  oDOMovimientoHistoriaClinica.IdMotivo = 7     'Transferencia
2360                  oDOMovimientoHistoriaClinica.IdServicioOrigen = oRsTmp1.Fields!idServicioDestino
2370                  oDOMovimientoHistoriaClinica.idServicioDestino = oDoAtencion.idServicioIngreso
2380                  oDOMovimientoHistoriaClinica.IdEmpleadoArchivo = oDoAtencion.IdUsuarioAuditoria
2390                  oDOMovimientoHistoriaClinica.IdEmpleadoTransporte = oDoAtencion.IdUsuarioAuditoria
2400                  oDOMovimientoHistoriaClinica.IdEmpleadoRecepcion = oDoAtencion.IdUsuarioAuditoria
2410                  oDOMovimientoHistoriaClinica.IdGrupoMovimiento = 1
2420                  oDOMovimientoHistoriaClinica.idAtencion = oDoAtencion.idAtencion

2430                  If Not oMovimientosHistoriaClinica.Insertar(oDOMovimientoHistoriaClinica) Then
2440                      bProcesoOK = False: ms_MensajeError = oMovimientosHistoriaClinica.MensajeError: GoTo Terminar
2450                  End If
2460                  lbYaSeTransfirioHCdeUnServicioAotro = True
2470              End If
2480          End If
2490          oRsTmp1.Close
2500      End If
          'datos adicionales
    
2510      mo_DoAtencionDatosAdicionales.idAtencion = oDoAtencion.idAtencion
2520      If Not oAtencionesDatosAdicionales.Insertar(mo_DoAtencionDatosAdicionales) Then
2530          bProcesoOK = False: ms_MensajeError = oAtencionesDatosAdicionales.MensajeError: GoTo Terminar
2540      End If
2550      If mo_DoAtencionDatosAdicionales.idSiaSis > 0 Then
2560         Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "A", oDoAtencion.idCuentaAtencion, "FUA", oConexion, 1345, lcNombrePc, "Cod.Prestacion FUA: " & mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion & " N° Cuenta: " & oDoAtencion.idCuentaAtencion)
2570      End If
          'Cupo Adicional
2580      If oDOCita.EsCitaAdicional = True Then
2590         oDOProgramacionMedica.idProgramacion = oDOCita.idProgramacion
2600         If Not oProgramacionMedica.SeleccionarPorId(oDOProgramacionMedica) Then
2610              bProcesoOK = False: ms_MensajeError = oProgramacionMedica.MensajeError: GoTo Terminar
2620         End If
2630         oDOProgramacionMedica.IdUsuarioAuditoria = oDOCita.IdUsuarioAuditoria
2640         oDOProgramacionMedica.horaFin = oDOCita.horaFin
      '       If Not oProgramacionMedica.modificar(oDOProgramacionMedica) Then
2650         If Not oProgramacionMedica.ModificarHoraFinPorCitaAdicional(oDOProgramacionMedica) Then
2660              bProcesoOK = False: ms_MensajeError = oProgramacionMedica.MensajeError: GoTo Terminar
2670         End If
2680      End If
          '
          mo_ReglasArchivoClinico.HistoriasClinicasActualizaFmovimiento oConexion, oDoAtencion.idPaciente
          '
          AsignaPorcentajeEPS oConexion, oDoAtencion.idCuentaAtencion, oDoAtencion.EpsPorcentaje, lnIdTablaLISTBARITEMS, _
                               sghAgregar, oDoFactOrdenServ, oFactOrdenesServicio, oDoFacturacionServicioDesp, _
                               oFacturacionServicioDespacho, oDOFactOrdenServicioPagos, oFactServicioPagos, _
                               oDoFacturacionServicioPagos, oFacturacionServicioPagos, oDoFactOrdenServ.FechaCreacion, _
                               oDoFactOrdenServ.IdServicioPaciente, oDoFactOrdenServ.idPaciente, _
                               oDoFactOrdenServ.IdUsuarioAuditoria, oDoFactOrdenServ.IdOrden, lnImporteTotal, _
                               sghPuntosCargaBasicos.sghPtoCargaAdmisionCE, 0, lnMontoCPTfijoEnsoles

          '
Terminar:
2690      If bProcesoOK Then
2700          AdmisionCEAgregar = True
2710          oConexion.CommitTrans
2720      Else
2730          AdmisionCEAgregar = False
2740          oConexion.RollbackTrans
2750          ms_MensajeError = ms_MensajeError & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                               "Function AdmisionCEAgregar", "reglasAdmision.cls") 'debb-02/05/2016
                     
2760      End If
                  
2770      oConexion.Close
2780      Set oConexion = Nothing
2790      Set oPaciente = Nothing
2800      Set oCuentaAtencion = Nothing
2810      Set oAtencion = Nothing
2820      Set oDOHistoriaSolicitada = Nothing
2830      Set oHistoriaSolicitada = Nothing
2840      Set oCita = Nothing
2850      Set oHistoria = Nothing
2860      Set oAtencionesDx = Nothing
2870      Set oCatalogoServicios = Nothing
2880      Set oFacturacionServicios = Nothing
2890      Set oFactOrdenesServicio = Nothing
2900      Set oFacturacionServicioDespacho = Nothing
2910      Set oFacturacionServicioPagos = Nothing
2920      Set oFacturacionServicioFinan = Nothing
2930      Set oFactServicioPagos = Nothing
2940      Set oDoFacturacionServicioPagos = Nothing
2950      Set oDoFacturacionServicioFinanc = Nothing
2960      Set oDoFacturacionServicioDesp = Nothing
2970      Set oDOFacturacionServicio = Nothing
2980      Set oDoCatalogoServicioHosp = Nothing
2990      Set oDOFactOrdenServicioPagos = Nothing
3000      Set ml_ReglasFacturacion = Nothing
3010      Set oBuscaCodigoNombre = Nothing
3020      Set oDOMovimientoHistoriaClinica = Nothing
3030      Set oMovimientosHistoriaClinica = Nothing
3040      Set oRsTmp1 = Nothing
3050      Set mo_ReglasArchivoClinico = Nothing
3060      Set oSunasaPacientesHistoricos = Nothing
3070      Set oAtencionesDatosAdicionales = Nothing
3080      Set oDOProgramacionMedica = Nothing
3090      Set oProgramacionMedica = Nothing
          Set oRsTmp2 = Nothing
End Function

Sub AsignaPorcentajeEPS(oConexion As Connection, lnIdCuentaAtencion As Long, lnEpsPorcentaje As Double, _
                        lnIdTablaLISTBARITEMS As sghOpcionGalenHos, lnOpcion As sghOpciones, _
                        oDoFactOrdenServ As DoFactOrdenServ, oFactOrdenesServicio As FactOrdenServicio, _
                        oDoFacturacionServicioDesp As DoFacturacionServicioDespacho, oFacturacionServicioDespacho As FacturacionServicioDespacho, _
                        oDOFactOrdenServicioPagos As DoFactOrdenServPagos, oFactServicioPagos As FactOrdenServicioPagos, _
                        oDoFacturacionServicioPagos As DoFacturacionServicioPagos, oFacturacionServicioPagos As FacturacionServicioPagos, _
                        ldFechaCreacion As Date, lnIdServicioIngreso As Long, lnIdPaciente As Long, lnIdUsuario As Long, _
                        lnIdOrden_Servicio As Long, lnImporteTotal As Double, lnIdPuntoCarga As Long, lnIdOrdenPago As Long, _
                        lnMontoCPTfijoEnsoles As Double)
    On Error GoTo ErrAsgPorc
    If lnEpsPorcentaje > 0 And lnEpsPorcentaje < 100 Then
       Dim lnImporteQpagara As Double, lnImporteCubreEPS As Double
       Select Case lnIdTablaLISTBARITEMS
       Case sghRegistroCitaCE, sghLaboratorioAP, sghLaboratorioPC, sghLaboratorioBS, _
            sghImagenTomografia, sghImagenRayosX, sghImagenEcogO, sghImagenEcogG
            If lnOpcion = sghAgregar Or lnOpcion = sghModificar Then
               'lnImporteQpagara = Round(lnImporteTotal * ((100 - lnEpsPorcentaje) / 100), 2)
               lnImporteQpagara = 0
               If lnMontoCPTfijoEnsoles > 0 Then
                   '****** el CPT tiene MONTO FIJO, no se toma en cuenta el PORCENTAJE (Ejm: pago de CITA)
                   oDoFactOrdenServ.IdEstadoFacturacion = sghEstadoFacturacion.sghAnulado
                   If oFactOrdenesServicio.modificar(oDoFactOrdenServ) Then
                      lnImporteQpagara = lnMontoCPTfijoEnsoles
                   End If
               Else
                    '****disminuir precios,Totales segun porcentaje en tabla Seguros -CPT
                    If EPSdisminuyePrecio(oConexion, lnEpsPorcentaje, lnIdOrden_Servicio, lnImporteCubreEPS) = True Then
                       lnImporteQpagara = lnImporteTotal - lnImporteCubreEPS
                    End If
               End If
               If lnImporteQpagara > 0 Then
                  '****crear ORDEN DE PAGO por el porcentaje que no cubre
                  EPScrearOrdenPago oDoFactOrdenServ, oFactOrdenesServicio, oDoFacturacionServicioDesp, oFacturacionServicioDespacho, _
                                oDOFactOrdenServicioPagos, oFactServicioPagos, oDoFacturacionServicioPagos, oFacturacionServicioPagos, _
                                ldFechaCreacion, lnImporteQpagara, lnIdCuentaAtencion, lnIdServicioIngreso, lnIdPaciente, lnIdUsuario, _
                                lnIdPuntoCarga, lnIdTablaLISTBARITEMS, lnIdOrdenPago
               End If
            ElseIf lnOpcion = sghEliminar Then
            
            End If
       End Select
    End If
ErrAsgPorc:
End Sub

Function EPS_CPTfijosSolesSeleccionarTodos(oConexion As Connection) As Recordset
    On Error GoTo ErrEPSDISMY1
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandType = adCmdStoredProc
        .CommandText = "EPS_CPTfijosSolesSeleccionarTodos"
        Set oRecordset = .Execute
    End With
    Set EPS_CPTfijosSolesSeleccionarTodos = oRecordset
    Set oCommand = Nothing
    Exit Function
ErrEPSDISMY1:
End Function

Function EPSdisminuyePrecio(oConexion As Connection, lnEpsPorcentaje As Double, lnIdOrden As Long, ByRef lnImporteCubreEPS As Double) As Boolean
    On Error GoTo ErrEPSDISMY
    EPSdisminuyePrecio = False
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandType = adCmdStoredProc
        .CommandText = "FacturacionServicioFinanciamientosActualizaPrecio"
        Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@porcentaje", adCurrency, adParamInput, 0, lnEpsPorcentaje): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ImporteCubreEPS", adCurrency, adParamOutput, 0): .Parameters.Append oParameter
        .Execute
        lnImporteCubreEPS = .Parameters("@ImporteCubreEPS")
    End With
    Set oCommand = Nothing
    EPSdisminuyePrecio = True
    Exit Function
ErrEPSDISMY:
End Function

Function EPSeliminaOrdenPago(oConexion As Connection, lnIdOrdenPago As Long, lnIdUsuario As Long) As Boolean
      On Error GoTo ErrEPScrearOP
      EPSeliminaOrdenPago = False
      Dim oDoFactOrdenServ As New DoFactOrdenServ, oFactOrdenesServicio As New FactOrdenServicio
      Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho, oFacturacionServicioDespacho As New FacturacionServicioDespacho
      Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos, oFactServicioPagos As New FactOrdenServicioPagos
      Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos, oFacturacionServicioPagos As New FacturacionServicioPagos
      Dim lnIdProductoEPS As Long, lnIdOrden As Long
      lnIdProductoEPS = Val(lcBuscaParametro.SeleccionaFilaParametro(566))
      Set oFactOrdenesServicio.Conexion = oConexion
      Set oFacturacionServicioDespacho.Conexion = oConexion
      Set oFactServicioPagos.Conexion = oConexion
      Set oFacturacionServicioPagos.Conexion = oConexion
      '
      If lnIdOrdenPago > 0 Then
            oDOFactOrdenServicioPagos.idOrdenPago = lnIdOrdenPago
            oDOFactOrdenServicioPagos.idUsuario = lnIdUsuario
            If oFactServicioPagos.SeleccionarPorId(oDOFactOrdenServicioPagos) = True Then
                lnIdOrden = oDOFactOrdenServicioPagos.IdOrden
                oDoFacturacionServicioPagos.idOrdenPago = lnIdOrdenPago
                oDoFacturacionServicioPagos.IdUsuarioAuditoria = lnIdUsuario
                oDoFacturacionServicioPagos.idProducto = lnIdProductoEPS
                If oFacturacionServicioPagos.Eliminar(oDoFacturacionServicioPagos) Then
                    oDOFactOrdenServicioPagos.idOrdenPago = lnIdOrdenPago
                    oDOFactOrdenServicioPagos.IdUsuarioAuditoria = lnIdUsuario
                    If oFactServicioPagos.Eliminar(oDOFactOrdenServicioPagos) = True Then
                       oDoFacturacionServicioDesp.IdOrden = lnIdOrden
                       oDoFacturacionServicioDesp.IdUsuarioAuditoria = lnIdUsuario
                       If oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then
                          oDoFactOrdenServ.IdOrden = lnIdOrden
                          oDoFactOrdenServ.IdUsuarioAuditoria = lnIdUsuario
                          If oFactOrdenesServicio.Eliminar(oDoFactOrdenServ) = True Then
                          End If
                       End If
                    End If
                End If
            End If
      End If
      EPSeliminaOrdenPago = True
      Set oDoFactOrdenServ = Nothing
      Set oFactOrdenesServicio = Nothing
      Set oDoFacturacionServicioDesp = Nothing
      Set oFacturacionServicioDespacho = Nothing
      Set oDOFactOrdenServicioPagos = Nothing
      Set oFactServicioPagos = Nothing
      Set oDoFacturacionServicioPagos = Nothing
      Set oFacturacionServicioPagos = Nothing
      Exit Function
ErrEPScrearOP:

End Function


'Function eps_cptfijosXfuenteFinanciamiento(lnIdFuenteFinanciamiento As Long, oConexion As Connection) As ADODB.Recordset
'On Error GoTo ManejadorDeError
'Dim oRecordset As New ADODB.Recordset
'Dim oCommand As New ADODB.Command
'Dim oParameter As ADODB.Parameter
'Dim ms_MensajeError As String
'    Set eps_cptfijosXfuenteFinanciamiento = Nothing
'    ms_MensajeError = ""
'    With oCommand
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = oConexion
'        .CommandTimeout = 150
'        .CommandText = "eps_cptfijosXfuenteFinanciamiento"
'        Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
'        Set oRecordset = .Execute
'   End With
'   Set eps_cptfijosXfuenteFinanciamiento = oRecordset
'   Set oCommand = Nothing
'Exit Function
'ManejadorDeError:
'   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
'Exit Function
'End Function



Function EPScrearOrdenPago(oDoFactOrdenServ As DoFactOrdenServ, oFactOrdenesServicio As FactOrdenServicio, _
                      oDoFacturacionServicioDesp As DoFacturacionServicioDespacho, oFacturacionServicioDespacho As FacturacionServicioDespacho, _
                      oDOFactOrdenServicioPagos As DoFactOrdenServPagos, oFactServicioPagos As FactOrdenServicioPagos, _
                      oDoFacturacionServicioPagos As DoFacturacionServicioPagos, oFacturacionServicioPagos As FacturacionServicioPagos, _
                      ldFechaCreacion As Date, lnImporteQpagara As Double, lnIdCuenta As Long, lnIdServicioIngreso As Long, _
                      lnIdPaciente As Long, lnIdUsuario As Long, lnIdPuntoCarga As Long, lnIdTablaLISTBARITEMS As sghOpcionGalenHos, _
                      lnIdOrdenPago As Long) As Boolean
      On Error GoTo ErrEPScrearOP
      EPScrearOrdenPago = False
      Dim lnIdProductoEPS As Long, lnIdOrden As Long
'      Dim oRsTmp9 As New Recordset
'      Set oRsTmp9 = eps_cptfijosXfuenteFinanciamiento(lnIdFuenteFinanciamiento)
      Select Case lnIdTablaLISTBARITEMS
      Case sghLaboratorioPC, sghLaboratorioBS, sghLaboratorioAP
           lnIdProductoEPS = Val(lcBuscaParametro.SeleccionaFilaParametro(564))
      Case sghImagenTomografia, sghImagenRayosX, sghImagenEcogO, sghImagenEcogG
           lnIdProductoEPS = Val(lcBuscaParametro.SeleccionaFilaParametro(565))
      Case sghRegistroCitaCE
           lnIdProductoEPS = Val(lcBuscaParametro.SeleccionaFilaParametro(563))
      End Select
      If lnIdOrdenPago > 0 Then
            oDOFactOrdenServicioPagos.idOrdenPago = lnIdOrdenPago
            oDOFactOrdenServicioPagos.idUsuario = lnIdUsuario
            If oFactServicioPagos.SeleccionarPorId(oDOFactOrdenServicioPagos) = True Then
                lnIdOrden = oDOFactOrdenServicioPagos.IdOrden
                oDoFacturacionServicioPagos.idOrdenPago = lnIdOrdenPago
                oDoFacturacionServicioPagos.IdUsuarioAuditoria = lnIdUsuario
                oDoFacturacionServicioPagos.idProducto = lnIdProductoEPS
                If oFacturacionServicioPagos.Eliminar(oDoFacturacionServicioPagos) Then
                    oDOFactOrdenServicioPagos.idOrdenPago = lnIdOrdenPago
                    oDOFactOrdenServicioPagos.IdUsuarioAuditoria = lnIdUsuario
                    If oFactServicioPagos.Eliminar(oDOFactOrdenServicioPagos) = True Then
                       oDoFacturacionServicioDesp.IdOrden = lnIdOrden
                       oDoFacturacionServicioDesp.IdUsuarioAuditoria = lnIdUsuario
                       If oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then
                          oDoFactOrdenServ.IdOrden = lnIdOrden
                          oDoFactOrdenServ.IdUsuarioAuditoria = lnIdUsuario
                          If oFactOrdenesServicio.Eliminar(oDoFactOrdenServ) = True Then
                          End If
                       End If
                    End If
                End If
            End If
      End If
      With oDoFactOrdenServ
          .FechaCreacion = ldFechaCreacion
          .idCuentaAtencion = lnIdCuenta
          .IdServicioPaciente = lnIdServicioIngreso
          .IdEstadoFacturacion = 1               'Registrado
          .idPaciente = lnIdPaciente
          .IdPuntoCarga = lnIdPuntoCarga
          .IdTipoFinanciamiento = sghTipoFinanciamiento.sghPacienteNormal
          .idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente
          .idUsuario = lnIdUsuario
          .IdUsuarioAuditoria = lnIdUsuario
          .fechaDespacho = .FechaCreacion
          .IdUsuarioDespacho = lnIdUsuario
      End With
      If oFactOrdenesServicio.Insertar(oDoFactOrdenServ) = True Then
         With oDoFacturacionServicioDesp
            .cantidad = 1
            .IdOrden = oDoFactOrdenServ.IdOrden
            .idProducto = lnIdProductoEPS
            .IdUsuarioAuditoria = lnIdUsuario
            .Precio = lnImporteQpagara
            .total = lnImporteQpagara
         End With
         If oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then
            With oDOFactOrdenServicioPagos
                .FechaCreacion = oDoFactOrdenServ.FechaCreacion
                .IdEstadoFacturacion = 1   'Registrado
                .IdOrden = oDoFactOrdenServ.IdOrden
                .idUsuario = lnIdUsuario
                .IdUsuarioAuditoria = lnIdUsuario
            End With
            If oFactServicioPagos.Insertar(oDOFactOrdenServicioPagos) = True Then
                   With oDoFacturacionServicioPagos
                       .cantidad = 1
                       .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
                       .idProducto = lnIdProductoEPS
                       .IdUsuarioAuditoria = lnIdUsuario
                       .Precio = lnImporteQpagara
                       .total = lnImporteQpagara
                   End With
                   If oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
                      mo_idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
                      EPScrearOrdenPago = True
                   End If
            End If
          End If
      End If
      Exit Function
ErrEPScrearOP:

End Function



Function AdmisionCEModificar(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, oDoPaciente As DOPaciente, _
                             oDOCita As doCita, oDoHistoria As DOHistoriaClinica, _
                             lIdTipoNumeracionAnterior As sghTipoNumeracionDeNroHistoria, oDiagnosticos As Collection, _
                             oProcedimientos As Collection, oExamenes As Collection, lIdHistoriaAnterior As Long, _
                             mo_FacturacionServicios As Collection, mo_FacturacionServiciosAEliminar As Collection, _
                             DOFacturacionPaquetesUltimo As DOFacturacionPaquetes, _
                             DOFacturacionPaquetesAnterior As DOFacturacionPaquetes, _
                             lbYaSeTransfirioHCdeUnServicioAotro As Boolean, lnIdTablaLISTBARITEMS As Long, _
                             lcNombrePc As String, lcObservacionesAuditoria As String, _
                             oDoSunasaPacientesHistoricos As DoSunasaPacientesHistoricos, _
                             mo_DoAtencionDatosAdicionales As DoAtencionDatosAdicionales, _
                             lbNecesitaCrearDatosEnTablasFacturacion As Boolean, lbEsUnEPSdesdeAgregarCE As Boolean, _
                             lnIdOrdenPago As Long, _
                             Optional oDoPacienteDatosAdd As DoPacienteDatosAdd = Nothing) As Boolean
      Dim oConexion As New ADODB.Connection
      Dim oPaciente As New Pacientes
      Dim oCuentaAtencion As New CuentasAtencion
      Dim oAtencion As New Atenciones
      Dim oDOHistoriaSolicitada As New DOHistoriaSolicitada
      Dim oHistoriaSolicitada As New HistoriasSolicitadas
      Dim oCita As New Citas
      Dim bProcesoOK As Boolean
      Dim oHistoria As New HistoriasClinicas
      Dim oAtencionesDx As New AtencionesDiagnosticos
      Dim oFacturacionServicios As New FacturacionServicios
      Dim oCatalogoServicios As New CatalogoServicios
      Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
      Dim oFacturacionServicioPagos As New FacturacionServicioPagos
      Dim oFacturacionServicioFinan As New FacturacionServicioFin
      Dim oFactServicioPagos As New FactOrdenServicioPagos
      Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
      Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
      Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
      Dim oDOFacturacionServicio As DOFacturacionServicios
      Dim oDoFactOrdenServ As New DoFactOrdenServ
      Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
      Dim oDOMovimientoHistoriaClinica As New DOMovimientoHistoriaClinica
      Dim oMovimientosHistoriaClinica As New MovimientosHistoriaClinica
      Dim oSunasaPacientesHistoricos As New SunasaPacientesHistoricos
      Dim oAtencionesDatosAdicionales As New AtencionesDatosAdicionales
      Dim oFactOrdenesServicio As New FactOrdenServicio
      Dim oRsTmp As New Recordset
      Dim oRsTmp1 As New Recordset
      Dim lcCodigoNombre As String
      Dim lnPrecioUnitario As Double
      Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
      Dim ml_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
      Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
      Dim oBuscaCodigoNombre As New SIGHNegocios.ReglasComunes
      Dim lcSql As String, lnImporteTotal As Double
      Dim oRsTmp2 As New Recordset
      Dim lnMontoCPTfijoEnsoles As Double
            oConexion.CommandTimeout = 900
            oConexion.CursorLocation = adUseClient
20        oConexion.Open sighentidades.CadenaConexion
          
40        Set oCuentaAtencion.Conexion = oConexion
50        Set oAtencion.Conexion = oConexion
60        Set oPaciente.Conexion = oConexion
70        Set oHistoriaSolicitada.Conexion = oConexion
80        Set oCita.Conexion = oConexion
90        Set oHistoria.Conexion = oConexion
100       Set oAtencionesDx.Conexion = oConexion
110       Set oFacturacionServicios.Conexion = oConexion
120       Set oCatalogoServicios.Conexion = oConexion
130       Set oFacturacionServicioPagos.Conexion = oConexion
140       Set oFacturacionServicioFinan.Conexion = oConexion
150       Set oFacturacionServicioDespacho.Conexion = oConexion
160       Set oFactServicioPagos.Conexion = oConexion
170       Set oMovimientosHistoriaClinica.Conexion = oConexion
180       Set oSunasaPacientesHistoricos.Conexion = oConexion
190       Set oAtencionesDatosAdicionales.Conexion = oConexion
          Set oFactOrdenesServicio.Conexion = oConexion
          
200       oConexion.BeginTrans
210       bProcesoOK = True
          
          Set oRsTmp2 = EPS_CPTfijosSolesSeleccionarTodos(oConexion)
          lnMontoCPTfijoEnsoles = 0
          
          
          'Agregar historia clinica (Si es tenia historia temporal y ahora tiene historia definitiva)
220       Select Case lIdTipoNumeracionAnterior
          Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
230           If lIdTipoNumeracionAnterior = sghHistoriaDefinitivaManual Then
240               If Not oHistoria.modificar(oDoHistoria) Then
250               End If
260           End If
              
270       Case Else
              
              'Si era temporal y ahora es definitiva => Genera y agrega la historia
280           Select Case oDoHistoria.IdTipoNumeracion
              Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
                  
290               If oDoHistoria.IdTipoNumeracion <> sghHistoriaDefinitivaManual Then
300                   oDoHistoria.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoPaciente.IdTipoNumeracion)
310               End If
                  
320               If oDoHistoria.NroHistoriaClinica = 0 Then
330                   If oDoHistoria.IdTipoNumeracion = sghHistoriaDefinitivaReciclada Then
340                       Trace "No se pudo generar el nro de de historia clínica reciclada", "AdmisionCEAgregar"
350                       bProcesoOK = False: ms_MensajeError = oHistoria.MensajeError: GoTo Terminar
360                   End If
370               End If
                  
380               oDoHistoria.NroHistoriaClinicaAnterior = lIdHistoriaAnterior
390               oDoHistoria.IdTipoNumeracionAnterior = lIdTipoNumeracionAnterior
                  
400               oDoPaciente.NroHistoriaClinica = oDoHistoria.NroHistoriaClinica
410               If Not oHistoria.Insertar(oDoHistoria) Then
420                   bProcesoOK = False: ms_MensajeError = oHistoria.MensajeError: GoTo Terminar
430               End If
440           End Select
450       End Select
          
460       If Not oPaciente.modificar(oDoPaciente, True) Then
470           bProcesoOK = False: GoTo Terminar
480       Else
490           If Not (oDoPacienteDatosAdd Is Nothing) Then
500               oDoPacienteDatosAdd.idPaciente = oDoPaciente.idPaciente
510           End If
520           Call PacientesDatosAdicionalesPersonalesAgregar(oDoPacienteDatosAdd, oConexion)
530       End If
          'Sunasa
540       oDoSunasaPacientesHistoricos.idPaciente = oDoPaciente.idPaciente
550       If oDoSunasaPacientesHistoricos.NuevoSeguro = True Then
560           If Not oSunasaPacientesHistoricos.Insertar(oDoSunasaPacientesHistoricos) Then
570               bProcesoOK = False: ms_MensajeError = oSunasaPacientesHistoricos.MensajeError: GoTo Terminar
580           End If
590       Else
600           If Not oSunasaPacientesHistoricos.modificar(oDoSunasaPacientesHistoricos) Then
610               bProcesoOK = False: ms_MensajeError = oSunasaPacientesHistoricos.MensajeError: GoTo Terminar
620           End If
630       End If
640       oDoAtencion.idSunasaPacienteHistorico = oDoSunasaPacientesHistoricos.idSunasaPacienteHistorico
          '
          'Modificar atencion
650       oDoAtencion.idPaciente = oDoPaciente.idPaciente
660       If lnIdTablaLISTBARITEMS = 103 Then
670           oDoAtencion.HoraEgreso = lcBuscaParametro.RetornaHoraServidorSQL
680           oDoAtencion.fechaEgreso = oDOCita.fecha
          'Else
          '    oDoAtencion.HoraEgreso = "__/__"
          '    oDoAtencion.FechaEgreso = 0
690       End If
700       If Not oAtencion.modificar(oDoAtencion) Then
710           bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
720       End If
          '
730       Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "M", oDoAtencion.idAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, lcObservacionesAuditoria)   'ListBarItems.idListItem
          '
740       oDOCuentaAtencion.FechaCreacion = CDate(oDoAtencion.FechaIngreso & " " & oDoAtencion.HoraIngreso)
750       oDOCuentaAtencion.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
760       If Not oCuentaAtencion.modificar(oDOCuentaAtencion) Then
770           bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
780       End If
          
790       If oDoAtencion.IdEstadoAtencion = sghEstadoTabla.sghAnulado Then
800           If Not oHistoriaSolicitada.EliminarPorCita(oDOCita) Then
810               bProcesoOK = False: ms_MensajeError = oHistoriaSolicitada.MensajeError: GoTo Terminar
820           End If
830       End If
          
          'Agregar diagnosticos
840       If Not oAtencionesDx.ActualizarDiagnosticosAtencion(oDiagnosticos, oDoAtencion.idAtencion) Then
850           bProcesoOK = False: ms_MensajeError = oAtencionesDx.MensajeError: GoTo Terminar
860       End If
                  
          'Modificar Cita
870       oDOCita.idAtencion = oDoAtencion.idAtencion
880       oDOCita.idPaciente = oDoPaciente.idPaciente
890       If oDoAtencion.IdEstadoAtencion = sghEstadoTabla.sghAnulado Then
900           If Not oCita.Eliminar(oDOCita) Then
910               bProcesoOK = False: ms_MensajeError = oCita.MensajeError: GoTo Terminar
920           End If
930       Else
940           If Not oCita.modificar(oDOCita) Then
950               bProcesoOK = False: GoTo Terminar
960           End If
970       End If
          
          '*************************Servicios***********************
          'elimina Codigo Servicio (cabecera/detalle) por Estancia de la Cuenta de Atencion}
980       If lbNecesitaCrearDatosEnTablasFacturacion = True Then
              Dim lbProsigue As Boolean
              lbProsigue = True
              'es un EPS (inicio)
              If lbEsUnEPSdesdeAgregarCE = True Then                             'es un EPS
                  Set oRsTmp = mo_ReglasFacturacion.FactOrdenServicioPagosPorIdCuentaAtencion(oDoAtencion.idCuentaAtencion, oConexion)
                  If oRsTmp.RecordCount > 0 Then
                       oRsTmp.MoveFirst
                       Do While Not oRsTmp.EOF
                             If oRsTmp.Fields!IdPuntoCarga = 6 Then    'admision consulta externa
                                oDoFactOrdenServ.IdOrden = oRsTmp.Fields!IdOrden
                                oDoFacturacionServicioFinanc.IdOrden = oRsTmp.Fields!IdOrden
                                oDoFacturacionServicioDesp.IdOrden = oRsTmp.Fields!IdOrden
                                If oRsTmp.Fields!IdEstadoFacturacion = 1 Then
                                     oDOFactOrdenServicioPagos.IdOrden = oRsTmp.Fields!IdOrden
                                     oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp.Fields!idOrdenPago
                                     oDoFacturacionServicioPagos.idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
                                     If Not oFacturacionServicioPagos.Eliminar(oDoFacturacionServicioPagos) Then
                                        bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
                                     End If
                                     If Not oFactServicioPagos.Eliminar(oDOFactOrdenServicioPagos) Then
                                        bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
                                     End If
                                End If
                                If Not oFacturacionServicioFinan.Eliminar(oDoFacturacionServicioFinanc) Then
                                   bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo Terminar
                                End If
                                If Not oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then
                                   bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
                                End If
                                oDoFacturacionServicioDesp.IdOrden = oRsTmp.Fields!IdOrden
                                oDoFacturacionServicioDesp.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                                If Not oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then
                                      bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
                                End If
                                oDoFactOrdenServ.IdOrden = oRsTmp.Fields!IdOrden
                                oDoFactOrdenServ.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                                If Not oFactOrdenesServicio.Eliminar(oDoFactOrdenServ) Then
                                     bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
                                End If
                                lbProsigue = False
                             End If
                            oRsTmp.MoveNext
                       Loop
                  End If
                  oRsTmp.Close
              End If
              'es un EPS (fin)
              '
              If lbProsigue = True Then
                 Set oRsTmp = mo_ReglasFacturacion.FactOrdenServicioPagosPorIdCuentaAtencion(oDoAtencion.idCuentaAtencion, oConexion)
                 If oRsTmp.RecordCount > 0 Then
                    lbProsigue = True
                 Else
                    lbProsigue = False
                 End If
              End If
          
              If lbProsigue = True Then
1010             oRsTmp.MoveFirst
1020             Do While Not oRsTmp.EOF
1030                   If oRsTmp.Fields!IdPuntoCarga = 6 Then    'admision consulta externa
1040                      oDoFactOrdenServ.IdOrden = oRsTmp.Fields!IdOrden
1050                      oDoFacturacionServicioFinanc.IdOrden = oRsTmp.Fields!IdOrden
1060                      oDoFacturacionServicioDesp.IdOrden = oRsTmp.Fields!IdOrden
1070                      If oRsTmp.Fields!IdEstadoFacturacion = 1 Then
1080                           oDOFactOrdenServicioPagos.IdOrden = oRsTmp.Fields!IdOrden
1090                           oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp.Fields!idOrdenPago
1100                           oDoFacturacionServicioPagos.idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
1110                           If Not oFacturacionServicioPagos.Eliminar(oDoFacturacionServicioPagos) Then
1120                              bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
1130                           End If
1140                           If Not oFactServicioPagos.Eliminar(oDOFactOrdenServicioPagos) Then
1150                              bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
1160                           End If
1170                      End If
1180                      If Not oFacturacionServicioFinan.Eliminar(oDoFacturacionServicioFinanc) Then
1190                         bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo Terminar
1200                      End If
1210                      If Not oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then
1220                         bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
1230                      End If
                          If lbEsUnEPSdesdeAgregarCE = True Then                             'es un EPS
                             oDoFacturacionServicioDesp.IdOrden = oRsTmp.Fields!IdOrden
                             oDoFacturacionServicioDesp.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                             If Not oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then
                                 bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
                             End If
                             oDoFactOrdenServ.IdOrden = oRsTmp.Fields!IdOrden
                             oDoFactOrdenServ.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                             If Not oFactOrdenesServicio.Eliminar(oDoFactOrdenServ) Then
                                bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
                             End If
                          End If
1240                   End If
1250                   oRsTmp.MoveNext
1260             Loop
1270          Else
1280             Set oRsTmp = mo_ReglasFacturacion.FactOrdenServicioFiltraPorIdCuenta(oDoAtencion.idCuentaAtencion, oConexion)
1290             If oRsTmp.RecordCount > 0 Then
1300                  oRsTmp.MoveFirst
1310                  Do While Not oRsTmp.EOF
1320                        If oRsTmp.Fields!IdPuntoCarga = 6 Then    'admision consulta externa
1330                              oDoFactOrdenServ.IdOrden = oRsTmp.Fields!IdOrden
1340                              oDoFacturacionServicioFinanc.IdOrden = oRsTmp.Fields!IdOrden
1350                              oDoFacturacionServicioDesp.IdOrden = oRsTmp.Fields!IdOrden
1360                              If Not oFacturacionServicioFinan.Eliminar(oDoFacturacionServicioFinanc) Then
1370                                 bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo Terminar
1380                              End If
1390                              If Not oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then
1400                                 bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
1410                              End If
                            
1420                        End If
1430                        oRsTmp.MoveNext
1440                  Loop
1450             Else
1460                  Set oRsTmp = mo_ReglasFacturacion.FactOrdenServicioSeleccionarPorIdCuenta(oDoAtencion.idCuentaAtencion)
1470                  If oRsTmp.RecordCount > 0 Then
1480                     oDoFactOrdenServ.IdOrden = oRsTmp.Fields!IdOrden
1490                  End If
1500             End If
1510          End If
1520          oRsTmp.Close
1530          Set oRsTmp = Nothing
1540          If lnIdTablaLISTBARITEMS <> 103 Then    '103=atencion pacientes ce
1550              If ml_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oDoAtencion.IdFormaPago, oConexion) = True Then
                     '***********Contado ***********
1560                  With oDOFactOrdenServicioPagos
1570                      .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL()
1580                      .IdEstadoFacturacion = 1   'Registrado
1590                      .IdOrden = oDoFactOrdenServ.IdOrden
1600                      .idUsuario = oDoAtencion.IdUsuarioAuditoria
1610                      .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1620                  End With
1630                  If Not oFactServicioPagos.Insertar(oDOFactOrdenServicioPagos) Then
1640                      bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
1650                  End If
1660              End If
                  lnImporteTotal = 0
1670              For Each oDOFacturacionServicio In mo_FacturacionServicios

                      If oDoAtencion.EpsPorcentaje > 0 Then
                        oRsTmp2.Filter = "idFuenteFinanciamiento=" & oDoAtencion.idFuenteFinanciamiento & _
                              " and idproducto=" & oDOFacturacionServicio.idProducto
                        If oRsTmp2.RecordCount > 0 Then
                            lnMontoCPTfijoEnsoles = oRsTmp2!Importe
                        End If
                      End If


1680                  lnPrecioUnitario = 0
1690                  oDoCatalogoServicioHosp.PrecioUnitario = 0
1700                  If oDOFacturacionServicio.idProducto > 0 Then
                          'Despachos
1710                      lnPrecioUnitario = 0
1720                      oDoCatalogoServicioHosp.PrecioUnitario = 0
1730                      Set oDoCatalogoServicioHosp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarPorId(oDOFacturacionServicio.idProducto, oDoAtencion.IdFormaPago)
1740                      If oDoCatalogoServicioHosp.PrecioUnitario = 0 And oDoCatalogoServicioHosp.SeUsaSinPrecio = False Then
1750                          lcCodigoNombre = oBuscaCodigoNombre.DevuelveNombreMedicamentoOServicioSegunId(oDOFacturacionServicio.idProducto, sghServicio)
1760                          ms_MensajeError = "Tiene problemas con el ID SERVICIO: " & lcCodigoNombre & Chr(13) & "para el PRODUCTO/PLAN: " & oDoAtencion.IdFormaPago & Chr(13) & "...consulte al ADMINISTRADOR DEL SISTEMA...."
1770                          bProcesoOK = False: GoTo Terminar
1780                      Else
1790                          lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
1800                          With oDoFacturacionServicioDesp
1810                              .cantidad = oDOFacturacionServicio.cantidad
1820                              .IdOrden = oDoFactOrdenServ.IdOrden
1830                              .idProducto = oDOFacturacionServicio.idProducto
1840                              .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1850                              .Precio = lnPrecioUnitario
1860                              .total = oDOFacturacionServicio.cantidad * lnPrecioUnitario
                                   'SCCQ 26-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
                                   'HRA 09/12/2020 Cambio 46 Inicio
                                   '.IdFinan = oDoAtencion.IdFormaPago
                                   '.CodPres = mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion1
                                   'HRA 09/12/2020 Cambio 46 Fin
                                   'SCCQ 26-04-2021 Cambio 64 Fin (Reversion Cambio 46)
                                 
1870                          End With
1880                          If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then
1890                                 bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
1900                          End If
1910                          Set oDoCatalogoServicioHosp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarPorId(oDOFacturacionServicio.idProducto, oDoAtencion.IdFormaPago)
1920                          If ml_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oDoAtencion.IdFormaPago, oConexion) = True Then
                                 '***********Contado ***********
1930                             With oDoFacturacionServicioPagos
1940                                 .cantidad = oDOFacturacionServicio.cantidad
1950                                 .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
1960                                 .idProducto = oDOFacturacionServicio.idProducto
1970                                 .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1980                                 .Precio = lnPrecioUnitario
1990                                 .total = oDOFacturacionServicio.cantidad * lnPrecioUnitario
2000                             End With
2010                             If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
2020                                 bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
2030                             End If
2040                          Else
                                 '***********Sis,Soat,exoneracion, convenio ***********
2050                             With oDoFacturacionServicioFinanc
2060                                 .CantidadFinanciada = oDOFacturacionServicio.cantidad
2070                                 .FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
2080                                 .IdOrden = oDoFactOrdenServ.IdOrden
2090                                 .idProducto = oDOFacturacionServicio.idProducto
2100                                 .IdTipoFinanciamiento = oDoAtencion.IdFormaPago
2110                                 .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2120                                 .IdUsuarioAutoriza = oDoAtencion.IdUsuarioAuditoria
2130                                 .PrecioFinanciado = lnPrecioUnitario
2140                                 .TotalFinanciado = oDOFacturacionServicio.cantidad * lnPrecioUnitario
                                     .idFuenteFinanciamiento = oDoAtencion.idFuenteFinanciamiento
                                     'SCCQ 26-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
                                     'HRA 09/12/2020 Cambio 46 Inicio
2150                                 '.idFuenteFinanciamiento = oDoAtencion.IdFormaPago
                                     '.CodPres = mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion1
                                     'HRA 09/12/2020 Cambio 46 Fin
                                    'SCCQ 26-04-2021 Cambio 64 Fin (Reversion Cambio 46)
2160                             End With
2170                             If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then
2180                                 bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo Terminar
2190                             End If
                                 lnImporteTotal = lnImporteTotal + oDoFacturacionServicioFinanc.TotalFinanciado
2200                          End If
2210                      End If
2220                  End If
2230              Next
2240          End If
2250      End If
          'Actualiza Paquete (si fuese el Caso)
2260      lbYaSeTransfirioHCdeUnServicioAotro = False
2270      If DOFacturacionPaquetesUltimo.IdComprobantePago > 0 And DOFacturacionPaquetesUltimo.IdComprobantePago <> DOFacturacionPaquetesAnterior.IdComprobantePago Then
2280          If FacturacionPaquetesActualizaXidComprobIdOrdenIdProducto(DOFacturacionPaquetesUltimo.idProducto, _
                                    oDoAtencion.idCuentaAtencion, DOFacturacionPaquetesUltimo.IdComprobantePago, _
                                    DOFacturacionPaquetesUltimo.idOrdenPago, True, oConexion) = False Then
2290             bProcesoOK = False:  GoTo Terminar
2300          End If
              'actualiza IdPaciente en Boleta
2310          If CajaComprobantesPagoActualizaXpacienteYComprobante(oDoAtencion.idPaciente, _
                                              DOFacturacionPaquetesUltimo.IdComprobantePago, False, oConexion) = False Then
2320             bProcesoOK = False:  GoTo Terminar
2330          End If
              'Agrega MOVIMIENTO DE ARCHIVO si no esta la Historia ubicada en Archivo Clinico}
2340          Set oRsTmp1 = mo_ReglasArchivoClinico.MovimientosHistoriaClinicaSeleccionaUltimoMovimientoPorPaciente(oDoAtencion.idPaciente)
2350          If oRsTmp1.RecordCount > 0 Then
2360              If oRsTmp1.Fields!idServicioDestino <> oBuscaCodigoNombre.ParametrosIdServicioArchivoClinico Then
2370                  oDOMovimientoHistoriaClinica.idPaciente = oDoAtencion.idPaciente
2380                  oDOMovimientoHistoriaClinica.FechaMovimiento = lcBuscaParametro.RetornaFechaHoraServidorSQL()
2390                  oDOMovimientoHistoriaClinica.IdMotivo = 7    'Transferencia
2400                  oDOMovimientoHistoriaClinica.IdServicioOrigen = oRsTmp1.Fields!idServicioDestino
2410                  oDOMovimientoHistoriaClinica.idServicioDestino = oDoAtencion.idServicioIngreso
2420                  oDOMovimientoHistoriaClinica.IdEmpleadoArchivo = oDoAtencion.IdUsuarioAuditoria
2430                  oDOMovimientoHistoriaClinica.IdEmpleadoTransporte = oDoAtencion.IdUsuarioAuditoria
2440                  oDOMovimientoHistoriaClinica.IdEmpleadoRecepcion = oDoAtencion.IdUsuarioAuditoria
2450                  oDOMovimientoHistoriaClinica.IdGrupoMovimiento = 1
2460                  oDOMovimientoHistoriaClinica.idAtencion = oDoAtencion.idAtencion
2470                  If Not oMovimientosHistoriaClinica.Insertar(oDOMovimientoHistoriaClinica) Then
2480                      bProcesoOK = False: ms_MensajeError = oMovimientosHistoriaClinica.MensajeError: GoTo Terminar
2490                  End If
2500                  lbYaSeTransfirioHCdeUnServicioAotro = True
2510              End If
2520          End If
2530          oRsTmp1.Close
2540      End If
          'Blanquea Dato de "AtencionId" del anterior Boleta grabada por error
2550      If DOFacturacionPaquetesAnterior.IdComprobantePago > 0 And DOFacturacionPaquetesUltimo.IdComprobantePago <> DOFacturacionPaquetesAnterior.IdComprobantePago Then
2560          If FacturacionPaquetesActualizaXidComprobIdOrdenIdProducto(DOFacturacionPaquetesAnterior.idProducto, _
                                       0, DOFacturacionPaquetesAnterior.IdComprobantePago, _
                                       DOFacturacionPaquetesAnterior.idOrdenPago, True, oConexion) = False Then
2570             bProcesoOK = False:  GoTo Terminar
2580          End If
              'actualiza IdPaciente en Boleta
2590          If CajaComprobantesPagoActualizaXpacienteYComprobante(0, _
                                    DOFacturacionPaquetesAnterior.IdComprobantePago, True, oConexion) = False Then
2600             bProcesoOK = False:  GoTo Terminar
2610          End If
2620      End If
          'datos adicionales
2630      mo_DoAtencionDatosAdicionales.idAtencion = oDoAtencion.idAtencion
2640      If Not oAtencionesDatosAdicionales.modificar(mo_DoAtencionDatosAdicionales) Then
2650          bProcesoOK = False: ms_MensajeError = oAtencionesDatosAdicionales.MensajeError: GoTo Terminar
2660      End If
2670      If mo_DoAtencionDatosAdicionales.idSiaSis > 0 Then
2680         Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "M", oDoAtencion.idCuentaAtencion, "FUA", oConexion, 1345, lcNombrePc, "Cod.Prestacion FUA " & mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion & " Cta: " & oDoAtencion.idCuentaAtencion)
2690      End If
      
          mo_ReglasArchivoClinico.HistoriasClinicasActualizaFmovimiento oConexion, oDoAtencion.idPaciente
          '
          AsignaPorcentajeEPS oConexion, oDoAtencion.idCuentaAtencion, oDoAtencion.EpsPorcentaje, lnIdTablaLISTBARITEMS, _
                               sghModificar, oDoFactOrdenServ, oFactOrdenesServicio, oDoFacturacionServicioDesp, _
                               oFacturacionServicioDespacho, oDOFactOrdenServicioPagos, oFactServicioPagos, _
                               oDoFacturacionServicioPagos, oFacturacionServicioPagos, CDate(lcBuscaParametro.RetornaFechaHoraServidorSQL()), _
                               oDoAtencion.idServicioIngreso, oDoAtencion.idPaciente, _
                               oDoAtencion.IdUsuarioAuditoria, oDoFactOrdenServ.IdOrden, lnImporteTotal, _
                               sghPuntosCargaBasicos.sghPtoCargaAdmisionCE, lnIdOrdenPago, lnMontoCPTfijoEnsoles
          
          '
Terminar:
2700      If bProcesoOK Then
2710          AdmisionCEModificar = True
2720          oConexion.CommitTrans
2730      Else
2740          AdmisionCEModificar = False
2750          oConexion.RollbackTrans
2760          ms_MensajeError = ms_MensajeError & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                                     "Function AdmisionCEModificar", "reglasAdmision.cls") 'debb-02/05/2016
                             
2770      End If
                  
2780      oConexion.Close
2790      Set oConexion = Nothing
2800      Set oPaciente = Nothing
2810      Set oCuentaAtencion = Nothing
2820      Set oAtencion = Nothing
2830      Set oDOHistoriaSolicitada = Nothing
2840      Set oHistoriaSolicitada = Nothing
2850      Set oCita = Nothing
2860      Set oHistoria = Nothing
2870      Set oAtencionesDx = Nothing
2880      Set oFacturacionServicios = Nothing
2890      Set oCatalogoServicios = Nothing
2900      Set oFacturacionServicioDespacho = Nothing
2910      Set oFacturacionServicioPagos = Nothing
2920      Set oFacturacionServicioFinan = Nothing
2930      Set oFactServicioPagos = Nothing
2940      Set oDoFacturacionServicioPagos = Nothing
2950      Set oDoFacturacionServicioFinanc = Nothing
2960      Set oDoFacturacionServicioDesp = Nothing
2970      Set oDOFacturacionServicio = Nothing
2980      Set oDoFactOrdenServ = Nothing
2990      Set oDOFactOrdenServicioPagos = Nothing
3000      Set oBuscaCodigoNombre = Nothing
3010      Set oDoCatalogoServicioHosp = Nothing
3020      Set ml_ReglasFacturacion = Nothing
3030      Set oRsTmp1 = Nothing
3040      Set oRsTmp = Nothing
3050      Set oDOMovimientoHistoriaClinica = Nothing
3060      Set oMovimientosHistoriaClinica = Nothing
3070      Set mo_ReglasArchivoClinico = Nothing
3080      Set oSunasaPacientesHistoricos = Nothing
3090      Set oAtencionesDatosAdicionales = Nothing
          Set oFactOrdenesServicio = Nothing
          Set oRsTmp2 = Nothing
End Function

Function AdmisionCEAnular(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, oDoPaciente As DOPaciente, _
                          oDOCita As doCita, oDoHistoria As DOHistoriaClinica, _
                          lIdTipoNumeracionAnterior As sghTipoNumeracionDeNroHistoria, oDiagnosticos As Collection, _
                          oProcedimientos As Collection, oExamenes As Collection, lIdHistoriaAnterior As Long, _
                          mo_FacturacionServicios As Collection, mo_FacturacionServiciosAEliminar As Collection, _
                          DOFacturacionPaquetesAnterior As DOFacturacionPaquetes, lnIdTablaLISTBARITEMS As Long, _
                          lcNombrePc As String, lcObservacionesAuditoria As String, oEpisodioClinico As EpisodioClinico) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim oDOHistoriaSolicitada As New DOHistoriaSolicitada
Dim oHistoriaSolicitada As New HistoriasSolicitadas
Dim oCita As New Citas
Dim bProcesoOK As Boolean
Dim oHistoria As New HistoriasClinicas
Dim oAtencionesDx As New AtencionesDiagnosticos
Dim oFacturacionServicios As New FacturacionServicios
Dim oCatalogoServicios As New CatalogoServicios
Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
Dim oFacturacionServicioPagos As New FacturacionServicioPagos
Dim oFacturacionServicioFinan As New FacturacionServicioFin
Dim oFactServicioPagos As New FactOrdenServicioPagos
Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
Dim oDOFacturacionServicio As DOFacturacionServicios
Dim oDoFactOrdenServ As New DoFactOrdenServ
Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
Dim oHistoriasSolicitadas As New HistoriasSolicitadas
Dim oDOMovimientoHistoriaClinica As New DOMovimientoHistoriaClinica
Dim oMovimientosHistoriaClinica As New MovimientosHistoriaClinica
Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim oDOProgramacionMedica As New DOProgramacionMedica, oProgramacionMedica As New ProgramacionMedica
Dim oDoAtencionesEpisodiosCab As New DoAtencionesEpisodiosCab, oAtencionesEpisodiosCab As New AtencionesEpisodiosCab
Dim oDoAtencionesEpisodiosDet As New DoAtencionesEpisodiosDet, oAtencionesEpisodiosDet As New AtencionesEpisodiosDet
Dim oRsTmp As New Recordset
Dim oRsTmp1 As New Recordset
Dim lcSql As String

    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    Set oHistoriaSolicitada.Conexion = oConexion
    Set oCita.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oAtencionesDx.Conexion = oConexion
    Set oFacturacionServicios.Conexion = oConexion
    Set oCatalogoServicios.Conexion = oConexion
    Set oFacturacionServicioPagos.Conexion = oConexion
    Set oFacturacionServicioFinan.Conexion = oConexion
    Set oFacturacionServicioDespacho.Conexion = oConexion
    Set oFactServicioPagos.Conexion = oConexion
    Set oMovimientosHistoriaClinica.Conexion = oConexion
    Set oProgramacionMedica.Conexion = oConexion
    Set oAtencionesEpisodiosCab.Conexion = oConexion
    Set oAtencionesEpisodiosDet.Conexion = oConexion
    
    oConexion.BeginTrans
    bProcesoOK = True
    
    
    'Modificar atencion
    If Not oAtencion.modificar(oDoAtencion) Then
        bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "E", oDoAtencion.idAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, lcObservacionesAuditoria)   'ListBarItems.idListItem
    '
    oDOCuentaAtencion.FechaCreacion = CDate(oDoAtencion.FechaIngreso & " " & oDoAtencion.HoraIngreso)
    oDOCuentaAtencion.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
    If Not oCuentaAtencion.modificar(oDOCuentaAtencion) Then
        bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
    End If
    'Eliminar solicitud historias
    If Not oHistoriaSolicitada.EliminarPorCita(oDOCita) Then
        bProcesoOK = False: ms_MensajeError = oHistoriasSolicitadas.MensajeError: GoTo Terminar
    End If
    '
    If Not oCita.Eliminar(oDOCita) Then
        bProcesoOK = False: ms_MensajeError = oCita.MensajeError: GoTo Terminar
    End If
    '
    '*************************Servicios***********************
    'elimina Codigo Servicio (cabecera/detalle) por Estancia de la Cuenta de Atencion}
    Set oRsTmp = mo_ReglasFacturacion.FactOrdenServicioPagosPorIdCuentaAtencion(oDoAtencion.idCuentaAtencion)
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
             If oRsTmp.Fields!IdPuntoCarga = 6 Then    'admision consulta externa
                oDoFactOrdenServ.IdOrden = oRsTmp.Fields!IdOrden
                oDoFacturacionServicioFinanc.IdOrden = oRsTmp.Fields!IdOrden
                oDoFacturacionServicioDesp.IdOrden = oRsTmp.Fields!IdOrden
                If oRsTmp.Fields!IdEstadoFacturacion = 1 Then
                     oDOFactOrdenServicioPagos.IdOrden = oRsTmp.Fields!IdOrden
                     oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp.Fields!idOrdenPago
                     oDoFacturacionServicioPagos.idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
                     If Not oFacturacionServicioPagos.Eliminar(oDoFacturacionServicioPagos) Then
                        bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
                     End If
                     If Not oFactServicioPagos.Eliminar(oDOFactOrdenServicioPagos) Then
                        bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
                     End If
                End If
                If Not oFacturacionServicioFinan.Eliminar(oDoFacturacionServicioFinanc) Then
                   bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo Terminar
                End If
                If Not oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then
                   bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
                End If
             End If
             oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    'Blanquea Dato de "AtencionId" de la Boleta
    If DOFacturacionPaquetesAnterior.IdComprobantePago > 0 Then
        If FacturacionPaquetesActualizaXidComprobIdOrdenIdProducto(DOFacturacionPaquetesAnterior.idProducto, _
                                 0, DOFacturacionPaquetesAnterior.IdComprobantePago, _
                                 DOFacturacionPaquetesAnterior.idOrdenPago, False, oConexion) = False Then
           bProcesoOK = False:  GoTo Terminar
        End If
        'actualiza IdPaciente en Boleta

        If CajaComprobantesPagoActualizaXpacienteYComprobante(0, _
                              DOFacturacionPaquetesAnterior.IdComprobantePago, True, oConexion) = False Then
           bProcesoOK = False:  GoTo Terminar
        End If
        'Elimina MOVIMIENTO DE ARCHIVO (si corresponde al Servicio de Ingreso}
        Set oRsTmp1 = mo_ReglasArchivoClinico.MovimientosHistoriaClinicaSeleccionaUltimoMovimientoPorPaciente(oDoAtencion.idPaciente)
        If oRsTmp1.RecordCount > 0 Then
            If oRsTmp1.Fields!idServicioDestino = oDoAtencion.idServicioIngreso Then
                oDOMovimientoHistoriaClinica.IdMovimiento = oRsTmp1.Fields!IdMovimiento
                oDOMovimientoHistoriaClinica.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                If Not oMovimientosHistoriaClinica.Eliminar(oDOMovimientoHistoriaClinica) Then
                    bProcesoOK = False: ms_MensajeError = oMovimientosHistoriaClinica.MensajeError: GoTo Terminar
                End If
            End If
        End If
        oRsTmp1.Close
    End If
    'Cupo Adicional
    If oDOCita.EsCitaAdicional = True Then
       oDOProgramacionMedica.idProgramacion = oDOCita.idProgramacion
       If Not oProgramacionMedica.SeleccionarPorId(oDOProgramacionMedica) Then
            bProcesoOK = False: ms_MensajeError = oProgramacionMedica.MensajeError: GoTo Terminar
       End If
'       If oDOProgramacionMedica.horaFin = oDOCita.horaFin Then
'            oDOProgramacionMedica.IdUsuarioAuditoria = oDOCita.IdUsuarioAuditoria
'            oDOProgramacionMedica.horaFin = oDOCita.horaInicio
'            If Not oProgramacionMedica.modificar(oDOProgramacionMedica) Then
'                 bProcesoOK = False: ms_MensajeError = oProgramacionMedica.MensajeError: GoTo Terminar
'            End If
'       Else
'            bProcesoOK = False: ms_MensajeError = "Sólo se puede eliminar la última CITA ADICIONAL": GoTo Terminar
'       End If
        'mgaray20141021
       oDOProgramacionMedica.IdUsuarioAuditoria = oDOCita.IdUsuarioAuditoria
       If oDOProgramacionMedica.horaFin = oDOCita.horaFin Then
            oDOProgramacionMedica.horaFin = oDOCita.HoraInicio
'       Else
'            bProcesoOK = False: ms_MensajeError = "Sólo se puede eliminar la última CITA ADICIONAL": GoTo Terminar
       End If
       If Not oProgramacionMedica.ModificarHoraFinPorCitaAdicional(oDOProgramacionMedica) Then
             bProcesoOK = False: ms_MensajeError = oProgramacionMedica.MensajeError: GoTo Terminar
        End If
    End If
    '
    mo_ReglasArchivoClinico.HistoriasClinicasActualizaFmovimiento oConexion, oDoAtencion.idPaciente
    'Episodio Clinico
    ms_MensajeError = EpisodioClinicoActualizar(oEpisodioClinico, oDoAtencionesEpisodiosCab, oAtencionesEpisodiosCab, _
                                                oDoAtencionesEpisodiosDet, oAtencionesEpisodiosDet, oDoAtencion, _
                                                oConexion, sghEliminar)
    If ms_MensajeError <> "" Then
       bProcesoOK = False: GoTo Terminar
    End If
    '
Terminar:
    If bProcesoOK Then
        AdmisionCEAnular = True
        oConexion.CommitTrans
    Else
        AdmisionCEAnular = False
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
    Set oPaciente = Nothing
    Set oCuentaAtencion = Nothing
    Set oAtencion = Nothing
    Set oDOHistoriaSolicitada = Nothing
    Set oHistoriaSolicitada = Nothing
    Set oCita = Nothing
    Set oHistoria = Nothing
    Set oAtencionesDx = Nothing
    Set oFacturacionServicios = Nothing
    Set oCatalogoServicios = Nothing
    Set oFacturacionServicioDespacho = Nothing
    Set oFacturacionServicioPagos = Nothing
    Set oFacturacionServicioFinan = Nothing
    Set oFactServicioPagos = Nothing
    Set oDoFacturacionServicioPagos = Nothing
    Set oDoFacturacionServicioFinanc = Nothing
    Set oDoFacturacionServicioDesp = Nothing
    Set oDOFacturacionServicio = Nothing
    Set oDoFactOrdenServ = Nothing
    Set oDOFactOrdenServicioPagos = Nothing
    Set oDOMovimientoHistoriaClinica = Nothing
    Set oMovimientosHistoriaClinica = Nothing
    Set mo_ReglasArchivoClinico = Nothing
    Set oDOProgramacionMedica = Nothing
    Set oProgramacionMedica = Nothing
    Set oDoAtencionesEpisodiosCab = Nothing
    Set oAtencionesEpisodiosCab = Nothing
    Set oDoAtencionesEpisodiosDet = Nothing
    Set oAtencionesEpisodiosDet = Nothing
    
End Function

Function AdmisionCEEliminar(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, oDOCita As doCita, lnIdTablaLISTBARITEMS As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim oCita As New Citas
Dim oAtencionesDx As New AtencionesDiagnosticos
Dim oFactOrdenesServicio As New FactOrdenesServicio
Dim oFacturacionServicios As New FacturacionServicios
Dim oHistoriasSolicitadas As New HistoriasSolicitadas
Dim rsTmp As New ADODB.Recordset
Dim bProcesoOK As Boolean

    oConexion.Open sighentidades.CadenaConexion
    
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oCita.Conexion = oConexion
    Set oAtencionesDx.Conexion = oConexion
    Set oFactOrdenesServicio.Conexion = oConexion
    Set oFacturacionServicios.Conexion = oConexion
    Set oHistoriasSolicitadas.Conexion = oConexion
    
    oConexion.BeginTrans
    bProcesoOK = True
    
    'Eliminar Cita
    If Not oCita.Eliminar(oDOCita) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    'Eliminar Facturacion de Servicios
    If Not oFacturacionServicios.EliminarServiciosDeLaCuenta(oDoAtencion.idAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    'Eliminar orden de servicio
    If Not oFactOrdenesServicio.EliminarPorIdAtencion(oDoAtencion.idAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    'Eliminar diagnosticos
    If Not oAtencionesDx.EliminarDiagnosticosPorAtencion(oDoAtencion.idAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    
    'Eliminar atencion
    If Not oAtencion.Eliminar(oDoAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    'Eliminar solicitud historias
    If Not oHistoriasSolicitadas.EliminarPorCita(oDOCita) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    'Eliminar cuenta atención   242658   idcuentaatencion=188581...24 atenc
    Set rsTmp = oAtencion.SeleccionarXidCuentaAtencion(oDOCuentaAtencion.idCuentaAtencion)
    If rsTmp.RecordCount = 0 Then
        If Not oCuentaAtencion.Eliminar(oDOCuentaAtencion) Then
            bProcesoOK = False: GoTo Terminar
        End If
    End If
    rsTmp.Close
    
Terminar:
    If bProcesoOK Then
        AdmisionCEEliminar = True
        oConexion.CommitTrans
    Else
        AdmisionCEEliminar = False
        ms_MensajeError = oCuentaAtencion.MensajeError + Chr(13) + oAtencion.MensajeError + Chr(13) + oPaciente.MensajeError + Chr(13) + oCita.MensajeError
        oConexion.RollbackTrans
    End If
            
    oConexion.Close
    Set oConexion = Nothing
    Set oPaciente = Nothing
    Set oCuentaAtencion = Nothing
    Set oAtencion = Nothing
    Set oCita = Nothing
    Set oAtencionesDx = Nothing
    Set oFactOrdenesServicio = Nothing
    Set oFacturacionServicios = Nothing
    Set oHistoriasSolicitadas = Nothing
    Set rsTmp = Nothing
End Function

Function AdmisionEmergAgregar(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, oDoPaciente As DOPaciente, oDoHistoria As DOHistoriaClinica, lIdTipoNumeracionAnterior As sghTipoNumeracionDeNroHistoria, oDiagnosticos As Collection, oProcedimientos As Collection, oExamenes As Collection, lIdHistoriaAnterior As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim bProcesoOK As Boolean
Dim oDOHistoriaSolicitada As New DOHistoriaSolicitada
Dim oHistoriaSolicitada As New HistoriasSolicitadas
Dim bPacienteNuevo As Boolean
Dim oHistoria As New HistoriasClinicas
Dim oCama As New Camas
Dim oAtencionesDx As New AtencionesDiagnosticos
Dim oFacturacionServicios As New FacturacionServicios


    oConexion.Open sighentidades.CadenaConexion
    
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    Set oHistoriaSolicitada.Conexion = oConexion
    Set oCama.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oAtencionesDx.Conexion = oConexion
    Set oFacturacionServicios.Conexion = oConexion
    
    oConexion.BeginTrans
    bProcesoOK = True
    bPacienteNuevo = False
    
    'Paciente nuevo
    If oDoPaciente.idPaciente = 0 Then
        bPacienteNuevo = True
        
        If oDoHistoria.IdTipoNumeracion <> sghHistoriaDefinitivaManual Then
            oDoHistoria.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoPaciente.IdTipoNumeracion)
        End If
        
        If oDoHistoria.NroHistoriaClinica = 0 Then
            If oDoHistoria.IdTipoNumeracion = sghHistoriaDefinitivaReciclada Then
                Trace "No se pudo generar el nro de de historia clínica reciclada", "AdmisionEmergAgregar"
                bProcesoOK = False: GoTo Terminar
            End If
        End If
        
        'Agregar paciente
        oDoPaciente.IdTipoNumeracion = oDoHistoria.IdTipoNumeracion
        oDoPaciente.NroHistoriaClinica = oDoHistoria.NroHistoriaClinica
        
        If Not oPaciente.Insertar(oDoPaciente) Then
            bProcesoOK = False: GoTo Terminar
        End If

        
        'Agregar historia clinica (Si es historia definitiva)
        'En caso de emergencia nuca va a entrar a esta condicion, alva que se modifiquen la reglas
        Select Case oDoHistoria.IdTipoNumeracion
        Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
            If Not oHistoria.Insertar(oDoHistoria) Then
                bProcesoOK = False: GoTo Terminar
            End If
        End Select
        
    Else
    'Paciente antiguo
        
        'Agregar historia clinica (Si es tenia historia temporal y ahora tiene historia definitiva)
        Select Case lIdTipoNumeracionAnterior
        Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
            
        Case Else
            
            'Si era temporal y ahora es definitiva => Genera y agrega la historia
            Select Case oDoHistoria.IdTipoNumeracion
            Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
                
                If oDoHistoria.IdTipoNumeracion <> sghHistoriaDefinitivaManual Then
                    oDoHistoria.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoPaciente.IdTipoNumeracion)
                End If
                
                If oDoHistoria.NroHistoriaClinica = 0 Then
                    If oDoHistoria.IdTipoNumeracion = sghHistoriaDefinitivaReciclada Then
                        Trace "No se pudo generar el nro de de historia clínica reciclada", "AdmisionEmergAgregar"
                        bProcesoOK = False: GoTo Terminar
                    End If
                End If
                
                oDoHistoria.NroHistoriaClinicaAnterior = lIdHistoriaAnterior
                oDoHistoria.IdTipoNumeracionAnterior = lIdTipoNumeracionAnterior
                
                oDoPaciente.NroHistoriaClinica = oDoHistoria.NroHistoriaClinica
                
                If Not oHistoria.Insertar(oDoHistoria) Then
                    bProcesoOK = False: GoTo Terminar
                End If
            End Select
        End Select
        
        If Not oPaciente.modificar(oDoPaciente, True) Then
            bProcesoOK = False: GoTo Terminar
        End If
        
    End If
    
    'Agregar solicitud de preestamo
    Select Case oDoHistoria.IdTipoNumeracion
    Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
        InicializarPrestamoAutomaticoEmerg oDoAtencion, oDoPaciente, oDOHistoriaSolicitada, bPacienteNuevo
        If Not oHistoriaSolicitada.Insertar(oDOHistoriaSolicitada) Then
            bProcesoOK = False: GoTo Terminar
        End If
    End Select
    
    'Agregar atencion
    oDoAtencion.idPaciente = oDoPaciente.idPaciente
    If Not oAtencion.Insertar(oDoAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    'Agregar diagnosticos
    If Not oAtencionesDx.ActualizarDiagnosticosAtencion(oDiagnosticos, oDoAtencion.idAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    'Agregar cuenta atención
    oDOCuentaAtencion.idPaciente = oDoPaciente.idPaciente
    If Not oCuentaAtencion.Insertar(oDOCuentaAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If

Terminar:
    If bProcesoOK Then
        AdmisionEmergAgregar = True
        oConexion.CommitTrans
    Else
        AdmisionEmergAgregar = False
        ms_MensajeError = oHistoriaSolicitada.MensajeError + Chr(13) + oCuentaAtencion.MensajeError + Chr(13) + oAtencion.MensajeError + Chr(13) + oPaciente.MensajeError + Chr(13) + oHistoria.MensajeError
        oConexion.RollbackTrans
    End If
            
    oConexion.Close

End Function
Function AdmisionEmergModificar(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, oDoPaciente As DOPaciente, oDoHistoria As DOHistoriaClinica, lIdTipoNumeracionAnterior As sghTipoNumeracionDeNroHistoria, oDiagnosticos As Collection, oProcedimientos As Collection, oExamenes As Collection, lIdHistoriaAnterior As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim oDOHistoriaSolicitada As New DOHistoriaSolicitada
Dim oHistoriaSolicitada As New HistoriasSolicitadas
Dim oCita As New Citas
Dim bProcesoOK As Boolean
Dim oHistoria As New HistoriasClinicas
Dim oCama As New Camas
Dim oAtencionesDx As New AtencionesDiagnosticos

    oConexion.Open sighentidades.CadenaConexion
    
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    Set oHistoriaSolicitada.Conexion = oConexion
    Set oCama.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oAtencionesDx.Conexion = oConexion
    
    oConexion.BeginTrans
    bProcesoOK = True
    
    'Agregar historia clinica (Si es tenia historia temporal y ahora tiene historia definitiva)
    Select Case lIdTipoNumeracionAnterior
    Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
        
    Case Else
        
        'Si era temporal y ahora es definitiva => Genera y agrega la historia
        Select Case oDoHistoria.IdTipoNumeracion
        Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
            
            If oDoHistoria.IdTipoNumeracion <> sghHistoriaDefinitivaManual Then
                oDoHistoria.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoPaciente.IdTipoNumeracion)
            End If
            
            If oDoHistoria.NroHistoriaClinica = 0 Then
                If oDoHistoria.IdTipoNumeracion = sghHistoriaDefinitivaReciclada Then
                    Trace "No se pudo generar el nro de de historia clínica reciclada", "AdmisionEmergModificar"
                    bProcesoOK = False: GoTo Terminar
                End If
            End If
            
            oDoHistoria.NroHistoriaClinicaAnterior = lIdHistoriaAnterior
            oDoHistoria.IdTipoNumeracionAnterior = lIdTipoNumeracionAnterior
            
            oDoPaciente.NroHistoriaClinica = oDoHistoria.NroHistoriaClinica
            
            If Not oHistoria.Insertar(oDoHistoria) Then
                bProcesoOK = False: GoTo Terminar
            End If
        End Select
    End Select
    
    'Modificar paciente
    If Not oPaciente.modificar(oDoPaciente, True) Then
        bProcesoOK = False: GoTo Terminar
    End If

    'Modificar solicitud de prestamo
    Select Case oDoHistoria.IdTipoNumeracion
    Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
        InicializarPrestamoAutomaticoEmerg oDoAtencion, oDoPaciente, oDOHistoriaSolicitada, False
        If Not oHistoriaSolicitada.modificar(oDOHistoriaSolicitada) Then
            bProcesoOK = False: GoTo Terminar
        End If
    End Select
    
    'Modificar atencion
    oDoAtencion.idPaciente = oDoPaciente.idPaciente
    If Not oAtencion.modificar(oDoAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    'Agregar diagnosticos
    If Not oAtencionesDx.ActualizarDiagnosticosAtencion(oDiagnosticos, oDoAtencion.idAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    'Modificar cuenta atención
    oDOCuentaAtencion.idPaciente = oDoPaciente.idPaciente
    If Not oCuentaAtencion.modificar(oDOCuentaAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    'Agregar procedimientos
    'If Not oAtencionesProc.ActualizarProcedimientosAtencion(oProcedimientos, oDOCuentaAtencion.IdCuentaAtencion) Then
    '    bProcesoOK = False: GoTo Terminar
    'End If
        
    'Agregar examenes
    'If Not oAtencionesExamenes.ActualizarExamenesAtencion(oExamenes, oDOCuentaAtencion.IdCuentaAtencion) Then
    '    bProcesoOK = False: GoTo Terminar
    'End If
    
    
Terminar:
    
    If bProcesoOK Then
        AdmisionEmergModificar = True
        oConexion.CommitTrans
    Else
        AdmisionEmergModificar = False
        ms_MensajeError = oHistoriaSolicitada.MensajeError + Chr(13) + oCuentaAtencion.MensajeError + Chr(13) + oAtencion.MensajeError + Chr(13) + oPaciente.MensajeError + Chr(13) + oHistoria.MensajeError
        oConexion.RollbackTrans
    End If
            
    oConexion.Close

End Function
Function AdmisionEmergEliminar(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim oCama As New Camas
Dim bProcesoOK As Boolean
Dim oAtencionesDx As New AtencionesDiagnosticos


    oConexion.Open sighentidades.CadenaConexion
    
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oCama.Conexion = oConexion
    Set oAtencionesDx.Conexion = oConexion
    
    oConexion.BeginTrans
    bProcesoOK = True
    
    
    
    'Eliminar cuenta atención
    If Not oCuentaAtencion.Eliminar(oDOCuentaAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    'Eliminar diagnosticos
    If Not oAtencionesDx.EliminarDiagnosticosPorAtencion(oDoAtencion.idAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    'Eliminar atencion
    If Not oAtencion.Eliminar(oDoAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
Terminar:
    If bProcesoOK Then
        AdmisionEmergEliminar = True
        oConexion.CommitTrans
    Else
        AdmisionEmergEliminar = False
        oConexion.RollbackTrans
    End If
            
    oConexion.Close
    Set oConexion = Nothing
    Set oPaciente = Nothing
    Set oCuentaAtencion = Nothing
    Set oAtencion = Nothing
    Set oCama = Nothing
    Set oAtencionesDx = Nothing

End Function


Function AdmisionHospAgregar(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, oDoPaciente As DOPaciente, _
                             oDoHistoria As DOHistoriaClinica, lIdTipoNumeracionAnterior As sghTipoNumeracionDeNroHistoria, _
                             oOcupaciones As Collection, oDiagnosticos As Collection, oProcedimientos As Collection, _
                             oExamenes As Collection, oNacimientos As Collection, lIdHistoriaAnterior As Long, _
                             oDOAtencionEmergencia As DOAtencionEmergencia, oDOAtencionPadre As DOAtencion, _
                             lbPacienteNN As Boolean, lnIdTablaLISTBARITEMS As Long, lcNombrePc As String, _
                             lcObservacionesAuditoria As String, lnIdNacimientoSeleccionado As Long, _
                             oDoSunasaPacientesHistoricos As DoSunasaPacientesHistoricos, _
                             mo_DoAtencionDatosAdicionales As DoAtencionDatosAdicionales, _
                             lbElServicioEsObservacionDeEmergencia As Boolean, _
                             Optional oDoPacienteDatosAdd As DoPacienteDatosAdd = Nothing) As Boolean
      Dim oConexion As New ADODB.Connection
      Dim oPaciente As New Pacientes
      Dim oCuentaAtencion As New CuentasAtencion
      Dim oAtencion As New Atenciones
      Dim bProcesoOK As Boolean
      Dim oDOHistoriaSolicitada As New DOHistoriaSolicitada
      Dim oHistoriaSolicitada As New HistoriasSolicitadas
      Dim bPacienteNuevo As Boolean
      Dim oHistoria As New HistoriasClinicas
      Dim oCama As New Camas
      Dim oAtencionesDx As New AtencionesDiagnosticos
      Dim oEstanciaHospitalaria As New AtencionesEstanciaHosp
      Dim oAtencionDiagnostico As New AtencionesDiagnosticos
      Dim oAtencionesNacimientos As New AtencionesNacimientos
      'WCG_20060320
      Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
      Dim oFacturacionServicioPagos As New FacturacionServicioPagos
      Dim oFacturacionServicioFinan As New FacturacionServicioFin
      Dim oFactServicioPagos As New FactOrdenServicioPagos
      Dim oCatalogoServicios As New CatalogoServicios
      Dim oDOFacturacionServicios As New DOFacturacionServicios
      Dim mo_ReglasFacturacion As New ReglasFacturacion
      Dim oFactOrdenesServicio As New FactOrdenServicio
      Dim oSunasaPacientesHistoricos As New SunasaPacientesHistoricos
      Dim oAtencionesDatosAdicionales As New AtencionesDatosAdicionales
      Dim oBuscaCodigoNombre As New SIGHNegocios.ReglasComunes
      Dim oDOAtencionNacimiento As New DOAtencionNacimiento
      Dim oDoAtencionesHijoMadre As New DoAtencionesHijoMadre, oAtencionesHijoMadre As New AtencionesHijoMadre
      Dim lcCodigoNombre As String
      Dim oRsTmp As New Recordset
        
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
          
20        Set oCuentaAtencion.Conexion = oConexion
30        Set oAtencion.Conexion = oConexion
40        Set oPaciente.Conexion = oConexion
50        Set oHistoriaSolicitada.Conexion = oConexion
60        Set oCama.Conexion = oConexion
70        Set oHistoria.Conexion = oConexion
80        Set oAtencionesDx.Conexion = oConexion
90        Set oEstanciaHospitalaria.Conexion = oConexion
100       Set oAtencionDiagnostico.Conexion = oConexion
110       Set oAtencionesNacimientos.Conexion = oConexion
120       Set oCatalogoServicios.Conexion = oConexion 'WCG_20060320
130       Set oFactOrdenesServicio.Conexion = oConexion
140       Set oFacturacionServicioPagos.Conexion = oConexion
150       Set oFacturacionServicioFinan.Conexion = oConexion
160       Set oFacturacionServicioDespacho.Conexion = oConexion
170       Set oFactOrdenesServicio.Conexion = oConexion
180       Set oFactServicioPagos.Conexion = oConexion
190       Set oSunasaPacientesHistoricos.Conexion = oConexion
200       Set oAtencionesDatosAdicionales.Conexion = oConexion
210       Set oAtencionesHijoMadre.Conexion = oConexion

          
220       oConexion.BeginTrans
230       bProcesoOK = True
240       bPacienteNuevo = False
          
          'Paciente nuevo
250       If oDoPaciente.idPaciente = 0 Then
260           bPacienteNuevo = True
              
270           If oDoHistoria.IdTipoNumeracion <> sghHistoriaDefinitivaManual Then
280               oDoHistoria.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoHistoria.IdTipoNumeracion)
290           End If
              
300           If oDoHistoria.NroHistoriaClinica = 0 Then
310               If oDoHistoria.IdTipoNumeracion = sghHistoriaDefinitivaReciclada Then
320                   Trace "No se pudo generar el nro de de historia clínica reciclada", "AdmisionHospAgregar"
330                   bProcesoOK = False: ms_MensajeError = oHistoria.MensajeError: GoTo Terminar
340               End If
350           End If
              
              'Agregar paciente
360           oDoPaciente.IdTipoNumeracion = oDoHistoria.IdTipoNumeracion
370           oDoPaciente.NroHistoriaClinica = oDoHistoria.NroHistoriaClinica
              
380           If Not PacientesAgregar(oDoPaciente, oDoPacienteDatosAdd) Then
390               bProcesoOK = False: GoTo Terminar
400           End If
              
              'Agregar historia clinica (Si es historia definitiva)
410           oDoHistoria.idPaciente = oDoPaciente.idPaciente
420           oDoHistoria.NroHistoriaClinica = oDoPaciente.NroHistoriaClinica
430           Select Case oDoHistoria.IdTipoNumeracion
              Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
440               If Not oHistoria.Insertar(oDoHistoria) Then
450                   bProcesoOK = False: ms_MensajeError = oHistoria.MensajeError: GoTo Terminar
460               End If
470           End Select
          
480       Else
          'Paciente antiguo
              
              'Agregar historia clinica (Si es tenia historia temporal y ahora tiene historia definitiva)
490           Select Case lIdTipoNumeracionAnterior
              Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
                  
500           Case Else
                  
                  'Si era temporal y ahora es definitiva => Genera y agrega la historia
510               Select Case oDoHistoria.IdTipoNumeracion
                  Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
                      
520                   If oDoHistoria.IdTipoNumeracion <> sghHistoriaDefinitivaManual Then
530                       oDoHistoria.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoPaciente.IdTipoNumeracion)
540                   End If
                      
550                   If oDoHistoria.NroHistoriaClinica = 0 Then
560                       If oDoHistoria.IdTipoNumeracion = sghHistoriaDefinitivaReciclada Then
570                           Trace "No se pudo generar el nro de de historia clínica reciclada", "AdmisionHospAgregar"
580                           bProcesoOK = False: GoTo Terminar
590                       End If
600                   End If
                      
610                   oDoPaciente.NroHistoriaClinica = oDoHistoria.NroHistoriaClinica
                      
620                   oDoHistoria.NroHistoriaClinicaAnterior = lIdHistoriaAnterior
630                   oDoHistoria.IdTipoNumeracionAnterior = lIdTipoNumeracionAnterior
                      
640                   If Not oHistoria.Insertar(oDoHistoria) Then
650                       bProcesoOK = False: ms_MensajeError = oHistoria.MensajeError: GoTo Terminar
660                   End If
670               End Select
680           End Select
              
690           If Not PacientesModificar(oDoPaciente, IIf(lbPacienteNN = True, False, True), oDoPacienteDatosAdd) Then
700               bProcesoOK = False: ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
710           End If
              
              
720       End If
          
         
         
          'Agregar cuenta atención
730       oDOCuentaAtencion.FechaCreacion = CDate(oDoAtencion.FechaIngreso & " " & oDoAtencion.HoraIngreso)
740       oDOCuentaAtencion.idPaciente = oDoPaciente.idPaciente
750       If Not oCuentaAtencion.Insertar(oDOCuentaAtencion) Then
760           bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
770       End If
          'Sunasa
780       If oDoSunasaPacientesHistoricos.NuevoSeguro = True Then
790           oDoSunasaPacientesHistoricos.idPaciente = oDoPaciente.idPaciente
800           If Not oSunasaPacientesHistoricos.Insertar(oDoSunasaPacientesHistoricos) Then
810               bProcesoOK = False: ms_MensajeError = oSunasaPacientesHistoricos.MensajeError: GoTo Terminar
820           End If
830       Else
840           If Not oSunasaPacientesHistoricos.modificar(oDoSunasaPacientesHistoricos) Then
850               bProcesoOK = False: ms_MensajeError = oSunasaPacientesHistoricos.MensajeError: GoTo Terminar
860           End If
870       End If
880       oDoAtencion.idSunasaPacienteHistorico = oDoSunasaPacientesHistoricos.idSunasaPacienteHistorico
          'Agregar atencion
890       oDoAtencion.idPaciente = oDoPaciente.idPaciente
900       oDoAtencion.idCuentaAtencion = oDOCuentaAtencion.idCuentaAtencion
910       If Not oAtencion.Insertar(oDoAtencion) Then
920           bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
930       End If
          '
940       Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "A", oDoAtencion.idAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, lcObservacionesAuditoria)    'ListBarItems.idListItem
          
          'Agregar solicitud de prestamo
950       Select Case oDoHistoria.IdTipoNumeracion
          Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
960           If oDoAtencion.idTipoServicio = 3 Then
970               InicializarPrestamoAutomaticoHosp oDoAtencion, oDoPaciente, oDOHistoriaSolicitada, bPacienteNuevo
980           Else
990               InicializarPrestamoAutomaticoEmerg oDoAtencion, oDoPaciente, oDOHistoriaSolicitada, bPacienteNuevo
1000          End If
1010          If Not oHistoriaSolicitada.Insertar(oDOHistoriaSolicitada) Then
1020              bProcesoOK = False: ms_MensajeError = oHistoriaSolicitada.MensajeError: GoTo Terminar
1030          End If
1040      End Select
          
          
          
          'Modificar la atencion asociada si es que la tuviera
          'Coloca los datos de alta
1050      If oDOAtencionPadre.idAtencion <> 0 Then
1060          If Not oAtencion.modificar(oDOAtencionPadre) Then
1070              bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
1080          End If
1090      End If
          
          'Agregar diagnosticos
1100      If Not oAtencionesDx.ActualizarDiagnosticosAtencion(oDiagnosticos, oDoAtencion.idAtencion) Then
1110          bProcesoOK = False: ms_MensajeError = oAtencionesDx.MensajeError: GoTo Terminar
1120      End If
              
          '**************************Solo para el caso de CONSULTORIO DE EMERGENCIA*******************
1130      If oDoAtencion.idTipoServicio = sghEmergenciaConsultorios Then
               Dim lnIdServicio As Long
               Dim lnPrecioUnitario As Double
               Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
1140           Dim lnIdtipoFinanciamiento As Long: Dim lnIdFuenteFinanciamiento As Long
1150           lnIdServicio = mo_ReglasFacturacion.ObtenerCodigoDeConsultaDeEmergencia()
1160           If lnIdServicio > 0 Then
1170               lnPrecioUnitario = 0
1180               oDoCatalogoServicioHosp.PrecioUnitario = 0
1190               lnIdtipoFinanciamiento = oDoAtencion.IdFormaPago
1200               lnIdFuenteFinanciamiento = oDoAtencion.idFuenteFinanciamiento
1210               Set oDoCatalogoServicioHosp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarPorId(lnIdServicio, lnIdtipoFinanciamiento)
1220               If oDoCatalogoServicioHosp.PrecioUnitario = 0 And oDoCatalogoServicioHosp.SeUsaSinPrecio = False Then
1230                  lcCodigoNombre = oBuscaCodigoNombre.DevuelveNombreMedicamentoOServicioSegunId(lnIdServicio, sghServicio)
1240                  ms_MensajeError = "Tiene problemas con el ID SERVICIO: " & lcCodigoNombre & Chr(13) & "para el PRODUCTO/PLAN: " & oDoAtencion.IdFormaPago & Chr(13) & "...consulte al ADMINISTRADOR DEL SISTEMA...."
1250                  bProcesoOK = False: GoTo Terminar
1260               Else
1270                  lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
1280               End If
                   'If lnPrecioUnitario > 0 Then
                          'Crear orden de servicio
                          Dim oDoFactOrdenServ As New DoFactOrdenServ
                          Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
1290                      With oDoFactOrdenServ
1300                          .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL()
1310                          .idCuentaAtencion = oDoAtencion.idCuentaAtencion
1320                          .IdServicioPaciente = oDoAtencion.idServicioIngreso
1330                          .IdEstadoFacturacion = 1               'Registrado
1340                          .idPaciente = oDoAtencion.idPaciente
1350                          .IdPuntoCarga = 10                      'Admisión de emergencia
1360                          .IdTipoFinanciamiento = lnIdtipoFinanciamiento
1370                          .idFuenteFinanciamiento = lnIdFuenteFinanciamiento
1380                          .idUsuario = oDoAtencion.IdUsuarioAuditoria
1390                          .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1400                          .fechaDespacho = .FechaCreacion
1410                          .IdUsuarioDespacho = oDoAtencion.IdUsuarioAuditoria
1420                      End With
1430                      If Not oFactOrdenesServicio.Insertar(oDoFactOrdenServ) Then
1440                          bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
1450                      End If
1460                      If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(lnIdtipoFinanciamiento, oConexion) = True Then
                            '***********Contado ***********
1470                        With oDOFactOrdenServicioPagos
1480                            .FechaCreacion = oDoFactOrdenServ.FechaCreacion
1490                            .IdEstadoFacturacion = 1   'Registrado
1500                            .IdOrden = oDoFactOrdenServ.IdOrden
1510                            .idUsuario = oDoAtencion.IdUsuarioAuditoria
1520                            .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1530                        End With
1540                        If Not oFactServicioPagos.Insertar(oDOFactOrdenServicioPagos) Then
1550                            bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
1560                        End If
1570                      End If
                          'Despachos
                          Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
1580                      If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(lnIdtipoFinanciamiento, oConexion) = False Then
1590                         lnPrecioUnitario = 0
1600                         oDoCatalogoServicioHosp.PrecioUnitario = 0
1610                         Set oDoCatalogoServicioHosp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarPorId(lnIdServicio, lnIdtipoFinanciamiento)
1620                         If oDoCatalogoServicioHosp.PrecioUnitario = 0 Then
1630                            lnPrecioUnitario = 0
1640                         Else
1650                            lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
1660                         End If
1670                      End If
1680                      With oDoFacturacionServicioDesp
1690                          .cantidad = 1
1700                          .IdOrden = oDoFactOrdenServ.IdOrden
1710                          .idProducto = lnIdServicio
1720                          .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1730                          .Precio = lnPrecioUnitario
1740                          .total = lnPrecioUnitario
1750                      End With
1760                      If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then
1770                             bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
1780                      End If
                          'Crear Detalle de servicio
1790                      If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(lnIdtipoFinanciamiento, oConexion) = True Then
                             '***********Contado ***********
                             Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
1800                         With oDoFacturacionServicioPagos
1810                             .cantidad = 1
1820                             .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
1830                             .idProducto = lnIdServicio
1840                             .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1850                             .Precio = lnPrecioUnitario
1860                             .total = lnPrecioUnitario
1870                         End With
1880                         If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
1890                             bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
1900                         End If
1910                      Else
                             '***********Sis,Soat,exoneracion, convenio ***********
                             Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
1920                         With oDoFacturacionServicioFinanc
1930                             .CantidadFinanciada = 1
1940                             .FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
1950                             .IdOrden = oDoFactOrdenServ.IdOrden
1960                             .idProducto = lnIdServicio
1970                             .IdTipoFinanciamiento = lnIdtipoFinanciamiento
1980                             .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1990                             .IdUsuarioAutoriza = oDoAtencion.IdUsuarioAuditoria
2000                             .PrecioFinanciado = lnPrecioUnitario
2010                             .TotalFinanciado = lnPrecioUnitario
2020                             .idFuenteFinanciamiento = lnIdFuenteFinanciamiento
2030                         End With
2040                         If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then
2050                             bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo Terminar
2060                         End If
2070                      End If
                   'End If
2080          End If
2090      End If
          
          'Agregar estancia en el servicio
2100      If Not oEstanciaHospitalaria.ActualizarEstanciaHospitalaria(oOcupaciones, oDoAtencion) Then
2110          bProcesoOK = False: ms_MensajeError = oEstanciaHospitalaria.MensajeError: GoTo Terminar
2120      End If
              
          'En el caso de consultorios de emrgencia no se registra nro de cama
2130      If oDoAtencion.idTipoServicio = sghHospitalizacion Or lbElServicioEsObservacionDeEmergencia = True Then
              'Actualizar cama
2140          If Not oCama.ActualizarIdPaciente(oDoPaciente.idPaciente, ObtenerUltimaCamaOcupada(oDoAtencion.IdCamaIngreso, oOcupaciones, oDOCuentaAtencion.IdEstado)) Then
2150              bProcesoOK = False: ms_MensajeError = oCama.MensajeError: GoTo Terminar
2160          End If
2170      End If
          
2180      If oDoAtencion.idTipoServicio = sghHospitalizacion Then
              'Agregar nacimientos
2190          If Not oAtencionesNacimientos.ActualizarNacimientosAtencion(oNacimientos, oDoAtencion.idAtencion) Then
2200              bProcesoOK = False: ms_MensajeError = oAtencionesNacimientos.MensajeError: GoTo Terminar
2210          End If
2220      End If
          
          'Agregar causa externa de morbilidad para paciente de emergencia
2230      If oDoAtencion.idTipoServicio = sghEmergenciaConsultorios Or oDoAtencion.idTipoServicio = sghEmergenciaObservacion Then
              Dim oAtencionEmergencia As New AtencionesEmergencia
2240          Set oAtencionEmergencia.Conexion = oConexion
2250          oDOAtencionEmergencia.idAtencion = oDoAtencion.idAtencion
2260          If Not oAtencionEmergencia.Insertar(oDOAtencionEmergencia) Then
2270              bProcesoOK = False: ms_MensajeError = oAtencionEmergencia.MensajeError: GoTo Terminar
2280          End If
2290      End If
          '
          '****Actualiza Dato del Recien Nacido (que está en la Cuenta de la Madre)....tabla: AtencionesNacimiento
          '****Asocia la Cuenta de la madre con la cuenta con el hijo..................tabla: AtencionesHijoMadre
2300      If oAtencionesNacimientos.AsignaNULLidPacienteNacido(oDoPaciente.idPaciente) = False Then
2310         ms_MensajeError = oAtencionesNacimientos.MensajeError: GoTo Terminar
2320      End If
2330      oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
2340      oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
2350      If oAtencionesHijoMadre.Eliminar(oDoAtencionesHijoMadre) = False Then
2360         ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
2370      End If
2380      If lnIdNacimientoSeleccionado > 0 Then
             Dim lnIdAtencionMadre As Long
2390         oDOAtencionNacimiento.idNacimiento = lnIdNacimientoSeleccionado
2400         If oAtencionesNacimientos.SeleccionarPorId(oDOAtencionNacimiento) = True Then
2410            lnIdAtencionMadre = oDOAtencionNacimiento.idAtencion
2420            oDOAtencionNacimiento.idPacienteNacido = oDoPaciente.idPaciente
2430            oDOAtencionNacimiento.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
2440            If oAtencionesNacimientos.modificar(oDOAtencionNacimiento) = True Then
2450               ms_MensajeError = oAtencionesNacimientos.MensajeError: GoTo Terminar
2460            End If
2470            oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
2480            If oAtencionesHijoMadre.SeleccionarPorId(oDoAtencionesHijoMadre) = True Then
2490                 oDoAtencionesHijoMadre.idAtencionMadre = lnIdAtencionMadre
2500                 oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
2510                 If oAtencionesHijoMadre.modificar(oDoAtencionesHijoMadre) = True Then
2520                    ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
2530                 End If
2540            Else
2550                 oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
2560                 oDoAtencionesHijoMadre.idAtencionMadre = lnIdAtencionMadre
2570                 oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
2580                 If oAtencionesHijoMadre.Insertar(oDoAtencionesHijoMadre) = True Then
2590                    ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
2600                 End If
2610            End If
2620         End If
2630      End If
          '
          If oDoAtencion.idTipoServicio = 3 Then
                mo_DoAtencionDatosAdicionales.emergenciaCorrelativo = ""
          ElseIf sighentidades.Parametro550 = "S" Then
                Dim lcNuevoCorrelativo As String
                Call mo_reglasComunes.ServiciosGrabaYdevueveCorrelativoEmergencia(False, lcNuevoCorrelativo, oDoAtencion.idServicioIngreso, oConexion)
                mo_DoAtencionDatosAdicionales.emergenciaCorrelativo = lcNuevoCorrelativo
          End If
          'datos adicionales
2640      mo_DoAtencionDatosAdicionales.idAtencion = oDoAtencion.idAtencion
2650      If Not oAtencionesDatosAdicionales.Insertar(mo_DoAtencionDatosAdicionales) Then
2660          bProcesoOK = False: ms_MensajeError = oAtencionesDatosAdicionales.MensajeError: GoTo Terminar
2670      End If
2680      If mo_DoAtencionDatosAdicionales.idSiaSis > 0 Then
2690         Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "A", oDoAtencion.idCuentaAtencion, "FUA", oConexion, 1345, lcNombrePc, "Cod.Prestacion FUA: " & mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion & " N° Cuenta: " & oDoAtencion.idCuentaAtencion)
2700      End If
          '
          mo_ReglasArchivoClinico.HistoriasClinicasActualizaFmovimiento oConexion, oDoAtencion.idPaciente
          '
2710      ms_MensajeError = ""
Terminar:
2720      If bProcesoOK Then
2730          AdmisionHospAgregar = True
2740          oConexion.CommitTrans
2750      Else
2760          If Err.Description <> "" Then
2770             ms_MensajeError = Err.Description
2780          End If
2790          AdmisionHospAgregar = False
2800          oConexion.RollbackTrans
2810          ms_MensajeError = ms_MensajeError & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                                                 "Function AdmisionHospAgregar", "reglasAdmision.cls") 'debb-02/05/2016
                         
2820      End If
2830      oConexion.Close
2840      Set oConexion = Nothing
2850      Set oPaciente = Nothing
2860      Set oCuentaAtencion = Nothing
2870      Set oAtencion = Nothing
2880      Set oDOHistoriaSolicitada = Nothing
2890      Set oHistoriaSolicitada = Nothing
2900      Set oHistoria = Nothing
2910      Set oCama = Nothing
2920      Set oAtencionesDx = Nothing
2930      Set oEstanciaHospitalaria = Nothing
2940      Set oAtencionDiagnostico = Nothing
2950      Set oAtencionesNacimientos = Nothing
2960      Set oFacturacionServicioDespacho = Nothing
2970      Set oFacturacionServicioPagos = Nothing
2980      Set oFacturacionServicioFinan = Nothing
2990      Set oFactServicioPagos = Nothing
3000      Set oCatalogoServicios = Nothing
3010      Set oDOFacturacionServicios = Nothing
3020      Set mo_ReglasFacturacion = Nothing
3030      Set oFactOrdenesServicio = Nothing
3040      Set oBuscaCodigoNombre = Nothing
3050      Set oRsTmp = Nothing
3060      Set oSunasaPacientesHistoricos = Nothing
3070      Set oAtencionesDatosAdicionales = Nothing
3080      Set oDOAtencionNacimiento = Nothing
3090      Set oDoAtencionesHijoMadre = Nothing
3100      Set oAtencionesHijoMadre = Nothing

End Function

Function ObtenerUltimaCamaOcupada(lIdCamaIngreso As Long, oOcupaciones As Collection, lEstadoCuenta As Long)
Dim lIdUltimaCama As Long
Dim oDOOcupacion As New DOEstanciaHospitalaria

    If lEstadoCuenta <> 1 Then
        'No hay camas ocupadas por el paciente
        lIdUltimaCama = 0
    Else
        'por defecto se considera la cama de ingreso
        lIdUltimaCama = lIdCamaIngreso
        
        'Y en el caso haya transferencias
        For Each oDOOcupacion In oOcupaciones
            lIdUltimaCama = oDOOcupacion.IdCama
        Next
    End If
    ObtenerUltimaCamaOcupada = lIdUltimaCama
    Set oDOOcupacion = Nothing
End Function


Function AdmisionHospModificar(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, oDoPaciente As DOPaciente, _
                               oDoHistoria As DOHistoriaClinica, lIdTipoNumeracionAnterior As sghTipoNumeracionDeNroHistoria, _
                               oOcupaciones As Collection, oDiagnosticos As Collection, oProcedimientos As Collection, _
                               oExamenes As Collection, oNacimientos As Collection, lIdHistoriaAnterior As Long, _
                               oDOAtencionEmergencia As DOAtencionEmergencia, lTipoAccionAdmision As sghTipoAccionEmergenciaYHospitalizacion, _
                               ldFechaEgresoMedicoAnterior As Date, lbPacienteNN As Boolean, lnIdTablaLISTBARITEMS As Long, _
                               lcNombrePc As String, lcObservacionesAuditoria As String, lnIdPlanAnterior As Long, _
                               lnIdTipoFinanciamientoAnterior As Long, oRsEstancia As Recordset, lnIdNacimientoSeleccionado As Long, _
                               oDoSunasaPacientesHistoricos As DoSunasaPacientesHistoricos, mo_DoAtencionDatosAdicionales As DoAtencionDatosAdicionales, _
                               lcUltimoCodigoDeServicioTransferido As String, lbElServicioEsObservacionDeEmergencia As Boolean, _
                               oEpisodioClinico As EpisodioClinico, _
                             Optional oDoPacienteDatosAdd As DoPacienteDatosAdd = Nothing) As Boolean
      Dim oConexion As New ADODB.Connection
      Dim oPaciente As New Pacientes
      Dim oCuentaAtencion As New CuentasAtencion
      Dim oAtencion As New Atenciones
      Dim oDOHistoriaSolicitada As New DOHistoriaSolicitada
      Dim oHistoriaSolicitada As New HistoriasSolicitadas
      Dim oCama As New Camas
      Dim bProcesoOK As Boolean
      Dim oHistoria As New HistoriasClinicas
      Dim oAtencionesDx As New AtencionesDiagnosticos
      Dim oEstanciaHospitalaria As New AtencionesEstanciaHosp
      Dim oAtencionDiagnostico As New AtencionesDiagnosticos
      Dim oAtencionesNacimientos As New AtencionesNacimientos
      Dim oFactOrdenesServicio As New FactOrdenServicio
      Dim oDOfactOrdenesServicio As New DOFactOrdenServicio
      Dim oFacturacionServicios As New FacturacionServicios
      Dim oDOFacturacionServicios As New DOFacturacionServicios
      Dim lcBuscaParametro As New SIGHDatos.Parametros
      Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
      Dim oFacturacionServicioPagos As New FacturacionServicioPagos
      Dim oFacturacionServicioFinan As New FacturacionServicioFin
      Dim oFactServicioPagos As New FactOrdenServicioPagos
      Dim oDOCama As New DOCama
      Dim oCamas As New Camas
      Dim oDOCamasMovimientos As New DOCamasMovimientos
      Dim oCamasMovimientos As New CamasMovimientos
      Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
      Dim oBuscaDiasPaciente As New SIGHDatos.Parametros
      Dim oDoFactOrdenServ As New DoFactOrdenServ
      Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
      Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
      Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
      Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
      Dim oSunasaPacientesHistoricos As New SunasaPacientesHistoricos
      Dim oAtencionesDatosAdicionales As New AtencionesDatosAdicionales
      Dim oDOAtencionNacimiento As New DOAtencionNacimiento
      Dim oDoAtencionesHijoMadre As New DoAtencionesHijoMadre, oAtencionesHijoMadre As New AtencionesHijoMadre
      Dim oDoAtencionesEpisodiosCab As New DoAtencionesEpisodiosCab, oAtencionesEpisodiosCab As New AtencionesEpisodiosCab
      Dim oDoAtencionesEpisodiosDet As New DoAtencionesEpisodiosDet, oAtencionesEpisodiosDet As New AtencionesEpisodiosDet
10    Dim lnIdServicio As Long: Dim lnPrecioUnitario As Double
      Dim oRsTmp As New ADODB.Recordset
      Dim oRsTmp1 As New Recordset
      Dim oRsTmp2 As New Recordset
      Dim oRsBoletaPagada As New Recordset
      Dim lnIdtipoFinanciamiento As Long
      Dim oReglasCaja As New SIGHNegocios.ReglasCaja
      Dim lcCodigoNombre As String, lcHoraEstanciaMax As String

          oConexion.CommandTimeout = 900
          oConexion.CursorLocation = adUseClient
30        oConexion.Open sighentidades.CadenaConexion
          'Chequea si hay Boleta  Activa, sino crear PAGO DE CONSULTA EMERGENCIA  (DEBB 27/04/2012)
50        lnIdServicio = mo_ReglasFacturacion.ObtenerCodigoDeConsultaDeEmergencia()
60        Set oRsBoletaPagada = oReglasCaja.CajaComprobantesPagoChequeaSiYaPago(oDoAtencion.idCuentaAtencion, lnIdServicio, oConexion)
70        If oRsBoletaPagada.RecordCount > 0 Then
80           lnIdServicio = 0
90        End If
100       oRsBoletaPagada.Close
          '
110       Set oCuentaAtencion.Conexion = oConexion
120       Set oAtencion.Conexion = oConexion
130       Set oPaciente.Conexion = oConexion
140       Set oHistoriaSolicitada.Conexion = oConexion
150       Set oCama.Conexion = oConexion
160       Set oHistoria.Conexion = oConexion
170       Set oAtencionesDx.Conexion = oConexion
180       Set oEstanciaHospitalaria.Conexion = oConexion
190       Set oAtencionDiagnostico.Conexion = oConexion
200       Set oAtencionesNacimientos.Conexion = oConexion
210       Set oFactOrdenesServicio.Conexion = oConexion
220       Set oFacturacionServicios.Conexion = oConexion
230       Set oFacturacionServicioPagos.Conexion = oConexion
240       Set oFacturacionServicioFinan.Conexion = oConexion
250       Set oFacturacionServicioDespacho.Conexion = oConexion
260       Set oFactServicioPagos.Conexion = oConexion
270       Set oCamas.Conexion = oConexion
280       Set oCamasMovimientos.Conexion = oConexion
290       Set oSunasaPacientesHistoricos.Conexion = oConexion
300       Set oAtencionesDatosAdicionales.Conexion = oConexion
310       Set oAtencionesHijoMadre.Conexion = oConexion
320       Set oAtencionesEpisodiosCab.Conexion = oConexion
330       Set oAtencionesEpisodiosDet.Conexion = oConexion
          
340       oConexion.BeginTrans
350       bProcesoOK = True
360       lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
          
          'Agregar historia clinica (Si es tenia historia temporal y ahora tiene historia definitiva)
370       Select Case lIdTipoNumeracionAnterior
          Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
380           If lIdTipoNumeracionAnterior = sghHistoriaDefinitivaManual Then
390               If Not oHistoria.modificar(oDoHistoria) Then
400               End If
410           End If
              
420       Case Else
              
              'Si era temporal y ahora es definitiva => Genera y agrega la historia
430           Select Case oDoHistoria.IdTipoNumeracion
              Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
                  
440               If oDoHistoria.IdTipoNumeracion <> sghHistoriaDefinitivaManual Then
450                   oDoHistoria.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoPaciente.IdTipoNumeracion)
460               End If
                  
470               If oDoHistoria.NroHistoriaClinica = 0 Then
480                   If oDoHistoria.IdTipoNumeracion = sghHistoriaDefinitivaReciclada Then
490                       Trace "No se pudo generar el nro de de historia clínica reciclada", "AdmisionHospModificar"
500                       bProcesoOK = False: GoTo Terminar
510                   End If
520               End If
                  
530               oDoPaciente.NroHistoriaClinica = oDoHistoria.NroHistoriaClinica
                   
540               oDoHistoria.NroHistoriaClinicaAnterior = lIdHistoriaAnterior
550               oDoHistoria.IdTipoNumeracionAnterior = lIdTipoNumeracionAnterior
                   
560               If Not oHistoria.Insertar(oDoHistoria) Then
570                   bProcesoOK = False: ms_MensajeError = oHistoria.MensajeError: GoTo Terminar
580               End If
590           End Select
600       End Select
          
          'Modificar paciente
610       If Not PacientesModificar(oDoPaciente, IIf(lbPacienteNN = True, False, True), oDoPacienteDatosAdd) Then
620           bProcesoOK = False: ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
630       End If
          
          'Modificar solicitud de prestamo
640       Select Case oDoHistoria.IdTipoNumeracion
          Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
650           InicializarPrestamoAutomaticoHosp oDoAtencion, oDoPaciente, oDOHistoriaSolicitada, False
660           If oDoAtencion.IdEstadoAtencion = sghEstadoTabla.sghAnulado Then
670               If Not oHistoriaSolicitada.Eliminar(oDOHistoriaSolicitada) Then
680                   bProcesoOK = False: ms_MensajeError = oHistoriaSolicitada.MensajeError: GoTo Terminar
690               End If
700           Else
710               If Not oHistoriaSolicitada.modificar(oDOHistoriaSolicitada) Then
720                   bProcesoOK = False: ms_MensajeError = oHistoriaSolicitada.MensajeError: GoTo Terminar
730               End If
740           End If
750       End Select
          'Sunasa
760       If oDoSunasaPacientesHistoricos.NuevoSeguro = True Then
770           oDoSunasaPacientesHistoricos.idPaciente = oDoPaciente.idPaciente
780           If Not oSunasaPacientesHistoricos.Insertar(oDoSunasaPacientesHistoricos) Then
790               bProcesoOK = False: ms_MensajeError = oSunasaPacientesHistoricos.MensajeError: GoTo Terminar
800           End If
810       Else
820           If Not oSunasaPacientesHistoricos.modificar(oDoSunasaPacientesHistoricos) Then
830               bProcesoOK = False: ms_MensajeError = oSunasaPacientesHistoricos.MensajeError: GoTo Terminar
840           End If
850       End If
860       oDoAtencion.idSunasaPacienteHistorico = oDoSunasaPacientesHistoricos.idSunasaPacienteHistorico
          'Modificar atencion
870       oDoAtencion.idPaciente = oDoPaciente.idPaciente
880       If Not oAtencion.modificar(oDoAtencion) Then
890           bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
900       End If
          '
910       Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "M", oDoAtencion.idAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, lcObservacionesAuditoria)  'ListBarItems.idListItem
          '
920       If sighentidades.EsFecha(oDoAtencion.fechaEgreso, "DD/MM/AAAA") Then
930          oDOCuentaAtencion.IdEstado = sghEstadoCuenta.sghConAltaMedica
940       End If
950       oDOCuentaAtencion.FechaCreacion = CDate(oDoAtencion.FechaIngreso & " " & oDoAtencion.HoraIngreso)
960       oDOCuentaAtencion.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
970       If Not oCuentaAtencion.modificar(oDOCuentaAtencion) Then
980           bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
990       End If
          '
1000      If oDoAtencion.IdEstadoAtencion = sghEstadoTabla.sghAnulado Then
              'Eliminar Id paciente de las camas ocupadas
1010          If oDoAtencion.idTipoServicio = sghHospitalizacion Or oDoAtencion.idTipoServicio = sghEmergenciaObservacion Then
1020              If Not oCama.ActualizarIdPaciente(oDoAtencion.idPaciente, 0) Then
1030                  bProcesoOK = False: ms_MensajeError = oCama.MensajeError: GoTo Terminar
1040              End If
1050          End If
1060      Else
1070          If Not oCama.ActualizarIdPaciente(oDoPaciente.idPaciente, oDoAtencion.IdCamaIngreso) Then
1080              bProcesoOK = False: ms_MensajeError = oCama.MensajeError: GoTo Terminar
1090          End If
1100      End If
          
          'Agregar diagnosticos
1110      If Not oAtencionesDx.ActualizarDiagnosticosAtencion(oDiagnosticos, oDoAtencion.idAtencion) Then
1120          bProcesoOK = False: ms_MensajeError = oAtencionesDx.MensajeError: GoTo Terminar
1130      End If
          
          'Modificar estancias hospitalaria solo si la cuenta aun esta abierta
1140      If oDOCuentaAtencion.IdEstado = sghEstadoCuenta.sghAbierto Or lcUltimoCodigoDeServicioTransferido <> "" Then
1150          If Not oEstanciaHospitalaria.ActualizarEstanciaHospitalaria(oOcupaciones, oDoAtencion) Then
1160              bProcesoOK = False: ms_MensajeError = oEstanciaHospitalaria.MensajeError: GoTo Terminar
1170          End If
1180      End If
          
          'Actualizar ocupacion o desocupacion de cama
1190      If oDoAtencion.idTipoServicio = sghHospitalizacion Or lbElServicioEsObservacionDeEmergencia = True Then
1200          If Not oCama.ActualizarIdPaciente(oDoPaciente.idPaciente, ObtenerUltimaCamaOcupada(oDoAtencion.IdCamaIngreso, oOcupaciones, oDOCuentaAtencion.IdEstado)) Then
1210              bProcesoOK = False: ms_MensajeError = oCama.MensajeError: GoTo Terminar
1220          End If
1230      End If
          
         ' If oDoAtencion.idTipoServicio = sghHospitalizacion Then
              'Agregar nacimientos
1240          If Not oAtencionesNacimientos.ActualizarNacimientosAtencion(oNacimientos, oDoAtencion.idAtencion) Then
1250              bProcesoOK = False: ms_MensajeError = oAtencionesNacimientos.MensajeError: GoTo Terminar
1260          End If
         ' End If
          
          'Modifica la causa externa de morbilidad para los pacientes de emergencia
1270      If oDoAtencion.idTipoServicio = sghEmergenciaConsultorios Or oDoAtencion.idTipoServicio = sghEmergenciaObservacion Then
              Dim oAtencionEmergencia As New AtencionesEmergencia
1280          Set oAtencionEmergencia.Conexion = oConexion
1290          oDOAtencionEmergencia.idAtencion = oDoAtencion.idAtencion
1300          If Not oAtencionEmergencia.modificar(oDOAtencionEmergencia) Then
1310              bProcesoOK = False: ms_MensajeError = oAtencionEmergencia.MensajeError: GoTo Terminar
1320          End If
1330      End If
          
          
       
          
          
          '***************daniel barrantes**************
          '***************Si el Paciente tiene ALTA MEDICA se genera
          '***************un SERVICIO a su CUENTA por los 'DIAS DE ESTANCIA'
1340      If sighentidades.EsFecha(oDoAtencion.fechaEgreso, "DD/MM/AAAA") Then
              'Desocupa cama
1350          If oDoAtencion.idTipoServicio = sghHospitalizacion Then
1360             If lcBuscaParametro.SeleccionaFilaParametro(213) = "1" Then
1370                  If Not oCama.ActualizarIdPaciente(oDoPaciente.idPaciente, 0) Then
1380                      bProcesoOK = False: ms_MensajeError = oCama.MensajeError: GoTo Terminar
1390                  End If
1400             End If
1410          ElseIf lcBuscaParametro.SeleccionaFilaParametro(214) = "1" Then
1420                  If Not oCama.ActualizarIdPaciente(oDoPaciente.idPaciente, 0) Then
1430                      bProcesoOK = False: ms_MensajeError = oCama.MensajeError: GoTo Terminar
1440                  End If
1450          End If
              'Genera Facturacion por ESTADIA
1460          If ldFechaEgresoMedicoAnterior <> oDoAtencion.fechaEgreso Then
1470              Dim lnDiasEstancia As Integer: Dim lnHorasEstancia As Integer
                  Dim lcCodServicio As String
                  Dim lbEliminaConsultaEmerg As Boolean
1480              lbEliminaConsultaEmerg = False
                  
                  'Dias de Estancia
1490              lnDiasEstancia = oBuscaDiasPaciente.DiasDelPacienteEnHospitalizacionEmergencia(oDoAtencion.FechaIngreso, oDoAtencion.HoraIngreso, oDoAtencion.fechaEgreso, oDoAtencion.HoraEgreso, lcHoraEstanciaMax)
          
                  'Horas de Estancia
1500              lnHorasEstancia = sighentidades.HorasDelPacienteEnHospitalizacionEmergencia(oDoAtencion.FechaIngreso, oDoAtencion.HoraIngreso, oDoAtencion.fechaEgreso, oDoAtencion.HoraEgreso)
                  
                  'Codigo Servicio
1510              If oDoAtencion.idTipoServicio = sghHospitalizacion Then
                     'lcCodServicio = "99221"
1520                 lnIdServicio = Val(lcBuscaParametro.SeleccionaFilaParametro(202))
1530              ElseIf lnHorasEstancia < 12 Then   'Emergencia Menor a 12 horas
                     'lcCodServicio = "99234"
1540                 lnIdServicio = Val(lcBuscaParametro.SeleccionaFilaParametro(203))
1550                 If oDoAtencion.idFuenteFinanciamiento = 2 Then  'solo SOAT
1560                    lbEliminaConsultaEmerg = True
1570                 End If
1580              ElseIf lnHorasEstancia <= 24 Then  'Emergencia Entre 12 y 24 horas
                     'lcCodServicio = "99235"
1590                 lnIdServicio = Val(lcBuscaParametro.SeleccionaFilaParametro(204))
1600                 If oDoAtencion.idFuenteFinanciamiento = 2 Then  'solo SOAT
1610                    lbEliminaConsultaEmerg = True
1620                 End If
1630              Else                               'Emergencia >24 horas
                     'lcCodServicio = "99235"
1640                 lnIdServicio = Val(lcBuscaParametro.SeleccionaFilaParametro(204))
1650                 If oDoAtencion.idFuenteFinanciamiento = 2 Then  'solo SOAT
1660                    lbEliminaConsultaEmerg = True
1670                 End If
1680              End If
                  
                  
                  
1690              If lnIdServicio > 0 Then
                      'elimina Codigo Servicio (cabecera/detalle) por Estancia de la Cuenta de Atencion
1700                  Set oRsTmp = mo_ReglasFacturacion.FactOrdenServicioFiltraPorIdCuenta(oDoAtencion.idCuentaAtencion)
1710                  If oRsTmp.RecordCount > 0 Then
1720                     oRsTmp.MoveFirst
1730                     Do While Not oRsTmp.EOF
1740                        If oRsTmp.Fields!IdPuntoCarga = 9 Then
                               '**********solo Hospitalizacion
1750                           oDoFactOrdenServ.IdOrden = oRsTmp.Fields!IdOrden
1760                           oDoFacturacionServicioFinanc.IdOrden = oRsTmp.Fields!IdOrden
1770                           oDoFacturacionServicioDesp.IdOrden = oRsTmp.Fields!IdOrden
1780                           Set oRsTmp1 = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrden(oRsTmp.Fields!IdOrden, oConexion)
1790                           If oRsTmp1.RecordCount > 0 Then
1800                             If oRsTmp1.Fields!IdEstadoFacturacion = 1 Then
1810                                  oDOFactOrdenServicioPagos.IdOrden = oRsTmp1.Fields!IdOrden
1820                                  oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp1.Fields!idOrdenPago
1830                                  oDoFacturacionServicioPagos.idOrdenPago = oRsTmp1.Fields!idOrdenPago
1840                                  If Not oFacturacionServicioPagos.Eliminar(oDoFacturacionServicioPagos) Then
1850                                     bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
1860                                  End If
1870                                  If Not oFactServicioPagos.Eliminar(oDOFactOrdenServicioPagos) Then
1880                                     bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
1890                                  End If
1900                             End If
1910                           End If
1920                           oDoFacturacionServicioFinanc.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1930                           If Not oFacturacionServicioFinan.Eliminar(oDoFacturacionServicioFinanc) Then
1940                              bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo Terminar
1950                           End If
1960                           oDoFacturacionServicioDesp.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1970                           If Not oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then
1980                              bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
1990                           End If
2000                           oDoFactOrdenServ.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2010                           If Not oFactOrdenesServicio.Eliminar(oDoFactOrdenServ) Then
2020                              bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
2030                           End If
2040                        ElseIf oRsTmp.Fields!IdPuntoCarga = 10 Then
                               '*********solo Emergencia
2050                           If lbEliminaConsultaEmerg = True And oRsTmp.Fields!idProducto = Val(lcBuscaParametro.SeleccionaFilaParametro(202)) Then
                                  'Consulta Emergencia/Solo SOAT
2060                              oDoFacturacionServicioFinanc.IdOrden = oRsTmp.Fields!IdOrden
2070                              oDoFacturacionServicioFinanc.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2080                              If Not oFacturacionServicioFinan.SeleccionarPorId(oDoFacturacionServicioFinanc) Then
2090                                 bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo Terminar
2100                              End If
2110                              oDoFacturacionServicioFinanc.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2120                              If Not oFacturacionServicioFinan.Eliminar(oDoFacturacionServicioFinanc) Then
2130                                 bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo Terminar
2140                              End If
2150                              oDoFactOrdenServ.IdOrden = oRsTmp.Fields!IdOrden
2160                              If Not oFactOrdenesServicio.SeleccionarPorId(oDoFactOrdenServ) Then
2170                                 bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
2180                              End If
2190                              oDoFactOrdenServ.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2200                              oDoFactOrdenServ.IdEstadoFacturacion = 9   'Anulado
2210                              If Not oFactOrdenesServicio.modificar(oDoFactOrdenServ) Then
2220                                 bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
2230                              End If
2240                           End If
2250                        End If
2260                        oRsTmp.MoveNext
2270                     Loop
2280                  End If
2290                  oRsTmp.Close
2300                  lnIdtipoFinanciamiento = oDoAtencion.IdFormaPago
                       'Crear orden de servicio
2310                   With oDoFactOrdenServ
2320                       .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL()
2330                       .idCuentaAtencion = oDoAtencion.idCuentaAtencion
2340                       .IdServicioPaciente = oDoAtencion.IdServicioEgreso
2350                       .IdEstadoFacturacion = 1               'Registrado
2360                       .idPaciente = oDoAtencion.idPaciente
2370                       .IdPuntoCarga = IIf(oDoAtencion.idTipoServicio = sghHospitalizacion, 9, 10)
2380                       .IdTipoFinanciamiento = lnIdtipoFinanciamiento
2390                       .idFuenteFinanciamiento = oDoAtencion.idFuenteFinanciamiento
2400                       .idUsuario = oDoAtencion.IdUsuarioAuditoria
2410                       .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2420                       .fechaDespacho = .FechaCreacion
2430                       .IdUsuarioDespacho = oDoAtencion.IdUsuarioAuditoria
2440                   End With
2450                   If Not oFactOrdenesServicio.Insertar(oDoFactOrdenServ) Then
2460                       bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
2470                   End If
2480                   If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(lnIdtipoFinanciamiento, oConexion) = True Then
                         '***********Contado ***********
2490                     With oDOFactOrdenServicioPagos
2500                         .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL()
2510                         .IdEstadoFacturacion = 1   'Registrado
2520                         .IdOrden = oDoFactOrdenServ.IdOrden
2530                         .idUsuario = oDoAtencion.IdUsuarioAuditoria
2540                         .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2550                     End With
2560                     If Not oFactServicioPagos.Insertar(oDOFactOrdenServicioPagos) Then
2570                         bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
2580                     End If
2590                   End If
2600                  oRsEstancia.MoveFirst
2610                  Do While Not oRsEstancia.EOF
                          'Crear Detalle de servicio
2620                      With oDoFacturacionServicioDesp
2630                          .cantidad = oRsEstancia.Fields!CantidadEstancia
2640                          .IdOrden = oDoFactOrdenServ.IdOrden
2650                          .idProducto = oRsEstancia.Fields!idProducto
2660                          .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2670                          .Precio = oRsEstancia.Fields!PrecioEstancia
2680                          .total = Round(oRsEstancia.Fields!CantidadEstancia * oRsEstancia.Fields!PrecioEstancia, 2)
2690                      End With
2700                      If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then
2710                             bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
2720                      End If
2730                      If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(lnIdtipoFinanciamiento, oConexion) = True Then
                             '***********Contado ***********
2740                         With oDoFacturacionServicioPagos
2750                             .cantidad = oRsEstancia.Fields!CantidadEstancia
2760                             .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
2770                             .idProducto = oRsEstancia.Fields!idProducto
2780                             .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2790                             .Precio = oRsEstancia.Fields!PrecioEstancia
2800                             .total = Round(oRsEstancia.Fields!CantidadEstancia * oRsEstancia.Fields!PrecioEstancia, 2)
2810                         End With
2820                         If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
2830                             bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
2840                         End If
2850                      Else
                             '***********Sis,Soat,exoneracion, convenio ***********
2860                         With oDoFacturacionServicioFinanc
2870                             .CantidadFinanciada = oRsEstancia.Fields!CantidadEstancia
2880                             .FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
2890                             .IdOrden = oDoFactOrdenServ.IdOrden
2900                             .idProducto = oRsEstancia.Fields!idProducto
2910                             .IdTipoFinanciamiento = lnIdtipoFinanciamiento
2920                             .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2930                             .IdUsuarioAutoriza = oDoAtencion.IdUsuarioAuditoria
2940                             .PrecioFinanciado = oRsEstancia.Fields!PrecioEstancia
2950                             .TotalFinanciado = Round(oRsEstancia.Fields!CantidadEstancia * oRsEstancia.Fields!PrecioEstancia, 2)
2960                             .idFuenteFinanciamiento = oDoAtencion.idFuenteFinanciamiento
2970                         End With
2980                         If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then
2990                             bProcesoOK = False: GoTo Terminar
3000                         End If
3010                      End If
3020                      oRsEstancia.MoveNext
3030                  Loop
3040              End If
3050          End If
3060      End If
          '
          '-----Chequea Si la "Cama de Ingreso" no pertenece al "Servicio de Ingreso"
          '-----de ser asi la "cama" pasará  al "servicio" como PRESTAMO en forma automatica
3070      If oDoAtencion.IdCamaIngreso > 0 Then
3080          oDOCama.IdCama = oDoAtencion.IdCamaIngreso
3090          If Not oCamas.SeleccionarPorId(oDOCama) Then
3100              bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
3110          End If
3120          If oDOCama.IdServicioUbicacionActual <> oDoAtencion.idServicioIngreso Then
3130             oDOCama.IdServicioUbicacionActual = oDoAtencion.idServicioIngreso
3140             oDOCama.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
3150             If Not oCamas.modificar(oDOCama) Then
3160                  bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
3170             End If
3180             oDOCamasMovimientos.IdCama = oDOCama.IdCama
3190             oDOCamasMovimientos.IdFechaIngreso = lcBuscaParametro.RetornaFechaHoraServidorSQL
3200             oDOCamasMovimientos.IdServicio = oDOCama.IdServicioUbicacionActual
3210             oDOCamasMovimientos.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
3220             If Not oCamasMovimientos.Insertar(oDOCamasMovimientos) Then
3230                  bProcesoOK = False: ms_MensajeError = oCamasMovimientos.MensajeError: GoTo Terminar
3240             End If
3250          End If
3260      End If
          '-----Chequea Si la "Cama de Egreso" no pertenece al "Servicio de Egreso"
          '-----de ser asi la "cama" pasará  al "servicio" como PRESTAMO en forma automatica
3270      If oDoAtencion.IdCamaEgreso > 0 Then
3280          oDOCama.IdCama = oDoAtencion.IdCamaEgreso
3290          If Not oCamas.SeleccionarPorId(oDOCama) Then
3300              bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
3310          End If
3320          If oDOCama.IdServicioUbicacionActual <> oDoAtencion.IdServicioEgreso Then
3330             oDOCama.IdServicioUbicacionActual = oDoAtencion.IdServicioEgreso
3340             oDOCama.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
3350             If Not oCamas.modificar(oDOCama) Then
3360                  bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
3370             End If
3380             oDOCamasMovimientos.IdCama = oDOCama.IdCama
3390             oDOCamasMovimientos.IdFechaIngreso = lcBuscaParametro.RetornaFechaHoraServidorSQL
3400             oDOCamasMovimientos.IdServicio = oDOCama.IdServicioUbicacionActual
3410             oDOCamasMovimientos.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
3420             If Not oCamasMovimientos.Insertar(oDOCamasMovimientos) Then
3430                  bProcesoOK = False: ms_MensajeError = oCamasMovimientos.MensajeError: GoTo Terminar
3440             End If
3450          End If
3460      End If
          '----Si cambio de Plan, debe cambiar: EMERG-> PAGO DE CONSULTA
3470      If (lnIdPlanAnterior <> oDoAtencion.idFuenteFinanciamiento Or lnIdTipoFinanciamientoAnterior <> oDoAtencion.IdFormaPago) And oDoAtencion.idTipoServicio = sghEmergenciaConsultorios Then
               Dim lnIdFuenteFinanciamiento As Long
              'Elimina Servicios
3480          Set oRsTmp2 = mo_ReglasFacturacion.FactOrdenServicioPagosPorIdCuentaAtencion(oDoAtencion.idCuentaAtencion)
3490          Set oRsTmp1 = mo_ReglasFacturacion.FactOrdenServicioFiltraPorIdCuenta(oDoAtencion.idCuentaAtencion)
3500          If oRsTmp1.RecordCount > 0 Then
3510             oRsTmp1.MoveFirst
3520             Do While Not oRsTmp1.EOF
3530                If oRsTmp1.Fields!IdPuntoCarga = 9 Or oRsTmp1.Fields!IdPuntoCarga = 10 Then   'Hosp/Emerg
3540                   oDoFactOrdenServ.IdOrden = oRsTmp1.Fields!IdOrden
3550                   If Not oFactOrdenesServicio.SeleccionarPorId(oDoFactOrdenServ) Then
3560                      bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
3570                   End If
3580                   If oDoFactOrdenServ.IdEstadoFacturacion = sghEstadoFacturacion.sghAtendido Then
3590                      oDoFactOrdenServ.IdEstadoFacturacion = sghEstadoFacturacion.sghAnulado
3600                      oDoFactOrdenServ.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
3610                      If Not oFactOrdenesServicio.modificar(oDoFactOrdenServ) Then
3620                         bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
3630                      End If
3640                      If Not oFacturacionServicioFinan.ActualizaIdEstadoFacturacion(oDoFactOrdenServ.IdOrden, 9) Then
3650                         bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo Terminar
3660                      End If
3670                   End If
3680                End If
3690                oRsTmp1.MoveNext
3700             Loop
3710           End If
               
3720           If oRsTmp2.RecordCount > 0 Then
3730             oRsTmp2.MoveFirst
3740             Do While Not oRsTmp2.EOF
3750                If oRsTmp2.Fields!IdPuntoCarga = 9 Or oRsTmp2.Fields!IdPuntoCarga = 10 Then   'Hosp/Emerg
3760                      oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp2.Fields!idOrdenPago
3770                      If Not oFactServicioPagos.SeleccionarPorId(oDOFactOrdenServicioPagos) Then
3780                         bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
3790                      End If
3800                      If oDOFactOrdenServicioPagos.IdEstadoFacturacion = sghEstadoFacturacion.sghAtendido Then
3810                          oDOFactOrdenServicioPagos.IdEstadoFacturacion = sghEstadoFacturacion.sghAnulado
3820                          oDOFactOrdenServicioPagos.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
3830                          If Not oFactServicioPagos.modificar(oDOFactOrdenServicioPagos) Then
3840                             bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
3850                          End If
3860                      End If
3870                End If
3880                oRsTmp2.MoveNext
3890             Loop
3900           End If
               'Crea Servicio
3910           ms_MensajeError = ""
3920           If lnIdServicio > 0 Then
3930               lnPrecioUnitario = 0
3940               oDoCatalogoServicioHosp.PrecioUnitario = 0
3950               lnIdtipoFinanciamiento = oDoAtencion.IdFormaPago
3960               lnIdFuenteFinanciamiento = oDoAtencion.idFuenteFinanciamiento
3970               Set oDoCatalogoServicioHosp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarPorId(lnIdServicio, lnIdtipoFinanciamiento)
3980               If oDoCatalogoServicioHosp.PrecioUnitario > 0 Then
3990                   lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
4000               End If
                   'If lnPrecioUnitario > 0 Then
                          'Crear orden de servicio
4010                      With oDoFactOrdenServ
4020                          .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL()
4030                          .idCuentaAtencion = oDoAtencion.idCuentaAtencion
4040                          .IdServicioPaciente = oDoAtencion.idServicioIngreso
4050                          .IdEstadoFacturacion = 1               'Registrado
4060                          .idPaciente = oDoAtencion.idPaciente
4070                          .IdPuntoCarga = 10                      'Admisión de emergencia
4080                          .IdTipoFinanciamiento = lnIdtipoFinanciamiento
4090                          .idFuenteFinanciamiento = lnIdFuenteFinanciamiento
4100                          .idUsuario = oDoAtencion.IdUsuarioAuditoria
4110                          .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
4120                          .fechaDespacho = .FechaCreacion
4130                          .IdUsuarioDespacho = oDoAtencion.IdUsuarioAuditoria
4140                      End With
4150                      If Not oFactOrdenesServicio.Insertar(oDoFactOrdenServ) Then
4160                          bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
4170                      End If
4180                      If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(lnIdtipoFinanciamiento, oConexion) = True Then
                            '***********Contado ***********
4190                        With oDOFactOrdenServicioPagos
4200                            .FechaCreacion = oDoFactOrdenServ.FechaCreacion
4210                            .IdEstadoFacturacion = 1   'Registrado
4220                            .IdOrden = oDoFactOrdenServ.IdOrden
4230                            .idUsuario = oDoAtencion.IdUsuarioAuditoria
4240                            .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
4250                        End With
4260                        If Not oFactServicioPagos.Insertar(oDOFactOrdenServicioPagos) Then
4270                            bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
4280                        End If
4290                      End If
                          'Despachos
4300                      If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(lnIdtipoFinanciamiento, oConexion) = False Then
4310                         lnPrecioUnitario = 0
4320                         oDoCatalogoServicioHosp.PrecioUnitario = 0
4330                         Set oDoCatalogoServicioHosp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarPorId(lnIdServicio, lnIdtipoFinanciamiento)
4340                         If oDoCatalogoServicioHosp.PrecioUnitario = 0 Then
4350                            lnPrecioUnitario = 0
4360                         Else
4370                            lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
4380                         End If
4390                      End If
4400                      With oDoFacturacionServicioDesp
4410                          .cantidad = 1
4420                          .IdOrden = oDoFactOrdenServ.IdOrden
4430                          .idProducto = lnIdServicio
4440                          .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
4450                          .Precio = lnPrecioUnitario
4460                          .total = lnPrecioUnitario
4470                      End With
4480                      If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then
4490                             bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
4500                      End If
                          'Crear Detalle de servicio
4510                      If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(lnIdtipoFinanciamiento, oConexion) = True Then
                             '***********Contado ***********
4520                         With oDoFacturacionServicioPagos
4530                             .cantidad = 1
4540                             .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
4550                             .idProducto = lnIdServicio
4560                             .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
4570                             .Precio = lnPrecioUnitario
4580                             .total = lnPrecioUnitario
4590                         End With
4600                         If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
4610                             bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
4620                         End If
4630                      Else
                             '***********Sis,Soat,exoneracion, convenio ***********
4640                         With oDoFacturacionServicioFinanc
4650                             .CantidadFinanciada = 1
4660                             .FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
4670                             .IdOrden = oDoFactOrdenServ.IdOrden
4680                             .idProducto = lnIdServicio
4690                             .IdTipoFinanciamiento = lnIdtipoFinanciamiento
4700                             .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
4710                             .IdUsuarioAutoriza = oDoAtencion.IdUsuarioAuditoria
4720                             .PrecioFinanciado = lnPrecioUnitario
4730                             .TotalFinanciado = lnPrecioUnitario
4740                             .idFuenteFinanciamiento = lnIdFuenteFinanciamiento
4750                         End With
4760                         If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then
4770                             bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo Terminar
4780                         End If
4790                      End If
                   'End If
4800          End If
4810      End If
          '
          '****Actualiza Dato del Recien Nacido (que está en la Cuenta de la Madre)....tabla: AtencionesNacimiento
          '****Asocia la Cuenta de la madre con la cuenta con el hijo..................tabla: AtencionesHijoMadre
4820      If oAtencionesNacimientos.AsignaNULLidPacienteNacido(oDoPaciente.idPaciente) = False Then
4830         ms_MensajeError = oAtencionesNacimientos.MensajeError: GoTo Terminar
4840      End If
4850      oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
4860      oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
4870      If oAtencionesHijoMadre.Eliminar(oDoAtencionesHijoMadre) = False Then
4880         ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
4890      End If
4900      If lnIdNacimientoSeleccionado > 0 Then
             Dim lnIdAtencionMadre As Long
4910         oDOAtencionNacimiento.idNacimiento = lnIdNacimientoSeleccionado
4920         If oAtencionesNacimientos.SeleccionarPorId(oDOAtencionNacimiento) = True Then
4930            lnIdAtencionMadre = oDOAtencionNacimiento.idAtencion
4940            oDOAtencionNacimiento.idPacienteNacido = oDoPaciente.idPaciente
4950            oDOAtencionNacimiento.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
4960            If oAtencionesNacimientos.modificar(oDOAtencionNacimiento) = True Then
4970               ms_MensajeError = oAtencionesNacimientos.MensajeError: GoTo Terminar
4980            End If
4990            oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
5000            If oAtencionesHijoMadre.SeleccionarPorId(oDoAtencionesHijoMadre) = True Then
5010                 oDoAtencionesHijoMadre.idAtencionMadre = lnIdAtencionMadre
5020                 oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
5030                 If oAtencionesHijoMadre.modificar(oDoAtencionesHijoMadre) = True Then
5040                    ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
5050                 End If
5060            Else
5070                 oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
5080                 oDoAtencionesHijoMadre.idAtencionMadre = lnIdAtencionMadre
5090                 oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
5100                 If oAtencionesHijoMadre.Insertar(oDoAtencionesHijoMadre) = True Then
5110                    ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
5120                 End If
5130            End If
5140         End If
5150      End If

          'datos adicionales
5160      mo_DoAtencionDatosAdicionales.idAtencion = oDoAtencion.idAtencion
          mo_DoAtencionDatosAdicionales.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
5170      If Not oAtencionesDatosAdicionales.modificar(mo_DoAtencionDatosAdicionales) Then
5180          bProcesoOK = False: ms_MensajeError = oAtencionesDatosAdicionales.MensajeError: GoTo Terminar
5190      End If
5200      If mo_DoAtencionDatosAdicionales.idSiaSis > 0 Then
5210         Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "M", oDoAtencion.idCuentaAtencion, "FUA", oConexion, 1345, lcNombrePc, "Cod.Prestacion FUA: " & mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion & " N° Cuenta: " & oDoAtencion.idCuentaAtencion)
5220      End If
          'Episodio Clinico
5230      If sighentidades.EsFecha(oDoAtencion.fechaEgreso, "DD/MM/AAAA") Then
5240         ms_MensajeError = EpisodioClinicoActualizar(oEpisodioClinico, oDoAtencionesEpisodiosCab, oAtencionesEpisodiosCab, _
                                                         oDoAtencionesEpisodiosDet, oAtencionesEpisodiosDet, oDoAtencion, _
                                                         oConexion, sghModificar)
5250         If ms_MensajeError <> "" Then
5260            bProcesoOK = False: GoTo Terminar
5270         End If
5280      End If
          '
          mo_ReglasArchivoClinico.HistoriasClinicasActualizaFmovimiento oConexion, oDoAtencion.idPaciente
5290      ms_MensajeError = ""
          '
Terminar:
5300      If bProcesoOK Then
5310          AdmisionHospModificar = True
5320          oConexion.CommitTrans
5330      Else
5340          If Err.Description <> "" Then
5350             ms_MensajeError = Err.Description & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                                 "Function AdmisionHospModificar", "reglasAdmision.cls")  'debb-02/05/2016
5360          End If
5370          AdmisionHospModificar = False
5380          oConexion.RollbackTrans
                            
5390      End If
                  
5400      oConexion.Close
5410      Set oRsBoletaPagada = Nothing
5420      Set oConexion = Nothing
5430      Set oPaciente = Nothing
5440      Set oCuentaAtencion = Nothing
5450      Set oAtencion = Nothing
5460      Set oDOHistoriaSolicitada = Nothing
5470      Set oHistoriaSolicitada = Nothing
5480      Set oCama = Nothing
5490      Set oHistoria = Nothing
5500      Set oAtencionesDx = Nothing
5510      Set oEstanciaHospitalaria = Nothing
5520      Set oAtencionDiagnostico = Nothing
5530      Set oAtencionesNacimientos = Nothing
5540      Set oFactOrdenesServicio = Nothing
5550      Set oDOfactOrdenesServicio = Nothing
5560      Set oFacturacionServicios = Nothing
5570      Set oDOFacturacionServicios = Nothing
5580      Set lcBuscaParametro = Nothing
5590      Set oFacturacionServicioDespacho = Nothing
5600      Set oFacturacionServicioPagos = Nothing
5610      Set oFacturacionServicioFinan = Nothing
5620      Set oFactServicioPagos = Nothing
5630      Set oDOCama = Nothing
5640      Set oCamas = Nothing
5650      Set oDOCamasMovimientos = Nothing
5660      Set oCamasMovimientos = Nothing
5670      Set oDoCatalogoServicioHosp = Nothing
5680      Set oBuscaDiasPaciente = Nothing
5690      Set oDoFactOrdenServ = Nothing
5700      Set oDoFacturacionServicioPagos = Nothing
5710      Set oDoFacturacionServicioFinanc = Nothing
5720      Set oDoFacturacionServicioDesp = Nothing
5730      Set oDOFactOrdenServicioPagos = Nothing
5740      Set oRsTmp = Nothing
5750      Set oRsTmp1 = Nothing
5760      Set oRsTmp2 = Nothing
5770      Set oSunasaPacientesHistoricos = Nothing
5780      Set oAtencionesDatosAdicionales = Nothing
5790      Set oDOAtencionNacimiento = Nothing
5800      Set oDoAtencionesHijoMadre = Nothing
5810      Set oAtencionesHijoMadre = Nothing
5820      Set oDoAtencionesEpisodiosCab = Nothing
5830      Set oAtencionesEpisodiosCab = Nothing
5840      Set oDoAtencionesEpisodiosDet = Nothing
5850      Set oAtencionesEpisodiosDet = Nothing
End Function
Function AdmisionHospAnular(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, oDoPaciente As DOPaciente, _
                            oDoHistoria As DOHistoriaClinica, lIdTipoNumeracionAnterior As sghTipoNumeracionDeNroHistoria, _
                            oOcupaciones As Collection, oDiagnosticos As Collection, oProcedimientos As Collection, _
                            oExamenes As Collection, oNacimientos As Collection, lIdHistoriaAnterior As Long, _
                            oDOAtencionEmergencia As DOAtencionEmergencia, lTipoAccionAdmision As sghTipoAccionEmergenciaYHospitalizacion, _
                            ldFechaEgresoMedicoAnterior As Date, lbPacienteNN As Boolean, lnIdTablaLISTBARITEMS As Long, _
                            lcNombrePc As String, lcObservacionesAuditoria As String, lnIdNacimientoSeleccionado As Long, _
                            oEpisodioClinico As EpisodioClinico) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim oDOHistoriaSolicitada As New DOHistoriaSolicitada
Dim oHistoriaSolicitada As New HistoriasSolicitadas
Dim oCama As New Camas
Dim bProcesoOK As Boolean
Dim oHistoria As New HistoriasClinicas
Dim oAtencionesDx As New AtencionesDiagnosticos
Dim oEstanciaHospitalaria As New AtencionesEstanciaHosp
Dim oAtencionDiagnostico As New AtencionesDiagnosticos
Dim oAtencionesNacimientos As New AtencionesNacimientos
Dim oFactOrdenesServicio As New FactOrdenServicio
Dim oDOfactOrdenesServicio As New DOFactOrdenServicio
Dim oFacturacionServicios As New FacturacionServicios
Dim oDOFacturacionServicios As New DOFacturacionServicios
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
Dim oFacturacionServicioPagos As New FacturacionServicioPagos
Dim oFacturacionServicioFinan As New FacturacionServicioFin
Dim oFactServicioPagos As New FactOrdenServicioPagos
Dim oDOCama As New DOCama
Dim oCamas As New Camas
Dim oDOCamasMovimientos As New DOCamasMovimientos
Dim oCamasMovimientos As New CamasMovimientos
Dim oRsTmp1 As New Recordset
Dim oRsTmp As New Recordset
Dim oRsTmp2 As New Recordset
Dim oDoFactOrdenServ As New DoFactOrdenServ
Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
Dim oFacturacionServicioFin As New FacturacionServicioFin
Dim oDOAtencionNacimiento As New DOAtencionNacimiento
Dim oDoAtencionesHijoMadre As New DoAtencionesHijoMadre, oAtencionesHijoMadre As New AtencionesHijoMadre
Dim oDoAtencionesEpisodiosCab As New DoAtencionesEpisodiosCab, oAtencionesEpisodiosCab As New AtencionesEpisodiosCab
Dim oDoAtencionesEpisodiosDet As New DoAtencionesEpisodiosDet, oAtencionesEpisodiosDet As New AtencionesEpisodiosDet

    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    Set oHistoriaSolicitada.Conexion = oConexion
    Set oCama.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oAtencionesDx.Conexion = oConexion
    Set oEstanciaHospitalaria.Conexion = oConexion
    Set oAtencionDiagnostico.Conexion = oConexion
    Set oAtencionesNacimientos.Conexion = oConexion
    Set oFactOrdenesServicio.Conexion = oConexion
    Set oFacturacionServicios.Conexion = oConexion
    Set oFacturacionServicioPagos.Conexion = oConexion
    Set oFacturacionServicioFinan.Conexion = oConexion
    Set oFacturacionServicioDespacho.Conexion = oConexion
    Set oFactServicioPagos.Conexion = oConexion
    Set oCamas.Conexion = oConexion
    Set oCamasMovimientos.Conexion = oConexion
    Set oFacturacionServicioFin.Conexion = oConexion
    Set oFacturacionServicioFin.Conexion = oConexion
    Set oAtencionesHijoMadre.Conexion = oConexion
    Set oAtencionesEpisodiosCab.Conexion = oConexion
    Set oAtencionesEpisodiosDet.Conexion = oConexion
    
    oConexion.BeginTrans
    bProcesoOK = True
    
    'Modificar atencion
    If Not oAtencion.modificar(oDoAtencion) Then
        bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "E", oDoAtencion.idAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, lcObservacionesAuditoria)  'ListBarItems.idListItem
    '
    oDOCuentaAtencion.FechaCreacion = CDate(oDoAtencion.FechaIngreso & " " & oDoAtencion.HoraIngreso)
    oDOCuentaAtencion.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
    If Not oCuentaAtencion.modificar(oDOCuentaAtencion) Then
        bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
    End If
    '
    'Desocupa cama
    If Not oCamas.ActualizarIdPaciente(oDoAtencion.idPaciente, 0) Then
        bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
    End If
    '
    'anula Codigo Servicio (cabecera) por Estancia/Consulta de la Cuenta de Atencion
    Set oRsTmp2 = mo_ReglasFacturacion.FactOrdenServicioPagosPorIdCuentaAtencion(oDoAtencion.idCuentaAtencion)
    Set oRsTmp1 = mo_ReglasFacturacion.FactOrdenServicioFiltraPorIdCuenta(oDoAtencion.idCuentaAtencion)
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          If oRsTmp1.Fields!IdPuntoCarga = 9 Or oRsTmp1.Fields!IdPuntoCarga = 10 Then   'Hosp/Emerg
             oDoFactOrdenServ.IdOrden = oRsTmp1.Fields!IdOrden
             If Not oFactOrdenesServicio.SeleccionarPorId(oDoFactOrdenServ) Then
                bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
             End If
             oDoFactOrdenServ.IdEstadoFacturacion = 9    'anulado
             oDoFactOrdenServ.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
             If Not oFactOrdenesServicio.modificar(oDoFactOrdenServ) Then
                bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
             End If
             If Not oFacturacionServicioFin.ActualizaIdEstadoFacturacion(oDoFactOrdenServ.IdOrden, 9) Then
                bProcesoOK = False: ms_MensajeError = oFacturacionServicioFin.MensajeError: GoTo Terminar
             End If
          End If
          oRsTmp1.MoveNext
       Loop
       If oRsTmp2.RecordCount > 0 Then
          oRsTmp2.MoveFirst
          Do While Not oRsTmp2.EOF
             If oRsTmp2.Fields!IdPuntoCarga = 9 Or oRsTmp2.Fields!IdPuntoCarga = 10 Then   'Hosp/Emerg
                oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp2.Fields!idOrdenPago
                If Not oFactServicioPagos.SeleccionarPorId(oDOFactOrdenServicioPagos) Then
                   bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
                End If
                oDOFactOrdenServicioPagos.IdEstadoFacturacion = 9    'anulado
                oDOFactOrdenServicioPagos.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                If Not oFactServicioPagos.modificar(oDOFactOrdenServicioPagos) Then
                   bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
                End If
             End If
             oRsTmp2.MoveNext
          Loop
       End If
    End If
    '
    '****Actualiza Dato del Recien Nacido (que está en la Cuenta de la Madre)....tabla: AtencionesNacimiento
    '****Elimina la Cuenta de la madre con la cuenta con el hijo..................tabla: AtencionesHijoMadre
    If oAtencionesNacimientos.AsignaNULLidPacienteNacido(oDoPaciente.idPaciente) = False Then
       ms_MensajeError = oAtencionesNacimientos.MensajeError: GoTo Terminar
    End If
    oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
    oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
    If oAtencionesHijoMadre.Eliminar(oDoAtencionesHijoMadre) = False Then
       ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
    End If
    '
    'Episodio Clinico
    If sighentidades.EsFecha(oDoAtencion.fechaEgreso, "DD/MM/AAAA") Then
       ms_MensajeError = EpisodioClinicoActualizar(oEpisodioClinico, oDoAtencionesEpisodiosCab, oAtencionesEpisodiosCab, _
                                                   oDoAtencionesEpisodiosDet, oAtencionesEpisodiosDet, oDoAtencion, _
                                                   oConexion, sghEliminar)
       If ms_MensajeError <> "" Then
          bProcesoOK = False: GoTo Terminar
       End If
    End If
    '
    mo_ReglasArchivoClinico.HistoriasClinicasActualizaFmovimiento oConexion, oDoAtencion.idPaciente
    '
    

    ms_MensajeError = ""
    '
Terminar:
    
    If bProcesoOK Then
        AdmisionHospAnular = True
        oConexion.CommitTrans
    Else
        If Err.Description <> "" Then
           ms_MensajeError = Err.Description
        End If
        AdmisionHospAnular = False
        oConexion.RollbackTrans
    End If
            
    oConexion.Close
    Set oConexion = Nothing
    Set oPaciente = Nothing
    Set oCuentaAtencion = Nothing
    Set oAtencion = Nothing
    Set oDOHistoriaSolicitada = Nothing
    Set oHistoriaSolicitada = Nothing
    Set oCama = Nothing
    Set oHistoria = Nothing
    Set oAtencionesDx = Nothing
    Set oEstanciaHospitalaria = Nothing
    Set oAtencionDiagnostico = Nothing
    Set oAtencionesNacimientos = Nothing
    Set oFactOrdenesServicio = Nothing
    Set oDOfactOrdenesServicio = Nothing
    Set oFacturacionServicios = Nothing
    Set oDOFacturacionServicios = Nothing
    Set lcBuscaParametro = Nothing
    Set oFacturacionServicioDespacho = Nothing
    Set oFacturacionServicioPagos = Nothing
    Set oFacturacionServicioFinan = Nothing
    Set oFactServicioPagos = Nothing
    Set oDOCama = Nothing
    Set oCamas = Nothing
    Set oDOCamasMovimientos = Nothing
    Set oCamasMovimientos = Nothing
    Set oDoFactOrdenServ = Nothing
    Set oDOFactOrdenServicioPagos = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set oFacturacionServicioFin = Nothing
    Set oDOAtencionNacimiento = Nothing
    Set oDoAtencionesHijoMadre = Nothing
    Set oAtencionesHijoMadre = Nothing
    Set oDoAtencionesEpisodiosCab = Nothing
    Set oAtencionesEpisodiosCab = Nothing
    Set oDoAtencionesEpisodiosDet = Nothing
    Set oAtencionesEpisodiosDet = Nothing
End Function


Function AdmisionHospEliminar(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, lnIdTablaLISTBARITEMS As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim oCama As New Camas
Dim bProcesoOK As Boolean
Dim oAtencionesDx As New AtencionesDiagnosticos
Dim oEstanciaHospitalaria As New AtencionesEstanciaHosp
Dim oAtencionDiagnostico As New AtencionesDiagnosticos
Dim oAtencionesNacimientos As New AtencionesNacimientos
Dim oFacturacionServicios As New FacturacionServicios
Dim oFactOrdenesServicio As New FactOrdenesServicio
Dim oFacturacionBienesInsumos As New FacturacionBienesInsumos
Dim oFactOrdenesBienesInsumo As New FactOrdenesBienesInsumo

    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oCama.Conexion = oConexion
    Set oAtencionesDx.Conexion = oConexion
    Set oEstanciaHospitalaria.Conexion = oConexion
    Set oAtencionDiagnostico.Conexion = oConexion
    Set oAtencionesNacimientos.Conexion = oConexion
    Set oFacturacionServicios.Conexion = oConexion
    Set oFactOrdenesServicio.Conexion = oConexion
    Set oFacturacionBienesInsumos.Conexion = oConexion
    Set oFactOrdenesBienesInsumo.Conexion = oConexion
    
    oConexion.BeginTrans
    bProcesoOK = True
    
    'Eliminar complicaciones
    If Not oAtencionDiagnostico.EliminarComplicaciones(oDoAtencion.idAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    'Eliminar ocupacion camas
    If Not oEstanciaHospitalaria.EliminarEstanciaHospitalaria(oDoAtencion.idAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    
    'Eliminar diagnosticos
    If Not oAtencionesDx.EliminarDiagnosticosPorAtencion(oDoAtencion.idAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    'Agregar nacimientos
    If oDoAtencion.idTipoServicio = sghHospitalizacion Then
        If Not oAtencionesNacimientos.EliminarNacimientosPorAtencion(oDoAtencion.idAtencion) Then
            bProcesoOK = False: GoTo Terminar
        End If
    End If
    
    'Eliminar Id paciente de las camas ocupadas
    If oDoAtencion.idTipoServicio = sghHospitalizacion Or oDoAtencion.idTipoServicio = sghEmergenciaObservacion Then
        If Not oCama.ActualizarIdPaciente(oDoAtencion.idPaciente, 0) Then
            bProcesoOK = False: GoTo Terminar
        End If
    End If
    
    'Eliminar causa externa de morbilidad para los pacientes de emergencia
    If oDoAtencion.idTipoServicio = sghEmergenciaConsultorios Or oDoAtencion.idTipoServicio = sghEmergenciaObservacion Then
        Dim oAtencionEmergencia As New AtencionesEmergencia
        Set oAtencionEmergencia.Conexion = oConexion
        If Not oAtencionEmergencia.EliminarAtencionEmergenciaPorIdAtencion(oDoAtencion.idAtencion) Then
            bProcesoOK = False: GoTo Terminar
        End If
    End If
    
    If Not oFacturacionServicios.EliminarPorIdAtencion(oDoAtencion.idAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    If Not oFactOrdenesServicio.EliminarPorIdAtencion(oDoAtencion.idAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    If Not oFacturacionBienesInsumos.EliminarPorIdAtencion(oDoAtencion.idAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    If Not oFactOrdenesBienesInsumo.EliminarPorIdAtencion(oDoAtencion.idAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    'Eliminar atencion
    If Not oAtencion.Eliminar(oDoAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    'Eliminar cuenta atención
    If Not oCuentaAtencion.Eliminar(oDOCuentaAtencion) Then
        bProcesoOK = False: GoTo Terminar
    End If
    
    
Terminar:
    If bProcesoOK Then
        AdmisionHospEliminar = True
        oConexion.CommitTrans
    Else
        AdmisionHospEliminar = False
        oConexion.RollbackTrans
    End If
            
    oConexion.Close

End Function

'-------------------------------------------------------------------------------------
'                                       Atenciones
'-------------------------------------------------------------------------------------

Function AtencionesAgregar(oDOCuentaAtencion As DOCuentaAtencion) As Boolean
Dim oConexion As New ADODB.Connection
Dim oAtenciones As New Atenciones
    
    AtencionesAgregar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oAtenciones.Conexion = oConexion
    If oAtenciones.Insertar(oDOCuentaAtencion) Then
        AtencionesAgregar = True
    Else
        ms_MensajeError = oAtenciones.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oAtenciones = Nothing
End Function

Function AtencionesModificar(oDOCuentaAtencion As DOCuentaAtencion) As Boolean
Dim oConexion As New ADODB.Connection
Dim oAtenciones As New Atenciones
    
    AtencionesModificar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oAtenciones.Conexion = oConexion
    If oAtenciones.modificar(oDOCuentaAtencion) Then
        AtencionesModificar = True
    Else
        ms_MensajeError = oAtenciones.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oAtenciones = Nothing
    
End Function

Function AtencionesEliminar(oDOCuentaAtencion As DOCuentaAtencion) As Boolean
Dim oConexion As New ADODB.Connection
Dim oAtenciones As New Atenciones
    
    AtencionesEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oAtenciones.Conexion = oConexion
    If oAtenciones.Eliminar(oDOCuentaAtencion) Then
        AtencionesEliminar = True
    Else
        ms_MensajeError = oAtenciones.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oAtenciones = Nothing
    
End Function

Function AtencionesSeleccionarPorId(IdAtenciones As Long, oConexion As Connection) As DOAtencion
Dim oAtenciones As New Atenciones
Dim oDoAtencion As New DOAtencion
    
    Set AtencionesSeleccionarPorId = Nothing
    Set oAtenciones.Conexion = oConexion
    oDoAtencion.idAtencion = IdAtenciones
    If oAtenciones.SeleccionarPorId(oDoAtencion) Then
        Set AtencionesSeleccionarPorId = oDoAtencion
    End If
    ms_MensajeError = oAtenciones.MensajeError
   
End Function
Function AtencionesFiltrarConsultaExterna(oDoPaciente As DOPaciente, oDoAtencion As DOAtencion, lcFechaIngreso As String, idMedicoIngreso As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Atenciones
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set AtencionesFiltrarConsultaExterna = oTabla.FiltrarConsultaExterna(oDoPaciente, oDoAtencion, lcFechaIngreso, idMedicoIngreso)
    ms_MensajeError = oTabla.MensajeError
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function AtencionesFiltrarEmergencia(oDoPaciente As DOPaciente, oDoAtencion As DOAtencion, lcFechaIngreso As String)
Dim oConexion As New ADODB.Connection
Dim oTabla As New Atenciones
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set AtencionesFiltrarEmergencia = oTabla.FiltrarEmergencia(oDoPaciente, oDoAtencion, lcFechaIngreso)
    ms_MensajeError = oTabla.MensajeError
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function AtencionesFiltrarConsultorioEmergencia(oDoPaciente As DOPaciente, oDoAtencion As DOAtencion)
Dim oConexion As New ADODB.Connection
Dim oTabla As New Atenciones
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set AtencionesFiltrarConsultorioEmergencia = oTabla.FiltrarConsultorioEmergencia(oDoPaciente, oDoAtencion)
    ms_MensajeError = oTabla.MensajeError
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function AtencionesFiltrarObservacionEmergencia(oDoPaciente As DOPaciente, oDoAtencion As DOAtencion) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Atenciones
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set AtencionesFiltrarObservacionEmergencia = oTabla.FiltrarObservacionEmergencia(oDoPaciente, oDoAtencion)
    ms_MensajeError = oTabla.MensajeError
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function AtencionesFiltrarHospitalizacion(oDoPaciente As DOPaciente, oDoAtencion As DOAtencion, lcFechaIngreso As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Atenciones
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set AtencionesFiltrarHospitalizacion = oTabla.FiltrarHospitalizacion(oDoPaciente, oDoAtencion, lcFechaIngreso)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function AtencionesSeleccionarCEPorCuentaPorHistoriaPorApellidosPorServicio(lnHistoriaClinica As Long, lnIdCuentaAtencion As Long, lcApellidoPaterno As String, ldFechaIngreso As String, lnIdServicio As Long, lcDNI As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Atenciones
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set AtencionesSeleccionarCEPorCuentaPorHistoriaPorApellidosPorServicio = oTabla.AtencionesSeleccionarCEPorCuentaPorHistoriaPorApellidosPorServicio(lnHistoriaClinica, lnIdCuentaAtencion, lcApellidoPaterno, ldFechaIngreso, lnIdServicio, lcDNI)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function
Function AtencionesSeleccionarEmergPorCuentaPorHistoriaPorApellidosPorServicio(lnHistoriaClinica As Long, lnIdCuentaAtencion As Long, lcApellidoPaterno As String, ldFechaIngreso As String, lnIdServicio As Long, lcDNI As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Atenciones
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set AtencionesSeleccionarEmergPorCuentaPorHistoriaPorApellidosPorServicio = oTabla.AtencionesSeleccionarEmergPorCuentaPorHistoriaPorApellidosPorServicio(lnHistoriaClinica, lnIdCuentaAtencion, lcApellidoPaterno, ldFechaIngreso, lnIdServicio, lcDNI)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function
Function AtencionesSeleccionarHospPorCuentaPorHistoriaPorApellidosPorServicio(lnHistoriaClinica As Long, lnIdCuentaAtencion As Long, lcApellidoPaterno As String, ldFechaIngreso As String, lnIdServicio As Long, lcDNI As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Atenciones
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set AtencionesSeleccionarHospPorCuentaPorHistoriaPorApellidosPorServicio = oTabla.AtencionesSeleccionarHospPorCuentaPorHistoriaPorApellidosPorServicio(lnHistoriaClinica, lnIdCuentaAtencion, lcApellidoPaterno, ldFechaIngreso, lnIdServicio, lcDNI)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function




Function AtencionesBusquedaDeAtenciones(oDoPaciente As DOPaciente) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oRrsTmp1 As New Recordset
Dim oCommand As New ADODB.Command
Dim oDODiagnostico As New doDiagnostico
Dim mo_AdminFacturacion As New ReglasFacturacion
Dim oConexion As New ADODB.Connection
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 150
        oConexion.Open sighentidades.CadenaConexion
        sWhere = ""
        With oRrsTmp1
              .Fields.Append "IdCuentaAtencion", adInteger, 4, adFldIsNullable
              .Fields.Append "EstadoCuenta", adVarChar, 100, adFldIsNullable
              .Fields.Append "IdPaciente", adInteger, 4, adFldIsNullable
              .Fields.Append "IdAtencion", adInteger, 4, adFldIsNullable
              .Fields.Append "FechaIngreso", adDate, 8, adFldIsNullable
              .Fields.Append "HoraIngreso", adVarChar, 5, adFldIsNullable
              .Fields.Append "NroHistoriaClinica", adInteger, 4, adFldIsNullable
              .Fields.Append "ApellidoPaterno", adVarChar, 40, adFldIsNullable
              .Fields.Append "ApellidoMaterno", adVarChar, 40, adFldIsNullable
              .Fields.Append "PrimerNombre", adVarChar, 40, adFldIsNullable
              .Fields.Append "SegundoNombre", adVarChar, 40, adFldIsNullable
              .Fields.Append "FechaEgreso", adDate, 8, adFldIsNullable
              .Fields.Append "HoraEgreso", adVarChar, 5, adFldIsNullable
              .Fields.Append "IdEstado", adInteger, 4, adFldIsNullable
              .Fields.Append "ServicioIngreso", adVarChar, 100, adFldIsNullable
              .Fields.Append "dTipoServicio", adVarChar, 100, adFldIsNullable
              .Fields.Append "Edad", adInteger, 4, adFldIsNullable
              .Fields.Append "Diagnostico", adVarChar, 100, adFldIsNullable
              .Fields.Append "IdTipoNumeracion", adInteger, 4, adFldIsNullable
              .Fields.Append "IdTipoServicio", adInteger, 4, adFldIsNullable
              .Fields.Append "IdServicioIngreso", adInteger, 4, adFldIsNullable
              .LockType = adLockOptimistic
              .Open
        End With

       sSql = ""
       sWhere = " Atenciones.idEstadoAtencion<>0  and "
       If oDoPaciente.ApellidoPaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoPaterno like '" + oDoPaciente.ApellidoPaterno + "%' and "
       End If
       If oDoPaciente.ApellidoMaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoMaterno like '" + oDoPaciente.ApellidoMaterno + "%' and "
       End If
       If oDoPaciente.PrimerNombre <> "" Then
            sWhere = sWhere + " PAcientes.PrimerNombre like '%" + oDoPaciente.PrimerNombre + "%' and "
       End If
       If oDoPaciente.NroHistoriaClinica <> 0 Then
            sWhere = sWhere + " Pacientes.NroHistoriaClinica = " & oDoPaciente.NroHistoriaClinica & " and "
       End If
       If oDoPaciente.IdTipoNumeracion <> 0 Then
            sWhere = sWhere + " Pacientes.IdTipoNumeracion = " & oDoPaciente.IdTipoNumeracion & " and "
       End If
       If oDoPaciente.idPaciente <> 0 Then
            sWhere = sWhere + " Pacientes.IdPaciente = " & oDoPaciente.idPaciente & " and "
       End If
       If (sWhere <> "") Then
            sSql = sSql + " where " + Left(sWhere, Len(sWhere) - 4)
       End If
       With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "AtencionesFiltra"
             Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 2000, sSql): .Parameters.Append oParameter
             Set oRecordset = .Execute
             Set oRecordset.ActiveConnection = Nothing
     End With
     If oRecordset.RecordCount > 0 Then
        oRecordset.MoveFirst
        Do While Not oRecordset.EOF
              Set oDODiagnostico = mo_AdminFacturacion.DevuelveDxAltaMedica(oRecordset.Fields!idAtencion, oRecordset.Fields!idTipoServicio, oConexion)
              oRrsTmp1.AddNew
              oRrsTmp1.Fields!idCuentaAtencion = oRecordset.Fields!idCuentaAtencion
              oRrsTmp1.Fields!EstadoCuenta = oRecordset.Fields!EstadoCuenta
              oRrsTmp1.Fields!idPaciente = oRecordset.Fields!idPaciente
              oRrsTmp1.Fields!idAtencion = oRecordset.Fields!idAtencion
              oRrsTmp1.Fields!FechaIngreso = oRecordset.Fields!FechaIngreso
              oRrsTmp1.Fields!HoraIngreso = oRecordset.Fields!HoraIngreso
              oRrsTmp1.Fields!NroHistoriaClinica = oRecordset.Fields!NroHistoriaClinica
              oRrsTmp1.Fields!ApellidoPaterno = oRecordset.Fields!ApellidoPaterno
              oRrsTmp1.Fields!ApellidoMaterno = oRecordset.Fields!ApellidoMaterno
              oRrsTmp1.Fields!PrimerNombre = oRecordset.Fields!PrimerNombre
              oRrsTmp1.Fields!SegundoNombre = IIf(IsNull(oRecordset.Fields!SegundoNombre), "", oRecordset.Fields!SegundoNombre)
              If Not IsNull(oRecordset.Fields!fechaEgreso) Then
                    oRrsTmp1.Fields!fechaEgreso = oRecordset.Fields!fechaEgreso
                    oRrsTmp1.Fields!HoraEgreso = oRecordset.Fields!HoraEgreso
              End If
              oRrsTmp1.Fields!IdEstado = oRecordset.Fields!IdEstado
              oRrsTmp1.Fields!ServicioIngreso = oRecordset.Fields!ServicioIngreso
              oRrsTmp1.Fields!dTipoServicio = oRecordset.Fields!dTipoServicio
              oRrsTmp1.Fields!Edad = oRecordset.Fields!Edad
              oRrsTmp1.Fields!IdTipoNumeracion = oRecordset.Fields!IdTipoNumeracion
              oRrsTmp1.Fields!idTipoServicio = oRecordset.Fields!idTipoServicio
              oRrsTmp1.Fields!idServicioIngreso = oRecordset.Fields!idServicioIngreso
              If oDODiagnostico.Descripcion <> "" Then
                 oRrsTmp1.Fields!diagnostico = Left(Trim(oDODiagnostico.CodigoCIE2004) & " " & oDODiagnostico.Descripcion, 100)
              Else
                 oRrsTmp1.Fields!diagnostico = Left(mo_AdminFacturacion.DevuelveDxAltaMedicaTodosDx(oRecordset.Fields!idAtencion, oRecordset.Fields!idTipoServicio, ""), 100)
              End If
              oRrsTmp1.Update
              oRecordset.MoveNext
        Loop
    End If
    oRrsTmp1.Sort = "FechaIngreso desc,HoraIngreso desc"
    Set AtencionesBusquedaDeAtenciones = oRrsTmp1
    oConexion.Close
    ms_MensajeError = ""
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oDODiagnostico = Nothing
    Set mo_AdminFacturacion = Nothing
    Set oConexion = Nothing
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "BusquedaDeAtenciones"
End Function

Function TiposOrigenAtencionSeleccionarViasDeConsultoriosExt() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposOrigenAtencion
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposOrigenAtencionSeleccionarViasDeConsultoriosExt = oTabla.SeleccionarViasDeConsultoriosExternos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function TiposOrigenAtencionSeleccionarViasDeHospitalizacion(lnTipoServicioHospitalizacion As sghTipoServicioHospitalizacion) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposOrigenAtencion
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposOrigenAtencionSeleccionarViasDeHospitalizacion = oTabla.SeleccionarViasDeHospitalizacion(lnTipoServicioHospitalizacion)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function TiposOrigenAtencionSeleccionarViasDeConsEmergencia() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposOrigenAtencion
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposOrigenAtencionSeleccionarViasDeConsEmergencia = oTabla.SeleccionarViasDeConsultorioEmergencia
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function TiposOrigenAtencionSeleccionarViasDeObsEmergencia() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposOrigenAtencion
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposOrigenAtencionSeleccionarViasDeObsEmergencia = oTabla.SeleccionarViasDeObservacionEmergencia
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function TiposOrigenAtencionSeleccionarPorCodigo(sCodigo As String, lIdTipoServicio As Long) As DOTipoOrigenAtencion
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposOrigenAtencion
Dim oDOOrigenAtencion As New DOTipoOrigenAtencion
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    oDOOrigenAtencion.Codigo = sCodigo
    oDOOrigenAtencion.idTipoServicio = lIdTipoServicio
    If oTabla.SeleccionarPorCodigo(oDOOrigenAtencion) Then
        Set TiposOrigenAtencionSeleccionarPorCodigo = oDOOrigenAtencion
    Else
        Set TiposOrigenAtencionSeleccionarPorCodigo = Nothing
    End If
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function TiposDestinoAtencionSeleccionarDestinosDeObsEmergencia() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposDestinosAtencion
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposDestinoAtencionSeleccionarDestinosDeObsEmergencia = oTabla.SeleccionarDestinosDeObsEmergencia
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function TiposDestinoAtencionSeleccionarDestinosDeConsultorioEmergencia() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposDestinosAtencion
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposDestinoAtencionSeleccionarDestinosDeConsultorioEmergencia = oTabla.SeleccionarDestinosDeConsultorioEmergencia
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function TiposDestinoAtencionSeleccionarDestinosDeObservacionEmergencia() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposDestinosAtencion
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposDestinoAtencionSeleccionarDestinosDeObservacionEmergencia = oTabla.SeleccionarDestinosDeObsEmergencia
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function TiposDestinoAtencionSeleccionarDestinosDeHospitalizacion(lnTipoServicioHospitalizacion As sghTipoServicioHospitalizacion) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposDestinosAtencion
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposDestinoAtencionSeleccionarDestinosDeHospitalizacion = oTabla.SeleccionarDestinosDeHospitalizacion(lnTipoServicioHospitalizacion)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function TiposDestinoAtencionSeleccionarDestinosDeConsultoriosExternos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposDestinosAtencion
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposDestinoAtencionSeleccionarDestinosDeConsultoriosExternos = oTabla.SeleccionarDestinosDeConsultoriosExternos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function TiposDestinoAtencionSeleccionarPorCodigo(sCodigo As String, lTipoServicio As Long) As DOTipoDestinoAtencion
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposDestinosAtencion
Dim oDODestinoAtencion As New DOTipoDestinoAtencion
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    oDODestinoAtencion.Codigo = sCodigo
    oDODestinoAtencion.idTipoServicio = lTipoServicio
    If oTabla.SeleccionarPorCodigo(oDODestinoAtencion) Then
        Set TiposDestinoAtencionSeleccionarPorCodigo = oDODestinoAtencion
    Else
        Set TiposDestinoAtencionSeleccionarPorCodigo = Nothing
    End If
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function AtencionesDiagnosticosSeleccionarComplicaciones(lIdAtencion As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New AtencionesDiagnosticos
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set AtencionesDiagnosticosSeleccionarComplicaciones = oTabla.SeleccionarComplicaciones(lIdAtencion)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function AtencionesDiagnosticosSeleccionarPorAtencion(lIdAtencion As Long, IdTipoDiagnostico As Long, oConexion As Connection) As Recordset
    Dim oTabla As New AtencionesDiagnosticos
    Set oTabla.Conexion = oConexion
    Set AtencionesDiagnosticosSeleccionarPorAtencion = oTabla.SeleccionarPorAtencion(lIdAtencion, IdTipoDiagnostico)
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing

End Function

Function AtencionesDiagnosticosSeleccionarPorNroCuenta(lnIdCuentaAtencion As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDiagnosticosXidCuentaAtencion"
        Set oParameter = .CreateParameter("@lnIdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesDiagnosticosSeleccionarPorNroCuenta = oRecordset
   oConexion.Close
   Set oCommand = Nothing
   Set oConexion = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function EstanciaHospitalariaSeleccionarPorAtencion(lIdAtencion As Long, lnSecuenciaMayorA As Long, oConexion As Connection) As Recordset
Dim oTabla As New AtencionesEstanciaHosp
    Set oTabla.Conexion = oConexion
    Set EstanciaHospitalariaSeleccionarPorAtencion = oTabla.SeleccionarPorAtencion(lIdAtencion, lnSecuenciaMayorA)
    ms_MensajeError = oTabla.MensajeError

End Function
Function EstanciaHospitalariaSeleccionarTodosPorCuentaAtencion(lIdCuentaAtencion As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New AtencionesEstanciaHosp
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EstanciaHospitalariaSeleccionarTodosPorCuentaAtencion = oTabla.SeleccionarTodosPorCuentaAtencion(lIdCuentaAtencion)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


Function AtencionesNacimientosSeleccionarPorAtencion(lIdAtencion As Long, oConexion As Connection) As Recordset
Dim oTabla As New AtencionesNacimientos
    Set oTabla.Conexion = oConexion
    Set AtencionesNacimientosSeleccionarPorAtencion = oTabla.SeleccionarPorAtencion(lIdAtencion)
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing

End Function

'*********************************************************
'            CITAS BLOQUEADAS
'*********************************************************
Function CitasBloqueadasAgregar(oDoCitaBloqueada As DOCitaBloqueada) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCitasBloqueadas As New CitasBloqueadas
    
    CitasBloqueadasAgregar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCitasBloqueadas.Conexion = oConexion
    If oCitasBloqueadas.Insertar(oDoCitaBloqueada) Then
        CitasBloqueadasAgregar = True
    Else
        ms_MensajeError = oCitasBloqueadas.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCitasBloqueadas = Nothing
    
End Function

Function CitasBloqueadasModificar(oDoCitaBloqueada As DOCitaBloqueada) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCitasBloqueadasClinica As New CitasBloqueadas
    
    CitasBloqueadasModificar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCitasBloqueadasClinica.Conexion = oConexion
    If oCitasBloqueadasClinica.modificar(oDoCitaBloqueada) Then
        CitasBloqueadasModificar = True
    Else
        ms_MensajeError = oCitasBloqueadasClinica.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCitasBloqueadasClinica = Nothing
    
End Function

Function CitasBloqueadasEliminar(oDoCitaBloqueada As DOCitaBloqueada) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCitasBloqueadasClinica As New CitasBloqueadas
    
    CitasBloqueadasEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCitasBloqueadasClinica.Conexion = oConexion
    If oCitasBloqueadasClinica.Eliminar(oDoCitaBloqueada) Then
        CitasBloqueadasEliminar = True
    Else
        ms_MensajeError = oCitasBloqueadasClinica.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCitasBloqueadasClinica = Nothing
    
End Function

Function CitasBloqueadasEliminarPorUsuario(lIdUsuario As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCitasBloqueadasClinica As New CitasBloqueadas
    
    CitasBloqueadasEliminarPorUsuario = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCitasBloqueadasClinica.Conexion = oConexion
    If oCitasBloqueadasClinica.EliminarPorUsuario(lIdUsuario) Then
        CitasBloqueadasEliminarPorUsuario = True
    Else
        ms_MensajeError = oCitasBloqueadasClinica.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCitasBloqueadasClinica = Nothing
    
End Function

Function CitasBloqueadasSeleccionarPorId(IdCitasBloqueadas As Long) As DOCitaBloqueada
Dim oConexion As New ADODB.Connection
Dim oCitasBloqueadas As New CitasBloqueadas
Dim oDOCitasBloqueadas As New DOCitaBloqueada
    
    Set CitasBloqueadasSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oCitasBloqueadas.Conexion = oConexion
    oDOCitasBloqueadas.idCitaBloqueada = IdCitasBloqueadas
    If oCitasBloqueadas.SeleccionarPorId(oDOCitasBloqueadas) Then
        Set CitasBloqueadasSeleccionarPorId = oDOCitasBloqueadas
    End If
    oConexion.Close

    ms_MensajeError = oCitasBloqueadas.MensajeError
    Set oConexion = Nothing
    Set oCitasBloqueadas = Nothing

End Function



Function EmergenciaCausaExternaMorbilidadSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New AtencionesEmergencia
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EmergenciaCausaExternaMorbilidadSeleccionarTodos = oTabla.CausaExternaMorbilidadSeleccionarTodos()
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EmergenciaClaseAccidenteSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New AtencionesEmergencia
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EmergenciaClaseAccidenteSeleccionarTodos = oTabla.ClaseAccidenteSeleccionarTodos()
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EmergenciaGrupoOcupacionalALABSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New AtencionesEmergencia
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EmergenciaGrupoOcupacionalALABSeleccionarTodos = oTabla.GrupoOcupacionalALABSeleccionarTodos()
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EmergenciaLugarEventoSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New AtencionesEmergencia
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EmergenciaLugarEventoSeleccionarTodos = oTabla.LugarEventoSeleccionarTodos()
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EmergenciaPosicionLesionadoSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New AtencionesEmergencia
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EmergenciaPosicionLesionadoSeleccionarTodos = oTabla.PosicionLesionadoALABSeleccionarTodos()
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EmergenciaRelacionAgresorVictimaSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New AtencionesEmergencia
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EmergenciaRelacionAgresorVictimaSeleccionarTodos = oTabla.RelacionAgresorVictimaSeleccionarTodos()
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EmergenciaSeguridadSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New AtencionesEmergencia
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EmergenciaSeguridadSeleccionarTodos = oTabla.SeguridadSeleccionarTodos()
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EmergenciaTipoAgenteAGANSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New AtencionesEmergencia
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EmergenciaTipoAgenteAGANSeleccionarTodos = oTabla.TipoAgenteAGANSeleccionarTodos()
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EmergenciaTipoEventoSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New AtencionesEmergencia
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EmergenciaTipoEventoSeleccionarTodos = oTabla.TipoEventoSeleccionarTodos()
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EmergenciaTipoTransporteSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New AtencionesEmergencia
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EmergenciaTipoTransporteSeleccionarTodos = oTabla.TipoTransporteSeleccionarTodos()
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


Function EmergenciaTipoVehiculoSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New AtencionesEmergencia
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EmergenciaTipoVehiculoSeleccionarTodos = oTabla.TipoVehiculoSeleccionarTodos()
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


Function EmergenciaUbicacionLesionadoSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New AtencionesEmergencia
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EmergenciaUbicacionLesionadoSeleccionarTodos = oTabla.UbicacionLesionadoSeleccionarTodos()
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function AtencionesEmergenciaSeleccionarPorId(IdAtencionEmergencia As Long, oConexion As Connection) As DOAtencionEmergencia
Dim oAtencionesEmergencia As New AtencionesEmergencia
Dim oDOAtencionEmergencia As New DOAtencionEmergencia
    
    Set AtencionesEmergenciaSeleccionarPorId = Nothing
    Set oAtencionesEmergencia.Conexion = oConexion
    oDOAtencionEmergencia.IdAtencionEmergencia = IdAtencionEmergencia
    If oAtencionesEmergencia.SeleccionarPorId(oDOAtencionEmergencia) Then
        Set AtencionesEmergenciaSeleccionarPorId = oDOAtencionEmergencia
    End If

    ms_MensajeError = oAtencionesEmergencia.MensajeError
    Set oAtencionesEmergencia = Nothing

End Function
Function AtencionHospObtenerUltimoPorIdPaciente(lidPaciente As Long, lTipoServicio As Long) As DOAtencion
Dim oConexion As New ADODB.Connection
Dim oAtenciones As New Atenciones
Dim oDoAtencion As New DOAtencion
    
    
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oAtenciones.Conexion = oConexion
    Set AtencionHospObtenerUltimoPorIdPaciente = oAtenciones.SeleccionarUltimaAtencionHospitalizacion(lidPaciente, lTipoServicio)
    
    oConexion.Close

    ms_MensajeError = oAtenciones.MensajeError
    Set oConexion = Nothing
    Set oAtenciones = Nothing
End Function
Function AtencionesEmergenciaSeleccionarIdPorIdAtencion(idAtencion As Long, oConexion As Connection) As Long
    Dim oAtencionesEmergencia As New AtencionesEmergencia
    AtencionesEmergenciaSeleccionarIdPorIdAtencion = 0
    Set oAtencionesEmergencia.Conexion = oConexion
    
    AtencionesEmergenciaSeleccionarIdPorIdAtencion = oAtencionesEmergencia.SeleccionarIdPorIdAtencion(idAtencion)
        
    ms_MensajeError = oAtencionesEmergencia.MensajeError
    Set oAtencionesEmergencia = Nothing
End Function

Function AtencionesFiltrarPacientesParaIngresarProcedimientos(oDoPaciente As DOPaciente, oDOCuentaAtencion As DOCuentaAtencion) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Atenciones
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set AtencionesFiltrarPacientesParaIngresarProcedimientos = oTabla.FiltrarPacientesParaIngresarProcedimientos(oDoPaciente, oDOCuentaAtencion)
    ms_MensajeError = oTabla.MensajeError
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function CuentasAtencionDatosPacientePorIdCuentaAtencion1(lIdCuentaAtencion As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CuentasAtencion
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set CuentasAtencionDatosPacientePorIdCuentaAtencion1 = oTabla.DatosPacientePorIdCuentaAtencion(lIdCuentaAtencion)
    ms_MensajeError = oTabla.MensajeError
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


Sub Trace(sMensaje As String, NombreDeMetodo As String)
    MsgBox "Mensaje: " & sMensaje + Chr(13) + "Metodo:" & NombreDeMetodo + Chr(13) + "Clase de Negocio: ReglasAdmision", vbInformation, "Capa de negocios, reglas de admisión"
End Sub

Function FuenteFinanciamientoSeleccionarPorIdAtencion(lnIdAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FuenteFinanciamientoSeleccionarPorIdAtencion = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FuenteFinanciamientoSeleccionarPorIdAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FuenteFinanciamientoSeleccionarPorIdAtencion = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function CentrosPobladosDevuelveDptoProvDistritoSegunIdCentroPoblado(lnIdCentroPoblado As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CentrosPobladosDevuelveDptoProvDistritoSegunIdCentroPoblado = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CentrosPobladosDevuelveDptoProvDistritoSegunIdCentroPoblado"
        Set oParameter = .CreateParameter("@IdCentroPoblado", adInteger, adParamInput, 0, lnIdCentroPoblado): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CentrosPobladosDevuelveDptoProvDistritoSegunIdCentroPoblado = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function DevuelveServiciosDelHospital(lcFiltro As String, lcEspecialidades As String, lnTipoEstado As sghTipoEstados, lnOrden As sghOrdenDeServiciosHospital) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lcSql As String, sWhere As String
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    lcSql = ""
    sWhere = ""
    If lnTipoEstado <> sghFiltraAnuladosYactivos Then
        lcSql = lcSql & " and dbo.Servicios.idEstado= " & lnTipoEstado
    End If
    If lcEspecialidades <> "" Then
       lcSql = lcSql & lcEspecialidades
    End If
    If lnOrden = sghPorDescServicio Then
       lcSql = lcSql & " ORDER BY dbo.Servicios.Nombre"
    Else
       lcSql = lcSql & " ORDER BY dbo.TiposServicio.Descripcion,dbo.Servicios.Nombre"
    End If
    lcSql = lcFiltro & lcSql
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveServiciosDelHospitalFiltro"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set DevuelveServiciosDelHospital = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
    
End Function

Function DevuelveEspecialidadesDelHospital(lcFiltro As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
oConexion.CommandTimeout = 300
oConexion.CursorLocation = adUseClient
oConexion.Open sighentidades.CadenaConexion
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveEspecialidadesDelHospitalfiltro"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set DevuelveEspecialidadesDelHospital = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function AtencionesSeleccionarPorIdAtencion(lnIdAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesSeleccionarPorIdAtencion = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarPorIdAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPorIdAtencion = oRecordset
   Set oConexion = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function ValidaEdadMaximaYSexoSegunServicioHosp(lnEdad As Long, lnTipoEdad As Long, lnTipoSexo As Long, _
                                               lnIdServicioHospital As Long, lbMuestraMensaje As Boolean) As Boolean
    If lnIdServicioHospital = 0 Then
       ValidaEdadMaximaYSexoSegunServicioHosp = True
       Exit Function
    End If
    Dim oRsTmp As New Recordset
    Dim lnEdadEnDias As Long
    Dim lcMensaje1 As String
    Dim oReglasComunes As New ReglasComunes
    lnEdadEnDias = 0
    ValidaEdadMaximaYSexoSegunServicioHosp = False
    With oRsTmp
         Set oRsTmp = oReglasComunes.ServiciosSeleccionarXidentificador(lnIdServicioHospital)
         If oRsTmp.RecordCount > 0 Then
            'edad
            Select Case lnTipoEdad
            Case 1   'años
                lnEdadEnDias = 365 * lnEdad
            Case 2   'mes
                lnEdadEnDias = 30 * lnEdad
            Case 3   'dias
                lnEdadEnDias = lnEdad
            Case Else 'horas
                lnEdadEnDias = 1
            End Select
            lcMensaje1 = "Para el Servicio: " & oRsTmp.Fields("nombre").value & Chr(13) & DevuelveRangoEdadesDiaMesAnio(Str(oRsTmp.Fields!minimaEdad), Str(oRsTmp.Fields!maximaEdad))
            If Not (lnEdadEnDias >= oRsTmp.Fields!minimaEdad And lnEdadEnDias <= oRsTmp.Fields!maximaEdad) Then
               If lbMuestraMensaje = True Then
                  MsgBox lcMensaje1, vbInformation, "Mensaje"
               End If
               Exit Function
            End If
            'sexo
            If oRsTmp.Fields!soloTipoSexo <> 3 Then
                If IsNull(oRsTmp.Fields!soloTipoSexo) Or oRsTmp.Fields!soloTipoSexo = 0 Then
                ElseIf oRsTmp.Fields!soloTipoSexo <> lnTipoSexo Then
                     If lbMuestraMensaje = True Then
                        MsgBox "Para el Servicio: " & oRsTmp.Fields!nombre & Chr(13) & "sólo acepta Pacientes con Sexo: " & IIf(oRsTmp.Fields!soloTipoSexo = 1, "Masculino", "Femenino"), vbInformation, "Mensaje"
                     End If
                     Exit Function
                End If
            End If
         Else
            If lbMuestraMensaje = True Then
               MsgBox "No existe el SERVICIO DEL HOSPITAL", vbInformation, "Mensaje"
            End If
            Exit Function
         End If
         oRsTmp.Close
    End With
    Set oRsTmp = Nothing
    ValidaEdadMaximaYSexoSegunServicioHosp = True
End Function

Function PacientesFiltraPorNroDocumentoYtipo(lcNroDocumento As String, lnIdDocIdentidad As Long, _
                                             Optional oConexion1 As Connection) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Pacientes
    If oConexion1 Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        Set oTabla.Conexion = oConexion
    Else
        Set oTabla.Conexion = oConexion1
    End If
    
    Set PacientesFiltraPorNroDocumentoYtipo = oTabla.PacientesFiltraPorNroDocumentoYtipo(lcNroDocumento, lnIdDocIdentidad)
    If oConexion1 Is Nothing Then
       oConexion.Close
    End If
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function





Function AtencionesSeleccionarPorIdAtencionTipoServicio(lnIdAtencion As Long, lnIdTipoServicio As Long, Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesSeleccionarPorIdAtencionTipoServicio = Nothing
    ms_MensajeError = ""
    If oConexion1 Is Nothing Then
     oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
   
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarPorIdAtencionTipoServicio"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, lnIdTipoServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPorIdAtencionTipoServicio = oRecordset
   
      Set oConexion = Nothing
   
    Set oCommand = Nothing
    Set oParameter = Nothing
   Set oConexion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FactOrdenServicioSeleccionarPorIdOrden(lnIdOrden As Long, Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioSeleccionarPorIdOrden"
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioSeleccionarPorIdOrden = oRecordset
   Set oConexion = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
   Set oConexion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function AtencionesSeleccionarPorIdPaciente(lnIdPaciente As Long, lnIdTipoServicio As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesSeleccionarPorIdPaciente = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarPorIdPaciente"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, lnIdTipoServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPorIdPaciente = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
   Set oConexion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function HistoriasClinicasSeleccionarPorFechaEgresoMedico(ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, lcOrdenFiltro As String, ml_TipoServicio As Integer) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String, lcSql As String
    lcSql = ""
    If ml_TipoServicio = 0 Then
           If lcOrdenFiltro = "HC" Then
              lcSql = lcSql & " order by dbo.Pacientes.NroHistoriaClinica"
           ElseIf lcOrdenFiltro = "FECHAEGRESO" Then
              lcSql = lcSql & " order by dbo.Atenciones.fechaEgreso,dbo.Atenciones.horaEgreso"
           Else
              lcSql = lcSql & " order by dbo.Pacientes.ApellidoPaterno,dbo.Pacientes.ApellidoMaterno,dbo.Pacientes.PrimerNombre"
           End If
    Else
           If lcOrdenFiltro = "HC" Then
              lcSql = lcSql & " order by dbo.Pacientes.NroHistoriaClinica"
           Else
              lcSql = lcSql & " order by dbo.Pacientes.ApellidoPaterno,dbo.Pacientes.ApellidoMaterno,dbo.Pacientes.PrimerNombre"
           End If
    End If
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        If ml_TipoServicio = 0 Then
           .CommandText = "AtencionesSeleccionarHospitalizadosPorFechasDeEgresoMedico"
        Else
           .CommandText = "AtencionesSeleccionarEmergenciaPorFechasDeEgresoMedico"
        End If
        Set oParameter = .CreateParameter("@FechaEgrIni", adVarChar, adParamInput, 10, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrFin", adVarChar, adParamInput, 10, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set HistoriasClinicasSeleccionarPorFechaEgresoMedico = oRecordset
   Set oCommand = Nothing
   Set oConexion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function AtencionesSeleccionarPorFechasDeIngresoYtipoServicio(ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, lnIdTipoServicio As Integer) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesSeleccionarPorFechasDeIngresoYtipoServicio = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarPorFechasDeIngresoYtipoServicio"
        Set oParameter = .CreateParameter("@FechaEgrIni", adDate, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrFin", adDate, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, lnIdTipoServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPorFechasDeIngresoYtipoServicio = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function AdmisionAlojadosAgregar(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, oDoPaciente As DOPaciente, oDOEstanciaHospitalaria As DOEstanciaHospitalaria, lbPacienteNN As Boolean, lnIdTablaLISTBARITEMS As Long, lcNombrePc As String, lcObservacionesAuditoria As String, lnIdNacimientoSeleccionado As Long, mo_DoAtencionDatosAdicionales As DoAtencionDatosAdicionales) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim oHistoria As New HistoriasClinicas
Dim oCamas As New Camas
Dim oDoCamas As New DOCama
Dim oEstanciaHospitalaria As New AtencionesEstanciaHosp
Dim oCamasMovimientos As New CamasMovimientos
Dim oDOCamasMovimientos As New DOCamasMovimientos
Dim oAtencionesDatosAdicionales As New AtencionesDatosAdicionales
Dim oDOAtencionNacimiento As New DOAtencionNacimiento, oAtencionesNacimientos As New AtencionesNacimientos
Dim oDoAtencionesHijoMadre As New DoAtencionesHijoMadre, oAtencionesHijoMadre As New AtencionesHijoMadre
Dim bProcesoOK As Boolean
Dim bPacienteNuevo As Boolean
Dim oRsTmp As New Recordset

    oConexion.Open sighentidades.CadenaConexion
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    Set oCamas.Conexion = oConexion
    Set oEstanciaHospitalaria.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oCamasMovimientos.Conexion = oConexion
    Set oAtencionesDatosAdicionales.Conexion = oConexion
    Set oAtencionesHijoMadre.Conexion = oConexion
    '
    oConexion.BeginTrans
    bProcesoOK = True
    '
    bPacienteNuevo = False
    If oDoPaciente.idPaciente = 0 Then
        'Paciente nuevo
        bPacienteNuevo = True
        'crea la historia clinica, si es difrenete a manual
         oDoPaciente.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoPaciente.IdTipoNumeracion)
        'Agregar paciente
        If Not PacientesAgregar(oDoPaciente) Then
            bProcesoOK = False: ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
        End If
    Else
        '********Paciente antiguo
        If Not PacientesModificar(oDoPaciente, IIf(lbPacienteNN = True, False, True)) Then
            bProcesoOK = False: ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
        End If
    End If
    '
    'Agregar cuenta atención
    oDOCuentaAtencion.FechaCreacion = CDate(oDoAtencion.FechaIngreso & " " & oDoAtencion.HoraIngreso)
    oDOCuentaAtencion.idPaciente = oDoPaciente.idPaciente
    If Not oCuentaAtencion.Insertar(oDOCuentaAtencion) Then
        bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
    End If
    '
    'Agregar atencion
    oDoAtencion.idPaciente = oDoPaciente.idPaciente
    oDoAtencion.idCuentaAtencion = oDOCuentaAtencion.idCuentaAtencion
    If Not oAtencion.Insertar(oDoAtencion) Then
        bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "A", oDoAtencion.idAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, lcObservacionesAuditoria)    'ListBarItems.idListItem
    '
    'Agregar estancia en el servicio
    oDOEstanciaHospitalaria.idAtencion = oDoAtencion.idAtencion
    If Not oEstanciaHospitalaria.Insertar(oDOEstanciaHospitalaria) Then
        bProcesoOK = False: ms_MensajeError = oEstanciaHospitalaria.MensajeError: GoTo Terminar
    End If
    '
    'Actualizar cama
    If Not oCamas.ActualizarIdPaciente(oDoPaciente.idPaciente, oDOEstanciaHospitalaria.IdCama) Then
        bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
    End If

    '-----Chequea Si la "Cama de Ingreso" no pertenece al "Servicio de Ingreso"
    '-----de ser asi la "cama" pasará  al "servicio" como PRESTAMO en forma automatica
    If oDoAtencion.IdCamaIngreso > 0 Then
        oDoCamas.IdCama = oDoAtencion.IdCamaIngreso
        If Not oCamas.SeleccionarPorId(oDoCamas) Then
            bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
        End If
        If oDoCamas.IdServicioUbicacionActual <> oDoAtencion.idServicioIngreso Then
           oDoCamas.IdServicioUbicacionActual = oDoAtencion.idServicioIngreso
           oDoCamas.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
           If Not oCamas.modificar(oDoCamas) Then
                bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
           End If
           oDOCamasMovimientos.IdCama = oDoCamas.IdCama
           oDOCamasMovimientos.IdFechaIngreso = lcBuscaParametro.RetornaFechaHoraServidorSQL
           oDOCamasMovimientos.IdServicio = oDoCamas.IdServicioUbicacionActual
           oDOCamasMovimientos.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
           If Not oCamasMovimientos.Insertar(oDOCamasMovimientos) Then
                bProcesoOK = False: ms_MensajeError = oCamasMovimientos.MensajeError: GoTo Terminar
           End If
        End If
    End If
    '
    '****Actualiza Dato del Recien Nacido (que está en la Cuenta de la Madre)....tabla: AtencionesNacimiento
    '****Asocia la Cuenta de la madre con la cuenta con el hijo..................tabla: AtencionesHijoMadre
    If oAtencionesNacimientos.AsignaNULLidPacienteNacido(oDoPaciente.idPaciente) = False Then
       ms_MensajeError = oAtencionesNacimientos.MensajeError: GoTo Terminar
    End If
    oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
    oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
    If oAtencionesHijoMadre.Eliminar(oDoAtencionesHijoMadre) = False Then
       ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
    End If
    If lnIdNacimientoSeleccionado > 0 Then
       Dim lnIdAtencionMadre As Long
       oDOAtencionNacimiento.idNacimiento = lnIdNacimientoSeleccionado
       If oAtencionesNacimientos.SeleccionarPorId(oDOAtencionNacimiento) = True Then
          lnIdAtencionMadre = oDOAtencionNacimiento.idAtencion
          oDOAtencionNacimiento.idPacienteNacido = oDoPaciente.idPaciente
          oDOAtencionNacimiento.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
          If oAtencionesNacimientos.modificar(oDOAtencionNacimiento) = True Then
             ms_MensajeError = oAtencionesNacimientos.MensajeError: GoTo Terminar
          End If
          oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
          If oAtencionesHijoMadre.SeleccionarPorId(oDoAtencionesHijoMadre) = True Then
               oDoAtencionesHijoMadre.idAtencionMadre = lnIdAtencionMadre
               oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
               If oAtencionesHijoMadre.modificar(oDoAtencionesHijoMadre) = True Then
                  ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
               End If
          Else
               oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
               oDoAtencionesHijoMadre.idAtencionMadre = lnIdAtencionMadre
               oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
               If oAtencionesHijoMadre.Insertar(oDoAtencionesHijoMadre) = True Then
                  ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
               End If
          End If
       End If
    End If

    'datos adicionales
    mo_DoAtencionDatosAdicionales.idAtencion = oDoAtencion.idAtencion
    If Not oAtencionesDatosAdicionales.Insertar(mo_DoAtencionDatosAdicionales) Then
        bProcesoOK = False: ms_MensajeError = oAtencionesDatosAdicionales.MensajeError: GoTo Terminar
    End If
    '
    ms_MensajeError = ""
Terminar:
    If bProcesoOK Then
        AdmisionAlojadosAgregar = True
        oConexion.CommitTrans
    Else
        AdmisionAlojadosAgregar = False
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPaciente = Nothing
    Set oCuentaAtencion = Nothing
    Set oAtencion = Nothing
    Set oHistoria = Nothing
    Set oCamas = Nothing
    Set oEstanciaHospitalaria = Nothing
    Set oCamasMovimientos = Nothing
    Set oAtencionesDatosAdicionales = Nothing
    Set oRsTmp = Nothing
    Set oDOAtencionNacimiento = Nothing
    Set oDoAtencionesHijoMadre = Nothing
    Set oAtencionesHijoMadre = Nothing
End Function


Function AdmisionAlojadosModificar(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, oDoPaciente As DOPaciente, oDOEstanciaHospitalaria As DOEstanciaHospitalaria, lbPacienteNN As Boolean, lnIdTablaLISTBARITEMS As Long, lcNombrePc As String, lcObservacionesAuditoria As String, lnIdNacimientoSeleccionado As Long, mo_DoAtencionDatosAdicionales As DoAtencionDatosAdicionales) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim oHistoria As New HistoriasClinicas
Dim oCamas As New Camas
Dim oDOCama As New DOCama
Dim oEstanciaHospitalaria As New AtencionesEstanciaHosp
Dim oDOCamasMovimientos As New DOCamasMovimientos
Dim oCamasMovimientos As New CamasMovimientos
Dim oDoFactOrdenServ As New DoFactOrdenServ
Dim oBuscaDiasPaciente As New SIGHDatos.Parametros
Dim oFactOrdenesServicio As New FactOrdenServicio
Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
Dim oFacturacionServicioPagos As New FacturacionServicioPagos
Dim oFacturacionServicioFinan As New FacturacionServicioFin
Dim oFactServicioPagos As New FactOrdenServicioPagos
Dim oAtencionesDatosAdicionales As New AtencionesDatosAdicionales
Dim oDOAtencionNacimiento As New DOAtencionNacimiento, oAtencionesNacimientos As New AtencionesNacimientos
Dim oDoAtencionesHijoMadre As New DoAtencionesHijoMadre, oAtencionesHijoMadre As New AtencionesHijoMadre
Dim bProcesoOK As Boolean
Dim bPacienteNuevo As Boolean
Dim lnIdServicio As Long
Dim lnIdtipoFinanciamiento As Long
Dim lnPrecioUnitario As Double
Dim oRsTmp As New Recordset
Dim oRsTmp1 As New Recordset
Dim lcHoraEstanciaMax As String

    oConexion.Open sighentidades.CadenaConexion
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    Set oCamas.Conexion = oConexion
    Set oEstanciaHospitalaria.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oCamasMovimientos.Conexion = oConexion
    Set oFactOrdenesServicio.Conexion = oConexion
    Set oFactServicioPagos.Conexion = oConexion
    Set oFacturacionServicioDespacho.Conexion = oConexion
    Set oFacturacionServicioPagos.Conexion = oConexion
    Set oFacturacionServicioFinan.Conexion = oConexion
    Set oAtencionesDatosAdicionales.Conexion = oConexion
    Set oAtencionesHijoMadre.Conexion = oConexion
    '
    oConexion.BeginTrans
    bProcesoOK = True
    '
    lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
    bPacienteNuevo = False
    '********Paciente antiguo
    If Not PacientesModificar(oDoPaciente, IIf(lbPacienteNN = True, False, True)) Then
        bProcesoOK = False: ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
    End If
    '
    'atencion
    oDoAtencion.idPaciente = oDoPaciente.idPaciente
    If Not oAtencion.modificar(oDoAtencion) Then
        bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "M", oDoAtencion.idAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, lcObservacionesAuditoria)     'ListBarItems.idListItem
    '
    'cuenta atención
    If Not oCuentaAtencion.modificar(oDOCuentaAtencion) Then
        bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
    End If
    '
    'Actualizar cama
    If oDoAtencion.IdEstadoAtencion = sghEstadoTabla.sghAnulado Then
        'Eliminar Id paciente de las camas ocupadas
        If Not oCamas.ActualizarIdPaciente(oDoAtencion.idPaciente, 0) Then
            bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
        End If
    Else
        If Not oCamas.ActualizarIdPaciente(oDoPaciente.idPaciente, oDOEstanciaHospitalaria.IdCama) Then
            bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
        End If
    End If
    '
    'estancia en el servicio
    If oDoAtencion.fechaEgreso <> 0 Then
        oDOEstanciaHospitalaria.idAtencion = oDoAtencion.idAtencion
        If Not oEstanciaHospitalaria.modificar(oDOEstanciaHospitalaria) Then
            bProcesoOK = False: ms_MensajeError = oEstanciaHospitalaria.MensajeError: GoTo Terminar
        End If
        '
    End If
    '
    '***************daniel barrantes**************
    '***************Si el Paciente tiene ALTA MEDICA se genera
    '***************un SERVICIO a su CUENTA por los 'DIAS DE ESTANCIA'
    If sighentidades.EsFecha(oDoAtencion.FechaEgresoAdministrativo, "DD/MM/AAAA") Then
        'Genera Facturacion por ESTADIA
        'If ldFechaEgresoMedicoAnterior <> oDoAtencion.FechaEgreso Then
            Dim lnDiasEstancia As Integer: Dim lnHorasEstancia As Integer
            Dim lcCodServicio As String
            
            'Eliminar Id paciente de las camas ocupadas
            If Not oCamas.ActualizarIdPaciente(oDoAtencion.idPaciente, 0) Then
                bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
            End If

            lnIdtipoFinanciamiento = oDoAtencion.IdFormaPago
            'Dias de Estancia
            lnDiasEstancia = oBuscaDiasPaciente.DiasDelPacienteEnHospitalizacionEmergencia(oDoAtencion.FechaIngreso, oDoAtencion.HoraIngreso, oDoAtencion.FechaEgresoAdministrativo, oDoAtencion.HoraEgresoAdministrativo, lcHoraEstanciaMax)
            'Horas de Estancia
            lnHorasEstancia = sighentidades.HorasDelPacienteEnHospitalizacionEmergencia(oDoAtencion.FechaIngreso, oDoAtencion.HoraIngreso, oDoAtencion.FechaEgresoAdministrativo, oDoAtencion.HoraEgresoAdministrativo)
            'lcCodServicio = "99221"
            lnIdServicio = Val(lcBuscaParametro.SeleccionaFilaParametro(202))   '4590
            If lnIdServicio > 0 Then
                'elimina Codigo Servicio (cabecera/detalle) por Estancia de la Cuenta de Atencion
                Set oRsTmp = mo_ReglasFacturacion.FactOrdenServicioFiltraPorIdCuenta(oDoAtencion.idCuentaAtencion)
                If oRsTmp.RecordCount > 0 Then
                   oRsTmp.MoveFirst
                   Do While Not oRsTmp.EOF
                      If oRsTmp.Fields!IdPuntoCarga = 9 Then   'solo Hospitalizacion
                         oDoFactOrdenServ.IdOrden = oRsTmp.Fields!IdOrden
                         oDoFacturacionServicioFinanc.IdOrden = oRsTmp.Fields!IdOrden
                         oDoFacturacionServicioDesp.IdOrden = oRsTmp.Fields!IdOrden
                         Set oRsTmp1 = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrden(oRsTmp.Fields!IdOrden, oConexion)
                         If oRsTmp1.RecordCount > 0 Then
                           If oRsTmp1.Fields!IdEstadoFacturacion = 1 Then
                                oDOFactOrdenServicioPagos.IdOrden = oRsTmp1.Fields!IdOrden
                                oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp1.Fields!idOrdenPago
                                oDoFacturacionServicioPagos.idOrdenPago = oRsTmp1.Fields!idOrdenPago
                                If Not oFacturacionServicioPagos.Eliminar(oDoFacturacionServicioPagos) Then
                                   bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
                                End If
                                If Not oFactServicioPagos.Eliminar(oDOFactOrdenServicioPagos) Then
                                   bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
                                End If
                           End If
                         End If
                         If Not oFacturacionServicioFinan.Eliminar(oDoFacturacionServicioFinanc) Then
                            bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo Terminar
                         End If
                         If Not oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then
                            bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
                         End If
                         If Not oFactOrdenesServicio.Eliminar(oDoFactOrdenServ) Then
                            bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
                         End If
                      End If
                      oRsTmp.MoveNext
                   Loop
                End If
                oRsTmp.Close
                Set oRsTmp = Nothing
                '********cabecera/detalle
                 lnPrecioUnitario = 0
                 lnIdtipoFinanciamiento = oDoAtencion.IdFormaPago
                 oDoCatalogoServicioHosp.PrecioUnitario = 0
                 Set oDoCatalogoServicioHosp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarPorId(lnIdServicio, lnIdtipoFinanciamiento)
                 If oDoCatalogoServicioHosp.PrecioUnitario > 0 Then
                     lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
                 End If
'                 If lnPrecioUnitario > 0 Then
                     'Crear orden de servicio
                     With oDoFactOrdenServ
                         .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL()
                         .idCuentaAtencion = oDoAtencion.idCuentaAtencion
                         .IdServicioPaciente = oDoAtencion.IdServicioEgreso
                         .IdEstadoFacturacion = 1               'Registrado
                         .idPaciente = oDoAtencion.idPaciente
                         .IdPuntoCarga = IIf(oDoAtencion.idTipoServicio = sghHospitalizacion, 9, 10)
                         .IdTipoFinanciamiento = lnIdtipoFinanciamiento
                         .idFuenteFinanciamiento = oDoAtencion.idFuenteFinanciamiento
                         .idUsuario = oDoAtencion.IdUsuarioAuditoria
                         .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                         .fechaDespacho = .FechaCreacion
                         .IdUsuarioDespacho = oDoAtencion.IdUsuarioAuditoria
                     End With
                     If Not oFactOrdenesServicio.Insertar(oDoFactOrdenServ) Then
                         bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
                     End If
    
                     If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(lnIdtipoFinanciamiento, oConexion) = True Then
                       '***********Contado ***********
                       With oDOFactOrdenServicioPagos
                           .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL()
                           .IdEstadoFacturacion = 1   'Registrado
                           .IdOrden = oDoFactOrdenServ.IdOrden
                           .idUsuario = oDoAtencion.IdUsuarioAuditoria
                           .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                       End With
                       If Not oFactServicioPagos.Insertar(oDOFactOrdenServicioPagos) Then
                           bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
                       End If
                     End If
                     'Crear Detalle de servicio
                     With oDoFacturacionServicioDesp
                         .cantidad = lnDiasEstancia
                         .IdOrden = oDoFactOrdenServ.IdOrden
                         .idProducto = lnIdServicio
                         .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                         .Precio = lnPrecioUnitario
                         .total = lnDiasEstancia * lnPrecioUnitario
                     End With
                     If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then
                            bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
                     End If
                     If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(lnIdtipoFinanciamiento, oConexion) = True Then
                        '***********Contado ***********
                        With oDoFacturacionServicioPagos
                            .cantidad = lnDiasEstancia
                            .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
                            .idProducto = lnIdServicio
                            .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                            .Precio = lnPrecioUnitario
                            .total = lnDiasEstancia * lnPrecioUnitario
                        End With
                        If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
                            bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
                        End If
                     Else
                        '***********Sis,Soat,exoneracion, convenio ***********
                        With oDoFacturacionServicioFinanc
                            .CantidadFinanciada = lnDiasEstancia
                            .FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
                            .IdOrden = oDoFactOrdenServ.IdOrden
                            .idProducto = lnIdServicio
                            .IdTipoFinanciamiento = lnIdtipoFinanciamiento
                            .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                            .IdUsuarioAutoriza = oDoAtencion.IdUsuarioAuditoria
                            .PrecioFinanciado = lnPrecioUnitario
                            .TotalFinanciado = lnDiasEstancia * lnPrecioUnitario
                            .idFuenteFinanciamiento = oDoAtencion.idFuenteFinanciamiento
                        End With
                        If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then
                            bProcesoOK = False: GoTo Terminar
                        End If
                     End If
                 'End If
            End If
        'End If
    End If
    '-----Chequea Si la "Cama de Ingreso" no pertenece al "Servicio de Ingreso"
    '-----de ser asi la "cama" pasará  al "servicio" como PRESTAMO en forma automatica
    If oDoAtencion.IdCamaIngreso > 0 Then
        oDOCama.IdCama = oDoAtencion.IdCamaIngreso
        If Not oCamas.SeleccionarPorId(oDOCama) Then
            bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
        End If
        If oDOCama.IdServicioUbicacionActual <> oDoAtencion.idServicioIngreso Then
           oDOCama.IdServicioUbicacionActual = oDoAtencion.idServicioIngreso
           oDOCama.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
           If Not oCamas.modificar(oDOCama) Then
                bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
           End If
           oDOCamasMovimientos.IdCama = oDOCama.IdCama
           oDOCamasMovimientos.IdFechaIngreso = lcBuscaParametro.RetornaFechaHoraServidorSQL
           oDOCamasMovimientos.IdServicio = oDOCama.IdServicioUbicacionActual
           oDOCamasMovimientos.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
           If Not oCamasMovimientos.Insertar(oDOCamasMovimientos) Then
                bProcesoOK = False: ms_MensajeError = oCamasMovimientos.MensajeError: GoTo Terminar
           End If
        End If
    End If
    '-----Chequea Si la "Cama de Egreso" no pertenece al "Servicio de Egreso"
    '-----de ser asi la "cama" pasará  al "servicio" como PRESTAMO en forma automatica
    If oDoAtencion.IdCamaEgreso > 0 Then
        oDOCama.IdCama = oDoAtencion.IdCamaEgreso
        If Not oCamas.SeleccionarPorId(oDOCama) Then
            bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
        End If
        If oDOCama.IdServicioUbicacionActual <> oDoAtencion.IdServicioEgreso Then
           oDOCama.IdServicioUbicacionActual = oDoAtencion.IdServicioEgreso
           oDOCama.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
           If Not oCamas.modificar(oDOCama) Then
                bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
           End If
           oDOCamasMovimientos.IdCama = oDOCama.IdCama
           oDOCamasMovimientos.IdFechaIngreso = lcBuscaParametro.RetornaFechaHoraServidorSQL
           oDOCamasMovimientos.IdServicio = oDOCama.IdServicioUbicacionActual
           oDOCamasMovimientos.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
           If Not oCamasMovimientos.Insertar(oDOCamasMovimientos) Then
                bProcesoOK = False: ms_MensajeError = oCamasMovimientos.MensajeError: GoTo Terminar
           End If
        End If
    End If
    '
    '****Actualiza Dato del Recien Nacido (que está en la Cuenta de la Madre)....tabla: AtencionesNacimiento
    '****Asocia la Cuenta de la madre con la cuenta con el hijo..................tabla: AtencionesHijoMadre
    If oAtencionesNacimientos.AsignaNULLidPacienteNacido(oDoPaciente.idPaciente) = False Then
       ms_MensajeError = oAtencionesNacimientos.MensajeError: GoTo Terminar
    End If
    oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
    oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
    If oAtencionesHijoMadre.Eliminar(oDoAtencionesHijoMadre) = False Then
       ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
    End If
    If lnIdNacimientoSeleccionado > 0 Then
       Dim lnIdAtencionMadre As Long
       oDOAtencionNacimiento.idNacimiento = lnIdNacimientoSeleccionado
       If oAtencionesNacimientos.SeleccionarPorId(oDOAtencionNacimiento) = True Then
          lnIdAtencionMadre = oDOAtencionNacimiento.idAtencion
          oDOAtencionNacimiento.idPacienteNacido = oDoPaciente.idPaciente
          oDOAtencionNacimiento.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
          If oAtencionesNacimientos.modificar(oDOAtencionNacimiento) = True Then
             ms_MensajeError = oAtencionesNacimientos.MensajeError: GoTo Terminar
          End If
          oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
          If oAtencionesHijoMadre.SeleccionarPorId(oDoAtencionesHijoMadre) = True Then
               oDoAtencionesHijoMadre.idAtencionMadre = lnIdAtencionMadre
               oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
               If oAtencionesHijoMadre.modificar(oDoAtencionesHijoMadre) = True Then
                  ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
               End If
          Else
               oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
               oDoAtencionesHijoMadre.idAtencionMadre = lnIdAtencionMadre
               oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
               If oAtencionesHijoMadre.Insertar(oDoAtencionesHijoMadre) = True Then
                  ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
               End If
          End If
       End If
    End If

    'datos adicionales
    mo_DoAtencionDatosAdicionales.idAtencion = oDoAtencion.idAtencion
    If Not oAtencionesDatosAdicionales.modificar(mo_DoAtencionDatosAdicionales) Then
        bProcesoOK = False: ms_MensajeError = oAtencionesDatosAdicionales.MensajeError: GoTo Terminar
    End If
    '
    ms_MensajeError = ""
Terminar:

    If bProcesoOK Then
        AdmisionAlojadosModificar = True
        oConexion.CommitTrans
    Else
        AdmisionAlojadosModificar = False
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPaciente = Nothing
    Set oCuentaAtencion = Nothing
    Set oAtencion = Nothing
    Set oHistoria = Nothing
    Set oCamas = Nothing
    Set oEstanciaHospitalaria = Nothing
    Set oCamasMovimientos = Nothing
    Set oDoFactOrdenServ = Nothing
    Set oBuscaDiasPaciente = Nothing
    Set oFactOrdenesServicio = Nothing
    Set oDoCatalogoServicioHosp = Nothing
    Set oDoFactOrdenServ = Nothing
    Set oDoFacturacionServicioPagos = Nothing
    Set oDoFacturacionServicioFinanc = Nothing
    Set oDoFacturacionServicioDesp = Nothing
    Set oDOFactOrdenServicioPagos = Nothing
    Set oFacturacionServicioDespacho = Nothing
    Set oFacturacionServicioPagos = Nothing
    Set oFacturacionServicioFinan = Nothing
    Set oFactServicioPagos = Nothing
    Set oAtencionesDatosAdicionales = Nothing
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    Set oDOAtencionNacimiento = Nothing
    Set oDoAtencionesHijoMadre = Nothing
    Set oAtencionesHijoMadre = Nothing
    Set oAtencionesNacimientos = Nothing
End Function

Function AdmisionAlojadosAnular(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, oDoPaciente As DOPaciente, oDOEstanciaHospitalaria As DOEstanciaHospitalaria, lbPacienteNN As Boolean, lnIdTablaLISTBARITEMS As Long, lcNombrePc As String, lcObservacionesAuditoria As String, lnIdNacimientoSeleccionado As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim oHistoria As New HistoriasClinicas
Dim oCamas As New Camas
Dim oDOCama As New DOCama
Dim oEstanciaHospitalaria As New AtencionesEstanciaHosp
Dim oDOCamasMovimientos As New DOCamasMovimientos
Dim oCamasMovimientos As New CamasMovimientos
Dim oDOAtencionNacimiento As New DOAtencionNacimiento, oAtencionesNacimientos As New AtencionesNacimientos
Dim oDoAtencionesHijoMadre As New DoAtencionesHijoMadre, oAtencionesHijoMadre As New AtencionesHijoMadre
Dim bProcesoOK As Boolean
Dim bPacienteNuevo As Boolean
Dim oRsTmp As New Recordset

    oConexion.Open sighentidades.CadenaConexion
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    Set oCamas.Conexion = oConexion
    Set oEstanciaHospitalaria.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oCamasMovimientos.Conexion = oConexion
    Set oAtencionesNacimientos.Conexion = oConexion
    Set oAtencionesHijoMadre.Conexion = oConexion
    
    '
    oConexion.BeginTrans
    bProcesoOK = True
    '
    bPacienteNuevo = False
    '
    'atencion
    If Not oAtencion.modificar(oDoAtencion) Then
        bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "E", oDoAtencion.idAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, lcObservacionesAuditoria)     'ListBarItems.idListItem
    '
    'cuenta atención
    If Not oCuentaAtencion.modificar(oDOCuentaAtencion) Then
        bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
    End If
    '
    'Desocupa cama
    If Not oCamas.ActualizarIdPaciente(oDoAtencion.idPaciente, 0) Then
        bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
    End If
    '
    '****Actualiza Dato del Recien Nacido (que está en la Cuenta de la Madre)....tabla: AtencionesNacimiento
    '****Elimina la Cuenta de la madre con la cuenta con el hijo..................tabla: AtencionesHijoMadre
    If oAtencionesNacimientos.AsignaNULLidPacienteNacido(oDoPaciente.idPaciente) = False Then
       ms_MensajeError = oAtencionesNacimientos.MensajeError: GoTo Terminar
    End If
    oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
    oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoPaciente.IdUsuarioAuditoria
    If oAtencionesHijoMadre.Eliminar(oDoAtencionesHijoMadre) = False Then
       ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
    End If
    
    
    ms_MensajeError = ""
    '
Terminar:
    If bProcesoOK Then
        AdmisionAlojadosAnular = True
        oConexion.CommitTrans
    Else
        AdmisionAlojadosAnular = False
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPaciente = Nothing
    Set oCuentaAtencion = Nothing
    Set oAtencion = Nothing
    Set oHistoria = Nothing
    Set oCamas = Nothing
    Set oEstanciaHospitalaria = Nothing
    Set oCamasMovimientos = Nothing
    Set oRsTmp = Nothing
    Set oDOAtencionNacimiento = Nothing
    Set oDoAtencionesHijoMadre = Nothing
    Set oAtencionesHijoMadre = Nothing
    Set oAtencionesNacimientos = Nothing
End Function

'debb-05/12/2015
Function BuscaSiEstaHospitalizado(lnIdPaciente As Long, oConexion As Connection, _
                                  lnTipoServicioLlamante As sghTipoServicio) As Boolean
    
    Dim oRsBuscaAtencPac As New Recordset
    Dim lcEstadosCta As String
    Dim lcEstadoActual As String
    Dim lnTotalDeuda As Double
    
    lcEstadosCta = ".4.5.9.11.13.14"
    BuscaSiEstaHospitalizado = False
    'Hospitalizacion
    Set oRsBuscaAtencPac = AtencionesSeleccionarPorIdPaciente(lnIdPaciente, 3)
    If oRsBuscaAtencPac.RecordCount > 0 Then
        oRsBuscaAtencPac.MoveLast
        If oRsBuscaAtencPac.Fields!IdEstado = 10 Then
           MsgBox "El paciente tiene ALTA MEDICA en Hospitalización, pero no tiene ALTA ADMINISTRATIVA (PENDIENTE PAGO SEGUROS o PAGO)" + Chr(13) + "Fecha Ingreso: " & oRsBuscaAtencPac.Fields!FechaIngreso & "  Cuenta: " & Trim(Str(oRsBuscaAtencPac.Fields!idCuentaAtencion)) & " " & oRsBuscaAtencPac.Fields!dplan & Chr(13) & "Debe dar de alta al paciente antes de agregar otra atención", vbInformation, "Mensaje"
        Else
            lcEstadoActual = "." & Trim(Str(oRsBuscaAtencPac.Fields!IdEstado)) & "."
            If InStr(lcEstadosCta, lcEstadoActual) = 0 Then
                MsgBox "El paciente tiene una atención pendiente en Hospitalización" + Chr(13) + "Fecha Ingreso: " & oRsBuscaAtencPac.Fields!FechaIngreso & "  Cuenta: " & Trim(Str(oRsBuscaAtencPac.Fields!idCuentaAtencion)) & " " & oRsBuscaAtencPac.Fields!dplan & Chr(13) & "Debe dar de alta al paciente antes de agregar otra atención", vbInformation, "Mensaje"
                BuscaSiEstaHospitalizado = True
                Exit Function
            End If
        End If
    End If
    'Emergencia -Consultorios
    Set oRsBuscaAtencPac = AtencionesSeleccionarPorIdPaciente(lnIdPaciente, 2)
    If oRsBuscaAtencPac.RecordCount > 0 Then
        oRsBuscaAtencPac.MoveLast
        lcEstadoActual = "." & Trim(Str(oRsBuscaAtencPac.Fields!IdEstado)) & "."
        If oRsBuscaAtencPac.Fields!IdEstado = 10 Then
           'con Alta Medica
           If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oRsBuscaAtencPac.Fields!IdFormaPago, oConexion) = False Then
              lnTotalDeuda = mo_ReglasFacturacion.RetornaConsumoFarmaciaServiciosPorNroCuenta(oRsBuscaAtencPac.Fields!idCuentaAtencion, oConexion)
'              If MsgBox("El paciente tiene ALTA MEDICA en Consultorio-Emergencia" + Chr(13) + "Fecha Ingreso: " & oRsBuscaAtencPac.Fields!FechaIngreso & "  Cuenta: " & Trim(Str(oRsBuscaAtencPac.Fields!IdCuentaAtencion)) & " - " & oRsBuscaAtencPac.Fields!dplan & Chr(13) & Chr(13) & "¿Desea realizar otra Admisión en Emergencia?", vbQuestion + vbYesNo, "Admisión") = vbNo Then
'                 BuscaSiEstaHospitalizado = True
'                 Exit Function
'              End If
           'Actualizado 16102014
              If MsgBox("El paciente tiene ALTA MEDICA en Consultorio-Emergencia" + Chr(13) + "Fecha Ingreso: " & oRsBuscaAtencPac.Fields!FechaIngreso & "  Cuenta: " & Trim(Str(oRsBuscaAtencPac.Fields!idCuentaAtencion)) & " - " & oRsBuscaAtencPac.Fields!dplan & Chr(13) & Chr(13) & "¿Desea realizar otra Admisión?", vbQuestion + vbYesNo, "Admisión") = vbNo Then 'A.Yañez 15/10/2014
                          BuscaSiEstaHospitalizado = True
                          Exit Function
              End If
           Else
              lnTotalDeuda = mo_ReglasFacturacion.RetornaTotalPagosPendientesPorNroCuentaDEBB(oRsBuscaAtencPac.Fields!idCuentaAtencion, oConexion)
              If lnTotalDeuda > 0 Then
                    If MsgBox("El paciente tiene ALTA MEDICA en Consultorio-Emergencia" + Chr(13) + "Fecha Ingreso: " & oRsBuscaAtencPac.Fields!FechaIngreso & "  Cuenta: " & Trim(Str(oRsBuscaAtencPac.Fields!idCuentaAtencion)) & " - " & oRsBuscaAtencPac.Fields!dplan + " - Debe: " & Format(lnTotalDeuda, "####,###,##0.00") & Chr(13) & Chr(13) & "¿Desea realizar otra Admisión?", vbQuestion + vbYesNo, "Admisión") = vbNo Then
                       BuscaSiEstaHospitalizado = True
                       Exit Function
                    End If
              End If
           End If
        Else
            If InStr(lcEstadosCta, lcEstadoActual) = 0 Then
                MsgBox "El paciente tiene una atención pendiente en Emergencia-Consultorios" + Chr(13) + "Fecha Ingreso: " & "  Cuenta: " & Trim(Str(oRsBuscaAtencPac.Fields!idCuentaAtencion)) & " " & oRsBuscaAtencPac.Fields!dplan & Chr(13) & "Debe dar de alta al paciente antes de agregar otra atención", vbInformation, "Mensaje"
                BuscaSiEstaHospitalizado = True
                Exit Function
            End If
        End If
    End If
    'CE
    If lnTipoServicioLlamante <> sghTipoServicio.sghConsultaExterna Then
        Set oRsBuscaAtencPac = AtencionesSeleccionarPorIdPaciente(lnIdPaciente, 1)
        If oRsBuscaAtencPac.RecordCount > 0 And lcBuscaParametro.SeleccionaFilaParametro(519) = "S" Then
            oRsBuscaAtencPac.MoveLast
            lcEstadoActual = "." & Trim(Str(oRsBuscaAtencPac.Fields!IdEstado)) & "."
            If InStr(lcEstadosCta, lcEstadoActual) = 0 Then
                If MsgBox("El paciente tiene una atención pendiente en Consultorios Externo" + Chr(13) + "Fecha Ingreso: " & oRsBuscaAtencPac.Fields!FechaIngreso & "  Cuenta: " & Trim(Str(oRsBuscaAtencPac.Fields!idCuentaAtencion)) & " " & oRsBuscaAtencPac.Fields!dplan & Chr(13) & Chr(13) & "¿Desea Cerrar la Cuenta?", vbQuestion + vbYesNo, "Admisión") = vbYes Then
                    If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oRsBuscaAtencPac.Fields!IdFormaPago, oConexion) = False Then
                       CuentaAtencionPendientePagoSeguros oRsBuscaAtencPac.Fields!idCuentaAtencion, oRsBuscaAtencPac.Fields!idPaciente, True
                    Else
                       CuentaAtencionCerradoAutomatico oRsBuscaAtencPac.Fields!idCuentaAtencion, oRsBuscaAtencPac.Fields!idPaciente, True
                    End If
                Else
                    BuscaSiEstaHospitalizado = True
                End If
            End If
        End If
    End If
    Set oRsBuscaAtencPac = Nothing
End Function



Function AtencionesEstanciaHospitalariaSeleccionarPorIdAtencion(lnIdAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesEstanciaHospitalariaSeleccionarPorIdAtencion = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesEstanciaHospitalariaSeleccionarPorIdAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesEstanciaHospitalariaSeleccionarPorIdAtencion = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function DevuelveEspecialidadesServicioSegunUsuarioSistema(lnIdAreaLaboraUsuarioSistema As sghAreasLaboraEmpleado, lnIdUsuario As Long) As String
    Dim rsIdAlmacen As Recordset
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Dim lcEspecialidadesDelUsuario  As String
    Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(lnIdAreaLaboraUsuarioSistema, lnIdUsuario)
    lcEspecialidadesDelUsuario = ""
    If rsIdAlmacen.RecordCount > 0 Then
       lcEspecialidadesDelUsuario = " and ("
       rsIdAlmacen.MoveFirst
       Do While Not rsIdAlmacen.EOF
          lcEspecialidadesDelUsuario = lcEspecialidadesDelUsuario & " dbo.Servicios.idEspecialidad=" & Trim(Str(rsIdAlmacen.Fields!idLaboraSubArea)) & " or "
          rsIdAlmacen.MoveNext
       Loop
       DevuelveEspecialidadesServicioSegunUsuarioSistema = Left(lcEspecialidadesDelUsuario, Len(lcEspecialidadesDelUsuario) - 4) & ")"
    Else
       DevuelveEspecialidadesServicioSegunUsuarioSistema = lcEspecialidadesDelUsuario
    End If
    Set rsIdAlmacen = Nothing
    Set oBuscaDondeLabora = Nothing
End Function



Function BuscaAtencionesCEparaFormatoHIS(idServicioIngreso As Long, IdMedico As Long, lcFechaIngreso As String, _
                                         lcFechaAlta As String, lcFiltro As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesCEparaFormatoHIS"
        Set oParameter = .CreateParameter("@idServicioIngreso", adInteger, adParamInput, 0, idServicioIngreso): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcFechaIngreso", adDBTimeStamp, adParamInput, 0, lcFechaIngreso): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcFechaAlta", adDBTimeStamp, adParamInput, 0, lcFechaAlta): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 500, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
        If lcFiltro <> "" Then
           oRecordset.Filter = lcFiltro
        End If
   End With
   Set BuscaAtencionesCEparaFormatoHIS = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function BuscaAtencionesDxCEparaFormatoHIS(idAtencion As Long) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDiagnosticosCEparaFormatoHIS"
        Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set BuscaAtencionesDxCEparaFormatoHIS = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function




Sub CuentaAtencionCerradoAutomatico(lnIdCuentaAtencion As Long, lnIdPaciente As Long, EsConsultaExterna As Boolean)
    Dim oConexion As New ADODB.Connection
    Dim oRsTmp As New ADODB.Recordset
    Dim sSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lcHora As String
    On Error GoTo ErrCierre
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    '

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CamasLimpiaIdPaciente"
        Set oParameter = .CreateParameter("@lIdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        .Execute
    End With
    '
    lcHora = Format(Time, "hh:mm")
    Set oCommand = Nothing
    Set oParameter = Nothing
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = oConexion
       .CommandText = "FacturacionCuentasAtencionCerradoAutomatico"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 738): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcHoraCierre", adChar, adParamInput, 5, lcHora): .Parameters.Append oParameter
       .Execute
    End With
    '
    oConexion.CommitTrans
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRsTmp = Nothing
    Exit Sub
ErrCierre:
    oConexion.RollbackTrans
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Sub

Sub CuentaAtencionPendientePagoSeguros(lnIdCuentaAtencion As Long, lnIdPaciente As Long, EsConsultaExterna As Boolean)
    Dim oConexion As New ADODB.Connection
    Dim oRsTmp As New ADODB.Recordset
    Dim sSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lcHora As String
    Dim lnTotalDeuda As Double
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    On Error GoTo ErrCierre
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    '
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CamasLimpiaIdPaciente"
        Set oParameter = .CreateParameter("@lIdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        .Execute
    End With
    '
    lnTotalDeuda = mo_ReglasFacturacion.RetornaConsumoFarmaciaServiciosPorNroCuenta(lnIdCuentaAtencion, oConexion)
    lcHora = Format(Time, "hh:mm")
    Set oCommand = Nothing
    Set oParameter = Nothing
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = oConexion
       .CommandText = "FacturacionCuentasAtencionPendientePagoSeguro"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 738): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcHoraCierre", adChar, adParamInput, 5, lcHora): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnDeudaPendiente", adCurrency, adParamInput, 0, lnTotalDeuda): .Parameters.Append oParameter
       .Execute
    End With
    '
    oConexion.CommitTrans
    oConexion.Close
    Set oCommand = Nothing
    Set oConexion = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set oRsTmp = Nothing
    Exit Sub
ErrCierre:
    oConexion.RollbackTrans
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Sub

Function CitasSeleccionarPorNroCuenta(lnIdCuentaAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CitasSeleccionarPorNroCuenta"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CitasSeleccionarPorNroCuenta = oRecordset
    Set oConexion = Nothing
    Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function AtencionesSeleccionarHospEmergPorFechasDeEgresoMedico(ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, lcOrdenFiltro As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarHospEmergPorFechasDeEgresoMedico"
        Set oParameter = .CreateParameter("@lcOrdenFiltro", adVarChar, adParamInput, 13, lcOrdenFiltro): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ldFechaEgresoIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ldFechaEgresoFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarHospEmergPorFechasDeEgresoMedico = oRecordset
   oConexion.Close
   Set oCommand = Nothing
   Set oConexion = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function DevuelveServiciosQueSonPuntosCarga(lcFiltro As String, lnIdTipoEstado As sghTipoEstados, lnOrden As sghOrdenDeServiciosHospital) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lcSql As String, sWhere As String
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion

    lcSql = ""
    sWhere = ""
    If lnIdTipoEstado <> sghFiltraAnuladosYactivos Then
       lcSql = lcSql & " and  dbo.Servicios.idEstado=" & lnIdTipoEstado
    End If
    If lnOrden = sghPorDescTipoServicio Then
       lcSql = lcSql & " ORDER BY dbo.TiposServicio.Descripcion,dbo.Servicios.Nombre"
    Else
       lcSql = lcSql & " ORDER BY dbo.Servicios.Nombre"
    End If
    lcSql = lcFiltro & lcSql
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveServiciosQueSonPuntosCarga"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set DevuelveServiciosQueSonPuntosCarga = oRecordset
   oConexion.Close
   Set oCommand = Nothing
   Set oConexion = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AdmisionPacienteExternoConSeguroAgregar(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, _
                                                 oDoPaciente As DOPaciente, lnIdTablaLISTBARITEMS As Long, _
                                                 lcNombrePc As String, mo_DoAtencionDatosAdicionales As DoAtencionDatosAdicionales) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim oHistoria As New HistoriasClinicas
Dim oDoHistoria As New DOHistoriaClinica
Dim oAtencionesDatosAdicionales As New AtencionesDatosAdicionales
Dim bProcesoOK As Boolean
Dim bPacienteNuevo As Boolean
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oAtencionesDatosAdicionales.Conexion = oConexion
    '
    oConexion.BeginTrans
    bProcesoOK = True
    '
    bPacienteNuevo = False
    If oDoPaciente.idPaciente = 0 Then
        'Paciente nuevo
        bPacienteNuevo = True
        'crea la historia clinica, si es difrenete a manual
        If oDoPaciente.IdTipoNumeracion <> sghHistoriaDefinitivaManual Then
           oDoPaciente.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoPaciente.IdTipoNumeracion)
        End If
        'Agregar paciente
        If Not PacientesAgregar(oDoPaciente) Then
            bProcesoOK = False: ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
        End If
        oDoHistoria.IdTipoNumeracion = oDoPaciente.IdTipoNumeracion
        'Agregar historia clinica (Si es historia definitiva)
        Select Case oDoHistoria.IdTipoNumeracion
        Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
            oDoHistoria.FechaCreacion = lcBuscaParametro.RetornaFechaServidorSQL
            oDoHistoria.idEstadoHistoria = 1
            oDoHistoria.idPaciente = oDoPaciente.idPaciente
            oDoHistoria.IdTipoHistoria = 1
            oDoHistoria.NroHistoriaClinica = oDoPaciente.NroHistoriaClinica
            oDoHistoria.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
            If Not oHistoria.Insertar(oDoHistoria) Then
                bProcesoOK = False: ms_MensajeError = oHistoria.MensajeError: GoTo Terminar
            End If
        End Select
        
    Else
        '********Paciente antiguo
'        If Not PacientesModificar(oDoPaciente, True) Then
'            bProcesoOK = False: ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
'        End If
    End If
    '
    'Agregar cuenta atención
    oDOCuentaAtencion.FechaCreacion = CDate(oDoAtencion.FechaIngreso & " " & oDoAtencion.HoraIngreso)
    oDOCuentaAtencion.idPaciente = oDoPaciente.idPaciente
    If Not oCuentaAtencion.Insertar(oDOCuentaAtencion) Then
        bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
    End If
    '
    'Agregar atencion
    oDoAtencion.idPaciente = oDoPaciente.idPaciente
    oDoAtencion.idCuentaAtencion = oDOCuentaAtencion.idCuentaAtencion
    If Not oAtencion.Insertar(oDoAtencion) Then
        bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
    End If
    '
    mo_DoAtencionDatosAdicionales.idAtencion = oDoAtencion.idAtencion
    If Not oAtencionesDatosAdicionales.Insertar(mo_DoAtencionDatosAdicionales) Then
        bProcesoOK = False: ms_MensajeError = oAtencionesDatosAdicionales.MensajeError: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "A", oDoAtencion.idCuentaAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, "N° Cuenta: " & Trim(Str(oDoAtencion.idCuentaAtencion)) & " Pac.Ext.Con.Seguro")

Terminar:
    If bProcesoOK Then
        AdmisionPacienteExternoConSeguroAgregar = True
        oConexion.CommitTrans
    Else
        AdmisionPacienteExternoConSeguroAgregar = False
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPaciente = Nothing
    Set oCuentaAtencion = Nothing
    Set oAtencion = Nothing
    Set oHistoria = Nothing
    Set oDoHistoria = Nothing
    Set oAtencionesDatosAdicionales = Nothing
End Function


Function AdmisionPacienteExternoConSeguroModificar(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, _
                                                   oDoPaciente As DOPaciente, lnIdTablaLISTBARITEMS As Long, _
                                                   lcNombrePc As String, mo_DoAtencionDatosAdicionales As DoAtencionDatosAdicionales) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim oHistoria As New HistoriasClinicas
Dim oAtencionesDatosAdicionales As New AtencionesDatosAdicionales
Dim bProcesoOK As Boolean
Dim bPacienteNuevo As Boolean
Dim lnIdServicio As Long
Dim lnIdtipoFinanciamiento As Long
Dim lnPrecioUnitario As Double
Dim oRsTmp As New Recordset
Dim oRsTmp1 As New Recordset

    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oAtencionesDatosAdicionales.Conexion = oConexion
    '
    oConexion.BeginTrans
    bProcesoOK = True
    '
    bPacienteNuevo = False
    '********Paciente antiguo
'    oDoPaciente.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
'    If Not PacientesModificar(oDoPaciente, True) Then
'        bProcesoOK = False: ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
'    End If
    '
    'atencion
    oDoAtencion.idPaciente = oDoPaciente.idPaciente
    If Not oAtencion.modificar(oDoAtencion) Then
        bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "M", oDoAtencion.idCuentaAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, "N° Cuenta: " & Trim(Str(oDoAtencion.idCuentaAtencion)) & " Pac.Ext.Con.Seguro")
    '
    'cuenta atención
    oDOCuentaAtencion.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
    If Not oCuentaAtencion.modificar(oDOCuentaAtencion) Then
        bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
    End If
    'datos adicionales
    mo_DoAtencionDatosAdicionales.idAtencion = oDoAtencion.idAtencion
    If Not oAtencionesDatosAdicionales.modificar(mo_DoAtencionDatosAdicionales) Then
        bProcesoOK = False: ms_MensajeError = oAtencionesDatosAdicionales.MensajeError: GoTo Terminar
    End If
    '
Terminar:
    If bProcesoOK Then
        AdmisionPacienteExternoConSeguroModificar = True
        oConexion.CommitTrans
    Else
        AdmisionPacienteExternoConSeguroModificar = False
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPaciente = Nothing
    Set oCuentaAtencion = Nothing
    Set oAtencion = Nothing
    Set oHistoria = Nothing
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    Set oAtencionesDatosAdicionales = Nothing
End Function

Function AdmisionPacienteExternoConSeguroAnular(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, _
                                                oDoPaciente As DOPaciente, lnIdTablaLISTBARITEMS As Long, _
                                                lcNombrePc As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim oHistoria As New HistoriasClinicas
Dim bProcesoOK As Boolean
Dim bPacienteNuevo As Boolean

    oConexion.Open sighentidades.CadenaConexion
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    '
    oConexion.BeginTrans
    bProcesoOK = True
    '
    bPacienteNuevo = False
    '
    'atencion
    If Not oAtencion.modificar(oDoAtencion) Then
        bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "E", oDoAtencion.idCuentaAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, "N° Cuenta: " & Trim(Str(oDoAtencion.idCuentaAtencion)) & " Pac.Ext.Con.Seguro")
    '
    'cuenta atención
    oDOCuentaAtencion.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
    If Not oCuentaAtencion.modificar(oDOCuentaAtencion) Then
        bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
    End If
Terminar:
    If bProcesoOK Then
        AdmisionPacienteExternoConSeguroAnular = True
        oConexion.CommitTrans
    Else
        AdmisionPacienteExternoConSeguroAnular = False
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPaciente = Nothing
    Set oCuentaAtencion = Nothing
    Set oAtencion = Nothing
    Set oHistoria = Nothing
End Function



Function VerificaSiTieneMovimientoFarmaciaOservicio(lnIdCuenta As Long, lnIdTipoServicio As Long, oConexion As Connection) As String
    Dim oRsTmp As New Recordset
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_Facturacion As New SIGHNegocios.ReglasFacturacion
    Dim oConexion1 As New Connection
    oConexion1.Open sighentidades.CadenaConexion
    oConexion1.CursorLocation = adUseClient
    VerificaSiTieneMovimientoFarmaciaOservicio = ""
    Set oRsTmp = mo_ReglasFarmacia.FarmMovimientoVentasDetalleSeleccionarPorCuenta(lnIdCuenta, oConexion1)
    If oRsTmp.RecordCount > 0 Then
       VerificaSiTieneMovimientoFarmaciaOservicio = "Existen despachos registrados en FARMACIA con esa Cuenta"
    Else
       Set oRsTmp = mo_Facturacion.FactOrdenServicioFiltraPorIdCuenta(lnIdCuenta)
       Select Case lnIdTipoServicio
       Case 1
          oRsTmp.Filter = "idPuntocarga<>6 and idPuntocarga<>1 and IdEstadoFacturacion <>9"   'CE
       Case 3
          oRsTmp.Filter = "idPuntoCarga<>9 and IdEstadoFacturacion = 1"   'HOSP
       Case 2
          oRsTmp.Filter = "idPuntoCarga<>10 and IdEstadoFacturacion = 1"  'EMERG
       Case Else
          oRsTmp.Filter = "idEstadoFacturacion=1 or idEstadoFacturacion=4"
       End Select
       If oRsTmp.RecordCount > 0 Then
          oRsTmp.MoveFirst
          Do While Not oRsTmp.EOF
             If oRsTmp.Fields!IdEstadoFacturacion = 1 Or oRsTmp.Fields!IdEstadoFacturacion = 4 Then
                VerificaSiTieneMovimientoFarmaciaOservicio = "Existen atenciones registradas en OTROS SERVICIOS DEL HOSPITAL con esa Cuenta"
                Exit Do
             End If
             oRsTmp.MoveNext
          Loop
       Else
            'NO hay movimientos
       End If
    End If
    Set mo_ReglasFarmacia = Nothing
    Set oRsTmp = Nothing
    Set mo_Facturacion = Nothing
    Set oConexion1 = Nothing
End Function

Function ExisteNroHistoria(lcNroHistoriaClinica As String) As String
  Dim oRecordset As New ADODB.Recordset
  On Error GoTo ManejadorDeError
  Set oRecordset = PacientesSeleccionarPorNroHistoria(Val(lcNroHistoriaClinica))
  If oRecordset.RecordCount > 0 Then
     ExisteNroHistoria = Trim(oRecordset!ApellidoPaterno) + " " + Trim(oRecordset!ApellidoMaterno) + ", " + oRecordset!PrimerNombre
  Else
     ExisteNroHistoria = ""
  End If
  oRecordset.Close
  Set oRecordset = Nothing
  Exit Function
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function


Function AtencionesSeleccionarPacExtPorCuentaHistoriaApellidosServPARTIC(lnHistoriaClinica As Long, lnIdCuentaAtencion As Long, lcApellidoPaterno As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Atenciones
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set AtencionesSeleccionarPacExtPorCuentaHistoriaApellidosServPARTIC = oTabla.AtencionesSeleccionarPacExtPorCuentaHistoriaApellidosServPARTIC(lnHistoriaClinica, lnIdCuentaAtencion, lcApellidoPaterno)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function AtencionesSeleccionarPacExtPorCuentaHistoriaApellidosServSEGUROS(lnHistoriaClinica As Long, lnIdCuentaAtencion As Long, lcApellidoPaterno As String, lcApellidoVacio As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Atenciones
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set AtencionesSeleccionarPacExtPorCuentaHistoriaApellidosServSEGUROS = oTabla.AtencionesSeleccionarPacExtPorCuentaHistoriaApellidosServSEGUROS(lnHistoriaClinica, lnIdCuentaAtencion, lcApellidoPaterno, lcApellidoVacio)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function


Function AtencionesSeleccionarPacExtPorFechas1(ldFechaIni As Date, ldFechaFin As Date) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Atenciones
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set AtencionesSeleccionarPacExtPorFechas1 = oTabla.AtencionesSeleccionarPacExtPorFechas1(ldFechaIni, ldFechaFin)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function


Function AdmisionPacienteExternoParticularAgregar(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, oDoPaciente As DOPaciente, oRsItems As Recordset, lnIdPuntoCarga As Long, lnIdTablaLISTBARITEMS As Long, lcNombrePc As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim oHistoria As New HistoriasClinicas
Dim oDoHistoria As New DOHistoriaClinica
Dim oDoFactOrdenServ As New DoFactOrdenServ
Dim oFactOrdenesServicio As New FactOrdenServicio
Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
Dim oFactServicioPagos As New FactOrdenServicioPagos
Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
Dim oFacturacionServicioPagos As New FacturacionServicioPagos
Dim oDoFactPreventa As New DoFactPreventa
Dim oFactPreventa As New FactPreventa
Dim bProcesoOK As Boolean
Dim bPacienteNuevo As Boolean
Dim lnIdServicio As Long

    oConexion.Open sighentidades.CadenaConexion
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oFactOrdenesServicio.Conexion = oConexion
    Set oFactServicioPagos.Conexion = oConexion
    Set oFacturacionServicioDespacho.Conexion = oConexion
    Set oFacturacionServicioPagos.Conexion = oConexion
    Set oFactPreventa.Conexion = oConexion
    '
    oConexion.BeginTrans
    bProcesoOK = True
    '
    bPacienteNuevo = False
    If oDoPaciente.idPaciente = 0 Then
        'Paciente nuevo
        bPacienteNuevo = True
        'crea la historia clinica, si es difrenete a manual
        If oDoPaciente.IdTipoNumeracion <> sghHistoriaDefinitivaManual Then
           oDoPaciente.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoPaciente.IdTipoNumeracion)
        End If
        'Agregar paciente
        If Not PacientesAgregar(oDoPaciente) Then
            bProcesoOK = False: ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
        End If
        oDoHistoria.IdTipoNumeracion = oDoPaciente.IdTipoNumeracion
        'Agregar historia clinica (Si es historia definitiva)
        Select Case oDoHistoria.IdTipoNumeracion
        Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
            oDoHistoria.FechaCreacion = lcBuscaParametro.RetornaFechaServidorSQL
            oDoHistoria.idEstadoHistoria = 1
            oDoHistoria.idPaciente = oDoPaciente.idPaciente
            oDoHistoria.IdTipoHistoria = 1
            oDoHistoria.NroHistoriaClinica = oDoPaciente.NroHistoriaClinica
            oDoHistoria.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
            If Not oHistoria.Insertar(oDoHistoria) Then
                bProcesoOK = False: ms_MensajeError = oHistoria.MensajeError: GoTo Terminar
            End If
        End Select
        
    Else
        '********Paciente antiguo
        If Not PacientesModificar(oDoPaciente, True) Then
            bProcesoOK = False: ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
        End If
    End If
    '
    'Agregar cuenta atención
    oDOCuentaAtencion.FechaCreacion = CDate(oDoAtencion.FechaIngreso & " " & oDoAtencion.HoraIngreso)
    oDOCuentaAtencion.idPaciente = oDoPaciente.idPaciente
    If Not oCuentaAtencion.Insertar(oDOCuentaAtencion) Then
        bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
    End If
    '
    'Agregar atencion
    oDoAtencion.idPaciente = oDoPaciente.idPaciente
    oDoAtencion.idCuentaAtencion = oDOCuentaAtencion.idCuentaAtencion
    If Not oAtencion.Insertar(oDoAtencion) Then
        bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "A", oDoAtencion.idCuentaAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, "Cuenta N° " & Trim(Str(oDoAtencion.idCuentaAtencion)) & " Pac.Ext.Particular")

    '
    oRsItems.Sort = "idPreventa"
    oRsItems.MoveFirst
    Do While Not oRsItems.EOF
       With oDoFactOrdenServ
            .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL()
            .idCuentaAtencion = oDoAtencion.idCuentaAtencion
            .IdServicioPaciente = oDoAtencion.idServicioIngreso
            .IdEstadoFacturacion = sghConPreVenta
            .idPaciente = oDoAtencion.idPaciente
            .IdPuntoCarga = lnIdPuntoCarga   'Cajero ó CE
            .IdTipoFinanciamiento = oDoAtencion.IdFormaPago
            .idFuenteFinanciamiento = oDoAtencion.idFuenteFinanciamiento
            .idUsuario = oDoAtencion.IdUsuarioAuditoria
            .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
            .fechaDespacho = .FechaCreacion
            .IdUsuarioDespacho = oDoAtencion.IdUsuarioAuditoria
       End With
       If Not oFactOrdenesServicio.Insertar(oDoFactOrdenServ) Then
            bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
       End If
       With oDOFactOrdenServicioPagos
           .FechaCreacion = oDoFactOrdenServ.FechaCreacion
           .IdEstadoFacturacion = sghConPreVenta
           .IdOrden = oDoFactOrdenServ.IdOrden
           .idUsuario = oDoAtencion.IdUsuarioAuditoria
           .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
       End With
       If Not oFactServicioPagos.Insertar(oDOFactOrdenServicioPagos) Then
           bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
       End If
       oDoFactPreventa.idFactPreventa = oRsItems.Fields!idPreVenta
       If Not oFactPreventa.SeleccionarPorId(oDoFactPreventa) Then
           bProcesoOK = False: ms_MensajeError = oFactPreventa.MensajeError: GoTo Terminar
       End If
       oDoFactPreventa.IdOrden = oDoFactOrdenServ.IdOrden
       oDoFactPreventa.idAtencion = oDoAtencion.idAtencion
       oDoFactPreventa.idEstadoPreventa = 2    'Con Documento
       If Not oFactPreventa.modificar(oDoFactPreventa) Then
           bProcesoOK = False: ms_MensajeError = oFactPreventa.MensajeError: GoTo Terminar
       End If
       lnIdServicio = oRsItems.Fields!IdServicio
       Do While Not oRsItems.EOF And lnIdServicio = oRsItems.Fields!IdServicio
            With oDoFacturacionServicioDesp
                .cantidad = oRsItems.Fields!cantidad
                .IdOrden = oDoFactOrdenServ.IdOrden
                .idProducto = oRsItems.Fields!idProducto
                .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                .Precio = oRsItems.Fields!Precio
                .total = oRsItems.Fields!total
            End With
            If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then
                   bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
            End If
            With oDoFacturacionServicioPagos
                .cantidad = oRsItems.Fields!cantidad
                .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
                .idProducto = oRsItems.Fields!idProducto
                .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                .Precio = oRsItems.Fields!Precio
                .total = oRsItems.Fields!total
            End With
            If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
                bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
            End If
            oRsItems.MoveNext
            If oRsItems.EOF Then
               Exit Do
            End If
       Loop
    Loop
Terminar:
    If bProcesoOK Then
        AdmisionPacienteExternoParticularAgregar = True
        oConexion.CommitTrans
    Else
        AdmisionPacienteExternoParticularAgregar = False
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPaciente = Nothing
    Set oCuentaAtencion = Nothing
    Set oAtencion = Nothing
    Set oHistoria = Nothing
    Set oDoHistoria = Nothing
    Set oDoFactOrdenServ = Nothing
    Set oFactOrdenesServicio = Nothing
    Set oDOFactOrdenServicioPagos = Nothing
    Set oFactServicioPagos = Nothing
    Set oDoFacturacionServicioDesp = Nothing
    Set oFacturacionServicioDespacho = Nothing
    Set oDoFacturacionServicioPagos = Nothing
    Set oFacturacionServicioPagos = Nothing
    Set oDoFactPreventa = Nothing
    Set oFactPreventa = Nothing
End Function


Function AdmisionPacienteExternoParticularModificar(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, oDoPaciente As DOPaciente, oRsItems As Recordset, lnIdPuntoCarga As Long, lnIdTablaLISTBARITEMS As Long, lcNombrePc As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim oHistoria As New HistoriasClinicas
Dim oDoFactOrdenServ As New DoFactOrdenServ
Dim oFactOrdenesServicio As New FactOrdenServicio
Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
Dim oFactServicioPagos As New FactOrdenServicioPagos
Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
Dim oFacturacionServicioPagos As New FacturacionServicioPagos
Dim oDoFactPreventa As New DoFactPreventa
Dim oFactPreventa As New FactPreventa
Dim bProcesoOK As Boolean
Dim bPacienteNuevo As Boolean
Dim lnIdServicio As Long
Dim lnIdtipoFinanciamiento As Long
Dim lnPrecioUnitario As Double
Dim oRsTmp As New Recordset


    oConexion.Open sighentidades.CadenaConexion
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oFacturacionServicioPagos.Conexion = oConexion
    Set oFactServicioPagos.Conexion = oConexion
    Set oFacturacionServicioDespacho.Conexion = oConexion
    Set oFactOrdenesServicio.Conexion = oConexion
    Set oFactPreventa.Conexion = oConexion
    '
    oConexion.BeginTrans
    bProcesoOK = True
    '
    bPacienteNuevo = False
    '********Paciente antiguo
    oDoPaciente.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
    If Not PacientesModificar(oDoPaciente, True) Then
        bProcesoOK = False: ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
    End If
    '
    'atencion
    oDoAtencion.idPaciente = oDoPaciente.idPaciente
    If Not oAtencion.modificar(oDoAtencion) Then
        bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "M", oDoAtencion.idCuentaAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, "Cuenta N° " & Trim(Str(oDoAtencion.idCuentaAtencion)) & " Pac.Ext.Particular")
    '
    'cuenta atención
    oDOCuentaAtencion.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
    If Not oCuentaAtencion.modificar(oDOCuentaAtencion) Then
        bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
    End If
    '*************************Servicios***********************
    'elimina Codigo Servicio (cabecera/detalle) por Estancia de la Cuenta de Atencion}
    Set oRsTmp = mo_ReglasFacturacion.FactOrdenServicioPagosPorIdCuentaAtencion(oDoAtencion.idCuentaAtencion)
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
            If oRsTmp.Fields!IdEstadoFacturacion = sghConPreVenta Then   'Preventa Sin Despacho
                 oDoFacturacionServicioPagos.idOrdenPago = oRsTmp.Fields!idOrdenPago
                 oDoFacturacionServicioPagos.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                 If Not oFacturacionServicioPagos.Eliminar(oDoFacturacionServicioPagos) Then
                    bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
                 End If
                 oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp.Fields!idOrdenPago
                 oDOFactOrdenServicioPagos.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                 If Not oFactServicioPagos.Eliminar(oDOFactOrdenServicioPagos) Then
                    bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
                 End If
                 oDoFacturacionServicioDesp.IdOrden = oRsTmp.Fields!IdOrden
                 oDoFacturacionServicioDesp.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                 If Not oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then
                    bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
                 End If
                 If Not oFactPreventa.SeleccionarPorIdOrden(oRsTmp.Fields!IdOrden, oDoFactPreventa) Then
                    bProcesoOK = False: ms_MensajeError = oFactPreventa.MensajeError: GoTo Terminar
                 End If
                 oDoFactPreventa.idEstadoPreventa = 1   'registrado
                 oDoFactPreventa.IdOrden = 0
                 oDoFactPreventa.idAtencion = 0
                 oDoFactPreventa.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                 If Not oFactPreventa.modificar(oDoFactPreventa) Then
                    bProcesoOK = False: ms_MensajeError = oFactPreventa.MensajeError: GoTo Terminar
                 End If
                 oDoFactOrdenServ.IdOrden = oRsTmp.Fields!IdOrden
                 oDoFactOrdenServ.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                 If Not oFactOrdenesServicio.Eliminar(oDoFactOrdenServ) Then
                    bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
                 End If
                 
            End If
            oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    'Agrega Servicios
    oRsItems.Sort = "idPreventa"
    oRsItems.MoveFirst
    Do While Not oRsItems.EOF
       With oDoFactOrdenServ
            .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL()
            .idCuentaAtencion = oDoAtencion.idCuentaAtencion
            .IdServicioPaciente = oDoAtencion.idServicioIngreso
            .IdEstadoFacturacion = sghConPreVenta
            .idPaciente = oDoAtencion.idPaciente
            .IdPuntoCarga = lnIdPuntoCarga   'Cajero ó CE
            .IdTipoFinanciamiento = oDoAtencion.IdFormaPago
            .idFuenteFinanciamiento = oDoAtencion.idFuenteFinanciamiento
            .idUsuario = oDoAtencion.IdUsuarioAuditoria
            .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
            .fechaDespacho = .FechaCreacion
            .IdUsuarioDespacho = oDoAtencion.IdUsuarioAuditoria
       End With
       If Not oFactOrdenesServicio.Insertar(oDoFactOrdenServ) Then
            bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
       End If
       With oDOFactOrdenServicioPagos
           .FechaCreacion = oDoFactOrdenServ.FechaCreacion
           .IdEstadoFacturacion = sghConPreVenta
           .IdOrden = oDoFactOrdenServ.IdOrden
           .idUsuario = oDoAtencion.IdUsuarioAuditoria
           .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
       End With
       If Not oFactServicioPagos.Insertar(oDOFactOrdenServicioPagos) Then
           bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
       End If
       oDoFactPreventa.idFactPreventa = oRsItems.Fields!idPreVenta
       If Not oFactPreventa.SeleccionarPorId(oDoFactPreventa) Then
           bProcesoOK = False: ms_MensajeError = oFactPreventa.MensajeError: GoTo Terminar
       End If
       oDoFactPreventa.IdOrden = oDoFactOrdenServ.IdOrden
       oDoFactPreventa.idAtencion = oDoAtencion.idAtencion
       oDoFactPreventa.idEstadoPreventa = 2    'Con Documento
       If Not oFactPreventa.modificar(oDoFactPreventa) Then
           bProcesoOK = False: ms_MensajeError = oFactPreventa.MensajeError: GoTo Terminar
       End If
       lnIdServicio = oRsItems.Fields!IdServicio
       Do While Not oRsItems.EOF And lnIdServicio = oRsItems.Fields!IdServicio
            With oDoFacturacionServicioDesp
                .cantidad = oRsItems.Fields!cantidad
                .IdOrden = oDoFactOrdenServ.IdOrden
                .idProducto = oRsItems.Fields!idProducto
                .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                .Precio = oRsItems.Fields!Precio
                .total = oRsItems.Fields!total
            End With
            If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then
                   bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
            End If
            With oDoFacturacionServicioPagos
                .cantidad = oRsItems.Fields!cantidad
                .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
                .idProducto = oRsItems.Fields!idProducto
                .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                .Precio = oRsItems.Fields!Precio
                .total = oRsItems.Fields!total
            End With
            If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
                bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
            End If
            oRsItems.MoveNext
            If oRsItems.EOF Then
               Exit Do
            End If
       Loop
    Loop

Terminar:
    If bProcesoOK Then
        AdmisionPacienteExternoParticularModificar = True
        oConexion.CommitTrans
    Else
        AdmisionPacienteExternoParticularModificar = False
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPaciente = Nothing
    Set oCuentaAtencion = Nothing
    Set oAtencion = Nothing
    Set oHistoria = Nothing
    Set oDoFactOrdenServ = Nothing
    Set oFactOrdenesServicio = Nothing
    Set oDOFactOrdenServicioPagos = Nothing
    Set oFactServicioPagos = Nothing
    Set oDoFacturacionServicioDesp = Nothing
    Set oFacturacionServicioDespacho = Nothing
    Set oDoFacturacionServicioPagos = Nothing
    Set oFacturacionServicioPagos = Nothing
    Set oDoFactPreventa = Nothing
    Set oFactPreventa = Nothing
    Set oRsTmp = Nothing
End Function

Function AdmisionPacienteExternoParticularAnular(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, oDoPaciente As DOPaciente, lnIdTablaLISTBARITEMS As Long, lcNombrePc As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim oHistoria As New HistoriasClinicas
Dim oDoFactOrdenServ As New DoFactOrdenServ
Dim oFactOrdenesServicio As New FactOrdenServicio
Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
Dim oFactServicioPagos As New FactOrdenServicioPagos
Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
Dim oFacturacionServicioPagos As New FacturacionServicioPagos
Dim oDoFactPreventa As New DoFactPreventa
Dim oFactPreventa As New FactPreventa
Dim bProcesoOK As Boolean
Dim bPacienteNuevo As Boolean
Dim oRsTmp As New Recordset

    oConexion.Open sighentidades.CadenaConexion
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oFacturacionServicioPagos.Conexion = oConexion
    Set oFactServicioPagos.Conexion = oConexion
    Set oFacturacionServicioDespacho.Conexion = oConexion
    Set oFactOrdenesServicio.Conexion = oConexion
    Set oFactPreventa.Conexion = oConexion
    '
    oConexion.BeginTrans
    bProcesoOK = True
    '
    bPacienteNuevo = False
    '
    'atencion
    If Not oAtencion.modificar(oDoAtencion) Then
        bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "E", oDoAtencion.idCuentaAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, "Cuenta N° " & Trim(Str(oDoAtencion.idCuentaAtencion)) & " Pac.Ext.Particular")
    '
    'cuenta atención
    oDOCuentaAtencion.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
    If Not oCuentaAtencion.modificar(oDOCuentaAtencion) Then
        bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
    End If
    'elimina Codigo Servicio (cabecera/detalle) por Estancia de la Cuenta de Atencion}
    Set oRsTmp = mo_ReglasFacturacion.FactOrdenServicioPagosPorIdCuentaAtencion(oDoAtencion.idCuentaAtencion)
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
            If oRsTmp.Fields!IdEstadoFacturacion = sghConPreVenta Then   'Preventa Sin Despacho
                 oDoFacturacionServicioPagos.idOrdenPago = oRsTmp.Fields!idOrdenPago
                 oDoFacturacionServicioPagos.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                 If Not oFacturacionServicioPagos.Eliminar(oDoFacturacionServicioPagos) Then
                    bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
                 End If
                 oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp.Fields!idOrdenPago
                 oDOFactOrdenServicioPagos.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                 If Not oFactServicioPagos.Eliminar(oDOFactOrdenServicioPagos) Then
                    bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
                 End If
                 oDoFacturacionServicioDesp.IdOrden = oRsTmp.Fields!IdOrden
                 oDoFacturacionServicioDesp.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                 If Not oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then
                    bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
                 End If
                 If Not oFactPreventa.SeleccionarPorIdOrden(oRsTmp.Fields!IdOrden, oDoFactPreventa) Then
                    bProcesoOK = False: ms_MensajeError = oFactPreventa.MensajeError: GoTo Terminar
                 End If
                 oDoFactPreventa.idEstadoPreventa = 0   'anulado
                 oDoFactPreventa.IdOrden = 0
                 oDoFactPreventa.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                 If Not oFactPreventa.modificar(oDoFactPreventa) Then
                    bProcesoOK = False: ms_MensajeError = oFactPreventa.MensajeError: GoTo Terminar
                 End If
                 oDoFactOrdenServ.IdOrden = oRsTmp.Fields!IdOrden
                 oDoFactOrdenServ.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                 If Not oFactOrdenesServicio.Eliminar(oDoFactOrdenServ) Then
                    bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
                 End If
            End If
            oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close


Terminar:
    If bProcesoOK Then
        AdmisionPacienteExternoParticularAnular = True
        oConexion.CommitTrans
    Else
        AdmisionPacienteExternoParticularAnular = False
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPaciente = Nothing
    Set oCuentaAtencion = Nothing
    Set oAtencion = Nothing
    Set oHistoria = Nothing
    Set oDoFactOrdenServ = Nothing
    Set oFactOrdenesServicio = Nothing
    Set oDOFactOrdenServicioPagos = Nothing
    Set oFactServicioPagos = Nothing
    Set oDoFacturacionServicioDesp = Nothing
    Set oFacturacionServicioDespacho = Nothing
    Set oDoFacturacionServicioPagos = Nothing
    Set oFacturacionServicioPagos = Nothing
    Set oDoFactPreventa = Nothing
    Set oFactPreventa = Nothing
    Set oRsTmp = Nothing
End Function



Sub GeneraEstanciaPorCadaServicioHospitalizado(lnIdCuentaAtencion As Long, ldFechaAltaMedica As Date, _
                                    lcHoraAltaMedica As String, ByRef mrs_FacturacionProductos As ADODB.Recordset, _
                                    ByRef lnTotalPagarEstancia As Double, ByRef lnDiasEstanciaTotal As Long, _
                                    oConexion As Connection, lbActualizaRecordset As Boolean, _
                                    lbBuscarEnWEBfechaAfiliacion As Boolean)
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oRs As New Recordset
    Dim oRsTmpProductos As New Recordset
    Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    Dim oBuscaDiasPaciente As New SIGHDatos.Parametros
    Dim lnNroTransferenciaServicios As Long, lnHorasEstanciaTotal As Long
    Dim lnComoSeTrabajaEnEstadoCuenta As sghComoSeTrabajaEnEstadoCuentaLosSeguros
    Dim lnIdtipoFinanciamiento As Long, lnIdFuenteFinanciamiento As Long
    Dim lnHorasAcumuladas As Long, lnTransferencia As Long
    Dim lnPrecioDespacho As Double, lnCantidad As Double
    Dim lnCantidadSIS As Long, lnPrecioSIS As Double, lnImporteSIS As Double
    Dim lnCantidadSOAT As Long: Dim lnPrecioSOAT As Double: Dim lnImporteSOAT As Double
    Dim lnCantidadCONV As Long, lnPrecioCONV As Double, lnImporteCONV As Double, lcEsCONVENIO As String
    Dim lnImporteEXO As Double, lnCantidadPagar As Long, lnPrecioPagar As Double, lnTotalPagar As Double
    Dim lnPrecioParticular As Double, lnCantidadEstancia As Long, lnPrecioEstancia As Double
    Dim lnidEstadoFacturacion As Long, lnImporteEnBoleta As Double, lbNuevoProducto As Boolean
    Dim lcHoraEstanciaMax As String
    Dim ldSISFechaAfiliacion As Date, lcSIScodigo As String, lnIdSiaSis As Long, lcAfiliacionSIS As String
    Dim wxParametro322 As String, wxParametro301 As String, wxParametro302 As String
    Dim wxParametro323 As String, wxParametroJAMO As String, ml_DNI  As String, lcMensaje111 As String
    Dim ldFechaHoraIngreso As Date, ldFechaHoraAlta As Date, lnEstanciaSIS As Integer
    '
    Set oRsTmp = mo_ReglasFacturacion.AtencionesEstanciaHospitalariaSeleccionarPorIdCuentaAtencion(lnIdCuentaAtencion)
    lnNroTransferenciaServicios = oRsTmp.RecordCount
    If lnNroTransferenciaServicios > 0 Then
       lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
       '
        With oRsTmpProductos
              .Fields.Append "IdProducto", adInteger, 4, adFldIsNullable
              .LockType = adLockOptimistic
              .Open
        End With
       '
       oRsTmp.MoveFirst
       lnDiasEstanciaTotal = oBuscaDiasPaciente.DiasDelPacienteEnHospitalizacionEmergencia(oRsTmp.Fields!FechaIngreso, _
                                     oRsTmp.Fields!HoraIngreso, ldFechaAltaMedica, lcHoraAltaMedica, lcHoraEstanciaMax)
       lnHorasEstanciaTotal = lnDiasEstanciaTotal * 24
       lnComoSeTrabajaEnEstadoCuenta = mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(oRsTmp.Fields!IdTipoFinanciamiento, oConexion)
       lnIdtipoFinanciamiento = oRsTmp.Fields!IdTipoFinanciamiento
       lnIdFuenteFinanciamiento = oRsTmp.Fields!idFuenteFinanciamiento
       lnHorasAcumuladas = 0
       lnTransferencia = 1
       lnTotalPagarEstancia = 0
       'debb-29/09/2017
       If lnIdtipoFinanciamiento = sghTipoFinanciamiento.sghSis And Not IsNull(oRsTmp!NroDocumento) Then
           wxParametro322 = IIf(lbBuscarEnWEBfechaAfiliacion = True, lcBuscaParametro.SeleccionaFilaParametro(322), "N")
           wxParametro301 = lcBuscaParametro.SeleccionaFilaParametro(301)
           wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
           wxParametro323 = lcBuscaParametro.SeleccionaFilaParametro(323)
           wxParametroJAMO = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
           ml_DNI = oRsTmp!NroDocumento
           ldSISFechaAfiliacion = mo_reglasComunes.DevuelveFechaAfiliacionSIS(lcSIScodigo, lnIdSiaSis, _
                                  lcAfiliacionSIS, lnIdtipoFinanciamiento, wxParametro301, wxParametro302, _
                                  wxParametro322, ml_DNI, wxParametro323, wxParametroJAMO, lcMensaje111)
       Else
           ldSISFechaAfiliacion = 0
       End If
       '
       Do While Not oRsTmp.EOF
            '
            lnPrecioParticular = mo_ReglasCaja.FactCatalogoServicioDevuelvePrecioUnitario(oRsTmp.Fields!idProducto, 1)
            lnPrecioParticular = lnPrecioParticular / 24
            lnPrecioDespacho = mo_ReglasCaja.FactCatalogoServicioDevuelvePrecioUnitario(oRsTmp.Fields!idProducto, _
                                                                                 oRsTmp.Fields!IdTipoFinanciamiento)
            lnPrecioDespacho = lnPrecioDespacho / 24
            '
            ldFechaHoraIngreso = Format(oRsTmp.Fields!fechaOcupacion, sighentidades.DevuelveFechaSoloFormato_DMY) & _
                                 " " & oRsTmp.Fields!HoraOcupacion
            If IsNull(oRsTmp.Fields!fechaDesocupacion) Then
               ldFechaHoraAlta = Format(ldFechaAltaMedica, sighentidades.DevuelveFechaSoloFormato_DMY) & _
                                 " " & lcHoraAltaMedica
            Else
               ldFechaHoraAlta = Format(oRsTmp.Fields!fechaDesocupacion, sighentidades.DevuelveFechaSoloFormato_DMY) & _
                                 " " & oRsTmp.Fields!HoraDesocupacion
            End If
            '
            If lnNroTransferenciaServicios = 1 Then
               'Solo estuvo en 1 Servicio
               lnCantidad = lnHorasEstanciaTotal
            ElseIf lnNroTransferenciaServicios = lnTransferencia Then
               'Es el ultimo servicio
               lnCantidad = lnHorasEstanciaTotal - lnHorasAcumuladas
               lnHorasAcumuladas = lnHorasAcumuladas + lnCantidad
               If ldFechaAltaMedica <> ldFechaHoraAlta Then
                    ldFechaHoraAlta = Format(ldFechaAltaMedica, sighentidades.DevuelveFechaSoloFormato_DMY) & _
                                      " " & lcHoraAltaMedica
               End If
            Else
               'primer Servicio, 2do servicio,....penUltimo Servicio
               If IsNull(oRsTmp.Fields!fechaDesocupacion) Then
                  lnCantidad = oBuscaDiasPaciente.HorasDelPacienteEnHospitalizacionEmergencia(oRsTmp.Fields!fechaOcupacion, _
                                                  oRsTmp.Fields!HoraOcupacion, ldFechaAltaMedica, lcHoraAltaMedica)
               Else
                  lnCantidad = oBuscaDiasPaciente.HorasDelPacienteEnHospitalizacionEmergencia(oRsTmp.Fields!fechaOcupacion, _
                                                  oRsTmp.Fields!HoraOcupacion, oRsTmp.Fields!fechaDesocupacion, _
                                                  oRsTmp.Fields!HoraDesocupacion)
               End If
               lnHorasAcumuladas = lnHorasAcumuladas + lnCantidad
            End If
            lnTransferencia = lnTransferencia + 1
            lnidEstadoFacturacion = 1
            lnImporteEnBoleta = 0
            lbNuevoProducto = True
            If oRsTmpProductos.RecordCount > 0 Then
               oRsTmpProductos.MoveFirst
               oRsTmpProductos.Find "idProducto=" & oRsTmp!idProducto
               If oRsTmpProductos.EOF Then
                  lbNuevoProducto = False
               End If
            End If
            If lbNuevoProducto = True Then
                 Set oRs = mo_ReglasFacturacion.FacturacionServicioPagosPorIdCuentaIdProducto(lnIdCuentaAtencion, _
                                                                                                    oRsTmp!idProducto)
                 If oRs.RecordCount > 0 Then
                    oRs.MoveFirst
                    lnPrecioDespacho = oRs.Fields!Precio
                    lnidEstadoFacturacion = 4
                    Do While Not oRs.EOF
                        lnImporteEnBoleta = Round(oRs.Fields!cantidad * oRs.Fields!Precio, 2)
                        oRs.MoveNext
                    Loop
                 End If
                 oRs.Close
                 oRsTmpProductos.AddNew
                 oRsTmpProductos.Fields!idProducto = oRsTmp!idProducto
                 oRsTmpProductos.Update
            End If
            '
            Select Case lnComoSeTrabajaEnEstadoCuenta
            Case sghTrabajaSeguroSIS
                 lnCantidadSIS = lnCantidad
                 lnPrecioSIS = lnPrecioDespacho
                 lnImporteSIS = Round(lnCantidad * lnPrecioDespacho, 2)
                 lnTotalPagarEstancia = lnTotalPagarEstancia + lnImporteSIS
                 lnCantidadEstancia = lnCantidadSIS
                 lnPrecioEstancia = lnPrecioSIS
                 '
                 If ldSISFechaAfiliacion <> 0 Then
                    If ldSISFechaAfiliacion >= ldFechaHoraIngreso And ldSISFechaAfiliacion <= ldFechaHoraAlta Then
                        lnEstanciaSIS = oBuscaDiasPaciente.HorasDelPacienteEnHospitalizacionEmergencia(CDate(Format(ldSISFechaAfiliacion, sighentidades.DevuelveFechaSoloFormato_DMY)), _
                                                  Format(ldSISFechaAfiliacion, "hh:mm"), CDate(Format(ldFechaHoraAlta, sighentidades.DevuelveFechaSoloFormato_DMY)), _
                                                  Format(ldFechaHoraAlta, "hh:mm"))
                        '
                        lnCantidadSIS = lnEstanciaSIS
                        lnPrecioSIS = lnPrecioDespacho
                        lnImporteSIS = Round(lnEstanciaSIS * lnPrecioDespacho, 2)
                        lnTotalPagarEstancia = lnTotalPagarEstancia + lnImporteSIS
                        'lnCantidadEstancia = lnCantidadSIS
                        'lnPrecioEstancia = lnPrecioSIS
                        '
                        lnImporteEXO = 0
                        lnCantidadPagar = lnCantidad - lnCantidadSIS
                        lnTotalPagar = Round((lnCantidad - lnCantidadSIS) * lnPrecioParticular, 2)
                        lnTotalPagarEstancia = lnTotalPagarEstancia + lnTotalPagar
                        'lnCantidadEstancia = lnCantidadPagar
                        'lnPrecioEstancia = lnPrecioDespacho
                    End If
                 End If
            Case sghTrabajaSeguroSOAT
                 lnCantidadSOAT = lnCantidad
                 lnPrecioSOAT = lnPrecioDespacho
                 lnImporteSOAT = Round(lnCantidad * lnPrecioDespacho, 2)
                 lnTotalPagarEstancia = lnTotalPagarEstancia + lnImporteSOAT
                 lnCantidadEstancia = lnCantidadSOAT
                 lnPrecioEstancia = lnPrecioSOAT
            Case sghTrabajaSeguroConvenios
                 lnPrecioCONV = lnPrecioDespacho
                 lcEsCONVENIO = "Si"
                 lnCantidadCONV = lnCantidad
                 lnImporteCONV = Round(lnCantidad * lnPrecioDespacho, 2)
                 lnTotalPagarEstancia = lnTotalPagarEstancia + lnImporteCONV
                 lnCantidadEstancia = lnCantidadCONV
                 lnPrecioEstancia = lnPrecioCONV
            Case Else           'exoneraciones/particular hospitalizado
                 lnImporteEXO = 0
                 lnCantidadPagar = lnCantidad
                 lnTotalPagar = Round(lnCantidad * lnPrecioDespacho, 2)
                 lnTotalPagarEstancia = lnTotalPagarEstancia + lnTotalPagar
                 lnCantidadEstancia = lnCantidadPagar
                 lnPrecioEstancia = lnPrecioDespacho
            End Select
             '
             If lbActualizaRecordset = True Then
                    mrs_FacturacionProductos.AddNew
                    mrs_FacturacionProductos!idProducto = oRsTmp!idProducto
                    mrs_FacturacionProductos!Codigo = Trim(Str(oRsTmp!IdServicio))
                    mrs_FacturacionProductos!nombreProducto = Left(Trim(oRsTmp!nombre) & "  (Estancia: " & Trim(Str(Round(lnCantidad / 24, 2))) & " días)", 254)
                    mrs_FacturacionProductos!CantidadPagar = lnCantidad  'cantidad inicial (no varia)
                    mrs_FacturacionProductos!PrecioUnitario = lnPrecioParticular    'rs!precio  'precio de venta
                    mrs_FacturacionProductos!TotalPagar = Round(lnCantidad * lnPrecioParticular, 2)    'rs!Total
                    mrs_FacturacionProductos!CantidadSIS = lnCantidadSIS
                    mrs_FacturacionProductos!PrecioSIS = lnPrecioSIS
                    mrs_FacturacionProductos!ImporteSIS = lnImporteSIS
                    mrs_FacturacionProductos!CantidadSOAT = lnCantidadSOAT
                    mrs_FacturacionProductos!PrecioSOAT = lnPrecioSOAT
                    mrs_FacturacionProductos!ImporteSOAT = lnImporteSOAT
                    mrs_FacturacionProductos!importeEXO = lnImporteEXO
                    mrs_FacturacionProductos!IdPuntoCarga = sghPtoCargaAdmisionHospitalizacion
                    mrs_FacturacionProductos!IdEstadoFacturacion = lnidEstadoFacturacion
                    mrs_FacturacionProductos!cantidad = lnCantidadPagar 'cantidad a pagar en caja (varia)
                    mrs_FacturacionProductos!totalPorPagar = lnTotalPagar  '(a pagar en caja)
                    'mrs_FacturacionProductos!IdComprobantePago = lnIdComprobantePago
                    mrs_FacturacionProductos!IdOrden = 0
                    mrs_FacturacionProductos!IdUsuarioAutorizaSeguro = sighentidades.Usuario
                    'If lnIdTipoConceptoFarmacia > 0 Then
                    '   mrs_FacturacionProductos!FechaAutorizaSeguro = ldFechaAutorizaSeguro
                    'End If
                    mrs_FacturacionProductos!IdUsuarioAutorizaDevolucion = sighentidades.Usuario
                    'mrs_FacturacionProductos!FechaAutorizaDevolucion = IIf(lnCantidadDev = 0, 0, Date)
                    'mrs_FacturacionProductos!IdComprobantePagoDevolucion = lnIdComprobDev
                    'mrs_FacturacionProductos!NroComprobante = IIf(lcDocumentoPago = "", "", lcDocumentoPago)  'si ya se PAGO muestra BOLETA sino muestra TICKET
                    mrs_FacturacionProductos!IdTipoFinanciamiento = lnIdtipoFinanciamiento
                    mrs_FacturacionProductos!PrecioConv = lnPrecioCONV
                    mrs_FacturacionProductos!esConvenio = lcEsCONVENIO
                    mrs_FacturacionProductos!FechaOrden = Date
                    mrs_FacturacionProductos!cantidadConv = lnCantidadCONV
                    mrs_FacturacionProductos!ImporteConv = lnImporteCONV
                    'mrs_FacturacionProductos!idTipoConceptoFarmacia = lnIdTipoConceptoFarmacia
                    mrs_FacturacionProductos!idFuenteFinanciamiento = lnIdFuenteFinanciamiento
                    'If Not IsNull(rs!IdServicioPaciente) Then
                    '    mrs_FacturacionProductos!ServicioDeEstancia = IIf(IsNull(rs!IdServicioPaciente), ".", mo_ReglasFacturacion.BuscaServicioActualDelPaciente(rs!IdServicioPaciente))
                    '    mrs_FacturacionProductos!idServicioDeEstancia = IIf(IsNull(rs!IdServicioPaciente), 0, rs!IdServicioPaciente)
                    'End If
                    'If ml_AgruparPor = 3 Or ml_AgruparPor = 5 Then
                    '   mrs_FacturacionProductos!Descripcion = rs!dfinanciamiento
                    'End If
                    mrs_FacturacionProductos!ImporteEnBoleta = lnImporteEnBoleta
                    'mrs_FacturacionProductos!nroDcto = lcNroDcto
                    mrs_FacturacionProductos!ComoSeTrabajaEnEstadoCuenta = lnComoSeTrabajaEnEstadoCuenta
                    'mrs_FacturacionProductos!idOrdenPago = lnIdOrdenPago
                    mrs_FacturacionProductos!fechaDespacho = Date
                    mrs_FacturacionProductos!CantidadEstancia = lnCantidadEstancia
                    mrs_FacturacionProductos!PrecioEstancia = lnPrecioEstancia
                    mrs_FacturacionProductos.Update
             End If
             oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    If oRsTmpProductos.State = 1 Then oRsTmpProductos.Close
    Set oRsTmpProductos = Nothing
    Set oRsTmp = Nothing
    Set mo_ReglasCaja = Nothing
    Set oRs = Nothing
    Set mo_reglasComunes = Nothing
End Sub


Function PacientesFiltrarTodosSoloHistoriasDefinitivas(ByVal oPaciente As DOPaciente, _
                                                       lcSinApellido As String, _
                                                       oConexion As Connection) As ADODB.Recordset
    Dim oTabla As New Pacientes
    Set oTabla.Conexion = oConexion
    Set PacientesFiltrarTodosSoloHistoriasDefinitivas = oTabla.FiltrarTodosSoloHistoriasDefinitivas(oPaciente, lcSinApellido)
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function


Function DevuelveDatosDeLaMadreDelPacienteActual(ByRef lnIdNacimientoSeleccionado1 As Long, lnTipoSexoPacienteActual As Long, oConexion As Connection) As String
       DevuelveDatosDeLaMadreDelPacienteActual = ""
       Dim oRsTmp As New Recordset
       Set oRsTmp = AtencionesNacimientosSeleccionarPorFiltro("idNacimiento=" & lnIdNacimientoSeleccionado1)
       If oRsTmp.RecordCount > 0 Then
          If oRsTmp.Fields!idTipoSexo <> lnTipoSexoPacienteActual Then
             MsgBox "El Sexo del Recién Nacido Elegido.....no corresponde con el Sexo del Paciente Ingresado", vbInformation, ""
             lnIdNacimientoSeleccionado1 = 0
             Exit Function
          End If
       Else
          lnIdNacimientoSeleccionado1 = 0
       End If
       oRsTmp.Close
       '
         Dim oCommand As New ADODB.Command
         Dim oParameter As ADODB.Parameter
         ms_MensajeError = ""
         With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "DevuelveDatosDeLaMadreDelPacienteActual "
             Set oParameter = .CreateParameter("@lnIdNacimientoSeleccionado1", adInteger, adParamInput, 0, lnIdNacimientoSeleccionado1): .Parameters.Append oParameter
             Set oRsTmp = .Execute
             Set oRsTmp.ActiveConnection = Nothing
        End With
       '
       ms_MensajeError = ""
       If oRsTmp.RecordCount > 0 Then
          DevuelveDatosDeLaMadreDelPacienteActual = Trim(oRsTmp.Fields!NroHistoriaClinica) & " " & Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & Trim(oRsTmp.Fields!PrimerNombre)
       Else
          lnIdNacimientoSeleccionado1 = 0
       End If
       oRsTmp.Close
       Set oRsTmp = Nothing
End Function


'debb-Jamo
Function AtencionCEAgregar(oDoAtencionTriaje As DOAtencionesCE, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcObservacion As String) As Boolean
10        On Error GoTo ErrCeAgr
          Dim oConexion As New ADODB.Connection
          Dim oConexionSIGH As New ADODB.Connection
          Dim oAtencionesCE As New AtencionesCE
20        AtencionCEAgregar = False
30        oConexionSIGH.Open sighentidades.CadenaConexion
40        oConexionSIGH.BeginTrans
50        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
60        oConexion.BeginTrans
70        Set oAtencionesCE.Conexion = oConexion
80        If oDoAtencionTriaje.TriajeFecha = 0 Then
90           oDoAtencionTriaje.TriajeFecha = lcBuscaParametro.RetornaFechaHoraServidorSQL
100       End If
110       If oAtencionesCE.Insertar(oDoAtencionTriaje) Then
120           Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencionTriaje.IdUsuarioAuditoria, "A", oDoAtencionTriaje.idAtencion, "AtencionesCE", oConexionSIGH, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcObservacion)          'ListBarItems.idListItem
130           AtencionCEAgregar = True
140           oConexion.CommitTrans
150           oConexionSIGH.CommitTrans
160       Else
170           ms_MensajeError = oAtencionesCE.MensajeError
180           oConexion.RollbackTrans
190           oConexionSIGH.RollbackTrans
200       End If
210       oConexion.Close
220       oConexionSIGH.Close
230       Set oConexionSIGH = Nothing
240       Set oConexion = Nothing
250       Set oAtencionesCE = Nothing
260       Exit Function
ErrCeAgr:
270       ms_MensajeError = ms_MensajeError & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                              "Function AtencionCEAgregar", "reglasadmision.cls") 'debb-02/05/2016
End Function

Function AtencionCEModificar(oDoAtencionTriaje As DOAtencionesCE, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcObservacion As String) As Boolean
10        On Error GoTo ErrCeAgr
          Dim oConexion As New ADODB.Connection
          Dim oConexionSIGH As New ADODB.Connection
          Dim oAtencionesCE As New AtencionesCE
20        AtencionCEModificar = False
30        oConexionSIGH.Open sighentidades.CadenaConexion
40        oConexionSIGH.BeginTrans
50        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
60        oConexion.BeginTrans
70        Set oAtencionesCE.Conexion = oConexion
80        If oDoAtencionTriaje.TriajeFecha = 0 Then
90           oDoAtencionTriaje.TriajeFecha = lcBuscaParametro.RetornaFechaHoraServidorSQL
100       End If
110       If oAtencionesCE.modificar(oDoAtencionTriaje) Then
              '
120           Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencionTriaje.IdUsuarioAuditoria, "M", oDoAtencionTriaje.idAtencion, "AtencionesCE", oConexionSIGH, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcObservacion)           'ListBarItems.idListItem
130           AtencionCEModificar = True
140           oConexionSIGH.CommitTrans
150           oConexion.CommitTrans
160       Else
170           ms_MensajeError = oAtencionesCE.MensajeError
180           oConexionSIGH.RollbackTrans
190           oConexion.RollbackTrans
200       End If
210       oConexion.Close
220       oConexionSIGH.Close
230       Set oConexion = Nothing
240       Set oConexionSIGH = Nothing
250       Set oAtencionesCE = Nothing
260       Exit Function
ErrCeAgr:
270       ms_MensajeError = ms_MensajeError & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                             "Function AtencionCEModificar", "reglasadmision.cls") 'debb-02/05/2016
End Function


Function AtencionCESeleccionarPorId(lnIdAtencion As Long) As DOAtencionesCE
Dim oAtencionesCE As New AtencionesCE
Dim oDOAtencionesCE As New DOAtencionesCE
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    oConexion.CursorLocation = adUseClient
    Set AtencionCESeleccionarPorId = Nothing
    Set oAtencionesCE.Conexion = oConexion
    oDOAtencionesCE.idAtencion = lnIdAtencion
    If oAtencionesCE.SeleccionarPorId(oDOAtencionesCE) Then
        Set AtencionCESeleccionarPorId = oDOAtencionesCE
    End If
    ms_MensajeError = oAtencionesCE.MensajeError
    oConexion.Close
    Set oAtencionesCE = Nothing
    Set oConexion = Nothing
End Function



Function AtencionesSeleccionarNroCuentaSegunIdAtencion(lnIdAtencion As Long) As Long
    Dim oRsTmp As New Recordset
    Set oRsTmp = AtencionesSeleccionarPorIdAtencion(lnIdAtencion)
    AtencionesSeleccionarNroCuentaSegunIdAtencion = 0
    If oRsTmp.RecordCount > 0 Then
       AtencionesSeleccionarNroCuentaSegunIdAtencion = oRsTmp.Fields!idCuentaAtencion
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function

Function AtencionCEeliminar(oDoAtencionTriaje As DOAtencionesCE, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcObservacion As String) As Boolean
10        On Error GoTo ErrCeAgr
          Dim oConexion As New ADODB.Connection
          Dim oConexionSIGH As New ADODB.Connection
          Dim oAtencionesCE As New AtencionesCE
20        AtencionCEeliminar = False
30        oConexionSIGH.Open sighentidades.CadenaConexion
40        oConexionSIGH.BeginTrans
50        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
60        oConexion.BeginTrans
70        Set oAtencionesCE.Conexion = oConexion
80        If oAtencionesCE.Eliminar(oDoAtencionTriaje) Then
              '
90            Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencionTriaje.IdUsuarioAuditoria, "E", oDoAtencionTriaje.idAtencion, "AtencionesCE", oConexionSIGH, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcObservacion)            'ListBarItems.idListItem
100           AtencionCEeliminar = True
110           oConexionSIGH.CommitTrans
120           oConexion.CommitTrans
130       Else
140           ms_MensajeError = oAtencionesCE.MensajeError
150           oConexionSIGH.RollbackTrans
160           oConexion.RollbackTrans
170       End If
180       oConexionSIGH.Close
190       oConexion.Close
200       Set oConexion = Nothing
210       Set oConexionSIGH = Nothing
220       Set oAtencionesCE = Nothing
230       Exit Function
ErrCeAgr:
240       ms_MensajeError = ms_MensajeError & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                              "Function AtencionCEeliminar", "reglasadmision.cls") 'debb-02/05/2016

End Function

'debb-jamo
Function AtencionesCEFiltrarPorPaciente(ByVal oTabla As DOPaciente, lnIdCuentaAtencion As Long, _
                              lcFechaTriaje As String, oConexion As Connection, _
                              lbUsaApellidoMaternoVacio As Boolean, _
                              lbUsaApellidoPaternoVacio As Boolean, lcApellidoVacio As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
'Dim oConexion As New ADODB.Connection
Dim sSql As String
Dim sWhere As String
Dim ldFechaInicio As Date, ldFechaFin As Date
Dim oParameter As ADODB.Parameter
     If lcFechaTriaje <> "" Then
        ldFechaInicio = CDate(lcFechaTriaje & " 00:00:01")
        ldFechaFin = CDate(lcFechaTriaje & " 23:23:59")
     End If
     Set AtencionesCEFiltrarPorPaciente = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandType = adCmdStoredProc
       .CommandText = "AtencionesCEFiltrarPorPaciente"
        Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, oTabla.NroHistoriaClinica): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, oTabla.ApellidoPaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, oTabla.ApellidoMaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, oTabla.PrimerNombre): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@dni", adVarChar, adParamInput, 12, oTabla.NroDocumento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcFechaTriaje", adVarChar, adParamInput, 10, lcFechaTriaje): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ldFechaTriaje1", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ldFechaTriaje2", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
        sWhere = ""
        If lbUsaApellidoPaternoVacio = True Then
           sWhere = sWhere & "ApellidoPaterno='" & lcApellidoVacio & "' and "
        End If
        If lbUsaApellidoMaternoVacio = True Then
           sWhere = sWhere & "ApellidoMaterno='" & lcApellidoVacio & "' and "
        End If
        If sWhere <> "" Then
           oRecordset.Filter = Left(sWhere, Len(sWhere) - 5)
        End If
     End With
     Set AtencionesCEFiltrarPorPaciente = oRecordset
     ms_MensajeError = ""
     Set oCommand = Nothing
     Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function PacientesFiltraPorHistoriaClinicaDefinitiva(ByVal oTabla As DOPaciente, Optional oConexion1 As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
 
   If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 150
        oConexion.Open sighentidades.CadenaConexion
   End If
   With oCommand
     .CommandType = adCmdStoredProc
     If oConexion1 Is Nothing Then
        Set .ActiveConnection = oConexion
     Else
        Set .ActiveConnection = oConexion1
     End If
     .CommandText = "PacientesSeleccionarPorHistoriaClinicaDefinitiva"
       Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, oTabla.NroHistoriaClinica): .Parameters.Append oParameter
     Set oRecordset = .Execute
     If oConexion1 Is Nothing Then
        Set oRecordset.ActiveConnection = Nothing
     End If
   End With
     Set PacientesFiltraPorHistoriaClinicaDefinitiva = oRecordset
     ms_MensajeError = ""
     If oConexion1 Is Nothing Then
        oConexion.Close
        Set oRecordset = Nothing
     End If
     Set oCommand = Nothing
     Set oConexion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function


Function ElServicioNecesitaTriaje(lnIdServicio As Long, oConexion As Connection, _
          ByRef lbElConsultorioUsaModuloPerinatal As Boolean, _
          ByRef lbElConsultorioUsaModuloMaterno As Boolean, _
          Optional ByRef lbElMedicoNOregistraDatosCE As Boolean) As Boolean
    Dim oRsTmp As New Recordset
    
    Set oRsTmp = mo_reglasComunes.ServiciosSeleccionarXidentificador(lnIdServicio)
    ElServicioNecesitaTriaje = False
    lbElConsultorioUsaModuloPerinatal = False
    lbElConsultorioUsaModuloMaterno = False
    lbElMedicoNOregistraDatosCE = False
    If oRsTmp.RecordCount > 0 Then
       If oRsTmp.Fields!triaje = True Then
          ElServicioNecesitaTriaje = True
       End If
       If oRsTmp.Fields!UsaModuloNinoSano = True Then
          lbElConsultorioUsaModuloPerinatal = True
       End If
       If oRsTmp.Fields!UsaModuloMaterno = True Then
          lbElConsultorioUsaModuloMaterno = True
       End If
       If oRsTmp.Fields!UsaGalenHos = True Then
          lbElMedicoNOregistraDatosCE = True
       End If
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function

Function EsServicioCostoCero(lnIdServicio As Long) As Boolean
    Dim oRsTmp As New Recordset
    EsServicioCostoCero = False
    Set oRsTmp = mo_reglasComunes.ServiciosSeleccionarXidentificador(lnIdServicio)

    If oRsTmp.RecordCount > 0 Then
       If oRsTmp.Fields!CostoCeroCE = "S" Then
          EsServicioCostoCero = True
       End If
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function

Function AtencionesDiagnosticosSeleccionarTodosPorIdAtencion(lnIdAtencion As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New AtencionesDiagnosticos
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set AtencionesDiagnosticosSeleccionarTodosPorIdAtencion = oTabla.SeleccionarTodosPorIdAtencion(lnIdAtencion)
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing

End Function

Function SunasaPacientesHistoricosSeleccionarPorIdPaciente(idPaciente As Long, oConexion As Connection) As DoSunasaPacientesHistoricos
Dim oSunasaPacientesHistoricos As New SunasaPacientesHistoricos
Dim oDoSunasaPacientesHistoricos As New DoSunasaPacientesHistoricos
    
    Set SunasaPacientesHistoricosSeleccionarPorIdPaciente = Nothing
    
    Set oSunasaPacientesHistoricos.Conexion = oConexion
    oDoSunasaPacientesHistoricos.idPaciente = idPaciente
    If oSunasaPacientesHistoricos.SeleccionarPorIdPaciente(oDoSunasaPacientesHistoricos) Then
        Set SunasaPacientesHistoricosSeleccionarPorIdPaciente = oDoSunasaPacientesHistoricos
    End If
    ms_MensajeError = oSunasaPacientesHistoricos.MensajeError
    Set oSunasaPacientesHistoricos = Nothing
    Set oDoSunasaPacientesHistoricos = Nothing
End Function


Function SunasaPacientesHistoricosSeleccionarPorId(idSunasaPacienteHistorico As Long) As DoSunasaPacientesHistoricos
Dim oConexion As New ADODB.Connection
Dim oSunasaPacientesHistoricos As New SunasaPacientesHistoricos
Dim oDoSunasaPacientesHistoricos As New DoSunasaPacientesHistoricos
    
    Set SunasaPacientesHistoricosSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oSunasaPacientesHistoricos.Conexion = oConexion
    oDoSunasaPacientesHistoricos.idSunasaPacienteHistorico = idSunasaPacienteHistorico
    If oSunasaPacientesHistoricos.SeleccionarPorId(oDoSunasaPacientesHistoricos) Then
        Set SunasaPacientesHistoricosSeleccionarPorId = oDoSunasaPacientesHistoricos
    End If
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oSunasaPacientesHistoricos.MensajeError
    Set oSunasaPacientesHistoricos = Nothing
    Set oDoSunasaPacientesHistoricos = Nothing
End Function

Function AtencionesDatosAdicionalesSeleccionarPorId(IdAtenciones As Long, oConexion As Connection) As DoAtencionDatosAdicionales
Dim oAtencionesDatosAdicionales As New AtencionesDatosAdicionales
Dim oDOAtencionDatosAdicionales As New DoAtencionDatosAdicionales
    
    Set AtencionesDatosAdicionalesSeleccionarPorId = Nothing
    Set oAtencionesDatosAdicionales.Conexion = oConexion
    oDOAtencionDatosAdicionales.idAtencion = IdAtenciones
    If oAtencionesDatosAdicionales.SeleccionarPorId(oDOAtencionDatosAdicionales) Then
        Set AtencionesDatosAdicionalesSeleccionarPorId = oDOAtencionDatosAdicionales
    End If
    ms_MensajeError = oAtencionesDatosAdicionales.MensajeError
    Set oAtencionesDatosAdicionales = Nothing

End Function

Function PerinatalCatalogoCptAutomaticosXidModulo(lnIdModulo As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PerinatalCatalogoCptAutomaticosXidModulo"
        Set oParameter = .CreateParameter("@idModulo", adInteger, adParamInput, 0, lnIdModulo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PerinatalCatalogoCptAutomaticosXidModulo = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function PerinatalCatalogoCptXidModulo(lnIdModulo As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PerinatalCatalogoCptXidModulo"
        Set oParameter = .CreateParameter("@idModulo", adInteger, adParamInput, 0, lnIdModulo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PerinatalCatalogoCptXidModulo = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function PerinatalCatalogoCie10xidModulo(lnIdModulo As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PerinatalCatalogoCie10xidModulo"
        Set oParameter = .CreateParameter("@idModulo", adInteger, adParamInput, 0, lnIdModulo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PerinatalCatalogoCie10xidModulo = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function PerinatalAtencionDiariaXIdAtencion(lnIdAtencion As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PerinatalAtencionDiariaXIdAtencion"
        Set oParameter = .CreateParameter("@lnIdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PerinatalAtencionDiariaXIdAtencion = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function PerinatalAtencionDiariaXidPerinatalAtencion(idPerinatalAtencion As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PerinatalAtencionDiariaXidPerinatalAtencion"
        Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamInput, 0, idPerinatalAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PerinatalAtencionDiariaXidPerinatalAtencion = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function FacturacionPaquetesActualizaXidComprobIdOrdenIdProducto(lnIdProducto As Long, lnIdCuentaAtencion As Long, _
                              lnIdComprobantePago As Long, lnIdOrdenPago As Long, lbNOActualizaConNULL As Boolean, oConexion As Connection) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionPaquetesActualizaXidComprobIdOrdenIdProducto"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idOrdenPago", adInteger, adParamInput, 0, lnIdOrdenPago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lbEsNull", adBoolean, adParamInput, 0, lbNOActualizaConNULL): .Parameters.Append oParameter
        .Execute
   End With
   FacturacionPaquetesActualizaXidComprobIdOrdenIdProducto = True
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function CajaComprobantesPagoActualizaXpacienteYComprobante(lnIdPaciente As Long, lnIdComprobantePago As Long, _
                                             lbActualizaConNULL As Boolean, oConexion As Connection) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantesPagoActualizaXpacienteYComprobante"
        Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lbEsNull", adBoolean, adParamInput, 0, lbActualizaConNULL): .Parameters.Append oParameter
        .Execute
   End With
   CajaComprobantesPagoActualizaXpacienteYComprobante = True
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function PerinatalCEAgregar(oDoPerinatalAtencion As DoPerinatalAtencion, oRsCptInmunizaciones As Recordset, oRsCptFrecuentes As Recordset, oRsDxDesarrollo As Recordset, oRsDxMorbilidadFrec As Recordset, oRsMedicamentos As Recordset, oRsCred As Recordset, oDoPerinatalAtencionCred1 As DoPerinatalAtencionCred1, lnIdAtencion As Long) As Boolean
    On Error GoTo ErrCeAgr
    Dim oConexion As New ADODB.Connection
    Dim oPerinatalAtencion As New PerinatalAtencion
    Dim oDoPerinatalAtencionCpt As New DoPerinatalAtencionCpt, oPerinatalAtencioncpt As New PerinatalAtencionCpt
    Dim oDoPerinatalAtencionDx As New DoPerinatalAtencionDx, oPerinatalAtencionDx As New PerinatalAtencionDx
    Dim oDoPerinatalAtencionMedicamentos As New DOPerinatalAtencionMed, oPerinatalAtencionMedicamentos As New PerinatalAtencionMed
    Dim oDoPerinatalAtencionCred As New DoPerinatalAtencionCred, oPerinatalAtencionCred As New PerinatalAtencionCred
    Dim oDOAtencionDiagnostico As New DOAtencionDiagnostico, oAtencionesDiagnosticos As New AtencionesDiagnosticos
    Dim oPerinatalAtencionCred1 As New PerinatalAtencionCred1
    Dim oPerinatalAtencionDiaria As New PerinatalAtencionDiaria, oDoPerinatalAtencionDiaria As New DoPerinatalAtencionDiaria
    Dim oRsTmpAutomatico As New Recordset, oRsTmpCodigoHis As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    PerinatalCEAgregar = False
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    '
    oConexion.BeginTrans
    '
    Set oPerinatalAtencion.Conexion = oConexion
    Set oPerinatalAtencioncpt.Conexion = oConexion
    Set oPerinatalAtencionDx.Conexion = oConexion
    Set oPerinatalAtencionMedicamentos.Conexion = oConexion
    Set oPerinatalAtencionCred.Conexion = oConexion
    Set oAtencionesDiagnosticos.Conexion = oConexion
    Set oPerinatalAtencionCred1.Conexion = oConexion
    Set oPerinatalAtencionDiaria.Conexion = oConexion
    '
    If Not oPerinatalAtencion.Insertar(oDoPerinatalAtencion) Then
        ms_MensajeError = oPerinatalAtencion.MensajeError: GoTo ErrCeAgr
    End If
    '
    oDoPerinatalAtencionDiaria.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
    oDoPerinatalAtencionDiaria.idAtencion = lnIdAtencion
    oDoPerinatalAtencionDiaria.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
    If Not oPerinatalAtencionDiaria.Insertar(oDoPerinatalAtencionDiaria) Then
        ms_MensajeError = oPerinatalAtencionDiaria.MensajeError: GoTo ErrCeAgr
    End If
    '

    Set oRsTmpAutomatico = PerinatalCatalogoCptAutomaticosXidModulo(oDoPerinatalAtencion.idModulo, oConexion)
    Set oRsTmpCodigoHis = PerinatalCatalogoCptXidModulo(oDoPerinatalAtencion.idModulo, oConexion)
    '
    If oRsCptInmunizaciones.RecordCount > 0 Then
       oRsTmpCodigoHis.Filter = "idLista=" & sghPerinatalListas.sighInmunizaciones
       oDoPerinatalAtencionCpt.idLista = sghPerinatalListas.sighInmunizaciones
       oDoPerinatalAtencionCpt.idModulo = oDoPerinatalAtencion.idModulo
       oDoPerinatalAtencionCpt.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
       oDoPerinatalAtencionCpt.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
       oRsCptInmunizaciones.MoveFirst
       Do While Not oRsCptInmunizaciones.EOF
          oDoPerinatalAtencionCpt.idProducto = oRsCptInmunizaciones.Fields!Id
          oDoPerinatalAtencionCpt.CptEsAutomatico = False
          oDoPerinatalAtencionCpt.idAtencion = oRsCptInmunizaciones.Fields!idAtencion
          'elegidos
          oDoPerinatalAtencionCpt.CodigoHIS = ""
          If oRsTmpCodigoHis.RecordCount > 0 Then
             oRsTmpCodigoHis.MoveFirst
             oRsTmpCodigoHis.Find "idProducto=" & oRsCptInmunizaciones.Fields!Id
             If Not oRsTmpCodigoHis.EOF Then
                oDoPerinatalAtencionCpt.CodigoHIS = oRsTmpCodigoHis.Fields!CodigoHIS
             End If
          End If
          '
          If Not oPerinatalAtencioncpt.Insertar(oDoPerinatalAtencionCpt) Then
             ms_MensajeError = oPerinatalAtencioncpt.MensajeError: GoTo ErrCeAgr
          End If
          'automaticos
          If oRsTmpAutomatico.RecordCount > 0 Then
             oRsTmpAutomatico.MoveFirst
             oRsTmpAutomatico.Find "idProducto=" & oRsCptInmunizaciones.Fields!Id
             If oRsTmpAutomatico.EOF = False Then
                oDoPerinatalAtencionCpt.idProducto = oRsTmpAutomatico.Fields!idProductoAutomatico
                oDoPerinatalAtencionCpt.CptEsAutomatico = True
                oDoPerinatalAtencionCpt.CodigoHIS = ""
                If Not oPerinatalAtencioncpt.Insertar(oDoPerinatalAtencionCpt) Then
                   ms_MensajeError = oPerinatalAtencioncpt.MensajeError: GoTo ErrCeAgr
                End If
             End If
          End If
          '
          oRsCptInmunizaciones.MoveNext
       Loop
    End If
    '
    If oRsCptFrecuentes.RecordCount > 0 Then
       oRsTmpCodigoHis.Filter = "idLista=" & sghPerinatalListas.sighCptFrecuentes
       oDoPerinatalAtencionCpt.idLista = sghPerinatalListas.sighCptFrecuentes
       oDoPerinatalAtencionCpt.idModulo = oDoPerinatalAtencion.idModulo
       oDoPerinatalAtencionCpt.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
       oDoPerinatalAtencionCpt.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
       oRsCptFrecuentes.MoveFirst
       Do While Not oRsCptFrecuentes.EOF
          oDoPerinatalAtencionCpt.idProducto = oRsCptFrecuentes.Fields!Id
          oDoPerinatalAtencionCpt.CptEsAutomatico = False
          oDoPerinatalAtencionCpt.idAtencion = oRsCptFrecuentes.Fields!idAtencion
          '
          oDoPerinatalAtencionCpt.CodigoHIS = ""
          If oRsTmpCodigoHis.RecordCount > 0 Then
             oRsTmpCodigoHis.MoveFirst
             oRsTmpCodigoHis.Find "idProducto=" & oRsCptFrecuentes.Fields!Id
             If Not oRsTmpCodigoHis.EOF Then
                oDoPerinatalAtencionCpt.CodigoHIS = oRsTmpCodigoHis.Fields!CodigoHIS
             End If
          End If
        
            If Not IsNull(oRsCptFrecuentes.Fields!labConfHIS) Then
                oDoPerinatalAtencionCpt.labConfHIS = oRsCptFrecuentes.Fields!labConfHIS
            Else
                oDoPerinatalAtencionCpt.labConfHIS = Space(3)
            End If
          If Not oPerinatalAtencioncpt.Insertar(oDoPerinatalAtencionCpt) Then
             ms_MensajeError = oPerinatalAtencioncpt.MensajeError: GoTo ErrCeAgr
          End If
          oRsCptFrecuentes.MoveNext
       Loop
       
    End If
    '
    oDOAtencionDiagnostico.idAtencion = lnIdAtencion
    oDOAtencionDiagnostico.IdClasificacionDx = 1 'no usuado despues de agregra tipo dx
    oDOAtencionDiagnostico.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
    '
    oRsTmpCodigoHis.Close
    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "PerinatalCatalogoCie10xidModulo"
      Set oParameter = .CreateParameter("@IdModulo", adInteger, adParamInput, 0, oDoPerinatalAtencion.idModulo): .Parameters.Append oParameter
      Set oRsTmpCodigoHis = .Execute
      Set oRsTmpCodigoHis.ActiveConnection = Nothing
    End With
    Set oParameter = Nothing
    Set oCommand = Nothing
    '
    If oRsDxDesarrollo.RecordCount > 0 Then
       oRsTmpCodigoHis.Filter = "idLista=" & sghPerinatalListas.sighMorbilidadDesarrollo
       oDoPerinatalAtencionDx.idLista = sghPerinatalListas.sighMorbilidadDesarrollo
       oDoPerinatalAtencionDx.idModulo = oDoPerinatalAtencion.idModulo
       oDoPerinatalAtencionDx.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
       oDoPerinatalAtencionDx.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
       oRsDxDesarrollo.MoveFirst
       Do While Not oRsDxDesarrollo.EOF
            oDoPerinatalAtencionDx.IdDiagnostico = oRsDxDesarrollo.Fields!Id
            oDoPerinatalAtencionDx.idAtencion = oRsDxDesarrollo.Fields!idAtencion
            oDoPerinatalAtencionDx.labConfHIS = IIf(IsNull(oRsDxDesarrollo.Fields!labConfHIS), "", oRsDxDesarrollo.Fields!labConfHIS)
            'mgaray201411a
            oDoPerinatalAtencionDx.IdClasificacionDx = IIf(IsNull(oRsDxDesarrollo.Fields!IdClasificacionDx), 0, oRsDxDesarrollo.Fields!IdClasificacionDx)
            oDoPerinatalAtencionDx.IdSubclasificacionDx = IIf(IsNull(oRsDxDesarrollo.Fields!IdSubclasificacionDx), 0, oRsDxDesarrollo.Fields!IdSubclasificacionDx)
            oDoPerinatalAtencionDx.CodigoHIS = ""
            If oRsTmpCodigoHis.RecordCount > 0 Then
               oRsTmpCodigoHis.MoveFirst
               oRsTmpCodigoHis.Find "idDiagnostico=" & oRsDxDesarrollo.Fields!Id
               If Not oRsTmpCodigoHis.EOF Then
                  oDoPerinatalAtencionDx.CodigoHIS = oRsTmpCodigoHis.Fields!CodigoHIS
               End If
            End If
            
            If Not oPerinatalAtencionDx.Insertar(oDoPerinatalAtencionDx) Then
                ms_MensajeError = oPerinatalAtencionDx.MensajeError: GoTo ErrCeAgr
            End If
            'mgaray201411a
            If Not AgregarDiagnosticosAtencionDesdePerinatal(oRsDxDesarrollo, oAtencionesDiagnosticos, oDOAtencionDiagnostico, oRsTmpCodigoHis) Then
                ms_MensajeError = oAtencionesDiagnosticos.MensajeError: GoTo ErrCeAgr
            End If
          
            oRsDxDesarrollo.MoveNext
       Loop
    End If
    '
    If oRsDxMorbilidadFrec.RecordCount > 0 Then
       oRsTmpCodigoHis.Filter = "idLista=" & sghPerinatalListas.sighMorbilidadFrecuente
       oDoPerinatalAtencionDx.idLista = sghPerinatalListas.sighMorbilidadFrecuente
       oDoPerinatalAtencionDx.idModulo = oDoPerinatalAtencion.idModulo
       oDoPerinatalAtencionDx.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
       oDoPerinatalAtencionDx.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
       oRsDxMorbilidadFrec.MoveFirst
       Do While Not oRsDxMorbilidadFrec.EOF
            oDoPerinatalAtencionDx.IdDiagnostico = oRsDxMorbilidadFrec.Fields!Id
            oDoPerinatalAtencionDx.idAtencion = oRsDxMorbilidadFrec.Fields!idAtencion
            oDoPerinatalAtencionDx.labConfHIS = IIf(IsNull(oRsDxMorbilidadFrec.Fields!labConfHIS), "", oRsDxMorbilidadFrec.Fields!labConfHIS)
            'mgaray201411a
            oDoPerinatalAtencionDx.IdClasificacionDx = IIf(IsNull(oRsDxMorbilidadFrec.Fields!IdClasificacionDx), 0, oRsDxMorbilidadFrec.Fields!IdClasificacionDx)
            oDoPerinatalAtencionDx.IdSubclasificacionDx = IIf(IsNull(oRsDxMorbilidadFrec.Fields!IdSubclasificacionDx), 0, oRsDxMorbilidadFrec.Fields!IdSubclasificacionDx)
            
            oDoPerinatalAtencionDx.CodigoHIS = ""
            If oRsTmpCodigoHis.RecordCount > 0 Then
               oRsTmpCodigoHis.MoveFirst
               oRsTmpCodigoHis.Find "idDiagnostico=" & oRsDxMorbilidadFrec.Fields!Id
               If Not oRsTmpCodigoHis.EOF Then
                  oDoPerinatalAtencionDx.CodigoHIS = oRsTmpCodigoHis.Fields!CodigoHIS
               End If
            End If
            
            If Not oPerinatalAtencionDx.Insertar(oDoPerinatalAtencionDx) Then
               ms_MensajeError = oPerinatalAtencionDx.MensajeError: GoTo ErrCeAgr
            End If
            'mgaray201410a
            If Not AgregarDiagnosticosAtencionDesdePerinatal(oRsDxMorbilidadFrec, oAtencionesDiagnosticos, oDOAtencionDiagnostico, oRsTmpCodigoHis) Then
                ms_MensajeError = oAtencionesDiagnosticos.MensajeError: GoTo ErrCeAgr
            End If
            
          oRsDxMorbilidadFrec.MoveNext
       Loop
    End If
    '
    oRsMedicamentos.Filter = "seleccionar=true"
    If oRsMedicamentos.RecordCount > 0 Then
       oDoPerinatalAtencionMedicamentos.idModulo = oDoPerinatalAtencion.idModulo
       oDoPerinatalAtencionMedicamentos.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
       oDoPerinatalAtencionMedicamentos.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
       oRsMedicamentos.MoveFirst
       Do While Not oRsMedicamentos.EOF
          oDoPerinatalAtencionMedicamentos.idProducto = oRsMedicamentos.Fields!Id
          oDoPerinatalAtencionMedicamentos.idAtencion = oRsMedicamentos.Fields!idAtencion
          If Not oPerinatalAtencionMedicamentos.Insertar(oDoPerinatalAtencionMedicamentos) Then
             ms_MensajeError = oPerinatalAtencionMedicamentos.MensajeError: GoTo ErrCeAgr
          End If
          oRsMedicamentos.MoveNext
       Loop
    End If
    '
    If oRsCred.RecordCount > 0 Then
       oDoPerinatalAtencionCred.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
       oDoPerinatalAtencionCred.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
       oRsCred.MoveFirst
       Do While Not oRsCred.EOF
          If oRsCred.Fields!idAtencion = lnIdAtencion Then
                oDoPerinatalAtencionCred.credCheck = oRsCred.Fields!credCheck
                oDoPerinatalAtencionCred.credNumero = oRsCred.Fields!credNumero
                oDoPerinatalAtencionCred.edadEnAnios = oRsCred.Fields!edadEnAnios
                oDoPerinatalAtencionCred.idAtencion = oRsCred.Fields!idAtencion
                If Not oPerinatalAtencionCred.Insertar(oDoPerinatalAtencionCred) Then
                   ms_MensajeError = oPerinatalAtencionCred.MensajeError: GoTo ErrCeAgr
                End If
          End If
          oRsCred.MoveNext
       Loop
    End If
    '
    If oDoPerinatalAtencionCred1.idAtencion > 0 Then
       oDoPerinatalAtencionCred1.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
       oDoPerinatalAtencionCred1.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
       If Not oPerinatalAtencionCred1.Insertar(oDoPerinatalAtencionCred1) Then
          ms_MensajeError = oPerinatalAtencionCred1.MensajeError: GoTo ErrCeAgr
       End If
    End If
    
    If grabarDatosPerinatalAtencionIntegral(oDoPerinatalAtencion, lnIdAtencion, oConexion) = False Then
        ms_MensajeError = "No se Puedieron Modificar los datos de atencion Integral"
    End If
    '
    PerinatalCEAgregar = True
    oRsTmpCodigoHis.Close
    oRsTmpAutomatico.Close
    oConexion.Close
ErrCeAgr:
'Resume
    If PerinatalCEAgregar = True Then
       oConexion.CommitTrans
    Else
       oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPerinatalAtencion = Nothing
    Set oDoPerinatalAtencionCpt = Nothing
    Set oPerinatalAtencioncpt = Nothing
    Set oDoPerinatalAtencionDx = Nothing
    Set oPerinatalAtencionDx = Nothing
    Set oDoPerinatalAtencionMedicamentos = Nothing
    Set oPerinatalAtencionMedicamentos = Nothing
    Set oDoPerinatalAtencionCred = Nothing
    Set oPerinatalAtencionCred = Nothing
    Set oRsTmpAutomatico = Nothing
    Set oRsTmpCodigoHis = Nothing
    Set oDOAtencionDiagnostico = Nothing
    Set oAtencionesDiagnosticos = Nothing
    Set oPerinatalAtencionCred1 = Nothing
    Set oPerinatalAtencionDiaria = Nothing
    Set oDoPerinatalAtencionDiaria = Nothing
    Set oCommand = Nothing
End Function

Function PerinatalCEModificar(oDoPerinatalAtencion As DoPerinatalAtencion, oRsCptInmunizaciones As Recordset, oRsCptFrecuentes As Recordset, oRsDxDesarrollo As Recordset, oRsDxMorbilidadFrec As Recordset, oRsMedicamentos As Recordset, oRsCred As Recordset, oDoPerinatalAtencionCred1 As DoPerinatalAtencionCred1, lnIdAtencion As Long) As Boolean
    On Error GoTo ErrCeAgr
    Dim oConexion As New ADODB.Connection
    Dim oPerinatalAtencion As New PerinatalAtencion
    Dim oDoPerinatalAtencionCpt As New DoPerinatalAtencionCpt, oPerinatalAtencioncpt As New PerinatalAtencionCpt
    Dim oDoPerinatalAtencionDx As New DoPerinatalAtencionDx, oPerinatalAtencionDx As New PerinatalAtencionDx
    Dim oDoPerinatalAtencionMedicamentos As New DOPerinatalAtencionMed, oPerinatalAtencionMedicamentos As New PerinatalAtencionMed
    Dim oDoPerinatalAtencionCred As New DoPerinatalAtencionCred, oPerinatalAtencionCred As New PerinatalAtencionCred
    Dim oDOAtencionDiagnostico As New DOAtencionDiagnostico, oAtencionesDiagnosticos As New AtencionesDiagnosticos
    Dim oPerinatalAtencionCred1 As New PerinatalAtencionCred1
    Dim oPerinatalAtencionDiaria As New PerinatalAtencionDiaria, oDoPerinatalAtencionDiaria As New DoPerinatalAtencionDiaria
    Dim oRsTmpAutomatico As New Recordset, oRsTmpCodigoHis As New Recordset, oRstmp9 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lbContinuar As Boolean
    PerinatalCEModificar = False
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    '
    Set oRsTmpAutomatico = PerinatalAtencionDiariaXIdAtencion(lnIdAtencion, oConexion)
    '
    oConexion.BeginTrans
    '
    Set oPerinatalAtencion.Conexion = oConexion
    Set oPerinatalAtencioncpt.Conexion = oConexion
    Set oPerinatalAtencionDx.Conexion = oConexion
    Set oPerinatalAtencionMedicamentos.Conexion = oConexion
    Set oPerinatalAtencionCred.Conexion = oConexion
    Set oAtencionesDiagnosticos.Conexion = oConexion
    Set oPerinatalAtencionCred1.Conexion = oConexion
    Set oPerinatalAtencionDiaria.Conexion = oConexion
    '
    If Not oPerinatalAtencion.modificar(oDoPerinatalAtencion) Then
        ms_MensajeError = oPerinatalAtencion.MensajeError: GoTo ErrCeAgr
    End If
    '
    '** esta al inicio del SUB
    '
    If oRsTmpAutomatico.RecordCount = 0 Then
        oDoPerinatalAtencionDiaria.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
        oDoPerinatalAtencionDiaria.idAtencion = lnIdAtencion
        oDoPerinatalAtencionDiaria.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
        If Not oPerinatalAtencionDiaria.Insertar(oDoPerinatalAtencionDiaria) Then
            ms_MensajeError = oPerinatalAtencionDiaria.MensajeError: GoTo ErrCeAgr
        End If
    End If
    oRsTmpAutomatico.Close
    
    Set oRsTmpAutomatico = PerinatalCatalogoCptAutomaticosXidModulo(oDoPerinatalAtencion.idModulo, oConexion)
    Set oRsTmpCodigoHis = PerinatalCatalogoCptXidModulo(oDoPerinatalAtencion.idModulo, oConexion)
    '
    oDoPerinatalAtencionCpt.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
    oDoPerinatalAtencionCpt.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
    If Not oPerinatalAtencioncpt.Eliminar(oDoPerinatalAtencionCpt) Then
       ms_MensajeError = oPerinatalAtencioncpt.MensajeError: GoTo ErrCeAgr
    End If
    If oRsCptInmunizaciones.RecordCount > 0 Then
       oRsTmpCodigoHis.Filter = "idLista=" & sghPerinatalListas.sighInmunizaciones
       oDoPerinatalAtencionCpt.idLista = sghPerinatalListas.sighInmunizaciones
       oDoPerinatalAtencionCpt.idModulo = oDoPerinatalAtencion.idModulo
       oRsCptInmunizaciones.MoveFirst
       Do While Not oRsCptInmunizaciones.EOF
          oDoPerinatalAtencionCpt.idProducto = oRsCptInmunizaciones.Fields!Id
          oDoPerinatalAtencionCpt.CptEsAutomatico = False
          oDoPerinatalAtencionCpt.idAtencion = oRsCptInmunizaciones.Fields!idAtencion
          'elegidos
          oDoPerinatalAtencionCpt.CodigoHIS = ""
          If oRsTmpCodigoHis.RecordCount > 0 Then
             oRsTmpCodigoHis.MoveFirst
             oRsTmpCodigoHis.Find "idProducto=" & oRsCptInmunizaciones.Fields!Id
             If Not oRsTmpCodigoHis.EOF Then
                oDoPerinatalAtencionCpt.CodigoHIS = oRsTmpCodigoHis.Fields!CodigoHIS
             End If
          End If
          '
          If Not oPerinatalAtencioncpt.Insertar(oDoPerinatalAtencionCpt) Then
             ms_MensajeError = oPerinatalAtencioncpt.MensajeError: GoTo ErrCeAgr
          End If
          'automaticos
          If oRsTmpAutomatico.RecordCount > 0 Then
             oRsTmpAutomatico.MoveFirst
             oRsTmpAutomatico.Find "idProducto=" & oRsCptInmunizaciones.Fields!Id
             If oRsTmpAutomatico.EOF = False Then
                oDoPerinatalAtencionCpt.idProducto = oRsTmpAutomatico.Fields!idProductoAutomatico
                oDoPerinatalAtencionCpt.CptEsAutomatico = True
                oDoPerinatalAtencionCpt.CodigoHIS = ""
                If Not oPerinatalAtencioncpt.Insertar(oDoPerinatalAtencionCpt) Then
                   ms_MensajeError = oPerinatalAtencioncpt.MensajeError: GoTo ErrCeAgr
                End If
             End If
          End If
          '
          oRsCptInmunizaciones.MoveNext
       Loop
    End If
    '
    If oRsCptFrecuentes.RecordCount > 0 Then
       oRsTmpCodigoHis.Filter = "idLista=" & sghPerinatalListas.sighCptFrecuentes
       oDoPerinatalAtencionCpt.idLista = sghPerinatalListas.sighCptFrecuentes
       oDoPerinatalAtencionCpt.idModulo = oDoPerinatalAtencion.idModulo
       oDoPerinatalAtencionCpt.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
       oDoPerinatalAtencionCpt.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
       oRsCptFrecuentes.MoveFirst
       Do While Not oRsCptFrecuentes.EOF
          oDoPerinatalAtencionCpt.idProducto = oRsCptFrecuentes.Fields!Id
          oDoPerinatalAtencionCpt.CptEsAutomatico = False
          oDoPerinatalAtencionCpt.idAtencion = oRsCptFrecuentes.Fields!idAtencion
          
          oDoPerinatalAtencionCpt.CodigoHIS = ""
          If oRsTmpCodigoHis.RecordCount > 0 Then
             oRsTmpCodigoHis.MoveFirst
             oRsTmpCodigoHis.Find "idProducto=" & oRsCptFrecuentes.Fields!Id
             If Not oRsTmpCodigoHis.EOF Then
                oDoPerinatalAtencionCpt.CodigoHIS = oRsTmpCodigoHis.Fields!CodigoHIS
             End If
          End If
          'mgaray201411a
            If Not IsNull(oRsCptFrecuentes.Fields!labConfHIS) Then
                oDoPerinatalAtencionCpt.labConfHIS = oRsCptFrecuentes.Fields!labConfHIS
            Else
                oDoPerinatalAtencionCpt.labConfHIS = Space(3)
            End If
            
          If Not oPerinatalAtencioncpt.Insertar(oDoPerinatalAtencionCpt) Then
             ms_MensajeError = oPerinatalAtencioncpt.MensajeError: GoTo ErrCeAgr
          End If
          oRsCptFrecuentes.MoveNext
       Loop
    End If
    '
    
    lbContinuar = True
    If oRsDxMorbilidadFrec.RecordCount = 0 Then
        Dim oRsTmp99 As New Recordset
        Set oRsTmp99 = oAtencionesDiagnosticos.SeleccionarTodosPorIdAtencion(lnIdAtencion)
        If oRsTmp99.RecordCount > 0 Then
           lbContinuar = False
        End If
        Set oRsTmp99 = Nothing
    End If
    oDOAtencionDiagnostico.idAtencion = lnIdAtencion
    oDOAtencionDiagnostico.IdClasificacionDx = 1
    oDOAtencionDiagnostico.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
'    mgaray201411a
    If lbContinuar = True Then
        If Not oAtencionesDiagnosticos.EliminarDiagnosticosPorAtencion(lnIdAtencion) Then
           ms_MensajeError = oAtencionesDiagnosticos.MensajeError: GoTo ErrCeAgr
        End If
    End If
    oRsTmpCodigoHis.Close
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PerinatalCatalogoCie10xidModulo"
        Set oParameter = .CreateParameter("@IdModulo", adInteger, adParamInput, 0, oDoPerinatalAtencion.idModulo): .Parameters.Append oParameter
        Set oRsTmpCodigoHis = .Execute
        Set oRsTmpCodigoHis.ActiveConnection = Nothing
    End With
    Set oParameter = Nothing
    Set oCommand = Nothing
    '
    oDoPerinatalAtencionDx.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
    oDoPerinatalAtencionDx.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
    If Not oPerinatalAtencionDx.Eliminar(oDoPerinatalAtencionDx) Then
       ms_MensajeError = oPerinatalAtencionDx.MensajeError: GoTo ErrCeAgr
    End If
    If oRsDxDesarrollo.RecordCount > 0 Then
       oRsTmpCodigoHis.Filter = "idLista=" & sghPerinatalListas.sighMorbilidadDesarrollo
       oDoPerinatalAtencionDx.idLista = sghPerinatalListas.sighMorbilidadDesarrollo
       oDoPerinatalAtencionDx.idModulo = oDoPerinatalAtencion.idModulo
       oRsDxDesarrollo.MoveFirst
       Do While Not oRsDxDesarrollo.EOF
          oDoPerinatalAtencionDx.IdDiagnostico = oRsDxDesarrollo.Fields!Id
          oDoPerinatalAtencionDx.idAtencion = oRsDxDesarrollo.Fields!idAtencion
          oDoPerinatalAtencionDx.labConfHIS = IIf(IsNull(oRsDxDesarrollo.Fields!labConfHIS), "", oRsDxDesarrollo.Fields!labConfHIS)
          'mgaray201411a
            oDoPerinatalAtencionDx.IdClasificacionDx = IIf(IsNull(oRsDxDesarrollo.Fields!IdClasificacionDx), 0, oRsDxDesarrollo.Fields!IdClasificacionDx)
            oDoPerinatalAtencionDx.IdSubclasificacionDx = IIf(IsNull(oRsDxDesarrollo.Fields!IdSubclasificacionDx), 0, oRsDxDesarrollo.Fields!IdSubclasificacionDx)
            
          oDoPerinatalAtencionDx.CodigoHIS = ""
          If oRsTmpCodigoHis.RecordCount > 0 Then
             oRsTmpCodigoHis.MoveFirst
             oRsTmpCodigoHis.Find "idDiagnostico=" & oRsDxDesarrollo.Fields!Id
             If Not oRsTmpCodigoHis.EOF Then
                oDoPerinatalAtencionDx.CodigoHIS = oRsTmpCodigoHis.Fields!CodigoHIS
             End If
          End If
          '
          If Not oPerinatalAtencionDx.Insertar(oDoPerinatalAtencionDx) Then
             ms_MensajeError = oPerinatalAtencionDx.MensajeError: GoTo ErrCeAgr
          End If
          'mgaray201411a
            If Not AgregarDiagnosticosAtencionDesdePerinatal(oRsDxDesarrollo, oAtencionesDiagnosticos, oDOAtencionDiagnostico, oRsTmpCodigoHis) Then
                ms_MensajeError = oAtencionesDiagnosticos.MensajeError: GoTo ErrCeAgr
            End If
          
'          If Not oRsTmpCodigoHis.EOF Then
'            If Not IsNull(oRsDxDesarrollo.Fields!LabConfHIS) Then
'                oDOAtencionDiagnostico.LabConfHIS = oRsDxDesarrollo.Fields!LabConfHIS
'            Else
'                oDOAtencionDiagnostico.LabConfHIS = Left(oRsTmpCodigoHis.Fields!LabHis, InStr(oRsTmpCodigoHis.Fields!LabHis, "-") - 1)
'             End If
'             oDOAtencionDiagnostico.IdDiagnostico = oRsDxDesarrollo.Fields!ID
'             oDOAtencionDiagnostico.IdSubClasificacionDX = 102 'Definitivo
'             oDOAtencionDiagnostico.idAtencion = oRsDxDesarrollo.Fields!idAtencion
'             If Not AtencionesDiagnosticoAgregarOModificar(oAtencionesDiagnosticos, oDOAtencionDiagnostico) Then
'                ms_MensajeError = oAtencionesDiagnosticos.MensajeError: GoTo ErrCeAgr
'             End If
'          Else
'            If Not IsNull(oRsDxDesarrollo.Fields!LabConfHIS) Then
'                oDOAtencionDiagnostico.LabConfHIS = oRsDxDesarrollo.Fields!LabConfHIS
'            Else
'                oDOAtencionDiagnostico.LabConfHIS = Space(3)
'             End If
'             oDOAtencionDiagnostico.IdDiagnostico = oRsDxDesarrollo.Fields!ID
'             oDOAtencionDiagnostico.IdSubClasificacionDX = 102 'Definitivo
'             oDOAtencionDiagnostico.idAtencion = oRsDxDesarrollo.Fields!idAtencion
'             If Not AtencionesDiagnosticoAgregarOModificar(oAtencionesDiagnosticos, oDOAtencionDiagnostico) Then
'                ms_MensajeError = oAtencionesDiagnosticos.MensajeError: GoTo ErrCeAgr
'             End If
'          End If

          oRsDxDesarrollo.MoveNext
       Loop
    End If
    '
    If oRsDxMorbilidadFrec.RecordCount > 0 Then
       oRsTmpCodigoHis.Filter = "idLista=" & sghPerinatalListas.sighMorbilidadFrecuente
       oDoPerinatalAtencionDx.idLista = sghPerinatalListas.sighMorbilidadFrecuente
       oDoPerinatalAtencionDx.idModulo = oDoPerinatalAtencion.idModulo
       oDoPerinatalAtencionDx.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
       oDoPerinatalAtencionDx.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
       oRsDxMorbilidadFrec.MoveFirst
       Do While Not oRsDxMorbilidadFrec.EOF
          oDoPerinatalAtencionDx.IdDiagnostico = oRsDxMorbilidadFrec.Fields!Id
          oDoPerinatalAtencionDx.idAtencion = oRsDxMorbilidadFrec.Fields!idAtencion
          oDoPerinatalAtencionDx.labConfHIS = IIf(IsNull(oRsDxMorbilidadFrec.Fields!labConfHIS), "", oRsDxMorbilidadFrec.Fields!labConfHIS)
          'mgaray201411a
            oDoPerinatalAtencionDx.IdClasificacionDx = IIf(IsNull(oRsDxMorbilidadFrec.Fields!IdClasificacionDx), 0, oRsDxMorbilidadFrec.Fields!IdClasificacionDx)
            oDoPerinatalAtencionDx.IdSubclasificacionDx = IIf(IsNull(oRsDxMorbilidadFrec.Fields!IdSubclasificacionDx), 0, oRsDxMorbilidadFrec.Fields!IdSubclasificacionDx)
            
          oDoPerinatalAtencionDx.CodigoHIS = ""
          If oRsTmpCodigoHis.RecordCount > 0 Then
             oRsTmpCodigoHis.MoveFirst
             oRsTmpCodigoHis.Find "idDiagnostico=" & oRsDxMorbilidadFrec.Fields!Id
             If Not oRsTmpCodigoHis.EOF Then
                oDoPerinatalAtencionDx.CodigoHIS = oRsTmpCodigoHis.Fields!CodigoHIS
             End If
          End If
          '
          If Not oPerinatalAtencionDx.Insertar(oDoPerinatalAtencionDx) Then
             ms_MensajeError = oPerinatalAtencionDx.MensajeError: GoTo ErrCeAgr
          End If
          'mgaray201411a
            If Not AgregarDiagnosticosAtencionDesdePerinatal(oRsDxMorbilidadFrec, oAtencionesDiagnosticos, oDOAtencionDiagnostico, oRsTmpCodigoHis) Then
                ms_MensajeError = oAtencionesDiagnosticos.MensajeError: GoTo ErrCeAgr
            End If
            
'          If Not oRsTmpCodigoHis.EOF Then
'            'mgaray20141003
'            If Not IsNull(oRsDxMorbilidadFrec.Fields!LabConfHIS) Then
'                oDOAtencionDiagnostico.LabConfHIS = oRsDxMorbilidadFrec.Fields!LabConfHIS
'            Else
'                'mgaray20141015
'                oDOAtencionDiagnostico.LabConfHIS = Left(oRsTmpCodigoHis.Fields!LabHis, InStr(oRsTmpCodigoHis.Fields!LabHis, "-") - 1)
'            End If
'             oDOAtencionDiagnostico.IdDiagnostico = oRsDxMorbilidadFrec.Fields!ID
'
'             oDOAtencionDiagnostico.idAtencion = oRsDxMorbilidadFrec.Fields!idAtencion
'             oDOAtencionDiagnostico.IdSubClasificacionDX = 102 'Definitivo
'             'mgaray20141015
''             If Not oAtencionesDiagnosticos.Insertar(oDOAtencionDiagnostico) Then
'             If Not AtencionesDiagnosticoAgregarOModificar(oAtencionesDiagnosticos, oDOAtencionDiagnostico) Then
'                 ms_MensajeError = oAtencionesDiagnosticos.MensajeError: GoTo ErrCeAgr
'             End If
'
''             If InStr(oRsTmpCodigoHis.Fields!TipoDx, "D") > 0 Then
''                oDOAtencionDiagnostico.IdSubclasificacionDx = 102 'Definitivo
''                If Not oAtencionesDiagnosticos.Insertar(oDOAtencionDiagnostico) Then
''                   ms_MensajeError = oAtencionesDiagnosticos.MensajeError: GoTo ErrCeAgr
''                End If
''             End If
''             If InStr(oRsTmpCodigoHis.Fields!TipoDx, "P") > 0 Then
''                oDOAtencionDiagnostico.IdSubclasificacionDx = 101 'Presuntivo
''                If Not oAtencionesDiagnosticos.Insertar(oDOAtencionDiagnostico) Then
''                   ms_MensajeError = oAtencionesDiagnosticos.MensajeError: GoTo ErrCeAgr
''                End If
''             End If
''             If InStr(oRsTmpCodigoHis.Fields!TipoDx, "R") > 0 Then
''                oDOAtencionDiagnostico.IdSubclasificacionDx = 103 'Repetido
''                If Not oAtencionesDiagnosticos.Insertar(oDOAtencionDiagnostico) Then
''                   ms_MensajeError = oAtencionesDiagnosticos.MensajeError: GoTo ErrCeAgr
''                End If
''             End If
'          Else
'            'mgaray20141003
'            If Not IsNull(oRsDxMorbilidadFrec.Fields!LabConfHIS) Then
'                oDOAtencionDiagnostico.LabConfHIS = oRsDxMorbilidadFrec.Fields!LabConfHIS
'            Else
'                oDOAtencionDiagnostico.LabConfHIS = Space(3)
'             End If
'             oDOAtencionDiagnostico.IdDiagnostico = oRsDxMorbilidadFrec.Fields!ID
'             oDOAtencionDiagnostico.IdSubClasificacionDX = 102 'Definitivo
'             oDOAtencionDiagnostico.idAtencion = oRsDxMorbilidadFrec.Fields!idAtencion
'             'mgaray20141015
''             If Not oAtencionesDiagnosticos.Insertar(oDOAtencionDiagnostico) Then
'             If Not AtencionesDiagnosticoAgregarOModificar(oAtencionesDiagnosticos, oDOAtencionDiagnostico) Then
'                ms_MensajeError = oAtencionesDiagnosticos.MensajeError: GoTo ErrCeAgr
'             End If
'          End If
          
          oRsDxMorbilidadFrec.MoveNext
       Loop
    End If
    '
    oDoPerinatalAtencionMedicamentos.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
    oDoPerinatalAtencionMedicamentos.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
    If Not oPerinatalAtencionMedicamentos.Eliminar(oDoPerinatalAtencionMedicamentos) Then
       ms_MensajeError = oPerinatalAtencionMedicamentos.MensajeError: GoTo ErrCeAgr
    End If
    oRsMedicamentos.Filter = "seleccionar=true"
    If oRsMedicamentos.RecordCount > 0 Then
       oDoPerinatalAtencionMedicamentos.idModulo = oDoPerinatalAtencion.idModulo
       oRsMedicamentos.MoveFirst
       Do While Not oRsMedicamentos.EOF
          oDoPerinatalAtencionMedicamentos.idProducto = oRsMedicamentos.Fields!Id
          oDoPerinatalAtencionMedicamentos.idAtencion = oRsMedicamentos.Fields!idAtencion
          If Not oPerinatalAtencionMedicamentos.Insertar(oDoPerinatalAtencionMedicamentos) Then
             ms_MensajeError = oPerinatalAtencionMedicamentos.MensajeError: GoTo ErrCeAgr
          End If
          oRsMedicamentos.MoveNext
       Loop
    End If
    '
    oDoPerinatalAtencionCred.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
    oDoPerinatalAtencionCred.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
    oDoPerinatalAtencionCred.idAtencion = lnIdAtencion
    If Not oPerinatalAtencionCred.EliminarXatencion(oDoPerinatalAtencionCred) Then
       ms_MensajeError = oPerinatalAtencionCred.MensajeError: GoTo ErrCeAgr
    End If
    If oRsCred.RecordCount > 0 Then
       oRsCred.MoveFirst
       Do While Not oRsCred.EOF
          If oRsCred.Fields!idAtencion = lnIdAtencion Then
                oDoPerinatalAtencionCred.credCheck = oRsCred.Fields!credCheck
                oDoPerinatalAtencionCred.credNumero = oRsCred.Fields!credNumero
                oDoPerinatalAtencionCred.edadEnAnios = oRsCred.Fields!edadEnAnios
                oDoPerinatalAtencionCred.idAtencion = oRsCred.Fields!idAtencion
                If Not oPerinatalAtencionCred.Insertar(oDoPerinatalAtencionCred) Then
                   ms_MensajeError = oPerinatalAtencionCred.MensajeError: GoTo ErrCeAgr
                End If
          End If
          oRsCred.MoveNext
       Loop
    End If
    '
    If oDoPerinatalAtencionCred1.idAtencion > 0 Then
       oDoPerinatalAtencionCred1.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
       Set oRstmp9 = oPerinatalAtencionCred1.PerinatalAtencionCred1SeleccionarTodoPorIdAtencion(lnIdAtencion)
       If oRstmp9.RecordCount = 0 Then
            If Not oPerinatalAtencionCred1.Insertar(oDoPerinatalAtencionCred1) Then
               ms_MensajeError = oPerinatalAtencionCred1.MensajeError: GoTo ErrCeAgr
            End If
       Else
            If Not oPerinatalAtencionCred1.modificar(oDoPerinatalAtencionCred1) Then
               ms_MensajeError = oPerinatalAtencionCred1.MensajeError: GoTo ErrCeAgr
            End If
       End If
       oRstmp9.Close
    End If
    
    If grabarDatosPerinatalAtencionIntegral(oDoPerinatalAtencion, lnIdAtencion, oConexion) = False Then
        ms_MensajeError = "No se Puedieron Modificar los datos de atencion Integral"
    End If
    '
    PerinatalCEModificar = True
    oRsTmpCodigoHis.Close
    oRsTmpAutomatico.Close
    oConexion.Close
ErrCeAgr:
    If PerinatalCEModificar = True Then
       oConexion.CommitTrans
    Else
       oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPerinatalAtencion = Nothing
    Set oDoPerinatalAtencionCpt = Nothing
    Set oPerinatalAtencioncpt = Nothing
    Set oDoPerinatalAtencionDx = Nothing
    Set oPerinatalAtencionDx = Nothing
    Set oDoPerinatalAtencionMedicamentos = Nothing
    Set oPerinatalAtencionMedicamentos = Nothing
    Set oDoPerinatalAtencionCred = Nothing
    Set oPerinatalAtencionCred = Nothing
    Set oRsTmpAutomatico = Nothing
    Set oRsTmpCodigoHis = Nothing
    Set oDOAtencionDiagnostico = Nothing
    Set oAtencionesDiagnosticos = Nothing
    Set oPerinatalAtencionCred1 = Nothing
    Set oPerinatalAtencionDiaria = Nothing
    Set oDoPerinatalAtencionDiaria = Nothing
    Set oRstmp9 = Nothing
    Set oCommand = Nothing
End Function


Function PerinatalCEeliminar(oDoPerinatalAtencion As DoPerinatalAtencion, lnIdAtencion As Long) As Boolean
    On Error GoTo ErrCeAgr
    Dim oConexion As New ADODB.Connection
    Dim oPerinatalAtencion As New PerinatalAtencion
    Dim oDoPerinatalAtencionCpt As New DoPerinatalAtencionCpt, oPerinatalAtencioncpt As New PerinatalAtencionCpt
    Dim oDoPerinatalAtencionDx As New DoPerinatalAtencionDx, oPerinatalAtencionDx As New PerinatalAtencionDx
    Dim oDoPerinatalAtencionMedicamentos As New DOPerinatalAtencionMed, oPerinatalAtencionMedicamentos As New PerinatalAtencionMed
    Dim oDoPerinatalAtencionCred As New DoPerinatalAtencionCred, oPerinatalAtencionCred As New PerinatalAtencionCred
    Dim oDOAtencionDiagnostico As New DOAtencionDiagnostico, oAtencionesDiagnosticos As New AtencionesDiagnosticos
    Dim oPerinatalAtencionCred1 As New PerinatalAtencionCred1
    Dim oPerinatalAtencionDiaria As New PerinatalAtencionDiaria, oDoPerinatalAtencionDiaria As New DoPerinatalAtencionDiaria
    Dim oRsTmpAutomatico As New Recordset
    PerinatalCEeliminar = False
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    '
    Set oRsTmpAutomatico = PerinatalAtencionDiariaXidPerinatalAtencion(oDoPerinatalAtencion.idPerinatalAtencion, oConexion)
    '
    oConexion.BeginTrans
    '
    Set oPerinatalAtencion.Conexion = oConexion
    Set oPerinatalAtencioncpt.Conexion = oConexion
    Set oPerinatalAtencionDx.Conexion = oConexion
    Set oPerinatalAtencionMedicamentos.Conexion = oConexion
    Set oPerinatalAtencionCred.Conexion = oConexion
    Set oAtencionesDiagnosticos.Conexion = oConexion
    Set oPerinatalAtencionCred1.Conexion = oConexion
    Set oPerinatalAtencionDiaria.Conexion = oConexion
    '
    oDoPerinatalAtencionCpt.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
    oDoPerinatalAtencionCpt.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
    If Not oPerinatalAtencioncpt.EliminarXatencion(oDoPerinatalAtencionCpt) Then
       ms_MensajeError = oPerinatalAtencioncpt.MensajeError: GoTo ErrCeAgr
    End If
    '
    oDoPerinatalAtencionDx.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
    oDoPerinatalAtencionDx.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
    If Not oPerinatalAtencionDx.EliminarXatencion(oDoPerinatalAtencionDx) Then
       ms_MensajeError = oPerinatalAtencionDx.MensajeError: GoTo ErrCeAgr
    End If
    '
    oDoPerinatalAtencionMedicamentos.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
    oDoPerinatalAtencionMedicamentos.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
    If Not oPerinatalAtencionMedicamentos.EliminarXatencion(oDoPerinatalAtencionMedicamentos) Then
       ms_MensajeError = oPerinatalAtencionMedicamentos.MensajeError: GoTo ErrCeAgr
    End If
    '
    oDoPerinatalAtencionCred.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
    oDoPerinatalAtencionCred.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
    oDoPerinatalAtencionCred.idAtencion = lnIdAtencion
    If Not oPerinatalAtencionCred.EliminarXatencion(oDoPerinatalAtencionCred) Then
       ms_MensajeError = oPerinatalAtencionCred.MensajeError: GoTo ErrCeAgr
    End If
    '
    If Not oAtencionesDiagnosticos.EliminarDiagnosticosPorAtencion(lnIdAtencion) Then
       ms_MensajeError = oAtencionesDiagnosticos.MensajeError: GoTo ErrCeAgr
    End If
    '
    oDoPerinatalAtencionDiaria.idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
    oDoPerinatalAtencionDiaria.idAtencion = lnIdAtencion
    oDoPerinatalAtencionDiaria.IdUsuarioAuditoria = oDoPerinatalAtencion.IdUsuarioAuditoria
    If Not oPerinatalAtencionDiaria.EliminarXidAtencion(oDoPerinatalAtencionDiaria) Then
        ms_MensajeError = oPerinatalAtencionDiaria.MensajeError: GoTo ErrCeAgr
    End If
    '***esta al inicio del SUB
    '
    If oRsTmpAutomatico.RecordCount = 0 Then
        If Not oPerinatalAtencion.Eliminar(oDoPerinatalAtencion) Then
            ms_MensajeError = oPerinatalAtencion.MensajeError: GoTo ErrCeAgr
        End If
    End If
    
    If eliminarDatosAtencionIntegral(oDoPerinatalAtencion, lnIdAtencion, oConexion) = False Then
        ms_MensajeError = "No se Pudieron Eliminar los Datos Relacionados al Plan Integral"
    End If
    oRsTmpAutomatico.Close
    '
    PerinatalCEeliminar = True
    oConexion.Close
ErrCeAgr:
    If PerinatalCEeliminar = True Then
       oConexion.CommitTrans
    Else
       oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPerinatalAtencion = Nothing
    Set oDoPerinatalAtencionCpt = Nothing
    Set oPerinatalAtencioncpt = Nothing
    Set oDoPerinatalAtencionDx = Nothing
    Set oPerinatalAtencionDx = Nothing
    Set oDoPerinatalAtencionMedicamentos = Nothing
    Set oPerinatalAtencionMedicamentos = Nothing
    Set oDoPerinatalAtencionCred = Nothing
    Set oPerinatalAtencionCred = Nothing
    Set oDOAtencionDiagnostico = Nothing
    Set oAtencionesDiagnosticos = Nothing
    Set oPerinatalAtencionCred1 = Nothing
    Set oPerinatalAtencionDiaria = Nothing
    Set oDoPerinatalAtencionDiaria = Nothing
    Set oRsTmpAutomatico = Nothing
End Function

Function AtencionesDePacienteElmismoDiaCE(lnIdPaciente As Long, ldFecha As Date) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDePacienteElmismoDiaCE"
        Set oParameter = .CreateParameter("@lnIdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ldFecha", adDBTimeStamp, adParamInput, 0, CDate(Format(ldFecha, sighentidades.FormatoFechaCorta))): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesDePacienteElmismoDiaCE = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
   Exit Function
ManejadorDeError:
   
End Function


Function RecetaAgregar(lnIdCuentaAtencion As Long, lnIdServicio As Long, lnIdUsuario As Long, _
                       ByRef lnRecetaRayosX As Long, ByRef lnRecetaEcografiaO As Long, ByRef lnRecetaEcografiaG As Long, _
                       ByRef lnRecetaTomografia As Long, ByRef lnRecetaAnatomiaP As Long, ByRef lnRecetaPatologiaC As Long, _
                       ByRef lnRecetaBancoS As Long, ByRef lnRecetaFarmacia As Long, _
                       oRsRayosX As Recordset, oRsEcografiaO As Recordset, oRsEcografiaG As Recordset, _
                       oRsTomografia As Recordset, oRsAnatomia As Recordset, oRsPatologia As Recordset, _
                       oRsBanco As Recordset, oRsFarmacia As Recordset, _
                       mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcObservacion As String, _
                       lnIdMedicoReceta As Long, oRsOtrosCpt As Recordset, lnRecetaOtrosCpt As Long) As Boolean
    On Error GoTo ErrCeAgr
    Dim oConexion As New ADODB.Connection
    Dim oRecetaCabecera As New RecetaCabecera, oDoRecetaCabecera As New DoRecetaCabecera
    Dim oRecetaDetalle As New RecetaDetalle, oDoRecetaDetalle As New DoRecetaDetalle
    RecetaAgregar = False
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.BeginTrans
    '
    Set oRecetaCabecera.Conexion = oConexion
    Set oRecetaDetalle.Conexion = oConexion
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(lnIdUsuario, "A", lnIdCuentaAtencion, "RecetaCabecera", oConexion, _
           mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "(" + lcBuscaParametro.RetornaFechaHoraServidorSQL() + ") " + lcObservacion)
    'Datos comunes para todos
    oDoRecetaCabecera.fechaReceta = lcBuscaParametro.RetornaFechaHoraServidorSQL()
    oDoRecetaCabecera.idCuentaAtencion = lnIdCuentaAtencion
    oDoRecetaCabecera.IdEstado = sghRecetaEstados.sighRecetaRegistrada
    oDoRecetaCabecera.idServicioReceta = lnIdServicio
    oDoRecetaCabecera.IdUsuarioAuditoria = lnIdUsuario
    oDoRecetaCabecera.idMedicoReceta = lnIdMedicoReceta
    oDoRecetaDetalle.IdUsuarioAuditoria = oDoRecetaCabecera.IdUsuarioAuditoria
    'Rayos X
    If oRsRayosX.RecordCount > 0 Then
       oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaRayosX
       If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
            ms_MensajeError = "Rx: " & oRecetaCabecera.MensajeError: GoTo ErrCeAgr
       End If
       lnRecetaRayosX = oDoRecetaCabecera.idReceta
       oDoRecetaDetalle.idReceta = oDoRecetaCabecera.idReceta
       oRsRayosX.MoveFirst
       Do While Not oRsRayosX.EOF
          oDoRecetaDetalle.dx = IIf(IsNull(oRsRayosX!dx), "", oRsRayosX!dx)
          oDoRecetaDetalle.cantidadPedida = oRsRayosX.Fields!cantidad
          oDoRecetaDetalle.idItem = oRsRayosX.Fields!Id
          oDoRecetaDetalle.Precio = oRsRayosX.Fields!Precio
          oDoRecetaDetalle.SaldoEnRegistroReceta = oRsRayosX.Fields!saldoActual
          oDoRecetaDetalle.total = Round(oRsRayosX.Fields!cantidad * oRsRayosX.Fields!Precio, 2)
          oDoRecetaDetalle.observaciones = IIf(IsNull(oRsRayosX.Fields!observaciones), "", oRsRayosX.Fields!observaciones)
          If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
                 ms_MensajeError = "Rx: " & oRecetaDetalle.MensajeError: GoTo ErrCeAgr
          End If
          oRsRayosX.MoveNext
       Loop
    End If
    'Ecografia Obstetrica
    If oRsEcografiaO.RecordCount > 0 Then
       oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaEcogObstetrica
       If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
            ms_MensajeError = "E.O: " & oRecetaCabecera.MensajeError: GoTo ErrCeAgr
       End If
       lnRecetaEcografiaO = oDoRecetaCabecera.idReceta
       oDoRecetaDetalle.idReceta = oDoRecetaCabecera.idReceta
       oRsEcografiaO.MoveFirst
       Do While Not oRsEcografiaO.EOF
          oDoRecetaDetalle.dx = IIf(IsNull(oRsEcografiaO!dx), "", oRsEcografiaO!dx)
          oDoRecetaDetalle.cantidadPedida = oRsEcografiaO.Fields!cantidad
          oDoRecetaDetalle.idItem = oRsEcografiaO.Fields!Id
          oDoRecetaDetalle.Precio = oRsEcografiaO.Fields!Precio
          oDoRecetaDetalle.SaldoEnRegistroReceta = oRsEcografiaO.Fields!saldoActual
          oDoRecetaDetalle.total = Round(oRsEcografiaO.Fields!cantidad * oRsEcografiaO.Fields!Precio, 2)
          oDoRecetaDetalle.observaciones = IIf(IsNull(oRsEcografiaO.Fields!observaciones), "", oRsEcografiaO.Fields!observaciones)
          If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
                 ms_MensajeError = "E.O: " & oRecetaDetalle.MensajeError: GoTo ErrCeAgr
          End If
          oRsEcografiaO.MoveNext
       Loop
    End If
    'Ecografia General
    If oRsEcografiaG.RecordCount > 0 Then
       oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaEcogGeneral
       If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
            ms_MensajeError = "Eg: " & oRecetaCabecera.MensajeError: GoTo ErrCeAgr
       End If
       lnRecetaEcografiaG = oDoRecetaCabecera.idReceta
       oDoRecetaDetalle.idReceta = oDoRecetaCabecera.idReceta
       oRsEcografiaG.MoveFirst
       Do While Not oRsEcografiaG.EOF
          oDoRecetaDetalle.dx = IIf(IsNull(oRsEcografiaG!dx), "", oRsEcografiaG!dx)
          oDoRecetaDetalle.cantidadPedida = oRsEcografiaG.Fields!cantidad
          oDoRecetaDetalle.idItem = oRsEcografiaG.Fields!Id
          oDoRecetaDetalle.Precio = oRsEcografiaG.Fields!Precio
          oDoRecetaDetalle.SaldoEnRegistroReceta = oRsEcografiaG.Fields!saldoActual
          oDoRecetaDetalle.total = Round(oRsEcografiaG.Fields!cantidad * oRsEcografiaG.Fields!Precio, 2)
          oDoRecetaDetalle.observaciones = IIf(IsNull(oRsEcografiaG.Fields!observaciones), "", oRsEcografiaG.Fields!observaciones)
          If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
                 ms_MensajeError = "Eg: " & oRecetaDetalle.MensajeError: GoTo ErrCeAgr
          End If
          oRsEcografiaG.MoveNext
       Loop
    End If
    'Tomografia
    If oRsTomografia.RecordCount > 0 Then
       oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaTomografia
       If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
            ms_MensajeError = "To: " & oRecetaCabecera.MensajeError: GoTo ErrCeAgr
       End If
       lnRecetaTomografia = oDoRecetaCabecera.idReceta
       oDoRecetaDetalle.idReceta = oDoRecetaCabecera.idReceta
       oRsTomografia.MoveFirst
       Do While Not oRsTomografia.EOF
          oDoRecetaDetalle.dx = IIf(IsNull(oRsTomografia!dx), "", oRsTomografia!dx)
          oDoRecetaDetalle.cantidadPedida = oRsTomografia.Fields!cantidad
          oDoRecetaDetalle.idItem = oRsTomografia.Fields!Id
          oDoRecetaDetalle.Precio = oRsTomografia.Fields!Precio
          oDoRecetaDetalle.SaldoEnRegistroReceta = oRsTomografia.Fields!saldoActual
          oDoRecetaDetalle.total = Round(oRsTomografia.Fields!cantidad * oRsTomografia.Fields!Precio, 2)
          oDoRecetaDetalle.observaciones = IIf(IsNull(oRsTomografia.Fields!observaciones), "", oRsTomografia.Fields!observaciones)
          If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
                 ms_MensajeError = "To: " & oRecetaDetalle.MensajeError: GoTo ErrCeAgr
          End If
          oRsTomografia.MoveNext
       Loop
    End If
    'Anatomia Patologica
    If oRsAnatomia.RecordCount > 0 Then
       oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaAnatomiaPatologica1
       If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
            ms_MensajeError = "Ap: " & oRecetaCabecera.MensajeError: GoTo ErrCeAgr
       End If
       lnRecetaAnatomiaP = oDoRecetaCabecera.idReceta
       oDoRecetaDetalle.idReceta = oDoRecetaCabecera.idReceta
       oRsAnatomia.MoveFirst
       Do While Not oRsAnatomia.EOF
          oDoRecetaDetalle.dx = IIf(IsNull(oRsAnatomia!dx), "", oRsAnatomia!dx)
          oDoRecetaDetalle.cantidadPedida = oRsAnatomia.Fields!cantidad
          oDoRecetaDetalle.idItem = oRsAnatomia.Fields!Id
          oDoRecetaDetalle.Precio = oRsAnatomia.Fields!Precio
          oDoRecetaDetalle.SaldoEnRegistroReceta = oRsAnatomia.Fields!saldoActual
          oDoRecetaDetalle.total = Round(oRsAnatomia.Fields!cantidad * oRsAnatomia.Fields!Precio, 2)
          oDoRecetaDetalle.observaciones = IIf(IsNull(oRsAnatomia.Fields!observaciones), "", oRsAnatomia.Fields!observaciones)
          If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
                 ms_MensajeError = "Ap: " & oRecetaDetalle.MensajeError: GoTo ErrCeAgr
          End If
          oRsAnatomia.MoveNext
       Loop
    End If
    'Patologia Clinica
    If oRsPatologia.RecordCount > 0 Then
       oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaPatologiaClinica
       If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
            ms_MensajeError = "Pc: " & oRecetaCabecera.MensajeError: GoTo ErrCeAgr
       End If
       lnRecetaPatologiaC = oDoRecetaCabecera.idReceta
       oDoRecetaDetalle.idReceta = oDoRecetaCabecera.idReceta
       oRsPatologia.MoveFirst
       Do While Not oRsPatologia.EOF
          oDoRecetaDetalle.dx = IIf(IsNull(oRsPatologia!dx), "", oRsPatologia!dx)
          oDoRecetaDetalle.cantidadPedida = oRsPatologia.Fields!cantidad
          oDoRecetaDetalle.idItem = oRsPatologia.Fields!Id
          oDoRecetaDetalle.Precio = oRsPatologia.Fields!Precio
          oDoRecetaDetalle.SaldoEnRegistroReceta = oRsPatologia.Fields!saldoActual
          oDoRecetaDetalle.total = Round(oRsPatologia.Fields!cantidad * oRsPatologia.Fields!Precio, 2)
          oDoRecetaDetalle.observaciones = IIf(IsNull(oRsPatologia.Fields!observaciones), "", oRsPatologia.Fields!observaciones)
          If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
                 ms_MensajeError = "Pc: " & oRecetaDetalle.MensajeError: GoTo ErrCeAgr
          End If
          oRsPatologia.MoveNext
       Loop
    End If
    'Banco de Sangre
    If oRsBanco.RecordCount > 0 Then
       oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaBancoSangre1
       If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
            ms_MensajeError = "Bs: " & oRecetaCabecera.MensajeError: GoTo ErrCeAgr
       End If
       lnRecetaBancoS = oDoRecetaCabecera.idReceta
       oDoRecetaDetalle.idReceta = oDoRecetaCabecera.idReceta
       oRsBanco.MoveFirst
       Do While Not oRsBanco.EOF
          oDoRecetaDetalle.dx = IIf(IsNull(oRsBanco!dx), "", oRsBanco!dx)
          oDoRecetaDetalle.cantidadPedida = oRsBanco.Fields!cantidad
          oDoRecetaDetalle.idItem = oRsBanco.Fields!Id
          oDoRecetaDetalle.Precio = oRsBanco.Fields!Precio
          oDoRecetaDetalle.SaldoEnRegistroReceta = oRsBanco.Fields!saldoActual
          oDoRecetaDetalle.total = Round(oRsBanco.Fields!cantidad * oRsBanco.Fields!Precio, 2)
          oDoRecetaDetalle.observaciones = IIf(IsNull(oRsBanco.Fields!observaciones), "", oRsBanco.Fields!observaciones)
          If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
                 ms_MensajeError = "Bs: " & oRecetaDetalle.MensajeError: GoTo ErrCeAgr
          End If
          oRsBanco.MoveNext
       Loop
    End If
    'Farmacia
    If oRsFarmacia.RecordCount > 0 Then
       'debb-24/06/2015
       oRsFarmacia.MoveFirst
       oDoRecetaCabecera.FechaVigencia = oRsFarmacia!FechaVigencia
       '
       oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaFarmacia
       If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
            ms_MensajeError = "Fa: " & oRecetaCabecera.MensajeError: GoTo ErrCeAgr
       End If
       lnRecetaFarmacia = oDoRecetaCabecera.idReceta
       oDoRecetaDetalle.idReceta = oDoRecetaCabecera.idReceta
       oRsFarmacia.MoveFirst
       Do While Not oRsFarmacia.EOF
          oDoRecetaDetalle.dx = IIf(IsNull(oRsFarmacia!dx), "", oRsFarmacia!dx)
          oDoRecetaDetalle.cantidadPedida = oRsFarmacia.Fields!cantidad
          oDoRecetaDetalle.idItem = oRsFarmacia.Fields!Id
          oDoRecetaDetalle.Precio = oRsFarmacia.Fields!Precio
          oDoRecetaDetalle.SaldoEnRegistroReceta = oRsFarmacia.Fields!saldoActual
          oDoRecetaDetalle.total = Round(oRsFarmacia.Fields!cantidad * oRsFarmacia.Fields!Precio, 2)
          oDoRecetaDetalle.observaciones = IIf(IsNull(oRsFarmacia.Fields!observaciones), "", oRsFarmacia.Fields!observaciones)
          oDoRecetaDetalle.idDosisRecetada = oRsFarmacia.Fields!idDosisRecetada
          oDoRecetaDetalle.IdViaAdministracion = oRsFarmacia.Fields!IdViaAdministracion
          If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
                 ms_MensajeError = "Fa: " & oRecetaDetalle.MensajeError: GoTo ErrCeAgr
          End If
          oRsFarmacia.MoveNext
       Loop
    End If
    
    '*********************************************Otros CPT**********************************
    oDoRecetaCabecera.IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion
    'solo agregar cpt
    If oRsOtrosCpt.RecordCount > 0 Then
         oRsOtrosCpt.MoveFirst
         Do While Not oRsOtrosCpt.EOF
            If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
                 ms_MensajeError = "Otros Cpt: " & oRecetaCabecera.MensajeError: GoTo ErrCeAgr
            End If
            oDoRecetaDetalle.IdPuntoCarga = oRsOtrosCpt!IdPuntoCarga
            oDoRecetaDetalle.idReceta = oDoRecetaCabecera.idReceta
            oDoRecetaDetalle.dx = IIf(IsNull(oRsOtrosCpt!dx), "", oRsOtrosCpt!dx)
            oDoRecetaDetalle.cantidadPedida = oRsOtrosCpt.Fields!cantidad
            oDoRecetaDetalle.idItem = oRsOtrosCpt.Fields!Id
            oDoRecetaDetalle.Precio = oRsOtrosCpt.Fields!Precio
            oDoRecetaDetalle.SaldoEnRegistroReceta = 0
            oDoRecetaDetalle.total = Round(oRsOtrosCpt.Fields!cantidad * oRsOtrosCpt.Fields!Precio, 2)
            oDoRecetaDetalle.idEstadoDetalle = sghRecetasEstadosDetalle.sghRecetado
            oDoRecetaDetalle.idDosisRecetada = 0
            oDoRecetaDetalle.IdViaAdministracion = 0 'Actualizado 2609
            oDoRecetaDetalle.observaciones = IIf(IsNull(oRsOtrosCpt.Fields!observaciones), "", oRsOtrosCpt.Fields!observaciones)
            If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
                   ms_MensajeError = "Otros Cpt: " & oRecetaDetalle.MensajeError: GoTo ErrCeAgr
            End If
            oRsOtrosCpt.MoveNext
         Loop
         lnRecetaOtrosCpt = oDoRecetaCabecera.idReceta
    End If
    '
    RecetaAgregar = True
ErrCeAgr:
'Resume
    If RecetaAgregar = True Then
       oConexion.CommitTrans
    Else
       oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oRecetaCabecera = Nothing
    Set oDoRecetaCabecera = Nothing
    Set oRecetaDetalle = Nothing
    Set oDoRecetaDetalle = Nothing
End Function

Function RecetaModificar(lnIdCuentaAtencion As Long, lnIdServicio As Long, lnIdUsuario As Long, _
                       ByRef lnRecetaRayosX As Long, ByRef lnRecetaEcografiaO As Long, ByRef lnRecetaEcografiaG As Long, _
                       ByRef lnRecetaTomografia As Long, ByRef lnRecetaAnatomiaP As Long, ByRef lnRecetaPatologiaC As Long, _
                       ByRef lnRecetaBancoS As Long, ByRef lnRecetaFarmacia As Long, _
                       oRsRayosX As Recordset, oRsEcografiaO As Recordset, oRsEcografiaG As Recordset, _
                       oRsTomografia As Recordset, oRsAnatomia As Recordset, oRsPatologia As Recordset, _
                       oRsBanco As Recordset, oRsFarmacia As Recordset, ldFechaReceta As Date, _
                       mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcObservacion As String, _
                       lnIdMedicoReceta As Long, lbModificaCabecera As Boolean, _
                       oRsOtrosCpt As Recordset, lnRecetaOtrosCpt As Long) As Boolean
10        On Error GoTo ErrModif
          Dim oConexion As New ADODB.Connection
          Dim oRecetaCabecera As New RecetaCabecera, oDoRecetaCabecera As New DoRecetaCabecera
          Dim oRecetaDetalle As New RecetaDetalle, oDoRecetaDetalle As New DoRecetaDetalle
20        RecetaModificar = False
30        oConexion.CommandTimeout = 300
40        oConexion.CursorLocation = adUseClient
50        oConexion.Open sighentidades.CadenaConexion
          
60        oConexion.BeginTrans
          '
70        Set oRecetaCabecera.Conexion = oConexion
80        Set oRecetaDetalle.Conexion = oConexion
          '
90        ms_MensajeError = IIf(lnRecetaRayosX > 0, "Rx: " + Trim(Str(lnRecetaRayosX)), "")
100       ms_MensajeError = ms_MensajeError + IIf(lnRecetaEcografiaO > 0, "Eco.Obs: " + Trim(Str(lnRecetaEcografiaO)), "")
110       ms_MensajeError = ms_MensajeError + IIf(lnRecetaEcografiaG > 0, "Eco.Gen: " + Trim(Str(lnRecetaEcografiaG)), "")
120       ms_MensajeError = ms_MensajeError + IIf(lnRecetaTomografia > 0, "Tomog: " + Trim(Str(lnRecetaTomografia)), "")
130       ms_MensajeError = ms_MensajeError + IIf(lnRecetaAnatomiaP > 0, "Ana.Pat: " + Trim(Str(lnRecetaAnatomiaP)), "")
140       ms_MensajeError = ms_MensajeError + IIf(lnRecetaPatologiaC > 0, "Pat.Cli: " + Trim(Str(lnRecetaPatologiaC)), "")
150       ms_MensajeError = ms_MensajeError + IIf(lnRecetaBancoS > 0, "Bco.San: " + Trim(Str(lnRecetaBancoS)), "")
160       ms_MensajeError = ms_MensajeError + IIf(lnRecetaFarmacia > 0, "Farm: " + Trim(Str(lnRecetaFarmacia)), "")
170       Call mo_ReglasSeguridad.AuditoriaAgregarV(lnIdUsuario, "M", lnIdCuentaAtencion, "RecetaCabecera", oConexion, _
              mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "(" + ms_MensajeError + ") " + lcObservacion)
180       ms_MensajeError = ""
          
          'Datos comunes para todos
190       oDoRecetaCabecera.fechaReceta = ldFechaReceta
200       oDoRecetaCabecera.idCuentaAtencion = lnIdCuentaAtencion
        
210       oDoRecetaCabecera.IdEstado = sghRecetaEstados.sighRecetaRegistrada
220       oDoRecetaCabecera.idServicioReceta = lnIdServicio
230       oDoRecetaCabecera.IdUsuarioAuditoria = lnIdUsuario
240       oDoRecetaCabecera.idMedicoReceta = lnIdMedicoReceta
250       oDoRecetaDetalle.IdUsuarioAuditoria = oDoRecetaCabecera.IdUsuarioAuditoria
          'Rayos X
260       If lnRecetaRayosX <> -100 Then
270           If lnRecetaRayosX = 0 And oRsRayosX.RecordCount > 0 Then
280               oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaRayosX
290               If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
300                    ms_MensajeError = "Rx: " & oRecetaCabecera.MensajeError: GoTo ErrModif
310               End If
320               lnRecetaRayosX = oDoRecetaCabecera.idReceta
330           ElseIf lbModificaCabecera = True Then
340               oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaRayosX
350               oDoRecetaCabecera.idReceta = lnRecetaRayosX
360               If Not oRecetaCabecera.modificar(oDoRecetaCabecera) Then
370                    ms_MensajeError = "Rx: " & oRecetaCabecera.MensajeError: GoTo ErrModif
380               End If
390           End If
400           oDoRecetaDetalle.idReceta = lnRecetaRayosX
410           If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
420                  ms_MensajeError = "Rx: " & oRecetaDetalle.MensajeError: GoTo ErrModif
430           End If
440           If oRsRayosX.RecordCount > 0 Then
450              oRsRayosX.MoveFirst
460              Do While Not oRsRayosX.EOF
                    oDoRecetaDetalle.dx = IIf(IsNull(oRsRayosX!dx), "", oRsRayosX!dx)
470                 oDoRecetaDetalle.cantidadPedida = oRsRayosX.Fields!cantidad
480                 oDoRecetaDetalle.idItem = oRsRayosX.Fields!Id
490                 oDoRecetaDetalle.Precio = oRsRayosX.Fields!Precio
500                 oDoRecetaDetalle.SaldoEnRegistroReceta = oRsRayosX.Fields!saldoActual
510                 oDoRecetaDetalle.total = Round(oRsRayosX.Fields!cantidad * oRsRayosX.Fields!Precio, 2)
520                 oDoRecetaDetalle.idEstadoDetalle = sghRecetasEstadosDetalle.sghRecetado
530                 oDoRecetaDetalle.idDosisRecetada = 0
540                 oDoRecetaDetalle.IdViaAdministracion = 0 'Actualizado 2609
550                 oDoRecetaDetalle.observaciones = IIf(IsNull(oRsRayosX.Fields!observaciones), "", oRsRayosX.Fields!observaciones)
560                 If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
570                        ms_MensajeError = "Rx: " & oRecetaDetalle.MensajeError: GoTo ErrModif
580                 End If
590                 oRsRayosX.MoveNext
600              Loop
610           End If
620       End If
          'Ecografia Obstetrica
630       If lnRecetaEcografiaO <> -100 Then
640           If lnRecetaEcografiaO = 0 And oRsEcografiaO.RecordCount > 0 Then
650               oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaEcogObstetrica
660               If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
670                    ms_MensajeError = "E.O: " & oRecetaCabecera.MensajeError: GoTo ErrModif
680               End If
690               lnRecetaEcografiaO = oDoRecetaCabecera.idReceta
700           ElseIf lbModificaCabecera = True Then
710               oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaEcogObstetrica
720               oDoRecetaCabecera.idReceta = lnRecetaEcografiaO
730               If Not oRecetaCabecera.modificar(oDoRecetaCabecera) Then
740                    ms_MensajeError = "E.O: " & oRecetaCabecera.MensajeError: GoTo ErrModif
750               End If
760           End If
770           oDoRecetaDetalle.idReceta = lnRecetaEcografiaO
780           If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
790                  ms_MensajeError = "Eo: " & oRecetaDetalle.MensajeError: GoTo ErrModif
800           End If
810           If oRsEcografiaO.RecordCount > 0 Then
820              oRsEcografiaO.MoveFirst
830              Do While Not oRsEcografiaO.EOF
                    oDoRecetaDetalle.dx = IIf(IsNull(oRsEcografiaO!dx), "", oRsEcografiaO!dx)
840                 oDoRecetaDetalle.cantidadPedida = oRsEcografiaO.Fields!cantidad
850                 oDoRecetaDetalle.idItem = oRsEcografiaO.Fields!Id
860                 oDoRecetaDetalle.Precio = oRsEcografiaO.Fields!Precio
870                 oDoRecetaDetalle.SaldoEnRegistroReceta = oRsEcografiaO.Fields!saldoActual
880                 oDoRecetaDetalle.total = Round(oRsEcografiaO.Fields!cantidad * oRsEcografiaO.Fields!Precio, 2)
890                 oDoRecetaDetalle.idEstadoDetalle = sghRecetasEstadosDetalle.sghRecetado
900                 oDoRecetaDetalle.idDosisRecetada = 0
910                 oDoRecetaDetalle.IdViaAdministracion = 0 'Actualizado 2609
920                 oDoRecetaDetalle.observaciones = IIf(IsNull(oRsEcografiaO.Fields!observaciones), "", oRsEcografiaO.Fields!observaciones)
930                 If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
940                        ms_MensajeError = "E.O: " & oRecetaDetalle.MensajeError: GoTo ErrModif
950                 End If
960                 oRsEcografiaO.MoveNext
970              Loop
980           End If
990       End If
          'Ecografia General
1000      If lnRecetaEcografiaG <> -100 Then
1010          If lnRecetaEcografiaG = 0 And oRsEcografiaG.RecordCount > 0 Then
1020              oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaEcogGeneral
1030              If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
1040                   ms_MensajeError = "Eg: " & oRecetaCabecera.MensajeError: GoTo ErrModif
1050              End If
1060              lnRecetaEcografiaG = oDoRecetaCabecera.idReceta
1070          ElseIf lbModificaCabecera = True Then
1080              oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaEcogGeneral
1090              oDoRecetaCabecera.idReceta = lnRecetaEcografiaG
1100              If Not oRecetaCabecera.modificar(oDoRecetaCabecera) Then
1110                   ms_MensajeError = "Eg: " & oRecetaCabecera.MensajeError: GoTo ErrModif
1120              End If
1130          End If
1140          oDoRecetaDetalle.idReceta = lnRecetaEcografiaG
1150          If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
1160                 ms_MensajeError = "Eg: " & oRecetaDetalle.MensajeError: GoTo ErrModif
1170          End If
1180          If oRsEcografiaG.RecordCount > 0 Then
1190             oRsEcografiaG.MoveFirst
1200             Do While Not oRsEcografiaG.EOF
                    oDoRecetaDetalle.dx = IIf(IsNull(oRsEcografiaG!dx), "", oRsEcografiaG!dx)
1210                oDoRecetaDetalle.cantidadPedida = oRsEcografiaG.Fields!cantidad
1220                oDoRecetaDetalle.idItem = oRsEcografiaG.Fields!Id
1230                oDoRecetaDetalle.Precio = oRsEcografiaG.Fields!Precio
1240                oDoRecetaDetalle.SaldoEnRegistroReceta = oRsEcografiaG.Fields!saldoActual
1250                oDoRecetaDetalle.total = Round(oRsEcografiaG.Fields!cantidad * oRsEcografiaG.Fields!Precio, 2)
1260                oDoRecetaDetalle.idEstadoDetalle = sghRecetasEstadosDetalle.sghRecetado
1270                oDoRecetaDetalle.idDosisRecetada = 0
1280                oDoRecetaDetalle.IdViaAdministracion = 0 'Actualizado 2609
1290                oDoRecetaDetalle.observaciones = IIf(IsNull(oRsEcografiaG.Fields!observaciones), "", oRsEcografiaG.Fields!observaciones)
1300                If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
1310                       ms_MensajeError = "Eg: " & oRecetaDetalle.MensajeError: GoTo ErrModif
1320                End If
1330                oRsEcografiaG.MoveNext
1340             Loop
1350          End If
1360      End If
          'Tomografia
1370      If lnRecetaTomografia <> -100 Then
1380          If lnRecetaTomografia = 0 And oRsTomografia.RecordCount > 0 Then
1390              oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaTomografia
1400              If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
1410                   ms_MensajeError = "To: " & oRecetaCabecera.MensajeError: GoTo ErrModif
1420              End If
1430              lnRecetaTomografia = oDoRecetaCabecera.idReceta
1440          ElseIf lbModificaCabecera = True Then
1450              oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaTomografia
1460              oDoRecetaCabecera.idReceta = lnRecetaTomografia
1470              If Not oRecetaCabecera.modificar(oDoRecetaCabecera) Then
1480                   ms_MensajeError = "To: " & oRecetaCabecera.MensajeError: GoTo ErrModif
1490              End If
1500          End If
1510          oDoRecetaDetalle.idReceta = lnRecetaTomografia
1520          If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
1530                 ms_MensajeError = "To: " & oRecetaDetalle.MensajeError: GoTo ErrModif
1540          End If
1550          If oRsTomografia.RecordCount > 0 Then
1560             oRsTomografia.MoveFirst
1570             Do While Not oRsTomografia.EOF
                    oDoRecetaDetalle.dx = IIf(IsNull(oRsTomografia!dx), "", oRsTomografia!dx)
1580                oDoRecetaDetalle.cantidadPedida = oRsTomografia.Fields!cantidad
1590                oDoRecetaDetalle.idItem = oRsTomografia.Fields!Id
1600                oDoRecetaDetalle.Precio = oRsTomografia.Fields!Precio
1610                oDoRecetaDetalle.SaldoEnRegistroReceta = oRsTomografia.Fields!saldoActual
1620                oDoRecetaDetalle.total = Round(oRsTomografia.Fields!cantidad * oRsTomografia.Fields!Precio, 2)
1630                oDoRecetaDetalle.idEstadoDetalle = sghRecetasEstadosDetalle.sghRecetado
1640                oDoRecetaDetalle.idDosisRecetada = 0
1650                oDoRecetaDetalle.IdViaAdministracion = 0 'Actualizado 2609
1660                oDoRecetaDetalle.observaciones = IIf(IsNull(oRsTomografia.Fields!observaciones), "", oRsTomografia.Fields!observaciones)
1670                If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
1680                       ms_MensajeError = "To: " & oRecetaDetalle.MensajeError: GoTo ErrModif
1690                End If
1700                oRsTomografia.MoveNext
1710             Loop
1720          End If
1730      End If
          'Anatomia Patologica
1740      If lnRecetaAnatomiaP <> -100 Then
1750          If lnRecetaAnatomiaP = 0 And oRsAnatomia.RecordCount > 0 Then
1760              oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaAnatomiaPatologica1
1770              If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
1780                   ms_MensajeError = "Ap: " & oRecetaCabecera.MensajeError: GoTo ErrModif
1790              End If
1800              lnRecetaAnatomiaP = oDoRecetaCabecera.idReceta
1810          ElseIf lbModificaCabecera = True Then
1820              oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaAnatomiaPatologica1
1830              oDoRecetaCabecera.idReceta = lnRecetaAnatomiaP
1840              If Not oRecetaCabecera.modificar(oDoRecetaCabecera) Then
1850                   ms_MensajeError = "Ap: " & oRecetaCabecera.MensajeError: GoTo ErrModif
1860              End If
1870          End If
1880          oDoRecetaDetalle.idReceta = lnRecetaAnatomiaP
1890          If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
1900                 ms_MensajeError = "Ap: " & oRecetaDetalle.MensajeError: GoTo ErrModif
1910          End If
1920          If oRsAnatomia.RecordCount > 0 Then
1930             oRsAnatomia.MoveFirst
1940             Do While Not oRsAnatomia.EOF
                    oDoRecetaDetalle.dx = IIf(IsNull(oRsAnatomia!dx), "", oRsAnatomia!dx)
1950                oDoRecetaDetalle.cantidadPedida = oRsAnatomia.Fields!cantidad
1960                oDoRecetaDetalle.idItem = oRsAnatomia.Fields!Id
1970                oDoRecetaDetalle.Precio = oRsAnatomia.Fields!Precio
1980                oDoRecetaDetalle.SaldoEnRegistroReceta = oRsAnatomia.Fields!saldoActual
1990                oDoRecetaDetalle.total = Round(oRsAnatomia.Fields!cantidad * oRsAnatomia.Fields!Precio, 2)
2000                oDoRecetaDetalle.idEstadoDetalle = sghRecetasEstadosDetalle.sghRecetado
2010                oDoRecetaDetalle.idDosisRecetada = 0
2020                oDoRecetaDetalle.IdViaAdministracion = 0 'Actualizado 2609
2030                oDoRecetaDetalle.observaciones = IIf(IsNull(oRsAnatomia.Fields!observaciones), "", oRsAnatomia.Fields!observaciones)
2040                If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
2050                       ms_MensajeError = "Ap: " & oRecetaDetalle.MensajeError: GoTo ErrModif
2060                End If
2070                oRsAnatomia.MoveNext
2080             Loop
2090          End If
2100      End If
          'Patologia Clinica
2110      If lnRecetaPatologiaC <> -100 Then
2120          If lnRecetaPatologiaC = 0 And oRsPatologia.RecordCount > 0 Then
2130              oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaPatologiaClinica
2140              If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
2150                   ms_MensajeError = "Pc: " & oRecetaCabecera.MensajeError: GoTo ErrModif
2160              End If
2170              lnRecetaPatologiaC = oDoRecetaCabecera.idReceta
2180          ElseIf lbModificaCabecera = True Then
2190              oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaPatologiaClinica
2200              oDoRecetaCabecera.idReceta = lnRecetaPatologiaC
2210              If Not oRecetaCabecera.modificar(oDoRecetaCabecera) Then
2220                   ms_MensajeError = "Pc: " & oRecetaCabecera.MensajeError: GoTo ErrModif
2230              End If
2240          End If
2250          oDoRecetaDetalle.idReceta = lnRecetaPatologiaC
2260          If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
2270                 ms_MensajeError = "Pc: " & oRecetaDetalle.MensajeError: GoTo ErrModif
2280          End If
2290          If oRsPatologia.RecordCount > 0 Then
2300             oRsPatologia.MoveFirst
2310             Do While Not oRsPatologia.EOF
                    oDoRecetaDetalle.dx = IIf(IsNull(oRsPatologia!dx), "", oRsPatologia!dx)
2320                oDoRecetaDetalle.cantidadPedida = oRsPatologia.Fields!cantidad
2330                oDoRecetaDetalle.idItem = oRsPatologia.Fields!Id
2340                oDoRecetaDetalle.Precio = oRsPatologia.Fields!Precio
2350                oDoRecetaDetalle.SaldoEnRegistroReceta = oRsPatologia.Fields!saldoActual
2360                oDoRecetaDetalle.total = Round(oRsPatologia.Fields!cantidad * oRsPatologia.Fields!Precio, 2)
2370                oDoRecetaDetalle.idEstadoDetalle = sghRecetasEstadosDetalle.sghRecetado
2380                oDoRecetaDetalle.idDosisRecetada = 0
2390                oDoRecetaDetalle.observaciones = IIf(IsNull(oRsPatologia.Fields!observaciones), "", oRsPatologia.Fields!observaciones)
2400                If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
2410                       ms_MensajeError = "Pc: " & oRecetaDetalle.MensajeError: GoTo ErrModif
2420                End If
2430                oRsPatologia.MoveNext
2440             Loop
2450          End If
2460      End If
          'Banco de Sangre
2470      If lnRecetaBancoS <> -100 Then
2480          If lnRecetaBancoS = 0 And oRsBanco.RecordCount > 0 Then
2490              oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaBancoSangre1
2500              If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
2510                   ms_MensajeError = "Bs: " & oRecetaCabecera.MensajeError: GoTo ErrModif
2520              End If
2530              lnRecetaBancoS = oDoRecetaCabecera.idReceta
2540          ElseIf lbModificaCabecera = True Then
2550              oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaBancoSangre1
2560              oDoRecetaCabecera.idReceta = lnRecetaBancoS
2570              If Not oRecetaCabecera.modificar(oDoRecetaCabecera) Then
2580                   ms_MensajeError = "Bs: " & oRecetaCabecera.MensajeError: GoTo ErrModif
2590              End If
2600          End If
2610          oDoRecetaDetalle.idReceta = lnRecetaBancoS
2620          If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
2630                 ms_MensajeError = "Bs: " & oRecetaDetalle.MensajeError: GoTo ErrModif
2640          End If
2650          If oRsBanco.RecordCount > 0 Then
2660             oRsBanco.MoveFirst
2670             Do While Not oRsBanco.EOF
                    oDoRecetaDetalle.dx = IIf(IsNull(oRsBanco!dx), "", oRsBanco!dx)
2680                oDoRecetaDetalle.cantidadPedida = oRsBanco.Fields!cantidad
2690                oDoRecetaDetalle.idItem = oRsBanco.Fields!Id
2700                oDoRecetaDetalle.Precio = oRsBanco.Fields!Precio
2710                oDoRecetaDetalle.SaldoEnRegistroReceta = oRsBanco.Fields!saldoActual
2720                oDoRecetaDetalle.total = Round(oRsBanco.Fields!cantidad * oRsBanco.Fields!Precio, 2)
2730                oDoRecetaDetalle.idEstadoDetalle = sghRecetasEstadosDetalle.sghRecetado
2740                oDoRecetaDetalle.idDosisRecetada = 0
2750                oDoRecetaDetalle.IdViaAdministracion = 0 'Actualizado 2609
2760                oDoRecetaDetalle.observaciones = IIf(IsNull(oRsBanco.Fields!observaciones), "", oRsBanco.Fields!observaciones)
2770                If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
2780                       ms_MensajeError = "Bs: " & oRecetaDetalle.MensajeError: GoTo ErrModif
2790                End If
2800                oRsBanco.MoveNext
2810             Loop
2820          End If
2830      End If
          'Farmacia
2840      If lnRecetaFarmacia <> -100 Then
              'debb-24/06/2015
2850          If oRsFarmacia.RecordCount > 0 Then
2860              oRsFarmacia.MoveFirst
2870              oDoRecetaCabecera.FechaVigencia = oRsFarmacia!FechaVigencia
2880          End If
              '
2890          If lnRecetaFarmacia = 0 And oRsFarmacia.RecordCount > 0 Then
2900              oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaFarmacia
2910              If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
2920                   ms_MensajeError = "Fa: " & oRecetaCabecera.MensajeError: GoTo ErrModif
2930              End If
2940              lnRecetaFarmacia = oDoRecetaCabecera.idReceta
2950          ElseIf lbModificaCabecera = True Then
2960              oDoRecetaCabecera.IdPuntoCarga = sghPtoCargaFarmacia
2970              oDoRecetaCabecera.idReceta = lnRecetaFarmacia
2980              If Not oRecetaCabecera.modificar(oDoRecetaCabecera) Then
2990                   ms_MensajeError = "Fa: " & oRecetaCabecera.MensajeError: GoTo ErrModif
3000              End If
3010          ElseIf oRsFarmacia.RecordCount > 0 Then                                                'debb-24/06/2015
3020              oDoRecetaCabecera.idReceta = lnRecetaFarmacia
3030              If Not oRecetaCabecera.SeleccionarPorId(oDoRecetaCabecera) Then
3040                   ms_MensajeError = "Fa: " & oRecetaCabecera.MensajeError: GoTo ErrModif
3050              End If
3060              oRsFarmacia.MoveFirst
3070              oDoRecetaCabecera.FechaVigencia = oRsFarmacia!FechaVigencia
3080              If Not oRecetaCabecera.modificar(oDoRecetaCabecera) Then
3090                   ms_MensajeError = "Fa: " & oRecetaCabecera.MensajeError: GoTo ErrModif
3100              End If
3110          End If
3120          oDoRecetaDetalle.idReceta = lnRecetaFarmacia
3130          If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
3140                 ms_MensajeError = "Fa: " & oRecetaDetalle.MensajeError: GoTo ErrModif
3150          End If
3160          If oRsFarmacia.RecordCount > 0 Then
3170             oRsFarmacia.MoveFirst
3180             Do While Not oRsFarmacia.EOF
                    oDoRecetaDetalle.dx = IIf(IsNull(oRsFarmacia!dx), "", oRsFarmacia!dx)
3190                oDoRecetaDetalle.cantidadPedida = oRsFarmacia.Fields!cantidad
3200                oDoRecetaDetalle.idItem = oRsFarmacia.Fields!Id
3210                oDoRecetaDetalle.Precio = oRsFarmacia.Fields!Precio
3220                oDoRecetaDetalle.SaldoEnRegistroReceta = oRsFarmacia.Fields!saldoActual
3230                oDoRecetaDetalle.total = Round(oRsFarmacia.Fields!cantidad * oRsFarmacia.Fields!Precio, 2)
3240                oDoRecetaDetalle.idEstadoDetalle = sghRecetasEstadosDetalle.sghRecetado
3250                oDoRecetaDetalle.idDosisRecetada = oRsFarmacia.Fields!idDosisRecetada
3260                oDoRecetaDetalle.IdViaAdministracion = oRsFarmacia.Fields!IdViaAdministracion 'Actualizado 2609
3270                oDoRecetaDetalle.observaciones = IIf(IsNull(oRsFarmacia.Fields!observaciones), "", oRsFarmacia.Fields!observaciones)
3280                If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
3290                       ms_MensajeError = "Fa: " & oRecetaDetalle.MensajeError: GoTo ErrModif
3300                End If
3310                oRsFarmacia.MoveNext
3320             Loop
3330          End If
3340      End If
          '*********************************************Otros CPT**********************************
          oDoRecetaCabecera.IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion
          If lnRecetaOtrosCpt = -100 Then
                   'ya hubo despacho de una de las recetas, no se puede hacer nada
          Else
                If lnRecetaOtrosCpt = 0 Then
                   'solo agregar cpt
                   If oRsOtrosCpt.RecordCount > 0 Then
                       oRsOtrosCpt.MoveFirst
                       Do While Not oRsOtrosCpt.EOF
                          If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
                               ms_MensajeError = "Otros Cpt: " & oRecetaCabecera.MensajeError: GoTo ErrModif
                          End If
                          oDoRecetaDetalle.IdPuntoCarga = oRsOtrosCpt!IdPuntoCarga
                          oDoRecetaDetalle.idReceta = oDoRecetaCabecera.idReceta
                          oDoRecetaDetalle.dx = IIf(IsNull(oRsOtrosCpt!dx), "", oRsOtrosCpt!dx)
                          oDoRecetaDetalle.cantidadPedida = oRsOtrosCpt.Fields!cantidad
                          oDoRecetaDetalle.idItem = oRsOtrosCpt.Fields!Id
                          oDoRecetaDetalle.Precio = oRsOtrosCpt.Fields!Precio
                          oDoRecetaDetalle.SaldoEnRegistroReceta = 0
                          oDoRecetaDetalle.total = Round(oRsOtrosCpt.Fields!cantidad * oRsOtrosCpt.Fields!Precio, 2)
                          oDoRecetaDetalle.idEstadoDetalle = sghRecetasEstadosDetalle.sghRecetado
                          oDoRecetaDetalle.idDosisRecetada = 0
                          oDoRecetaDetalle.IdViaAdministracion = 0 'Actualizado 2609
                          oDoRecetaDetalle.observaciones = IIf(IsNull(oRsOtrosCpt.Fields!observaciones), "", oRsOtrosCpt.Fields!observaciones)
                          If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
                                 ms_MensajeError = "Otros Cpt: " & oRecetaDetalle.MensajeError: GoTo ErrModif
                          End If
                          oRsOtrosCpt.MoveNext
                       Loop
                       lnRecetaOtrosCpt = oDoRecetaCabecera.idReceta
                   End If
                Else
                   'Modificar
                   Dim oRsTmp1 As New Recordset
                   Set oRsTmp1 = oRecetaCabecera.SeleccionarPorIdCuentaAtencion(oDoRecetaCabecera.idCuentaAtencion)
                   'elimina recetas generadas, que se quitaron de la Grilla
                   If oRsTmp1.RecordCount > 0 Then
                      oRsTmp1.MoveFirst
                      Do While Not oRsTmp1.EOF
                        If oRsTmp1!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion Then
                            oDoRecetaDetalle.idReceta = oRsTmp1!idReceta
                            oRsOtrosCpt.Filter = "Receta=" & oRsTmp1!idReceta
                            If oRsOtrosCpt.RecordCount = 0 Then
                                If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
                                       ms_MensajeError = "Bs: " & oRecetaDetalle.MensajeError: GoTo ErrModif
                                End If
                                oDoRecetaCabecera.idReceta = oRsTmp1!idReceta
                                If Not oRecetaCabecera.Eliminar(oDoRecetaCabecera) Then
                                     ms_MensajeError = "Otros Cpt: " & oRecetaCabecera.MensajeError: GoTo ErrModif
                                End If
                            End If
                        End If
                        oRsTmp1.MoveNext
                      Loop
                   End If
                   'agrega/modificar recetas
                   oRsOtrosCpt.Filter = ""
                   If oRsOtrosCpt.RecordCount > 0 Then
                       oRsOtrosCpt.MoveFirst
                       Do While Not oRsOtrosCpt.EOF
                          If oRsOtrosCpt!receta > 0 Then
                            oDoRecetaCabecera.idReceta = oRsOtrosCpt!receta
                            If Not oRecetaCabecera.SeleccionarPorId(oDoRecetaCabecera) Then
                                 ms_MensajeError = "Otros Cpt: " & oRecetaCabecera.MensajeError: GoTo ErrModif
                            End If
                          Else
                            If Not oRecetaCabecera.Insertar(oDoRecetaCabecera) Then
                                 ms_MensajeError = "Otros Cpt: " & oRecetaCabecera.MensajeError: GoTo ErrModif
                            End If
                          End If
                          oDoRecetaDetalle.IdPuntoCarga = oRsOtrosCpt!IdPuntoCarga
                          oDoRecetaDetalle.idReceta = oDoRecetaCabecera.idReceta
                          oDoRecetaDetalle.dx = IIf(IsNull(oRsOtrosCpt!dx), "", oRsOtrosCpt!dx)
                          oDoRecetaDetalle.cantidadPedida = oRsOtrosCpt.Fields!cantidad
                          oDoRecetaDetalle.idItem = oRsOtrosCpt.Fields!Id
                          oDoRecetaDetalle.Precio = oRsOtrosCpt.Fields!Precio
                          oDoRecetaDetalle.SaldoEnRegistroReceta = 0
                          oDoRecetaDetalle.total = Round(oRsOtrosCpt.Fields!cantidad * oRsOtrosCpt.Fields!Precio, 2)
                          oDoRecetaDetalle.idEstadoDetalle = sghRecetasEstadosDetalle.sghRecetado
                          oDoRecetaDetalle.idDosisRecetada = 0
                          oDoRecetaDetalle.IdViaAdministracion = 0 'Actualizado 2609
                          oDoRecetaDetalle.observaciones = IIf(IsNull(oRsOtrosCpt.Fields!observaciones), "", oRsOtrosCpt.Fields!observaciones)
                          If oRsOtrosCpt!receta > 0 Then
                            If Not oRecetaDetalle.modificar(oDoRecetaDetalle) Then
                                   ms_MensajeError = "Otros Cpt: " & oRecetaDetalle.MensajeError: GoTo ErrModif
                            End If
                          Else
                            If Not oRecetaDetalle.Insertar(oDoRecetaDetalle) Then
                                   ms_MensajeError = "Otros Cpt: " & oRecetaDetalle.MensajeError: GoTo ErrModif
                            End If
                          End If
                          oRsOtrosCpt.MoveNext
                       Loop
                       lnRecetaOtrosCpt = oDoRecetaCabecera.idReceta
                   End If
                   oRsTmp1.Close
                   Set oRsTmp1 = Nothing
                End If
          End If
          '
3350      RecetaModificar = True
ErrModif:
      'Resume
3360      If RecetaModificar = True Then
3370         oConexion.CommitTrans
3380      Else
      'Resume
3390         oConexion.RollbackTrans
3400         ms_MensajeError = ms_MensajeError & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                                        "Function RecetaModificar", "reglasadmision.cls") 'debb-02/05/2016
             
3410      End If
3420      oConexion.Close
3430      Set oConexion = Nothing
3440      Set oRecetaCabecera = Nothing
3450      Set oDoRecetaCabecera = Nothing
3460      Set oRecetaDetalle = Nothing
3470      Set oDoRecetaDetalle = Nothing
End Function

Function RecetaEliminar(lnIdCuentaAtencion As Long, lnIdServicio As Long, lnIdUsuario As Long, _
                       lnRecetaRayosX As Long, lnRecetaEcografiaO As Long, lnRecetaEcografiaG As Long, _
                       lnRecetaTomografia As Long, lnRecetaAnatomiaP As Long, lnRecetaPatologiaC As Long, _
                       lnRecetaBancoS As Long, lnRecetaFarmacia As Long, _
                       oRsRayosX As Recordset, oRsEcografiaO As Recordset, oRsEcografiaG As Recordset, _
                       oRsTomografia As Recordset, oRsAnatomia As Recordset, oRsPatologia As Recordset, _
                       oRsBanco As Recordset, oRsFarmacia As Recordset, _
                       mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcObservacion As String) As Boolean
                       
    On Error GoTo ErrEliminar
    Dim oConexion As New ADODB.Connection
    Dim oRecetaCabecera As New RecetaCabecera, oDoRecetaCabecera As New DoRecetaCabecera
    Dim oRecetaDetalle As New RecetaDetalle, oDoRecetaDetalle As New DoRecetaDetalle
    RecetaEliminar = False
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.BeginTrans
    '
    Set oRecetaCabecera.Conexion = oConexion
    Set oRecetaDetalle.Conexion = oConexion
    '
    ms_MensajeError = IIf(lnRecetaRayosX > 0, "Rx: " + Trim(Str(lnRecetaRayosX)), "")
    ms_MensajeError = ms_MensajeError + IIf(lnRecetaEcografiaO > 0, "Eco.Obs: " + Trim(Str(lnRecetaEcografiaO)), "")
    ms_MensajeError = ms_MensajeError + IIf(lnRecetaEcografiaG > 0, "Eco.Gen: " + Trim(Str(lnRecetaEcografiaG)), "")
    ms_MensajeError = ms_MensajeError + IIf(lnRecetaTomografia > 0, "Tomog: " + Trim(Str(lnRecetaTomografia)), "")
    ms_MensajeError = ms_MensajeError + IIf(lnRecetaAnatomiaP > 0, "Ana.Pat: " + Trim(Str(lnRecetaAnatomiaP)), "")
    ms_MensajeError = ms_MensajeError + IIf(lnRecetaPatologiaC > 0, "Pat.Cli: " + Trim(Str(lnRecetaPatologiaC)), "")
    ms_MensajeError = ms_MensajeError + IIf(lnRecetaBancoS > 0, "Bco.San: " + Trim(Str(lnRecetaBancoS)), "")
    ms_MensajeError = ms_MensajeError + IIf(lnRecetaFarmacia > 0, "Farm: " + Trim(Str(lnRecetaFarmacia)), "")
    Call mo_ReglasSeguridad.AuditoriaAgregarV(lnIdUsuario, "E", lnIdCuentaAtencion, "RecetaCabecera", oConexion, _
                                               mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "(" + ms_MensajeError + ") " + lcObservacion)
    ms_MensajeError = ""
    'Datos comunes para todos
    oDoRecetaDetalle.IdUsuarioAuditoria = lnIdUsuario
    'Rayos X
    If lnRecetaRayosX <> -100 Then
        oDoRecetaDetalle.idReceta = lnRecetaRayosX
        If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
               ms_MensajeError = "Rx: " & oRecetaDetalle.MensajeError: GoTo ErrEliminar
        End If
        oDoRecetaCabecera.idReceta = lnRecetaRayosX
        If Not oRecetaCabecera.Eliminar(oDoRecetaCabecera) Then
               ms_MensajeError = "Rx: " & oRecetaDetalle.MensajeError: GoTo ErrEliminar
        End If
    End If
    'Ecografia Obstetrica
    If lnRecetaEcografiaO <> -100 Then
        oDoRecetaDetalle.idReceta = lnRecetaEcografiaO
        If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
               ms_MensajeError = "Eo: " & oRecetaDetalle.MensajeError: GoTo ErrEliminar
        End If
        oDoRecetaCabecera.idReceta = lnRecetaEcografiaO
        If Not oRecetaCabecera.Eliminar(oDoRecetaCabecera) Then
               ms_MensajeError = "Eo: " & oRecetaDetalle.MensajeError: GoTo ErrEliminar
        End If
    End If
    'Ecografia General
    If lnRecetaEcografiaG <> -100 Then
        oDoRecetaDetalle.idReceta = lnRecetaEcografiaG
        If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
               ms_MensajeError = "Eg: " & oRecetaDetalle.MensajeError: GoTo ErrEliminar
        End If
        oDoRecetaCabecera.idReceta = lnRecetaEcografiaG
        If Not oRecetaCabecera.Eliminar(oDoRecetaCabecera) Then
               ms_MensajeError = "Eg: " & oRecetaDetalle.MensajeError: GoTo ErrEliminar
        End If
    End If
    'Tomografia
    If lnRecetaTomografia <> -100 Then
        oDoRecetaDetalle.idReceta = lnRecetaTomografia
        If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
               ms_MensajeError = "To: " & oRecetaDetalle.MensajeError: GoTo ErrEliminar
        End If
        oDoRecetaCabecera.idReceta = lnRecetaTomografia
        If Not oRecetaCabecera.Eliminar(oDoRecetaCabecera) Then
               ms_MensajeError = "To: " & oRecetaDetalle.MensajeError: GoTo ErrEliminar
        End If
    End If
    'Anatomia Patologica
    If lnRecetaAnatomiaP <> -100 Then
        oDoRecetaDetalle.idReceta = lnRecetaAnatomiaP
        If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
               ms_MensajeError = "Ap: " & oRecetaDetalle.MensajeError: GoTo ErrEliminar
        End If
        oDoRecetaCabecera.idReceta = lnRecetaAnatomiaP
        If Not oRecetaCabecera.Eliminar(oDoRecetaCabecera) Then
               ms_MensajeError = "Ap: " & oRecetaDetalle.MensajeError: GoTo ErrEliminar
        End If
    End If
    'Patologia Clinica
    If lnRecetaPatologiaC <> -100 Then
        oDoRecetaDetalle.idReceta = lnRecetaPatologiaC
        If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
               ms_MensajeError = "Pc: " & oRecetaDetalle.MensajeError: GoTo ErrEliminar
        End If
        oDoRecetaCabecera.idReceta = lnRecetaPatologiaC
        If Not oRecetaCabecera.Eliminar(oDoRecetaCabecera) Then
               ms_MensajeError = "Pc: " & oRecetaDetalle.MensajeError: GoTo ErrEliminar
        End If
    End If
    'Banco de Sangre
    If lnRecetaBancoS <> -100 Then
        oDoRecetaDetalle.idReceta = lnRecetaBancoS
        If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
               ms_MensajeError = "Bs: " & oRecetaDetalle.MensajeError: GoTo ErrEliminar
        End If
        oDoRecetaCabecera.idReceta = lnRecetaBancoS
        If Not oRecetaCabecera.Eliminar(oDoRecetaCabecera) Then
               ms_MensajeError = "Bs: " & oRecetaDetalle.MensajeError: GoTo ErrEliminar
        End If
    End If
    'Farmacia
    If lnRecetaFarmacia <> -100 Then
        oDoRecetaDetalle.idReceta = lnRecetaFarmacia
        If Not oRecetaDetalle.Eliminar(oDoRecetaDetalle) Then
               ms_MensajeError = "Fa: " & oRecetaDetalle.MensajeError: GoTo ErrEliminar
        End If
        oDoRecetaCabecera.idReceta = lnRecetaFarmacia
        If Not oRecetaCabecera.Eliminar(oDoRecetaCabecera) Then
               ms_MensajeError = "Fa: " & oRecetaDetalle.MensajeError: GoTo ErrEliminar
        End If
    End If
    '
    RecetaEliminar = True
ErrEliminar:
'Resume
    If RecetaEliminar = True Then
       oConexion.CommitTrans
    Else
       oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oRecetaCabecera = Nothing
    Set oDoRecetaCabecera = Nothing
    Set oRecetaDetalle = Nothing
    Set oDoRecetaDetalle = Nothing
End Function

Function RecetaFiltrar(oPaciente As DOPaciente, lnNreceta As Long, lnIdCuenta As Long, _
                       lnIdTipoServicio As Long, lcSinApellido As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim lcSql As String
    lcSql = ""
    If lnIdTipoServicio = 3 Then
       lcSql = lcSql & " dbo.Atenciones.idTipoServicio=3 and"
    ElseIf lnIdTipoServicio = 1 Then
       lcSql = lcSql & " dbo.Atenciones.idTipoServicio=1 and"
    Else
       lcSql = lcSql & " (dbo.Atenciones.idTipoServicio=2 or dbo.Atenciones.idTipoServicio=4) and"
    End If
    If oPaciente.ApellidoPaterno <> "" Then
       If lcSinApellido = oPaciente.ApellidoPaterno Then
          lcSql = lcSql & "  dbo.Pacientes.ApellidoPaterno='" & oPaciente.ApellidoPaterno & "' and"
       Else
          lcSql = lcSql & "  dbo.Pacientes.ApellidoPaterno like '%" & oPaciente.ApellidoPaterno & "%' and"
       End If
    End If
    If oPaciente.ApellidoMaterno <> "" Then
       If lcSinApellido = oPaciente.ApellidoPaterno Then
          lcSql = lcSql & "  dbo.Pacientes.ApellidoMaterno='" & oPaciente.ApellidoMaterno & "' and"
       Else
          lcSql = lcSql & "  dbo.Pacientes.ApellidoMaterno like '%" & oPaciente.ApellidoMaterno & "%' and"
       End If
    End If
    If oPaciente.NroHistoriaClinica <> 0 Then
       lcSql = lcSql & "  dbo.Pacientes.NroHistoriaClinica= " & oPaciente.NroHistoriaClinica & " and"
    End If
    If oPaciente.NroDocumento <> "" Then
       lcSql = lcSql & "  dbo.Pacientes.NroDocumento= '" & oPaciente.NroDocumento & "' and"
    End If
    If lnNreceta > 0 Then
       lcSql = lcSql & "  dbo.RecetaCabecera.idReceta= " & lnNreceta & " and"
    End If
    If lnIdCuenta > 0 Then
       lcSql = lcSql & "  dbo.RecetaCabecera.idCuentaAtencion= " & lnIdCuenta & " and"
    End If
    lcSql = Left(lcSql, Len(lcSql) - 4)
    lcSql = lcSql & " Order by dbo.RecetaCabecera.FechaReceta desc,dbo.Pacientes.ApellidoPaterno,dbo.Pacientes.ApellidoMaterno,dbo.Pacientes.PrimerNombre"
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaFiltrar"
        Set oParameter = .CreateParameter("@filtro", adVarChar, adParamInput, 1000, lcSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set RecetaFiltrar = oRecordset
   oConexion.Close
   'Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function



Function TiposDestinoAtencionDevuelveIdSis(lnIdDestinoAtencionGalenHos As Long) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposDestinoAtencionDevuelveIdSis"
        Set oParameter = .CreateParameter("@lnIdDestinoAtencionGalenHos", adInteger, adParamInput, 0, lnIdDestinoAtencionGalenHos): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   TiposDestinoAtencionDevuelveIdSis = "0"
    If oRecordset.RecordCount > 0 Then
           TiposDestinoAtencionDevuelveIdSis = oRecordset.Fields!id_DestinoAseguradoSIS
   End If
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function TiposDocIdentidadDevuelveIdSis(lnIdTipoDocumentoGalenHos As Long) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposDocIdentidadDevuelveIdSis"
        Set oParameter = .CreateParameter("@lnIdTipoDocumentoGalenHos", adInteger, adParamInput, 0, lnIdTipoDocumentoGalenHos): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   TiposDocIdentidadDevuelveIdSis = "2"
    If oRecordset.RecordCount > 0 Then
           TiposDocIdentidadDevuelveIdSis = oRecordset.Fields!CodigoSIS
   End If
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function BuscaAtencionesCptCEparaFormatoHIS(lnIdCuentaAtencion As Long, lnIdPuntoCarga As Long) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "BuscaAtencionesCptCEparaFormatoHIS"
        Set oParameter = .CreateParameter("@lnIdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set BuscaAtencionesCptCEparaFormatoHIS = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



Function ExisteFichaFamiliar(lcFichaFamiliar As String, lnIdPaciente As Long) As String
  Dim oRecordset As New ADODB.Recordset
  On Error GoTo ManejadorDeError
  Set oRecordset = PacientesSeleccionarPorFichaFamiliar(lcFichaFamiliar)
  ExisteFichaFamiliar = ""
  If oRecordset.RecordCount > 0 Then
     oRecordset.MoveFirst
     Do While Not oRecordset.EOF
        If oRecordset!idPaciente <> lnIdPaciente Then
           ExisteFichaFamiliar = Trim(oRecordset!ApellidoPaterno) + " " + Trim(oRecordset!ApellidoMaterno) + ", " + oRecordset!PrimerNombre
           Exit Do
        End If
        oRecordset.MoveNext
     Loop
  End If
  oRecordset.Close
  Set oRecordset = Nothing
  Exit Function
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function


Function AtencionesSeleccionarPorIdFuenteFinanciamiento(lnIdFuenteFinanciamiento As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarPorIdFuenteFinanciamiento"
        Set oParameter = .CreateParameter("@lnIdFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPorIdFuenteFinanciamiento = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function PacientesSeleccionarPorFichaFamiliar(lcFichaFamiliar As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesSeleccionarPorFichaFamiliar"
        Set oParameter = .CreateParameter("@lcFichaFamiliar", adVarChar, adParamInput, 30, lcFichaFamiliar): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PacientesSeleccionarPorFichaFamiliar = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesNacimientosSeleccionarPorFiltro(lcFiltro As String, Optional oConexion1 As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "AtencionesNacimientosSeleccionarPorFiltro"
        Set oParameter = .CreateParameter("@filtro", adVarChar, adParamInput, 500, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set AtencionesNacimientosSeleccionarPorFiltro = oRecordset


      Set oConexion = Nothing

   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function AtencionesCEseleccionarPorFechaServicioMedico(lcFechaIngreso As String, lnIdServicioIngreso As Long, lnIdMedicoIngreso As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = " AtencionesCEseleccionarPorFechaServicioMedico"
        Set oParameter = .CreateParameter("@lnIdMedicoIngreso", adInteger, adParamInput, 0, lnIdMedicoIngreso): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdServicioIngreso", adInteger, adParamInput, 0, lnIdServicioIngreso): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcFechaIngreso", adDBTimeStamp, adParamInput, 0, lcFechaIngreso): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesCEseleccionarPorFechaServicioMedico = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesCEseleccionarPorFechaServicioMedicoCuenta(lcFechaIngreso As String, lnIdServicioIngreso As Long, lnIdMedicoIngreso As Long, lnIdCuentaAtencion As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesCEseleccionarPorFechaServicioMedicoCuenta"
        Set oParameter = .CreateParameter("@lnIdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdMedicoIngreso", adInteger, adParamInput, 0, lnIdMedicoIngreso): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdServicioIngreso", adInteger, adParamInput, 0, lnIdServicioIngreso): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcFechaIngreso", adDBTimeStamp, adParamInput, 0, lcFechaIngreso): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesCEseleccionarPorFechaServicioMedicoCuenta = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function PacientesSeleccionarPorNroHistoria(lnNroHistoriaClinica As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesSeleccionarPorNroHistoria"
        Set oParameter = .CreateParameter("@lnNroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PacientesSeleccionarPorNroHistoria = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
'debb-20/06/2016
Function AtencionesDiagnosticosSeleccionarXidAtencion(ml_idAtencion As Long, Optional oConexion1 As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
    End If
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "AtencionesDiagnosticosSeleccionarXidAtencion"
        Set oParameter = .CreateParameter("@ml_idAtencion", adInteger, adParamInput, 0, ml_idAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set AtencionesDiagnosticosSeleccionarXidAtencion = oRecordset
   If oConexion1 Is Nothing Then
      oConexion.Close
      Set oRecordset = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description

End Function


Function AtencionesSeleccionarPorIdTipoFinanciamiento(lnIdtipoFinanciamiento As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "atencionesSeleccionarXtipoFinanciamiento"
        Set oParameter = .CreateParameter("@lnIdtipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPorIdTipoFinanciamiento = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function PacientesSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PacientesSeleccionarTodos = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionCESeleccionarPorIdAtencion(lnIdAtencion As Long, Optional oConexion1 As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "atencionesCESeleccionarPorId"
        Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set AtencionCESeleccionarPorIdAtencion = oRecordset

        Set oConexion = Nothing

   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function PacientesSeleccionarPorDNI(lcDNI As String) As DOPaciente
    Dim oPaciente As New Pacientes, oDoPaciente As New DOPaciente
    Dim oConexion As New ADODB.Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    Set PacientesSeleccionarPorDNI = Nothing
    Set oPaciente.Conexion = oConexion
    If oPaciente.PacientesSeleccionarPorDNI(lcDNI, oDoPaciente) Then
        Set PacientesSeleccionarPorDNI = oDoPaciente
    End If
    ms_MensajeError = oPaciente.MensajeError
    Set oPaciente = Nothing
    Set oDoPaciente = Nothing
End Function


Function PacientesDatosAdicionalesSeleccionarPorId(idPaciente As Long, oConexion As Connection) As DoPacienteDatosAdd
    Dim oPacientesDatosAdd As New PacientesDatosAdd
    Dim oDoPacienteDatosAdd As New DoPacienteDatosAdd
    Set PacientesDatosAdicionalesSeleccionarPorId = Nothing
    Set oPacientesDatosAdd.Conexion = oConexion
    oDoPacienteDatosAdd.idPaciente = idPaciente
    If oPacientesDatosAdd.SeleccionarPorId(oDoPacienteDatosAdd) Then
        Set PacientesDatosAdicionalesSeleccionarPorId = oDoPacienteDatosAdd
    End If
    ms_MensajeError = oPacientesDatosAdd.MensajeError
    Set oPacientesDatosAdd = Nothing
End Function

Function AtencionesDatosAdicionalesSeleccionarPorIdCuenta(lnIdCuentaAtencion As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDatosAdicionalesSeleccionarPorIdCuenta"
        Set oParameter = .CreateParameter("@lnIdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesDatosAdicionalesSeleccionarPorIdCuenta = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



Function TiposOrigenAtencionDevuelveIdSis(lnIdOrigenAtencionGalenHos As Long) As String
    TiposOrigenAtencionDevuelveIdSis = "0"
    If lnIdOrigenAtencionGalenHos > 0 Then
         Dim oRsTmp As New Recordset
         Dim oCommand As New ADODB.Command
         Dim oParameter As ADODB.Parameter
         Dim oConexion As New ADODB.Connection
         oConexion.CursorLocation = adUseClient
         oConexion.CommandTimeout = 300
         oConexion.Open sighentidades.CadenaConexion
         With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "TiposOrigenAtencionXorigenAtencion"
             Set oParameter = .CreateParameter("@lnIdOrigenAtencionGalenHos", adInteger, adParamInput, 0, lnIdOrigenAtencionGalenHos): .Parameters.Append oParameter
             Set oRsTmp = .Execute
             Set oRsTmp.ActiveConnection = Nothing
        End With
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
        '
        If oRsTmp.RecordCount > 0 Then
           If Not IsNull(oRsTmp.Fields!id_origenAseguradoSIS) Then
              TiposOrigenAtencionDevuelveIdSis = oRsTmp.Fields!id_origenAseguradoSIS
           End If
        End If
        oRsTmp.Close
        Set oRsTmp = Nothing
    End If
End Function

'debb-12/08/2015
Function AtencionesDatosAdicionalesSItieneCodigoPrestacionSIS(lnIdCuentaAtencion As Long, wxParametro302 As String, _
                                                              lnIdFuenteFinanciamiento As Long) As Boolean
    If lnIdCuentaAtencion > 0 And wxParametro302 = "S" And lnIdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim lcMensaje As String
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        AtencionesDatosAdicionalesSItieneCodigoPrestacionSIS = False
         oConexion.CursorLocation = adUseClient
         oConexion.CommandTimeout = 300
         oConexion.Open sighentidades.CadenaConexion
         ms_MensajeError = ""
         With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "AtencionesDatosAdicionalesSItieneCodigoPrestacionSIS"
            Set oParameter = .CreateParameter("@lnIdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
'        oConexion.Close
'        Set oConexion = Nothing
'        Set oCommand = Nothing
        '
        If oRsTmp1.RecordCount > 0 Then
           '****Editado por DBB 30032015
           If oRsTmp1.Fields!idSiaSis = 0 Then
                AtencionesDatosAdicionalesSItieneCodigoPrestacionSIS = True
           End If
           '*****************************
           If Not IsNull(oRsTmp1.Fields!FuaCodigoPrestacion) Then
              If Len(Trim(oRsTmp1.Fields!FuaCodigoPrestacion)) > 0 Then
                 AtencionesDatosAdicionalesSItieneCodigoPrestacionSIS = True
              End If
           End If
           If sghTipoServicio.sghConsultaExterna = oRsTmp1.Fields!idTipoServicio Or _
              sghTipoServicio.sghEmergenciaConsultorios = oRsTmp1.Fields!idTipoServicio Or _
              sghTipoServicio.sghHospitalizacion = oRsTmp1.Fields!idTipoServicio Then
                If IsNull(oRsTmp1.Fields!UsaFua) Or oRsTmp1.Fields!UsaFua = False Then
                   AtencionesDatosAdicionalesSItieneCodigoPrestacionSIS = True
                End If
           End If
        Else
            AtencionesDatosAdicionalesSItieneCodigoPrestacionSIS = True  '****Editado por DBB 30032015
        End If
        If AtencionesDatosAdicionalesSItieneCodigoPrestacionSIS = False Then
           'debb-09/07/2015 (inicio)
           If oRsTmp1.RecordCount > 0 Then
                lcMensaje = ""
                Select Case oRsTmp1.Fields!idTipoServicio
                Case sghEmergenciaConsultorios
                     If oRsTmp1.Fields!UsaGalenHos = False Then
                        lcMensaje = "El Médico de Admisión debe asignarle el 'Código de Prestación' opción SIGH: Admisión de Emergencia"
                     End If
                Case sghHospitalizacion
                     If oRsTmp1.Fields!UsaGalenHos = False Then
                        lcMensaje = "El Médico de Admisión debe asignarle el 'Código de Prestación' opción SIGH: Admisión de Hospitalización"
                     End If
                Case sghConsultaExterna
                     If oRsTmp1.Fields!UsaGalenHos = False Then
                        Dim lbHayProblemas As Boolean
                        lbHayProblemas = True
                        If oRsTmp1!idSiaSis > 0 And IsNull(oRsTmp1.Fields!FuaCodigoPrestacion) Then
                            Dim oRsTmp2 As New Recordset
                            Dim oRsTmp3 As New Recordset
                            Dim oConexionExterna As New Connection
                            Dim oAtencionesDatosAdicionales As New AtencionesDatosAdicionales
                            Dim oDOAtencionDatosAdicionales As New DoAtencionDatosAdicionales
                            oConexionExterna.CommandTimeout = 300
                            oConexionExterna.CursorLocation = adUseClient
                            oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
                            With oCommand
                                 .CommandType = adCmdStoredProc
                                 Set .ActiveConnection = oConexionExterna
                                 .CommandTimeout = 150
                                 .CommandText = "SisFuaAtencionSeleccionarPorId"
                                 Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
                                 Set oRsTmp2 = .Execute
                                 Set oRsTmp2.ActiveConnection = Nothing
                            End With
                            Set oCommand = Nothing
                            Set oParameter = Nothing
                            If oRsTmp2.RecordCount > 0 Then
                               If Not IsNull(oRsTmp2!FuaCodigoPrestacion) Then
                                    With oCommand
                                         .CommandType = adCmdStoredProc
                                         Set .ActiveConnection = oConexion
                                         .CommandTimeout = 150
                                         .CommandText = "atencionesXIdCuentaAtencion"
                                         Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
                                         Set oRsTmp3 = .Execute
                                         Set oRsTmp3.ActiveConnection = Nothing
                                    End With
                                    Set oCommand = Nothing
                                    Set oParameter = Nothing
                                    oDOAtencionDatosAdicionales.idAtencion = oRsTmp3!idAtencion
                                    Set oAtencionesDatosAdicionales.Conexion = oConexion
                                    If oAtencionesDatosAdicionales.SeleccionarPorId(oDOAtencionDatosAdicionales) = True Then
                                       oDOAtencionDatosAdicionales.FuaCodigoPrestacion = oRsTmp2!FuaCodigoPrestacion
                                       If oAtencionesDatosAdicionales.modificar(oDOAtencionDatosAdicionales) = True Then
                                       End If
                                    End If
                                    lbHayProblemas = False
                               End If
                            End If
                            oConexionExterna.Close
                            Set oRsTmp2 = Nothing
                            Set oRsTmp3 = Nothing
                            Set oConexionExterna = Nothing
                            Set oAtencionesDatosAdicionales = Nothing
                            Set oDOAtencionDatosAdicionales = Nothing
                        End If
                        If lbHayProblemas = True Then
                           lcMensaje = "El Médico de Admisión debe asignarle el 'Código de Prestación', opción SIGH: 'Registro de Atenciones'"
                        End If
                     End If
                Case Else
                     'debb-18/08/2015 (inicio)
                     'lcMensaje = "En el Documento debe aparecer 'Código de Prestación' opción SIGH: 'Pacientes Externos con cuenta (seguro)'"
                     'debb-18/08/2015 (fin)
                End Select
                If lcMensaje <> "" Then
                   MsgBox "Esa Cuenta no tiene 'Código de Prestación' (Formato FUA del SIS)" & Chr(13) & lcMensaje, vbInformation, "SIS"
                Else
                   AtencionesDatosAdicionalesSItieneCodigoPrestacionSIS = True
                End If
           End If
           'debb-09/07/2015 (fin)
        End If
        oRsTmp1.Close
        Set oRsTmp1 = Nothing
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
    Else
        AtencionesDatosAdicionalesSItieneCodigoPrestacionSIS = True
    End If
End Function

Function AtencionesPacientesCitasDatosadicionalesSeleccionarPorId(oDoPaciente As DOPaciente, _
                                                   oDoAtencion As DOAtencion, _
                                                   oDOAtencionDatosAdicionales As DoAtencionDatosAdicionales, _
                                                   oConexion As Connection, mo_CuentasAtencion As DOCuentaAtencion, _
                                                   lbConsidearTablaCita As Boolean, Optional oDOCita As doCita) As Boolean
    Dim oAtenciones As New Atenciones
    Set oAtenciones.Conexion = oConexion
    If oAtenciones.AtencionesPacientesCitasDatosadicionalesSeleccionarPorId(oDoAtencion, _
                                                           oDOAtencionDatosAdicionales, oDoPaciente, _
                                                           mo_CuentasAtencion, lbConsidearTablaCita, oDOCita) Then
            AtencionesPacientesCitasDatosadicionalesSeleccionarPorId = True
    End If
    ms_MensajeError = oAtenciones.MensajeError
    Set oAtenciones = Nothing
End Function

Function CitasBloqueadasSeleccionarPorIdCitaBloqueada(lnIdCitaBloqueada As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CitasBloqueadasSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdCitaBloqueada", adInteger, adParamInput, 0, lnIdCitaBloqueada): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CitasBloqueadasSeleccionarPorIdCitaBloqueada = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function PacientesDatosAdicionalesSeleccionarPorIdPaciente(lnIdPaciente As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesDatosAdicionalesSeleccionarPorIdPaciente"
        Set oParameter = .CreateParameter("@lnIdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PacientesDatosAdicionalesSeleccionarPorIdPaciente = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesCEXSeleccionarPorId(lnIdAtencion As Long, oConexionExterna As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionExterna
        .CommandTimeout = 150
        .CommandText = "atencionesCESeleccionarPorId"
        Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesCEXSeleccionarPorId = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description

End Function

'debb-29/04/2016
Function AtencionesSeleccionarPorPaciente(lnIdPaciente As Long, lbMuestraColumnaDx As Boolean) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset, oRsTmp3 As New Recordset, oRsTmp4 As New Recordset
    Dim mo_AdminFacturacion As New ReglasFacturacion
    Dim mo_ReglasServiciosHosp As New ReglasServiciosHosp
    Dim mo_ReglasAdmision As New ReglasAdmision
    Dim oDODiagnostico As New doDiagnostico
    Dim lcDx As String, lcMotivo As String, lcTratamiento As String, lcUPS As String
'    Dim oCommand As New ADODB.Command
'    Dim oParameter As ADODB.Parameter
    Dim oConexionExterna As New Connection
    Dim oConexion As New Connection
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    With oRsTmp3
          .Fields.Append "IdCuentaAtencion", adInteger, 4, adFldIsNullable
          .Fields.Append "IdAtencion", adInteger, 4, adFldIsNullable
          .Fields.Append "FechaIngreso", adDate, 8, adFldIsNullable
          .Fields.Append "HoraIngreso", adVarChar, 5, adFldIsNullable
          .Fields.Append "TipoServicio", adVarChar, 50, adFldIsNullable
          .Fields.Append "Especialidad", adVarChar, 50, adFldIsNullable
          If lbMuestraColumnaDx = True Then
                .Fields.Append "Motivo", adVarChar, 1000, adFldIsNullable
                .Fields.Append "Dx", adVarChar, 100, adFldIsNullable
                .Fields.Append "Tratamiento", adVarChar, 1000, adFldIsNullable
          End If
          .Fields.Append "Medico", adVarChar, 100, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    
    Set oRsTmp1 = mo_ReglasAdmision.AtencionesXidPaciente(lnIdPaciente, oConexion)
'    Set oCommand = Nothing
'    Set oParameter = Nothing
'    With oCommand
'           .CommandType = adCmdStoredProc
'           Set .ActiveConnection = oConexion
'           .CommandTimeout = 150
'           .CommandText = "AtencionesXidPaciente"
'           Set oParameter = .CreateParameter("@lnIdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
'           Set oRsTmp1 = .Execute
'    End With
    '
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          '
          Set oDODiagnostico = mo_AdminFacturacion.DevuelveDxAltaMedica(oRsTmp1.Fields!idAtencion, oRsTmp1.Fields!idTipoServicio)
          If oDODiagnostico.Descripcion <> "" Then
               lcDx = Left(Trim(oDODiagnostico.CodigoCIE2004) & " " & oDODiagnostico.Descripcion, 100)
          Else
               lcDx = Left(mo_AdminFacturacion.DevuelveDxAltaMedicaTodosDx(oRsTmp1.Fields!idAtencion, oRsTmp1.Fields!idTipoServicio, ""), 100)
          End If
          '
          lcMotivo = "": lcTratamiento = ""
          If oRsTmp1.Fields!idTipoServicio = sghTipoServicio.sghConsultaExterna Then
             Set oRsTmp2 = AtencionesCEXSeleccionarPorId(oRsTmp1.Fields!idAtencion, oConexionExterna)
             '
             If oRsTmp2.RecordCount > 0 Then
                lcMotivo = IIf(IsNull(oRsTmp2.Fields!citaMotivo), "", oRsTmp2.Fields!citaMotivo)
                lcTratamiento = IIf(IsNull(oRsTmp2.Fields!citaTratamiento), "", oRsTmp2.Fields!citaTratamiento)
             End If
             oRsTmp2.Close
          End If
          '
          oRsTmp3.AddNew
          oRsTmp3.Fields!idCuentaAtencion = oRsTmp1.Fields!idCuentaAtencion
          oRsTmp3.Fields!idAtencion = oRsTmp1.Fields!idAtencion
          oRsTmp3.Fields!FechaIngreso = oRsTmp1.Fields!FechaIngreso
          oRsTmp3.Fields!HoraIngreso = oRsTmp1.Fields!HoraIngreso
          oRsTmp3.Fields!TipoServicio = IIf(oRsTmp1.Fields!idTipoServicio = sghTipoServicio.sghConsultaExterna, "Cons. Externa", _
                                        IIf(oRsTmp1.Fields!idTipoServicio = sghTipoServicio.sghEmergenciaConsultorios, "Emergencia     ", _
                                        IIf(oRsTmp1.Fields!idTipoServicio = sghTipoServicio.sghHospitalizacion, "Hospitalización", "Externo")))
          oRsTmp3.Fields!Especialidad = oRsTmp1.Fields!Especialidad
          If lbMuestraColumnaDx = True Then
                oRsTmp3.Fields!Motivo = lcMotivo
                oRsTmp3.Fields!dx = lcDx
                oRsTmp3.Fields!Tratamiento = lcTratamiento
          End If
          oRsTmp3.Fields!Medico = oRsTmp1.Fields!Medico
          oRsTmp3.Update
          oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    'Atenciones Historicas desde el HIS
    Set oRsTmp4 = mo_reglasComunes.UPServiciosSeleccionarTodos
    Set oRsTmp1 = mo_reglasComunes.his_historicoAtencionesSeleccionarPorId(lnIdPaciente, oConexion)
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          lcUPS = ""
          If Len(Trim(oRsTmp1.Fields!ups)) > 0 Then
            oRsTmp4.MoveFirst
            oRsTmp4.Find "upsHis='" & Trim(oRsTmp1.Fields!ups) & "'"
            If Not oRsTmp4.EOF Then
               lcUPS = Left(oRsTmp4.Fields!Descripcion, 50)
            End If
          End If
          '
          lcDx = ""
          If Len(Trim(oRsTmp1.Fields!diagnost)) > 0 Then
             Set oRsTmp2 = mo_reglasComunes.DiagnosticosSeleccionarXcodigoCIEsinPto(oRsTmp1.Fields!diagnost)
             If oRsTmp2.RecordCount > 0 Then
                lcDx = Left(Trim(oRsTmp1.Fields!diagnost) & " " & oRsTmp2.Fields!Descripcion, 100)
             End If
             oRsTmp2.Close
          End If
          '
          oRsTmp3.AddNew
          oRsTmp3.Fields!idCuentaAtencion = 0
          oRsTmp3.Fields!idAtencion = 0
          oRsTmp3.Fields!FechaIngreso = oRsTmp1.Fields!fecha
          oRsTmp3.Fields!HoraIngreso = ""
          oRsTmp3.Fields!TipoServicio = "Cons. Externa"
          oRsTmp3.Fields!Especialidad = lcUPS
          If lbMuestraColumnaDx = True Then
                oRsTmp3.Fields!Motivo = ""
                oRsTmp3.Fields!dx = lcDx
                oRsTmp3.Fields!Tratamiento = ""
          End If
          oRsTmp3.Fields!Medico = "(hist.HIS)"
          oRsTmp3.Update
          oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    
    
    
    If oRsTmp3.RecordCount > 0 Then
       oRsTmp3.MoveFirst
    End If
    oConexionExterna.Close
    Set oConexionExterna = Nothing
    Set AtencionesSeleccionarPorPaciente = oRsTmp3
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set mo_AdminFacturacion = Nothing
    Set oDODiagnostico = Nothing
    Set mo_ReglasServiciosHosp = Nothing
    Set mo_ReglasAdmision = Nothing
End Function

'***debb-14/05/2015
Function ServiciosIntermediosSeleccionarPorPaciente(lnIdPaciente As Long, _
                                                    lbMuestraReceta As Boolean, _
                                                    lbMuestraConsumoEnServicio As Boolean, _
                                                    lnIdCuenta As Long, lbMuestraRecetaYusuario As Boolean, _
                                                    lbSoloConResultados As Boolean) As Recordset
    
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset, oRsTmp3 As New Recordset, oRsTmp4 As New Recordset
    Dim oRsUsuarioDespacho As New Recordset
    Dim mo_ReglasDeSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Dim mo_AdminFacturacion As New ReglasFacturacion
    Dim mo_ReglasLaboratorio As New ReglasLaboratorio
    Dim mo_reglasImagen As New SIGHNegocios.ReglasImagenes
    Dim rsResultados As New Recordset
    Dim mo_ReglasCaja As New ReglasCaja
    Dim lcDx As String, lcMotivo As String, lcTratamiento As String, lnCantidadDev As Long
    Dim lcFecha As String, lcHora As String, lcResultado As String, lcEspecialista As String, lcServicioApDx As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection, lcUsuarioDespacho As String
    Dim lbAdiciona As Boolean, lcDocDespacho As String, lbAgregaFila As Boolean
    Dim lnIdReceta1 As Long, ldFechaReceta1 As Date, lcDocDespacho1 As String, lnCantidadPedida1 As Long
    Const lcLaboratorio As String = "Laboratorio", lcImagenologia As String = "Imagenología"
    Const lcConsumoServicio As String = "Cons.en Servicio", lcFarmaciaVta As String = "Farmacia-Vtas"
    Const lcFarmaciaIS As String = "Farmacia-IS"
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    

    With oRsTmp3
          If lbMuestraReceta = True Then
             .Fields.Append "Receta", adInteger, 4, adFldIsNullable
             .Fields.Append "FReceta", adDate, 8, adFldIsNullable
             .Fields.Append "CantPedida", adInteger, 4, adFldIsNullable
             .Fields.Append "DocDespacho", adVarChar, 20, adFldIsNullable
          End If
          .Fields.Append "IdCuentaAtencion", adInteger, 4, adFldIsNullable
          .Fields.Append "idOrden", adInteger, 4, adFldIsNullable
          .Fields.Append "IdProducto", adInteger, 4, adFldIsNullable
          .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "Fecha", adDate, 8, adFldIsNullable
          .Fields.Append "Hora", adVarChar, 5, adFldIsNullable
          .Fields.Append "ServicioApDx", adVarChar, 50, adFldIsNullable
          .Fields.Append "Item", adVarChar, 300, adFldIsNullable                          'debb-22/12/2016
          .Fields.Append "Cantidad", adInteger, 4, adFldIsNullable
          .Fields.Append "Resultado", adVarChar, 3000, adFldIsNullable
          .Fields.Append "Especialista", adVarChar, 100, adFldIsNullable
          If lbMuestraRecetaYusuario = True Then
             .Fields.Append "Receta", adInteger, 4, adFldIsNullable
             .Fields.Append "UsuarioDespacho", adVarChar, 30, adFldIsNullable
          End If
          .Fields.Append "Resultado1", adVarChar, 3000, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    If lbMuestraReceta = True Then
        'Cpt- se añade Recetas sin despacho
        Set oCommand = Nothing
        Set oParameter = Nothing
        With oCommand
               .CommandType = adCmdStoredProc
               Set .ActiveConnection = oConexion
               .CommandTimeout = 150
               .CommandText = "RecetasCPTconSinDespachoXcuenta"
               Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuenta): .Parameters.Append oParameter
               Set oRsTmp4 = .Execute
        End With
        oRsTmp4.Filter = "docDespacho=null"
        If oRsTmp4.RecordCount > 0 Then
           oRsTmp4.MoveFirst
           Do While Not oRsTmp4.EOF
                Select Case oRsTmp4.Fields!IdPuntoCarga
                Case sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2, _
                     sghPuntosCargaBasicos.sghPtoCargaBancoSangre1, sghPuntosCargaBasicos.sghPtoCargaBancoSangre2, _
                     sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica
                     lcServicioApDx = lcLaboratorio
                Case sghPuntosCargaBasicos.sghPtoCargaEcogGeneral, sghPuntosCargaBasicos.sghPtoCargaEcogObstetrica, _
                     sghPuntosCargaBasicos.sghPtoCargaRayosX, sghPuntosCargaBasicos.sghPtoCargaTomografia
                     lcServicioApDx = lcImagenologia
                Case Else
                     lcServicioApDx = lcConsumoServicio
                End Select
                oRsTmp3.AddNew
                oRsTmp3.Fields!idCuentaAtencion = lnIdCuenta
                oRsTmp3.Fields!IdOrden = 0
                oRsTmp3.Fields!idProducto = oRsTmp4.Fields!idItem
                oRsTmp3.Fields!Codigo = oRsTmp4.Fields!Codigo
                'oRsTmp3.Fields!fecha = CDate(lcFecha)
                'oRsTmp3.Fields!Hora = lcHora
                oRsTmp3.Fields!ServicioApDx = lcServicioApDx
                oRsTmp3.Fields!Item = oRsTmp4.Fields!nombre
                'oRsTmp3.Fields!cantidad = oRsTmp1.Fields!cantidad
                'oRsTmp3.Fields!Resultado = lcResultado
                'oRsTmp3.Fields!Especialista = lcEspecialista
                oRsTmp3.Fields!receta = oRsTmp4.Fields!idReceta
                oRsTmp3.Fields!FReceta = oRsTmp4.Fields!fechaReceta
                oRsTmp3.Fields!docDespacho = oRsTmp4.Fields!docDespacho
                oRsTmp3.Fields!CantPedida = oRsTmp4.Fields!cantidadPedida
                oRsTmp3.Update
                oRsTmp4.MoveNext
           Loop
        End If
        oRsTmp4.Filter = "docDespacho<>null"
    End If
    
    'Cpt
    Set oCommand = Nothing
    Set oParameter = Nothing
    With oCommand
           .CommandType = adCmdStoredProc
           Set .ActiveConnection = oConexion
           .CommandTimeout = 150
           .CommandText = "ProcedimientosXpaciente"
           Set oParameter = .CreateParameter("@lnIdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
           Set oRsTmp1 = .Execute
    End With
    '
    If lnIdCuenta > 0 Then
       oRsTmp1.Filter = "idCuentaAtencion=" & lnIdCuenta
    End If
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          If lbMuestraRecetaYusuario = True And Not IsNull(oRsTmp1!idCuentaAtencion) Then
                Set oCommand = Nothing
                Set oParameter = Nothing
                With oCommand
                       .CommandType = adCmdStoredProc
                       Set .ActiveConnection = oConexion
                       .CommandTimeout = 150
                       .CommandText = "RecetasCPTconSinDespachoXcuenta"
                       Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsTmp1!idCuentaAtencion): .Parameters.Append oParameter
                       Set oRsTmp4 = .Execute
                End With
                oRsTmp4.Filter = "docDespacho<>null"
                '
                lcUsuarioDespacho = ""
                If Not IsNull(oRsTmp1!UsuarioDespacho) Then
                   lcUsuarioDespacho = Left(oRsTmp1!UsuarioDespacho, 30)
                End If
          End If
          '
          lbAdiciona = True: lcDocDespacho = ""
          lcFecha = "": lcHora = "": lcResultado = "": lcEspecialista = ""
          Select Case oRsTmp1.Fields!IdPuntoCarga
          Case sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2, _
               sghPuntosCargaBasicos.sghPtoCargaBancoSangre1, sghPuntosCargaBasicos.sghPtoCargaBancoSangre2, _
               sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica
               lcServicioApDx = lcLaboratorio
               
               mo_ReglasLaboratorio.ResultadosAutomaticosActualizaHaciaGalenhos oRsTmp1!IdOrden
               
               If mo_ReglasLaboratorio.PruebaTieneResultado(oRsTmp1!Codigo, oRsTmp1!IdOrden) = True Then
                   lcResultado = "SI"
               Else
                   lcResultado = "NO"
               End If

               Set oCommand = Nothing
               Set oParameter = Nothing
               With oCommand
                      .CommandType = adCmdStoredProc
                      Set .ActiveConnection = oConexion
                      .CommandTimeout = 150
                      .CommandText = "LaboratorioXidOrden"
                      Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, oRsTmp1.Fields!IdOrden): .Parameters.Append oParameter
                      Set oRsTmp2 = .Execute
               End With
               '
               If oRsTmp2.RecordCount > 0 Then
                    lcFecha = Format(oRsTmp2.Fields!fecha, "dd/mm/yyyy")
                    lcHora = Format(oRsTmp2.Fields!fecha, "hh:mm")
                    lcEspecialista = oRsTmp2.Fields!especialista
                    lcDocDespacho = Trim(Str(oRsTmp2.Fields!IdMovimiento))
               End If
               oRsTmp2.Close
          Case sghPuntosCargaBasicos.sghPtoCargaEcogGeneral, sghPuntosCargaBasicos.sghPtoCargaEcogObstetrica, _
               sghPuntosCargaBasicos.sghPtoCargaRayosX, sghPuntosCargaBasicos.sghPtoCargaTomografia
               lcServicioApDx = lcImagenologia

               

               Set oCommand = Nothing
               Set oParameter = Nothing
               With oCommand
                      .CommandType = adCmdStoredProc
                      Set .ActiveConnection = oConexion
                      .CommandTimeout = 150
                      .CommandText = "ImagenesXidOrden"
                      Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, oRsTmp1.Fields!IdOrden): .Parameters.Append oParameter
                      Set oRsTmp2 = .Execute
               End With
               '
               If oRsTmp2.RecordCount > 0 Then
                    lcFecha = Format(oRsTmp2.Fields!fecha, "dd/mm/yyyy")
                    lcHora = Format(oRsTmp2.Fields!fecha, "hh:mm")
                    lcResultado = IIf(IsNull(oRsTmp2.Fields!resultadoFinal), "", oRsTmp2.Fields!resultadoFinal)
                    lcEspecialista = oRsTmp2.Fields!especialista
                    lcDocDespacho = Trim(Str(oRsTmp2.Fields!IdMovimiento))
               End If
               oRsTmp2.Close
               'x item
               Set oRsTmp2 = mo_reglasComunes.ImagMovimientoSeleccionarIdOrden(oRsTmp1!IdOrden)
               If oRsTmp2.RecordCount > 0 Then
                    Set rsResultados = mo_reglasImagen.ImagMovimientoResultadosSeleccionarPorId(oRsTmp2!IdMovimiento)
                    rsResultados.Filter = "idProductoCpt=" & oRsTmp1!idProducto
                    If rsResultados.RecordCount > 0 Then
                        lcResultado = "SI"
                        lcFecha = Format(rsResultados!resultadoFecha, "dd/mm/yyyy")
                        lcHora = Format(rsResultados!resultadoFecha, "hh:mm")
                        lcEspecialista = rsResultados!resultadodni
                        lcDocDespacho = Trim(Str(rsResultados!IdMovimiento))
                    End If
               End If
               oRsTmp2.Close
               
          Case Else
               If lbMuestraConsumoEnServicio = False Then
                  lbAdiciona = False
               End If
               lcServicioApDx = lcConsumoServicio
               lcFecha = Format(oRsTmp1.Fields!fechaDespacho, "dd/mm/yyyy")
               lcHora = Format(oRsTmp1.Fields!fechaDespacho, "hh:mm")
               lcResultado = ""
               lcEspecialista = ""
          End Select
          '
          If lbAdiciona = True Then
                lnIdReceta1 = 0: ldFechaReceta1 = 0: lcDocDespacho1 = "": lnCantidadPedida1 = 0
                If lbMuestraReceta = True And lcDocDespacho <> "" Then
                    If oRsTmp4.RecordCount > 0 Then
                       oRsTmp4.MoveFirst
                       Do While Not oRsTmp4.EOF
                          If oRsTmp4.Fields!IdPuntoCarga = oRsTmp1.Fields!IdPuntoCarga And Trim(oRsTmp4.Fields!docDespacho) = lcDocDespacho Then
                                lnIdReceta1 = oRsTmp4.Fields!idReceta
                                ldFechaReceta1 = oRsTmp4.Fields!fechaReceta
                                lcDocDespacho1 = oRsTmp4.Fields!docDespacho
                                lnCantidadPedida1 = oRsTmp4.Fields!cantidadPedida
                                Exit Do
                          End If
                          oRsTmp4.MoveNext
                       Loop
                    End If
                ElseIf lbMuestraRecetaYusuario = True And lcDocDespacho <> "" Then
                    If oRsTmp4.RecordCount > 0 Then
                       oRsTmp4.MoveFirst
                       Do While Not oRsTmp4.EOF
                          If oRsTmp4.Fields!IdPuntoCarga = oRsTmp1.Fields!IdPuntoCarga And Trim(oRsTmp4.Fields!docDespacho) = lcDocDespacho Then
                                lnIdReceta1 = oRsTmp4.Fields!idReceta
                                ldFechaReceta1 = oRsTmp4.Fields!fechaReceta
                                lcDocDespacho1 = oRsTmp4.Fields!docDespacho
                                lnCantidadPedida1 = oRsTmp4.Fields!cantidadPedida
                                Exit Do
                          End If
                          oRsTmp4.MoveNext
                       Loop
                    End If
                End If
                lbAgregaFila = True
                If lbSoloConResultados = True Then
                   If lcResultado = "NO" Or lcResultado = "" Then
                      lbAgregaFila = False
                   End If
                End If
                If lbAgregaFila Then
                    oRsTmp3.AddNew
                    oRsTmp3.Fields!idCuentaAtencion = oRsTmp1.Fields!idCuentaAtencion
                    oRsTmp3.Fields!IdOrden = oRsTmp1.Fields!IdOrden
                    oRsTmp3.Fields!idProducto = oRsTmp1.Fields!idProducto
                    oRsTmp3.Fields!Codigo = oRsTmp1.Fields!Codigo
                    If lcFecha <> "" Then
                       oRsTmp3.Fields!fecha = CDate(lcFecha)
                    End If
                    oRsTmp3.Fields!Hora = lcHora
                    oRsTmp3.Fields!ServicioApDx = lcServicioApDx
                    oRsTmp3.Fields!Item = oRsTmp1.Fields!Item
                    oRsTmp3.Fields!cantidad = oRsTmp1.Fields!cantidad
                    If lcResultado = "SI" Or lcResultado = "NO" Then
                       oRsTmp3.Fields!Resultado = lcResultado
                    ElseIf Len(lcResultado) > 0 Then
                       oRsTmp3.Fields!Resultado = "SI"
                    End If
                    oRsTmp3.Fields!Resultado1 = lcResultado
                    oRsTmp3.Fields!especialista = lcEspecialista
                    If lbMuestraReceta = True And lcDocDespacho <> "" And lnIdReceta1 > 0 Then
                        oRsTmp3.Fields!receta = lnIdReceta1
                        oRsTmp3.Fields!FReceta = ldFechaReceta1
                        oRsTmp3.Fields!docDespacho = lcDocDespacho1
                        oRsTmp3.Fields!CantPedida = lnCantidadPedida1
                    ElseIf lbMuestraRecetaYusuario = True Then
                        oRsTmp3.Fields!receta = lnIdReceta1
                        oRsTmp3.Fields!UsuarioDespacho = lcUsuarioDespacho
                    End If
                    oRsTmp3.Update
                End If
          End If
          oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    
    If lbSoloConResultados = True Then
        oRsTmp3.Sort = "fecha desc,hora desc"
        Set ServiciosIntermediosSeleccionarPorPaciente = oRsTmp3
        Set oRsTmp1 = Nothing
        Set oRsTmp2 = Nothing
        Set mo_AdminFacturacion = Nothing
        Set mo_ReglasLaboratorio = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set mo_ReglasCaja = Nothing
        Set oRsTmp4 = Nothing
        Set oRsUsuarioDespacho = Nothing
        Set mo_ReglasDeSeguridad = Nothing
        Set mo_reglasImagen = Nothing
        Set rsResultados = Nothing
        Exit Function
    End If
    
    '******************************* farmacia *********************
    lcServicioApDx = lcFarmaciaVta
    If lbMuestraReceta = True Then
        'Farmacia- se añade Recetas sin despacho
        Set oCommand = Nothing
        Set oParameter = Nothing
        With oCommand
               .CommandType = adCmdStoredProc
               Set .ActiveConnection = oConexion
               .CommandTimeout = 150
               .CommandText = "RecetasFarmaciaConSinDespachoXcuenta"
               Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuenta): .Parameters.Append oParameter
               Set oRsTmp4 = .Execute
        End With
        oRsTmp4.Filter = "docDespacho=null"
        If oRsTmp4.RecordCount > 0 Then
           oRsTmp4.MoveFirst
           Do While Not oRsTmp4.EOF
                oRsTmp3.AddNew
                oRsTmp3.Fields!idCuentaAtencion = lnIdCuenta
                oRsTmp3.Fields!IdOrden = 0
                oRsTmp3.Fields!idProducto = oRsTmp4.Fields!idItem
                oRsTmp3.Fields!Codigo = oRsTmp4.Fields!Codigo
                'oRsTmp3.Fields!fecha = CDate(lcFecha)
                'oRsTmp3.Fields!Hora = lcHora
                oRsTmp3.Fields!ServicioApDx = lcServicioApDx
                oRsTmp3.Fields!Item = oRsTmp4.Fields!nombre
                'oRsTmp3.Fields!cantidad = oRsTmp1.Fields!cantidad
                'oRsTmp3.Fields!Resultado = lcResultado
                'oRsTmp3.Fields!Especialista = lcEspecialista
                oRsTmp3.Fields!receta = oRsTmp4.Fields!idReceta
                oRsTmp3.Fields!FReceta = oRsTmp4.Fields!fechaReceta
                oRsTmp3.Fields!docDespacho = oRsTmp4.Fields!docDespacho
                oRsTmp3.Fields!CantPedida = oRsTmp4.Fields!cantidadPedida
                oRsTmp3.Update
                oRsTmp4.MoveNext
           Loop
        End If
        oRsTmp4.Filter = "docDespacho<>null"
    End If
    
    'Farmacia - despachos Ventas
    Set oCommand = Nothing
    Set oParameter = Nothing
    With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "FarmVentasXpaciente"
          Set oParameter = .CreateParameter("@lnIdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
          Set oRsTmp1 = .Execute
    End With
    '
    If lnIdCuenta > 0 Then
       oRsTmp1.Filter = "idCuentaAtencion=" & lnIdCuenta
    End If
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
If oRsTmp1!idCuentaAtencion = 206697 Then
lnCantidadDev = 0
End If
           If lbMuestraRecetaYusuario = True And Not IsNull(oRsTmp1!idCuentaAtencion) Then
                Set oCommand = Nothing
                Set oParameter = Nothing
                With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = oConexion
                    .CommandTimeout = 150
                    .CommandText = "RecetasFarmaciaConSinDespachoXcuenta"
                    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsTmp1!idCuentaAtencion): .Parameters.Append oParameter
                    Set oRsTmp4 = .Execute
                End With
                oRsTmp4.Filter = "docDespacho<>null"
          End If
          'Devoluciones
          Set oCommand = Nothing
          Set oParameter = Nothing
          With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "DevolucionesMedicamentosXcuentaIdproductoMovNumero"
                Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsTmp1.Fields!idCuentaAtencion): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, oRsTmp1.Fields!idProducto): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, oRsTmp1.Fields!movNumero): .Parameters.Append oParameter
                Set oRsTmp2 = .Execute
          End With
          '
          lnCantidadDev = 0
          If oRsTmp2.RecordCount > 0 Then
            Do While Not oRsTmp2.EOF
               lnCantidadDev = lnCantidadDev + oRsTmp2.Fields!cantidadAdevolver
               oRsTmp2.MoveNext
            Loop
          End If
          oRsTmp2.Close
          lcFecha = Format(oRsTmp1.Fields!FechaCreacion, "dd/mm/yyyy")
          lcHora = Format(oRsTmp1.Fields!FechaCreacion, "hh:mm")
          lcDocDespacho = oRsTmp1.Fields!DocumentoNumero
          If lnCantidadDev < oRsTmp1.Fields!cantidad Then
                lnIdReceta1 = 0: ldFechaReceta1 = 0: lcDocDespacho1 = "": lnCantidadPedida1 = 0
                If lbMuestraReceta = True And lcDocDespacho <> "" Then
                     If oRsTmp4.RecordCount > 0 Then
                       oRsTmp4.MoveFirst
                       Do While Not oRsTmp4.EOF
                          If Trim(oRsTmp4.Fields!docDespacho) = lcDocDespacho And oRsTmp4.Fields!idItem = oRsTmp1.Fields!idProducto Then
                                lnIdReceta1 = oRsTmp4.Fields!idReceta
                                ldFechaReceta1 = oRsTmp4.Fields!fechaReceta
                                lcDocDespacho1 = oRsTmp4.Fields!docDespacho
                                lnCantidadPedida1 = oRsTmp4.Fields!cantidadPedida
                                Exit Do
                          End If
                          oRsTmp4.MoveNext
                       Loop
                    End If
                ElseIf lbMuestraRecetaYusuario = True And lcDocDespacho <> "" Then
                     If oRsTmp4.RecordCount > 0 Then
                       oRsTmp4.MoveFirst
                       Do While Not oRsTmp4.EOF
                          If Trim(oRsTmp4.Fields!docDespacho) = lcDocDespacho And oRsTmp4.Fields!idItem = oRsTmp1.Fields!idProducto Then
                                lnIdReceta1 = oRsTmp4.Fields!idReceta
                                ldFechaReceta1 = oRsTmp4.Fields!fechaReceta
                                lcDocDespacho1 = oRsTmp4.Fields!docDespacho
                                lnCantidadPedida1 = oRsTmp4.Fields!cantidadPedida
                                Exit Do
                          End If
                          oRsTmp4.MoveNext
                       Loop
                    End If
                End If
                '
                If lbMuestraRecetaYusuario = True Then
                   lcUsuarioDespacho = ""
                   If Not IsNull(oRsTmp1!especialista) Then
                      lcUsuarioDespacho = Left(oRsTmp1!especialista, 30)
                   End If
                End If
                '
                oRsTmp3.AddNew
                oRsTmp3.Fields!idCuentaAtencion = oRsTmp1.Fields!idCuentaAtencion
                oRsTmp3.Fields!fecha = CDate(lcFecha)
                oRsTmp3.Fields!Hora = lcHora
                oRsTmp3.Fields!ServicioApDx = lcServicioApDx
                oRsTmp3.Fields!Item = oRsTmp1.Fields!Item
                oRsTmp3.Fields!cantidad = oRsTmp1.Fields!cantidad - lnCantidadDev
                oRsTmp3.Fields!Resultado = ""
                oRsTmp3.Fields!especialista = oRsTmp1.Fields!especialista
                If lbMuestraReceta = True And lcDocDespacho <> "" And lnIdReceta1 > 0 Then
                    oRsTmp3.Fields!receta = lnIdReceta1
                    oRsTmp3.Fields!FReceta = ldFechaReceta1
                    oRsTmp3.Fields!docDespacho = lcDocDespacho1
                    oRsTmp3.Fields!CantPedida = lnCantidadPedida1
                ElseIf lbMuestraRecetaYusuario = True Then
                    oRsTmp3.Fields!receta = lnIdReceta1
                    oRsTmp3.Fields!UsuarioDespacho = lcUsuarioDespacho
                End If
                oRsTmp3.Update
          End If
          oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    'Farmacia-Intervenciones Sanitarias
    If lnIdCuenta = 0 Then
        lcServicioApDx = lcFarmaciaIS
    
        Set oCommand = Nothing
        Set oParameter = Nothing
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "farmMovimientoProgramasXpaciente"
            Set oParameter = .CreateParameter("@lnIdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
        End With
        '
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              lcFecha = Format(oRsTmp1.Fields!FechaCreacion, "dd/mm/yyyy")
              lcHora = Format(oRsTmp1.Fields!FechaCreacion, "hh:mm")
              oRsTmp3.AddNew
              oRsTmp3.Fields!idCuentaAtencion = 0
              oRsTmp3.Fields!fecha = CDate(lcFecha)
              oRsTmp3.Fields!Hora = lcHora
              oRsTmp3.Fields!ServicioApDx = lcServicioApDx
              oRsTmp3.Fields!Item = oRsTmp1.Fields!Item
              oRsTmp3.Fields!cantidad = oRsTmp1.Fields!cantidad
              oRsTmp3.Fields!Resultado = oRsTmp1.Fields!Descripcion
              oRsTmp3.Fields!especialista = oRsTmp1.Fields!especialista
              oRsTmp3.Update
              oRsTmp1.MoveNext
           Loop
        End If
        oRsTmp1.Close
        
        'Atenciones Historicas desde el HIS
        Set oRsTmp1 = mo_reglasComunes.his_historicoAtencionesSeleccionarPorId(lnIdPaciente, oConexion)
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              lcServicioApDx = ""
              If Len(Trim(oRsTmp1.Fields!cpt)) > 0 Then
                Set oRsTmp2 = mo_ReglasCaja.FactCatalogoServiciosSeleccionarPorCodigoOnombre(Trim(oRsTmp1.Fields!cpt), "", oConexion)
                If oRsTmp2.RecordCount > 0 Then
                   lcServicioApDx = Left(oRsTmp2.Fields!nombre, 250)
                End If
                oRsTmp2.Close
              End If
              '
              If lcServicioApDx <> "" Then
                oRsTmp3.AddNew
                oRsTmp3.Fields!idCuentaAtencion = 0
                oRsTmp3.Fields!fecha = oRsTmp1.Fields!fecha
                oRsTmp3.Fields!Hora = ""
                oRsTmp3.Fields!ServicioApDx = ""
                oRsTmp3.Fields!Item = lcServicioApDx
                oRsTmp3.Fields!cantidad = 1
                oRsTmp3.Fields!Resultado = ""
                oRsTmp3.Fields!especialista = "(hist.HIS)"
                oRsTmp3.Update
              End If
              oRsTmp1.MoveNext
           Loop
        End If
        oRsTmp1.Close
    End If
    '
    If oRsTmp3.RecordCount > 0 Then
       If lbMuestraReceta = True Then
          oRsTmp3.Sort = "freceta desc,fecha desc,hora desc"
       Else
          oRsTmp3.Sort = "fecha desc,hora desc"
       End If
       oRsTmp3.MoveFirst
    End If
    oConexion.Close
    Set ServiciosIntermediosSeleccionarPorPaciente = oRsTmp3
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set mo_AdminFacturacion = Nothing
    Set mo_ReglasLaboratorio = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set mo_ReglasCaja = Nothing
    Set oRsTmp4 = Nothing
    Set oRsUsuarioDespacho = Nothing
    Set mo_ReglasDeSeguridad = Nothing
    Set mo_reglasImagen = Nothing
    Set rsResultados = Nothing
    
    
    
End Function

Function CentrosPobladosSeleccionarTodos() As Recordset
   On Error GoTo ErrCPSt
   Dim oCentroPoblados As New CentroPoblados
   Dim oConexion As New Connection
   oConexion.CommandTimeout = 300
   oConexion.CursorLocation = adUseClient
   oConexion.Open sighentidades.CadenaConexion
   Set oCentroPoblados.Conexion = oConexion
   Set CentrosPobladosSeleccionarTodos = oCentroPoblados.SeleccionarTodos
   Exit Function
ErrCPSt:
   MsgBox Err.Description
End Function

Function PacientesSeleccionarPorIdentificador(lnIdPaciente As Long, Optional oConexion1 As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "PacientesSeleccionarPorIdentificador"
        Set oParameter = .CreateParameter("@lnIdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set PacientesSeleccionarPorIdentificador = oRecordset
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function BuscaAtencionesCE(lcFechaAtencionInicial As String, lcFechaAtencionFinal As String, lnIdResponsable As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "BuscaAtencionesCE"
        Set oParameter = .CreateParameter("@lcFechaAtencionInicial", adDBTimeStamp, adParamInput, 0, lcFechaAtencionInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcFechaAtencionFinal", adDBTimeStamp, adParamInput, 0, lcFechaAtencionFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount > 0 And lnIdResponsable > 0 Then
      oRecordset.Filter = "IdMedicoIngreso=" & lnIdResponsable
   End If
   Set BuscaAtencionesCE = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesNacimientosSeleccionarXidPaciente(lnIdPaciente As Long, oConexion As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesNacimientosSeleccionarXidPaciente"
        Set oParameter = .CreateParameter("@lnIdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesNacimientosSeleccionarXidPaciente = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ExisteAlgunPaciente(ap As String, am As String, PN As String, SN As String, TN As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
Dim C1 As String, C2 As String
Dim lcFiltro As String
    lcFiltro = ""
    If SN <> "" Then
      lcFiltro = lcFiltro & " AND segundoNombre='" & SN & "' "
    End If
    If TN <> "" Then
      lcFiltro = lcFiltro & " AND tercernombre='" & TN & "' "
    End If
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesFiltraPorApellidosYnombre"
        Set oParameter = .CreateParameter("@ap", adVarChar, adParamInput, 40, ap): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@am", adVarChar, adParamInput, 40, am): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@PN", adVarChar, adParamInput, 40, PN): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If lcFiltro <> "" Then
      lcFiltro = Mid(lcFiltro, 5)
      oRecordset.Filter = lcFiltro
   End If
   Set ExisteAlgunPaciente = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesCeXnrohistoria(ml_NroHistoriaClinica As Long, oConexion As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesCeXnrohistoria"
        Set oParameter = .CreateParameter("@ml_NroHistoriaClinica", adInteger, adParamInput, 0, ml_NroHistoriaClinica): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesCeXnrohistoria = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesFiltraDatosParaFormatoSIS(lnIdAtencion As Long, oConexion As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesFiltraDatosParaFormatoSIS"
        Set oParameter = .CreateParameter("@lnIdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesFiltraDatosParaFormatoSIS = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function PacientesDatosParaImpresionFiliacion(lnIdPaciente As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesDatosParaImpresionFiliacion"
        Set oParameter = .CreateParameter("@lnIdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set PacientesDatosParaImpresionFiliacion = oRecordset
   'Set oRecordSet = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesXidAtencionCondicionAlta(ml_idAtencion As Long, ml_dCondicionAlta As String, oConexion As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesXidAtencionCondicionAlta"
        Set oParameter = .CreateParameter("@ml_dCondicionAlta", adVarChar, adParamInput, 100, ml_dCondicionAlta): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ml_idAtencion", adInteger, adParamInput, 0, ml_idAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesXidAtencionCondicionAlta = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesListaCuentasXpaciente(ml_idPaciente As Long, oConexion As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesListaCuentasXpaciente"
        Set oParameter = .CreateParameter("@ml_idPaciente", adInteger, adParamInput, 0, ml_idPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesListaCuentasXpaciente = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesFiltraDatosCabecera(lnIdCuentaAtencion As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesFiltraDatosCabecera"
        Set oParameter = .CreateParameter("@lnIdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesFiltraDatosCabecera = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function PacientesXdni(lcDNI As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesXdni"
        Set oParameter = .CreateParameter("@lcDni", adVarChar, adParamInput, 12, lcDNI): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set PacientesXdni = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function PacientesXnroHistoriaTipoNumeracion(lnNroHistoriaClinica As Long, lnIdTipoNumeracion As Long, Optional oConexion1 As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New Connection
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
    End If
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "PacientesXnroHistoriaTipoNumeracion"
        Set oParameter = .CreateParameter("@lnNroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdTipoNumeracion", adInteger, adParamInput, 0, lnIdTipoNumeracion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set PacientesXnroHistoriaTipoNumeracion = oRecordset
   Set oCommand = Nothing
   If oConexion1 Is Nothing Then
      oConexion.Close
      Set oConexion = Nothing
   End If
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function atencionesXtipoFinanciamiento(lIdCuentaAtencion As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "atencionesXtipoFinanciamiento"
        Set oParameter = .CreateParameter("@lIdCuentaAtencion", adInteger, adParamInput, 0, lIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set atencionesXtipoFinanciamiento = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function FacturacionCuentasAtencionXnroCuentaConexion(lnIdCuentaAtencion As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionCuentasAtencionXnroCuentaConexion"
        Set oParameter = .CreateParameter("@lnIdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set FacturacionCuentasAtencionXnroCuentaConexion = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



Function TiposDocIdentidadXfiltro(lcFiltro As String, oConexion As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposDocIdentidadXfiltro"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 500, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set TiposDocIdentidadXfiltro = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function atencionesDatosAdicionalesXfiltro(lcFiltro As String, oConexion As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "atencionesDatosAdicionalesXfiltro"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 500, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set atencionesDatosAdicionalesXfiltro = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesDiagnosticosXidAtencion(lnIdAtencion As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDiagnosticosXidAtencion"
        Set oParameter = .CreateParameter("@lnIdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesDiagnosticosXidAtencion = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesSeleccionarPorFechasDeIngresoYtipoServicioOconexion(lcMes As String, lcAnio As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarPorFechasDeIngresoYtipoServicioOconexion"
        Set oParameter = .CreateParameter("@lcMes", adInteger, adParamInput, 0, Val(lcMes)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcAnio", adInteger, adParamInput, 0, Val(lcAnio)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesSeleccionarPorFechasDeIngresoYtipoServicioOconexion = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesEmergenciaXidAtencion(lnIdAtencion As Long, oConexion As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesEmergenciaXidAtencion"
        Set oParameter = .CreateParameter("@lnIdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesEmergenciaXidAtencion = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function TiposDestinoAtencionXidDestinoAtencion(lnIdDestinoAtencion As Long, oConexion As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposDestinoAtencionXidDestinoAtencion"
        Set oParameter = .CreateParameter("@lnIdDestinoAtencion", adInteger, adParamInput, 0, lnIdDestinoAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set TiposDestinoAtencionXidDestinoAtencion = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesEstanciaHospitalariaXidAtencionIdServicioIngreso(lnIdServicioIngreso As Long, lnIdAtencion As Long, _
                         oConexion As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesEstanciaHospitalariaXidAtencionIdServicioIngreso"
        Set oParameter = .CreateParameter("@lnIdServicioIngreso", adInteger, adParamInput, 0, lnIdServicioIngreso): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesEstanciaHospitalariaXidAtencionIdServicioIngreso = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesParaSip2000(lcFechaInicio As String, lcFechaFinal As String, oConexion As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesParaSip2000"
        Set oParameter = .CreateParameter("@lcFechaInicio", adDBTimeStamp, adParamInput, 0, lcFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcFechaFinal", adDBTimeStamp, adParamInput, 0, lcFechaFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesParaSip2000 = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesDiagnosticosParaSip2000(lnIdAtencion As Long, oConexion As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDiagnosticosParaSip2000"
        Set oParameter = .CreateParameter("@lnIdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesDiagnosticosParaSip2000 = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description

End Function

Function AtencionesNacimientosXidAtencion(lnIdAtencion As Long, oConexion As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesNacimientosXidAtencion"
        Set oParameter = .CreateParameter("@lnIdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesNacimientosXidAtencion = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesSisPorFechas(lcFechaInicio As String, lcFechaFinal As String) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSisPorFechas"
        Set oParameter = .CreateParameter("@lcFechaInicio", adDBTimeStamp, adParamInput, 0, lcFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcFechaFinal", adDBTimeStamp, adParamInput, 0, lcFechaFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSisPorFechas = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function CitasBloqueadasSeleccionarTodas(oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CitasBloqueadasSeleccionarTodas"
        Set oRecordset = .Execute
   End With
   Set CitasBloqueadasSeleccionarTodas = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function CitasWebCuposSeleccionarTodas(oConexionExterna As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionExterna
        .CommandTimeout = 150
        .CommandText = "CitasWebCuposSeleccionarTodas"
        Set oRecordset = .Execute
   End With
   Set CitasWebCuposSeleccionarTodas = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function atencionesCExServicio(lnIdServicioIngreso As Long, lcFechaIngreso As Date, oConexionExterna As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionExterna
        .CommandTimeout = 150
        .CommandText = "atencionesCExServicio"
        Set oParameter = .CreateParameter("@lnIdServicioIngreso", adInteger, adParamInput, 0, lnIdServicioIngreso): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaIngreso", adDBTimeStamp, adParamInput, 0, lcFechaIngreso): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set atencionesCExServicio = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description


End Function

Function AtencionesParaAtencionPagantesDelMedico(lnIdServicioIngreso As Long, ldFechaIngreso As Date, oConexion As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesParaAtencionPagantesDelMedico"
        Set oParameter = .CreateParameter("@ldFechaIngreso", adDBTimeStamp, adParamInput, 0, ldFechaIngreso): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdServicioIngreso", adInteger, adParamInput, 0, lnIdServicioIngreso): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesParaAtencionPagantesDelMedico = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesTFinanciamientoXidAtencion(lnIdAtencion As Long, oConexion As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesTFinanciamientoXidAtencion"
        Set oParameter = .CreateParameter("@lnIdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesTFinanciamientoXidAtencion = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function atencionesSeleccionarPorIdCuentaAtencion(lnIdCuentaAtencion As Long, mda_FechaInicio As Date, mda_FechaFin As Date) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "atencionesSeleccionarPorIdCuentaAtencion"
        Set oParameter = .CreateParameter("@lnIdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@mda_FechaInicio", adDBTimeStamp, adParamInput, 0, mda_FechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@mda_FechaFin", adDBTimeStamp, adParamInput, 0, mda_FechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set atencionesSeleccionarPorIdCuentaAtencion = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function BuscaCEAtencionesDx(idAtencion As Long) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "BuscaCEAtencionesDx"
        Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set BuscaCEAtencionesDx = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesActivasFiltraXanioIngreso(lcAnio As String) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesActivasFiltraXanioIngreso"
        Set oParameter = .CreateParameter("@lnAnio", adInteger, adParamInput, 0, Val(lcAnio)): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesActivasFiltraXanioIngreso = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function TiposDestinoAtencionSoloSIS(lnIdTipoServicio As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposDestinoAtencionSoloSIS"
        Set oParameter = .CreateParameter("@lnIdTipoServicio", adInteger, adParamInput, 0, lnIdTipoServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set TiposDestinoAtencionSoloSIS = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ChequeaSiTuboDxParto(mo_Diagnosticos As Collection) As Boolean
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    ChequeaSiTuboDxParto = False
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DiagnosticosSoloPartosSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount > 0 Then
        Dim oDODiagnostico As DOAtencionDiagnostico
        For Each oDODiagnostico In mo_Diagnosticos
            If oDODiagnostico.IdClasificacionDx = sghHospitalizacionEgreso Then
                oRecordset.MoveFirst
                oRecordset.Find "idDiagnostico=" & oDODiagnostico.IdDiagnostico
                If Not oRecordset.EOF Then
                   ChequeaSiTuboDxParto = True
                   Exit For
                End If
            End If
        Next
   End If
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
End Function

Function AtencionesNacimientosXidPacienteDeMadre(ml_IdPacienteMadre As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
Dim oConexion As New ADODB.Connection
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesNacimientosXidPacienteDeMadre"
        Set oParameter = .CreateParameter("@ml_IdPacienteMadre", adInteger, adParamInput, 0, ml_IdPacienteMadre): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesNacimientosXidPacienteDeMadre = oRecordset
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function AtencionesSeleccionarSoloNULLenEpisodiosDetalle(oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarSoloNULLenEpisodiosDetalle"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarSoloNULLenEpisodiosDetalle = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function

Function AtencionesEpisodiosCabeceraSeleccionarPorId(lnIdPaciente As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesEpisodiosCabeceraSeleccionarPorId"
        Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesEpisodiosCabeceraSeleccionarPorId = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function


Function AtencionesEpisodiosDetalleSeleccionarXpaciente(lnIdPaciente As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesEpisodiosDetalleSeleccionarXpaciente"
        Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesEpisodiosDetalleSeleccionarXpaciente = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function


Function EpisodioClinicoActualizar(oEpisodioClinico As EpisodioClinico, oDoAtencionesEpisodiosCab As DoAtencionesEpisodiosCab, _
                              oAtencionesEpisodiosCab As AtencionesEpisodiosCab, oDoAtencionesEpisodiosDet As DoAtencionesEpisodiosDet, _
                              oAtencionesEpisodiosDet As AtencionesEpisodiosDet, oDoAtencion As DOAtencion, oConexion As Connection, _
                              mi_opcion As sghOpciones) As String
        Dim oRsTmp2 As New Recordset
        Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
        Dim lnIdEpisodio As Long
        EpisodioClinicoActualizar = ""
        If mi_opcion = sghEliminar Then
                Set oRsTmp2 = AtencionesEpisodiosDetalleSeleccionarXpaciente(oDoAtencion.idPaciente, oConexion)
                oDoAtencionesEpisodiosDet.idAtencion = oDoAtencion.idAtencion
                oDoAtencionesEpisodiosDet.idEpisodio = oEpisodioClinico.idEpisodio
                oDoAtencionesEpisodiosDet.idPaciente = oDoAtencion.idPaciente
                If oAtencionesEpisodiosDet.Eliminar(oDoAtencionesEpisodiosDet) = False Then
                   EpisodioClinicoActualizar = oAtencionesEpisodiosDet.MensajeError: GoTo Terminar
                End If
                oRsTmp2.Filter = "idEpisodio=" & oEpisodioClinico.idEpisodio
                If oRsTmp2.RecordCount = 1 Then
                    oDoAtencionesEpisodiosCab.idEpisodio = oEpisodioClinico.idEpisodio
                    oDoAtencionesEpisodiosCab.idPaciente = oDoAtencion.idPaciente
                    oDoAtencionesEpisodiosCab.IdUsuarioAuditoria = 1
                    If oAtencionesEpisodiosCab.Eliminar(oDoAtencionesEpisodiosCab) = False Then
                       EpisodioClinicoActualizar = oAtencionesEpisodiosCab.MensajeError: GoTo Terminar
                    End If
                End If
        ElseIf mi_opcion = sghModificar Then
                If oEpisodioClinico.lbNuevoEpisodio = True Then
                    Set oRsTmp2 = AtencionesEpisodiosDetalleSeleccionarXpaciente(oDoAtencion.idPaciente, oConexion)
                    oRsTmp2.Filter = "idEpisodio=" & oEpisodioClinico.idEpisodio & " and idAtencion=" & oDoAtencion.idAtencion
                    If oRsTmp2.RecordCount = 0 Then
                        Set oRsTmp2 = mo_ReglasAdmision.AtencionesEpisodiosCabeceraSeleccionarPorId(oDoAtencion.idPaciente, oConexion)
                        If oRsTmp2.RecordCount = 0 Then
                           lnIdEpisodio = 1
                        Else
                           oRsTmp2.MoveLast
                           lnIdEpisodio = oRsTmp2!idEpisodio + 1
                        End If
                        oRsTmp2.Close
                        oDoAtencionesEpisodiosCab.FechaApertura = oDoAtencion.FechaIngreso
                        If oEpisodioClinico.lbCierreEpisodio = True Then
                           oDoAtencionesEpisodiosCab.FechaCierre = oDoAtencion.fechaEgreso
                        End If
                        oDoAtencionesEpisodiosCab.idEpisodio = lnIdEpisodio
                        oDoAtencionesEpisodiosCab.idPaciente = oDoAtencion.idPaciente
                        oDoAtencionesEpisodiosCab.IdUsuarioAuditoria = 1
                        If oAtencionesEpisodiosCab.Insertar(oDoAtencionesEpisodiosCab) = False Then
                           EpisodioClinicoActualizar = oAtencionesEpisodiosCab.MensajeError: GoTo Terminar
                        End If
                        oDoAtencionesEpisodiosDet.idAtencion = oDoAtencion.idAtencion
                        oDoAtencionesEpisodiosDet.idEpisodio = lnIdEpisodio
                        oDoAtencionesEpisodiosDet.idPaciente = oDoAtencion.idPaciente
                        oDoAtencionesEpisodiosDet.IdUsuarioAuditoria = 1
                        If oAtencionesEpisodiosDet.Insertar(oDoAtencionesEpisodiosDet) = False Then
                           EpisodioClinicoActualizar = oAtencionesEpisodiosDet.MensajeError: GoTo Terminar
                        End If
                     Else
                        'Ya existe: Paciente/Episodio/N°Atencion
                            oDoAtencionesEpisodiosCab.idPaciente = oDoAtencion.idPaciente
                            oDoAtencionesEpisodiosCab.idEpisodio = oEpisodioClinico.idEpisodio
                            If oAtencionesEpisodiosCab.SeleccionarPorId(oDoAtencionesEpisodiosCab) = True Then
                                If oEpisodioClinico.lbCierreEpisodio = True Then
                                   oDoAtencionesEpisodiosCab.FechaCierre = oDoAtencion.fechaEgreso
                                Else
                                   oDoAtencionesEpisodiosCab.FechaCierre = 0
                                End If
                                oDoAtencionesEpisodiosCab.IdUsuarioAuditoria = 1
                                If oAtencionesEpisodiosCab.modificar(oDoAtencionesEpisodiosCab) = False Then
                                   EpisodioClinicoActualizar = oAtencionesEpisodiosCab.MensajeError: GoTo Terminar
                                End If
                            End If
                        
                     End If
                ElseIf oEpisodioClinico.lbCierreEpisodio = True And oEpisodioClinico.idEpisodio > 0 Then
                    oDoAtencionesEpisodiosCab.idPaciente = oDoAtencion.idPaciente
                    oDoAtencionesEpisodiosCab.idEpisodio = oEpisodioClinico.idEpisodio
                    If oAtencionesEpisodiosCab.SeleccionarPorId(oDoAtencionesEpisodiosCab) = True Then
                        oDoAtencionesEpisodiosCab.FechaCierre = oDoAtencion.fechaEgreso
                        oDoAtencionesEpisodiosCab.IdUsuarioAuditoria = 1
                        If oAtencionesEpisodiosCab.modificar(oDoAtencionesEpisodiosCab) = False Then
                           EpisodioClinicoActualizar = oAtencionesEpisodiosCab.MensajeError: GoTo Terminar
                        End If
                    End If
                End If
         End If
         Set oRsTmp2 = Nothing
         Set mo_ReglasAdmision = Nothing
Terminar:
End Function


Function AtencionesSeleccionaCitasPagadas(lnIdAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionaCitasPagadas"
        Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionaCitasPagadas = oRecordset
   Set oCommand = Nothing
   Set oConexion = Nothing
Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function


Function AdmisionCEModificarAM(oDoAtencion As DOAtencion, oDiagnosticos As Collection, _
                             mo_DoAtencionDatosAdicionales As DoAtencionDatosAdicionales, lnIdTablaLISTBARITEMS As Long, _
                             lcNombrePc As String, lcObservacionesAuditoria As String, _
                             oEpisodioClinico As EpisodioClinico, ldFechaCita As Date, lnIdCita As Long, _
                             oRsSoloLabActividades As Recordset, lb_YaSeRegistroDatos As Boolean) As Boolean
      Dim oConexion As New ADODB.Connection
      Dim oPaciente As New Pacientes
      Dim oAtencion As New Atenciones
      Dim bProcesoOK As Boolean
      Dim oAtencionesDx As New AtencionesDiagnosticos, oAtencionesDatosAdicionales As New AtencionesDatosAdicionales
      Dim oDoAtencionesEpisodiosCab As New DoAtencionesEpisodiosCab, oAtencionesEpisodiosCab As New AtencionesEpisodiosCab
      Dim oDoAtencionesEpisodiosDet As New DoAtencionesEpisodiosDet, oAtencionesEpisodiosDet As New AtencionesEpisodiosDet
      Dim oDOCita As New doCita, oCitas As New Citas
      Dim oDOAtencionesLaboratorio As New DOAtencionesLaboratorio, oAtencionesLaboratorio As New AtencionesLaboratorio
      Dim oRsTmp As New Recordset
      Dim oRsTmp1 As New Recordset
      Dim lcCodigoNombre As String
      Dim lnPrecioUnitario As Double
      Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
      Dim ml_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
      Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
      Dim oBuscaCodigoNombre As New SIGHNegocios.ReglasComunes
      Dim lcSql As String

          oConexion.CommandTimeout = 900
          oConexion.CursorLocation = adUseClient
20        oConexion.Open sighentidades.CadenaConexion
          
40        Set oAtencion.Conexion = oConexion
50        Set oAtencionesDx.Conexion = oConexion
60        Set oAtencionesEpisodiosCab.Conexion = oConexion
70        Set oAtencionesEpisodiosDet.Conexion = oConexion
80        Set oAtencionesDatosAdicionales.Conexion = oConexion
          Set oCitas.Conexion = oConexion
          Set oAtencionesLaboratorio.Conexion = oConexion
          
90        oConexion.BeginTrans
100       bProcesoOK = True
If lnIdCita Then
   bProcesoOK = True
End If
          
          '
          'Modificar atencion
110       If lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
              If oDoAtencion.HoraEgreso = "99:99" Then
                oDoAtencion.HoraEgreso = sighentidades.HORA_VACIA_HM
                oDoAtencion.fechaEgreso = 0
              ElseIf lb_YaSeRegistroDatos = False Then
                oDoAtencion.HoraEgreso = lcBuscaParametro.RetornaHoraServidorSQL
                oDoAtencion.fechaEgreso = ldFechaCita
              End If
140       End If
150       If Not oAtencion.modificar(oDoAtencion) Then
160           bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
170       End If
'
          If lnIdCita > 0 Then
              oDOCita.IdCita = lnIdCita
              If oCitas.SeleccionarPorId(oDOCita) Then
                 oDOCita.IdEstadoCita = 2                    'cita atendida
                 If oCitas.modificar(oDOCita) = False Then
                      bProcesoOK = False: ms_MensajeError = oCitas.MensajeError: GoTo Terminar
                 End If
              End If
          End If
          '
180       Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "M", oDoAtencion.idAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, lcObservacionesAuditoria)   'ListBarItems.idListItem
          'Modificar Datos Adicionales
          
190       If Not oAtencionesDatosAdicionales.modificar(mo_DoAtencionDatosAdicionales) Then
200           bProcesoOK = False: ms_MensajeError = oAtencionesDatosAdicionales.MensajeError: GoTo Terminar
210       End If
          'Agregar diagnosticos
220       If Not oAtencionesDx.ActualizarDiagnosticosAtencion(oDiagnosticos, oDoAtencion.idAtencion) Then
230           bProcesoOK = False: ms_MensajeError = oAtencionesDx.MensajeError: GoTo Terminar
240       End If
                  
          
          'Episodio Clinico
250       If lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
260          ms_MensajeError = EpisodioClinicoActualizar(oEpisodioClinico, oDoAtencionesEpisodiosCab, oAtencionesEpisodiosCab, _
                                                         oDoAtencionesEpisodiosDet, oAtencionesEpisodiosDet, oDoAtencion, _
                                                         oConexion, sghModificar)
270          If ms_MensajeError <> "" Then
280             bProcesoOK = False: GoTo Terminar
290          End If
300       End If
          '
          oDOAtencionesLaboratorio.idAtencion = oDoAtencion.idAtencion
          oDOAtencionesLaboratorio.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
          If oAtencionesLaboratorio.Eliminar(oDOAtencionesLaboratorio) Then
          End If
          If oRsSoloLabActividades.RecordCount > 0 Then
             oRsSoloLabActividades.MoveFirst
             Do While Not oRsSoloLabActividades.EOF
                oDOAtencionesLaboratorio.Grupo = oRsSoloLabActividades!Grupo
                oDOAtencionesLaboratorio.lab = oRsSoloLabActividades!lab
                oDOAtencionesLaboratorio.subGrupo = oRsSoloLabActividades!subGrupo
                If oAtencionesLaboratorio.Insertar(oDOAtencionesLaboratorio) = False Then
                   ms_MensajeError = oAtencionesLaboratorio.MensajeError: bProcesoOK = False: GoTo Terminar
                End If
                oRsSoloLabActividades.MoveNext
             Loop
          End If

Terminar:
310       If bProcesoOK Then
320           AdmisionCEModificarAM = True
330           oConexion.CommitTrans
350       Else
360           AdmisionCEModificarAM = False
370           oConexion.RollbackTrans
340           ms_MensajeError = ms_MensajeError & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                                  "Function AdmisionCEModificarAM", "reglasadmision.cls") 'debb-02/05/2016
380       End If

          mo_ReglasArchivoClinico.HistoriasClinicasActualizaFmovimiento oConexion, oDoAtencion.idPaciente

                  
390       oConexion.Close
400       Set oConexion = Nothing
410       Set oPaciente = Nothing
420       Set oAtencion = Nothing
430       Set oAtencionesDx = Nothing
440       Set oBuscaCodigoNombre = Nothing
450       Set oDoCatalogoServicioHosp = Nothing
460       Set ml_ReglasFacturacion = Nothing
470       Set oRsTmp1 = Nothing
480       Set oRsTmp = Nothing
490       Set mo_ReglasArchivoClinico = Nothing
500       Set oDoAtencionesEpisodiosCab = Nothing
510       Set oAtencionesEpisodiosCab = Nothing
520       Set oDoAtencionesEpisodiosDet = Nothing
530       Set oAtencionesEpisodiosDet = Nothing
540       Set oAtencionesDatosAdicionales = Nothing
          Set oDOAtencionesLaboratorio = Nothing
          Set oAtencionesLaboratorio = Nothing
End Function

Function AdmisionHospModEGRESO(oDOCuentaAtencion As DOCuentaAtencion, oDoAtencion As DOAtencion, _
                               oOcupaciones As Collection, oDiagnosticos As Collection, oProcedimientos As Collection, _
                               oExamenes As Collection, oNacimientos As Collection, _
                               lTipoAccionAdmision As sghTipoAccionEmergenciaYHospitalizacion, _
                               ldFechaEgresoMedicoAnterior As Date, lbPacienteNN As Boolean, lnIdTablaLISTBARITEMS As Long, _
                               lcNombrePc As String, lcObservacionesAuditoria As String, _
                               oRsEstancia As Recordset, lnIdNacimientoSeleccionado As Long, _
                               mo_DoAtencionDatosAdicionales As DoAtencionDatosAdicionales, _
                               lbElServicioEsObservacionDeEmergencia As Boolean, _
                               oEpisodioClinico As EpisodioClinico, lcParametro511 As String, _
                               mo_AtencionesEmergencia As DOAtencionEmergencia) As Boolean
      Dim oConexion As New ADODB.Connection
      Dim oPaciente As New Pacientes
      Dim oCuentaAtencion As New CuentasAtencion
      Dim oAtencion As New Atenciones
      Dim oCama As New Camas
      Dim bProcesoOK As Boolean
      Dim oAtencionesDx As New AtencionesDiagnosticos
      Dim oEstanciaHospitalaria As New AtencionesEstanciaHosp
      Dim oAtencionDiagnostico As New AtencionesDiagnosticos
      Dim oAtencionesNacimientos As New AtencionesNacimientos
      Dim oFactOrdenesServicio As New FactOrdenServicio
          
      
      Dim oDOfactOrdenesServicio As New DOFactOrdenServicio
      
      Dim oFacturacionServicioDespacho_1 As New FacturacionServicioDespacho
      Dim oFactOrdenesServicio_1 As New FactOrdenServicio
      Dim oDoFactOrdenServ_1 As New DoFactOrdenServ
      Dim oDoFacturacionServicioDesp_1 As New DoFacturacionServicioDespacho

      
      Dim oFacturacionServicios As New FacturacionServicios
      Dim oDOFacturacionServicios As New DOFacturacionServicios
      Dim lcBuscaParametro As New SIGHDatos.Parametros
      Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
      Dim oFacturacionServicioPagos As New FacturacionServicioPagos
      Dim oFacturacionServicioFinan As New FacturacionServicioFin
      Dim oFactServicioPagos As New FactOrdenServicioPagos
      Dim oDOCama As New DOCama
      Dim oCamas As New Camas
      Dim oDOCamasMovimientos As New DOCamasMovimientos
      Dim oCamasMovimientos As New CamasMovimientos
      Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
      Dim oBuscaDiasPaciente As New SIGHDatos.Parametros
      Dim oDoFactOrdenServ As New DoFactOrdenServ
      Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
      Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
      Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
      Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
      Dim oAtencionesDatosAdicionales As New AtencionesDatosAdicionales
      Dim oDOAtencionNacimiento As New DOAtencionNacimiento
      Dim oDoAtencionesHijoMadre As New DoAtencionesHijoMadre, oAtencionesHijoMadre As New AtencionesHijoMadre
      Dim oDoAtencionesEpisodiosCab As New DoAtencionesEpisodiosCab, oAtencionesEpisodiosCab As New AtencionesEpisodiosCab
      Dim oDoAtencionesEpisodiosDet As New DoAtencionesEpisodiosDet, oAtencionesEpisodiosDet As New AtencionesEpisodiosDet
      Dim mo_DoAtencionDatosAdic1 As New DoAtencionDatosAdicionales
10    Dim lnIdServicio As Long: Dim lnPrecioUnitario As Double
      Dim oRsTmp As New ADODB.Recordset
      Dim oRsTmp1 As New Recordset
      Dim oRsTmp2 As New Recordset
      Dim oRsBoletaPagada As New Recordset
      Dim lnIdtipoFinanciamiento As Long
      Dim oReglasCaja As New SIGHNegocios.ReglasCaja
      
      
      Dim lcCodigoNombre As String, lcHoraEstanciaMax As String
      Dim lnCantidadXdia As Long, lnPrecioEstanciaXdia As Double, lbGeneraCabeceraPago As Boolean

          oConexion.CommandTimeout = 900
          oConexion.CursorLocation = adUseClient
30        oConexion.Open sighentidades.CadenaConexion
          'Chequea si hay Boleta  Activa, sino crear PAGO DE CONSULTA EMERGENCIA  (DEBB 27/04/2012)
50        lnIdServicio = mo_ReglasFacturacion.ObtenerCodigoDeConsultaDeEmergencia()
60        Set oRsBoletaPagada = oReglasCaja.CajaComprobantesPagoChequeaSiYaPago(oDoAtencion.idCuentaAtencion, lnIdServicio, oConexion)
70        If oRsBoletaPagada.RecordCount > 0 Then
80           lnIdServicio = 0
90        End If
100       oRsBoletaPagada.Close
          '
110       Set oAtencion.Conexion = oConexion
120       Set oCuentaAtencion.Conexion = oConexion
130       Set oCama.Conexion = oConexion
140       Set oAtencionesDx.Conexion = oConexion
150       Set oEstanciaHospitalaria.Conexion = oConexion
160       Set oAtencionDiagnostico.Conexion = oConexion
170       Set oAtencionesNacimientos.Conexion = oConexion
180       Set oAtencionesHijoMadre.Conexion = oConexion
190       Set oFactOrdenesServicio.Conexion = oConexion
200       Set oFacturacionServicios.Conexion = oConexion
210       Set oFacturacionServicioPagos.Conexion = oConexion
220       Set oFacturacionServicioFinan.Conexion = oConexion
230       Set oFacturacionServicioDespacho.Conexion = oConexion
240       Set oFactServicioPagos.Conexion = oConexion
250       Set oCamas.Conexion = oConexion
260       Set oCamasMovimientos.Conexion = oConexion
270       Set oAtencionesDatosAdicionales.Conexion = oConexion
280       Set oAtencionesEpisodiosCab.Conexion = oConexion
290       Set oAtencionesEpisodiosDet.Conexion = oConexion
          Set oFacturacionServicioDespacho_1.Conexion = oConexion
          Set oFactOrdenesServicio_1.Conexion = oConexion
          
          
          
300       oConexion.BeginTrans
310       bProcesoOK = True
320       lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
          
          'Modificar atencion
330       If Not oAtencion.modificar(oDoAtencion) Then
340           bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
350       End If
          '
360       Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "M", oDoAtencion.idAtencion, "Atenciones", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, lcObservacionesAuditoria)  'ListBarItems.idListItem
          '
370       If sighentidades.EsFecha(oDoAtencion.fechaEgreso, "DD/MM/AAAA") Then
380          oDOCuentaAtencion.IdEstado = sghEstadoCuenta.sghConAltaMedica
390       End If
      '    oDOCuentaAtencion.FechaCreacion = CDate(oDoAtencion.FechaIngreso & " " & oDoAtencion.HoraIngreso)
400       oDOCuentaAtencion.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
410       If Not oCuentaAtencion.modificar(oDOCuentaAtencion) Then
420           bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
430       End If
          '
440       If oDoAtencion.IdEstadoAtencion = sghEstadoTabla.sghAnulado Then
              'Eliminar Id paciente de las camas ocupadas
450           If oDoAtencion.idTipoServicio = sghHospitalizacion Or oDoAtencion.idTipoServicio = sghEmergenciaObservacion Then
460               If Not oCama.ActualizarIdPaciente(oDoAtencion.idPaciente, 0) Then
470                   bProcesoOK = False: ms_MensajeError = oCama.MensajeError: GoTo Terminar
480               End If
490           End If
500       Else
510           If Not oCama.ActualizarIdPaciente(oDoAtencion.idPaciente, oDoAtencion.IdCamaEgreso) Then
520               bProcesoOK = False: ms_MensajeError = oCama.MensajeError: GoTo Terminar
530           End If
540       End If
          
          'Agregar diagnosticos
550       If Not oAtencionesDx.ActualizarDiagnosticosAtencion(oDiagnosticos, oDoAtencion.idAtencion) Then
560           bProcesoOK = False: ms_MensajeError = oAtencionesDx.MensajeError: GoTo Terminar
570       End If
          
          'Modificar estancias hospitalaria solo si la cuenta aun esta abierta
580       If oDOCuentaAtencion.IdEstado = sghEstadoCuenta.sghAbierto Or oDOCuentaAtencion.IdEstado = sghEstadoCuenta.sghConAltaMedica Then
590           If Not oEstanciaHospitalaria.ActualizarEstanciaHospitalaria(oOcupaciones, oDoAtencion) Then
600               bProcesoOK = False: ms_MensajeError = oEstanciaHospitalaria.MensajeError: GoTo Terminar
610           End If
620       End If
          
          'Actualizar ocupacion o desocupacion de cama
      '    If oDoAtencion.idTipoServicio = sghHospitalizacion Or lbElServicioEsObservacionDeEmergencia = True Then
      '        If Not oCama.ActualizarIdPaciente(oDoAtencion.idPaciente, ObtenerUltimaCamaOcupada(oDoAtencion.IdCamaEgreso, oOcupaciones, oDOCuentaAtencion.idEstado)) Then
      '            bProcesoOK = False: ms_MensajeError = oCama.MensajeError: GoTo Terminar
      '        End If
      '    End If
          
         ' If oDoAtencion.idTipoServicio = sghHospitalizacion Then
              'Agregar nacimientos
630           If Not oAtencionesNacimientos.ActualizarNacimientosAtencion(oNacimientos, oDoAtencion.idAtencion) Then
640               bProcesoOK = False: ms_MensajeError = oAtencionesNacimientos.MensajeError: GoTo Terminar
650           End If
         ' End If
          
          If oDoAtencion.idTipoServicio = sghEmergenciaConsultorios Then
               Dim oAtencionEmergencia As New AtencionesEmergencia
               Set oAtencionEmergencia.Conexion = oConexion
               If Not oAtencionEmergencia.modificar(mo_AtencionesEmergencia) Then
                   bProcesoOK = False: ms_MensajeError = oAtencionEmergencia.MensajeError: GoTo Terminar
               End If
               Set oAtencionEmergencia = Nothing
          End If
          
          
       
          
          
          '***************daniel barrantes**************
          '***************Si el Paciente tiene ALTA MEDICA se genera
          '***************un SERVICIO a su CUENTA por los 'DIAS DE ESTANCIA'
660       If sighentidades.EsFecha(oDoAtencion.fechaEgreso, "DD/MM/AAAA") Then
              'Desocupa cama
670           If oDoAtencion.idTipoServicio = sghHospitalizacion Then
680              If lcBuscaParametro.SeleccionaFilaParametro(213) = "1" Then
690                   If Not oCama.ActualizarIdPaciente(oDoAtencion.idPaciente, 0) Then
700                       bProcesoOK = False: ms_MensajeError = oCama.MensajeError: GoTo Terminar
710                   End If
720              End If
730           ElseIf lcBuscaParametro.SeleccionaFilaParametro(214) = "1" Then
740                   If Not oCama.ActualizarIdPaciente(oDoAtencion.idPaciente, 0) Then
750                       bProcesoOK = False: ms_MensajeError = oCama.MensajeError: GoTo Terminar
760                   End If
770           End If
              'Genera Facturacion por ESTADIA
780           If ldFechaEgresoMedicoAnterior <> oDoAtencion.fechaEgreso Then
790               Dim lnDiasEstancia As Integer: Dim lnHorasEstancia As Integer
                  Dim lcCodServicio As String
                  Dim lbEliminaConsultaEmerg As Boolean
800               lbEliminaConsultaEmerg = False
                  
                  'Dias de Estancia
810               lnDiasEstancia = oBuscaDiasPaciente.DiasDelPacienteEnHospitalizacionEmergencia(oDoAtencion.FechaIngreso, oDoAtencion.HoraIngreso, oDoAtencion.fechaEgreso, oDoAtencion.HoraEgreso, lcHoraEstanciaMax)
          
                  'Horas de Estancia
820               lnHorasEstancia = sighentidades.HorasDelPacienteEnHospitalizacionEmergencia(oDoAtencion.FechaIngreso, oDoAtencion.HoraIngreso, oDoAtencion.fechaEgreso, oDoAtencion.HoraEgreso)
                  
                  'Codigo Servicio
830               If oDoAtencion.idTipoServicio = sghHospitalizacion Then
                     'lcCodServicio = "99221"
840                  lnIdServicio = Val(lcBuscaParametro.SeleccionaFilaParametro(202))
                     '****paciente pagante, que no se calcula estancia, pero si calcula estancia para pacientes sis, soat,...
                     If sighentidades.Parametro560 = "S" And lnIdServicio > 0 Then
                        If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oDoAtencion.IdFormaPago, oConexion) = True Then
                           lnIdServicio = 0
                        End If
                     End If
850               ElseIf lnHorasEstancia < 12 Then   'Emergencia Menor a 12 horas
                     'lcCodServicio = "99234"
860                  lnIdServicio = Val(lcBuscaParametro.SeleccionaFilaParametro(203))
870                  If oDoAtencion.idFuenteFinanciamiento = 2 Then  'solo SOAT
880                     lbEliminaConsultaEmerg = True
890                  End If
900               ElseIf lnHorasEstancia <= 24 Then  'Emergencia Entre 12 y 24 horas
                     'lcCodServicio = "99235"
910                  lnIdServicio = Val(lcBuscaParametro.SeleccionaFilaParametro(204))
920                  If oDoAtencion.idFuenteFinanciamiento = 2 Then  'solo SOAT
930                     lbEliminaConsultaEmerg = True
940                  End If
950               Else                               'Emergencia >24 horas
                     'lcCodServicio = "99235"
960                  lnIdServicio = Val(lcBuscaParametro.SeleccionaFilaParametro(204))
970                  If oDoAtencion.idFuenteFinanciamiento = 2 Then  'solo SOAT
980                     lbEliminaConsultaEmerg = True
990                  End If
1000              End If
                  
                  
1010              If lnIdServicio > 0 Then
                      'elimina Codigo Servicio (cabecera/detalle) por Estancia de la Cuenta de Atencion
1020                  Set oRsTmp = mo_ReglasFacturacion.FactOrdenServicioFiltraPorIdCuenta(oDoAtencion.idCuentaAtencion)
1030                  If oRsTmp.RecordCount > 0 Then
1040                     oRsTmp.MoveFirst
1050                     Do While Not oRsTmp.EOF
1060                        If oRsTmp.Fields!IdPuntoCarga = 9 Then
                               '**********solo Hospitalizacion
1070                           oDoFactOrdenServ.IdOrden = oRsTmp.Fields!IdOrden
1080                           oDoFacturacionServicioFinanc.IdOrden = oRsTmp.Fields!IdOrden
1090                           oDoFacturacionServicioDesp.IdOrden = oRsTmp.Fields!IdOrden
1100                           Set oRsTmp1 = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrden(oRsTmp.Fields!IdOrden, oConexion)
1110                           If oRsTmp1.RecordCount > 0 Then
1120                             If oRsTmp1.Fields!IdEstadoFacturacion = 1 Then
1130                                  oDOFactOrdenServicioPagos.IdOrden = oRsTmp1.Fields!IdOrden
1140                                  oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp1.Fields!idOrdenPago
1150                                  oDoFacturacionServicioPagos.idOrdenPago = oRsTmp1.Fields!idOrdenPago
1160                                  If Not oFacturacionServicioPagos.Eliminar(oDoFacturacionServicioPagos) Then
1170                                     bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
1180                                  End If
1190                                  If Not oFactServicioPagos.Eliminar(oDOFactOrdenServicioPagos) Then
1200                                     bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
1210                                  End If
1220                             End If
1230                           End If
1240                           oDoFacturacionServicioFinanc.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1250                           If Not oFacturacionServicioFinan.Eliminar(oDoFacturacionServicioFinanc) Then
1260                              bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo Terminar
1270                           End If
1280                           oDoFacturacionServicioDesp.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1290                           If Not oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then
1300                              bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
1310                           End If
1320                           oDoFactOrdenServ.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1330                           If Not oFactOrdenesServicio.Eliminar(oDoFactOrdenServ) Then
1340                              bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
1350                           End If
1360                        ElseIf oRsTmp.Fields!IdPuntoCarga = 10 Then
                               '*********solo Emergencia
1370                           If lbEliminaConsultaEmerg = True And oRsTmp.Fields!idProducto = Val(lcBuscaParametro.SeleccionaFilaParametro(202)) Then
                                  'Consulta Emergencia/Solo SOAT
1380                              oDoFacturacionServicioFinanc.IdOrden = oRsTmp.Fields!IdOrden
1390                              oDoFacturacionServicioFinanc.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1400                              If Not oFacturacionServicioFinan.SeleccionarPorId(oDoFacturacionServicioFinanc) Then
1410                                 bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo Terminar
1420                              End If
1430                              oDoFacturacionServicioFinanc.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1440                              If Not oFacturacionServicioFinan.Eliminar(oDoFacturacionServicioFinanc) Then
1450                                 bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo Terminar
1460                              End If
1470                              oDoFactOrdenServ.IdOrden = oRsTmp.Fields!IdOrden
1480                              If Not oFactOrdenesServicio.SeleccionarPorId(oDoFactOrdenServ) Then
1490                                 bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
1500                              End If
1510                              oDoFactOrdenServ.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1520                              oDoFactOrdenServ.IdEstadoFacturacion = 9   'Anulado
1530                              If Not oFactOrdenesServicio.modificar(oDoFactOrdenServ) Then
1540                                 bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
1550                              End If
1560                           End If
1570                        End If
1580                        oRsTmp.MoveNext
1590                     Loop
1600                  End If
1610                  oRsTmp.Close
1620                  lnIdtipoFinanciamiento = oDoAtencion.IdFormaPago
                       'Crear orden de servicio
1630                   With oDoFactOrdenServ
1640                       .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL()
1650                       .idCuentaAtencion = oDoAtencion.idCuentaAtencion
1660                       .IdServicioPaciente = oDoAtencion.IdServicioEgreso
1670                       .IdEstadoFacturacion = 1               'Registrado
1680                       .idPaciente = oDoAtencion.idPaciente
1690                       .IdPuntoCarga = IIf(oDoAtencion.idTipoServicio = sghHospitalizacion, 9, 10)
1700                       .IdTipoFinanciamiento = lnIdtipoFinanciamiento
1710                       .idFuenteFinanciamiento = oDoAtencion.idFuenteFinanciamiento
1720                       .idUsuario = oDoAtencion.IdUsuarioAuditoria
1730                       .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1740                       .fechaDespacho = .FechaCreacion
1750                       .IdUsuarioDespacho = oDoAtencion.IdUsuarioAuditoria
1760                   End With
1770                   If Not oFactOrdenesServicio.Insertar(oDoFactOrdenServ) Then
1780                       bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo Terminar
1790                   End If
                       lbGeneraCabeceraPago = False
                       If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(lnIdtipoFinanciamiento, oConexion) = True Then
                          lbGeneraCabeceraPago = True
                       Else
                          oRsEstancia.MoveFirst
                          oRsEstancia.Find "idTipoFinanciamiento=1"
                          If Not oRsEstancia.EOF Then
                             lbGeneraCabeceraPago = True
                          End If
                       End If
1800                   If lbGeneraCabeceraPago = True Then
                         '***********Contado ***********
1810                     With oDOFactOrdenServicioPagos
1820                         .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL()
1830                         .IdEstadoFacturacion = 1   'Registrado
1840                         .IdOrden = oDoFactOrdenServ.IdOrden
1850                         .idUsuario = oDoAtencion.IdUsuarioAuditoria
1860                         .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
1870                     End With
1880                     If Not oFactServicioPagos.Insertar(oDOFactOrdenServicioPagos) Then
1890                         bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo Terminar
1900                     End If
1910                   End If
1920                  oRsEstancia.MoveFirst
1930                  Do While Not oRsEstancia.EOF
                          If lcParametro511 = "S" Then
                             lnPrecioEstanciaXdia = oReglasCaja.DevuelvePrecioCPTporProductoTfinanciamiento(oRsEstancia!idProducto, oRsEstancia!IdTipoFinanciamiento, oConexion)
                          End If
                          'Crear Detalle de servicio
1940                      With oDoFacturacionServicioDesp
1960                          .IdOrden = oDoFactOrdenServ.IdOrden
1970                          .idProducto = oRsEstancia.Fields!idProducto
1980                          .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                              If lcParametro511 = "S" Then
                                    lnCantidadXdia = Round(oRsEstancia.Fields!CantidadEstancia / 24, 0)
                                    .cantidad = lnCantidadXdia
                                    .Precio = lnPrecioEstanciaXdia
                                    .total = Round(lnCantidadXdia * lnPrecioEstanciaXdia, 2)
                              Else
                                    .cantidad = oRsEstancia.Fields!CantidadEstancia
                                    .Precio = oRsEstancia.Fields!PrecioEstancia
                                    .total = Round(oRsEstancia.Fields!CantidadEstancia * oRsEstancia.Fields!PrecioEstancia, 2)
                              End If
2010                      End With
2020                      If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then
                                 If Left(oFacturacionServicioDespacho.MensajeError, 11) = "-2147217873" Then
'                                    With oDoFactOrdenServ_1
'                                        .FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL()
'                                        .idCuentaAtencion = oDoAtencion.idCuentaAtencion
'                                        .IdServicioPaciente = oDoAtencion.IdServicioEgreso
'                                        .IdEstadoFacturacion = 1               'Registrado
'                                        .idPaciente = oDoAtencion.idPaciente
'                                        .IdPuntoCarga = IIf(oDoAtencion.idTipoServicio = sghHospitalizacion, 9, 10)
'                                        .IdTipoFinanciamiento = oRsEstancia!IdTipoFinanciamiento
'                                        .idFuenteFinanciamiento = oRsEstancia!idFuenteFinanciamiento
'                                        .idUsuario = oDoAtencion.IdUsuarioAuditoria
'                                        .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
'                                        .fechaDespacho = .FechaCreacion
'                                        .IdUsuarioDespacho = oDoAtencion.IdUsuarioAuditoria
'                                    End With
'                                    If Not oFactOrdenesServicio_1.Insertar(oDoFactOrdenServ_1) Then
'                                        bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio_1.MensajeError: GoTo Terminar
'                                    End If
'                                    With oDoFacturacionServicioDesp_1
'                                          .IdOrden = oDoFactOrdenServ_1.IdOrden
'                                          .idProducto = oRsEstancia.Fields!idProducto
'                                          .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
'                                              If lcParametro511 = "S" Then
'                                                    lnCantidadXdia = Round(oRsEstancia.Fields!CantidadEstancia / 24, 0)
'                                                    .cantidad = lnCantidadXdia
'                                                    .Precio = lnPrecioEstanciaXdia
'                                                    .total = Round(lnCantidadXdia * lnPrecioEstanciaXdia, 2)
'                                              Else
'                                                    .cantidad = oRsEstancia.Fields!CantidadEstancia
'                                                    .Precio = oRsEstancia.Fields!PrecioEstancia
'                                                    .total = Round(oRsEstancia.Fields!CantidadEstancia * oRsEstancia.Fields!PrecioEstancia, 2)
'                                              End If
'                                    End With
'                                    If Not oFacturacionServicioDespacho_1.Insertar(oDoFacturacionServicioDesp_1) Then
'                                       bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho_1.MensajeError: GoTo Terminar
'                                    End If
                                 Else
                                    bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo Terminar
                                 End If
2040                      End If
2050                      If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oRsEstancia!IdTipoFinanciamiento, oConexion) = True Then
                             '***********Contado ***********
2060                         With oDoFacturacionServicioPagos
2080                             .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
2090                             .idProducto = oRsEstancia.Fields!idProducto
2100                             .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
                                 If lcParametro511 = "S" Then
                                    lnCantidadXdia = Round(oRsEstancia.Fields!CantidadEstancia / 24, 0)
                                    .cantidad = lnCantidadXdia
                                    .Precio = lnPrecioEstanciaXdia
                                    .total = Round(lnCantidadXdia * lnPrecioEstanciaXdia, 2)
                                 Else
                                    .cantidad = oRsEstancia.Fields!CantidadEstancia
                                    .Precio = oRsEstancia.Fields!PrecioEstancia
                                    .total = Round(oRsEstancia.Fields!CantidadEstancia * oRsEstancia.Fields!PrecioEstancia, 2)
                                 End If
2130                         End With
2140                         If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
2150                             bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo Terminar
2160                         End If
2170                      Else
                             '***********Sis,Soat,exoneracion, convenio ***********
2180                         With oDoFacturacionServicioFinanc
2200                             .FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
2210                             .IdOrden = oDoFactOrdenServ.IdOrden
2220                             .idProducto = oRsEstancia.Fields!idProducto
2230                             .IdTipoFinanciamiento = lnIdtipoFinanciamiento
2240                             .IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2250                             .IdUsuarioAutoriza = oDoAtencion.IdUsuarioAuditoria
2280                             .idFuenteFinanciamiento = oDoAtencion.idFuenteFinanciamiento
                                 If lcParametro511 = "S" Then
                                    lnCantidadXdia = Round(oRsEstancia.Fields!CantidadEstancia / 24, 0)
                                    .CantidadFinanciada = lnCantidadXdia
                                    .PrecioFinanciado = lnPrecioEstanciaXdia
                                    .TotalFinanciado = Round(lnCantidadXdia * lnPrecioEstanciaXdia, 2)
                                 Else
                                    .CantidadFinanciada = oRsEstancia.Fields!CantidadEstancia
                                    .PrecioFinanciado = oRsEstancia.Fields!PrecioEstancia
                                    .TotalFinanciado = Round(oRsEstancia.Fields!CantidadEstancia * oRsEstancia.Fields!PrecioEstancia, 2)
                                 End If
2290                         End With
2300                         If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then
2310                             bProcesoOK = False: GoTo Terminar
2320                         End If
2330                      End If
2340                      oRsEstancia.MoveNext
2350                  Loop
2360              End If
2370          End If
2380      End If
          '
          '-----Chequea Si la "Cama de Egreso" no pertenece al "Servicio de Egreso"
          '-----de ser asi la "cama" pasará  al "servicio" como PRESTAMO en forma automatica
          '-----si tiene ALTA MEDICA, volvera a su SERVICIO PROPIETARIO en forma automática  (debb-13/09/2016)
2390      If oDoAtencion.IdCamaEgreso > 0 Then
2400          oDOCama.IdCama = oDoAtencion.IdCamaEgreso
2410          If Not oCamas.SeleccionarPorId(oDOCama) Then
2420              bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
2430          End If
2440          If oDOCama.IdServicioUbicacionActual <> oDoAtencion.IdServicioEgreso Then
                 If sighentidades.EsFecha(oDoAtencion.fechaEgreso, "DD/MM/AAAA") Then
                    oDOCama.IdServicioUbicacionActual = oDOCama.IdServicioPropietario
                 Else
2450                oDOCama.IdServicioUbicacionActual = oDoAtencion.IdServicioEgreso
                 End If
2460             oDOCama.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2470             If Not oCamas.modificar(oDOCama) Then
2480                  bProcesoOK = False: ms_MensajeError = oCamas.MensajeError: GoTo Terminar
2490             End If
2500             oDOCamasMovimientos.IdCama = oDOCama.IdCama
2510             oDOCamasMovimientos.IdFechaIngreso = lcBuscaParametro.RetornaFechaHoraServidorSQL
                 If sighentidades.EsFecha(oDoAtencion.fechaEgreso, "DD/MM/AAAA") Then
                    oDOCamasMovimientos.IdServicio = oDOCama.IdServicioPropietario
                 Else
2520                oDOCamasMovimientos.IdServicio = oDOCama.IdServicioUbicacionActual
                 End If
2530             oDOCamasMovimientos.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2540             If Not oCamasMovimientos.Insertar(oDOCamasMovimientos) Then
2550                  bProcesoOK = False: ms_MensajeError = oCamasMovimientos.MensajeError: GoTo Terminar
2560             End If
2570          End If
2580      End If
          '
          '****Actualiza Dato del Recien Nacido (que está en la Cuenta de la Madre)....tabla: AtencionesNacimiento
          '****Asocia la Cuenta de la madre con la cuenta con el hijo..................tabla: AtencionesHijoMadre
2590      If oAtencionesNacimientos.AsignaNULLidPacienteNacido(oDoAtencion.idPaciente) = False Then
2600         bProcesoOK = False: ms_MensajeError = oAtencionesNacimientos.MensajeError: GoTo Terminar
2610      End If
2620      oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
2630      oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2640      If oAtencionesHijoMadre.Eliminar(oDoAtencionesHijoMadre) = False Then
2650         bProcesoOK = False: ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
2660      End If
2670      If lnIdNacimientoSeleccionado > 0 Then
             Dim lnIdAtencionMadre As Long
2680         oDOAtencionNacimiento.idNacimiento = lnIdNacimientoSeleccionado
2690         If oAtencionesNacimientos.SeleccionarPorId(oDOAtencionNacimiento) = True Then
2700            lnIdAtencionMadre = oDOAtencionNacimiento.idAtencion
2710            oDOAtencionNacimiento.idPacienteNacido = oDoAtencion.idPaciente
2720            oDOAtencionNacimiento.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2730            If oAtencionesNacimientos.modificar(oDOAtencionNacimiento) = True Then
2740               bProcesoOK = False: ms_MensajeError = oAtencionesNacimientos.MensajeError: GoTo Terminar
2750            End If
2760            oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
2770            If oAtencionesHijoMadre.SeleccionarPorId(oDoAtencionesHijoMadre) = True Then
2780                 oDoAtencionesHijoMadre.idAtencionMadre = lnIdAtencionMadre
2790                 oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2800                 If oAtencionesHijoMadre.modificar(oDoAtencionesHijoMadre) = True Then
2810                    bProcesoOK = False: ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
2820                 End If
2830            Else
2840                 oDoAtencionesHijoMadre.idAtencion = oDoAtencion.idAtencion
2850                 oDoAtencionesHijoMadre.idAtencionMadre = lnIdAtencionMadre
2860                 oDoAtencionesHijoMadre.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
2870                 If oAtencionesHijoMadre.Insertar(oDoAtencionesHijoMadre) = True Then
2880                    bProcesoOK = False: ms_MensajeError = oAtencionesHijoMadre.MensajeError: GoTo Terminar
2890                 End If
2900            End If
2910         End If
2920      End If

          'datos adicionales
2930      mo_DoAtencionDatosAdicionales.idAtencion = oDoAtencion.idAtencion
             mo_DoAtencionDatosAdic1.idAtencion = oDoAtencion.idAtencion
             mo_DoAtencionDatosAdic1.IdUsuarioAuditoria = oDoAtencion.IdUsuarioAuditoria
             If oAtencionesDatosAdicionales.SeleccionarPorId(mo_DoAtencionDatosAdic1) = True Then
                If Not oAtencionesDatosAdicionales.modificar(mo_DoAtencionDatosAdicionales) Then
                   bProcesoOK = False: ms_MensajeError = oAtencionesDatosAdicionales.MensajeError: GoTo Terminar
                End If
             Else
                mo_DoAtencionDatosAdicionales.idAtencion = oDoAtencion.idAtencion
                If Not oAtencionesDatosAdicionales.Insertar(mo_DoAtencionDatosAdicionales) Then
                   bProcesoOK = False: ms_MensajeError = oAtencionesDatosAdicionales.MensajeError: GoTo Terminar
                End If
             End If
2970      If mo_DoAtencionDatosAdicionales.idSiaSis > 0 Then
2980         Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencion.IdUsuarioAuditoria, "M", oDoAtencion.idCuentaAtencion, "FUA", oConexion, 1345, lcNombrePc, "Cod.Prestacion FUA: " & mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion & " N° Cuenta: " & oDoAtencion.idCuentaAtencion)
2990      End If
          'Episodio Clinico
3000      If sighentidades.EsFecha(oDoAtencion.fechaEgreso, "DD/MM/AAAA") Then
3010         ms_MensajeError = EpisodioClinicoActualizar(oEpisodioClinico, oDoAtencionesEpisodiosCab, oAtencionesEpisodiosCab, _
                                                         oDoAtencionesEpisodiosDet, oAtencionesEpisodiosDet, oDoAtencion, _
                                                         oConexion, sghModificar)
3020         If ms_MensajeError <> "" Then
3030            bProcesoOK = False: GoTo Terminar
3040         End If
3050      End If
          '
          mo_ReglasArchivoClinico.HistoriasClinicasActualizaFmovimiento oConexion, oDoAtencion.idPaciente
3060      ms_MensajeError = ""
          '
Terminar:

3070      If bProcesoOK Then
3080          AdmisionHospModEGRESO = True
3090          oConexion.CommitTrans
3100      Else
3110          If Err.Description <> "" Then
3120             ms_MensajeError = Err.Description & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                                    "Function AdmisionHospModEGRESO", "reglasAdmision.cls")  'debb-02/05/2016
3130          End If
3140          AdmisionHospModEGRESO = False
3150          oConexion.RollbackTrans
              
3160      End If
                  
3170      oConexion.Close
3180      Set oRsBoletaPagada = Nothing
3190      Set oConexion = Nothing
3200      Set oPaciente = Nothing
3210      Set oCuentaAtencion = Nothing
3220      Set oAtencion = Nothing
3230      Set oCama = Nothing
3240      Set oAtencionesDx = Nothing
3250      Set oEstanciaHospitalaria = Nothing
3260      Set oAtencionDiagnostico = Nothing
3270      Set oAtencionesNacimientos = Nothing
3280      Set oFactOrdenesServicio = Nothing
3290      Set oDOfactOrdenesServicio = Nothing
3300      Set oFacturacionServicios = Nothing
3310      Set oDOFacturacionServicios = Nothing
3320      Set lcBuscaParametro = Nothing
3330      Set oFacturacionServicioDespacho = Nothing
3340      Set oFacturacionServicioPagos = Nothing
3350      Set oFacturacionServicioFinan = Nothing
3360      Set oFactServicioPagos = Nothing
3370      Set oDOCama = Nothing
3380      Set oCamas = Nothing
3390      Set oDOCamasMovimientos = Nothing
3400      Set oCamasMovimientos = Nothing
3410      Set oDoCatalogoServicioHosp = Nothing
3420      Set oBuscaDiasPaciente = Nothing
3430      Set oDoFactOrdenServ = Nothing
3440      Set oDoFacturacionServicioPagos = Nothing
3450      Set oDoFacturacionServicioFinanc = Nothing
3460      Set oDoFacturacionServicioDesp = Nothing
3470      Set oDOFactOrdenServicioPagos = Nothing
3480      Set oRsTmp = Nothing
3490      Set oRsTmp1 = Nothing
3500      Set oRsTmp2 = Nothing
3510      Set oAtencionesDatosAdicionales = Nothing
3520      Set oDOAtencionNacimiento = Nothing
3530      Set oDoAtencionesHijoMadre = Nothing
3540      Set oAtencionesHijoMadre = Nothing
3550      Set oDoAtencionesEpisodiosCab = Nothing
3560      Set oAtencionesEpisodiosCab = Nothing
3570      Set oDoAtencionesEpisodiosDet = Nothing
3580      Set oAtencionesEpisodiosDet = Nothing
          Set mo_DoAtencionDatosAdic1 = Nothing
          Set oFacturacionServicioDespacho_1 = Nothing
          Set oFactOrdenesServicio_1 = Nothing
          Set oDoFactOrdenServ_1 = Nothing
          Set oDoFacturacionServicioDesp_1 = Nothing
End Function

'CREADO POR MARIO 21/03/14

Function AtencionesTiempoCE(lcAnio As String, ldFechaInicial As Date, ldFechaFinal As Date) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesTiempoAtencionCE"
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, Val(lcAnio)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 8, ldFechaInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 8, ldFechaFinal): .Parameters.Append oParameter
        
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesTiempoCE = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'CREADO POR MARIO 02/04/2014

Function UsuarioPC(lnIdAtencion As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesUsuarioPC"
        Set oParameter = .CreateParameter("@Atencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set UsuarioPC = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Sub his_historicoAtencionesEliminarTodas(oConexion As Connection)
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "his_historicoAtencionesEliminarTodas"
        .Execute
   End With
   Set oCommand = Nothing
   Exit Sub
ManejadorDeError:
    MsgBox Err.Description
End Sub


Function AtencionesDatosAdicionalesSeleccionarXidAtencion(lnIdAtencion As Long, Optional oConexion1 As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    If oConexion1 Is Nothing Then
        sighentidades.AbreConexionSIGH oConexion
    End If
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "AtencionesDatosAdicionalesSeleccionarXidAtencion"
        Set oParameter = .CreateParameter("@lnIdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
            Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set AtencionesDatosAdicionalesSeleccionarXidAtencion = oRecordset
   If oConexion1 Is Nothing Then
      oConexion.Close
      Set oRecordset = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'Franklin Cachay 06/03/2014 --- Modulo de Enfermerias
Function Enfermeria_CatalogoDominiosSeleccionar(oConexion As ADODB.Connection) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Set Enfermeria_CatalogoDominiosSeleccionar = Nothing

    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "Enfermeria_CatalogoDominiosSeleccionar"
         Set oRsTmp = .Execute
    End With
    Set Enfermeria_CatalogoDominiosSeleccionar = oRsTmp
    
    Set oRsTmp = Nothing
    Set oCommand = Nothing
End Function

Function Enfermeria_CatalogoVariablesSeleccionar(oConexion As ADODB.Connection) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Set Enfermeria_CatalogoVariablesSeleccionar = Nothing
    
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "Enfermeria_CatalogoVariablesSeleccionar"
         Set oRsTmp = .Execute
    End With
    Set Enfermeria_CatalogoVariablesSeleccionar = oRsTmp

    Set oRsTmp = Nothing
    Set oCommand = Nothing
End Function

Function Enfermeria_TotalDatosGrafico(oConexion As ADODB.Connection) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Set Enfermeria_TotalDatosGrafico = Nothing
    
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "Enfermeria_TotalDatosGrafico"
         Set oRsTmp = .Execute
    End With
    Set Enfermeria_TotalDatosGrafico = oRsTmp

    Set oRsTmp = Nothing
    Set oCommand = Nothing
End Function

Function Enfermeria_VariablesComboSeleccionarValores(oConexion As ADODB.Connection, lnIdVariable As Long) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As New ADODB.Parameter
    Set Enfermeria_VariablesComboSeleccionarValores = Nothing
    
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "Enfermeria_VariablesComboSeleccionarValores"
         Set oParameter = .CreateParameter("@IdVariable", adInteger, adParamInput, 0, lnIdVariable): .Parameters.Append oParameter
         Set oRsTmp = .Execute
    End With
    Set Enfermeria_VariablesComboSeleccionarValores = oRsTmp

    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Function Enfermeria_ConsultarUltimaVisita(oConexion As ADODB.Connection, lnIdCuentaAtencion As Long) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As New ADODB.Parameter
    Set Enfermeria_ConsultarUltimaVisita = Nothing
    
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "Enfermeria_ConsultarUltimaVisita"
         Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
         Set oRsTmp = .Execute
    End With
    Set Enfermeria_ConsultarUltimaVisita = oRsTmp

    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Function Enfermeria_ConsultarDatosActualesPaciente(oConexion As ADODB.Connection, lnIdCuentaAtencion As Long) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As New ADODB.Parameter
    Set Enfermeria_ConsultarDatosActualesPaciente = Nothing
    
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "Enfermeria_ConsultarDatosActualesPaciente"
         Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
         Set oRsTmp = .Execute
    End With
    Set Enfermeria_ConsultarDatosActualesPaciente = oRsTmp

    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Function EnfermeriaCargarDatosPacientePorVisita(oConexion As ADODB.Connection, lnIdCuentaAtencion As Long, lnIdVisita As Integer) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As New ADODB.Parameter
    Set EnfermeriaCargarDatosPacientePorVisita = Nothing
    
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "EnfermeriaCargarDatosPacientePorVisita"
         Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@IdVisita", adInteger, adParamInput, 0, lnIdVisita): .Parameters.Append oParameter
         Set oRsTmp = .Execute
    End With
    Set EnfermeriaCargarDatosPacientePorVisita = oRsTmp

    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Function Enfermeria_CatalogoVisitas(oConexion As ADODB.Connection, lnIdCuentaAtencion As Long) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As New ADODB.Parameter
    Set Enfermeria_CatalogoVisitas = Nothing
    
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "Enfermeria_CatalogoVisitas"
         Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
         Set oRsTmp = .Execute
    End With
    Set Enfermeria_CatalogoVisitas = oRsTmp

    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Function EnfermeriaCargarDatosVariablePorVisita(oConexion As ADODB.Connection, lnIdCuentaAtencion As Long, lnIdVisita As Integer) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As New ADODB.Parameter
    Set EnfermeriaCargarDatosVariablePorVisita = Nothing
    
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "EnfermeriaCargarDatosVariablePorVisita"
         Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@IdVisita", adInteger, adParamInput, 0, lnIdVisita): .Parameters.Append oParameter
         Set oRsTmp = .Execute
    End With
    Set EnfermeriaCargarDatosVariablePorVisita = oRsTmp

    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Function EnfermeriaDatosComboPorVisita(oConexion As ADODB.Connection, lnIdCuentaAtencion As Long, lnIdVisita As Integer) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As New ADODB.Parameter
    Set EnfermeriaDatosComboPorVisita = Nothing
    
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "EnfermeriaDatosComboPorVisita"
         Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@IdVisita", adInteger, adParamInput, 0, lnIdVisita): .Parameters.Append oParameter
         Set oRsTmp = .Execute
    End With
    Set EnfermeriaDatosComboPorVisita = oRsTmp

    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Function Enfermeria_ConsultaValoresVariableGrafico(oConexion As ADODB.Connection, lnIdCuentaAtencion As Long, lnIdVariable As Integer, lnIdVisita As Integer) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As New ADODB.Parameter
    Set Enfermeria_ConsultaValoresVariableGrafico = Nothing
    
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "Enfermeria_ConsultaValoresVariableGrafico"
         Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@IdVariable", adInteger, adParamInput, 0, lnIdVariable): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@IdVisita", adInteger, adParamInput, 0, lnIdVisita): .Parameters.Append oParameter
         Set oRsTmp = .Execute
    End With
    Set Enfermeria_ConsultaValoresVariableGrafico = oRsTmp

    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function


Function EnfermeriaVisitasAgregar(oDOEnfermeria_Visitas As DOEnfermeria_Visitas, _
         oRsDatosEnfermeriaVariables As Recordset, oRsDatosEnfDatosValoresCombo As Recordset, _
         oRsEnfermeriaTratamientoDosis As Recordset, oDoPacienteDatosAdd As DoPacienteDatosAdd) As Boolean
    On Error GoTo ErrCeAgr
    Dim oConexion As New ADODB.Connection
    Dim oEnfermeria_Visitas As New Enfermeria_Visitas
    Dim oEnfermeria_Variables As New Enfermeria_Variables, oDOEnfermeria_Variables As New DOEnfermeria_Variables
    Dim oEnfermeria_ValoresCombo As New Enfermeria_ValoresCombo, oDoEnfermeria_ValoresCombo As New DoEnfermeria_ValoresCombo
    Dim oEnfermeria_TratamientoDosis As New Enfermeria_TratamientoDosis, oDOEnfermeria_TratamientoDosis As New DOEnfermeria_TratamientoDosis
    Dim oPacientesDatosAdd As New PacientesDatosAdd
       
    Dim orsTemp As New ADODB.Recordset
    EnfermeriaVisitasAgregar = False
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    '
    oConexion.BeginTrans
    '
    Set oEnfermeria_Visitas.Conexion = oConexion
    Set oEnfermeria_Variables.Conexion = oConexion
    Set oEnfermeria_ValoresCombo.Conexion = oConexion
    Set oEnfermeria_TratamientoDosis.Conexion = oConexion
    Set oPacientesDatosAdd.Conexion = oConexion

    If Not oPacientesDatosAdd.modificar(oDoPacienteDatosAdd) Then
        ms_MensajeError = oPacientesDatosAdd.MensajeError: GoTo ErrCeAgr
    End If

    If Not oEnfermeria_Visitas.Insertar(oDOEnfermeria_Visitas) Then
        ms_MensajeError = oEnfermeria_Visitas.MensajeError: GoTo ErrCeAgr
    End If
    'Inserta los datos de las variables
    If oRsDatosEnfermeriaVariables.RecordCount > 0 Then
        oRsDatosEnfermeriaVariables.MoveFirst
        Do While Not oRsDatosEnfermeriaVariables.EOF
           oDOEnfermeria_Variables.idCuentaAtencion = oDOEnfermeria_Visitas.idCuentaAtencion
           oDOEnfermeria_Variables.IdUsuarioAuditoria = oDOEnfermeria_Visitas.IdUsuarioAuditoria
           oDOEnfermeria_Variables.IdVisita = oDOEnfermeria_Visitas.IdVisita
           oDOEnfermeria_Variables.IdVariable = oRsDatosEnfermeriaVariables.Fields!IdVariable
           oDOEnfermeria_Variables.VariableDato = IIf(IsNull(oRsDatosEnfermeriaVariables.Fields!VariableDato), "", oRsDatosEnfermeriaVariables.Fields!VariableDato)
           If Not oEnfermeria_Variables.Insertar(oDOEnfermeria_Variables) Then
             ms_MensajeError = oEnfermeria_Variables.MensajeError: GoTo ErrCeAgr
           End If
           oRsDatosEnfermeriaVariables.MoveNext
        Loop
    End If
    'Inserta los datos de la variable combo
    If oRsDatosEnfDatosValoresCombo.RecordCount > 0 Then
        oRsDatosEnfDatosValoresCombo.MoveFirst
        Do While Not oRsDatosEnfDatosValoresCombo.EOF
           oDoEnfermeria_ValoresCombo.idCuentaAtencion = oDOEnfermeria_Visitas.idCuentaAtencion
           oDoEnfermeria_ValoresCombo.IdVisita = oDOEnfermeria_Visitas.IdVisita
           oDoEnfermeria_ValoresCombo.IdVariable = oRsDatosEnfDatosValoresCombo.Fields!IdVariable
           oDoEnfermeria_ValoresCombo.IdValorCombo = oRsDatosEnfDatosValoresCombo.Fields!IdValorCombo
           oDoEnfermeria_ValoresCombo.IdUsuarioAuditoria = oDOEnfermeria_Visitas.IdUsuarioAuditoria
           If Not oEnfermeria_ValoresCombo.Insertar(oDoEnfermeria_ValoresCombo) Then
             ms_MensajeError = oEnfermeria_ValoresCombo.MensajeError: GoTo ErrCeAgr
           End If
           oRsDatosEnfDatosValoresCombo.MoveNext
        Loop
    End If
    
     'Inserta los datos de medicamentos de la visita
    If oRsEnfermeriaTratamientoDosis.RecordCount > 0 Then
        oRsEnfermeriaTratamientoDosis.MoveFirst
        Do While Not oRsEnfermeriaTratamientoDosis.EOF
           oDOEnfermeria_TratamientoDosis.idCuentaAtencion = oRsEnfermeriaTratamientoDosis.Fields!idCuentaAtencion
           oDOEnfermeria_TratamientoDosis.IdVisita = oRsEnfermeriaTratamientoDosis.Fields!IdVisita
           oDOEnfermeria_TratamientoDosis.IdDiaVisita = oRsEnfermeriaTratamientoDosis.Fields!IdDiaVisita
           oDOEnfermeria_TratamientoDosis.idReceta = oRsEnfermeriaTratamientoDosis.Fields!idReceta
           oDOEnfermeria_TratamientoDosis.idItem = oRsEnfermeriaTratamientoDosis.Fields!idItem
           If Not IsNull(oRsEnfermeriaTratamientoDosis.Fields!Dosis) Then
                oDOEnfermeria_TratamientoDosis.Dosis = oRsEnfermeriaTratamientoDosis.Fields!Dosis
           Else
                If Not IsNull(oRsEnfermeriaTratamientoDosis.Fields!DatoProrenata) Then
                    oDOEnfermeria_TratamientoDosis.DatoProrenata = oRsEnfermeriaTratamientoDosis.Fields!DatoProrenata
                End If
           End If
           If Not oEnfermeria_TratamientoDosis.Insertar(oDOEnfermeria_TratamientoDosis) Then
             ms_MensajeError = oEnfermeria_TratamientoDosis.MensajeError: GoTo ErrCeAgr
           End If
           
           'Consultar si la dosificacion esta completa para cambiar estado
            If Not IsNull(oRsEnfermeriaTratamientoDosis.Fields!Dosis) Then
                 Set orsTemp = Enfermeria_ConsultarTotalDosificadoMedicamento(oConexion, oRsEnfermeriaTratamientoDosis.Fields!idReceta, oRsEnfermeriaTratamientoDosis.Fields!idItem)
                 If orsTemp.RecordCount > 0 Then
                     If orsTemp.Fields!CantidadDespachada = orsTemp.Fields!CantidadDosificada Then
                         If Not Enfermeria_RecDetalleModificarEstadoConcluido(oConexion, oRsEnfermeriaTratamientoDosis.Fields!idReceta, oRsEnfermeriaTratamientoDosis.Fields!idItem) Then
                         End If
                     End If
                 End If
            Else
                 If Not IsNull(oRsEnfermeriaTratamientoDosis.Fields!DatoProrenata) Then
                     Set orsTemp = Enfermeria_ConsultarTotalDosificadoProrenata(oConexion, oRsEnfermeriaTratamientoDosis.Fields!idReceta, oRsEnfermeriaTratamientoDosis.Fields!idItem)
                     If orsTemp.RecordCount > 0 Then
                         If orsTemp.Fields!CantidadDespachada <= orsTemp.Fields!CantidadDosificada Then
                             If Not Enfermeria_RecDetalleModificarEstadoConcluido(oConexion, oRsEnfermeriaTratamientoDosis.Fields!idReceta, oRsEnfermeriaTratamientoDosis.Fields!idItem) Then
                             End If
                         End If
                     End If
                 End If
            End If
            oRsEnfermeriaTratamientoDosis.MoveNext
        Loop
    End If
    
    

    EnfermeriaVisitasAgregar = True
    oConexion.Close
ErrCeAgr:
'Resume
    If EnfermeriaVisitasAgregar = True Then
       oConexion.CommitTrans
    Else
       oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oEnfermeria_Visitas = Nothing
    Set oEnfermeria_Variables = Nothing
    Set oDOEnfermeria_Variables = Nothing
    Set oEnfermeria_ValoresCombo = Nothing
    Set oDoEnfermeria_ValoresCombo = Nothing
End Function

Function EnfermeriaVisitasModificar(oDOEnfermeria_Visitas As DOEnfermeria_Visitas, oRsDatosEnfermeriaVariables As Recordset, oRsDatosEnfDatosValoresCombo As Recordset, oRsEnfermeriaTratamientoDosis As Recordset) As Boolean
    On Error GoTo ErrCeAgr
    Dim oConexion As New ADODB.Connection
    Dim oEnfermeria_Visitas As New Enfermeria_Visitas
    Dim oEnfermeria_Variables As New Enfermeria_Variables, oDOEnfermeria_Variables As New DOEnfermeria_Variables
    Dim oEnfermeria_ValoresCombo As New Enfermeria_ValoresCombo, oDoEnfermeria_ValoresCombo As New DoEnfermeria_ValoresCombo
    Dim oEnfermeria_TratamientoDosis As New Enfermeria_TratamientoDosis, oDOEnfermeria_TratamientoDosis As New DOEnfermeria_TratamientoDosis
    Dim orsTemp As New ADODB.Recordset
    
    EnfermeriaVisitasModificar = False
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    '
    oConexion.BeginTrans
    '
    Set oEnfermeria_Visitas.Conexion = oConexion
    Set oEnfermeria_Variables.Conexion = oConexion
    Set oEnfermeria_ValoresCombo.Conexion = oConexion
    Set oEnfermeria_TratamientoDosis.Conexion = oConexion

    If Not oEnfermeria_Visitas.modificar(oDOEnfermeria_Visitas) Then
        ms_MensajeError = oEnfermeria_Visitas.MensajeError: GoTo ErrCeAgr
    End If
    'Modificar los datos de las variables
    If oRsDatosEnfermeriaVariables.RecordCount > 0 Then
        oRsDatosEnfermeriaVariables.MoveFirst
        Do While Not oRsDatosEnfermeriaVariables.EOF
           oDOEnfermeria_Variables.idCuentaAtencion = oDOEnfermeria_Visitas.idCuentaAtencion
           oDOEnfermeria_Variables.IdUsuarioAuditoria = oDOEnfermeria_Visitas.IdUsuarioAuditoria
           oDOEnfermeria_Variables.IdVisita = oDOEnfermeria_Visitas.IdVisita
           oDOEnfermeria_Variables.IdVariable = oRsDatosEnfermeriaVariables.Fields!IdVariable
           oDOEnfermeria_Variables.VariableDato = IIf(IsNull(oRsDatosEnfermeriaVariables.Fields!VariableDato), "", oRsDatosEnfermeriaVariables.Fields!VariableDato)
           If Not oEnfermeria_Variables.modificar(oDOEnfermeria_Variables) Then
             ms_MensajeError = oEnfermeria_Variables.MensajeError: GoTo ErrCeAgr
           End If
           oRsDatosEnfermeriaVariables.MoveNext
        Loop
    End If
    
    If Not oEnfermeria_ValoresCombo.Eliminar(oDOEnfermeria_Visitas.idCuentaAtencion, oDOEnfermeria_Visitas.IdVisita) Then
      ms_MensajeError = oEnfermeria_ValoresCombo.MensajeError: GoTo ErrCeAgr
    End If
    'Inserta los datos de la variable combo
    If oRsDatosEnfDatosValoresCombo.RecordCount > 0 Then
        oRsDatosEnfDatosValoresCombo.MoveFirst
        Do While Not oRsDatosEnfDatosValoresCombo.EOF
           oDoEnfermeria_ValoresCombo.idCuentaAtencion = oDOEnfermeria_Visitas.idCuentaAtencion
           oDoEnfermeria_ValoresCombo.IdVisita = oDOEnfermeria_Visitas.IdVisita
           oDoEnfermeria_ValoresCombo.IdVariable = oRsDatosEnfDatosValoresCombo.Fields!IdVariable
           oDoEnfermeria_ValoresCombo.IdValorCombo = oRsDatosEnfDatosValoresCombo.Fields!IdValorCombo
           oDoEnfermeria_ValoresCombo.IdUsuarioAuditoria = oDOEnfermeria_Visitas.IdUsuarioAuditoria
           If Not oEnfermeria_ValoresCombo.Insertar(oDoEnfermeria_ValoresCombo) Then
             ms_MensajeError = oEnfermeria_ValoresCombo.MensajeError: GoTo ErrCeAgr
           End If
           oRsDatosEnfDatosValoresCombo.MoveNext
        Loop
    End If
    
     'Inserta los datos de medicamentos de la visita
    If oRsEnfermeriaTratamientoDosis.RecordCount > 0 Then
        oRsEnfermeriaTratamientoDosis.MoveFirst
        Do While Not oRsEnfermeriaTratamientoDosis.EOF
           oDOEnfermeria_TratamientoDosis.idCuentaAtencion = oRsEnfermeriaTratamientoDosis.Fields!idCuentaAtencion
           oDOEnfermeria_TratamientoDosis.IdVisita = oRsEnfermeriaTratamientoDosis.Fields!IdVisita
           oDOEnfermeria_TratamientoDosis.IdDiaVisita = oRsEnfermeriaTratamientoDosis.Fields!IdDiaVisita
           oDOEnfermeria_TratamientoDosis.idReceta = oRsEnfermeriaTratamientoDosis.Fields!idReceta
           oDOEnfermeria_TratamientoDosis.idItem = oRsEnfermeriaTratamientoDosis.Fields!idItem
           oDOEnfermeria_TratamientoDosis.Dosis = 0
           oDOEnfermeria_TratamientoDosis.DatoProrenata = 0
           If Not IsNull(oRsEnfermeriaTratamientoDosis.Fields!Dosis) Then
                oDOEnfermeria_TratamientoDosis.Dosis = oRsEnfermeriaTratamientoDosis.Fields!Dosis
           Else
                If Not IsNull(oRsEnfermeriaTratamientoDosis.Fields!DatoProrenata) Then
                    oDOEnfermeria_TratamientoDosis.DatoProrenata = oRsEnfermeriaTratamientoDosis.Fields!DatoProrenata
                End If
           End If
           If Not oEnfermeria_TratamientoDosis.modificar(oDOEnfermeria_TratamientoDosis) Then
             ms_MensajeError = oEnfermeria_TratamientoDosis.MensajeError: GoTo ErrCeAgr
           End If
           
           'Consultar si la dosificacion esta completa para cambiar estado
            If Not IsNull(oRsEnfermeriaTratamientoDosis.Fields!Dosis) Then
                 Set orsTemp = Enfermeria_ConsultarTotalDosificadoMedicamento(oConexion, oRsEnfermeriaTratamientoDosis.Fields!idReceta, oRsEnfermeriaTratamientoDosis.Fields!idItem)
                 If orsTemp.RecordCount > 0 Then
                     If orsTemp.Fields!CantidadDespachada = orsTemp.Fields!CantidadDosificada Then
                         If Not Enfermeria_RecDetalleModificarEstadoConcluido(oConexion, oRsEnfermeriaTratamientoDosis.Fields!idReceta, oRsEnfermeriaTratamientoDosis.Fields!idItem) Then
                         End If
                     End If
                 End If
            Else
                 If Not IsNull(oRsEnfermeriaTratamientoDosis.Fields!DatoProrenata) Then
                     Set orsTemp = Enfermeria_ConsultarTotalDosificadoProrenata(oConexion, oRsEnfermeriaTratamientoDosis.Fields!idReceta, oRsEnfermeriaTratamientoDosis.Fields!idItem)
                     If orsTemp.RecordCount > 0 Then
                         If orsTemp.Fields!CantidadDespachada <= orsTemp.Fields!CantidadDosificada Then
                             If Not Enfermeria_RecDetalleModificarEstadoConcluido(oConexion, oRsEnfermeriaTratamientoDosis.Fields!idReceta, oRsEnfermeriaTratamientoDosis.Fields!idItem) Then
                             End If
                         End If
                     End If
                 End If
            End If
           oRsEnfermeriaTratamientoDosis.MoveNext
        Loop
    End If

    EnfermeriaVisitasModificar = True
    oConexion.Close
ErrCeAgr:
'Resume
    If EnfermeriaVisitasModificar = True Then
       oConexion.CommitTrans
    Else
       oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oEnfermeria_Visitas = Nothing
    Set oEnfermeria_Variables = Nothing
    Set oDOEnfermeria_Variables = Nothing
    Set oEnfermeria_ValoresCombo = Nothing
    Set oDoEnfermeria_ValoresCombo = Nothing
End Function

Function EnfermeriaVisitasEliminar(lnIdCuentaAtencion As Long, lnIdVisita As Integer) As Boolean
    On Error GoTo ErrCeAgr
    Dim oConexion As New ADODB.Connection
    Dim oEnfermeria_Visitas As New Enfermeria_Visitas
       
    EnfermeriaVisitasEliminar = False
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    '
    oConexion.BeginTrans
    '
    Set oEnfermeria_Visitas.Conexion = oConexion

    If Not oEnfermeria_Visitas.Eliminar(lnIdCuentaAtencion, lnIdVisita) Then
        ms_MensajeError = oEnfermeria_Visitas.MensajeError: GoTo ErrCeAgr
    End If

    EnfermeriaVisitasEliminar = True
    oConexion.Close
ErrCeAgr:
'Resume
    If EnfermeriaVisitasEliminar = True Then
       oConexion.CommitTrans
    Else
       oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oEnfermeria_Visitas = Nothing
End Function

Function Enfermeria_RecetasDespachadasPorCuenta(oConexion As ADODB.Connection, lnIdCuentaAtencion As Long, ldFechaVisita As String) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As New ADODB.Parameter
    Set Enfermeria_RecetasDespachadasPorCuenta = Nothing
    
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "Enfermeria_RecetasDespachadasPorCuenta"
         Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@FechaVisita", adDBTimeStamp, adParamInput, 0, ldFechaVisita): .Parameters.Append oParameter
         Set oRsTmp = .Execute
    End With
    Set Enfermeria_RecetasDespachadasPorCuenta = oRsTmp

    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Function Enfermeria_ConsultarMedicamentosVisita(oConexion As ADODB.Connection, lnIdCuentaAtencion As Long, ldIdVisita As Integer) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As New ADODB.Parameter
    Set Enfermeria_ConsultarMedicamentosVisita = Nothing
    
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "Enfermeria_ConsultarMedicamentosVisita"
         Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@IdVisita", adInteger, adParamInput, 0, ldIdVisita): .Parameters.Append oParameter
         Set oRsTmp = .Execute
    End With
    Set Enfermeria_ConsultarMedicamentosVisita = oRsTmp

    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Function Enfermeria_ConsultarMedicamentoDosisDia(oConexion As ADODB.Connection, lnIdCuentaAtencion As Long, lnIdReceta As Long, _
                                                lnIdItem As Long, lnIdDiaVIsita As Integer, lnDosis As Integer) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As New ADODB.Parameter
    Set Enfermeria_ConsultarMedicamentoDosisDia = Nothing
    
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "Enfermeria_ConsultarMedicamentoDosisDia"
         Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@IdItem", adInteger, adParamInput, 0, lnIdItem): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@IdDiaVIsita", adInteger, adParamInput, 0, lnIdDiaVIsita): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@Dosis", adInteger, adParamInput, 0, lnDosis): .Parameters.Append oParameter
         Set oRsTmp = .Execute
    End With
    Set Enfermeria_ConsultarMedicamentoDosisDia = oRsTmp

    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Function Enfermeria_ConsultarTotalDosificadoProrenata(oConexion As ADODB.Connection, lnIdReceta As Long, lnIdItem As Long) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As New ADODB.Parameter
    Set Enfermeria_ConsultarTotalDosificadoProrenata = Nothing
    
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "Enfermeria_ConsultarTotalDosificadoProrenata"
         Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@IdItem", adInteger, adParamInput, 0, lnIdItem): .Parameters.Append oParameter
         Set oRsTmp = .Execute
    End With
    Set Enfermeria_ConsultarTotalDosificadoProrenata = oRsTmp

    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Function Enfermeria_ConsultarTotalDosificadoMedicamento(oConexion As ADODB.Connection, lnIdReceta As Long, lnIdItem As Long) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As New ADODB.Parameter
    Set Enfermeria_ConsultarTotalDosificadoMedicamento = Nothing
    
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "Enfermeria_ConsultarTotalDosificadoMedicamento"
         Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@IdItem", adInteger, adParamInput, 0, lnIdItem): .Parameters.Append oParameter
         Set oRsTmp = .Execute
    End With
    Set Enfermeria_ConsultarTotalDosificadoMedicamento = oRsTmp

    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function


Function Enfermeria_RecDetalleModificarEstadoConcluido(oConexion As ADODB.Connection, lnIdReceta As Long, lnIdItem As Long) As Boolean
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As New ADODB.Parameter
    Enfermeria_RecDetalleModificarEstadoConcluido = False
    
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "Enfermeria_RecDetalleModificarEstadoConcluido"
         Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@IdItem", adInteger, adParamInput, 0, lnIdItem): .Parameters.Append oParameter
         .Execute
    End With
    Enfermeria_RecDetalleModificarEstadoConcluido = True

    Set oCommand = Nothing
    Set oParameter = Nothing
End Function


Function AtencionesSeleccionarXanioIngreso(lcAnio As String) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarXanioIngreso"
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, Val(lcAnio)): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarXanioIngreso = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesSeleccionarXanio(lcAnio As String) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarXanio"
        Set oParameter = .CreateParameter("@lnAnio", adInteger, adParamInput, 0, Val(lcAnio)): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarXanio = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'mgaray
Function PacientesDatosAdicionalesPersonalesAgregar(oDOPacienteAdd As DoPacienteDatosAdd, _
        oConexion As ADODB.Connection) As Boolean
    'Dim oConexion As New ADODB.Connection
    If Not (oDOPacienteAdd Is Nothing) Then
        Dim oPacienteAdd As New PacientesDatosAdd
        
        
        
        Set oPacienteAdd.Conexion = oConexion
        PacientesDatosAdicionalesPersonalesAgregar = oPacienteAdd.DatosPersonalesAgregar(oDOPacienteAdd)
        If PacientesDatosAdicionalesPersonalesAgregar = False Then
            ms_MensajeError = oPacienteAdd.MensajeError
        End If
        Set oPacienteAdd = Nothing
    End If
End Function


Function PacientesSeleccionarMayor30Menor60(lcAnioCreacionHC As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim ms_MensajeError  As String
Dim oParameter As ADODB.Parameter
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesSeleccionarMayor30Menor60"
        Set oParameter = .CreateParameter("@lnAnio", adInteger, adParamInput, 0, Val(lcAnioCreacionHC)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set PacientesSeleccionarMayor30Menor60 = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function PacientesSeleccionarMayor18Menor30(lcAnioCreacionHC As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim ms_MensajeError  As String
Dim oParameter As ADODB.Parameter
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesSeleccionarMayor18Menor30"
        Set oParameter = .CreateParameter("@lnAnio", adInteger, adParamInput, 0, Val(lcAnioCreacionHC)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set PacientesSeleccionarMayor18Menor30 = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function PacientesSeleccionarMayor12Menor18(lcAnioCreacionHC As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim ms_MensajeError  As String
Dim oParameter As ADODB.Parameter
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesSeleccionarMayor12Menor18"
        Set oParameter = .CreateParameter("@lnAnio", adInteger, adParamInput, 0, Val(lcAnioCreacionHC)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set PacientesSeleccionarMayor12Menor18 = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function PacientesSelecionarMayor10Menor12(lcAnioCreacionHC As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim ms_MensajeError  As String
Dim oParameter As ADODB.Parameter
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesSelecionarMayor10Menor12"
        Set oParameter = .CreateParameter("@lnAnio", adInteger, adParamInput, 0, Val(lcAnioCreacionHC)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set PacientesSelecionarMayor10Menor12 = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function PacientesSeleccionarMayor5Menor10(lcAnioCreacionHC As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim ms_MensajeError  As String
Dim oParameter As ADODB.Parameter
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesSeleccionarMayor5Menor10"
        Set oParameter = .CreateParameter("@lnAnio", adInteger, adParamInput, 0, Val(lcAnioCreacionHC)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set PacientesSeleccionarMayor5Menor10 = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function PacientesSeleccionarMayor1menor5(lcAnioCreacionHC As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim ms_MensajeError  As String
Dim oParameter As ADODB.Parameter
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesSeleccionarMayor1menor5"
        Set oParameter = .CreateParameter("@lnAnio", adInteger, adParamInput, 0, Val(lcAnioCreacionHC)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set PacientesSeleccionarMayor1menor5 = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function PacientesSeleccionarMenores1anio(lcAnioCreacionHC As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim ms_MensajeError  As String
Dim oParameter As ADODB.Parameter
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesSeleccionarMenores1anio"
        Set oParameter = .CreateParameter("@lnAnio", adInteger, adParamInput, 0, Val(lcAnioCreacionHC)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set PacientesSeleccionarMenores1anio = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function PacientesSeleccionarRn(lcAnioCreacionHC As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim ms_MensajeError  As String
Dim oParameter As ADODB.Parameter
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesSeleccionarRn"
        Set oParameter = .CreateParameter("@lnAnio", adInteger, adParamInput, 0, Val(lcAnioCreacionHC)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set PacientesSeleccionarRn = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function PacientesSeleccionarMujeresEdadFertil(lcAnioCreacionHC As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim ms_MensajeError  As String
Dim oParameter As ADODB.Parameter
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesSeleccionarMujeresEdadFertil"
        Set oParameter = .CreateParameter("@lnAnio", adInteger, adParamInput, 0, Val(lcAnioCreacionHC)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set PacientesSeleccionarMujeresEdadFertil = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function PacientesSeleccionarTotalFemenino(lcAnioCreacionHC As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim ms_MensajeError  As String
Dim oParameter As ADODB.Parameter
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesSeleccionarTotalFemenino"
        Set oParameter = .CreateParameter("@lnAnio", adInteger, adParamInput, 0, Val(lcAnioCreacionHC)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set PacientesSeleccionarTotalFemenino = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function PacientesSeleccionarTotalMasculino(lcAnioCreacionHC As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesSeleccionarTotalMasculino"
        Set oParameter = .CreateParameter("@lnAnio", adInteger, adParamInput, 0, Val(lcAnioCreacionHC)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set PacientesSeleccionarTotalMasculino = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Public Function RealizarBusquedaPacienteSiNo(ByVal BusquedaRapidaHabilitada As String, _
        ByVal ApellidoPaterno As String, ByVal ApellidoMaterno As String) As Boolean
        ' ByVal PrimerNombre As String, ByVal SegundoNombre As String) As Boolean
        
    Dim RealizarBusqueda As Boolean
    RealizarBusqueda = False
    
'    If Trim(PrimerNombre) <> "" Or Trim(SegundoNombre) <> "" Then
'        RealizarBusqueda = True
'    Else
    If Trim(ApellidoPaterno) <> "" Or Trim(ApellidoMaterno) <> "" Then
        If BusquedaRapidaHabilitada = "S" Then
            If Trim(ApellidoPaterno) <> "" And Trim(ApellidoMaterno) <> "" Then
                RealizarBusqueda = True
            End If
        Else
            RealizarBusqueda = True
        End If
    End If
    
    RealizarBusquedaPacienteSiNo = RealizarBusqueda
End Function

Private Function grabarDatosPerinatalAtencionIntegral(oDoPerinatalAtencion As DoPerinatalAtencion, _
                lnIdAtencion As Long, oConexion As ADODB.Connection) As Boolean
    Call grabarDatosAtencionIntegralInmunizaciones(oDoPerinatalAtencion, lnIdAtencion, oConexion)
    Call grabarDatosAtencionIntegralTamizajes(oDoPerinatalAtencion, lnIdAtencion, oConexion)
    Call grabarDatosAtencionIntegralDesarrollo(oDoPerinatalAtencion, lnIdAtencion, oConexion)
    Call grabarDatosAtencionIntegralCrecimiento(oDoPerinatalAtencion, lnIdAtencion, oConexion)
    Call grabarDatosAtencionIntegralSuplemento(oDoPerinatalAtencion, lnIdAtencion, oConexion)
    
    Call grabarDatosPlanInmunizaciones(oDoPerinatalAtencion, lnIdAtencion, oConexion)
    Call grabarDatosPlanTamizajes(oDoPerinatalAtencion, lnIdAtencion, oConexion)
    Call grabarDatosPlanDesarrollo(oDoPerinatalAtencion, lnIdAtencion, oConexion)
    Call grabarDatosPlanCrecimiento(oDoPerinatalAtencion, lnIdAtencion, oConexion)
    Call grabarDatosPlanSuplemento(oDoPerinatalAtencion, lnIdAtencion, oConexion)
    
    grabarDatosPerinatalAtencionIntegral = True
End Function

'====================================================================================
'
'====================================================================================

Private Function getCodigoHIS(oRsTmpCodigoHis As ADODB.Recordset, lIdProducto As Long) As String
    getCodigoHIS = ""
    If oRsTmpCodigoHis.RecordCount > 0 Then
        oRsTmpCodigoHis.MoveFirst
        oRsTmpCodigoHis.Find "idProducto=" & lIdProducto
        If Not oRsTmpCodigoHis.EOF Then
            getCodigoHIS = oRsTmpCodigoHis.Fields!CodigoHIS
        End If
    End If
End Function

Private Function grabarDatosAtencionIntegralProcedimiento(oRsAtenInteProcedimientos As ADODB.Recordset, _
                oDoPerinatalAtencion As DoPerinatalAtencion, _
                lnIdAtencion As Long, oConexion As ADODB.Connection) As Boolean
    Dim GraboDatos As Boolean
    Dim oDOAtenIntePlanProcedPaciente As New DOAtenIntePlanProcedPaciente
    Dim oAtenIntePlanProcedPaciente As New AtenIntePlanProcedPaciente
    Dim cCodigoHIS As String
    
    Dim oRsTmpCodigoHis As ADODB.Recordset
    
    Set oRsTmpCodigoHis = PerinatalCatalogoCptXidModulo(oDoPerinatalAtencion.idModulo, oConexion)
    oRsTmpCodigoHis.Filter = "idLista=" & sghPerinatalListas.sighInmunizaciones
    
    GraboDatos = True
    
    If Not (oRsAtenInteProcedimientos Is Nothing) Then
        If oRsAtenInteProcedimientos.RecordCount > 0 Then
            Set oAtenIntePlanProcedPaciente.Conexion = oConexion
            oRsAtenInteProcedimientos.MoveFirst
            While oRsAtenInteProcedimientos.EOF = False
                If IsNull(oRsAtenInteProcedimientos!CodigoHIS) Then
                    cCodigoHIS = getCodigoHIS(oRsTmpCodigoHis, oRsAtenInteProcedimientos!idProducto)
                ElseIf Trim(oRsAtenInteProcedimientos!CodigoHIS) = "" Then
                    cCodigoHIS = getCodigoHIS(oRsTmpCodigoHis, oRsAtenInteProcedimientos!idProducto)
                Else
                    cCodigoHIS = oRsAtenInteProcedimientos!CodigoHIS
                End If
                oDOAtenIntePlanProcedPaciente.IdPlanProcedimientoPaciente = IIf(IsNull(oRsAtenInteProcedimientos!IdPlanProcedimientoPaciente), 0, oRsAtenInteProcedimientos!IdPlanProcedimientoPaciente)
                oDOAtenIntePlanProcedPaciente.IdPlanIntegralPaciente = oRsAtenInteProcedimientos!IdPlanIntegralPaciente
                oDOAtenIntePlanProcedPaciente.idProducto = IIf(IsNull(oRsAtenInteProcedimientos!idProducto), 0, oRsAtenInteProcedimientos!idProducto)
                oDOAtenIntePlanProcedPaciente.IdPlanAtencion = IIf(IsNull(oRsAtenInteProcedimientos!IdPlanAtencion), 0, oRsAtenInteProcedimientos!IdPlanAtencion)
                'oDOAtenIntePlanProcedPaciente.IdAtenInteItemPlan = IIf(IsNull(oRsAtenInteProcedimientos!IdAtenInteItemPlan), 0, oRsAtenInteProcedimientos!IdAtenInteItemPlan)
                oDOAtenIntePlanProcedPaciente.FechaProgramada = IIf(IsNull(oRsAtenInteProcedimientos!FechaProgramada), 0, oRsAtenInteProcedimientos!FechaProgramada)
                oDOAtenIntePlanProcedPaciente.FechaEjecucion = IIf(IsNull(oRsAtenInteProcedimientos!FechaEjecucion), 0, oRsAtenInteProcedimientos!FechaEjecucion)
                oDOAtenIntePlanProcedPaciente.NumeroDosis = oRsAtenInteProcedimientos!NumeroDosis
                oDOAtenIntePlanProcedPaciente.CodigoHIS = cCodigoHIS
                oDOAtenIntePlanProcedPaciente.idAtencion = IIf(IsNull(oRsAtenInteProcedimientos!idAtencion), 0, oRsAtenInteProcedimientos!idAtencion)
                oDOAtenIntePlanProcedPaciente.IdEstablecimiento = IIf(IsNull(oRsAtenInteProcedimientos!IdEstablecimiento), 0, oRsAtenInteProcedimientos!IdEstablecimiento)
                
                If oAtenIntePlanProcedPaciente.modificar(oDOAtenIntePlanProcedPaciente) = False Then
                    GraboDatos = False
                    Exit Function
                End If
                
                oRsAtenInteProcedimientos.MoveNext
            Wend
        End If
    End If
    grabarDatosAtencionIntegralProcedimiento = GraboDatos
End Function

Private Function grabarDatosPlanProcedimiento(oRsPlanProcedimientos As ADODB.Recordset, _
                oDoPerinatalAtencion As DoPerinatalAtencion, _
                lnIdAtencion As Long, oConexion As ADODB.Connection) As Boolean
    
    Dim GraboDatos As Boolean
    Dim oDOAtenIntePlanProcedPaciente As New DOAtenIntePlanProcedPaciente
    Dim oAtenIntePlanProcedPaciente As New AtenIntePlanProcedPaciente
    
    GraboDatos = True
    
    If Not (oRsPlanProcedimientos Is Nothing) Then
'        Set oDOAtenIntePlanDesPaciente = New DOAtenIntePlanDesPaciente
        If oRsPlanProcedimientos.RecordCount > 0 Then
            Set oAtenIntePlanProcedPaciente.Conexion = oConexion
            oRsPlanProcedimientos.MoveFirst
            While oRsPlanProcedimientos.EOF = False
            
                oDOAtenIntePlanProcedPaciente.IdPlanProcedimientoPaciente = IIf(IsNull(oRsPlanProcedimientos!IdPlanProcedimientoPaciente), 0, oRsPlanProcedimientos!IdPlanProcedimientoPaciente)
                oDOAtenIntePlanProcedPaciente.IdPlanIntegralPaciente = oRsPlanProcedimientos!IdPlanIntegralPaciente
                oDOAtenIntePlanProcedPaciente.FechaProgramada = IIf(IsNull(oRsPlanProcedimientos!FechaProgramada), 0, oRsPlanProcedimientos!FechaProgramada)
                
                If oAtenIntePlanProcedPaciente.ModificarFechaProgramada(oDOAtenIntePlanProcedPaciente) = False Then
                    GraboDatos = False
                    Exit Function
                End If
                oRsPlanProcedimientos.MoveNext
            Wend
        End If
    End If
    
    grabarDatosPlanProcedimiento = GraboDatos
End Function

Private Function grabarDatosAtencionIntegralInmunizaciones(oDoPerinatalAtencion As DoPerinatalAtencion, _
                lnIdAtencion As Long, oConexion As ADODB.Connection) As Boolean
    grabarDatosAtencionIntegralInmunizaciones = grabarDatosAtencionIntegralProcedimiento(mo_rsAtenInteInmunizaciones, _
                                                oDoPerinatalAtencion, lnIdAtencion, oConexion)
'    Dim GraboDatos As Boolean
'    Dim oDOAtenIntePlanProcedPaciente As New DOAtenIntePlanProcedPaciente
'    Dim oAtenIntePlanProcedPaciente As New AtenIntePlanProcedPaciente
'
'    GraboDatos = True
'
'    If Not (mo_rsAtenInteInmunizaciones Is Nothing) Then
'        If mo_rsAtenInteInmunizaciones.RecordCount > 0 Then
'            Set oAtenIntePlanProcedPaciente.Conexion = oConexion
'            mo_rsAtenInteInmunizaciones.MoveFirst
'            While mo_rsAtenInteInmunizaciones.EOF = False
'
'                oDOAtenIntePlanProcedPaciente.IdPlanProcedimientoPaciente = IIf(IsNull(mo_rsAtenInteInmunizaciones!IdPlanProcedimientoPaciente), 0, mo_rsAtenInteInmunizaciones!IdPlanProcedimientoPaciente)
'                oDOAtenIntePlanProcedPaciente.IdPlanIntegralPaciente = mo_rsAtenInteInmunizaciones!IdPlanIntegralPaciente
'                oDOAtenIntePlanProcedPaciente.idProducto = IIf(IsNull(mo_rsAtenInteInmunizaciones!idProducto), 0, mo_rsAtenInteInmunizaciones!idProducto)
'                oDOAtenIntePlanProcedPaciente.IdPlanAtencion = IIf(IsNull(mo_rsAtenInteInmunizaciones!IdPlanAtencion), 0, mo_rsAtenInteInmunizaciones!IdPlanAtencion)
'                'oDOAtenIntePlanProcedPaciente.IdAtenInteItemPlan = IIf(IsNull(mo_rsAtenInteInmunizaciones!IdAtenInteItemPlan), 0, mo_rsAtenInteInmunizaciones!IdAtenInteItemPlan)
'                oDOAtenIntePlanProcedPaciente.FechaProgramada = IIf(IsNull(mo_rsAtenInteInmunizaciones!FechaProgramada), 0, mo_rsAtenInteInmunizaciones!FechaProgramada)
'                oDOAtenIntePlanProcedPaciente.FechaEjecucion = IIf(IsNull(mo_rsAtenInteInmunizaciones!FechaEjecucion), 0, mo_rsAtenInteInmunizaciones!FechaEjecucion)
'                oDOAtenIntePlanProcedPaciente.NumeroDosis = mo_rsAtenInteInmunizaciones!NumeroDosis
'                oDOAtenIntePlanProcedPaciente.CodigoHIS = IIf(IsNull(mo_rsAtenInteInmunizaciones!CodigoHIS), "", mo_rsAtenInteInmunizaciones!CodigoHIS)
'                oDOAtenIntePlanProcedPaciente.idAtencion = IIf(IsNull(mo_rsAtenInteInmunizaciones!idAtencion), 0, mo_rsAtenInteInmunizaciones!idAtencion)
'                oDOAtenIntePlanProcedPaciente.IdEstablecimiento = IIf(IsNull(mo_rsAtenInteInmunizaciones!IdEstablecimiento), 0, mo_rsAtenInteInmunizaciones!IdEstablecimiento)
'
'                If oAtenIntePlanProcedPaciente.modificar(oDOAtenIntePlanProcedPaciente) = False Then
'                    GraboDatos = False
'                    Exit Function
'                End If
'
'                mo_rsAtenInteInmunizaciones.MoveNext
'            Wend
'        End If
'    End If
'    grabarDatosAtencionIntegralInmunizaciones = GraboDatos
End Function

Private Function grabarDatosPlanInmunizaciones(oDoPerinatalAtencion As DoPerinatalAtencion, _
                lnIdAtencion As Long, oConexion As ADODB.Connection) As Boolean
    grabarDatosPlanInmunizaciones = grabarDatosPlanProcedimiento(mo_rsPlanInmunizaciones, oDoPerinatalAtencion, _
                            lnIdAtencion, oConexion)
'    Dim GraboDatos As Boolean
'    Dim oDOAtenIntePlanProcedPaciente As New DOAtenIntePlanProcedPaciente
'    Dim oAtenIntePlanProcedPaciente As New AtenIntePlanProcedPaciente
'
'    GraboDatos = True
'
'    If Not (mo_rsPlanInmunizaciones Is Nothing) Then
''        Set oDOAtenIntePlanDesPaciente = New DOAtenIntePlanDesPaciente
'        If mo_rsPlanInmunizaciones.RecordCount > 0 Then
'            Set oAtenIntePlanProcedPaciente.Conexion = oConexion
'            mo_rsPlanInmunizaciones.MoveFirst
'            While mo_rsPlanInmunizaciones.EOF = False
'
'                oDOAtenIntePlanProcedPaciente.IdPlanDesarrolloPaciente = IIf(IsNull(mo_rsPlanInmunizaciones!IdPlanDesarrolloPaciente), 0, mo_rsPlanInmunizaciones!IdPlanDesarrolloPaciente)
'                oDOAtenIntePlanProcedPaciente.IdPlanIntegralPaciente = mo_rsPlanInmunizaciones!IdPlanIntegralPaciente
'                oDOAtenIntePlanProcedPaciente.FechaProgramada = IIf(IsNull(mo_rsPlanInmunizaciones!FechaProgramada), 0, mo_rsPlanInmunizaciones!FechaProgramada)
'
'                If oAtenIntePlanProcedPaciente.ModificarFechaProgramada(oDOAtenIntePlanProcedPaciente) = False Then
'                    GraboDatos = False
'                    Exit Function
'                End If
'                mo_rsPlanInmunizaciones.MoveNext
'            Wend
'        End If
'    End If
'
'    grabarDatosPlanInmunizaciones = GraboDatos
End Function

'====================================================================================
'
'====================================================================================
Private Function grabarDatosAtencionIntegralTamizajes(oDoPerinatalAtencion As DoPerinatalAtencion, _
                lnIdAtencion As Long, oConexion As ADODB.Connection) As Boolean
    grabarDatosAtencionIntegralTamizajes = grabarDatosAtencionIntegralProcedimiento(mo_rsAtenInteTamizaje, _
                                                oDoPerinatalAtencion, lnIdAtencion, oConexion)
    'grabarDatosAtencionIntegralTamizajes = True
End Function

Private Function grabarDatosPlanTamizajes(oDoPerinatalAtencion As DoPerinatalAtencion, _
                lnIdAtencion As Long, oConexion As ADODB.Connection) As Boolean
    grabarDatosPlanTamizajes = grabarDatosPlanProcedimiento(mo_rsPlanTamizaje, oDoPerinatalAtencion, _
                            lnIdAtencion, oConexion)
    'grabarDatosPlanTamizajes = True
End Function

'====================================================================================
'
'====================================================================================


Private Function grabarAtencionIntegralDesarrolloCabecera(oDoPerinatalAtencion As DoPerinatalAtencion, _
                lnIdAtencion As Long, oConexion As ADODB.Connection) As Boolean
    Dim GraboDatos As Boolean
    GraboDatos = False
    
    If Not (mo_rsAtenInteDesarrolloCabecera Is Nothing) Then
        Dim oDOAtenIntePlanDesPaciente As New DOAtenIntePlanDesPaciente
        Dim oAtenIntePlanDesPaciente As New AtenIntePlanDesPaciente
        Set oAtenIntePlanDesPaciente.Conexion = oConexion
        
        oDOAtenIntePlanDesPaciente.IdPlanDesarrolloPaciente = IIf(IsNull(mo_rsAtenInteDesarrolloCabecera!IdPlanDesarrolloPaciente), 0, mo_rsAtenInteDesarrolloCabecera!IdPlanDesarrolloPaciente)
        oDOAtenIntePlanDesPaciente.IdPlanIntegralPaciente = mo_rsAtenInteDesarrolloCabecera!IdPlanIntegralPaciente
        oDOAtenIntePlanDesPaciente.Evaluacion = IIf(IsNull(mo_rsAtenInteDesarrolloCabecera!Evaluacion), 0, mo_rsAtenInteDesarrolloCabecera!Evaluacion)
        oDOAtenIntePlanDesPaciente.IdPlanAtencion = IIf(IsNull(mo_rsAtenInteDesarrolloCabecera!IdPlanAtencion), 0, mo_rsAtenInteDesarrolloCabecera!IdPlanAtencion)
        'oDOAtenIntePlanDesPaciente.IdAtenInteItemPlan = IIf(IsNull(mo_rsAtenInteDesarrolloCabecera!IdAtenInteItemPlan), 0, mo_rsAtenInteDesarrolloCabecera!IdAtenInteItemPlan)
        oDOAtenIntePlanDesPaciente.FechaProgramada = IIf(IsNull(mo_rsAtenInteDesarrolloCabecera!FechaProgramada), 0, mo_rsAtenInteDesarrolloCabecera!FechaProgramada)
        oDOAtenIntePlanDesPaciente.FechaEjecucion = IIf(IsNull(mo_rsAtenInteDesarrolloCabecera!FechaEjecucion), 0, mo_rsAtenInteDesarrolloCabecera!FechaEjecucion)
        oDOAtenIntePlanDesPaciente.NumeroSesion = IIf(IsNull(mo_rsAtenInteDesarrolloCabecera!NumeroSesion), 0, mo_rsAtenInteDesarrolloCabecera!NumeroSesion)
        oDOAtenIntePlanDesPaciente.idAtencion = IIf(IsNull(mo_rsAtenInteDesarrolloCabecera!idAtencion), 0, mo_rsAtenInteDesarrolloCabecera!idAtencion)
        oDOAtenIntePlanDesPaciente.IdEstablecimiento = IIf(IsNull(mo_rsAtenInteDesarrolloCabecera!IdEstablecimiento), 0, mo_rsAtenInteDesarrolloCabecera!IdEstablecimiento)
       
        If oAtenIntePlanDesPaciente.modificar(oDOAtenIntePlanDesPaciente) = True Then
            GraboDatos = True
        End If
    End If
    grabarAtencionIntegralDesarrolloCabecera = GraboDatos
End Function

'mgaray20141012
Public Function grabarAtencionIntegralDesarrolloVencidos(lIdAtencion As Long, _
            oDOAtenIntePlanDesPaciente As DOAtenIntePlanDesPaciente, _
            oRsItemDesarrollo As ADODB.Recordset)
On Error GoTo miError
    Dim GraboDatos As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oAtenIntePlanDesPaciente As New AtenIntePlanDesPaciente
    Dim oRsDesarrolloCabecera As ADODB.Recordset
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    Set oAtenIntePlanDesPaciente.Conexion = oConexion
    Set oRsDesarrolloCabecera = oAtenIntePlanDesPaciente.setRsSeleccionarPorId(oDOAtenIntePlanDesPaciente)
    
    If Not (oRsDesarrolloCabecera Is Nothing) Then
        oRsDesarrolloCabecera.Fields!FechaEjecucion = oDOAtenIntePlanDesPaciente.FechaEjecucion
        oRsDesarrolloCabecera.Fields!IdEstablecimiento = IIf(oDOAtenIntePlanDesPaciente.IdEstablecimiento = 0, Null, oDOAtenIntePlanDesPaciente.IdEstablecimiento)
        oRsDesarrolloCabecera.Fields!Evaluacion = oDOAtenIntePlanDesPaciente.Evaluacion
    End If
    'hacer refactory de esta funcionalidad
    Set mo_rsAtenInteDesarrollo = oRsItemDesarrollo
    Set mo_rsAtenInteDesarrolloCabecera = oRsDesarrolloCabecera
    
    GraboDatos = grabarDatosAtencionIntegralDesarrollo(Nothing, 0, oConexion)
    
    
    grabarAtencionIntegralDesarrolloVencidos = GraboDatos
     
    oConexion.Close
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbInformation, "Error"
    End If
    Set oConexion = Nothing
End Function

Private Function grabarDatosAtencionIntegralDesarrollo(oDoPerinatalAtencion As DoPerinatalAtencion, _
                lnIdAtencion As Long, oConexion As ADODB.Connection) As Boolean
    
    Dim GraboDatos As Boolean
    Dim oDOAtenIntePlanDesPacienteDet As New DOAtenIntePlanDesPacienteDet
    Dim oAtenIntePlanDesPacienteDet As New AtenIntePlanDesPacienteDet
    
    GraboDatos = True
    
    If Not (mo_rsAtenInteDesarrollo Is Nothing) Then
        
        If mo_rsAtenInteDesarrollo.RecordCount > 0 Then
            If grabarAtencionIntegralDesarrolloCabecera(oDoPerinatalAtencion, lnIdAtencion, oConexion) = True Then
                Set oAtenIntePlanDesPacienteDet.Conexion = oConexion
                mo_rsAtenInteDesarrollo.MoveFirst
                While mo_rsAtenInteDesarrollo.EOF = False
                
                    oDOAtenIntePlanDesPacienteDet.IdPlanDesarrolloPaciente = IIf(IsNull(mo_rsAtenInteDesarrollo!IdPlanDesarrolloPaciente), 0, mo_rsAtenInteDesarrollo!IdPlanDesarrolloPaciente)
                    oDOAtenIntePlanDesPacienteDet.IdPlanIntegralPaciente = mo_rsAtenInteDesarrollo!IdPlanIntegralPaciente
                    oDOAtenIntePlanDesPacienteDet.IdItemDesarrollo = IIf(IsNull(mo_rsAtenInteDesarrollo!IdItemDesarrollo), 0, mo_rsAtenInteDesarrollo!IdItemDesarrollo)
                    oDOAtenIntePlanDesPacienteDet.OrdenItem = IIf(IsNull(mo_rsAtenInteDesarrollo!OrdenItem), 0, mo_rsAtenInteDesarrollo!OrdenItem)
                    oDOAtenIntePlanDesPacienteDet.EjecutaAccion = IIf(IsNull(mo_rsAtenInteDesarrollo!EjecutaAccion), 0, mo_rsAtenInteDesarrollo!EjecutaAccion)
                    oDOAtenIntePlanDesPacienteDet.RespondioEjecutaAccion = IIf(IsNull(mo_rsAtenInteDesarrollo!EjecutaAccion), False, True)
                    If oAtenIntePlanDesPacienteDet.modificar(oDOAtenIntePlanDesPacienteDet) = False Then
                        GraboDatos = False
                        Exit Function
                    End If
                    mo_rsAtenInteDesarrollo.MoveNext
                Wend
            End If
        End If
    End If
    
    grabarDatosAtencionIntegralDesarrollo = GraboDatos
End Function

Private Function grabarDatosPlanDesarrollo(oDoPerinatalAtencion As DoPerinatalAtencion, _
                lnIdAtencion As Long, oConexion As ADODB.Connection) As Boolean
    
    Dim GraboDatos As Boolean
    Dim oDOAtenIntePlanDesPaciente As New DOAtenIntePlanDesPaciente
    Dim oAtenIntePlanDesPaciente As New AtenIntePlanDesPaciente
    
    GraboDatos = True
    
    If Not (mo_rsPlanDesarrollo Is Nothing) Then
'        Set oDOAtenIntePlanDesPaciente = New DOAtenIntePlanDesPaciente
        If mo_rsPlanDesarrollo.RecordCount > 0 Then
            Set oAtenIntePlanDesPaciente.Conexion = oConexion
            mo_rsPlanDesarrollo.MoveFirst
            While mo_rsPlanDesarrollo.EOF = False
            
                oDOAtenIntePlanDesPaciente.IdPlanDesarrolloPaciente = IIf(IsNull(mo_rsPlanDesarrollo!IdPlanDesarrolloPaciente), 0, mo_rsPlanDesarrollo!IdPlanDesarrolloPaciente)
                oDOAtenIntePlanDesPaciente.IdPlanIntegralPaciente = mo_rsPlanDesarrollo!IdPlanIntegralPaciente
                oDOAtenIntePlanDesPaciente.FechaProgramada = IIf(IsNull(mo_rsPlanDesarrollo!FechaProgramada), 0, mo_rsPlanDesarrollo!FechaProgramada)
                
                If oAtenIntePlanDesPaciente.ModificarFechaProgramada(oDOAtenIntePlanDesPaciente) = False Then
                    GraboDatos = False
                    Exit Function
                End If
                mo_rsPlanDesarrollo.MoveNext
            Wend
        End If
    End If
    
    grabarDatosPlanDesarrollo = GraboDatos
End Function

'====================================================================================
'
'====================================================================================

Private Function grabarDatosAtencionIntegralCrecimiento(oDoPerinatalAtencion As DoPerinatalAtencion, _
                lnIdAtencion As Long, oConexion As ADODB.Connection) As Boolean
    grabarDatosAtencionIntegralCrecimiento = True
End Function

Private Function grabarDatosPlanCrecimiento(oDoPerinatalAtencion As DoPerinatalAtencion, _
                lnIdAtencion As Long, oConexion As ADODB.Connection) As Boolean
    grabarDatosPlanCrecimiento = True
End Function

'====================================================================================
'
'====================================================================================

Private Function grabarDatosAtencionIntegralSuplemento(oDoPerinatalAtencion As DoPerinatalAtencion, _
                lnIdAtencion As Long, oConexion As ADODB.Connection) As Boolean
    Dim GraboDatos As Boolean
    Dim oDOAtenIntePlanSuplemPaciente As New DOAtenIntePlanSuplemPaciente ' DoAtenIntePlanSuplPaciente
    Dim oAtenIntePlanSuplPaciente  As New AtenIntePlanSuplPaciente ' AtenIntePlanProcedPaciente
    
    GraboDatos = True
    
    If Not (mo_rsAtenInteSuplemento Is Nothing) Then
        If mo_rsAtenInteSuplemento.RecordCount > 0 Then
            Set oAtenIntePlanSuplPaciente.Conexion = oConexion
            mo_rsAtenInteSuplemento.MoveFirst
            While mo_rsAtenInteSuplemento.EOF = False
                oDOAtenIntePlanSuplemPaciente.IdPlanSuplementoPaciente = IIf(IsNull(mo_rsAtenInteSuplemento!IdPlanSuplementoPaciente), 0, mo_rsAtenInteSuplemento!IdPlanSuplementoPaciente)
                oDOAtenIntePlanSuplemPaciente.IdPlanIntegralPaciente = mo_rsAtenInteSuplemento!IdPlanIntegralPaciente
                oDOAtenIntePlanSuplemPaciente.idProducto = IIf(IsNull(mo_rsAtenInteSuplemento!idProducto), 0, mo_rsAtenInteSuplemento!idProducto)
                oDOAtenIntePlanSuplemPaciente.IdPlanAtencion = IIf(IsNull(mo_rsAtenInteSuplemento!IdPlanAtencion), 0, mo_rsAtenInteSuplemento!IdPlanAtencion)
                oDOAtenIntePlanSuplemPaciente.FechaProgramada = IIf(IsNull(mo_rsAtenInteSuplemento!FechaProgramada), 0, mo_rsAtenInteSuplemento!FechaProgramada)
                oDOAtenIntePlanSuplemPaciente.FechaEjecucion = IIf(IsNull(mo_rsAtenInteSuplemento!FechaEjecucion), 0, mo_rsAtenInteSuplemento!FechaEjecucion)
                oDOAtenIntePlanSuplemPaciente.NumeroDosis = mo_rsAtenInteSuplemento!NumeroDosis
                oDOAtenIntePlanSuplemPaciente.idAtencion = IIf(IsNull(mo_rsAtenInteSuplemento!idAtencion), 0, mo_rsAtenInteSuplemento!idAtencion)
                oDOAtenIntePlanSuplemPaciente.IdEstablecimiento = IIf(IsNull(mo_rsAtenInteSuplemento!IdEstablecimiento), 0, mo_rsAtenInteSuplemento!IdEstablecimiento)
                
                If oAtenIntePlanSuplPaciente.modificar(oDOAtenIntePlanSuplemPaciente) = False Then
                    GraboDatos = False
                    Exit Function
                End If
                
                mo_rsAtenInteSuplemento.MoveNext
            Wend
        End If
    End If
    grabarDatosAtencionIntegralSuplemento = GraboDatos
End Function

Private Function grabarDatosPlanSuplemento(oDoPerinatalAtencion As DoPerinatalAtencion, _
                lnIdAtencion As Long, oConexion As ADODB.Connection) As Boolean
    Dim GraboDatos As Boolean
    Dim oDOAtenIntePlanSuplemPaciente As New DOAtenIntePlanSuplemPaciente
    Dim oAtenIntePlanSuplPaciente As New AtenIntePlanSuplPaciente
    
    GraboDatos = True
    
    If Not (mo_rsPlanSuplemento Is Nothing) Then
'        Set oDOAtenIntePlanDesPaciente = New DOAtenIntePlanDesPaciente
        If mo_rsPlanSuplemento.RecordCount > 0 Then
            Set oAtenIntePlanSuplPaciente.Conexion = oConexion
            mo_rsPlanSuplemento.MoveFirst
            While mo_rsPlanSuplemento.EOF = False
            
                oDOAtenIntePlanSuplemPaciente.IdPlanSuplementoPaciente = IIf(IsNull(mo_rsPlanSuplemento!IdPlanSuplementoPaciente), 0, mo_rsPlanSuplemento!IdPlanSuplementoPaciente)
                oDOAtenIntePlanSuplemPaciente.IdPlanIntegralPaciente = mo_rsPlanSuplemento!IdPlanIntegralPaciente
                oDOAtenIntePlanSuplemPaciente.FechaProgramada = IIf(IsNull(mo_rsPlanSuplemento!FechaProgramada), 0, mo_rsPlanSuplemento!FechaProgramada)
                
                If oAtenIntePlanSuplPaciente.ModificarFechaProgramada(oDOAtenIntePlanSuplemPaciente) = False Then
                    GraboDatos = False
                    Exit Function
                End If
                mo_rsPlanSuplemento.MoveNext
            Wend
        End If
    End If
    
    grabarDatosPlanSuplemento = GraboDatos
End Function

'====================================================================================
'
'====================================================================================

Private Function eliminarDatosAtencionIntegral(oDoPerinatalAtencion As DoPerinatalAtencion, _
            lnIdAtencion As Long, oConexion As ADODB.Connection) As Boolean
            
    Dim oDoAtencion As New DOAtencion
    Dim oAtenIntePlanDesPaciente As New AtenIntePlanDesPaciente
    
    oDoAtencion.idAtencion = lnIdAtencion
    oDoAtencion.idPaciente = oDoPerinatalAtencion.idPaciente
    
    oAtenIntePlanDesPaciente.Conexion = oConexion
    eliminarDatosAtencionIntegral = oAtenIntePlanDesPaciente.EliminarAtencionIntegral(oDoAtencion)
End Function

'Actualizado 14102014
Function AdmisionBuscarPacientesAltaMedica(lcFecha As String, lnIdServicio As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AdmisionBuscarPacientesAltaMedica = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AdmisionBuscarPacientesAltaMedica"
        Set oParameter = .CreateParameter("@lcFecha", adVarChar, adParamInput, 10, lcFecha): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AdmisionBuscarPacientesAltaMedica = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'mgaray20141015
Function AtencionesDiagnosticoAgregarOModificar(oAtencionesDiagnosticos As AtencionesDiagnosticos, _
            oDOAtencionDiagnostico As DOAtencionDiagnostico) As Boolean
    Dim oRsAtencionesDiagnotico As ADODB.Recordset
    
    Set oRsAtencionesDiagnotico = oAtencionesDiagnosticos.SeleccionPorIdAtencionDiagnosticoLab(oDOAtencionDiagnostico)
    
    If oRsAtencionesDiagnotico.RecordCount = 0 Then
        AtencionesDiagnosticoAgregarOModificar = oAtencionesDiagnosticos.Insertar(oDOAtencionDiagnostico)
    Else
        oRsAtencionesDiagnotico.MoveFirst
        oDOAtencionDiagnostico.IdAtencionDiagnostico = oRsAtencionesDiagnotico.Fields!IdAtencionDiagnostico
        AtencionesDiagnosticoAgregarOModificar = oAtencionesDiagnosticos.ModificarTodosCampos(oDOAtencionDiagnostico)
    End If
End Function
'mgaray20141024
Function PerinatalModicarOrdenServicio(lIdAtencion As Long, lidOrden As Long, _
                IdUsuarioAuditoria As Long, oConexion As Connection) As Boolean
    On Error GoTo ErrCeAgr
'    Dim oConexion As New ADODB.Connection
    Dim oPerinatalAtencioncpt As New PerinatalAtencionCpt
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
    Dim oDoPerinatalAtencionCpt As New DoPerinatalAtencionCpt
    
    PerinatalModicarOrdenServicio = False
    
    oDoPerinatalAtencionCpt.idLista = sighItemPlanIntegral.Inmunizacion
    oDoPerinatalAtencionCpt.idAtencion = lIdAtencion
    oDoPerinatalAtencionCpt.IdOrden = lidOrden
    oDoPerinatalAtencionCpt.IdUsuarioAuditoria = IdUsuarioAuditoria
    
'    oConexion.CommandTimeout = 300
'    oConexion.CursorLocation = adUseClient
'    oConexion.Open sighentidades.CadenaConexion

    Set oPerinatalAtencioncpt.Conexion = oConexion
    PerinatalModicarOrdenServicio = oPerinatalAtencioncpt.ModificarOrdenServicio(oDoPerinatalAtencionCpt)
'    oConexion.Close
ErrCeAgr:
'    oConexion.Close
'    Set oConexion = Nothing
    Set oPerinatalAtencioncpt = Nothing
End Function

Function PerinatalBuscarOrdenServicioInmunizacion(lnIdAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    Dim oPerinatalAtencioncpt As New PerinatalAtencionCpt
    Dim DoPerinatalAtencionCpt As New DoPerinatalAtencionCpt
    
    Set PerinatalBuscarOrdenServicioInmunizacion = Nothing
    
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oPerinatalAtencioncpt.Conexion = oConexion
    DoPerinatalAtencionCpt.idAtencion = lnIdAtencion
    
    Set oRecordset = oPerinatalAtencioncpt.PerinatalAtenBuscarOrdenServicioInmunizaciones(DoPerinatalAtencionCpt)
 
   Set PerinatalBuscarOrdenServicioInmunizacion = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'mgaray201410c
Public Function ValidarLabDiagnosticoAtencionCREDFromRs(oRsDiagnosticos As ADODB.Recordset, sNombreColumnaCie As String, _
                ByRef lNroSesion As String, sNombreColumnaLab As String, sCodigoCIE As String) As Boolean
'    Dim sCodigoCIE As String
    Dim bResultado As Boolean
    
    bResultado = True
'    sCodigoCIE = ObtenerCodigoCIEParaAtencionCRED
    
    If oRsDiagnosticos.RecordCount > 0 Then
        oRsDiagnosticos.MoveFirst
        While oRsDiagnosticos.EOF = False
            bResultado = ValidarLabDiagnosticoAtencionCRED(sCodigoCIE, oRsDiagnosticos.Fields(sNombreColumnaCie).value, _
                                lNroSesion, IIf(IsNull(oRsDiagnosticos.Fields(sNombreColumnaLab).value), "", oRsDiagnosticos.Fields(sNombreColumnaLab).value))
            If bResultado = False Then
                bResultado = False
                oRsDiagnosticos.MoveLast
            End If
            oRsDiagnosticos.MoveNext
        Wend
        oRsDiagnosticos.MoveFirst
    End If
    ValidarLabDiagnosticoAtencionCREDFromRs = bResultado
End Function

'validar si en la atenciòn no hay una sesion por ejecutar validar que no se ingrese el codigo de diagnostico correspondiente a control rutinario del niño
Public Function ValidarIngresoDiagnosticoAtencionCREDFromRs(oRsDiagnosticos As ADODB.Recordset, sNombreColumnaCie As String, _
                ByRef lNroSesion As String, sCodigoCIE As String) As Boolean

    Dim bResultado As Boolean
    bResultado = True
    'si no hay sesion por ejecutar comprobar que no se este ingresando el diagnostico de control rutinario del niño
    If lNroSesion = "" Then
        If oRsDiagnosticos.RecordCount > 0 Then
            oRsDiagnosticos.MoveFirst
            While oRsDiagnosticos.EOF = False
                If UCase(Trim(oRsDiagnosticos.Fields(sNombreColumnaCie).value)) = UCase(Trim(sCodigoCIE)) Then
                    bResultado = False
                    oRsDiagnosticos.MoveLast
                End If
                oRsDiagnosticos.MoveNext
            Wend
            oRsDiagnosticos.MoveFirst
        End If
    End If
    ValidarIngresoDiagnosticoAtencionCREDFromRs = bResultado
End Function

Public Function ObtenerCodigoCIEParaAtencionCRED() As String
    ObtenerCodigoCIEParaAtencionCRED = lcBuscaParametro.SeleccionaFilaParametro(349)
End Function

'validar que el codigo lab ingresado para el control rutinario del niño corresponda a la sesion que se esta ejecutando
Private Function ValidarLabDiagnosticoAtencionCRED(sCodigoCIE As String, sCodigoCIEDx As String, lNroSesion As String, sLab As String) As Boolean
    Dim bResultado As Boolean
    bResultado = True
    'solo verifica el lab si el codigo cie pertenece a control rutinario del niño
    If UCase(Trim(sCodigoCIEDx)) = UCase(Trim(sCodigoCIE)) Then
        If UCase(Trim(sLab)) <> UCase(Trim(lNroSesion)) Then
            bResultado = False
        End If
    End If
    ValidarLabDiagnosticoAtencionCRED = bResultado
End Function
'
Private Function AgregarDiagnosticosAtencionDesdePerinatal(oRsDxDesarrollo As ADODB.Recordset, _
            oAtencionesDiagnosticos As AtencionesDiagnosticos, oDOAtencionDiagnostico As DOAtencionDiagnostico, _
            oRsTmpCodigoHis As ADODB.Recordset) As Boolean
            
    oDOAtencionDiagnostico.IdDiagnostico = oRsDxDesarrollo.Fields!Id
    oDOAtencionDiagnostico.idAtencion = oRsDxDesarrollo.Fields!idAtencion
    oDOAtencionDiagnostico.IdClasificacionDx = oRsDxDesarrollo.Fields!IdClasificacionDx
    oDOAtencionDiagnostico.IdSubclasificacionDx = oRsDxDesarrollo.Fields!IdSubclasificacionDx
    If Not IsNull(oRsDxDesarrollo.Fields!labConfHIS) Then
        oDOAtencionDiagnostico.labConfHIS = oRsDxDesarrollo.Fields!labConfHIS
    Else
        If Not oRsTmpCodigoHis.EOF Then
            oDOAtencionDiagnostico.labConfHIS = Left(oRsTmpCodigoHis.Fields!LabHis, InStr(oRsTmpCodigoHis.Fields!LabHis, "-") - 1)
        Else
            oDOAtencionDiagnostico.labConfHIS = Space(3)
        End If
    End If
    AgregarDiagnosticosAtencionDesdePerinatal = AtencionesDiagnosticoAgregarOModificar(oAtencionesDiagnosticos, oDOAtencionDiagnostico)
End Function

'debb-15/06/2016
Function BuscaOtrosCPTyExamenesparaFormatoHIS(lnIdCuentaAtencion As Long, lnIdPuntoCarga As Long, Optional oConexion1 As Connection) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
    End If
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
            Set .ActiveConnection = oConexion
        Else
            Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "BuscaOtrosCPTyExamenesparaFormatoHIS"
        Set oParameter = .CreateParameter("@lnIdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdPuntoCArga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set BuscaOtrosCPTyExamenesparaFormatoHIS = oRecordset
   If oConexion1 Is Nothing Then
        oConexion.Close
        Set oRecordset = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'mgaray201503
Public Function UltimoNroHistoriaGenerado() As Long
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oHistoriasClinicas As New HistoriasClinicas
    
    Dim oRsDesarrolloCabecera As ADODB.Recordset
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    Set oHistoriasClinicas.Conexion = oConexion
    UltimoNroHistoriaGenerado = oHistoriasClinicas.UltimoNroHistoria()
    oConexion.Close
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbInformation, "Error"
    End If
    Set oConexion = Nothing
End Function






'********************************debb-27/05/2015
Function AdmisionCEatenderAlaVezEnVariosCE(oDoAtencion As DOAtencion, oDiagnosticos As Collection, _
                             mo_DoAtencionDatosAdicionales As DoAtencionDatosAdicionales, _
                             oDoPaciente As DOPaciente, oDOCuentaAtencion As DOCuentaAtencion, _
                             mi_opcion As sghOpciones) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCuentaAtencion As New CuentasAtencion
Dim oAtencion As New Atenciones
Dim oAtencionesDx As New AtencionesDiagnosticos
Dim oAtencionesDatosAdicionales As New AtencionesDatosAdicionales
Dim ml_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim oBuscaCodigoNombre As New SIGHNegocios.ReglasComunes
Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim lcMinutosCita As String
Dim lnPrecioUnitario As Double
Dim lcCodigoNombre As String
Dim oRsTmp1 As New Recordset
Dim lcSql As String
Dim bProcesoOK As Boolean
Dim bPacienteNuevo As Boolean
Dim lIdCuentaAtencion As Long
Dim lnFor As Long
'Dim rs As New Recordset

    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    
    Set oCuentaAtencion.Conexion = oConexion
    Set oAtencion.Conexion = oConexion
    Set oAtencionesDx.Conexion = oConexion
    Set oAtencionesDatosAdicionales.Conexion = oConexion
    bProcesoOK = True
    '
    If mi_opcion = sghAgregar Then
        oDOCuentaAtencion.idPaciente = oDoPaciente.idPaciente
        oDOCuentaAtencion.FechaCreacion = CDate(oDoAtencion.FechaIngreso & " " & oDoAtencion.HoraIngreso)
        If Not oCuentaAtencion.Insertar(oDOCuentaAtencion) Then
            bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
        End If
    ElseIf mi_opcion = sghEliminar Then
        oDOCuentaAtencion.IdEstado = 9
        If Not oCuentaAtencion.modificar(oDOCuentaAtencion) Then
            bProcesoOK = False: ms_MensajeError = oCuentaAtencion.MensajeError: GoTo Terminar
        End If
    End If
    '
    'IMPORTANTE!!!!!:    Aqui se crea ATENCION
    If mi_opcion = sghAgregar Then
        oDoAtencion.idCuentaAtencion = oDOCuentaAtencion.idCuentaAtencion   'Se acocia a su atención
        If Not oAtencion.Insertar(oDoAtencion) Then
            bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
        End If
    Else
        If mi_opcion = sghEliminar Then
           oDoAtencion.IdEstadoAtencion = 0
        End If
        If Not oAtencion.modificar(oDoAtencion) Then
            bProcesoOK = False: ms_MensajeError = oAtencion.MensajeError: GoTo Terminar
        End If
    End If
    'Agregar diagnosticos
    If mi_opcion = sghModificar Then
         If Not oAtencionesDx.ActualizarDiagnosticosAtencion(oDiagnosticos, oDoAtencion.idAtencion) Then
             bProcesoOK = False: GoTo Terminar
         End If
    End If
    'datos adicionales
    If mi_opcion <> sghEliminar Then
        mo_DoAtencionDatosAdicionales.idAtencion = oDoAtencion.idAtencion
        If mi_opcion = sghAgregar Then
            If Not oAtencionesDatosAdicionales.Insertar(mo_DoAtencionDatosAdicionales) Then
                bProcesoOK = False: ms_MensajeError = oAtencionesDatosAdicionales.MensajeError: GoTo Terminar
            End If
        Else
            If Not oAtencionesDatosAdicionales.modificar(mo_DoAtencionDatosAdicionales) Then
                bProcesoOK = False: ms_MensajeError = oAtencionesDatosAdicionales.MensajeError: GoTo Terminar
            End If
        End If
    End If
    '
Terminar:
    If bProcesoOK Then
        AdmisionCEatenderAlaVezEnVariosCE = True
        oConexion.CommitTrans
    Else
        AdmisionCEatenderAlaVezEnVariosCE = False
        oConexion.RollbackTrans
    End If
            
    oConexion.Close
    Set oConexion = Nothing
    Set oCuentaAtencion = Nothing
    Set oAtencion = Nothing
    Set oAtencionesDx = Nothing
    Set ml_ReglasFacturacion = Nothing
    Set oBuscaCodigoNombre = Nothing
    Set oRsTmp1 = Nothing
    Set mo_ReglasArchivoClinico = Nothing
    Set oAtencionesDatosAdicionales = Nothing
End Function
'debb-27/05/2015
Function ServiciosAtenSimultaneaSeleccionarXidServicio(lnIdServicio As Long, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaSeleccionarXidServicio"
        Set oParameter = .CreateParameter("@idServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServiciosAtenSimultaneaSeleccionarXidServicio = oRecordset
   'oConexion.Close
   Set oRecordset = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
'debb-27/05/2015
Function ServiciosAtenSimultaneaMovUltimo(oConexion As Connection) As Long
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaMovUltimo"
        Set oRecordset = .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount > 0 Then
     ServiciosAtenSimultaneaMovUltimo = oRecordset!AScorrelativo
   Else
     ServiciosAtenSimultaneaMovUltimo = 0
   End If
   'oConexion.Close
   Set oRecordset = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'debb-27/05/2015
Function ServiciosAtenSimultaneaMovAgregar(lnIdCorrelativo As Long, lnIdAtencion As Long, oConexion As Connection) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaMovAgregar"
        Set oParameter = .CreateParameter("@IdCorrelativo", adInteger, adParamInput, 0, lnIdCorrelativo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   ServiciosAtenSimultaneaMovAgregar = True
   'oConexion.Close
   Set oRecordset = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function
'debb-27/05/2015
Function ServiciosAtenSimultaneaMovXidatencion(lnIdAtencion As Long, oConexion As Connection) As Long
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaMovXidatencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount = 0 Then
      ServiciosAtenSimultaneaMovXidatencion = 0
   Else
      ServiciosAtenSimultaneaMovXidatencion = oRecordset!AScorrelativo
   End If
   'oConexion.Close
   Set oRecordset = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function
'debb-27/05/2015
Function AtencionesEnOtrosConsultoriosAlMismoTiempo(mo_Paciente As DOPaciente, mo_Atenciones As DOAtencion, _
                                                    mo_DoAtencionDatosAdicionales As DoAtencionDatosAdicionales, _
                                                    mo_CuentasAtencion As DOCuentaAtencion, _
                                                    mo_DOAtencionesCE As DOAtencionesCE, _
                                                    mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, _
                                                    mi_opcion As sighentidades.sghOpciones, _
                                                    Optional oRsDx As Recordset, _
                                                    Optional ml_TipoDiagnostico As sghTiposDiagnostico, _
                                                    Optional oRsCPT As Recordset, _
                                                    Optional lbSePulsoBOTON_CANCELAR As Boolean, _
                                                    Optional ml_AScorrelativo As Long) As Boolean
                                                          
                                                          
10        On Error GoTo ErrImpCWEB
          Dim oConexion As New Connection
          Dim oRsAtencionSimultanea As New Recordset
          Dim oRsTmpBuscaAtencion As New Recordset
          Dim oRsTmp1 As New Recordset
          Dim oRsAtencionesExistentes As New Recordset
          Dim mo_Paciente1 As New DOPaciente
          Dim mo_Atencion1 As New DOAtencion
          Dim mo_Diagnostico1 As New Collection
          Dim mo_DoAtencionDatosAdicional1 As New DoAtencionDatosAdicionales
          Dim oDOCuentaAtencion1 As New DOCuentaAtencion
          Dim mo_DOAtencionCE1 As New DOAtencionesCE
          Dim oDOAtencionDiagnostico As DOAtencionDiagnostico
          Dim lnIdAtencionNueva As Long
          Dim lbGrabaAtencionJamo As Boolean
          Dim lnUltimoCorrelativo As Long
          Dim mi_opcion1 As sghOpciones
          Dim lnFor As Integer, lbYaExistenAtenciones As Boolean, lbNuevo As Boolean
          Dim lnIdOrden As Long, lnIdProducto As Long, lcLabConfHIS As String, lbSeEliminaAtencion As Boolean
          '
20        oConexion.CursorLocation = adUseClient
30        oConexion.CommandTimeout = 300
40        oConexion.Open sighentidades.CadenaConexion
          '
50        If ml_AScorrelativo > 0 Then
60            Set oRsAtencionSimultanea = ServiciosAtenSimultaneaSeleccionarXidServicio(mo_Atenciones.idServicioIngreso, oConexion)
70            If oRsAtencionSimultanea.RecordCount > 0 Then
80                Set mo_Paciente1 = mo_Paciente
90                Set mo_Atencion1 = mo_Atenciones
                  
100               Set mo_DoAtencionDatosAdicional1 = mo_DoAtencionDatosAdicionales
110               Set oDOCuentaAtencion1 = mo_CuentasAtencion
120               Set mo_DOAtencionCE1 = mo_DOAtencionesCE
                  '
130               If mi_opcion = sghAgregar Then
140                   lbYaExistenAtenciones = False
150                   lnUltimoCorrelativo = ServiciosAtenSimultaneaMovXidatencion(mo_Atenciones.idAtencion, oConexion)
160                   If lnUltimoCorrelativo > 0 Then
170                       Set oRsAtencionesExistentes = ServiciosAtenSimultaneaMovXcorrelativo(lnUltimoCorrelativo, False)
180                       oRsAtencionesExistentes.Filter = "idatencion<>" & mo_Atenciones.idAtencion
190                       If oRsAtencionesExistentes.RecordCount > 0 Then
200                          lbYaExistenAtenciones = True
210                       End If
220                   Else
230                       lnUltimoCorrelativo = ServiciosAtenSimultaneaMovUltimo(oConexion) + 1
240                       If ServiciosAtenSimultaneaMovAgregar(lnUltimoCorrelativo, mo_Atenciones.idAtencion, oConexion) = True Then
250                       End If
260                   End If
270                   oRsAtencionSimultanea.MoveFirst
280                   Do While Not oRsAtencionSimultanea.EOF
290                       lbNuevo = True
300                       If lbYaExistenAtenciones = True Then
310                          oRsAtencionesExistentes.MoveFirst
320                          oRsAtencionesExistentes.Find "idServicioIngreso=" & oRsAtencionSimultanea!idServicioAtencionSimultanea
330                          If Not oRsAtencionesExistentes.EOF Then
340                             lbNuevo = False
350                          End If
360                       End If
370                       If lbNuevo = True Then
380                           mo_Atencion1.idServicioIngreso = oRsAtencionSimultanea!idServicioAtencionSimultanea
390                           mi_opcion1 = sghAgregar
400                           If AdmisionCEatenderAlaVezEnVariosCE(mo_Atencion1, mo_Diagnostico1, mo_DoAtencionDatosAdicional1, _
                                                                                    mo_Paciente1, oDOCuentaAtencion1, mi_opcion1) = False Then
410                                MsgBox ms_MensajeError: GoTo ErrImpCWEB
420                           End If
430                           If ServiciosAtenSimultaneaMovAgregar(lnUltimoCorrelativo, mo_Atencion1.idAtencion, oConexion) = True Then
440                           End If
450                       End If
460                       oRsAtencionSimultanea.MoveNext
470                   Loop
480               Else
490                   lbSeEliminaAtencion = False
500                   lnUltimoCorrelativo = ServiciosAtenSimultaneaMovXidatencion(mo_Atenciones.idAtencion, oConexion)
510                   Set oRsAtencionSimultanea = ServiciosAtenSimultaneaMovXcorrelativo(lnUltimoCorrelativo, False)
520                   oRsAtencionSimultanea.Filter = "idatencion<>" & mo_Atenciones.idAtencion
530                   If oRsAtencionSimultanea.RecordCount > 0 Then
540                       oRsAtencionSimultanea.MoveFirst
550                       Do While Not oRsAtencionSimultanea.EOF
560                           Set oRsTmp1 = AtencionesSeleccionarPorIdAtencion(oRsAtencionSimultanea.Fields("idatencion").value)
570                           If oRsTmp1.RecordCount > 0 Then
580                                mo_Atencion1.idAtencion = oRsTmp1!idAtencion
590                                mo_Atencion1.idCuentaAtencion = oRsTmp1!idCuentaAtencion
600                                mo_Atencion1.idServicioIngreso = oRsTmp1!idServicioIngreso
610                                mo_Atencion1.IdEstadoAtencion = oRsTmp1!IdEstadoAtencion
620                                oDOCuentaAtencion1.idCuentaAtencion = oRsTmp1!idCuentaAtencion
630                                oDOCuentaAtencion1.IdEstado = 1
640                                mi_opcion1 = mi_opcion
650                           End If
660                           oRsTmp1.Close
                              '*** agrega Dx que corresponden
670                           If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
680                               If mo_Diagnostico1.Count > 0 Then
690                                   For lnFor = mo_Diagnostico1.Count To 1
700                                       mo_Diagnostico1.Remove (lnFor)
710                                   Next
720                               End If
730                               oRsDx.Filter = "idcuentaatencion=" & mo_Atencion1.idCuentaAtencion
740                               If oRsDx.RecordCount > 0 Then
750                                  oRsDx.MoveFirst
760                                  Do While Not oRsDx.EOF
770                                       Set oDOAtencionDiagnostico = New DOAtencionDiagnostico
780                                       oDOAtencionDiagnostico.IdAtencionDiagnostico = 0
790                                       oDOAtencionDiagnostico.idAtencion = mo_Atencion1.idAtencion
800                                       oDOAtencionDiagnostico.IdDiagnostico = oRsDx!IdDiagnostico
810                                       oDOAtencionDiagnostico.IdClasificacionDx = ml_TipoDiagnostico
820                                       oDOAtencionDiagnostico.IdSubclasificacionDx = oRsDx!IdTipoDiagnostico
830                                       oDOAtencionDiagnostico.IdUsuarioAuditoria = mo_Atencion1.IdUsuarioAuditoria
840                                       oDOAtencionDiagnostico.labConfHIS = IIf(IsNull(oRsDx!labConfHIS), "", oRsDx!labConfHIS)
850                                       mo_Diagnostico1.Add oDOAtencionDiagnostico
860                                       oRsDx.MoveNext
870                                  Loop
880                               Else
890                                  lbSeEliminaAtencion = True
900                               End If
910                            End If
                               '
920                            If oRsCPT.RecordCount > 0 Then
930                                  oRsCPT.MoveFirst
940                                  Do While Not oRsCPT.EOF
950                                     If oRsCPT.Fields!idCuentaAtencion = mo_Atencion1.idCuentaAtencion Then
960                                        lbSeEliminaAtencion = False
970                                        Exit Do
980                                     End If
990                                     oRsCPT.MoveNext
1000                                 Loop
1010                           End If
                               '
1020                           If lbSeEliminaAtencion = True Then
1030                              mi_opcion1 = sghEliminar
1040                           End If
1050                           If lbSePulsoBOTON_CANCELAR = True Then
1060                              If mi_opcion1 = sghEliminar Then
1070                                  If AdmisionCEatenderAlaVezEnVariosCE(mo_Atencion1, mo_Diagnostico1, mo_DoAtencionDatosAdicional1, _
                                                                                           mo_Paciente1, oDOCuentaAtencion1, mi_opcion1) = False Then
1080                                      MsgBox ms_MensajeError: GoTo ErrImpCWEB
1090                                  End If
1100                              End If
1110                           Else
1120                              If AdmisionCEatenderAlaVezEnVariosCE(mo_Atencion1, mo_Diagnostico1, mo_DoAtencionDatosAdicional1, _
                                                                                       mo_Paciente1, oDOCuentaAtencion1, mi_opcion1) = False Then
1130                                  MsgBox ms_MensajeError: GoTo ErrImpCWEB
1140                              End If
1150                              If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE And lbSeEliminaAtencion = False Then
                                    
                                     '***Jamo
1160                                 lnIdAtencionNueva = mo_Atencion1.idAtencion
1170                                 Set oRsTmpBuscaAtencion = AtencionCESeleccionarPorIdAtencion(lnIdAtencionNueva)
1180                                 mo_DOAtencionCE1.idAtencion = lnIdAtencionNueva
1190                                 With mo_DOAtencionCE1
1200                                      .CitaIdServicio = mo_Atencion1.idServicioIngreso
1210                                      .CitaServicioJamo = mo_ReglasFacturacion.BuscaServicioActualDelPaciente(mo_Atencion1.idServicioIngreso)
1220                                 End With
1230                                 If oRsTmpBuscaAtencion.RecordCount = 0 Then
1240                                     lbGrabaAtencionJamo = AtencionCEAgregar(mo_DOAtencionCE1, mo_lnIdTablaLISTBARITEMS, _
                                                                      mo_lcNombrePc, "IdAtencion: " & Trim(Str(lnIdAtencionNueva)) & "(desde Atención)")
1250                                 Else
1260                                     lbGrabaAtencionJamo = AtencionCEModificar(mo_DOAtencionCE1, mo_lnIdTablaLISTBARITEMS, _
                                                                      mo_lcNombrePc, "IdAtencion: " & Trim(Str(lnIdAtencionNueva)) & "(desde Atención)")
1270                                 End If
1280                              End If
1290                           End If
1300                           If lbSeEliminaAtencion = True Then
1310                              If ServiciosAtenSimultaneaMovElimXatencion(lnUltimoCorrelativo, mo_Atencion1.idAtencion, oConexion) = True Then
1320                              End If
1330                           End If
1340                           oRsAtencionSimultanea.MoveNext
1350                      Loop
1360                      If mi_opcion = sghEliminar Then
1370                         If ServiciosAtenSimultaneaMovEliminar(lnUltimoCorrelativo, oConexion) Then
1380                            MsgBox ms_MensajeError: GoTo ErrImpCWEB
1390                         End If
1400                      End If
1410                  End If
1420              End If
1430          End If
1440          oRsAtencionSimultanea.Close
1450      End If
          '
1460      If lbSePulsoBOTON_CANCELAR = False Then
1470          If Not (oRsCPT Is Nothing) Then
1480              If oRsCPT.RecordCount > 0 Then
1490                 oRsCPT.MoveFirst
1500                 Do While Not oRsCPT.EOF
1510                    lnIdOrden = oRsCPT!IdOrden
1520                    lnIdProducto = oRsCPT!idProducto
1530                    lcLabConfHIS = IIf(IsNull(oRsCPT!labConfHIS), "", oRsCPT!labConfHIS)
1540                    If mo_ReglasFacturacion.FacturacionServicioDespachoActualizaLAB(lnIdOrden, lnIdProducto, lcLabConfHIS, oConexion) = False Then
1550                       MsgBox ms_MensajeError: GoTo ErrImpCWEB
1560                    End If
1570                    oRsCPT.MoveNext
1580                 Loop
1590              End If
1600          End If
1610      End If
          '
1620      oConexion.Close
1630      AtencionesEnOtrosConsultoriosAlMismoTiempo = True
ErrImpCWEB:
1640      If AtencionesEnOtrosConsultoriosAlMismoTiempo = False Then
1650         MsgBox "Atenciones en otros CONSULTORIOS:  " & Err.Description & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                         "Function AtencionesEnOtrosConsultoriosAlMismoTiempo", "reglasAdmision.cls") 'debb-02/05/2016
             
      'Resume
1660      End If
1670      Set oConexion = Nothing
1680      Set oRsAtencionSimultanea = Nothing
1690      Set oRsTmpBuscaAtencion = Nothing
1700      Set oRsTmp1 = Nothing
1710      Set mo_Paciente1 = Nothing
1720      Set mo_Atencion1 = Nothing
1730      Set mo_Diagnostico1 = Nothing
1740      Set mo_DoAtencionDatosAdicional1 = Nothing
1750      Set oDOCuentaAtencion1 = Nothing
1760      Set mo_DOAtencionCE1 = Nothing
1770      Set oDOAtencionDiagnostico = Nothing
1780      Set mo_Diagnostico1 = Nothing
1790      Set oRsAtencionesExistentes = Nothing
          'Resume
End Function
'debb-27/05/2015
Function ServiciosAtenSimultaneaMovEliminar(lnIdCorrelativo As Long, oConexion As Connection) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaMovEliminar"
        Set oParameter = .CreateParameter("@IdCorrelativo", adInteger, adParamInput, 0, lnIdCorrelativo): .Parameters.Append oParameter
        .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   ServiciosAtenSimultaneaMovEliminar = True
   'oConexion.Close
   Set oRecordset = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function
'debb-27/05/2015
Function ServiciosAtenSimultaneaMovXcorrelativo(lnIdCorrelativo As Long, lbOrdenadoPorConsultorio As Boolean) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaMovXcorrelativo"
        Set oParameter = .CreateParameter("@IdCorrelativo", adInteger, adParamInput, 0, lnIdCorrelativo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@OrdenadoPorConsultorio", adInteger, adParamInput, 0, IIf(lbOrdenadoPorConsultorio = True, 1, 0)): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServiciosAtenSimultaneaMovXcorrelativo = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function
'debb-27/05/2015
Function ServiciosAtenSimultaneaMovCpt(lnIdCorrelativo As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaMovCpt"
        Set oParameter = .CreateParameter("@IdCorrelativo", adInteger, adParamInput, 0, lnIdCorrelativo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServiciosAtenSimultaneaMovCpt = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function


Function ActualizaVariosFUASdeVariosConsultorios(lnIdAtencion As Long, oRsDx As Recordset, oRsCPT As Recordset, _
                       oRsRayosX As Recordset, oRsEcografiaO As Recordset, oRsEcografiaG As Recordset, _
                       oRsTomografia As Recordset, oRsAnatomia As Recordset, oRsPatologia As Recordset, _
                       oRsBanco As Recordset, oRsFarmacia As Recordset, lnAScorrelativo As Long) As Boolean
10        On Error GoTo ErrImpCWEB
          Dim oConexion As New Connection
          Dim rsConsultorios As New Recordset
          Dim rsFuas As New Recordset
          Dim lnIdAtencion1 As Long, lnIdFuaIdCuentaAtencion As Long, lcFuaCodigoPrestacion As String
          '
20        Set rsFuas = ServiciosAtenSimultaneaFuaXcorrelativo(lnAScorrelativo)
30        rsFuas.Filter = "idFuaIdCuentaAtencion>0"
          '
40        oConexion.CursorLocation = adUseClient
50        oConexion.CommandTimeout = 300
60        oConexion.Open sighentidades.CadenaConexion
70        oConexion.BeginTrans
          '
80        If lnAScorrelativo = 0 Then
90            If ServiciosAtenSimultaneaFuaElimXatencion(lnAScorrelativo, lnIdAtencion, oConexion) = False Then
100              GoTo ErrImpCWEB
110           End If
120       Else
130           If ServiciosAtenSimultaneaFuaEliminar(lnAScorrelativo, oConexion) = False Then
140              GoTo ErrImpCWEB
150           End If
160       End If
          '
170       Set rsConsultorios = ServiciosAtenSimultaneaMovXcorrelativo(lnAScorrelativo, False)
          '
180       If oRsDx.RecordCount > 0 Then
190          oRsDx.MoveFirst
200          Do While Not oRsDx.EOF
210             lnIdFuaIdCuentaAtencion = 0
220             If rsFuas.RecordCount > 0 Then
230                rsFuas.MoveFirst
240                rsFuas.Find "idFuaCorrelativo=" & oRsDx!fua
250                If Not rsFuas.EOF Then
260                   lnIdFuaIdCuentaAtencion = rsFuas!idFuaIdCuentaAtencion
270                End If
280             End If
                '
290             If lnAScorrelativo = 0 Then
300                   lnIdAtencion1 = lnIdAtencion
310             Else
320                   rsConsultorios.MoveFirst
330                   rsConsultorios.Find "idCuentaAtencion=" & oRsDx!idCuentaAtencion
340                   lnIdAtencion1 = rsConsultorios!idAtencion
350             End If
                '
360             lcFuaCodigoPrestacion = IIf(IsNull(oRsDx!FuaCodigoPrestacion), "", oRsDx!FuaCodigoPrestacion)
370             If ServiciosAtenSimultaneaFuaAgregar(lnAScorrelativo, lnIdAtencion1, oRsDx!IdDiagnostico, _
                                                     3, oRsDx!fua, lnIdFuaIdCuentaAtencion, oConexion, _
                                                     lcFuaCodigoPrestacion) = False Then
380                GoTo ErrImpCWEB
390             End If
400             oRsDx.MoveNext
410          Loop
420       End If
          '
430       If oRsCPT.RecordCount > 0 Then
440          oRsCPT.MoveFirst
450          Do While Not oRsCPT.EOF
460             lnIdFuaIdCuentaAtencion = 0
470             If rsFuas.RecordCount > 0 Then
480                rsFuas.MoveFirst
490                rsFuas.Find "idFuaCorrelativo=" & oRsCPT!fua
500                If Not rsFuas.EOF Then
510                   lnIdFuaIdCuentaAtencion = rsFuas!idFuaIdCuentaAtencion
520                End If
530             End If
                '
540             If lnAScorrelativo = 0 Then
550                   lnIdAtencion1 = lnIdAtencion
560             Else
570                   rsConsultorios.MoveFirst
580                   rsConsultorios.Find "idCuentaAtencion=" & oRsCPT!idCuentaAtencion
590                   lnIdAtencion1 = rsConsultorios!idAtencion
600             End If
                '
610             If ServiciosAtenSimultaneaFuaAgregar(lnAScorrelativo, lnIdAtencion1, oRsCPT!idProducto, _
                                                  1, oRsCPT!fua, lnIdFuaIdCuentaAtencion, oConexion, "") = False Then
620                GoTo ErrImpCWEB
630             End If
640             oRsCPT.MoveNext
650          Loop
660       End If
          '
670       If oRsRayosX.RecordCount > 0 Then
680          oRsRayosX.MoveFirst
690          Do While Not oRsRayosX.EOF
700             lnIdFuaIdCuentaAtencion = 0
710             If rsFuas.RecordCount > 0 Then
720                rsFuas.MoveFirst
730                rsFuas.Find "idFuaCorrelativo=" & oRsRayosX!fua
740                If Not rsFuas.EOF Then
750                   lnIdFuaIdCuentaAtencion = rsFuas!idFuaIdCuentaAtencion
760                End If
770             End If
                '
780             If ServiciosAtenSimultaneaFuaAgregar(lnAScorrelativo, lnIdAtencion, oRsRayosX!Id, _
                                                     1, oRsRayosX!fua, lnIdFuaIdCuentaAtencion, oConexion, "") = False Then
790                GoTo ErrImpCWEB
800             End If
810             oRsRayosX.MoveNext
820          Loop
830       End If
          '
840       If oRsEcografiaO.RecordCount > 0 Then
850          oRsEcografiaO.MoveFirst
860          Do While Not oRsEcografiaO.EOF
870             lnIdFuaIdCuentaAtencion = 0
880             If rsFuas.RecordCount > 0 Then
890                rsFuas.MoveFirst
900                rsFuas.Find "idFuaCorrelativo=" & oRsEcografiaO!fua
910                If Not rsFuas.EOF Then
920                   lnIdFuaIdCuentaAtencion = rsFuas!idFuaIdCuentaAtencion
930                End If
940             End If
                '
950             If ServiciosAtenSimultaneaFuaAgregar(lnAScorrelativo, lnIdAtencion, oRsEcografiaO!Id, _
                                                     1, oRsEcografiaO!fua, lnIdFuaIdCuentaAtencion, oConexion, "") = False Then
960                GoTo ErrImpCWEB
970             End If
980             oRsEcografiaO.MoveNext
990          Loop
1000      End If
          '
1010      If oRsEcografiaG.RecordCount > 0 Then
1020         oRsEcografiaG.MoveFirst
1030         Do While Not oRsEcografiaG.EOF
1040            lnIdFuaIdCuentaAtencion = 0
1050            If rsFuas.RecordCount > 0 Then
1060               rsFuas.MoveFirst
1070               rsFuas.Find "idFuaCorrelativo=" & oRsEcografiaG!fua
1080               If Not rsFuas.EOF Then
1090                  lnIdFuaIdCuentaAtencion = rsFuas!idFuaIdCuentaAtencion
1100               End If
1110            End If
                '
1120            If ServiciosAtenSimultaneaFuaAgregar(lnAScorrelativo, lnIdAtencion, oRsEcografiaG!Id, _
                                                     1, lnIdFuaIdCuentaAtencion, oRsEcografiaG!fua, oConexion, "") = False Then
1130               GoTo ErrImpCWEB
1140            End If
1150            oRsEcografiaG.MoveNext
1160         Loop
1170      End If
          '
1180      If oRsTomografia.RecordCount > 0 Then
1190         oRsTomografia.MoveFirst
1200         Do While Not oRsTomografia.EOF
1210            lnIdFuaIdCuentaAtencion = 0
1220            If rsFuas.RecordCount > 0 Then
1230               rsFuas.MoveFirst
1240               rsFuas.Find "idFuaCorrelativo=" & oRsTomografia!fua
1250               If Not rsFuas.EOF Then
1260                  lnIdFuaIdCuentaAtencion = rsFuas!idFuaIdCuentaAtencion
1270               End If
1280            End If
                '
1290            If ServiciosAtenSimultaneaFuaAgregar(lnAScorrelativo, lnIdAtencion, oRsTomografia!Id, _
                                                     1, oRsTomografia!fua, lnIdFuaIdCuentaAtencion, oConexion, "") = False Then
1300               GoTo ErrImpCWEB
1310            End If
1320            oRsTomografia.MoveNext
1330         Loop
1340      End If
          '
1350      If oRsAnatomia.RecordCount > 0 Then
1360         oRsAnatomia.MoveFirst
1370         Do While Not oRsAnatomia.EOF
1380            lnIdFuaIdCuentaAtencion = 0
1390            If rsFuas.RecordCount > 0 Then
1400               rsFuas.MoveFirst
1410               rsFuas.Find "idFuaCorrelativo=" & oRsAnatomia!fua
1420               If Not rsFuas.EOF Then
1430                  lnIdFuaIdCuentaAtencion = rsFuas!idFuaIdCuentaAtencion
1440               End If
1450            End If
                '
1460            If ServiciosAtenSimultaneaFuaAgregar(lnAScorrelativo, lnIdAtencion, oRsAnatomia!Id, _
                                                     1, oRsAnatomia!fua, lnIdFuaIdCuentaAtencion, oConexion, "") = False Then
1470               GoTo ErrImpCWEB
1480            End If
1490            oRsAnatomia.MoveNext
1500         Loop
1510      End If
          '
1520      If oRsPatologia.RecordCount > 0 Then
1530         oRsPatologia.MoveFirst
1540         Do While Not oRsPatologia.EOF
1550            lnIdFuaIdCuentaAtencion = 0
1560            If rsFuas.RecordCount > 0 Then
1570               rsFuas.MoveFirst
1580               rsFuas.Find "idFuaCorrelativo=" & oRsPatologia!fua
1590               If Not rsFuas.EOF Then
1600                  lnIdFuaIdCuentaAtencion = rsFuas!idFuaIdCuentaAtencion
1610               End If
1620            End If
                '
1630            If ServiciosAtenSimultaneaFuaAgregar(lnAScorrelativo, lnIdAtencion, oRsPatologia!Id, _
                                                     1, oRsPatologia!fua, lnIdFuaIdCuentaAtencion, oConexion, "") = False Then
1640               GoTo ErrImpCWEB
1650            End If
1660            oRsPatologia.MoveNext
1670         Loop
1680      End If
          '
1690      If oRsBanco.RecordCount > 0 Then
1700         oRsBanco.MoveFirst
1710         Do While Not oRsBanco.EOF
1720            lnIdFuaIdCuentaAtencion = 0
1730            If rsFuas.RecordCount > 0 Then
1740               rsFuas.MoveFirst
1750               rsFuas.Find "idFuaCorrelativo=" & oRsBanco!fua
1760               If Not rsFuas.EOF Then
1770                  lnIdFuaIdCuentaAtencion = rsFuas!idFuaIdCuentaAtencion
1780               End If
1790            End If
                '
1800            If ServiciosAtenSimultaneaFuaAgregar(lnAScorrelativo, lnIdAtencion, oRsBanco!Id, _
                                                      1, oRsBanco!fua, lnIdFuaIdCuentaAtencion, oConexion, "") = False Then
1810               GoTo ErrImpCWEB
1820            End If
1830            oRsBanco.MoveNext
1840         Loop
1850      End If
          '
1860      If oRsFarmacia.RecordCount > 0 Then
1870         oRsFarmacia.MoveFirst
1880         Do While Not oRsFarmacia.EOF
1890            lnIdFuaIdCuentaAtencion = 0
1900            If rsFuas.RecordCount > 0 Then
1910               rsFuas.MoveFirst
1920               rsFuas.Find "idFuaCorrelativo=" & oRsFarmacia!fua
1930               If Not rsFuas.EOF Then
1940                  lnIdFuaIdCuentaAtencion = rsFuas!idFuaIdCuentaAtencion
1950               End If
1960            End If
                '
1970            If ServiciosAtenSimultaneaFuaAgregar(lnAScorrelativo, lnIdAtencion, oRsFarmacia!Id, _
                                                      2, oRsFarmacia!fua, lnIdFuaIdCuentaAtencion, oConexion, "") = False Then
1980               GoTo ErrImpCWEB
1990            End If
2000            oRsFarmacia.MoveNext
2010         Loop
2020      End If
          '
2030      ActualizaVariosFUASdeVariosConsultorios = True
ErrImpCWEB:
2040      If ActualizaVariosFUASdeVariosConsultorios = False Then
2050         MsgBox "Al grabar tabla ServiciosAtenSimultaneaFUA " & Chr(13) & Err.Description & _
              sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                               "Function ActualizaVariosFUASdeVariosConsultorios", "reglasAdmision.cls") 'debb-02/05/2016
             
             
             
2060         oConexion.RollbackTrans
2070      Else
2080         oConexion.CommitTrans
2090      End If
2100      oConexion.Close
2110      Set oConexion = Nothing
2120      Exit Function
2130      Resume
End Function
Function ServiciosAtenSimultaneaFuaEliminar(lnIdCorrelativo As Long, oConexion As Connection) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaFuaEliminar"
        Set oParameter = .CreateParameter("@IdCorrelativo", adInteger, adParamInput, 0, lnIdCorrelativo): .Parameters.Append oParameter
        .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   ServiciosAtenSimultaneaFuaEliminar = True
   'oConexion.Close
   Set oRecordset = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function
Function ServiciosAtenSimultaneaFuaAgregar(lnIdCorrelativo As Long, lnIdAtencion As Long, lnItem As Long, _
                                           lnIdTipo As Long, lnIdFuaCorrelativo As Long, lnIdFuaIdCuentaAtencion As Long, _
                                           oConexion As Connection, lcFuaCodigoPrestacion As String) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaFuaAgregar"
        Set oParameter = .CreateParameter("@IdCorrelativo", adInteger, adParamInput, 0, lnIdCorrelativo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@item", adInteger, adParamInput, 0, lnItem): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipo", adInteger, adParamInput, 0, lnIdTipo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idFuaCorrelativo", adInteger, adParamInput, 0, lnIdFuaCorrelativo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idFuaIdCuentaAtencion", adInteger, adParamInput, 0, IIf(lnIdFuaIdCuentaAtencion = 0, Null, lnIdFuaIdCuentaAtencion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FuaCodigoPrestacion", adVarChar, adParamInput, 3, lcFuaCodigoPrestacion): .Parameters.Append oParameter
        .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   ServiciosAtenSimultaneaFuaAgregar = True
   'oConexion.Close
   Set oRecordset = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function
Function ServiciosAtenSimultaneaFuaXcorrelativo(lnIdCorrelativo As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaFuaXcorrelativo"
        Set oParameter = .CreateParameter("@IdCorrelativo", adInteger, adParamInput, 0, lnIdCorrelativo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServiciosAtenSimultaneaFuaXcorrelativo = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function


Function ServiciosAtenSimultaneaFuaSoloDx(lnIdCorrelativo As Long, lbConInstruccionDISTINCT As Boolean) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaFuaSoloDx"
        Set oParameter = .CreateParameter("@IdCorrelativo", adInteger, adParamInput, 0, lnIdCorrelativo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@conDistinct", adInteger, adParamInput, 0, IIf(lbConInstruccionDISTINCT = True, 1, 0)): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServiciosAtenSimultaneaFuaSoloDx = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function


Function ServiciosAtenSimultaneaFuaModificar(lnAScorrelativo As Long, lnIdFuaCorrelativo As Long, _
                                             lnIdFuaIdCuentaAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
If lnAScorrelativo > 0 And lnIdFuaCorrelativo > 0 Then
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim ms_MensajeError  As String
    Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
        ms_MensajeError = ""
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ServiciosAtenSimultaneaFuaModificar"
            Set oParameter = .CreateParameter("@IdCorrelativo", adInteger, adParamInput, 0, lnAScorrelativo): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdFuaCorrelativo", adInteger, adParamInput, 0, lnIdFuaCorrelativo): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@idFuaIdCuentaAtencion", adInteger, adParamInput, 0, lnIdFuaIdCuentaAtencion): .Parameters.Append oParameter
            .Execute
            
       End With
       ServiciosAtenSimultaneaFuaModificar = True
       oConexion.Close
       Set oRecordset = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
    End If
    Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function
'debb-27/05/2015
Function ServiciosAtenSimultaneaAgregar(lnIdServicio As Long, lnIdServicioAtencionSimultanea As Long, _
                                         oConexion As Connection) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaAgregar"
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idServicioAtencionSimultanea", adInteger, adParamInput, 0, lnIdServicioAtencionSimultanea): .Parameters.Append oParameter
        .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   ServiciosAtenSimultaneaAgregar = True
   'oConexion.Close
   Set oRecordset = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function
Function ServiciosAtenSimultaneaEliminar(lnIdServicio As Long, oConexion As Connection) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaEliminar"
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
        .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   ServiciosAtenSimultaneaEliminar = True
   'oConexion.Close
   Set oRecordset = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function
'debb-27/05/2015
Function ServiciosAtenSimultaneaSeleccionarXServicioAtencionSimultanea(lnIdServicioAtencionSimultanea As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaSeleccionarXConsultorio"
        Set oParameter = .CreateParameter("@IdServicioAtencionSimultanea", adInteger, adParamInput, 0, lnIdServicioAtencionSimultanea): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServiciosAtenSimultaneaSeleccionarXServicioAtencionSimultanea = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
'debb-27/05/2015
Function ServiciosAtenSimultaneaMovElimXatencion(lnIdCorrelativo As Long, lnIdAtencion As Long, oConexion As Connection) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaMovElimXatencion"
        Set oParameter = .CreateParameter("@IdCorrelativo", adInteger, adParamInput, 0, lnIdCorrelativo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   ServiciosAtenSimultaneaMovElimXatencion = True
   'oConexion.Close
   Set oRecordset = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function

'debb-27/05/2015
Function ServiciosAtenSimultaneaImpHISxUPS(lcUPS As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaImpHISxUPS"
        Set oParameter = .CreateParameter("@ups", adVarChar, adParamInput, 6, lcUPS): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServiciosAtenSimultaneaImpHISxUPS = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
'debb-27/05/2015
Function ServiciosAtenSimultaneaFUAequivXups(lcUPS As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaFUAequivXups"
        Set oParameter = .CreateParameter("@ups", adVarChar, adParamInput, 6, lcUPS): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServiciosAtenSimultaneaFUAequivXups = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Sub CreaTemporalYcargaDxYcptXgrupo(ByRef rsDx As Recordset, oRecordsetDx As Recordset, _
                                   ByRef oRecordsetCpt As Recordset, lnIdAtencion As Long, _
                                   lbDesdeReporte As Boolean, lcAd040 As String)
       Dim lnGrupo As Integer, lnSubGrupo As Integer, lnFilas As Integer, lnFor As Integer
       Dim lnPesoKg As Double, oEdad As Edad, lnEdad As Long, lcLabConfHIS As String, lbSigue As Boolean
       Dim oRsAtencionCE As New Recordset, lnIndiceDxCPT As Long, lcLab As String
       Dim rsDxCopia As New Recordset, lcCptDx As String, lbCambioGrupo As Boolean
       Dim lnGrupoAnt As Integer, lnSubGrupoAnt As Integer
       Dim oRsAtencionesLabSinDxSinCpt As New Recordset
       If rsDx.State = 1 Then Set rsDx = Nothing
       Set rsDx = New Recordset
       With rsDx
              .Fields.Append "dDiagnostico", adVarChar, 255, adFldIsNullable
              .Fields.Append "idSubClasificacionDx", adInteger, 4, adFldIsNullable
              .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable
              .Fields.Append "codigoCie10", adVarChar, 20, adFldIsNullable
              .Fields.Append "grupo", adInteger
              .Fields.Append "subgrupo", adInteger
              .Fields.Append "IdDxCpt", adInteger
              .Fields.Append "Dx", adVarChar, 255, adFldIsNullable
              .Fields.Append "TDx", adVarChar, 255, adFldIsNullable
              .Fields.Append "CDx", adVarChar, 255, adFldIsNullable
              .Fields.Append "LDx", adVarChar, 255, adFldIsNullable
              .CursorType = adOpenKeyset
              .LockType = adLockOptimistic
              .Open
        End With
        lnIndiceDxCPT = 1
        '
        If oRecordsetDx.RecordCount > 0 Then
           oRecordsetDx.MoveFirst
           Do While Not oRecordsetDx.EOF

              rsDx.AddNew
              rsDx.Fields!dDiagnostico = Left(oRecordsetDx!dDiagnostico, 255)
              rsDx.Fields!IdSubclasificacionDx = oRecordsetDx!IdSubclasificacionDx
              rsDx.Fields!labConfHIS = IIf(IsNull(oRecordsetDx!labConfHIS), "", oRecordsetDx!labConfHIS)
              rsDx.Fields!CodigoCIE10 = oRecordsetDx!CodigoCIE10
              rsDx.Fields!Grupo = IIf(IsNull(oRecordsetDx!GrupoHIS), 0, oRecordsetDx!GrupoHIS)
              rsDx.Fields!subGrupo = IIf(IsNull(oRecordsetDx!SubGrupoHIS), 0, oRecordsetDx!SubGrupoHIS)
              If lbDesdeReporte = False Then
                    rsDx.Fields!IdDxCpt = lnIndiceDxCPT
                    rsDx.Fields!dx = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRecordsetDx!CodigoCIE10), 6)
                    rsDx.Fields!TDx = IIf(oRecordsetDx!IdSubclasificacionDx = 101, "P", IIf(oRecordsetDx!IdSubclasificacionDx = 102, "D", "R"))
                    rsDx.Fields!CDx = IIf(IsNull(oRecordsetDx.Fields!ClaseDxHIS), "1", Left(oRecordsetDx.Fields!ClaseDxHIS, 1))
                    rsDx.Fields!LDx = IIf(IsNull(oRecordsetDx!labConfHIS), "", oRecordsetDx!labConfHIS)
                    lnIndiceDxCPT = lnIndiceDxCPT + 1
              End If
              rsDx.Update
              oRecordsetDx.MoveNext
           Loop
        End If
        If oRecordsetCpt.State <> 0 Then                        'debb-09/05/2016
            If oRecordsetCpt.RecordCount > 0 Then
               oRecordsetCpt.MoveFirst
               Do While Not oRecordsetCpt.EOF
                  
                  rsDx.AddNew
                  rsDx.Fields!dDiagnostico = IIf(UCase(Trim(oRecordsetCpt!Codigo)) = lcAd040, " ", Left(oRecordsetCpt!nombre, 255))
                  rsDx.Fields!IdSubclasificacionDx = IIf(UCase(Trim(oRecordsetCpt!Codigo)) = lcAd040, 0, 102)
                  rsDx.Fields!labConfHIS = IIf(IsNull(oRecordsetCpt!labConfHIS), "", oRecordsetCpt!labConfHIS)
                  rsDx.Fields!CodigoCIE10 = IIf(UCase(Trim(oRecordsetCpt!Codigo)) = lcAd040, " ", oRecordsetCpt!Codigo)
                  rsDx.Fields!Grupo = IIf(IsNull(oRecordsetCpt!GrupoHIS), 0, oRecordsetCpt!GrupoHIS)
                  rsDx.Fields!subGrupo = IIf(IsNull(oRecordsetCpt!SubGrupoHIS), 0, oRecordsetCpt!SubGrupoHIS)
                  If lbDesdeReporte = False Then
                        rsDx.Fields!IdDxCpt = lnIndiceDxCPT
                        rsDx.Fields!dx = IIf(UCase(Trim(oRecordsetCpt!Codigo)) = lcAd040, " ", oRecordsetCpt!Codigo)
                        rsDx.Fields!TDx = IIf(UCase(Trim(oRecordsetCpt!Codigo)) = lcAd040, " ", "D")
                        rsDx.Fields!CDx = ""
                        rsDx.Fields!LDx = IIf(IsNull(oRecordsetCpt!labConfHIS), "", oRecordsetCpt!labConfHIS)
                        lnIndiceDxCPT = lnIndiceDxCPT + 1
                  End If
                  rsDx.Update
                  oRecordsetCpt.MoveNext
               Loop
            End If
        End If                                      'debb-09/05/2016
        '
        Set oRsAtencionesLabSinDxSinCpt = AtencionesLabSinDxSinCptSeleccionarPorId(lnIdAtencion)
        If oRsAtencionesLabSinDxSinCpt.RecordCount > 0 Then
           oRsAtencionesLabSinDxSinCpt.MoveFirst
           Do While Not oRsAtencionesLabSinDxSinCpt.EOF
              rsDx.AddNew
              rsDx.Fields!dDiagnostico = sighentidades.Lx_LabVacio
              rsDx.Fields!IdSubclasificacionDx = 0
              rsDx!labConfHIS = oRsAtencionesLabSinDxSinCpt!lab
              rsDx.Fields!CodigoCIE10 = sighentidades.Lx_LabVacio
              rsDx!Grupo = oRsAtencionesLabSinDxSinCpt!Grupo
              rsDx!subGrupo = oRsAtencionesLabSinDxSinCpt!subGrupo
              If lbDesdeReporte = False Then
                  rsDx.Fields!IdDxCpt = lnIndiceDxCPT
                  rsDx.Fields!dx = sighentidades.Lx_LabVacio
                  rsDx.Fields!TDx = " "
                  rsDx.Fields!CDx = ""
                  rsDx.Fields!LDx = oRsAtencionesLabSinDxSinCpt!lab
                  lnIndiceDxCPT = lnIndiceDxCPT + 1
              End If
              rsDx.Update
              oRsAtencionesLabSinDxSinCpt.MoveNext
           Loop
        End If
        '
        If rsDx.RecordCount > 0 Then
              rsDx.Sort = "grupo,subgrupo"
              rsDx.MoveLast
              If rsDx!Grupo > 0 Then
                 'agrega lineas en blanco
                 Do While True
                    lnIndiceDxCPT = 1
                    lbSigue = False
                    rsDx.MoveFirst
                    Do While Not rsDx.EOF
                       lnGrupo = rsDx!Grupo
                       lnSubGrupo = 1
                       lnFilas = 0
                       Do While Not rsDx.EOF And lnGrupo = rsDx!Grupo
                          lnFilas = lnFilas + 1
                          If lbDesdeReporte = False Then
                             rsDx!subGrupo = lnSubGrupo                               '*********
                             rsDx.Fields!IdDxCpt = lnIndiceDxCPT
                             rsDx.Update
                             lnIndiceDxCPT = lnIndiceDxCPT + 1
                             lnSubGrupo = lnSubGrupo + 1                              '*********
                          Else
                             rsDx!subGrupo = lnSubGrupo
                             lnSubGrupo = lnSubGrupo + 1
                          End If
                          rsDx.MoveNext
                          If rsDx.EOF Then
                             Exit Do
                          End If
                       Loop
                       lnFor = lnFilas Mod 6
                       If lnFor <> 0 Then
                          lnFor = 6 - lnFor
                          Do While lnFor <> 0
                                lnFilas = lnFilas + 1
                                rsDx.AddNew
                                rsDx.Fields!dDiagnostico = " "
                                rsDx.Fields!IdSubclasificacionDx = 0
                                rsDx.Fields!labConfHIS = " "
                                rsDx.Fields!CodigoCIE10 = " "
                                rsDx.Fields!Grupo = lnGrupo
                                rsDx.Fields!subGrupo = lnFilas
                                If lbDesdeReporte = False Then
                                      rsDx.Fields!IdDxCpt = 1
                                      rsDx.Fields!dx = " "
                                      rsDx.Fields!TDx = " "
                                      rsDx.Fields!CDx = ""
                                      rsDx.Fields!LDx = ""
                                End If
                                rsDx.Update
                                lnFor = lnFor - 1
                          Loop
                          lbSigue = True
                          Exit Do
                       End If
                    Loop
                    If lbSigue = False Then
                       Exit Do
                    End If
                 Loop
           End If

        End If
        Set oRecordsetCpt = BuscaOtrosCPTyExamenesparaFormatoHIS(-10, sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion)
        Set oRsAtencionCE = Nothing
        Set rsDxCopia = Nothing
        Set oRsAtencionesLabSinDxSinCpt = Nothing
End Sub



'Sub CreaTemporalYcargaDxYcpt(ByRef rsDx As Recordset, oRecordsetDx As Recordset, _
'                             ByRef oRecordsetCpt As Recordset, oRsImpHIStodos As Recordset, _
'                             lbAgregaLineasEnBlanco As Boolean, lnIdAtencion As Long, _
'                             ldFechaNacimiento As Date, ldFechaAtencion As Date, oConexionExterna As Connection, _
'                             lbDesdeReporte As Boolean)
'       Dim lnGrupo As Integer, lnSubGrupo As Integer, lnFilas As Integer, lnFor As Integer
'       Dim lnPesoKg As Double, oEdad As Edad, lnEdad As Long, lcLabConfHIS As String, lbSigue As Boolean
'       Dim oRsAtencionCE As New Recordset, lnIndiceDxCPT As Long, lcLab As String
'       Dim rsDxCopia As New Recordset, lcCptDx As String, lbCambioGrupo As Boolean
'       Dim lnGrupoAnt As Integer, lnSubGrupoAnt As Integer
'
'       Const lcAd040 As String = "AD040"
'       If rsDx.State = 1 Then Set rsDx = Nothing
'       Set rsDx = New Recordset
'       With rsDx
'              .Fields.Append "dDiagnostico", adVarChar, 255, adFldIsNullable
'              .Fields.Append "idSubClasificacionDx", adInteger, 4, adFldIsNullable
'              .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable
'              .Fields.Append "codigoCie10", adVarChar, 20, adFldIsNullable
'              .Fields.Append "grupo", adInteger
'              .Fields.Append "subgrupo", adInteger
'              .Fields.Append "IdDxCpt", adInteger
'              .Fields.Append "Dx", adVarChar, 255, adFldIsNullable
'              .Fields.Append "TDx", adVarChar, 255, adFldIsNullable
'              .Fields.Append "CDx", adVarChar, 255, adFldIsNullable
'              .Fields.Append "LDx", adVarChar, 255, adFldIsNullable
'              .CursorType = adOpenKeyset
'              .LockType = adLockOptimistic
'              .Open
'        End With
'        lnIndiceDxCPT = 1
'        '
'        lnPesoKg = 0
'        Set oRsAtencionCE = AtencionCESeleccionarPorIdAtencion(lnIdAtencion, oConexionExterna)
'        If oRsAtencionCE.RecordCount > 0 Then
'           If Not IsNull(oRsAtencionCE!TriajePeso) Then
'              lnPesoKg = oRsAtencionCE!TriajePeso
'           End If
'        End If
'        oRsAtencionCE.Close
'        '
'        If oRecordsetDx.RecordCount > 0 Then
'           oRecordsetDx.MoveFirst
'           Do While Not oRecordsetDx.EOF
'
'              lnGrupo = 0
'              lnSubGrupo = 0
'              lcLabConfHIS = ""
'              If Not IsNull(oRecordsetDx!labConfHIS) Then
'                  lcLabConfHIS = oRecordsetDx!labConfHIS
'              End If
'              oRsImpHIStodos.MoveFirst
'              Do While Not oRsImpHIStodos.EOF
'
'                 lcLab = ""
'                 If Not IsNull(oRsImpHIStodos!lab) Then
'                    If Trim(oRsImpHIStodos!lab) <> "" Then
'                       If UCase(Trim(oRsImpHIStodos!lab)) = UCase(Trim(lcLabConfHIS)) Then
'                          lcLab = ""
'                       Else
'                          lcLab = "x"
'                       End If
'                    End If
'                 End If
'                 If lcLab = "" Then
'                    Select Case oRsImpHIStodos!IdTipoEdad
'                    Case 1
'                          lnEdad = DateDiff("yyyy", ldFechaNacimiento, ldFechaAtencion)
'                    Case 2
'                          lnEdad = DateDiff("m", ldFechaNacimiento, ldFechaAtencion)
'                    Case 3
'                          lnEdad = DateDiff("d", ldFechaNacimiento, ldFechaAtencion)
'                    End Select
'                    If UCase(Trim(oRsImpHIStodos!cpt_dx)) = UCase(Trim(oRecordsetDx!CodigoCIE10)) _
'                                      And (lnEdad >= oRsImpHIStodos!EdadInicio And lnEdad <= oRsImpHIStodos!EdadFinal) _
'                                      And (lnPesoKg >= oRsImpHIStodos!pesoKgMenor And lnPesoKg <= oRsImpHIStodos!pesoKgMayor) Then
'                        lnGrupo = oRsImpHIStodos!Grupo
'                        lnSubGrupo = oRsImpHIStodos!subgrupoorden
'                        '
'                        If rsDx.RecordCount > 0 Then
'                           lcLab = ""
'                           rsDx.MoveFirst
'                           Do While Not rsDx.EOF
'                                If rsDx!CodigoCIE10 = oRecordsetDx!CodigoCIE10 And rsDx!Grupo = lnGrupo And rsDx!SubGrupo = lnSubGrupo Then
'                                   lcLab = "x"
'                                   Exit Do
'                                End If
'                                rsDx.MoveNext
'                           Loop
'                           If lcLab = "" Then
'                              Exit Do
'                           End If
'                        Else
'                           Exit Do
'                        End If
''                       Exit Do
'                        '
'                    End If
'                 End If
'                 oRsImpHIStodos.MoveNext
'              Loop
'
'              rsDx.AddNew
'              rsDx.Fields!dDiagnostico = Left(oRecordsetDx!dDiagnostico, 255)
'              rsDx.Fields!IdSubclasificacionDx = oRecordsetDx!IdSubclasificacionDx
'              rsDx.Fields!labConfHIS = IIf(IsNull(oRecordsetDx!labConfHIS), "", oRecordsetDx!labConfHIS)
'              rsDx.Fields!CodigoCIE10 = oRecordsetDx!CodigoCIE10
'              rsDx.Fields!Grupo = lnGrupo
'              rsDx.Fields!SubGrupo = lnSubGrupo
'              If lbDesdeReporte = False Then
'                    rsDx.Fields!IdDxCpt = lnIndiceDxCPT
'                    rsDx.Fields!dx = sighentidades.DevuelveCodigoDxSinPUNTO(Left(oRecordsetDx!CodigoCIE10, 6))
'                    rsDx.Fields!TDx = IIf(oRecordsetDx!IdSubclasificacionDx = 101, "P", IIf(oRecordsetDx!IdSubclasificacionDx = 102, "D", "R"))
'                    rsDx.Fields!CDx = IIf(IsNull(oRecordsetDx.Fields!ClaseDxHIS), "1", Left(oRecordsetDx.Fields!ClaseDxHIS, 1))
'                    rsDx.Fields!LDx = IIf(IsNull(oRecordsetDx!labConfHIS), "", oRecordsetDx!labConfHIS)
'                    lnIndiceDxCPT = lnIndiceDxCPT + 1
'              End If
'              rsDx.Update
'              oRecordsetDx.MoveNext
'           Loop
'        End If
'        If oRecordsetCpt.RecordCount > 0 Then
'           oRecordsetCpt.MoveFirst
'           Do While Not oRecordsetCpt.EOF
'              lcLabConfHIS = ""
'              If Not IsNull(oRecordsetCpt!labConfHIS) Then
'                   lcLabConfHIS = oRecordsetCpt!labConfHIS
'              End If
'              lnGrupo = 0
'              lnSubGrupo = 0
'              oRsImpHIStodos.MoveFirst
'              Do While Not oRsImpHIStodos.EOF
'                 lcLab = ""
'                 If Not IsNull(oRsImpHIStodos!lab) Then
'                    If Trim(oRsImpHIStodos!lab) <> "" Then
'                       If UCase(Trim(oRsImpHIStodos!lab)) = UCase(Trim(lcLabConfHIS)) Then
'                          lcLab = ""
'                       Else
'                          lcLab = "x"
'                       End If
'                    End If
'                 ElseIf UCase(Trim(oRecordsetCpt!Codigo)) = lcAd040 Then
'                     If lcLabConfHIS = "" Then
'                        lcLab = "x"
'                     End If
'                 End If
'                 If lcLab = "" Then
'                    Select Case oRsImpHIStodos!IdTipoEdad
'                    Case 1
'                          lnEdad = DateDiff("yyyy", ldFechaNacimiento, ldFechaAtencion)
'                    Case 2
'                          lnEdad = DateDiff("m", ldFechaNacimiento, ldFechaAtencion)
'                    Case 3
'                          lnEdad = DateDiff("d", ldFechaNacimiento, ldFechaAtencion)
'                    End Select
'                    If UCase(Trim(oRsImpHIStodos!cpt_dx)) = UCase(Trim(oRecordsetCpt!Codigo)) _
'                                And (lnEdad >= oRsImpHIStodos!EdadInicio And lnEdad <= oRsImpHIStodos!EdadFinal) _
'                                And (lnPesoKg >= oRsImpHIStodos!pesoKgMenor And lnPesoKg <= oRsImpHIStodos!pesoKgMayor) Then
'                       lnGrupo = oRsImpHIStodos!Grupo
'                       lnSubGrupo = oRsImpHIStodos!subgrupoorden
'                       Exit Do
'                    End If
'                 End If
'                 oRsImpHIStodos.MoveNext
'              Loop
'
'              rsDx.AddNew
'              rsDx.Fields!dDiagnostico = IIf(UCase(Trim(oRecordsetCpt!Codigo)) = lcAd040, " ", Left(oRecordsetCpt!nombre, 255))
'              rsDx.Fields!IdSubclasificacionDx = IIf(UCase(Trim(oRecordsetCpt!Codigo)) = lcAd040, 0, 102)
'              rsDx.Fields!labConfHIS = IIf(IsNull(oRecordsetCpt!labConfHIS), "", oRecordsetCpt!labConfHIS)
'              rsDx.Fields!CodigoCIE10 = IIf(UCase(Trim(oRecordsetCpt!Codigo)) = lcAd040, " ", oRecordsetCpt!Codigo)
'              rsDx.Fields!Grupo = lnGrupo
'              rsDx.Fields!SubGrupo = lnSubGrupo
'              If lbDesdeReporte = False Then
'                    rsDx.Fields!IdDxCpt = lnIndiceDxCPT
'                    rsDx.Fields!dx = IIf(UCase(Trim(oRecordsetCpt!Codigo)) = lcAd040, " ", oRecordsetCpt!Codigo)
'                    rsDx.Fields!TDx = IIf(UCase(Trim(oRecordsetCpt!Codigo)) = lcAd040, " ", "D")
'                    rsDx.Fields!CDx = ""
'                    rsDx.Fields!LDx = IIf(IsNull(oRecordsetCpt!labConfHIS), "", oRecordsetCpt!labConfHIS)
'                    lnIndiceDxCPT = lnIndiceDxCPT + 1
'              End If
'              rsDx.Update
'              oRecordsetCpt.MoveNext
'           Loop
'        End If
'        If rsDx.RecordCount > 0 Then
'           rsDx.Sort = "grupo,subgrupo"
'           If lbAgregaLineasEnBlanco = True Then
'              rsDx.MoveLast
'              If rsDx!Grupo > 0 Then
''                 'busca dx solo con 1 registro y lo mete en otro grupo existente
''                 Set rsDxCopia = rsDx.Clone
''                 lbSigue = True
''                 Do While True
''                    rsDx.MoveFirst
''                    Do While Not rsDx.EOF
''                       lnGrupo = rsDx!Grupo
''                       lnSubGrupo = rsDx!subGrupo
''                       lnFilas = 0
''                       lcCptDx = rsDx!CodigoCIE10
''                       Do While Not rsDx.EOF And lnGrupo = rsDx!Grupo
''                          lnFilas = lnFilas + 1
''                          rsDx.MoveNext
''                          If rsDx.EOF Then
''                             lbSigue = False
''                             Exit Do
''                          End If
''                       Loop
''                       lbCambioGrupo = False
''                       If lnFilas = 1 Then
''                          lnGrupoAnt = lnGrupo
''                          lnSubGrupoAnt = lnSubGrupo
''
''                          rsDxCopia.Filter = "grupo>" & lnGrupo
''                          Do While Not rsDxCopia
''                             oRsImpHIStodos.Filter = "grupo=" & rsDxCopia!Grupo & " and cpt_dx='" & lcCptDx & "'"
''                             If oRsImpHIStodos.RecordCount > 0 Then
''                                   oRsImpHIStodos.MoveFirst
''                                   Do While Not oRsImpHIStodos.EOF
''                                         lcLab = ""
''                                         If Not IsNull(oRsImpHIStodos!lab) Then
''                                            If Trim(oRsImpHIStodos!lab) <> "" Then
''                                               If UCase(Trim(oRsImpHIStodos!lab)) = UCase(Trim(lcLabConfHIS)) Then
''                                                  lcLab = ""
''                                               Else
''                                                  lcLab = "x"
''                                               End If
''                                            End If
''                                         End If
''                                         If lcLab = "" Then
''                                            If oRsImpHIStodos!IdTipoEdad = 1 Then
''
''                                                  lnEdad = DateDiff("yyyy", ldFechaNacimiento, ldFechaAtencion)
''                                            ElseIf oRsImpHIStodos!IdTipoEdad = 2 Then
''                                                  lnEdad = DateDiff("m", ldFechaNacimiento, ldFechaAtencion)
''                                            Else
''                                                  lnEdad = DateDiff("d", ldFechaNacimiento, ldFechaAtencion)
''                                            End If
''                                            If UCase(Trim(oRsImpHIStodos!cpt_dx)) = UCase(Trim(lcCptDx)) _
''                                                              And (lnEdad >= oRsImpHIStodos!EdadInicio And lnEdad <= oRsImpHIStodos!EdadFinal) _
''                                                              And (lnPesoKg >= oRsImpHIStodos!pesoKgMenor And lnPesoKg <= oRsImpHIStodos!pesoKgMayor) Then
''                                               lnGrupo = oRsImpHIStodos!Grupo
''                                               lnSubGrupo = oRsImpHIStodos!subgrupoorden
''                                               lbSigue = True
''                                               lbCambioGrupo = True
''                                               Exit Do
''                                            End If
''                                         End If
''                                         oRsImpHIStodos.MoveNext
''                                   Loop
''                                   If lbSigue = True Then
''                                      Exit Do
''                                   End If
''                             End If
''                             rsDxCopia.MoveNext
''                          Loop
''                       End If
''                       If lbSigue = True And lbCambioGrupo = True Then
''                           Exit Do
''                       End If
''
''                    Loop
''                    If lbSigue = False And lbCambioGrupo = False Then
''                       Exit Do
''                    End If
''                    If lbCambioGrupo = True Then
''                        rsDx.Sort = ""
''                        rsDx.MoveFirst
''                        Do While Not rsDx.EOF
''                           If lcCptDx = rsDx!CodigoCIE10 And lnGrupoAnt = rsDx!Grupo And lnSubGrupoAnt = rsDx!subGrupo Then
''                              rsDx!Grupo = lnGrupo
''                              rsDx!subGrupo = lnSubGrupo
''                              rsDx.Update
''                              Exit Do
''                           End If
''                           rsDx.MoveNext
''                        Loop
''                        rsDx.Sort = "grupo,subgrupo"
''                    End If
''                 Loop
''                 oRsImpHIStodos.Filter = ""
'                 'agrega lineas en blanco
'                 Do While True
'                    lnIndiceDxCPT = 1
'                    lbSigue = False
'                    rsDx.MoveFirst
'                    Do While Not rsDx.EOF
'                       lnGrupo = rsDx!Grupo
'                       lnSubGrupo = 1
'                       lnFilas = 0
'                       Do While Not rsDx.EOF And lnGrupo = rsDx!Grupo
'                          lnFilas = lnFilas + 1
'                          If lbDesdeReporte = False Then
'                             rsDx.Fields!IdDxCpt = lnIndiceDxCPT
'                             rsDx.Update
'                             lnIndiceDxCPT = lnIndiceDxCPT + 1
'                          Else
'                             rsDx!SubGrupo = lnSubGrupo
'                             lnSubGrupo = lnSubGrupo + 1
'                          End If
'                          rsDx.MoveNext
'                          If rsDx.EOF Then
'                             Exit Do
'                          End If
'                       Loop
'                       lnFor = lnFilas Mod 6
'                       If lnFor <> 0 Then
'                          lnFor = 6 - lnFor
'                          Do While lnFor <> 0
'                                lnFilas = lnFilas + 1
'                                rsDx.AddNew
'                                rsDx.Fields!dDiagnostico = " "
'                                rsDx.Fields!IdSubclasificacionDx = 0
'                                rsDx.Fields!labConfHIS = " "
'                                rsDx.Fields!CodigoCIE10 = " "
'                                rsDx.Fields!Grupo = lnGrupo
'                                rsDx.Fields!SubGrupo = lnFilas
'                                If lbDesdeReporte = False Then
'                                      rsDx.Fields!IdDxCpt = 1
'                                      rsDx.Fields!dx = " "
'                                      rsDx.Fields!TDx = " "
'                                      rsDx.Fields!CDx = ""
'                                      rsDx.Fields!LDx = ""
'                                End If
'                                rsDx.Update
'                                lnFor = lnFor - 1
'                          Loop
'                          lbSigue = True
'                          Exit Do
'                       End If
'                    Loop
'                    If lbSigue = False Then
'                       Exit Do
'                    End If
'                 Loop
'              End If
'           End If
'
'        End If
'        Set oRecordsetCpt = BuscaOtrosCPTyExamenesparaFormatoHIS(-10)
'        Set oRsAtencionCE = Nothing
'        Set rsDxCopia = Nothing
'End Sub

Function ServiciosAtenSimultaneaFuaElimXatencion(lnIdCorrelativo As Long, lnIdAtencion As Long, oConexion As Connection) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaFuaElimXatencion"
        Set oParameter = .CreateParameter("@IdCorrelativo", adInteger, adParamInput, 0, lnIdCorrelativo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   ServiciosAtenSimultaneaFuaElimXatencion = True
   'oConexion.Close
   Set oRecordset = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function
Function ServiciosAtenSimultaneaFuaXidatencion(lnIdCorrelativo As Long, lnIdAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaFuaXidatencion"
        Set oParameter = .CreateParameter("@IdCorrelativo", adInteger, adParamInput, 0, lnIdCorrelativo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServiciosAtenSimultaneaFuaXidatencion = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function


Function BuscaUPSactualDelPaciente(ml_IdServicioPaciente As Long) As String
    Dim oBuscaNombreServicioPaciente As New SIGHNegocios.ReglasServiciosHosp
    Dim DOServicio As New DOServicio
    Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set DOServicio = oBuscaNombreServicioPaciente.ServiciosSeleccionarPorId(ml_IdServicioPaciente, oConexion)
    BuscaUPSactualDelPaciente = DOServicio.codigoServicioHIS
    oConexion.Close
    Set oConexion = Nothing
    Set DOServicio = Nothing
    Set oBuscaNombreServicioPaciente = Nothing
End Function

'debb-06/08/2015
Function ServiciosAtenSimultaneaMovXcuenta(lnIdCuentaAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosAtenSimultaneaMovXcuenta"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServiciosAtenSimultaneaMovXcuenta = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function


'debb-17/03/2016
Function AtencionesPorIdAtencionConTurnoYfuentefinanciamiento(lnIdAtencion As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String

    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "atencionesPorIdAtencionConTurnoYfuentefinanciamiento"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesPorIdAtencionConTurnoYfuentefinanciamiento = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function


'debb-19/04/2016
Function AtencionesSoloCEconFuenteFinanciamiento(FI As Date, FF As Date, lnIdFuenteFinanciamiento As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSoloCEconFuenteFinanciamiento"
        Set oParameter = .CreateParameter("@FechaIni", adDBTimeStamp, adParamInput, 8, FI): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 8, FF): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSoloCEconFuenteFinanciamiento = oRecordset
   Set oCommand = Nothing
   oConexion.Close
   Set oConexion = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function

'debb-29/04/2016
Function AtencionesXidPaciente(lnIdPaciente As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesXidPaciente"
           Set oParameter = .CreateParameter("@lnIdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesXidPaciente = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function


'debb-15/07/2016
Function AtencionesDatosAdicionalesXfechas(ldFechaInicial As Date, ldFechaFinal As Date, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDatosAdicionalesXfechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 8, ldFechaInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 8, ldFechaFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesDatosAdicionalesXfechas = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function


Function PacientesFiltrarTodosSoloHistoriasDefinit_rap(ByVal oPaciente As DOPaciente, _
                                                       lcSinApellido As String, _
                                                       oConexion As Connection) As ADODB.Recordset
    Dim oTabla As New Pacientes
    Set oTabla.Conexion = oConexion
    Set PacientesFiltrarTodosSoloHistoriasDefinit_rap = oTabla.FiltrarTodosSoloHistoriasDefinit_rap(oPaciente, lcSinApellido)
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function

'debb-21/11/2016
Function SisFuaAtencionSeleccionarPorId(lnCuenta, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SisFuaAtencionSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdCuenta", adInteger, adParamInput, 0, lnCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set SisFuaAtencionSeleccionarPorId = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function
'debb-21/11/2016
Function SisFuaAtencionDevuelveNumeroFUA(lnCuenta As Long) As String
    Dim oConexion As New Connection
    Dim oRsTmp1 As New Recordset
    SisFuaAtencionDevuelveNumeroFUA = ""
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    Set oRsTmp1 = SisFuaAtencionSeleccionarPorId(lnCuenta, oConexion)
    If oRsTmp1.RecordCount > 0 Then
       If Not IsNull(oRsTmp1!fuaNumero) Then
          SisFuaAtencionDevuelveNumeroFUA = oRsTmp1!fuaDisa & "-" & oRsTmp1!fuaLote & "-" & Trim(oRsTmp1!fuaNumero)
       End If
    End If
    oRsTmp1.Close
    oConexion.Close
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
End Function
'debb-21/11/2016
Function DevuelveServicioAlta(lnCuenta As Long, Optional oConexion1 As Connection) As String
    Dim oRsTmp1 As New Recordset
    DevuelveServicioAlta = ""
    Set oRsTmp1 = AtencionesSeleccionarServicioEgresoXCuenta(lnCuenta, oConexion1)
    If oRsTmp1.RecordCount > 0 Then
       If Not IsNull(oRsTmp1!servicioE) Then
          DevuelveServicioAlta = oRsTmp1!servicioE
       End If
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
End Function
'debb-21/11/2016
Function AtencionesSeleccionarServicioEgresoXCuenta(lnIdCuentaAtencion As Long, Optional oConexion1 As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
    End If
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarServicioEgresoXCuenta"
        Set oParameter = .CreateParameter("@lnIdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set AtencionesSeleccionarServicioEgresoXCuenta = oRecordset
   If oConexion1 Is Nothing Then
        oConexion.Close
        Set oRecordset = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function DevuelveServicioIngreso(lnCuenta As Long, oConexion As Connection) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    DevuelveServicioIngreso = ""
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesFiltra"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 2000, " where dbo.Atenciones.idCuentaAtencion=" & Trim(Str(lnCuenta))): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount > 0 Then
      DevuelveServicioIngreso = oRecordset!ServicioIngreso
   End If
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
    
End Function

Function AtencionesDatosAdicionalesSeleccionarDxRefOrigen(lnIdAtencion As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDatosAdicionalesSeleccionarDxRefOrigen"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesDatosAdicionalesSeleccionarDxRefOrigen = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
    
End Function

Function LaFechaDespachoEsMenorAfechaCita(ldFechaCita As Date) As Boolean
         Dim ldHoy As Date
         ldHoy = CDate(lcBuscaParametro.RetornaFechaHoraServidorSQL)
         If ldFechaCita > ldHoy Then
            If Format(ldFechaCita, sighentidades.DevuelveFechaSoloFormato_DMY) = Format(ldHoy, sighentidades.DevuelveFechaSoloFormato_DMY) Then
                LaFechaDespachoEsMenorAfechaCita = False
            Else
                LaFechaDespachoEsMenorAfechaCita = True
                MsgBox "La FECHA DE DESPACHO es menor a: " & ldFechaCita, vbInformation, ""
            End If
         Else
            LaFechaDespachoEsMenorAfechaCita = False
         End If
End Function



Function PacientesSinAltaMedicaEnHospEmerg() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesSinAltaMedicaEnHospEmerg"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PacientesSinAltaMedicaEnHospEmerg = oRecordset
   Set oCommand = Nothing
   oConexion.Close
   Set oConexion = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function




Function AtencionesEstanciaHospitalariaListaMedicos(ml_idAtencion As Long, Optional oConexion1 As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
    End If
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "AtencionesEstanciaHospitalariaListaMedicos"
        Set oParameter = .CreateParameter("@ml_idAtencion", adInteger, adParamInput, 0, ml_idAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set AtencionesEstanciaHospitalariaListaMedicos = oRecordset
   If oConexion1 Is Nothing Then
      oConexion.Close
      Set oRecordset = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description

End Function

Function PacientesSeleccionarXdni(lcNroDocumento As String, lnIdDocIdentidad As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
     Set PacientesSeleccionarXdni = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandType = adCmdStoredProc
        .CommandText = "PacientesFiltraPorNroDocumentoYtipo"
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 12, lcNroDocumento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idDocIdentidad", adInteger, adParamInput, 0, lnIdDocIdentidad): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    Set PacientesSeleccionarXdni = oRecordset
    ms_MensajeError = ""
    oConexion.Close
    Set oConexion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'debb-29/05/2018
Function AtencionesDatosAdicionalesSeleccionarPorIdPaciente(lnIdPaciente As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
     Set AtencionesDatosAdicionalesSeleccionarPorIdPaciente = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandType = adCmdStoredProc
        .CommandText = "AtencionesDatosAdicionalesSeleccionarPorIdPaciente"
        Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    Set AtencionesDatosAdicionalesSeleccionarPorIdPaciente = oRecordset
    ms_MensajeError = ""
    oConexion.Close
    Set oConexion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function AtencionesLabSinDxSinCptSeleccionarPorId(lnIdAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
     Set AtencionesLabSinDxSinCptSeleccionarPorId = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandType = adCmdStoredProc
        .CommandText = "AtencionesLabSinDxSinCptSeleccionarPorId"
        Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    Set AtencionesLabSinDxSinCptSeleccionarPorId = oRecordset
    ms_MensajeError = ""
    oConexion.Close
    Set oConexion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function AtencionesCEseleccionarCITADOS(ldFechaIngreso As Date) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
     Set AtencionesCEseleccionarCITADOS = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandType = adCmdStoredProc
        .CommandText = "AtencionesCEseleccionarCITADOS"
        Set oParameter = .CreateParameter("@lcFechaIngreso", adDBTimeStamp, adParamInput, 0, ldFechaIngreso): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    Set AtencionesCEseleccionarCITADOS = oRecordset
    ms_MensajeError = ""
    oConexion.Close
    Set oConexion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function AtencionesCEseleccionarCitadosYAtendidos(ldFechaIngreso As Date, lnIdServicioIngreso As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim sSql As String
Dim sWhere As String
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
     Set AtencionesCEseleccionarCitadosYAtendidos = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandType = adCmdStoredProc
        .CommandText = "AtencionesCEseleccionarCitadosYAtendidos"
        Set oParameter = .CreateParameter("@lcFechaIngreso", adDBTimeStamp, adParamInput, 0, ldFechaIngreso): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idServicio", adInteger, adParamInput, 0, lnIdServicioIngreso): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    Set AtencionesCEseleccionarCitadosYAtendidos = oRecordset
    ms_MensajeError = ""
    oConexion.Close
    Set oConexion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function TienePermisosParaModificarAtencionesPasadas() As Boolean
        TienePermisosParaModificarAtencionesPasadas = False
        Dim oRsPermisos As New Recordset
        Dim mo_ReglasDeSeguridad As New SIGHNegocios.ReglasDeSeguridad
        Set oRsPermisos = mo_ReglasDeSeguridad.UsuariosRolesSeleccionarPermisosTodos(sighentidades.Usuario)
        oRsPermisos.Filter = "idPermiso=408"
        If oRsPermisos.RecordCount = 1 Then
           TienePermisosParaModificarAtencionesPasadas = True
        End If
        oRsPermisos.Close
        Set oRsPermisos = Nothing
        Set mo_ReglasDeSeguridad = Nothing
End Function



Sub PasaNuevaHC(lcHistoriaActual As String, lcHistoriaNueva As String, oConexion As Connection, oConexionExterna As Connection)
                Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oRsTmp As New Recordset
        Dim oRsTmp2 As New Recordset
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "HistoriasClinicasActualizarNroHistoriaClinica"
            Set oParameter = .CreateParameter("@NroHistoriaNew", adInteger, adParamInput, 0, CLng(lcHistoriaNueva)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@NroHistoriaActual", adInteger, adParamInput, 0, CLng(lcHistoriaActual)): .Parameters.Append oParameter
            Set oRsTmp = .Execute
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "PacientesActualizarNroHistoriaClinica"
            Set oParameter = .CreateParameter("@NroHistoriaNew", adInteger, adParamInput, 0, CLng(lcHistoriaNueva)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@NroHistoriaActual", adInteger, adParamInput, 0, CLng(lcHistoriaActual)): .Parameters.Append oParameter
            Set oRsTmp = .Execute
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
        
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexionExterna
            .CommandTimeout = 150
            .CommandText = "atencionesCEactualizaHistoria"
            Set oParameter = .CreateParameter("@NroHistoriaActual", adInteger, adParamInput, 0, CLng(lcHistoriaActual)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@NroHistoriaNew", adInteger, adParamInput, 0, CLng(lcHistoriaNueva)): .Parameters.Append oParameter
            Set oRsTmp = .Execute
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
        
        
        Set oRsTmp = Nothing
        Set oRsTmp2 = Nothing
   

End Sub

Sub ActualizaHistoriaIgualDNI(lcIdPaciente As String)
   On Error GoTo errActHCDNI
    Dim oConexion As New Connection
    Dim oConexionExterna As New Connection
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexionExterna.CommandTimeout = 900
    oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
'    If lcIdPaciente = "" Then
'       mo_ReglasArchivoClinico.ActualizaIDenTablaLabResultadoPorItems oConexion, 0
'    End If
    '
    If lcBuscaParametro.SeleccionaFilaParametro(351) = "S" Then
      ' On Error Resume Next
       Dim lcSql As String
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim lcNroHistoriaActual As String, lcNroHistoriaNew As String
       If lcIdPaciente = "" Then
          lcSql = "select IdPaciente,NroDocumento,IdTipoNumeracion,NroHistoriaClinica,IdDocIdentidad   from pacientes where  " & _
                " nroDocumento=right('0000000'+ltrim(convert(varchar(12),nroHistoriaClinica)),8) and " & _
                " NroHistoriaClinica <900000000 and IdTipoNumeracion =2 " & _
                " order by IdPaciente desc "
       Else
          lcSql = "select IdPaciente,NroDocumento,IdTipoNumeracion,NroHistoriaClinica,IdDocIdentidad   from pacientes where  " & _
                " idTipoNumeracion=2 and not (NroDocumento is null) and idPaciente=" & lcIdPaciente
       End If
       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       If oRsTmp1.RecordCount > 0 Then
          oRsTmp1.MoveFirst
          Do While Not oRsTmp1.EOF
             lcNroHistoriaActual = oRsTmp1!NroHistoriaClinica
             'GLCC 22/12/20 CAMBIO36 INICIO
             'ANTERIOR:lcNroHistoriaNew = "9" & Right("0000000000" & Trim(oRsTmp1!NroDocumento), 8)
             lcNroHistoriaNew = Right("0000000000" & Trim(oRsTmp1!NroDocumento), 8)
             'GLCC 22/12/20 CAMBIO36 FIN
'             lcSql = "update atencionesCE set nroHistoriaClinica=" & lcNroHistoriaNew & _
'                     " where nroHistoriaClinica=" & lcNroHistoriaActual
'             If oRsTmp2.State = 1 Then oRsTmp2.Close
'             oRsTmp2.Open lcSql, oConexionExterna, adOpenKeyset, adLockOptimistic
             PasaNuevaHC lcNroHistoriaActual, lcNroHistoriaNew, oConexion, oConexionExterna
             oRsTmp1.MoveNext
          Loop
       End If
       oRsTmp1.Close
       Set oRsTmp1 = Nothing
       Set oRsTmp2 = Nothing
    End If
    oConexion.Close
    oConexionExterna.Close
    Set oConexion = Nothing
    Set oConexionExterna = Nothing
    Unload Me
    Exit Sub
errActHCDNI:
    Set oConexion = Nothing
    Set oConexionExterna = Nothing
    Exit Sub
    Resume

End Sub


Sub PacientesActualizarDNI(lcDNI As String, lnIdPaciente As Long, lnIdTipoNumeracion As Long)
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim sSql As String
Dim sWhere As String
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandType = adCmdStoredProc
        .CommandText = "PacientesActualizarDNI"
        Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 12, lcDNI): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoNumeracion", adInteger, adParamInput, 0, lnIdTipoNumeracion): .Parameters.Append oParameter
        .Execute
     End With
    ms_MensajeError = ""
    oConexion.Close
    Set oConexion = Nothing
Exit Sub
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Sub
End Sub



Function AtencionesSeleccionarTSyFFporCuenta(lnIdCuentaAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim sSql As String
Dim sWhere As String
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
     Set AtencionesSeleccionarTSyFFporCuenta = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandType = adCmdStoredProc
        .CommandText = "AtencionesSeleccionarTSyFFporCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    Set AtencionesSeleccionarTSyFFporCuenta = oRecordset
    ms_MensajeError = ""
    oConexion.Close
    Set oConexion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function CitasFormaCitaSeleccionaXcuenta(lnIdCuentaAtencion As Long, oConexion As Connection) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim sSql As String
Dim sWhere As String
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
     CitasFormaCitaSeleccionaXcuenta = "N"
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandType = adCmdStoredProc
        .CommandText = "citasFormaCitaSeleccionaXcuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
     If oRecordset.RecordCount > 0 Then
        If Not IsNull(oRecordset!FormaCita) Then
           CitasFormaCitaSeleccionaXcuenta = oRecordset!FormaCita
        End If
     End If
    ms_MensajeError = ""
    Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Sub CitasFormaCitaActualiza(lnIdCuentaAtencion As Long, lcFormaCita As String, oConexion As Connection)
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim sSql As String
Dim sWhere As String
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandType = adCmdStoredProc
        .CommandText = "CitasFormaCitaActualiza"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@formaCita", adVarChar, adParamInput, 1, lcFormaCita): .Parameters.Append oParameter
        .Execute
     End With
    ms_MensajeError = ""
    Set oRecordset = Nothing
Exit Sub
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Sub
End Sub

Function atencionesSeleccionarXIdCuentaAtencion(lnIdCuentaAtencion As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "atencionesSeleccionarXIdCuentaAtencion"
        Set oParameter = .CreateParameter("@lnIdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set atencionesSeleccionarXIdCuentaAtencion = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

