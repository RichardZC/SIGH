VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasArchivoClinico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Archivo Clínico
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_Teclado As New sighentidades.Teclado
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

'------------------------------------------------------------------------------------------------
'                            REGLAS DE NEGOCIOS
'------------------------------------------------------------------------------------------------
'*****MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function HistoriaClinicaAgregar(oDoHistoria As DOHistoriaClinica, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcHC As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oHistoria As New HistoriasClinicas
Dim oPaciente As New Pacientes
Dim oDoPaciente As New DOPaciente
Dim bProcesoOK As Boolean

    bProcesoOK = True
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oHistoria.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    oConexion.BeginTrans
    If oDoHistoria.IdTipoNumeracion <> sghHistoriaDefinitivaManual Then
        oDoHistoria.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDoHistoria.IdTipoNumeracion)
    End If
    If oDoHistoria.NroHistoriaClinica = 0 Then
        If oDoHistoria.IdTipoNumeracion = sghHistoriaDefinitivaReciclada Then
            MsgBox "No se pudo generar el nro de de historia clínica reciclada", vbInformation, "Atenciones"
            bProcesoOK = False: GoTo Terminar
        End If
    End If
    
    If Not oHistoria.Insertar(oDoHistoria) Then
        bProcesoOK = False: ms_MensajeError = oHistoria.MensajeError: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoHistoria.IdUsuarioAuditoria, "A", oDoHistoria.NroHistoriaClinica, "HistoriasClinicas", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcHC)        'ListBarItems.idListItem
    
    oDoPaciente.idPaciente = oDoHistoria.idPaciente
    If oPaciente.SeleccionarPorId(oDoPaciente) Then
        oDoPaciente.IdTipoNumeracion = oDoHistoria.IdTipoNumeracion
        oDoPaciente.NroHistoriaClinica = oDoHistoria.NroHistoriaClinica
        If Not oPaciente.modificar(oDoPaciente, False) Then
            bProcesoOK = False: ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
        End If
    Else
        bProcesoOK = False: GoTo Terminar
    End If
    HistoriasClinicasActualizaFmovimiento oConexion, oDoHistoria.idPaciente
    
Terminar:
    If bProcesoOK Then
        oConexion.CommitTrans
    Else
        oConexion.RollbackTrans
        'ms_MensajeError = oHistoria.MensajeError
    End If
    
    oConexion.Close
    HistoriaClinicaAgregar = bProcesoOK
    Set oConexion = Nothing
    Set oHistoria = Nothing
    Set oPaciente = Nothing
    
    
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function HistoriaClinicaModificar(oDoHistoriaClinica As DOHistoriaClinica, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcHC As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oHistoriaClinica As New HistoriasClinicas
    
    HistoriaClinicaModificar = False
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
    Set oHistoriaClinica.Conexion = oConexion
    If oHistoriaClinica.modificar(oDoHistoriaClinica) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoHistoriaClinica.IdUsuarioAuditoria, "M", oDoHistoriaClinica.NroHistoriaClinica, "HistoriasClinicas", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcHC)        'ListBarItems.idListItem
        HistoriaClinicaModificar = True
    Else
        ms_MensajeError = oHistoriaClinica.MensajeError
    End If
    HistoriasClinicasActualizaFmovimiento oConexion, oDoHistoriaClinica.idPaciente
    oConexion.Close
    Set oConexion = Nothing
    Set oHistoriaClinica = Nothing
    
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function HistoriaClinicaEliminar(oDoHistoria As DOHistoriaClinica, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcHC As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oHistoriaClinica As New HistoriasClinicas
Dim oPaciente As New Pacientes
Dim oDoPaciente As New DOPaciente
Dim bProcesoOK As Boolean

    
    bProcesoOK = True
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    Set oHistoriaClinica.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    
    oConexion.BeginTrans
    
    If Not oHistoriaClinica.Eliminar(oDoHistoria) Then
        bProcesoOK = False: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoHistoria.IdUsuarioAuditoria, "E", oDoHistoria.NroHistoriaClinica, "HistoriasClinicas", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcHC)        'ListBarItems.idListItem
    
    oDoPaciente.idPaciente = oDoHistoria.idPaciente
    If oPaciente.SeleccionarPorId(oDoPaciente) Then
        oDoPaciente.IdTipoNumeracion = 9
        oDoPaciente.NroHistoriaClinica = 0
        If Not oPaciente.modificar(oDoPaciente, False) Then
            bProcesoOK = False: GoTo Terminar
        End If
    Else
        bProcesoOK = False: GoTo Terminar
    End If
    HistoriasClinicasActualizaFmovimiento oConexion, oDoHistoria.idPaciente
Terminar:
    If bProcesoOK Then
        oConexion.CommitTrans
    Else
        oConexion.RollbackTrans
        ms_MensajeError = oHistoriaClinica.MensajeError
    End If
    
    oConexion.Close
    HistoriaClinicaEliminar = bProcesoOK

    Set oConexion = Nothing
    Set oHistoriaClinica = Nothing
    Set oPaciente = Nothing
    
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function HistoriaClinicaSeleccionarPorId(NroHistoriaClinica As Long) As DOHistoriaClinica
Dim oConexion As New ADODB.Connection
Dim oHistoriaClinica As New HistoriasClinicas
Dim oDoHistoriaClinica As New DOHistoriaClinica
    
    Set HistoriaClinicaSeleccionarPorId = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    
    Set oHistoriaClinica.Conexion = oConexion
    oDoHistoriaClinica.NroHistoriaClinica = NroHistoriaClinica
    If oHistoriaClinica.SeleccionarPorId(oDoHistoriaClinica) Then
        Set HistoriaClinicaSeleccionarPorId = oDoHistoriaClinica
    End If
    oConexion.Close

    ms_MensajeError = oHistoriaClinica.MensajeError

    Set oConexion = Nothing
    Set oHistoriaClinica = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function HistoriaClinicaFiltrar(ByVal oPaciente As DOPaciente, lcSinApellido As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New HistoriasClinicas
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set HistoriaClinicaFiltrar = oTabla.Filtrar(oPaciente, lcSinApellido)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function GenerarNroHistoriaClinica(lTipoGeneracion As sghTipoNumeracionDeNroHistoria) As Long
Dim oConexion As New ADODB.Connection
Dim oHistoriaClinica As New HistoriasClinicas
    
    GenerarNroHistoriaClinica = 0
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oHistoriaClinica.Conexion = oConexion
    GenerarNroHistoriaClinica = oHistoriaClinica.GenerarNroHistoria(lTipoGeneracion)
    oConexion.Close

    ms_MensajeError = oHistoriaClinica.MensajeError
    Set oConexion = Nothing
    Set oHistoriaClinica = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function TiposHistoriaClinicaSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposHistoriaClinica
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set TiposHistoriaClinicaSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MotivosPrestamoHistoriaSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New MotivoPrestamoHistoria
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set MotivosPrestamoHistoriaSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MotivosMovimientoHistoriaSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New MotivosMovimientoHistoria
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set MotivosMovimientoHistoriaSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function EstadosPrestamoHistoriaSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New EstadosPrestamoHistoria
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set EstadosPrestamoHistoriaSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function EstadosHistoriaClinicaSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New EstadosHistoriaClinica
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set EstadosHistoriaClinicaSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function TiposGeneracionHistoriasSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposNumeracionHistoria
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set TiposGeneracionHistoriasSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oConexion = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function TiposGeneracionHistoriasSeleccionarDeConsultaExterna() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposNumeracionHistoria
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set TiposGeneracionHistoriasSeleccionarDeConsultaExterna = oTabla.SeleccionarDeConsultaExterna
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function TiposGeneracionHistoriasSeleccionarDeEmergencia() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposNumeracionHistoria
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set TiposGeneracionHistoriasSeleccionarDeEmergencia = oTabla.SeleccionarDeEmergencia
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function TiposGeneracionHistoriasSeleccionarDeHospitalizacion() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposNumeracionHistoria
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set TiposGeneracionHistoriasSeleccionarDeHospitalizacion = oTabla.SeleccionarDeHospitalizacion
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function TiposGeneracionHistoriaSeleccionarDefinitivos(lIdAdicional As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposNumeracionHistoria
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set TiposGeneracionHistoriaSeleccionarDefinitivos = oTabla.SeleccionarDefinitivos(lIdAdicional)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


'*********************************************************
'            HISTORIAS SOLICITADAS
'*********************************************************
'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function HistoriasSolicitadasAgregar(oDOHistoriasSolicitadas As DOHistoriaSolicitada, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPaciente As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oHistoriasSolicitadas As New HistoriasSolicitadas
    
    HistoriasSolicitadasAgregar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    Set oHistoriasSolicitadas.Conexion = oConexion
    If oHistoriasSolicitadas.Insertar(oDOHistoriasSolicitadas) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOHistoriasSolicitadas.IdUsuarioAuditoria, "A", oDOHistoriasSolicitadas.IdHistoriaSolicitada, "HistoriasSolicitadas", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPaciente)            'ListBarItems.idListItem
        HistoriasSolicitadasAgregar = True
    Else
        ms_MensajeError = oHistoriasSolicitadas.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oHistoriasSolicitadas = Nothing
    
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function HistoriasSolicitadasModificar(oDOHistoriasSolicitadas As DOHistoriaSolicitada, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPaciente As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oHistoriasSolicitadasClinica As New HistoriasSolicitadas
    
    HistoriasSolicitadasModificar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    Set oHistoriasSolicitadasClinica.Conexion = oConexion
    If oHistoriasSolicitadasClinica.modificar(oDOHistoriasSolicitadas) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOHistoriasSolicitadas.IdUsuarioAuditoria, "M", oDOHistoriasSolicitadas.IdHistoriaSolicitada, "HistoriasSolicitadas", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPaciente)            'ListBarItems.idListItem
        HistoriasSolicitadasModificar = True
    Else
        ms_MensajeError = oHistoriasSolicitadasClinica.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oHistoriasSolicitadasClinica = Nothing
    
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function HistoriasSolicitadasEliminar(oDoHistoria As DOHistoriaSolicitada, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPaciente As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oHistoriasSolicitadasClinica As New HistoriasSolicitadas
    
    HistoriasSolicitadasEliminar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    Set oHistoriasSolicitadasClinica.Conexion = oConexion
    If oHistoriasSolicitadasClinica.Eliminar(oDoHistoria) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoHistoria.IdUsuarioAuditoria, "E", oDoHistoria.IdHistoriaSolicitada, "HistoriasSolicitadas", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPaciente)            'ListBarItems.idListItem
        HistoriasSolicitadasEliminar = True
    Else
        ms_MensajeError = oHistoriasSolicitadasClinica.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oHistoriasSolicitadasClinica = Nothing
    
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function HistoriasSolicitadasSeleccionarPorId(IdHistoriasSolicitadas As Long) As DOHistoriaSolicitada
Dim oConexion As New ADODB.Connection
Dim oHistoriasSolicitadas As New HistoriasSolicitadas
Dim oDOHistoriasSolicitadas As New DOHistoriaSolicitada
    
    Set HistoriasSolicitadasSeleccionarPorId = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oHistoriasSolicitadas.Conexion = oConexion
    oDOHistoriasSolicitadas.IdHistoriaSolicitada = IdHistoriasSolicitadas
    If oHistoriasSolicitadas.SeleccionarPorId(oDOHistoriasSolicitadas) Then
        Set HistoriasSolicitadasSeleccionarPorId = oDOHistoriasSolicitadas
    End If
    oConexion.Close

    ms_MensajeError = oHistoriasSolicitadas.MensajeError
    Set oConexion = Nothing
    Set oHistoriasSolicitadas = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function HistoriasSolicitadasFiltrar(ByVal oPaciente As DOPaciente, ByVal oHistoria As DOHistoriaSolicitada, ByVal oDoArchiveroServ As DOArchiveroServicio) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New HistoriasSolicitadas
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set HistoriasSolicitadasFiltrar = oTabla.Filtrar(oPaciente, oHistoria, oDoArchiveroServ)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function HistoriasSolicitadasSeleccionarPorIdPaciente(ByVal lidPaciente As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New HistoriasSolicitadas
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set HistoriasSolicitadasSeleccionarPorIdPaciente = oTabla.SeleccionarPorIdPaciente(lidPaciente)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function HistoriasSolicitadasSeleccionarPorIdMovimiento(ByVal lIdMovimiento As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New HistoriasSolicitadas
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set HistoriasSolicitadasSeleccionarPorIdMovimiento = oTabla.SeleccionarPorIdMovimiento(lIdMovimiento)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function HistoriasSolicitadasSeleccionarPorArchivero(ByVal lIdEmpleado As Long, lIdMotivo As Long, lIdServicioDestino As Long, sOperador As String, sFechaDesde As String, sFechaHasta As String, lnNroHistoriaClinica As Long, lbEsFechaRequerimiento As Boolean, lnIdArchivoClinico As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New HistoriasSolicitadas
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set HistoriasSolicitadasSeleccionarPorArchivero = oTabla.SeleccionarPorArchivero(lIdEmpleado, lIdMotivo, lIdServicioDestino, sOperador, sFechaDesde, sFechaHasta, lnNroHistoriaClinica, lbEsFechaRequerimiento, lnIdArchivoClinico)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'*********************************************************
'            MOVIMIENTO DE HISTORIAS
'*********************************************************
Function MovimientosHistoriaClinicaAgregar(oMovimiento As DOMovimientoHistoriaClinica, rsMovimientos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPaciente As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oMovimientosHistoriaClinica As New MovimientosHistoriaClinica
Dim oHistoriaSolicitadas As New HistoriasSolicitadas
    
    MovimientosHistoriaClinicaAgregar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    oConexion.BeginTrans
    Set oMovimientosHistoriaClinica.Conexion = oConexion
    Set oHistoriaSolicitadas.Conexion = oConexion
    
    If oMovimientosHistoriaClinica.AgregarVarios(oMovimiento, rsMovimientos) Then
        If oHistoriaSolicitadas.ActualizarIdMovimiento(rsMovimientos) Then
            '
            Call mo_ReglasSeguridad.AuditoriaAgregarV(oMovimiento.IdUsuarioAuditoria, "A", oMovimiento.IdMovimiento, "MovimientosHistoriaClinica", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPaciente)          'ListBarItems.idListItem
            MovimientosHistoriaClinicaAgregar = True
            oConexion.CommitTrans
        End If
    Else
        ms_MensajeError = oMovimientosHistoriaClinica.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimientosHistoriaClinica = Nothing
    Set oHistoriaSolicitadas = Nothing
    
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosHistoriaClinicaModificar(oMovimiento As DOMovimientoHistoriaClinica, rsMovimientos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPaciente As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oMovimientosHistoriaClinicaClinica As New MovimientosHistoriaClinica
Dim oHistoriaSolicitadas As New HistoriasSolicitadas
    
    MovimientosHistoriaClinicaModificar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    oConexion.BeginTrans
    Set oMovimientosHistoriaClinicaClinica.Conexion = oConexion
    Set oHistoriaSolicitadas.Conexion = oConexion
    
    If oHistoriaSolicitadas.ActualizarIdMovimiento(rsMovimientos) Then
        If oMovimientosHistoriaClinicaClinica.ModificarVarios(oMovimiento, rsMovimientos) Then
            '
            Call mo_ReglasSeguridad.AuditoriaAgregarV(oMovimiento.IdUsuarioAuditoria, "M", oMovimiento.IdMovimiento, "MovimientosHistoriaClinica", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPaciente)        'ListBarItems.idListItem
            MovimientosHistoriaClinicaModificar = True
            oConexion.CommitTrans
        End If
    Else
        ms_MensajeError = oMovimientosHistoriaClinicaClinica.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oHistoriaSolicitadas = Nothing
    
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosHistoriaClinicaEliminar(oMovimiento As DOMovimientoHistoriaClinica, rsMovimientos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPaciente As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oMovimientosHistoriaClinicaClinica As New MovimientosHistoriaClinica
Dim oHistoriaSolicitadas As New HistoriasSolicitadas
    
    MovimientosHistoriaClinicaEliminar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    oConexion.BeginTrans
    Set oHistoriaSolicitadas.Conexion = oConexion
    
    Set oMovimientosHistoriaClinicaClinica.Conexion = oConexion
    If oHistoriaSolicitadas.EliminarIdMovimientoIdPaciente(rsMovimientos) Then
        If oMovimientosHistoriaClinicaClinica.EliminarVarios(rsMovimientos) Then
            '
            Call mo_ReglasSeguridad.AuditoriaAgregarV(oMovimiento.IdUsuarioAuditoria, "E", oMovimiento.IdMovimiento, "MovimientosHistoriaClinica", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPaciente)        'ListBarItems.idListItem
            MovimientosHistoriaClinicaEliminar = True
            oConexion.CommitTrans
        End If
    Else
        ms_MensajeError = oMovimientosHistoriaClinicaClinica.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oHistoriaSolicitadas = Nothing
    
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosHistoriaClinicaSeleccionarPorId(IdMovimientosHistoriaClinica As Long) As DOMovimientoHistoriaClinica
Dim oConexion As New ADODB.Connection
Dim oMovimientosHistoriaClinica As New MovimientosHistoriaClinica
Dim oDOMovimientosHistoriaClinica As New DOMovimientoHistoriaClinica
    
    Set MovimientosHistoriaClinicaSeleccionarPorId = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    oConexion.CursorLocation = adUseClient
    
    Set oMovimientosHistoriaClinica.Conexion = oConexion
    oDOMovimientosHistoriaClinica.IdMovimiento = IdMovimientosHistoriaClinica
    If oMovimientosHistoriaClinica.SeleccionarPorId(oDOMovimientosHistoriaClinica) Then
        Set MovimientosHistoriaClinicaSeleccionarPorId = oDOMovimientosHistoriaClinica
    End If
    oConexion.Close

    ms_MensajeError = oMovimientosHistoriaClinica.MensajeError
    Set oConexion = Nothing
    Set oMovimientosHistoriaClinica = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosHistoriaClinicaFiltrar(ByVal oPaciente As DOPaciente, _
                                           ByVal oHistoria As DOMovimientoHistoriaClinica, _
                                           lcApellidoVacio As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New MovimientosHistoriaClinica
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set MovimientosHistoriaClinicaFiltrar = oTabla.Filtrar(oPaciente, oHistoria, lcApellidoVacio)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosFormatosHCFiltrar(ByVal oPaciente As DOPaciente, ByVal oHistoria As DOMovimientoHistoriaClinica) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New MovimientoFormatosHC
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set MovimientosFormatosHCFiltrar = oTabla.Filtrar(oPaciente, oHistoria)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosHistoriaUltimaUbicacion(lidPaciente As Long) As Long
Dim oConexion As New ADODB.Connection
Dim oTabla As New MovimientosHistoriaClinica

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    MovimientosHistoriaUltimaUbicacion = oTabla.UltimaUbicacionDeHistoria(lidPaciente)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosHistoriasClinicasPorIdGrupo(lIdGrupo As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New MovimientosHistoriaClinica
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set MovimientosHistoriasClinicasPorIdGrupo = oTabla.MovimientosHistoriasClinicasPorIdGrupo(lIdGrupo)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosFormatosHCPorIdGrupo(lIdGrupo As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New MovimientoFormatosHC
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set MovimientosFormatosHCPorIdGrupo = oTabla.MovimientosFormatosHCPorIdGrupo(lIdGrupo)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosHistoriasClinicasDetallePorIdPaciente(lidPaciente As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New MovimientosHistoriaClinica
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set MovimientosHistoriasClinicasDetallePorIdPaciente = oTabla.MovimientosHistoriasClinicasDetallePorIdPaciente(lidPaciente)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function HistoriasSolicitadasSeleccionarDetallePorIdPaciente(lidPaciente As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New HistoriasSolicitadas
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set HistoriasSolicitadasSeleccionarDetallePorIdPaciente = oTabla.HistoriasSolicitadasSeleccionarDetallePorIdPaciente(lidPaciente)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function HistoriasSolicitadasSeleccionarPorIdAtencion(lnIdAtencion As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New HistoriasSolicitadas
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    Set HistoriasSolicitadasSeleccionarPorIdAtencion = oTabla.HistoriasSolicitadasSeleccionarPorIdAtencion(lnIdAtencion)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosHistoriaEsUltimoMovimiento(lidPaciente As Long, lIdMovimiento As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTabla As New MovimientosHistoriaClinica
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    MovimientosHistoriaEsUltimoMovimiento = oTabla.EsUltimoMovimiento(lidPaciente, lIdMovimiento)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function


'*********************************************************
'            ARCHIVEROS SERVICIOS
'*********************************************************
'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function ArchiveroServicioAgregar(oArchiverosServicio As Collection, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oArchiveroServicio As New ArchiveroServicio
    
    ArchiveroServicioAgregar = False
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
    Set oArchiveroServicio.Conexion = oConexion
    If oArchiveroServicio.AgregarVarios(oArchiverosServicio) Then
        ArchiveroServicioAgregar = True
    Else
        ms_MensajeError = oArchiveroServicio.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oArchiveroServicio = Nothing
    
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function ArchiveroServicioModificar(oArchiveroServicio As Collection, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oArchiveroServicioClinica As New ArchiveroServicio
    
    ArchiveroServicioModificar = False
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    Set oArchiveroServicioClinica.Conexion = oConexion
    If oArchiveroServicioClinica.ModificarVarios(oArchiveroServicio) Then
        ArchiveroServicioModificar = True
    Else
        ms_MensajeError = oArchiveroServicioClinica.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oArchiveroServicioClinica = Nothing
    
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function ArchiveroServicioEliminar(oArchiveroServicio As Collection, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lnIdUsuarioAuditoria As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oArchiveroServicioClinica As New ArchiveroServicio
    
    ArchiveroServicioEliminar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set oArchiveroServicioClinica.Conexion = oConexion
    If oArchiveroServicioClinica.EliminarVarios(oArchiveroServicio) Then
        ArchiveroServicioEliminar = True
    Else
        ms_MensajeError = oArchiveroServicioClinica.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oArchiveroServicioClinica = Nothing
    
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function ArchiveroServicioSeleccionarPorId(IdArchiveroServicio As Long) As DOMovimientoHistoriaClinica
Dim oConexion As New ADODB.Connection
Dim oArchiveroServicio As New ArchiveroServicio
Dim oDOArchiveroServicio As New DOArchiveroServicio
    
    Set ArchiveroServicioSeleccionarPorId = Nothing
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oArchiveroServicio.Conexion = oConexion
    oDOArchiveroServicio.IdArchivero = IdArchiveroServicio
    If oArchiveroServicio.SeleccionarPorId(oDOArchiveroServicio) Then
        Set ArchiveroServicioSeleccionarPorId = oDOArchiveroServicio
    End If
    oConexion.Close

    ms_MensajeError = oArchiveroServicio.MensajeError
    Set oConexion = Nothing
    Set oArchiveroServicio = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function ArchiveroServicioFiltrar(ByVal oEmpleado As DOEmpleado, lbEsConsultorioAsignado As Boolean) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New ArchiveroServicio
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set ArchiveroServicioFiltrar = oTabla.Filtrar(oEmpleado, lbEsConsultorioAsignado)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function ArchiveroServicioFiltrarPorEmpleado(lIdEmpleado As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New ArchiveroServicio
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set ArchiveroServicioFiltrarPorEmpleado = oTabla.FiltrarPorEmpleado(lIdEmpleado)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function MovimientosHistoriasClinicasParaDevolver(lIdServicioDestino As Long, sFechaMovimiento As String, lIdUsuario As Long, lnNroHistoriaClinica As Long, lnIdArchivoClinico As Long, lnIdPacienteSolo As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New MovimientosHistoriaClinica
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set MovimientosHistoriasClinicasParaDevolver = oTabla.MovimientosHistoriasClinicasParaDevolver2(lIdServicioDestino, sFechaMovimiento, lIdUsuario, lnNroHistoriaClinica, lnIdArchivoClinico, lnIdPacienteSolo)
    
    
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'Yamill 2307
Function MovimientosHistoriasClinicasParaDevolverPorNroHistoria(lIdServicioDestino As Long, sFechaMovimiento As String, _
        lIdUsuario As Long, lnNroHistoriaClinica As Long, lnIdArchivoClinico As Long, _
        lnIdPacienteSolo As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New MovimientosHistoriaClinica
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set MovimientosHistoriasClinicasParaDevolverPorNroHistoria = oTabla.MovimientosHistoriasClinicasParaDevolverPorNroHistoria(lIdServicioDestino, sFechaMovimiento, _
                                    lIdUsuario, lnNroHistoriaClinica, lnIdArchivoClinico, lnIdPacienteSolo)

    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'Yamill 2307
Function MovimientosHistoriasClinicasParaDevolverPorServicio(lIdServicioDestino As Long, sFechaMovimiento As String, _
        lIdUsuario As Long, lnNroHistoriaClinica As Long, lnIdArchivoClinico As Long, _
        lnIdPacienteSolo As Long, FechaTipoFiltro As Integer, operador As String, fechaInicio As String, _
        fechaFin As String, fechaFinRequerida As Boolean) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New MovimientosHistoriaClinica
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set MovimientosHistoriasClinicasParaDevolverPorServicio = oTabla.MovimientosHistoriasClinicasParaDevolverPorServicio(lIdServicioDestino, sFechaMovimiento, _
                                    lIdUsuario, lnNroHistoriaClinica, lnIdArchivoClinico, lnIdPacienteSolo, _
                                    FechaTipoFiltro, operador, fechaInicio, fechaFin, fechaFinRequerida)

    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function ServiciosPorArchivero(lIdEmpleado As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Servicios
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set ServiciosPorArchivero = oTabla.SeleccionarPorIdArchivero(lIdEmpleado)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function ServiciosPorArchiveroTipoServicio(lIdEmpleado As Long, lcTipoServicio As String) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Servicios
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set ServiciosPorArchiveroTipoServicio = oTabla.SeleccionarPorIdArchiveroTipoServicio(lIdEmpleado, lcTipoServicio)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosHistoriaClinicaMovimientosPorDia(ldFechaIni As Date, ldFechaFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set MovimientosHistoriaClinicaMovimientosPorDia = Nothing
    ms_MensajeError = ""
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "MovimientosHistoriaClinicaMovimientosPorDia"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set MovimientosHistoriaClinicaMovimientosPorDia = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosFormatoHCMovimientosPorDia(ldFechaIni As Date, ldFechaFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set MovimientosFormatoHCMovimientosPorDia = Nothing
    ms_MensajeError = ""
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "MovimientosFormatoHCMovimientosPorDia"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set MovimientosFormatoHCMovimientosPorDia = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


'*********************************************************
'            MOVIMIENTO DE FORMATOS DE HISTORIAS
'*********************************************************
'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosFormatosHCAgregar(oMovimiento As DOMovimientoFormatoHC, rsMovimientos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPaciente As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oMovimientosHistoriaClinica As New MovimientoFormatosHC
    MovimientosFormatosHCAgregar = False
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oMovimientosHistoriaClinica.Conexion = oConexion
    If oMovimientosHistoriaClinica.AgregarVarios(oMovimiento, rsMovimientos) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oMovimiento.IdUsuarioAuditoria, "A", oMovimiento.IdMovimiento, "MovimientosFormatosHistoriaClinica", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPaciente)          'ListBarItems.idListItem
        MovimientosFormatosHCAgregar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oMovimientosHistoriaClinica.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimientosHistoriaClinica = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosFormatosHCModificar(oMovimiento As DOMovimientoFormatoHC, rsMovimientos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPaciente As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oMovimientosHistoriaClinicaClinica As New MovimientoFormatosHC
    MovimientosFormatosHCModificar = False
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oMovimientosHistoriaClinicaClinica.Conexion = oConexion
    If oMovimientosHistoriaClinicaClinica.ModificarVarios(oMovimiento, rsMovimientos) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oMovimiento.IdUsuarioAuditoria, "M", oMovimiento.IdMovimiento, "MovimientosFormatosHistoriaClinica", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPaciente)        'ListBarItems.idListItem
        MovimientosFormatosHCModificar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oMovimientosHistoriaClinicaClinica.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimientosHistoriaClinicaClinica = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosFormatosHCEliminar(oMovimiento As DOMovimientoFormatoHC, rsMovimientos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPaciente As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oMovimientosHistoriaClinicaClinica As New MovimientoFormatosHC
    MovimientosFormatosHCEliminar = False
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oMovimientosHistoriaClinicaClinica.Conexion = oConexion
    If oMovimientosHistoriaClinicaClinica.EliminarVarios(rsMovimientos) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oMovimiento.IdUsuarioAuditoria, "E", oMovimiento.IdMovimiento, "MovimientosFormatosHistoriaClinica", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPaciente)        'ListBarItems.idListItem
        MovimientosFormatosHCEliminar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oMovimientosHistoriaClinicaClinica.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimientosHistoriaClinicaClinica = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosFormatosHCSeleccionarPorId(IdMovimientosHistoriaClinica As Long) As DOMovimientoFormatoHC
Dim oConexion As New ADODB.Connection
Dim oMovimientosHistoriaClinica As New MovimientoFormatosHC
Dim oDOMovimientosHistoriaClinica As New DOMovimientoFormatoHC
    Set MovimientosFormatosHCSeleccionarPorId = Nothing
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    Set oMovimientosHistoriaClinica.Conexion = oConexion
    oDOMovimientosHistoriaClinica.IdMovimiento = IdMovimientosHistoriaClinica
    If oMovimientosHistoriaClinica.SeleccionarPorId(oDOMovimientosHistoriaClinica) Then
        Set MovimientosFormatosHCSeleccionarPorId = oDOMovimientosHistoriaClinica
    End If
    oConexion.Close
    ms_MensajeError = oMovimientosHistoriaClinica.MensajeError
    Set oConexion = Nothing
    Set oMovimientosHistoriaClinica = Nothing
End Function


'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosFormatoHCMovimientosRetornoHCPorFechasYtipoServicio(ldFechaIni As Date, ldFechaFin As Date, lnIdTipoServicio As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set MovimientosFormatoHCMovimientosRetornoHCPorFechasYtipoServicio = Nothing
    ms_MensajeError = ""
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "MovimientosFormatoHCMovimientosRetornoHCPorFechasYtipoServicio"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoServicio", adInteger, adParamInput, 0, lnIdTipoServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set MovimientosFormatoHCMovimientosRetornoHCPorFechasYtipoServicio = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MovimientosHistoriaClinicaSeleccionaUltimoMovimientoPorPaciente(lnIdPaciente As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set MovimientosHistoriaClinicaSeleccionaUltimoMovimientoPorPaciente = Nothing
    ms_MensajeError = ""
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "MovimientosHistoriaClinicaSeleccionaUltimoMovimientoPorPaciente"
        Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set MovimientosHistoriaClinicaSeleccionaUltimoMovimientoPorPaciente = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 15/10/2013 - se agrego oConexion.CommandTimeout = 300
Function ArchiverosSeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
'Dim ssql As String
'Dim sWhere As String
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
Set ArchiverosSeleccionarTodos = Nothing
        
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
        

        With oCommand
'           .CommandType = adCmdText
           .CommandType = adCmdStoredProc
           Set .ActiveConnection = oConexion
           .CommandTimeout = 150
'           .CommandText = ssql
           .CommandText = "ArchiverosSeleccionarTodos"
           Set oRecordset = .Execute
           Set oRecordset.ActiveConnection = Nothing
        End With

        Set ArchiverosSeleccionarTodos = oRecordset
        ms_MensajeError = ""
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
        Exit Function
ManejadorDeError:
        ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
        Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Sub ActualizaIdPacienteEnTodasLasTablas(lnIdPacienteEncontradoConHC As Long, lnIdPacienteActualConHCoTemp As Long, lnIdHistoriaClinica As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lnHistoriaNueva As Long, ml_idUsuario As Long)
          Dim oConexion As New ADODB.Connection
          Dim oCommand As New ADODB.Command
          Dim oParameter As ADODB.Parameter
          Dim lcSql As String
          Dim oRsTmp As New Recordset
          Dim lcBuscaParametro As New SIGHDatos.Parametros
          Dim oConexionExterna As New Connection
          oConexionExterna.CommandTimeout = 300
          oConexionExterna.CursorLocation = adUseClient
          
          lcSql = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
          oConexionExterna.Open lcSql
          oConexion.CommandTimeout = 300
          oConexion.Open sighentidades.CadenaConexion
          
          With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "ActualizaIdPacienteEnTodasLasTablas"
                Set oParameter = .CreateParameter("@IdPacienteEncontradoConHC", adInteger, adParamInput, 0, lnIdPacienteEncontradoConHC): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdPacienteActualConHCoTemp", adInteger, adParamInput, 0, lnIdPacienteActualConHCoTemp): .Parameters.Append oParameter
                .Execute
          End With
          Set oParameter = Nothing
        '
          '
          Set oCommand = Nothing
          Set oParameter = Nothing
          With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexionExterna
                .CommandTimeout = 150
                .CommandText = "ActualizaAtencionesCEXNroHistoriaClinica"
                Set oParameter = .CreateParameter("@NroHistoriaNueva", adInteger, adParamInput, 0, lnHistoriaNueva): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdHistoriaClinica", adInteger, adParamInput, 0, lnIdHistoriaClinica): .Parameters.Append oParameter
                .Execute
          End With
          
          If ml_idUsuario > 0 Then
             Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "M", lnIdHistoriaClinica, "HistoriasClinicas", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "Se remplazó el N° Historia: " & Trim(Str(lnIdHistoriaClinica)) & " por " & Trim(Str(lnHistoriaNueva)))        'ListBarItems.idListItem
          End If
          '
'          oConexion.CommitTrans
          oConexion.Close
          oConexionExterna.Close
          Set oConexion = Nothing
'          Set oRsTmp = Nothing
          Set lcBuscaParametro = Nothing
          Set oConexionExterna = Nothing
          Set oCommand = Nothing
          Exit Sub
ErrActPaciente:
          MsgBox Err.Description
          oConexion.RollbackTrans
          oConexion.Close
          Set oConexion = Nothing
          oConexionExterna.Close
          Set oCommand = Nothing
          Set oConexionExterna = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Function HistoriaClinicaSeleccionarXhistoriaYtipoNumeracion(lnNroHistoriaClinica As Long, lnIdTipoNumeracion As Long) As Recordset
'    Dim lcSql As String
    Dim oRsTmp1 As New Recordset
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
      .CommandText = "HistoriaClinicaSeleccionarXhistoriaYtipoNumeracion"
      Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, lnIdTipoNumeracion): .Parameters.Append oParameter
      Set oRsTmp1 = .Execute
      Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set HistoriaClinicaSeleccionarXhistoriaYtipoNumeracion = oRsTmp1
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Function HistoriasSolicitadasSeleccionarXmotivosYfechas(lcIdMotivo As String, sOperador As String, lcFechaDesde As String, lcFechaHasta As String, lnNroHistoriaClinica As Long, lnIdServicio As Long) As Recordset
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim lcWhereIdTipoServicio As String
    Dim lcWhereFechaIngreso As String
    Dim lcWhereNroHistoriaClinica As String
    Dim lcWhereIdServicioIngreso As String
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
      
    lcWhereIdTipoServicio = "dbo.Atenciones.idTipoServicio in " & IIf(lcIdMotivo = "1", "(1)", IIf(lcIdMotivo = "2", "(3)", "(2,4)"))
    lcWhereFechaIngreso = ""
    lcWhereNroHistoriaClinica = ""
    lcWhereIdServicioIngreso = ""
    
    If sOperador <> "" Then
        If lcFechaHasta = sighentidades.FECHA_VACIA_DMY Then
            lcWhereFechaIngreso = " and dbo.Atenciones.FechaIngreso " & sOperador & "'" & lcFechaDesde & "'"
        Else
            lcWhereFechaIngreso = " and  Atenciones.FechaIngreso Between (CONVERT(DATETIME,'" & lcFechaDesde & "',103)) and (CONVERT(DATETIME,'" & lcFechaHasta & "',103))"
        End If
    End If
    If lnNroHistoriaClinica > 0 Then
        lcWhereNroHistoriaClinica = " and Pacientes.NroHistoriaClinica = " & lnNroHistoriaClinica
    End If
    If lnIdServicio > 0 Then
       lcWhereIdServicioIngreso = " and dbo.Atenciones.idServicioIngreso = " & lnIdServicio
    End If
    
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "HistoriasSolicitadasSeleccionarXmotivosYfechas"
    Set oParameter = .CreateParameter("@WhereIdTipoServicio", adVarChar, adParamInput, 200, lcWhereIdTipoServicio): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@WhereFechaIngreso", adVarChar, adParamInput, 200, lcWhereFechaIngreso): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@WhereNroHistoriaClinica", adVarChar, adParamInput, 200, lcWhereNroHistoriaClinica): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@WhereIdServicioIngreso", adVarChar, adParamInput, 200, lcWhereIdServicioIngreso): .Parameters.Append oParameter
    Set oRsTmp1 = .Execute
    Set oRsTmp1.ActiveConnection = Nothing
  End With

    Set HistoriasSolicitadasSeleccionarXmotivosYfechas = oRsTmp1
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Function HistoriaUltimoServicioDondeEstubo(lnIdPaciente As Long) As Recordset
'    Dim lcSql As String
    Dim oRsTmp1 As New Recordset
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
      .CommandText = "HistoriaUltimoServicioDondeEstubo"
      Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
      Set oRsTmp1 = .Execute
      Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set HistoriaUltimoServicioDondeEstubo = oRsTmp1
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Function HistoriaUltimoMovimiento(lnIdPaciente As Long) As Recordset
    Dim oRsTmp1 As New Recordset
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
      .CommandText = "HistoriaUltimoMovimiento"
      Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
      Set oRsTmp1 = .Execute
      Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set HistoriaUltimoMovimiento = oRsTmp1
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Function HistoriaPagoCita(lnIdAtencion As Long, lcMotivo As String) As Recordset
'    Dim lcSql As String
    Dim oRsTmp1 As New Recordset
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
        .CommandText = "HistoriaPagoCita"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Motivo", adChar, adParamInput, 1, lcMotivo): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
  Set HistoriaPagoCita = oRsTmp1
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  
End Function

'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Function HistoriasPagoCitaDescripcionTarifa(lnIdAtencion As Long, lcMotivo As String) As Recordset
'    Dim lcSql As String
    Dim oRsTmp1 As New Recordset
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
        .CommandText = "HistoriasPagoCitaDescripcionTarifa"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Motivo", adChar, adParamInput, 1, lcMotivo): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
  Set HistoriasPagoCitaDescripcionTarifa = oRsTmp1
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Function HistoriasMovimientos(lnIdPaciente As Long) As Recordset
'    Dim lcSql As String
    Dim oRsTmp1 As New Recordset
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
      .CommandText = "HistoriasMovimientos"
      Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
      Set oRsTmp1 = .Execute
      Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set HistoriasMovimientos = oRsTmp1
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'Datos con Problemas
Sub ActualizaDatosConProblemas(lbDesdeActualizacionEstructuras As Boolean)
    Dim oConexion As New Connection, oConexionExterna As New Connection, oConexionSIS As New Connection
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim mo_AdminServiciosComunes As New ReglasComunes
    Dim lcMensajeLicencia As String
    oConexion.CommandTimeout = 300: oConexion.CursorLocation = adUseClient: oConexion.Open sighentidades.CadenaConexion
    oConexionExterna.CommandTimeout = 300: oConexionExterna.CursorLocation = adUseClient: oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    oConexionSIS.CommandTimeout = 300: oConexionSIS.CursorLocation = adUseClient: oConexionSIS.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    
    ArreglaIdPacienteEnAtenciones oConexion
    ActualizaGrupoSanguineoYfactorRH1 oConexion
    HistoriasClinicasActualizaFmovimiento oConexion, 0
    AutogeneradoConNULL oConexion
    AtencionesEpisodiosDetalleEliminaNull oConexion
    generadorNroHistoriaClinicaActualizaNroAutomaticoDeHistoriaClinica oConexion
    'ActualizaIDenTablaLabResultadoPorItems oConexion, 0
    ActualizaIdServicioNULL oConexion
    ActualizaNullhistoriaClinica oConexion
    AtencionesConIdPacienteSinKeyEnTablaPacientes oConexion
    TarifasNULLplanesNULLlesAsignaValor oConexion
    ActualizaCuentasAnuladasEnTablaAtenciones oConexion
    CierraPreVentasFarmaciaDelDiaAnterior oConexion
    ActualizaFechaEgresoEnCEparaLosQtienenDx oConexion
    ActualizaDxParaCE oConexionExterna
    AnulaPagosConsultaCEporPaquetes oConexion
    If lbDesdeActualizacionEstructuras = True Then
        ActualizaHistoricoDePreciosFarmacia oConexion     'debb2014b
        ActualizaAUTOGENERADOconHORA oConexion            'debb2014b
        ActualizaCampoDxSINpunto oConexion
        ActualizaTiempoPromedioAtencion oConexion
        ActualizaRecetasGeneradas oConexion
        ActualizaEDAD_en_movimientos oConexion, 0          'debb-18/05/2016
        ActualizaTildesMayusculasAFILIADOS_SIS
        ActualizaDireccionYdistritoDomicilio oConexion
        farmActualizaPrecioPonderadoMenorAcero oConexion
    End If
    ActualizaEpisodioClinico oConexion
    ActualizaNULLenIdPartidaIdCCostoCon999
    ActualizaSeUsaSinPrecioNULL oConexion
    ActualizaAUTOGENERADOconHORA oConexion
    ActualizaUltimoNroExoneracion oConexion
    ActualizaLabResultadosXItemsENtablaLabPruebas oConexion
    ActualizaTipoTarifaSinDetalle oConexion
    ActualizaIdEstanciaHospitalariaEnServios oConexion
    ActualizaEstablecimientosSIS oConexion, oConexionSIS
    HistoriasSolicitadasNOexistentes oConexion   'debb-07/10/2015
    CitasNOexistentes oConexion    'debb-15/09/2015
    ItemsBoletaFarmaciaConCantidadCERO oConexion  'debb-07/10/2015
    Fua2015conNULLenFuaTipoAnexo2015 oConexion, oConexionExterna   'debb-07/10/2015
    ActualizaCPTsinTipoTarifa oConexion   'debb-14/04/2016
    ActualizaLOTES_his_conNULL oConexion   'debb-02/05/2016
    ActualizaPerinatalEdadEnSemanas oConexion   'debb-09/06/2016
    
    GeneraIdPacienteAlosNacidosEnElHospital oConexion    'debb-19/09/2016
    '
'    If true Then 'licencia
'    End If
    '
    actualizaCamasDeObservacionConNull oConexion     'franklin 2017
    oConexion.Close: Set oConexion = Nothing
    oConexionExterna.Close: Set oConexionExterna = Nothing
    Set mo_AdminServiciosComunes = Nothing
End Sub

Sub GeneraIdPacienteAlosNacidosEnElHospital(oConexion As Connection)
       On Error GoTo ErrLabIt1
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oAtencionesNacimientos As New AtencionesNacimientos
       Dim oDOAtencionNacimiento As New DOAtencionNacimiento
       Dim oPacientes As New Pacientes
       Dim oDoPaciente As New DOPaciente
       Dim lcSql As String, lcDireccion As String, lcDNI As String, wxParametro323 As String, lnIdDistrito As Long
       lcSql = "select idPaciente,NroOrdenHijo,madreDocumento from pacientes where not (madreDocumento is null)"
       If oRsTmp2.State = 1 Then oRsTmp2.Close
       oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       lcSql = "SELECT     dbo.AtencionesNacimientos.IdNacimiento, dbo.AtencionesNacimientos.docIdentidad, dbo.AtencionesNacimientos.IdDocIdentidad, " & _
            "                      dbo.AtencionesNacimientos.IdPacienteNacido, dbo.Atenciones.IdAtencion, dbo.Pacientes.NroDocumento AS mamaDocumento," & _
            "                      dbo.Pacientes.IdDocIdentidad AS mamaIdTipoDocumento,  dbo.AtencionesNacimientos.NroOrdenHijo" & _
            " FROM         dbo.Pacientes RIGHT OUTER JOIN" & _
            "                      dbo.Atenciones ON dbo.Pacientes.IdPaciente = dbo.Atenciones.IdPaciente RIGHT OUTER JOIN" & _
            "                      dbo.AtencionesNacimientos ON dbo.Atenciones.IdAtencion = dbo.AtencionesNacimientos.IdAtencion" & _
            " Where (Not (dbo.Pacientes.NroDocumento Is Null)) And (dbo.Pacientes.IdDocIdentidad = 1) And (dbo.AtencionesNacimientos.IdPacienteNacido Is Null)" & _
            "                      AND (dbo.Atenciones.idEstadoAtencion <> 0)"
      oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
      If oRsTmp1.RecordCount > 0 And oRsTmp2.RecordCount > 0 Then
         Set oAtencionesNacimientos.Conexion = oConexion
         Set oPacientes.Conexion = oConexion
         oDOAtencionNacimiento.IdUsuarioAuditoria = sighentidades.Usuario
         oDoPaciente.IdUsuarioAuditoria = sighentidades.Usuario
         oRsTmp1.MoveFirst
         Do While Not oRsTmp1.EOF
            oRsTmp2.Filter = "madreDocumento='" & oRsTmp1!mamaDocumento & "' and NroOrdenHijo=" & oRsTmp1!NroOrdenHijo
            If oRsTmp2.RecordCount > 0 Then
               oDOAtencionNacimiento.idNacimiento = oRsTmp1!idNacimiento
               If oAtencionesNacimientos.SeleccionarPorId(oDOAtencionNacimiento) = True Then
                  oDOAtencionNacimiento.idPacienteNacido = oRsTmp2!idPaciente
                  If oAtencionesNacimientos.modificar(oDOAtencionNacimiento) = True Then
                  End If
               End If
               If Not IsNull(oRsTmp1!idDocIdentidad) And Not IsNull(oRsTmp1!docIdentidad) Then
                  oDoPaciente.idPaciente = oRsTmp2!idPaciente
                  If oPacientes.SeleccionarPorId(oDoPaciente) = True Then
                     If oDoPaciente.idDocIdentidad = 0 And oDoPaciente.NroDocumento = "" Then
                        oDoPaciente.idDocIdentidad = oRsTmp2!idDocIdentidad
                        oDoPaciente.NroDocumento = oRsTmp2!docIdentidad
                        If oPacientes.modificar(oDoPaciente, False) = True Then
                        End If
                     End If
                  End If
               End If
            End If
            oRsTmp1.MoveNext
         Loop
      End If
      oRsTmp1.Close
ErrLabIt1:
       Set oRsTmp1 = Nothing
       Set oRsTmp2 = Nothing
       Set oAtencionesNacimientos = Nothing
       Set oDOAtencionNacimiento = Nothing
       Set oPacientes = Nothing
       Set oDoPaciente = Nothing
End Sub

Sub ActualizaDireccionYdistritoDomicilio(oConexion As Connection)
    On Error GoTo ErrLabIt1
       Dim lcBuscaParametro As New SIGHDatos.Parametros
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oRsBuscaPacientesSis As New Recordset
       Dim mo_SisConsumoWeb As New SIGHNegocios.SisConsumoWeb
       Dim lcSql As String, lcDireccion As String, lcDNI As String, wxParametro323 As String, lnIdDistrito As Long
    lcSql = "SELECT     dbo.Pacientes.IdPaciente, dbo.Pacientes.DireccionDomicilio, dbo.Pacientes.IdDistritoDomicilio, dbo.Atenciones.IdAtencion, " & _
"                      dbo.Atenciones.IdCuentaAtencion , dbo.Atenciones.idTipoServicio, dbo.Atenciones.FechaIngreso," & _
"                      dbo.Pacientes.NroDocumento" & _
" FROM         dbo.Atenciones INNER JOIN" & _
"                      dbo.Pacientes ON dbo.Atenciones.IdPaciente = dbo.Pacientes.IdPaciente" & _
" WHERE     (dbo.Atenciones.IdTipoServicio > 1) AND (YEAR(dbo.Atenciones.FechaIngreso) = 2016) AND" & _
"           (MONTH(dbo.Atenciones.FechaIngreso) > 6) and  dbo.Atenciones.idEstadoAtencion<>0 AND" & _
"            not (dbo.Pacientes.NroDocumento is null) and" & _
"                      ((dbo.Pacientes.IdDistritoDomicilio IS NULL) OR (dbo.Pacientes.DireccionDomicilio IS NULL)) and" & _
"                      (dbo.Atenciones.IdTipoServicio > 1) AND (YEAR(dbo.Atenciones.FechaIngreso) = 2016) AND" & _
"                       (MONTH(dbo.Atenciones.FechaIngreso) > 6) "
      oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
      If oRsTmp1.RecordCount > 0 Then
         wxParametro323 = lcBuscaParametro.SeleccionaFilaParametro(323)
         oRsTmp1.MoveFirst
         Do While Not oRsTmp1.EOF
            lcDNI = oRsTmp1!NroDocumento
            lcDireccion = "": lnIdDistrito = 0
            Set oRsBuscaPacientesSis = mo_SisConsumoWeb.WebServiceSISBuscarAfiliado(lcDNI, "", _
                                                "", "", "", "", wxParametro323)
            If oRsBuscaPacientesSis.RecordCount > 0 Then
               lnIdDistrito = Val(oRsBuscaPacientesSis.Fields!DistritoDomicilio)
            End If
            oRsBuscaPacientesSis.Close
            lcSql = ""
            If lcDireccion <> "" Then
                If IsNull(oRsTmp1!DireccionDomicilio) Then
                   lcSql = "update Pacientes set Direcciondomicilio='" & lcDireccion & "'"
                End If
            End If
            If lnIdDistrito > 0 Then
                If IsNull(oRsTmp1!IdDistritoDomicilio) Then
                   If lcSql = "" Then
                      lcSql = "update Pacientes set idDistritoDomicilio=" & lnIdDistrito
                   Else
                      lcSql = lcSql & ",idDistritoDomicilio=" & lnIdDistrito
                   End If
                End If
            End If
            If lcSql <> "" Then
                lcSql = lcSql & " where idPaciente=" & oRsTmp1!idPaciente
                oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                If lcDireccion <> "" Then
                   lcSql = "update Pacientes set Direcciondomicilio='" & lcDireccion & "' where idAtencion=" & oRsTmp1!idAtencion
                   oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                End If
            End If
            oRsTmp1.MoveNext
         Loop
      End If
      oRsTmp1.Close
      '*** DIRECCION desde Pacientes
      lcSql = "SELECT        dbo.Pacientes.DireccionDomicilio, dbo.AtencionesDatosAdicionales.DireccionDomicilio AS ddirecion, dbo.Atenciones.IdPaciente, dbo.Atenciones.IdAtencion" & _
" FROM            dbo.Atenciones LEFT OUTER JOIN" & _
"                         dbo.AtencionesDatosAdicionales ON dbo.Atenciones.IdAtencion = dbo.AtencionesDatosAdicionales.idAtencion LEFT OUTER JOIN" & _
"                         dbo.Pacientes ON dbo.Atenciones.IdPaciente = dbo.Pacientes.IdPaciente" & _
" Where (dbo.Pacientes.DireccionDomicilio Is Not Null) And (dbo.AtencionesDatosAdicionales.DireccionDomicilio Is Null)"
      oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
      If oRsTmp1.RecordCount > 0 Then
         oRsTmp1.MoveFirst
         Do While Not oRsTmp1.EOF
            lcDireccion = oRsTmp1!DireccionDomicilio
            lcSql = "update AtencionesDatosAdicionales set Direcciondomicilio='" & lcDireccion & _
                    "' where idAtencion=" & oRsTmp1!idAtencion
            oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
            oRsTmp1.MoveNext
         Loop
      End If
      oRsTmp1.Close
      
ErrLabIt1:
       Set oRsTmp1 = Nothing
       Set oRsTmp2 = Nothing
       Set oRsBuscaPacientesSis = Nothing
       Set mo_SisConsumoWeb = Nothing
       Set lcBuscaParametro = Nothing
End Sub

Sub ActualizaTildesMayusculasAFILIADOS_SIS()
'On Error GoTo ErrLabIt1
'   Dim oRsTmp1 As New Recordset
'   Dim oRsTmp2 As New Recordset
'   Dim oConexion1 As New Connection
'   Dim lcSql As String
'   oConexion1.CommandTimeout = 900
'   oConexion1.CursorLocation = adUseClient
'   oConexion1.Open "dsn=galenhosExterna"
'   oRsTmp1.Open "select idsiasis,codigo,paterno,materno,pnombre,onombres from sisfiliaciones", oConexion1, adOpenKeyset, adLockOptimistic
'   If oRsTmp1.RecordCount > 0 Then
'      oRsTmp1.MoveFirst
'      Do While Not oRsTmp1.EOF
'
'         oRsTmp1!paterno = mo_Teclado.DevuelveTextoSINtildes(UCase(oRsTmp1!paterno))
'         oRsTmp1!materno = mo_Teclado.DevuelveTextoSINtildes(UCase(oRsTmp1!materno))
'         oRsTmp1!Pnombre = mo_Teclado.DevuelveTextoSINtildes(UCase(oRsTmp1!Pnombre))
'         If Not IsNull(oRsTmp1!oNombres) Then
'            oRsTmp1!oNombres = mo_Teclado.DevuelveTextoSINtildes(mo_Teclado.CapitalizarNombres(oRsTmp1!oNombres))
'         End If
'         oRsTmp1.Update
'         oRsTmp1.MoveNext
'      Loop
'   End If
'   oRsTmp1.Close
'ErrLabIt1:
'   Set oRsTmp1 = Nothing
'   Set oRsTmp2 = Nothing
'   Set oConexion1 = Nothing
End Sub
Sub ActualizaPerinatalEdadEnSemanas(oConexion As Connection)
    On Error GoTo ErrLabIt1
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim lcSql As String, lnIdPaciente As Long
    
    lcSql = "select * from PerinatalAtencion where GrafXedadEnSemanas is null"
    oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          lnIdPaciente = oRsTmp1!idPaciente
          Set oRsTmp2 = mo_ReglasAdmision.PacientesSeleccionarPorIdentificador(lnIdPaciente, oConexion)
          If oRsTmp2.RecordCount > 0 Then
             oRsTmp1.Fields!GrafXedadEnSemanas = sighentidades.DevuelveEdadEnSemanas(oRsTmp2!FechaNacimiento, oRsTmp1!fechaAtencion)
             oRsTmp1.Update
          End If
          oRsTmp2.Close
          
          oRsTmp1.MoveNext
       Loop
       
    End If
    oRsTmp1.Close
ErrLabIt1:
    Set mo_ReglasAdmision = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing

End Sub

'debb-18/05/2016
Sub ActualizaEDAD_en_movimientos(oConexion As Connection, lnIdPaciente1 As Long)
    On Error GoTo ErrLabIt1
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim lcSql As String, lnIdPaciente As Long, lcDireccion As String
    Dim oEdad As Edad
    lcSql = "SELECT     dbo.ListBarItems.Texto, dbo.Auditoria.Accion, dbo.Auditoria.idListItem, dbo.Auditoria.IdRegistro, " & _
            "    dbo.Pacientes.FechaNacimiento " & _
            " FROM         dbo.Auditoria INNER JOIN " & _
            "                      dbo.ListBarItems ON dbo.Auditoria.idListItem = dbo.ListBarItems.IdListItem LEFT OUTER JOIN " & _
            "                      dbo.Pacientes ON dbo.Auditoria.IdRegistro = dbo.Pacientes.IdPaciente " & _
            " WHERE   (dbo.Auditoria.Accion = 'M') and   ((dbo.ListBarItems.IdListItem = 101)  OR " & _
            "                      (dbo.ListBarItems.IdListItem = 201) OR " & _
            "                      (dbo.ListBarItems.IdListItem = 301)) " & _
            IIf(lnIdPaciente1 = 0, "", " and dbo.Auditoria.IdRegistro=" & lnIdPaciente1) & _
            " order by dbo.Auditoria.IdRegistro"
    oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          lnIdPaciente = oRsTmp1!idRegistro
          lcSql = "select * from atenciones where idPaciente=" & lnIdPaciente
          oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
          If oRsTmp2.RecordCount > 0 Then
             oRsTmp2.MoveFirst
             Do While Not oRsTmp2.EOF
                oEdad = sighentidades.CalcularEdad(oRsTmp1!FechaNacimiento, oRsTmp2!FechaIngreso)
                oRsTmp2!IdTipoEdad = oEdad.TipoEdad
                oRsTmp2!Edad = oEdad.Edad
                oRsTmp2.Update
                oRsTmp2.MoveNext
             Loop
          End If
          oRsTmp2.Close
          Do While Not oRsTmp1.EOF And lnIdPaciente = oRsTmp1!idRegistro
                oRsTmp1.MoveNext
                If oRsTmp1.EOF Then
                   Exit Do
                End If
          Loop
       Loop
       
    End If
    oRsTmp1.Close
   '*** DIRECCION desde Pacientes
    If lnIdPaciente1 > 0 Then
   
      lcSql = "SELECT        dbo.Pacientes.DireccionDomicilio, dbo.AtencionesDatosAdicionales.DireccionDomicilio AS ddirecion, dbo.Atenciones.IdPaciente, dbo.Atenciones.IdAtencion" & _
" FROM            dbo.Atenciones LEFT OUTER JOIN" & _
"                         dbo.AtencionesDatosAdicionales ON dbo.Atenciones.IdAtencion = dbo.AtencionesDatosAdicionales.idAtencion LEFT OUTER JOIN" & _
"                         dbo.Pacientes ON dbo.Atenciones.IdPaciente = dbo.Pacientes.IdPaciente" & _
" Where (dbo.Pacientes.DireccionDomicilio Is Not Null) And (dbo.AtencionesDatosAdicionales.DireccionDomicilio Is Null)" & _
" and  dbo.Atenciones.idpaciente=" & lnIdPaciente1
      oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
      If oRsTmp1.RecordCount > 0 Then
         oRsTmp1.MoveFirst
         Do While Not oRsTmp1.EOF
            lcDireccion = oRsTmp1!DireccionDomicilio
            lcSql = "update AtencionesDatosAdicionales set Direcciondomicilio='" & lcDireccion & _
                    "' where idAtencion=" & oRsTmp1!idAtencion
            oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
            oRsTmp1.MoveNext
         Loop
      End If
      oRsTmp1.Close
    
    End If
    
    
ErrLabIt1:
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
End Sub
'debb-02/05/2016
Sub ActualizaLOTES_his_conNULL(oConexion As Connection)
    On Error GoTo ErrLabIt1
    Dim oRsTmp1 As New Recordset
    Dim mo_ReglasDeProgMedica As New ReglasDeProgMedica
    Dim oRsTmp2 As New Recordset
    Dim oMedicos As New Medicos
    Dim oDoMedico As New DOMedico
    Dim lcSql As String, lcLote As String, lcAsc As String, lcNumero As String, lcMedico As String
    Set oRsTmp2 = mo_ReglasDeProgMedica.MedicosSeleccionarTodos
    oRsTmp2.Filter = "lotehis<>null"
    Set oRsTmp1 = mo_ReglasDeProgMedica.MedicosSeleccionarTodos
    oRsTmp1.Filter = "lotehis=null"
    If oRsTmp1.RecordCount > 0 Then
       Set oMedicos.Conexion = oConexion
       lcAsc = "A": lcNumero = "0"
       lcLote = lcAsc + lcNumero
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          Do While True
                oRsTmp2.Filter = "lotehis='" & lcLote & "'"
                If oRsTmp2.RecordCount > 0 Then
                   lcLote = IncrementaLote(lcLote)
                Else
                   Exit Do
                End If
          Loop
          If Left(lcLote, 1) = "a" Then
             Exit Do
          End If
          oDoMedico.IdMedico = oRsTmp1!IdMedico
          oDoMedico.IdUsuarioAuditoria = sighentidades.Usuario
          If oMedicos.SeleccionarPorId(oDoMedico) Then
                oDoMedico.LoteHis = lcLote
                If oMedicos.modificar(oDoMedico) Then
                End If
          End If
          lcLote = IncrementaLote(lcLote)
          lcMedico = oRsTmp1!ApellidoPaterno + oRsTmp1!ApellidoMaterno + oRsTmp1!Nombres
          Do While Not oRsTmp1.EOF And lcMedico = oRsTmp1!ApellidoPaterno + oRsTmp1!ApellidoMaterno + oRsTmp1!Nombres
                oRsTmp1.MoveNext
                If oRsTmp1.EOF Then
                   Exit Do
                End If
          Loop
       Loop
    End If
ErrLabIt1:
    Set oRsTmp1 = Nothing
    Set mo_ReglasDeProgMedica = Nothing
    Set oRsTmp2 = Nothing
    Set oMedicos = Nothing
    Set oDoMedico = Nothing
End Sub
'debb-22/04/2016
Function IncrementaLote(lcLoteActual As String) As String
          Dim lcAsc As String, lcNumero As String
          lcNumero = Right(lcLoteActual, 1)
          lcAsc = Left(lcLoteActual, 1)
          lcNumero = Trim(Str(Val(lcNumero) + 1))
          If Val(lcNumero) > 9 Then
             lcNumero = "0"
             lcAsc = Chr(Asc(lcAsc) + 1)
             If UCase(lcAsc) = "V" Then
                lcAsc = Chr(Asc(lcAsc) + 1)
             End If
          End If
          If Asc(lcAsc) = 91 Then
             lcAsc = "a"
          ElseIf Asc(lcAsc) = 123 Then
             lcAsc = Chr(161)
          End If
          IncrementaLote = lcAsc & lcNumero
End Function


'debb-07/10/2015
Sub Fua2015conNULLenFuaTipoAnexo2015(oConexion As Connection, oConexionExterna As Connection)
    On Error GoTo ErrLabIt1
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oRsTmp3 As New Recordset
    Dim oRsTmp4 As New Recordset
    Dim oReglasFarmacia As New ReglasFarmacia
    Dim oReglasComunes As New ReglasComunes
    Dim lcSql As String
    lcSql = "select idCuentaAtencion from sisFuaAtencion where FuaVersionFormato <>'A' and FuaTipoAnexo2015 is null"
    oRsTmp1.Open lcSql, oConexionExterna, adOpenKeyset, adLockOptimistic
    If oRsTmp1.RecordCount > 0 Then
       Set oRsTmp2 = oReglasFarmacia.AtencionesSelecionarPorCuenta(oRsTmp1!idCuentaAtencion, oConexion)
       If oRsTmp2.RecordCount > 0 Then
          Set oRsTmp3 = oReglasComunes.ServiciosSeleccionarXidentificador(oRsTmp2!idServicioIngreso)
          If oRsTmp3.RecordCount > 0 Then
             lcSql = "update sisFuaAtencion set FuaTipoAnexo2015=" & oRsTmp3!FuaTipoAnexo2015 & " where idCuentaAtencion=" & oRsTmp1!idCuentaAtencion
             oRsTmp4.Open lcSql, oConexionExterna, adOpenKeyset, adLockOptimistic
          End If
          oRsTmp3.Close
       End If
       oRsTmp2.Close
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
ErrLabIt1:
End Sub

'debb-07/10/2015
'......Franklin.....convierte en PROCEDIMIENTO la CONSULTA del oRsTmp1
Sub ItemsBoletaFarmaciaConCantidadCERO(oConexion As Connection)
    On Error GoTo ErrLabIt1
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oReglasFarmacia As New ReglasFarmacia
    Dim oDoFacturacionBienesPagos As New DoFacturacionBienesPagos
    Dim oFacturacionBienesPagos As New FacturacionBienesPagos
    Dim lcSql As String
    lcSql = "SELECT     dbo.FacturacionBienesPagos.*,dbo.FactOrdenesBienes.idComprobantePago,  " & _
            "                       dbo.FactOrdenesBienes.MovNumero, dbo.FactOrdenesBienes.MovTipo" & _
            " FROM         dbo.FacturacionBienesPagos INNER JOIN" & _
            "                      dbo.FactOrdenesBienes ON dbo.FacturacionBienesPagos.IdOrden = dbo.FactOrdenesBienes.idOrden" & _
            " Where (dbo.FacturacionBienesPagos.CantidadPagar = 0) And (Not (dbo.FactOrdenesBienes.IdComprobantePago Is Null))"
    oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp1.RecordCount > 0 Then
       Set oFacturacionBienesPagos.Conexion = oConexion
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          Set oRsTmp2 = oReglasFarmacia.farmMovimientoVentasDetalleSeleccionarPorId(oRsTmp1!movNumero, oRsTmp1!MovTipo)
          oRsTmp2.Filter = "idProducto=" & oRsTmp1!idProducto
          If oRsTmp2.RecordCount > 0 Then
                oDoFacturacionBienesPagos.CantidadPagar = oRsTmp2!cantidad
                oDoFacturacionBienesPagos.IdOrden = oRsTmp1!IdOrden
                oDoFacturacionBienesPagos.idProducto = oRsTmp1!idProducto
                oDoFacturacionBienesPagos.IdUsuarioAuditoria = Val(sighentidades.Usuario)
                oDoFacturacionBienesPagos.PrecioVenta = oRsTmp2!Precio
                oDoFacturacionBienesPagos.TotalPagar = oRsTmp2!total
                If oFacturacionBienesPagos.ModificarXidOrdenIdProducto(oDoFacturacionBienesPagos) = True Then
                End If
          End If
          oRsTmp2.Close
          oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set oReglasFarmacia = Nothing
    Set oDoFacturacionBienesPagos = Nothing
    Set oFacturacionBienesPagos = Nothing
ErrLabIt1:
End Sub

'debb-15/09/2015
Sub CitasNOexistentes(oConexion As Connection)
   On Error GoTo ErrLabIt1
   Dim oRsTmp1 As New Recordset
   Dim oRsTmp2 As New Recordset
   Dim oRsTmp3 As New Recordset
   Dim oCommand As New ADODB.Command
   Dim oParameter As ADODB.Parameter
   Dim oCitas As New Citas, oDocitas As New doCita
   Dim mo_ReglasServiciosHosp As New ReglasServiciosHosp
   Dim mo_ReglasDeProgMedica As New ReglasDeProgMedica

   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CitasNOexistentes"
        Set oRsTmp1 = .Execute
   End With
   If oRsTmp1.RecordCount > 0 Then
      Set oCitas.Conexion = oConexion
      oRsTmp1.MoveFirst
      Do While Not oRsTmp1.EOF
         If oRsTmp1!IdMedico > 0 And Not IsNull(oRsTmp1!fechaRequerida) And oRsTmp1!IdServicio > 0 Then
             Set oRsTmp3 = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarPorMedicoFechaServicio(oRsTmp1!IdMedico, _
                                                            oRsTmp1!fechaRequerida, oRsTmp1!IdServicio)
             If oRsTmp3.RecordCount > 0 Then
               oRsTmp3.MoveFirst
               Do While Not oRsTmp3.EOF
                  If oRsTmp3!HoraInicio <= oRsTmp1!horaRequerida And oRsTmp3!horaFin >= oRsTmp1!horaRequerida Then
                     Exit Do
                  End If
                  oRsTmp3.MoveNext
               Loop
             End If
             If Not oRsTmp3.EOF Then
                 Set oRsTmp2 = mo_ReglasServiciosHosp.EspecialidadCESeleccionarPorIdServicio(oRsTmp1!IdServicio)
    
                 'oDocitas.EsCitaAdicional = oRsTmp1!
                 oDocitas.fecha = oRsTmp1!fechaRequerida
                 oDocitas.fechaSolicitud = oRsTmp1!fechaSolicitud
                 oDocitas.horaFin = mo_ReglasDeProgMedica.ConvertirAHora(mo_ReglasDeProgMedica.ConvertirAMinutos(oRsTmp1!horaRequerida) + oRsTmp2!TiempoPromedioAtencion)
                 oDocitas.HoraInicio = oRsTmp1!horaRequerida
                 oDocitas.horaSolicitud = oRsTmp1!horaSolicitud
                 oDocitas.idAtencion = oRsTmp1!idAtencion
                ' oDocitas.IdCita = oRsTmp1!
                 oDocitas.IdEspecialidad = oRsTmp2!IdEspecialidad
                 oDocitas.IdEstadoCita = 1
                 oDocitas.IdMedico = oRsTmp1!IdMedico
                 oDocitas.idPaciente = oRsTmp1!idPaciente
                 oDocitas.idProducto = oRsTmp2!IdProductoConsulta
                 oDocitas.idProgramacion = oRsTmp3!idProgramacion
                 oDocitas.IdServicio = oRsTmp1!IdServicio
                 oDocitas.IdUsuarioAuditoria = sighentidades.Usuario
                 If oCitas.Insertar(oDocitas) = True Then
                 End If
                 oRsTmp2.Close
                 
             End If
             oRsTmp3.Close
         End If
         oRsTmp1.MoveNext
      Loop
   End If
   oRsTmp1.Close
ErrLabIt1:
   Set oCitas = Nothing
   Set oDocitas = Nothing
   Set mo_ReglasServiciosHosp = Nothing
   Set mo_ReglasDeProgMedica = Nothing
   Set oRsTmp1 = Nothing
   Set oRsTmp2 = Nothing
   Set oRsTmp3 = Nothing
   Set oCommand = Nothing

End Sub

'debb-07/10/2015
Sub HistoriasSolicitadasNOexistentes(oConexion As Connection)
   On Error GoTo ErrLabIt1
   Dim oRsTmp1 As New Recordset
   Dim oCommand As New ADODB.Command
   Dim oParameter As ADODB.Parameter
   Dim oHistoriasSolicitadas As New HistoriasSolicitadas, oDOHistoriaSolicitada As New DOHistoriaSolicitada
   Dim lcSql As String
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HistoriasSolicitadasNOexistentes"
        Set oRsTmp1 = .Execute
   End With
   If oRsTmp1.RecordCount > 0 Then
      Set oHistoriasSolicitadas.Conexion = oConexion
      oRsTmp1.MoveFirst
      Do While Not oRsTmp1.EOF
         oDOHistoriaSolicitada.fechaRequerida = oRsTmp1!fecha
         oDOHistoriaSolicitada.fechaSolicitud = oRsTmp1!fechaSolicitud
         oDOHistoriaSolicitada.horaRequerida = oRsTmp1!HoraInicio
         oDOHistoriaSolicitada.horaSolicitud = oRsTmp1!horaSolicitud
         oDOHistoriaSolicitada.idAtencion = oRsTmp1!idAtencion
         'oDOHistoriaSolicitada.IdEmpleadoSolicita = oRsTmp1!
         'oDOHistoriaSolicitada.IdHistoriaSolicitada = oRsTmp1!
         oDOHistoriaSolicitada.IdMotivo = 1
         'oDOHistoriaSolicitada.IdMovimiento = oRsTmp1!
         oDOHistoriaSolicitada.idPaciente = oRsTmp1!idPaciente
         oDOHistoriaSolicitada.IdServicio = oRsTmp1!IdServicio
         oDOHistoriaSolicitada.IdUsuarioAuditoria = sighentidades.Usuario
         oDOHistoriaSolicitada.Observacion = "Cita"
         If oHistoriasSolicitadas.Insertar(oDOHistoriaSolicitada) = True Then
         End If
         oRsTmp1.MoveNext
      Loop
   End If
   oRsTmp1.Close
   lcSql = "      update HistoriasSolicitadas set IdMovimiento =" & _
            "                             (select top 1 IdMovimiento from MovimientosHistoriaClinica where" & _
            "                              MovimientosHistoriaClinica.IdMotivo =1 and" & _
            "                              MovimientosHistoriaClinica.IdMotivo   =HistoriasSolicitadas.IdMotivo   and" & _
            "                              MovimientosHistoriaClinica.IdServicioDestino  =HistoriasSolicitadas.IdServicio  and" & _
            "                              MovimientosHistoriaClinica.IdAtencion =HistoriasSolicitadas.IdAtencion and" & _
            "                              MovimientosHistoriaClinica.IdPaciente  =HistoriasSolicitadas.IdPaciente)" & _
            "      where HistoriasSolicitadas.IdMovimiento is null and" & _
            "                              (select top 1 IdMovimiento from MovimientosHistoriaClinica where" & _
            "                              MovimientosHistoriaClinica.IdMotivo =1 and" & _
            "                              MovimientosHistoriaClinica.IdMotivo   =HistoriasSolicitadas.IdMotivo   and" & _
            "                              MovimientosHistoriaClinica.IdServicioDestino  =HistoriasSolicitadas.IdServicio  and" & _
            "                              MovimientosHistoriaClinica.IdAtencion =HistoriasSolicitadas.IdAtencion and" & _
            "                              MovimientosHistoriaClinica.IdPaciente  =HistoriasSolicitadas.IdPaciente )>0"
   oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
   
ErrLabIt1:
   Set oHistoriasSolicitadas = Nothing
   Set oDOHistoriaSolicitada = Nothing
   Set oRsTmp1 = Nothing
   Set oCommand = Nothing
End Sub

Sub ActualizaEstablecimientosSIS(oConexion As Connection, oConexionSIS As Connection)
   On Error GoTo ErrLabIt1
   Dim oRsTmp1 As New Recordset
   Dim oRsTmp2 As New Recordset
   Dim oCommand As New ADODB.Command
   Dim oParameter As ADODB.Parameter
   Dim mo_reglasComunes As New ReglasComunes
   Dim oTiposTarifaCpt As New TiposTarifaCpt
   Dim oDOTiposTarifaCPT As New doTiposTarifaCpt
   Dim lnIdCpt As Long
   Dim lcCodigoRenaes As String, lcIdUbigeo As Long, lcIdCategoriaEESS As String, lcNombre As String
   Dim ml_idUsuario As Long, lnUltimoIdEstablecimiento As Long
   Dim oDOEstablecimiento As New DOEstablecimiento, oEstablecimiento As New Establecimientos
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionSIS
        .CommandTimeout = 150
        .CommandText = "EstablecimientosSinRENAESenSIS"
        Set oRsTmp1 = .Execute
   End With
   If oRsTmp1.RecordCount > 0 Then
      Set oRsTmp2 = mo_reglasComunes.EstablecimientosSeleccionarPorFiltro(" 1=1 order by idEstablecimiento desc", oConexion)
      lnUltimoIdEstablecimiento = 1
      If oRsTmp2.RecordCount > 0 Then
         lnUltimoIdEstablecimiento = oRsTmp2.Fields!IdEstablecimiento + 1
      End If
      Set oEstablecimiento.Conexion = oConexion
      oRsTmp1.MoveFirst
      Do While Not oRsTmp1.EOF
        lcCodigoRenaes = oRsTmp1!pre_CodigoRENAES
        lcIdUbigeo = oRsTmp1!pre_idUbigeo
        lcIdCategoriaEESS = oRsTmp1!pre_idCategoriaEESS
        lcNombre = oRsTmp1!pre_nombre
        oDOEstablecimiento.Codigo = Right(lcCodigoRenaes, 5)
        oDOEstablecimiento.IdDistrito = Val(lcIdUbigeo)
        oDOEstablecimiento.IdEstablecimiento = lnUltimoIdEstablecimiento
        oDOEstablecimiento.IdTipo = IIf(InStr(sighentidades.SISdevuelveCategoriaPS, lcIdCategoriaEESS) > 0, sghCategoriaEstablecimiento.sghPS, _
                            IIf(InStr(sighentidades.SISdevuelveCategoriaCS, lcIdCategoriaEESS) > 0, sghCategoriaEstablecimiento.sghCS, _
                            IIf(InStr(sighentidades.SISdevuelveCategoriaHospital, lcIdCategoriaEESS) > 0, sghCategoriaEstablecimiento.sghHospital, sghCategoriaEstablecimiento.sghOtros)))
                            '3=PS,2=CS,1=Hosp
        oDOEstablecimiento.IdUsuarioAuditoria = 0
        oDOEstablecimiento.nombre = lcNombre
        If oEstablecimiento.Insertar(oDOEstablecimiento) Then
        End If
        lnUltimoIdEstablecimiento = lnUltimoIdEstablecimiento + 1
        oRsTmp1.MoveNext
      Loop
   End If

   Set oRsTmp1 = Nothing
   Set oRsTmp2 = Nothing
   Set oCommand = Nothing
   Set mo_reglasComunes = Nothing
   Set oTiposTarifaCpt = Nothing
   Set oDOTiposTarifaCPT = Nothing
   Set oDOEstablecimiento = Nothing
   Set oEstablecimiento = Nothing
ErrLabIt1:

End Sub

Sub ActualizaIdEstanciaHospitalariaEnServios(oConexion As Connection)
   On Error GoTo ErrLabIt1
   Dim oRsTmp1 As New Recordset
   Dim oRsTmp2 As New Recordset
   Dim oCommand As New ADODB.Command
   Dim oParameter As ADODB.Parameter
   Dim mo_reglasComunes As New ReglasComunes
   Dim oTiposTarifaCpt As New TiposTarifaCpt
   Dim oDOTiposTarifaCPT As New doTiposTarifaCpt
   Dim lnIdCpt As Long
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosActualizaIdEstancia"
        Set oRsTmp1 = .Execute
   End With

   Set oRsTmp1 = Nothing
   Set oRsTmp2 = Nothing
   Set oCommand = Nothing
   Set mo_reglasComunes = Nothing
   Set oTiposTarifaCpt = Nothing
   Set oDOTiposTarifaCPT = Nothing
      
ErrLabIt1:

End Sub

Sub ActualizaTipoTarifaSinDetalle(oConexion As Connection)
    On Error GoTo ErrLabIt1
   Dim oRsTmp1 As New Recordset
   Dim oRsTmp2 As New Recordset
   Dim oCommand As New ADODB.Command
   Dim oParameter As ADODB.Parameter
   Dim mo_reglasComunes As New ReglasComunes
   Dim oTiposTarifaCpt As New TiposTarifaCpt
   Dim oDOTiposTarifaCPT As New doTiposTarifaCpt
   Dim lnIdCpt As Long
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposTarifaSinDetalle"
        Set oRsTmp1 = .Execute
   End With

   If oRsTmp1.RecordCount > 0 Then
      Set oRsTmp2 = mo_reglasComunes.TiposTarifaSeleccionarTodos(oConexion)
      lnIdCpt = 23
      If oRsTmp2.RecordCount > 0 Then
         oRsTmp2.MoveFirst
         oRsTmp2.Find "idTipoTarifa=" & lnIdCpt
         If oRsTmp2.EOF Then
            oRsTmp2.MoveFirst
            lnIdCpt = oRsTmp2.Fields!IdTipoTarifa
         End If
      End If
    
      
      Set oTiposTarifaCpt.Conexion = oConexion
      oRsTmp1.MoveFirst
      Do While Not oRsTmp1.EOF
         oDOTiposTarifaCPT.idProductoCPT = lnIdCpt
         oDOTiposTarifaCPT.IdTipoTarifa = oRsTmp1!IdTipoTarifa
         If oTiposTarifaCpt.Insertar(oDOTiposTarifaCPT) = True Then
         End If
         oRsTmp1.MoveNext
      Loop
   End If
   oRsTmp1.Close
   Set oRsTmp1 = Nothing
   Set oRsTmp2 = Nothing
   Set oCommand = Nothing
   Set mo_reglasComunes = Nothing
   Set oTiposTarifaCpt = Nothing
   Set oDOTiposTarifaCPT = Nothing
      
ErrLabIt1:
End Sub

Sub ActualizaLabResultadosXItemsENtablaLabPruebas(oConexion As Connection)
   On Error GoTo ErrLabIt
   Dim oRsTmp1 As New Recordset
   Dim oCommand As New ADODB.Command
   Dim oParameter As ADODB.Parameter
   Dim oDoLabPruebas As New DoLabPruebas
   Dim oLabPruebas As New LabPruebas
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioActualizaIdFuenteFinanciamientoPagantes"
        Set oRsTmp1 = .Execute
   End With
   Set oCommand = Nothing
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "LabResultadosItemsSINlabPruebas"
        Set oRsTmp1 = .Execute
   End With

   If oRsTmp1.RecordCount > 0 Then
      Set oLabPruebas.Conexion = oConexion
      oRsTmp1.MoveFirst
      Do While Not oRsTmp1.EOF
         oDoLabPruebas.codigoCPT = oRsTmp1.Fields!cpt
         oDoLabPruebas.codigoPrueba = "x"
         If oLabPruebas.Insertar(oDoLabPruebas) = True Then
         End If
         oRsTmp1.MoveNext
      Loop
   End If
   oRsTmp1.Close
   Set oRsTmp1 = Nothing
   Set oCommand = Nothing
   Set oDoLabPruebas = Nothing
   Set oLabPruebas = Nothing
ErrLabIt:
End Sub

Sub ActualizaUltimoNroExoneracion(oConexion As Connection)
   On Error GoTo ErrPqtes
   Dim oRsTmp1 As New Recordset
   Dim oCommand As New ADODB.Command
   Dim oParameter As ADODB.Parameter
   Dim lcAutogenerado As String, lnUltimoNroExoneracion As Long
    Dim oDOPArametro As New DOPArametro, oParametros As New Parametros
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionCuentasAtencionExonUltima"
        Set oRsTmp1 = .Execute
   End With
   lnUltimoNroExoneracion = 0
   If oRsTmp1.RecordCount > 0 Then
     lnUltimoNroExoneracion = oRsTmp1!NumeroExoneracion
   End If
   oRsTmp1.Close
   '
    Set oParametros.Conexion = oConexion
    oDOPArametro.IdParametro = 337  'Ultimo N° exoneracion
    If oParametros.SeleccionarPorId(oDOPArametro) Then
    End If
    oDOPArametro.valorTexto = Trim(Str(lnUltimoNroExoneracion))
    If Not oParametros.modificar(oDOPArametro) Then
       
    End If
   
   Set oRsTmp1 = Nothing
   Set oCommand = Nothing
   Set oDOPArametro = Nothing
   Set oParametros = Nothing
ErrPqtes:

End Sub




'Proceso: llena tabla 'RecetaDetalleItem'
'         elimina columna 'recetaCabecera.DocuentoDespacho'
'         llena con 2 a la columna 'recetadetalle.idestadoDetalle'
'         llena con 0 a la columna 'recetadetalle.idDosisRecetada'
'         genera union entre RecetaDetalle y RecetaDosis
'         genera union entre RecetaDetalle y RecetaEstadosDetalle
Sub ActualizaRecetasGeneradas(oConexion As Connection)
    On Error GoTo ErrActRec
    Dim oReglasComunes As New ReglasComunes
    Dim oRsTmp1 As New Recordset
    Dim oRecetaDetalleItem As New RecetaDetalleItem
    Dim oDoRecetaDetalleItem As New DoRecetaDetalleItem
    Dim lcSql As String, oConexion1 As New Connection
    'llena tabla 'RecetaDetalleItem'
    lcSql = "SELECT     dbo.RecetaCabecera.idReceta, dbo.RecetaDetalleItem.DocumentoDespacho AS DocDetalle, dbo.RecetaDetalle.CantidadDespachada, " & _
            "                      dbo.RecetaDetalle.idItem" & _
            " FROM         dbo.RecetaCabecera INNER JOIN" & _
            "                      dbo.RecetaDetalle ON dbo.RecetaCabecera.idReceta = dbo.RecetaDetalle.idReceta LEFT OUTER JOIN" & _
            "                      dbo.RecetaDetalleItem ON dbo.RecetaCabecera.idReceta = dbo.RecetaDetalleItem.idReceta" & _
            " WHERE     (dbo.RecetaDetalleItem.DocumentoDespacho IS NULL) AND (NOT (dbo.RecetaCabecera.DocumentoDespacho IS NULL)) AND" & _
            "                      (dbo.RecetaDetalle.CantidadDespachada > 0)" & _
            " ORDER BY dbo.RecetaDetalle.CantidadDespachada"
    oRsTmp1.Open lcSql, oConexion1, adOpenKeyset, adLockOptimistic
    If oRsTmp1.RecordCount > 0 Then
       Set oRecetaDetalleItem.Conexion = oConexion
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          oDoRecetaDetalleItem.CantidadDespachada = oRsTmp1!CantidadDespachada
          oDoRecetaDetalleItem.DocumentoDespacho = oRsTmp1!DocumentoDespacho
          oDoRecetaDetalleItem.idItem = oRsTmp1!idItem
          oDoRecetaDetalleItem.idReceta = oRsTmp1!idReceta
          oDoRecetaDetalleItem.IdUsuarioAuditoria = 0
          If oRecetaDetalleItem.Insertar(oDoRecetaDetalleItem) = True Then
          End If
          oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    'elimina columna 'recetaCabecera.DocuentoDespacho'
    oConexion1.CommandTimeout = 300
    oConexion1.Open "dsn=GALENHOS"
    lcSql = "alter table recetaCabecera drop column DocumentoDespacho"
    oRsTmp1.Open lcSql, oConexion1, adOpenKeyset, adLockOptimistic
    'llena con 2 a la columna 'recetadetalle.idestadoDetalle'
    lcSql = "update recetaDetalle set idEstadoDetalle=2 where idEstadoDetalle is null"
    oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
    'llena con 0 a la columna 'recetadetalle.idDosisRecetada'
    lcSql = "update recetaDetalle set idDosisRecetada=0 where idDosisRecetada is null"
    oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
    'genera union entre RecetaDetalle y RecetaDosis
    'genera union entre RecetaDetalle y RecetaEstadosDetalle
    lcSql = "ALTER TABLE [dbo].[RecetaDetalle] ADD " & _
            " CONSTRAINT [FK_RecetaDetalle_RecetaDosis] FOREIGN KEY" & _
            " (" & _
            "    [idDosisRecetada]" & _
            " ) REFERENCES [dbo].[RecetaDosis] (" & _
            "    [idDosis]" & _
            " )," & _
            " CONSTRAINT [FK_RecetaDetalle_RecetaEstadosDetalle] FOREIGN KEY" & _
            " (" & _
            "    [idEstadoDetalle]" & _
            " ) REFERENCES [dbo].[RecetaEstadosDetalle] (" & _
            "    [idEstadoDetalle]" & _
            " )"
    oRsTmp1.Open lcSql, oConexion1, adOpenKeyset, adLockOptimistic
ErrActRec:
    Set oReglasComunes = Nothing
    Set oRsTmp1 = Nothing
    Set oRecetaDetalleItem = Nothing
    Set oDoRecetaDetalleItem = Nothing
    Set oConexion1 = Nothing
End Sub

Sub ActualizaSeUsaSinPrecioNULL(oConexion As Connection)
   On Error GoTo ErrPqtes
   Dim oCommand As New ADODB.Command
   Dim oParameter As ADODB.Parameter
   Dim lnTiempoPromedioAtencion As Long
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosHospActualizaSeUsaSinPrecio"
        .Execute
    End With
    Set oCommand = Nothing
    Exit Sub
ErrPqtes:
End Sub

Sub ActualizaTiempoPromedioAtencion(oConexion As Connection)
   On Error GoTo ErrPqtes
   Dim oRsTmp1 As New Recordset
   Dim oRsTmp2 As New Recordset
   Dim oCommand As New ADODB.Command
   Dim oParameter As ADODB.Parameter
   Dim lnTiempoPromedioAtencion As Long
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProgramacionMedicaConTiempoPromedioAtencionNULL"
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
   
   If oRsTmp1.RecordCount > 0 Then
      oRsTmp1.MoveFirst
      Do While Not oRsTmp1.EOF
        Set oCommand = Nothing
        Set oParameter = Nothing
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "AtencionesSeleccionarXidProgramacion"
            Set oParameter = .CreateParameter("@idProgramacion", adInteger, adParamInput, 0, oRsTmp1.Fields!idProgramacion): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
        End With
        If oRsTmp2.RecordCount > 0 Then
            lnTiempoPromedioAtencion = DateDiff("n", CDate(oRsTmp2.Fields!HoraInicio), CDate(oRsTmp2.Fields!horaFin))
        Else
             oRsTmp2.Close
             Set oCommand = Nothing
             Set oParameter = Nothing
             With oCommand
                 .CommandType = adCmdStoredProc
                 Set .ActiveConnection = oConexion
                 .CommandTimeout = 150
                 .CommandText = "EspecialidadCEseleccionarIdServicio"
                 Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, oRsTmp1.Fields!IdServicio): .Parameters.Append oParameter
                 Set oRsTmp2 = .Execute
            End With
            If oRsTmp2.RecordCount > 0 Then
               lnTiempoPromedioAtencion = IIf(IsNull(oRsTmp2.Fields!TiempoPromedioAtencion), 0, oRsTmp2.Fields!TiempoPromedioAtencion)
            Else
               lnTiempoPromedioAtencion = 0
            End If
        End If
        Set oCommand = Nothing
        Set oParameter = Nothing
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ProgramacionMedicaActualizaTiempoPromedioAtencion"
            Set oParameter = .CreateParameter("@idProgramacion", adInteger, adParamInput, 0, oRsTmp1.Fields!idProgramacion): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@TiempoPromedioAtencion", adInteger, adParamInput, 0, lnTiempoPromedioAtencion): .Parameters.Append oParameter
             .Execute
       End With
       oRsTmp2.Close
       oRsTmp1.MoveNext
      Loop
   End If
   oRsTmp1.Close
   Set oRsTmp1 = Nothing
   Set oRsTmp2 = Nothing
   Set oCommand = Nothing
   Exit Sub
ErrPqtes:
   'MsgBox Err.Description
   'Resume
End Sub

Sub ActualizaNULLenIdPartidaIdCCostoCon999()
    On Error GoTo ErrActPc
    Dim oRsTmp As New Recordset
    Dim mo_ReglasFacturacion As New ReglasFacturacion
    Dim mo_reglasComunes As New ReglasComunes
    Dim lnIdPartida As Long
    '************Servicios**********
    Set oRsTmp = mo_reglasComunes.PartidasPresupuestalesSeleccionarTodos
    If oRsTmp.RecordCount > 0 Then
        lnIdPartida = 999
        oRsTmp.MoveFirst
        oRsTmp.Find "idPartidaPresupuestal=" & lnIdPartida
        If oRsTmp.EOF Then
           oRsTmp.MoveFirst
           lnIdPartida = oRsTmp!IdPartidaPresupuestal
        End If
        oRsTmp.Close
        Set oRsTmp = mo_ReglasFacturacion.FactCatalogoServiciosConPrecioMayor
        oRsTmp.Filter = "precio>0"
        If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
              If IsNull(oRsTmp!idPartida) Then
                 mo_reglasComunes.CatalogoServiciosActualizaPartidaXidProducto lnIdPartida, oRsTmp!idProducto, False
              End If
              If IsNull(oRsTmp!IdCentroCosto) Then
                 mo_ReglasFacturacion.FactCatalogoServiciosActualizaCentroCostoXproducto lnIdPartida, oRsTmp!idProducto
              End If
              oRsTmp.MoveNext
           Loop
        End If
    End If
    '
ErrActPc:
    Set oRsTmp = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_reglasComunes = Nothing
    
End Sub
'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Sub AtencionesConIdPacienteSinKeyEnTablaPacientes(oConexion As Connection)
   On Error GoTo ErrPqtes
   Dim oRsTmp1 As New Recordset
'   Dim lcSql As String
   Dim oRsTmp2 As New Recordset
   Dim oCommand As New ADODB.Command
   Dim oParameter As ADODB.Parameter
   

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarNroHistoriaClinicaNull"
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
   
   If oRsTmp1.RecordCount > 0 Then
      oRsTmp1.MoveFirst
      Do While Not oRsTmp1.EOF

        Set oCommand = Nothing
        Set oParameter = Nothing
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "PacientesFacturacionCuentasAtencionSeleccionarXIdCuentaAtencion"
            Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsTmp1.Fields!idCuentaAtencion): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
            Set oRsTmp2.ActiveConnection = Nothing
        End With
        If oRsTmp2.RecordCount > 0 Then
                 Set oCommand = Nothing
                 Set oParameter = Nothing
                 With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = oConexion
                    .CommandTimeout = 150
                    .CommandText = "AtencionesActualizaIdPaciente"
                    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsTmp1.Fields!idCuentaAtencion): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oRsTmp1.Fields!idPaciente): .Parameters.Append oParameter
                     .Execute
                End With
         End If
         oRsTmp2.Close
         oRsTmp1.MoveNext
      Loop
   End If
   oRsTmp1.Close
   Set oRsTmp1 = Nothing
   Set oRsTmp2 = Nothing
ErrPqtes:
End Sub

'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Sub AnulaPagosConsultaCEporPaquetes(oConexion As Connection)
   On Error GoTo ErrPqtes
   Dim oRsTmp1 As New Recordset
'   dim lcSql As String
   Dim oRsTmp2 As New Recordset
   Dim oCommand As New ADODB.Command
   Dim oParameter As ADODB.Parameter
   

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionPaquetesSeleccionaXIdPuntoCArga"
        Set oParameter = .CreateParameter("@IdPuntoCArga", adInteger, adParamInput, 0, sghPuntosCargaBasicos.sghPtoCargaAdmisionCE): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
   
   If oRsTmp1.RecordCount > 0 Then
      oRsTmp1.MoveFirst
      Do While Not oRsTmp1.EOF
        Set oCommand = Nothing
        Set oParameter = Nothing
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "AnulaPagosConsultaCEporPaquetes"
            Set oParameter = .CreateParameter("@Anulado", adInteger, adParamInput, 0, sghEstadoFacturacion.sghAnulado): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Atendido", adInteger, adParamInput, 0, sghEstadoFacturacion.sghAtendido): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsTmp1.Fields!atencionId): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdPuntoCArga", adInteger, adParamInput, 0, oRsTmp1.Fields!IdPuntoCarga): .Parameters.Append oParameter
            .Execute
        End With
        Set oParameter = Nothing
        oRsTmp1.MoveNext
      Loop
   End If
   oRsTmp1.Close
   Set oRsTmp1 = Nothing
   Set oRsTmp2 = Nothing
ErrPqtes:
End Sub

'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Sub ActualizaDxParaCE(oConexion As Connection)
   On Error GoTo ErrCE
   Const lcLineaChar As String = "¨"
   Dim lcBuscaParametro As New SIGHDatos.Parametros
   Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
   Dim oRsTmp1 As New Recordset, lcDxTexto As String, lnPos As Integer, lcCitaDiagMed As String
'   dim lcSql As String
   Dim oCommand As New ADODB.Command
   Dim oParameter As ADODB.Parameter


   
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesCESeleccionarCitaDiagMedNull"
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
      
   If oRsTmp1.RecordCount > 0 Then
      oRsTmp1.MoveFirst
      Do While Not oRsTmp1.EOF
         If Not IsNull(oRsTmp1.Fields!CitaDiagMed) Then
            lnPos = InStr(oRsTmp1.Fields!CitaDiagMed, lcLineaChar)
            lcDxTexto = ""
            If lnPos > 0 Then
               lcDxTexto = Mid(oRsTmp1.Fields!CitaDiagMed, lnPos + 1, 1000)
            End If
            lcCitaDiagMed = Left(mo_AdminFacturacion.DevuelveDxAltaMedicaTodosDx(oRsTmp1.Fields!idAtencion, 1, ""), 600) & Chr(13) & Chr(10) & lcLineaChar & lcDxTexto
            'oRsTmp1.Fields!CitaDiagMed = lcCitaDiagMed
            'oRsTmp1.Update
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "AtencionesCEactualizaDx"
                Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, oRsTmp1.Fields!idAtencion): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@citaDiagMed", adVarChar, adParamInput, 1000, lcCitaDiagMed): .Parameters.Append oParameter
                .Execute
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
         End If
         oRsTmp1.MoveNext
      Loop
   End If
   oRsTmp1.Close
   Set oRsTmp1 = Nothing
   Set oCommand = Nothing
   Set mo_AdminFacturacion = Nothing
   Set lcBuscaParametro = Nothing
ErrCE:
End Sub

'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Sub generadorNroHistoriaClinicaActualizaNroAutomaticoDeHistoriaClinica(oConexion As Connection)
    Dim oCommand As New ADODB.Command
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "generadorNroHistoriaClinicaActualizaNroAutomaticoDeHistoriaClinica"
        .Execute
    End With
    Set oCommand = Nothing
End Sub

'actualiza Cuentas ANULADAS en tabla ATENCIONES.idEstadoAtencion=0
'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Sub ActualizaCuentasAnuladasEnTablaAtenciones(oConexion As Connection)
       Dim mo_ReglasFacturacion  As New SIGHNegocios.ReglasFacturacion
       Dim lnFor As Long, lcHoraInicio As String, lcHoraActual As String
       Dim oRsTmp1 As New Recordset
'       dim oRsTmp2 As New Recordset
       Dim oCommand As New ADODB.Command
       Dim oParameter As ADODB.Parameter

        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "FacturacionCuentasAtencionSeleccionarPoridEstado9"
          Set oRsTmp1 = .Execute
          Set oRsTmp1.ActiveConnection = Nothing
        End With
        
       If oRsTmp1.RecordCount > 0 Then
          oRsTmp1.MoveFirst
          Do While Not oRsTmp1.EOF
              Set oCommand = Nothing
              Set oParameter = Nothing
              With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = oConexion
                  .CommandTimeout = 150
                  .CommandText = "ActualizarAtencionesIdEstadoAtencion0"
                  Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsTmp1.Fields!idCuentaAtencion): .Parameters.Append oParameter
                  .Execute
              End With
              Set oParameter = Nothing
              oRsTmp1.MoveNext
          Loop
       End If
       oRsTmp1.Close
       Set oRsTmp1 = Nothing
'       Set oRsTmp2 = Nothing
       Set mo_ReglasFacturacion = Nothing
End Sub

'Para  las cuentas con Tarifas=null o Planes=null (si estan anuladas, les asigna un valor)
'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Sub TarifasNULLplanesNULLlesAsignaValor(oConexion As Connection)
       Dim mo_ReglasFacturacion  As New SIGHNegocios.ReglasFacturacion
       Dim lnFor As Long, lcHoraInicio As String, lcHoraActual As String
       Dim oCommand As New ADODB.Command
       Dim oParameter As ADODB.Parameter
       Dim oRsTmp1 As New Recordset, oRsTmp2 As New Recordset
       
'       Dim lcSql As String
       Dim lnIdFormaPago As Long, lnIdFuenteFinanciamiento As Long

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "AtencionesFacturacionCuentasAtencionidFuenteFinanciamientoNull"
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
       
       If oRsTmp1.RecordCount > 0 Then
          oRsTmp1.MoveFirst
          Do While Not oRsTmp1.EOF
             lnIdFormaPago = IIf(IsNull(oRsTmp1.Fields!IdFormaPago), 0, oRsTmp1.Fields!IdFormaPago)
             lnIdFuenteFinanciamiento = IIf(IsNull(oRsTmp1.Fields!idFuenteFinanciamiento), 0, oRsTmp1.Fields!idFuenteFinanciamiento)
            Set oCommand = Nothing
            Set oParameter = Nothing
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "FactOrdenServicioSeleccionarXIdCuentaAtencion"
                Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsTmp1.Fields!idCuentaAtencion): .Parameters.Append oParameter
                Set oRsTmp2 = .Execute
                Set oRsTmp2.ActiveConnection = Nothing
            End With
            Set oParameter = Nothing
             
            If oRsTmp2.RecordCount > 0 Then
                If Not IsNull(oRsTmp2.Fields!IdTipoFinanciamiento) And lnIdFormaPago = 0 Then
                   lnIdFormaPago = oRsTmp2.Fields!IdTipoFinanciamiento
                End If
                If Not IsNull(oRsTmp2.Fields!idFuenteFinanciamiento) And lnIdFuenteFinanciamiento = 0 Then
                   lnIdFuenteFinanciamiento = oRsTmp2.Fields!idFuenteFinanciamiento
                End If
             End If
             oRsTmp2.Close
             
             Set oCommand = Nothing
             Set oParameter = Nothing
             With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = oConexion
                    .CommandTimeout = 150
                    .CommandText = "TarifasNULLplanesNULLlesAsignaValor"
                    Set oParameter = .CreateParameter("@IdFormaPago", adInteger, adParamInput, 0, lnIdFormaPago): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, oRsTmp1.Fields!IdEstado): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdEstadoAtencion", adInteger, adParamInput, 0, oRsTmp1.Fields!IdEstadoAtencion): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsTmp1.Fields!idCuentaAtencion): .Parameters.Append oParameter
                    .Execute
             End With
             Set oParameter = Nothing
            oRsTmp1.MoveNext
          Loop
       End If
       oRsTmp1.Close
       Set oRsTmp1 = Nothing
       Set oRsTmp2 = Nothing
       Set mo_ReglasFacturacion = Nothing
End Sub

'Actualiza SERVICIO DONDE ESTA EL PACIENTE con Valor NULL, En Tablas Servicios y Farmacia
'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
 Sub ActualizaIdServicioNULL(oConexion As Connection)
    Dim oRsTmp As New Recordset
'    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim lnIdServicioPaciente As Long
'    Dim lcSql As String
    'Servicios
    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "FactOrdenServicioIdServicioPacienteNull"
      Set oRsTmp = .Execute
      Set oRsTmp.ActiveConnection = Nothing
    End With
  
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          lnIdServicioPaciente = mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(oRsTmp.Fields!idCuentaAtencion, CDate(Format(oRsTmp.Fields!fechaDespacho, "dd/mm/yyyy")), Format(oRsTmp.Fields!fechaDespacho, "hh:mm"))
          If lnIdServicioPaciente = 0 Then
             lnIdServicioPaciente = oRsTmp.Fields!idServicioIngreso
          End If
          Set oCommand = Nothing
          Set oParameter = Nothing
          With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = oConexion
              .CommandTimeout = 150
              .CommandText = "ActualizaIdServicioNULL"
              Set oParameter = .CreateParameter("@IdServicioPaciente", adInteger, adParamInput, 0, lnIdServicioPaciente): .Parameters.Append oParameter
              Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, oRsTmp.Fields!IdOrden): .Parameters.Append oParameter
              .Execute
          End With
          Set oParameter = Nothing
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    'Farmacia
    Set oCommand = Nothing
    Set oParameter = Nothing
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "FarmMovimientoVentasSeleccionarXIdServicioPacienteNull"
         Set oRsTmp = .Execute
         Set oRsTmp.ActiveConnection = Nothing
    End With
        
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          lnIdServicioPaciente = mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(oRsTmp.Fields!idCuentaAtencion, CDate(Format(oRsTmp.Fields!FechaCreacion, "dd/mm/yyyy")), Format(oRsTmp.Fields!FechaCreacion, "hh:mm"))
          If lnIdServicioPaciente = 0 Then
             lnIdServicioPaciente = oRsTmp.Fields!idServicioIngreso
          End If
          With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "ActualizaFarmMovimientoVentasXmovNumero"
             Set oParameter = .CreateParameter("@IdServicioPaciente", adInteger, adParamInput, 0, lnIdServicioPaciente): .Parameters.Append oParameter
             Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, oRsTmp.Fields!movNumero): .Parameters.Append oParameter
             Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, oRsTmp.Fields!MovTipo): .Parameters.Append oParameter
             .Execute
          End With
          Set oParameter = Nothing
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
'    Set oRsTmp1 = Nothing
    Set mo_ReglasFarmacia = Nothing
End Sub

'busca N° Historia para Pacientes con Historia=NULL
'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Private Sub ActualizaNullhistoriaClinica(oConexion As Connection)
    Dim oPacientes As New SIGHDatos.Pacientes
    Set oPacientes.Conexion = oConexion
    oPacientes.ActualizaCorrelativosDeNroHistoria
    Set oPacientes = Nothing
'    Dim oRsTmp As New Recordset
'    Dim oRsTmp1 As New Recordset
'    Dim oCommand As New ADODB.Command
'    Dim oParameter As ADODB.Parameter
'    Dim lnIdPaciente As Long
'
'    With oCommand
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = oConexion
'        .CommandTimeout = 150
'        .CommandText = "PacientesSeleccionarXTipoNumeracionHistoriaClinica"
'        Set oRsTmp = .Execute
'        Set oRsTmp.ActiveConnection = Nothing
'    End With
'
'    If oRsTmp.RecordCount > 0 Then
'       oRsTmp.MoveFirst
'       Do While Not oRsTmp.EOF
'
'          Set oRsTmp1 = HistoriasClinicasXIdPaciente(oRsTmp.Fields!idPaciente, oConexion)
'
'          If oRsTmp1.RecordCount > 0 Then
'                Set oCommand = Nothing
'                Set oParameter = Nothing
'                With oCommand
'                      .CommandType = adCmdStoredProc
'                      Set .ActiveConnection = oConexion
'                      .CommandTimeout = 150
'                      .CommandText = "PacientesActualizarNroHistoriaYTipoNumeracion"
'                      Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, oRsTmp1.Fields!NroHistoriaClinica): .Parameters.Append oParameter
'                      Set oParameter = .CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, oRsTmp.Fields!IdTipoNumeracion): .Parameters.Append oParameter
'                      Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oRsTmp.Fields!idPaciente): .Parameters.Append oParameter
'                      .Execute
'                End With
'
'
'
'          End If
'          oRsTmp1.Close
'          oRsTmp.MoveNext
'       Loop
'    End If
'    oRsTmp.Close
'    Set oRsTmp = Nothing
'    Set oRsTmp1 = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Sub ActualizaIdRecienNacido(lnIdAtencion As Long, oConexion As Connection)
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lnIdPaciente As Long
    Dim lcSql As String
    

    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "AtencionesDatosAdicionalesSeleccionar1eraFila"
      Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
      Set oRsTmp = .Execute
      Set oRsTmp.ActiveConnection = Nothing
    End With
    Set oParameter = Nothing
    
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF

           Set oCommand = Nothing
           Set oParameter = Nothing
           With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "PacientesAtencionesSeleccionarPorID"
                Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, oRsTmp.Fields!idAtencion): .Parameters.Append oParameter
                Set oRsTmp1 = .Execute
                Set oRsTmp1.ActiveConnection = Nothing
           End With
           Set oParameter = Nothing
          
          If oRsTmp1.RecordCount > 0 Then
                Set oCommand = Nothing
                Set oParameter = Nothing
                With oCommand
                     .CommandType = adCmdStoredProc
                     Set .ActiveConnection = oConexion
                     .CommandTimeout = 150
                     .CommandText = "AtencionesDatosAdicionalesActualizaRN"
                     Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, oRsTmp.Fields!idAtencion): .Parameters.Append oParameter
                     If sighentidades.CalculaSiEsRecienNacido(oRsTmp1.Fields!FechaNacimiento, CDate(oRsTmp1.Fields!FechaIngreso & " " & oRsTmp1.Fields!HoraIngreso)) = 1 Then
                        Set oParameter = .CreateParameter("@RecienNacido", adBoolean, adParamInput, 0, 1): .Parameters.Append oParameter
                     Else
                        Set oParameter = .CreateParameter("@RecienNacido", adBoolean, adParamInput, 0, 0): .Parameters.Append oParameter
                     End If
                      .Execute
                End With
          End If
          oRsTmp1.Close
          
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    Set oRsTmp = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 16/10/2013 - se cambio a store procedure
Sub CierraPreVentasFarmaciaDelDiaAnterior(oConexion As Connection)
'    Dim oRsTmp1 As New Recordset
'    Dim oRsTmp2 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CierraPreVentasFarmaciaDelDiaAnterior"
        Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, lcBuscaParametro.RetornaFechaServidorSQL): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@HoraCreacion", adChar, adParamInput, 5, lcBuscaParametro.RetornaHoraServidorSQL): .Parameters.Append oParameter
        .Execute
    End With
    Set oCommand = Nothing
    Set lcBuscaParametro = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 15/10/2013 - se cambio a store procedure
Sub ActualizaFechaEgresoEnCEparaLosQtienenDx(oConexion As Connection)
    Dim oRsTmp1 As New Recordset
'    Dim oRsTmp2 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
  

    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "ConsultaFechaEgresoEnCEparaLosQtienenDx"
      Set oRsTmp1 = .Execute
      Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set oCommand = Nothing
  
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
            Set oCommand = Nothing
            Set oParameter = Nothing
            With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = oConexion
                  .CommandTimeout = 150
                  .CommandText = "ActualizaFechaEgresoEnCEparaLosQtienenDx"
                  Set oParameter = .CreateParameter("@horaEgreso", adChar, adParamInput, 5, Format(Time, "hh:mm")): .Parameters.Append oParameter
                  Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oRsTmp1.Fields!idAtencion): .Parameters.Append oParameter
                  .Execute
            End With
            Set oParameter = Nothing
          oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
End Sub

Sub ActualizaIdRecienNacidoEnTablaAtenciones(lnIdAtencion As Long)
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    ActualizaIdRecienNacido lnIdAtencion, oConexion
    oConexion.Close
    Set oConexion = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 15/10/2013 - se cambio a store procedure
Function ServiciosSeleccionarSoloCE() As Recordset
    Dim oRsTmp As New Recordset
'    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
  
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    

    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "ServiciosSeleccionarSoloCE"
      Set oRsTmp = .Execute
      Set oRsTmp.ActiveConnection = Nothing
    End With
    Set ServiciosSeleccionarSoloCE = oRsTmp
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 15/10/2013 - se cambio a store procedure
Function ServicioEsPuntoCarga(lnIdServicio As Long) As Boolean
    Dim oRsTmp As New Recordset
'    Dim lcSql As String
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
      .CommandText = "ServicioEsPuntoCarga"
      Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
      Set oRsTmp = .Execute
      Set oRsTmp.ActiveConnection = Nothing
    End With
  
    ServicioEsPuntoCarga = IIf(oRsTmp.RecordCount > 0, True, False)
    oRsTmp.Close
    Set oRsTmp = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 15/10/2013 - se cambio a store procedure
Sub ServicioSeUsanModulosPerinatalyMaterno(lnIdCita As Long, ByRef lbElConsultorioUsaModuloPerinatal As Boolean, _
                                           ByRef lbElConsultorioUsaModuloMaterno As Boolean, _
                                           ByRef lbElConsultorioUsaFUA As Boolean)
    Dim oRsTmp As New Recordset
'    Dim lcSql As String
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
      .CommandText = "ServicioSeUsanModulosPerinatalyMaterno"
      Set oParameter = .CreateParameter("@IdCita", adInteger, adParamInput, 0, lnIdCita): .Parameters.Append oParameter
      Set oRsTmp = .Execute
      Set oRsTmp.ActiveConnection = Nothing
    End With
  
    lbElConsultorioUsaModuloPerinatal = False
    lbElConsultorioUsaModuloMaterno = False
    lbElConsultorioUsaFUA = False
    If oRsTmp.RecordCount > 0 Then
       If oRsTmp.Fields!UsaModuloNinoSano = True Then
          lbElConsultorioUsaModuloPerinatal = True
       End If
       If oRsTmp.Fields!UsaModuloMaterno = True Then
          lbElConsultorioUsaModuloMaterno = True
       End If
       If oRsTmp.Fields!UsaFua = True Then
          lbElConsultorioUsaFUA = True
       End If
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 15/10/2013 - se cambio a store procedure
Function ServicioUsaGalenHos(lnIdServicio As Long) As String
    Dim oRsTmp As New Recordset
'    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
  
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    ServicioUsaGalenHos = "N"
    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "ServicioUsaGalenHos"
      Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
      Set oRsTmp = .Execute
      Set oRsTmp.ActiveConnection = Nothing
    End With
    
    If oRsTmp.RecordCount > 0 Then
       If Not IsNull(oRsTmp.Fields!UsaGalenHos) Then
          ServicioUsaGalenHos = IIf(oRsTmp.Fields!UsaGalenHos = True, "F", "S")
       End If
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 15/10/2013 - se cambio a store procedure
Function MovimientosHistoriaClinicaSeleccionarPorIdentificador(lnIdMovimiento As Long, Optional oConexion1 As Connection) As Recordset
'   Dim lcSql As String
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexionShape As New ADODB.Connection
    Set MovimientosHistoriaClinicaSeleccionarPorIdentificador = Nothing
    If oConexion1 Is Nothing Then
        oConexionShape.CursorLocation = adUseClient
        oConexionShape.CommandTimeout = 300
        oConexionShape.Open sighentidades.CadenaConexionShape
    End If

    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexionShape
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "MovimientosHistoriaClinicaSeleccionarPorIdentificador"
        Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set MovimientosHistoriaClinicaSeleccionarPorIdentificador = oRsTmp1
    
    Set oConexionShape = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 15/10/2013 - se cambio a store procedure
Function MovimientosFormatosSeleccionarPorIdentificador(lnIdMovimientoF As Long, Optional oConexion1 As Connection) As Recordset
'   Dim lcSql As String
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexionShape As New ADODB.Connection
    If oConexion1 Is Nothing Then
        oConexionShape.CursorLocation = adUseClient
        oConexionShape.CommandTimeout = 300
        oConexionShape.Open sighentidades.CadenaConexionShape
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexionShape
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "MovimientosFormatosSeleccionarPorIdentificador"
        Set oParameter = .CreateParameter("@IdMovimientoF", adInteger, adParamInput, 0, lnIdMovimientoF): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set MovimientosFormatosSeleccionarPorIdentificador = oRsTmp1
    Set oConexionShape = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 15/10/2013 - se cambio a store procedure
Function HistoriaClinicasSeleccionarXnrohistoria(lnNroHistoriaClinica As Long, Optional oConexion1 As Connection) As Recordset
'   Dim lcSql As String
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexionShape As New ADODB.Connection
    If oConexion1 Is Nothing Then

        oConexionShape.CursorLocation = adUseClient
        oConexionShape.CommandTimeout = 300
        oConexionShape.Open sighentidades.CadenaConexionShape
    End If

    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexionShape
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "HistoriaClinicasSeleccionarXnrohistoria"
        Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    
    Set HistoriaClinicasSeleccionarXnrohistoria = oRsTmp1
    Set oConexionShape = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 15/10/2013 - se cambio a store procedure
Function ServiciosSeleccionarXidentificador(lnIdServicio As Long) As Recordset
'   Dim lcSql As String
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Set ServiciosSeleccionarXidentificador = Nothing

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosSeleccionarXidentificador"
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    
    Set ServiciosSeleccionarXidentificador = oRsTmp1
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 15/10/2013 - se cambio a store procedure
Function HistoriasSolicitadasXidentificador(lnIdMovimiento As Long) As Recordset
'   Dim lcSql As String
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Set HistoriasSolicitadasXidentificador = Nothing

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HistoriasSolicitadasXidentificador"
        Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    
    Set HistoriasSolicitadasXidentificador = oRsTmp1
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 15/10/2013 - se cambio a store procedure
Function GeneradorNroHistoriaClinicaXidTipoNumeracion(lnIdTipoNumeracion As Long, oConexion As Connection) As Recordset
'   Dim lcSql As String
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set GeneradorNroHistoriaClinicaXidTipoNumeracion = Nothing
    

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "GeneradorNroHistoriaClinicaXidTipoNumeracion"
        Set oParameter = .CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, lnIdTipoNumeracion): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    
    Set GeneradorNroHistoriaClinicaXidTipoNumeracion = oRsTmp1
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 15/10/2013 - se cambio a store procedure
Function HistoriaClinicasXidentificador(lnNroHistoriaClinica As Long, oConexion As Connection) As Recordset
'   Dim lcSql As String
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set HistoriaClinicasXidentificador = Nothing
    

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HistoriaClinicasXidentificador"
        Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    
    Set HistoriaClinicasXidentificador = oRsTmp1
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 15/10/2013 - se cambio a store procedure
Function PacientesConHistoriasTemporales() As Recordset
'   Dim lcSql As String
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Set PacientesConHistoriasTemporales = Nothing

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesConHistoriasTemporales"
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    
    Set PacientesConHistoriasTemporales = oRsTmp1
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 15/10/2013 - se cambio a store procedure
Function MovimientosHistoriaClinicaNoLlegaAlArchivoClinicoDebb(mda_FechaInicio As String, mda_FechaFin As String) As Recordset
'   Dim lcSql As String
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Set MovimientosHistoriaClinicaNoLlegaAlArchivoClinicoDebb = Nothing

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion


    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "MovimientosHistoriaClinicaNoLlegaAlArchivoClinico"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, mda_FechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, mda_FechaFin): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    
    Set MovimientosHistoriaClinicaNoLlegaAlArchivoClinicoDebb = oRsTmp1
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

Function HistoriasClinicasSeleccionarPorAnios(ml_EdadMaxima As Long) As ADODB.Recordset

    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim oRsTmp2 As New Recordset
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "PacientesMenoresAunaEdad"
            Set oParameter = .CreateParameter("@ml_EdadMaxima", adInteger, adParamInput, 0, ml_EdadMaxima - 1): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
            Set oRsTmp2.ActiveConnection = Nothing
       End With
       Set HistoriasClinicasSeleccionarPorAnios = oRsTmp2
       oConexion.Close
       Set oRsTmp2 = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
       Exit Function
ManejadorDeError:
       MsgBox Err.Description
End Function

Function PacientesSeleccionarPorSectorComunidad(lcSector As String, lcComunidad As String) As Recordset
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim oRsTmp2 As New Recordset
    Dim lcSectorComunidad As String
        lcSectorComunidad = Trim(lcSector) & "-"
        If lcComunidad <> "" Then
           lcSectorComunidad = lcSectorComunidad & Trim(lcComunidad) & "-"
        End If
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "PacientesSeleccionarPorSectorComunidad"
            Set oParameter = .CreateParameter("@lcSectorComunidad", adVarChar, adParamInput, 30, lcSectorComunidad): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Left", adInteger, adParamInput, 0, Len(lcSectorComunidad)): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
            Set oRsTmp2.ActiveConnection = Nothing
       End With
       Set PacientesSeleccionarPorSectorComunidad = oRsTmp2
       oConexion.Close
       Set oRsTmp2 = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
       Exit Function
ManejadorDeError:
       MsgBox Err.Description
End Function

Function HistoriasClinicasXIdPaciente(lnIdPaciente As Long, oConexion As Connection) As Recordset
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRsTmp1 As New Recordset
    With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "HistoriasClinicasXIdPaciente"
          Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
          Set oRsTmp1 = .Execute
    End With
    Set HistoriasClinicasXIdPaciente = oRsTmp1
    Set oCommand = Nothing
    Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Sub ActualizaCampoDxSINpunto(oConexion As Connection)
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRsTmp As New Recordset
        With oCommand
                 .CommandType = adCmdStoredProc
                 Set .ActiveConnection = oConexion
                 .CommandTimeout = 150
                 .CommandText = "DiagnosticosSeleccionarTodos"
                 Set oRsTmp = .Execute
                 Set oRsTmp.ActiveConnection = Nothing
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
        
        If oRsTmp.RecordCount >= 0 Then
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
              
                With oCommand
                     .CommandType = adCmdStoredProc
                     Set .ActiveConnection = oConexion
                     .CommandTimeout = 150
                     .CommandText = "DiagnosticosModificarCodigoCIEsinPto"
                     Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, oRsTmp.Fields!IdDiagnostico): .Parameters.Append oParameter
                     Set oParameter = .CreateParameter("@codigoCIEsinPto", adVarChar, adParamInput, 7, Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCIE2004), 7)): .Parameters.Append oParameter
                     .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
                        
              oRsTmp.MoveNext
           Loop
        End If
    Set oCommand = Nothing
    Exit Sub
ManejadorDeError:
    MsgBox Err.Description

End Sub


Sub ActualizaEpisodioClinico(oConexion As Connection)
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim mo_ReglasAdmision As New ReglasAdmision
    Dim lnIdEpisodio As Long
    Dim oDoAtencionesEpisodiosCab As New DoAtencionesEpisodiosCab, oAtencionesEpisodiosCab As New AtencionesEpisodiosCab
    Dim oDoAtencionesEpisodiosDet As New DoAtencionesEpisodiosDet, oAtencionesEpisodiosDet As New AtencionesEpisodiosDet
    Set oRsTmp1 = mo_ReglasAdmision.AtencionesSeleccionarSoloNULLenEpisodiosDetalle(oConexion)
    If oRsTmp1.RecordCount > 0 Then
       Set oAtencionesEpisodiosCab.Conexion = oConexion
       Set oAtencionesEpisodiosDet.Conexion = oConexion
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
                Set oRsTmp2 = mo_ReglasAdmision.AtencionesEpisodiosCabeceraSeleccionarPorId(oRsTmp1!idPaciente, oConexion)
                If oRsTmp2.RecordCount = 0 Then
                   lnIdEpisodio = 1
                Else
                   oRsTmp2.MoveLast
                   lnIdEpisodio = oRsTmp2!idEpisodio + 1
                End If
                oRsTmp2.Close
                oDoAtencionesEpisodiosCab.FechaApertura = oRsTmp1!FechaIngreso
                oDoAtencionesEpisodiosCab.FechaCierre = oRsTmp1!fechaEgreso
                oDoAtencionesEpisodiosCab.idEpisodio = lnIdEpisodio
                oDoAtencionesEpisodiosCab.idPaciente = oRsTmp1!idPaciente
                oDoAtencionesEpisodiosCab.IdUsuarioAuditoria = 1
                If oAtencionesEpisodiosCab.Insertar(oDoAtencionesEpisodiosCab) = False Then
                   ms_MensajeError = oAtencionesEpisodiosCab.MensajeError
                End If
                oDoAtencionesEpisodiosDet.idAtencion = oRsTmp1!idAtencion
                oDoAtencionesEpisodiosDet.idEpisodio = lnIdEpisodio
                oDoAtencionesEpisodiosDet.idPaciente = oRsTmp1!idPaciente
                oDoAtencionesEpisodiosDet.IdUsuarioAuditoria = 1
                If oAtencionesEpisodiosDet.Insertar(oDoAtencionesEpisodiosDet) = False Then
                   ms_MensajeError = oAtencionesEpisodiosDet.MensajeError
                End If
          oRsTmp1.MoveNext
       Loop
    End If
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set mo_ReglasAdmision = Nothing
    Set oDoAtencionesEpisodiosCab = Nothing
    Set oAtencionesEpisodiosCab = Nothing
    Set oDoAtencionesEpisodiosDet = Nothing
    Set oAtencionesEpisodiosDet = Nothing
End Sub



Function HistoriaClinicaEsNueva(lnNroHistoriaClinica As Long, lnIdTipoNumeracion As Long, Optional oConexion1 As Connection) As String
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRsTmp1 As New Recordset
    Dim lcEsNueva As String
    Dim oConexion As New ADODB.Connection
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
          .CommandText = "HistoriaClinicaEsNueva"
          Set oParameter = .CreateParameter("@esNueva", adVarChar, adParamOutput, 3, lcEsNueva): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@idTipoNumeracion", adInteger, adParamInput, 0, lnIdTipoNumeracion): .Parameters.Append oParameter
          .Execute
          HistoriaClinicaEsNueva = .Parameters("@esNueva")
    End With
    Set oCommand = Nothing
    Set oConexion = Nothing
    Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function HistoriasClinicasSeleccionarTodosFcreacion(lcAnioCreacionHC As String, oConexion As Connection) As Recordset
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRsTmp1 As New Recordset
    With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "HistoriasClinicasSeleccionarTodosFcreacion"
          Set oParameter = .CreateParameter("@lnAnio", adInteger, adParamInput, 0, Val(lcAnioCreacionHC)): .Parameters.Append oParameter
          Set oRsTmp1 = .Execute
    End With
    Set HistoriasClinicasSeleccionarTodosFcreacion = oRsTmp1
    Set oCommand = Nothing
    Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



'debb2014b
Sub ActualizaAUTOGENERADOconHORA(oConexion As Connection)
   On Error GoTo ErrPqtes
   Dim oRsTmp1 As New Recordset
   Dim oCommand As New ADODB.Command
   Dim oParameter As ADODB.Parameter
   Dim oDoPaciente As New DOPaciente, oPacientes As New Pacientes
   Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
   Dim lcAutogenerado As String
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesAdicionaAutogeneradoHoraNacimiento"
        .Execute
   End With
   'Actualiza AUTOGENERADO=NULL           debb2014b
   Set oRsTmp1 = PacientesAutogeneradoIGUALnull
   If oRsTmp1.RecordCount > 0 Then
      Set oPacientes.Conexion = oConexion
      oDoPaciente.IdUsuarioAuditoria = sighentidades.Usuario
      oRsTmp1.MoveFirst
      Do While Not oRsTmp1.EOF
         oDoPaciente.idPaciente = oRsTmp1!idPaciente
         If oPacientes.SeleccionarPorId(oDoPaciente) = True Then
            oDoPaciente.autogenerado = mo_AdminAdmision.PacienteCrearNroAutogenerado(oDoPaciente)
            If oPacientes.modificar(oDoPaciente, False) = True Then
            End If
         End If
         oRsTmp1.MoveNext
      Loop
   End If
   oRsTmp1.Close
   '
   Set oRsTmp1 = Nothing
   Set oCommand = Nothing
   Set oDoPaciente = Nothing
   Set oPacientes = Nothing
   Set mo_AdminAdmision = Nothing
ErrPqtes:
'Resume
End Sub
'debb2014b
Sub ActualizaHistoricoDePreciosFarmacia(oConexion As Connection)
    On Error GoTo errAcHistPr
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    Dim oDoFarmHistPrecio As New DoFarmHistPrecio, oFarmHistPrecio As New FarmHistPrecio
    Dim ldFechaMod As Date
    Set oRsTmp2 = mo_ReglasFarmacia.FarmHistPrecioSeleccionarPorCodigo("", "")
    If oRsTmp2.RecordCount = 0 Then
        Set oRsTmp = mo_ReglasFarmacia.ExportaPreciosSismed(oConexion)
        If oRsTmp.RecordCount > 0 Then
            oDoFarmHistPrecio.idUsuario = sighentidades.Usuario
            oDoFarmHistPrecio.IdUsuarioAuditoria = sighentidades.Usuario
            Set oFarmHistPrecio.Conexion = oConexion
            oRsTmp.MoveFirst
            Do While Not oRsTmp.EOF
                ldFechaMod = Now
                Set oRsTmp1 = mo_ReglasFarmacia.ExportaPreciosSismedRegSant(oRsTmp.Fields!idProducto, oConexion)
                Set oRsTmp2 = mo_ReglasFarmacia.ExportaPreciosSismedFecha(oRsTmp.Fields!idProducto, oConexion)
                If oRsTmp2.RecordCount > 0 Then
                    If Not IsNull(oRsTmp2!FechaHora) Then
                        ldFechaMod = oRsTmp2!FechaHora
                    Else
                        If oRsTmp1.RecordCount > 0 Then
                            ldFechaMod = oRsTmp1!FechaCreacion
                        End If
                    End If
                End If
                With oDoFarmHistPrecio
                     .fecha = ldFechaMod
                     .idProducto = oRsTmp.Fields!idProducto
                     .PrecioCompra = IIf(IsNull(oRsTmp.Fields!PrecioCompra), 0, oRsTmp.Fields!PrecioCompra)
                     .PrecioDistribucion = IIf(IsNull(oRsTmp.Fields!PrecioDistribucion), 0, oRsTmp.Fields!PrecioDistribucion)
                     .PrecioDonacion = IIf(IsNull(oRsTmp.Fields!PrecioDonacion), 0, oRsTmp.Fields!PrecioDonacion)
                     .PrecioVenta = IIf(IsNull(oRsTmp.Fields!PrecioVenta), 0, oRsTmp.Fields!PrecioVenta)
                End With
                If oFarmHistPrecio.Insertar(oDoFarmHistPrecio) = True Then
                End If
                oRsTmp.MoveNext
            Loop
        End If
    End If
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_reglasComunes = Nothing
errAcHistPr:

End Sub


'debb2014b
Function PacientesAutogeneradoIGUALnull() As Recordset
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesAutogeneradoIGUALnull"
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    
    Set PacientesAutogeneradoIGUALnull = oRsTmp1
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function


'debb-20/10/2015
Function HistoriasSolicitadasSegunFiltro(lcFiltro As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HistoriasSolicitadasSegunFiltro"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set HistoriasSolicitadasSegunFiltro = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'debb-14/04/2016
Sub ActualizaCPTsinTipoTarifa(oConexion As Connection)
    On Error GoTo errorTar
    Dim mo_ReglasCaja As New ReglasCaja
    Dim mo_reglasComunes As New ReglasComunes
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oTiposTarifaCpt As New SIGHDatos.TiposTarifaCpt, oDOTiposTarifaCPT As New doTiposTarifaCpt
    Dim lnIdTipoTarifa As Long

    Set oRsTmp1 = mo_ReglasCaja.CajaComprobantesPagoParaTipoTarifaFechas(CDate("01/01/1990"), Date + 1, oConexion)
    oRsTmp1.Filter = "idtipotarifa = null"
    If oRsTmp1.RecordCount > 0 Then
       Set oRsTmp2 = mo_reglasComunes.TiposTarifaSeleccionarTodos(oConexion)
       If oRsTmp2.RecordCount > 0 Then
            oRsTmp2.MoveFirst
            lnIdTipoTarifa = oRsTmp2!IdTipoTarifa
            oRsTmp2.Find "tipotarifa= 'otros'"
            If Not oRsTmp2.EOF Then
               lnIdTipoTarifa = oRsTmp2!IdTipoTarifa
            End If
            oRsTmp2.Close
            Set oTiposTarifaCpt.Conexion = oConexion
            oDOTiposTarifaCPT.IdTipoTarifa = lnIdTipoTarifa
            oDOTiposTarifaCPT.IdUsuarioAuditoria = sighentidades.Usuario
            oRsTmp1.MoveFirst
            Do While Not oRsTmp1.EOF
               oDOTiposTarifaCPT.idProductoCPT = oRsTmp1!idProducto

               If oTiposTarifaCpt.Insertar(oDOTiposTarifaCPT) Then
               End If
               oRsTmp1.MoveNext
            Loop
       Else
            oRsTmp2.Close
       End If
    End If
errorTar:
    oRsTmp1.Close

    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set mo_ReglasCaja = Nothing
    Set oTiposTarifaCpt = Nothing
    Set oDOTiposTarifaCPT = Nothing
    Exit Sub
Resume
End Sub


'franklin 2017
Sub actualizaCamasDeObservacionConNull(oConexion As Connection)
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRsTmp As New Recordset
        With oCommand
                 .CommandType = adCmdStoredProc
                 Set .ActiveConnection = oConexion
                 .CommandTimeout = 150
                 .CommandText = "actualizaCamasDeObservacionConNull"
                 Set oRsTmp = .Execute
                 
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
        
    Exit Sub
ManejadorDeError:

End Sub

'debb-10/07/2018
Sub ActualizaIDenTablaLabResultadoPorItems(oConexion As Connection, lnIdOrden As Long)
    On Error GoTo errActIdEnTabl
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oRsTmp3 As New Recordset
    Dim lcSql As String, lbActualiza As Boolean
    Dim lbDesdeConexionVacia As Boolean
    lbDesdeConexionVacia = False
    If oConexion Is Nothing Then
       Set oConexion = New Connection
       oConexion.CommandTimeout = 900
       oConexion.CursorLocation = adUseClient
       oConexion.Open sighentidades.CadenaConexion
       lbDesdeConexionVacia = True
    End If
    
    If lnIdOrden = 0 Then
       lcSql = "select top 1000000 * FROM   dbo.LabResultadoPorItems where idGrupo is null order by fecha desc"
    Else
       lcSql = "select  * FROM   dbo.LabResultadoPorItems where idorden=" & Trim(Str(lnIdOrden)) & " and idGrupo is null order by fecha desc"
    End If
    If oRsTmp1.State = 1 Then oRsTmp1.Close
    oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
If oRsTmp1.AbsolutePosition = 6000 Then
lcSql = ""
End If
           lcSql = "SELECT     dbo.LabMovimientoCPT.ResultadoAutomatico, dbo.LabMovimientoLaboratorio.IdOrden, dbo.LabMovimientoCPT.idProductoCPT,dbo.LabMovimientoLaboratorio.idMovimiento" & _
               " FROM         dbo.LabMovimientoLaboratorio RIGHT OUTER JOIN" & _
               "  dbo.LabMovimientoCPT ON dbo.LabMovimientoLaboratorio.IdMovimiento = dbo.LabMovimientoCPT.idMovimiento" & _
               " WHERE dbo.LabMovimientoLaboratorio.IdOrden=" & oRsTmp1!IdOrden & " and dbo.LabMovimientoCPT.idProductoCPT=" & oRsTmp1!idProductoCPT
           oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
           If oRsTmp2.RecordCount > 0 Then
              lbActualiza = True
              If Not IsNull(oRsTmp2!resultadoAutomatico) Then
                 If oRsTmp2!resultadoAutomatico = 1 Then
                    lbActualiza = False
                 End If
              End If
              If lbActualiza = True Then
                 lcSql = "select * from LabItemsCpt where idProductoCpt=" & oRsTmp1!idProductoCPT & " and ordenXresultado=" & oRsTmp1!ordenXresultado
                 oRsTmp3.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                 If oRsTmp3.RecordCount > 0 Then
                    oRsTmp1!IdGrupo = oRsTmp3!idItemGrupo
                    oRsTmp1!idItem = oRsTmp3!idItem
                    oRsTmp1!ValorReferencial = oRsTmp3!ValorReferencial
                    oRsTmp1!Metodo = oRsTmp3!Metodo
                    oRsTmp1.Update
                 End If
                 oRsTmp3.Close
              End If
           Else
           lcSql = ""
           End If
           oRsTmp2.Close
           
           oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    If lbDesdeConexionVacia = True Then
       oConexion.Close
       Set oConexion = Nothing
    End If
errActIdEnTabl:
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set oRsTmp3 = Nothing
End Sub

Sub AtencionesEpisodiosDetalleEliminaNull(oConexion As Connection)
    On Error GoTo ManejadorDeError1
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRsTmp As New Recordset
        With oCommand
                 .CommandType = adCmdStoredProc
                 Set .ActiveConnection = oConexion
                 .CommandTimeout = 150
                 .CommandText = "AtencionesEpisodiosDetalleEliminaNull"
                 Set oRsTmp = .Execute
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
        
    Exit Sub
ManejadorDeError1:

End Sub


Function ListaHistoriasParaPasivo(lnIdEstadoArchivo As Integer) As Recordset
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ListaHistoriasParaPasivo"
        Set oParameter = .CreateParameter("@IdEstadoArchivo", adInteger, adParamInput, 0, lnIdEstadoArchivo): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set ListaHistoriasParaPasivo = oRsTmp1
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

Sub CreaTemporalPacientesAdepurar()
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CreaTemporalPacientesAdepurar"
         .Execute
    End With
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Sub
Function PacientesRepetidosSegunSoundex(lcApellidoPaterno As String, lcApellidoMaterno As String, lcPrimerNombre As String) As Recordset
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesRepetidosSegunSoundex"
        Set oParameter = .CreateParameter("@ApellidoPaterno   ", adVarChar, adParamInput, 40, lcApellidoPaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, lcApellidoMaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, lcPrimerNombre): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set PacientesRepetidosSegunSoundex = oRsTmp1
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function


Function AutogeneradoConNULL(oConexion As Connection)
    On Error GoTo AutogenErr
    Dim oRsPacientes As New Recordset
    Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
    Dim oDoPaciente As New DOPaciente
    Dim oPacientes As New Pacientes
    Dim lcSql As String
    lcSql = "select * from Pacientes where autogenerado is null"
    oRsPacientes.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
    If oRsPacientes.RecordCount > 0 Then
       Set oPacientes.Conexion = oConexion
       oRsPacientes.MoveFirst
       Do While Not oRsPacientes.EOF
          oDoPaciente.idPaciente = oRsPacientes!idPaciente
          oDoPaciente.IdUsuarioAuditoria = sighentidades.Usuario
          If oPacientes.SeleccionarPorId(oDoPaciente) = True Then
             oRsPacientes!autogenerado = mo_AdminAdmision.PacienteCrearNroAutogenerado(oDoPaciente)
             oRsPacientes.Update
          End If
          oRsPacientes.MoveNext
       Loop
    End If
    oRsPacientes.Close
AutogenErr:
    Set oRsPacientes = Nothing
    Set mo_AdminAdmision = Nothing
    Set oDoPaciente = Nothing
    Set oPacientes = Nothing
End Function


Function HistoriasClinicasActualizaFmovimiento(oConexion As Connection, lnIdPaciente As Long)
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HistoriasClinicasActualizaFmovimiento"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
ManejadorDeError:
End Function



Function HistoriasnormasSeleccionarTodos() As Recordset
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "historiasnormasSeleccionarTodos"
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set HistoriasnormasSeleccionarTodos = oRsTmp1
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function



Function HistoriasNormasModificar(lcNoResolucion As String, lcNoDirectiva As String, lnInstitucion As Integer, _
                                  ldFechaNorma As Date, lnAnioPasivo_N1 As Integer, lnAnioElimin_N2 As Integer, _
                                  ldFechaVigencia As Date, lnEstado As Integer) As Boolean
    On Error GoTo ErrHNM
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    HistoriasNormasModificar = False
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HistoriasNormasModificar"
        Set oParameter = .CreateParameter("@NoResolucion", adVarChar, adParamInput, 100, lcNoResolucion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NoDirectiva", adVarChar, adParamInput, 150, lcNoDirectiva): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Institucion", adInteger, adParamInput, 0, lnInstitucion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaNorma", adDBTimeStamp, adParamInput, 0, ldFechaNorma): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@AnioPasivo_N1", adInteger, adParamInput, 0, lnAnioPasivo_N1): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@AnioElimin_N2", adInteger, adParamInput, 0, lnAnioElimin_N2): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaVigencia", adDBTimeStamp, adParamInput, 0, ldFechaVigencia): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Estado", adInteger, adParamInput, 0, lnEstado): .Parameters.Append oParameter
        .Execute
    End With
    HistoriasNormasModificar = True
    oConexion.Close
ErrHNM:
    If Err.Number > 0 Then
       ms_MensajeError = Err.Description
    End If
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function


Function HistoriasClinicasSegunFiltro(lcFiltro As String) As Recordset
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HistoriasClinicasSegunFiltro"
        Set oParameter = .CreateParameter("@filtro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set HistoriasClinicasSegunFiltro = oRsTmp1
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function


Sub ActualizaGrupoSanguineoYfactorRH1(oConexion As Connection)
    Dim mo_ReglasLaboratorio As New ReglasLaboratorio
    Dim oRsTmp1 As New Recordset
    Dim lcSql As String, lcGrupoSanguineo As String, lcFactorRh As String, lnIdProducto As Long, lnIdPaciente As Long
    Dim lcProducto As String
    On Error GoTo ErrGSFHR
    
    lcSql = "SELECT     dbo.LabResultado.*, dbo.FactOrdenServicio.IdPaciente, " & _
"                      dbo.FactOrdenServicio.IdEstadoFacturacion" & _
" FROM         dbo.LabResultado LEFT OUTER JOIN" & _
"                      dbo.FactOrdenServicio ON dbo.LabResultado.idOrden = dbo.FactOrdenServicio.IdOrden" & _
" WHERE     (dbo.FactOrdenServicio.IdEstadoFacturacion <> 9) AND " & _
"      (dbo.LabResultado.idAnalisis = '86900' or dbo.LabResultado.idAnalisis = '86899')" & _
"          AND (dbo.FactOrdenServicio.IdPaciente IS NOT NULL)" & _
" Order By dbo.FactOrdenServicio.IdPaciente,dbo.LabResultado.idAnalisis,dbo.LabResultado.fecha desc"
    oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          lnIdPaciente = oRsTmp1!idPaciente
          lcGrupoSanguineo = ""
          lcFactorRh = ""
          Do While Not oRsTmp1.EOF And lnIdPaciente = oRsTmp1!idPaciente
             If Not IsNull(oRsTmp1!ResultadoAnalisis) Then
                If InStr(oRsTmp1!ResultadoAnalisis, "\") > 0 And lcGrupoSanguineo = "" Then
                   lcGrupoSanguineo = Left(oRsTmp1!ResultadoAnalisis, InStr(oRsTmp1!ResultadoAnalisis, "\") - 1)
                   lcFactorRh = Trim(Mid(oRsTmp1!ResultadoAnalisis, InStr(oRsTmp1!ResultadoAnalisis, "\") + 1, 100))
                   'Exit Do
                End If
             End If
             oRsTmp1.MoveNext
             If oRsTmp1.EOF Then
                Exit Do
             End If
          Loop
          mo_ReglasLaboratorio.ActualizaGrupoSanguineoYfactorRH lnIdPaciente, lcGrupoSanguineo, lcFactorRh, oConexion
          If oRsTmp1.EOF Then
             Exit Do
          End If
       Loop
    End If
    oRsTmp1.Close
    
    lcSql = "SELECT     dbo.FactOrdenServicio.IdPaciente, dbo.FactOrdenServicio.IdEstadoFacturacion, dbo.LabResultadoPorItems.*, " & _
"                      dbo.FactOrdenServicio.FechaCreacion" & _
" FROM         dbo.FactOrdenServicio RIGHT OUTER JOIN" & _
"                      dbo.LabResultadoPorItems ON dbo.FactOrdenServicio.IdOrden = dbo.LabResultadoPorItems.idOrden" & _
" Where (dbo.FactOrdenServicio.IdEstadoFacturacion <> 9) And (dbo.FactOrdenServicio.idPaciente Is Not Null)" & _
"       and (dbo.LabResultadoPorItems.idProductoCpt = 3841 or (dbo.LabResultadoPorItems.idProductoCpt = 50020))       " & _
" ORDER BY dbo.FactOrdenServicio.IdPaciente,dbo.LabResultadoPorItems.idproductoCpt ,dbo.FactOrdenServicio.FechaCreacion DESC"
    oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          lnIdPaciente = oRsTmp1!idPaciente
          
          lcGrupoSanguineo = ""
          lcFactorRh = ""
          Do While Not oRsTmp1.EOF And lnIdPaciente = oRsTmp1!idPaciente
                lnIdProducto = oRsTmp1!idProductoCPT
                If oRsTmp1!idProductoCPT = 3841 Then
                   If IsNull(oRsTmp1!valorTexto) Then
                      lcGrupoSanguineo = oRsTmp1!ValorCombo
                   Else
                      lcGrupoSanguineo = oRsTmp1!valorTexto
                   End If
                Else
                   If IsNull(oRsTmp1!valorTexto) Then
                      lcFactorRh = oRsTmp1!ValorCombo
                   Else
                      lcFactorRh = oRsTmp1!valorTexto
                   End If
                End If
                
                Do While Not oRsTmp1.EOF And lnIdPaciente = oRsTmp1!idPaciente And lnIdProducto = oRsTmp1!idProductoCPT
                   oRsTmp1.MoveNext
                   If oRsTmp1.EOF Then
                      Exit Do
                   End If
                Loop
                If oRsTmp1.EOF Then
                   Exit Do
                End If
          Loop
          mo_ReglasLaboratorio.ActualizaGrupoSanguineoYfactorRH lnIdPaciente, lcGrupoSanguineo, lcFactorRh, oConexion
       Loop
    End If
    oRsTmp1.Close
ErrGSFHR:
    Set mo_ReglasLaboratorio = Nothing
    Set oRsTmp1 = Nothing
    Exit Sub
    Resume
End Sub

Sub ArreglaIdPacienteEnAtenciones(oConexion As Connection)
    Dim oRsTmp1 As New Recordset
    Dim lcSql As String
    On Error GoTo ErrGSFHR11
    lcSql = "update atenciones set IdPaciente =(select IdPaciente  from FacturacionCuentasAtencion where atenciones.IdCuentaAtencion =FacturacionCuentasAtencion.IdCuentaAtencion   )" & _
"       where  atenciones.IdPaciente    <>" & _
"              (select IdPaciente  from FacturacionCuentasAtencion where" & _
"    atenciones.IdCuentaAtencion =FacturacionCuentasAtencion.IdCuentaAtencion   )"
    oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
ErrGSFHR11:
    Set oRsTmp1 = Nothing
    Exit Sub
    Resume
     
End Sub


Function farmActualizaPrecioPonderadoMenorAcero(oConexion As Connection)
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmActualizaPrecioPonderadoMenorAcero"
        .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
ManejadorDeError:
End Function
'SCCQ 27/08/2020 Cambio27 Inicio
Function SeleccionarHCSinDevolver(horas As Integer) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New MovimientosHistoriaClinica
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set SeleccionarHCSinDevolver = oTabla.SeleccionarHCSinDevolver(horas)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'SCCQ 27/08/2020 Cambio27 Fin
