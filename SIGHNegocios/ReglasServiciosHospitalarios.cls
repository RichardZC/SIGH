VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasServiciosHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para SERVICIOS HOSPITALARIOS
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim ms_MensajeError As String
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Function TiposServicioSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposServicio
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposServicioSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function TiposServicioSeleccionarAsistenciales() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposServicio
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposServicioSeleccionarAsistenciales = oTabla.SeleccionarAsistenciales
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function TiposServicioSeleccionarDeEmergencia() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposServicio
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposServicioSeleccionarDeEmergencia = oTabla.SeleccionarDeEmergencia
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EspecialidadesSeleccionarAsistenciales() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Especialidades
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EspecialidadesSeleccionarAsistenciales = oTabla.SeleccionarAsistenciales()
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function EspecialidadesSeleccionarporDepartamento(IdDepartamento As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Especialidades
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EspecialidadesSeleccionarporDepartamento = oTabla.SeleccionarPorDepartamento(IdDepartamento)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EspecialidadesSeleccionarporDepartamentoV2(IdDepartamento As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Especialidades
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EspecialidadesSeleccionarporDepartamentoV2 = oTabla.SeleccionarPorDepartamentoV2(IdDepartamento)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EspecialidadesSeleccionarporMedico(IdMedico As Long, oConexion As Connection) As ADODB.Recordset
Dim oTabla As New Especialidades
    Set oTabla.Conexion = oConexion
    Set EspecialidadesSeleccionarporMedico = oTabla.SeleccionarPorMedico(IdMedico)
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function
Function EspecialidadesSeleccionarPorId(IdEspecialidad As Long) As DOEspecialidades
Dim oConexion As New ADODB.Connection
Dim oTabla As New Especialidades
Dim doEspecialidad As New DOEspecialidades
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    doEspecialidad.IdEspecialidad = IdEspecialidad
    If oTabla.SeleccionarPorId(doEspecialidad) Then
        Set EspecialidadesSeleccionarPorId = doEspecialidad
    Else
        Set EspecialidadesSeleccionarPorId = Nothing
    End If
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EspecialidadesSeleccionarTodos(IdDepartamento As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Especialidades
Dim doEspecialidad As New DOEspecialidades
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    
    Set oTabla.Conexion = oConexion
    
    Set EspecialidadesSeleccionarTodos = oTabla.SeleccionarTodos(IdDepartamento)

    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
    Set doEspecialidad = Nothing
End Function


Function DepartamentosSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New DepartamentosHospital
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set DepartamentosSeleccionarTodos = oTabla.SeleccionarTodos()
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    
    If ms_MensajeError <> "" Then
        MsgBox ms_MensajeError, vbInformation, "Reglas de servicios hospitalarios"
    End If
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function DepartamentosSeleccionarPorId(IdDepartamento As Long) As DODepartamentoHospital
Dim oConexion As New ADODB.Connection
Dim oTabla As New DepartamentosHospital
Dim doDepartamento As New DODepartamentoHospital
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    doDepartamento.IdDepartamento = IdDepartamento
    If oTabla.SeleccionarPorId(doDepartamento) Then
        Set DepartamentosSeleccionarPorId = doDepartamento
    Else
        Set DepartamentosSeleccionarPorId = Nothing
    End If
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'-------------------------------------------------------------------------------------
'                                       Servicios
'-------------------------------------------------------------------------------------

Function ServiciosAgregar(oDOServicio As DOServicio, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNservicio As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oServicio As New Servicios
    
    ServiciosAgregar = False
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oServicio.Conexion = oConexion
    If oServicio.Insertar(oDOServicio) Then
        If oDOServicio.idTipoServicio <> 1 Then
           oServicio.Tabla_FactPuntosCarga_Actualiza
        End If
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOServicio.IdUsuarioAuditoria, "A", oDOServicio.IdServicio, "Servicios", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNservicio)            'ListBarItems.idListItem
        ServiciosAgregar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oServicio.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oServicio = Nothing
    
End Function

Function ServiciosModificar(oDOServicio As DOServicio, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNservicio As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oServicio As New Servicios
    
    ServiciosModificar = False
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oServicio.Conexion = oConexion
    
    If oServicio.modificar(oDOServicio) Then
        If oDOServicio.idTipoServicio <> 1 Then
           oServicio.Tabla_FactPuntosCarga_Actualiza
        End If
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOServicio.IdUsuarioAuditoria, "M", oDOServicio.IdServicio, "Servicios", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNservicio)            'ListBarItems.idListItem
        
        ServiciosModificar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oServicio.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oServicio = Nothing
    
End Function

Function ServiciosEliminar(oDOServicio As DOServicio, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNservicio As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oServicio As New Servicios
    ServiciosEliminar = False
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oServicio.Conexion = oConexion
    If oDOServicio.idTipoServicio <> 1 Then
       oServicio.Tabla_FactPuntosCarga_Elimina (oDOServicio.IdServicio)
    End If
    If oServicio.Eliminar(oDOServicio) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOServicio.IdUsuarioAuditoria, "E", oDOServicio.IdServicio, "Servicios", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNservicio)            'ListBarItems.idListItem
        ServiciosEliminar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oServicio.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oServicio = Nothing
    
End Function

Function ServiciosSeleccionarPorId(IdServicio As Long, oConexion As Connection) As DOServicio
Dim oTabla As New Servicios
Dim DOServicio As New DOServicio
    Set oTabla.Conexion = oConexion
    DOServicio.IdServicio = IdServicio
    If oTabla.SeleccionarPorId(DOServicio) Then
        Set ServiciosSeleccionarPorId = DOServicio
    Else
        Set ServiciosSeleccionarPorId = Nothing
    End If
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function
Function ServiciosSeleccionarIdDepartamento(IdServicio As Long) As Long
Dim oConexion As New ADODB.Connection
Dim oTabla As New Servicios
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    ServiciosSeleccionarIdDepartamento = oTabla.SeleccionarIdDepartamento(IdServicio)
    
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function ServiciosSeleccionarPorCodigo(sCodigo As String) As DOServicio
Dim oConexion As New ADODB.Connection
Dim oTabla As New Servicios
Dim DOServicio As New DOServicio
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    DOServicio.Codigo = sCodigo
    If oTabla.SeleccionarPorCodigo(DOServicio) Then
        Set ServiciosSeleccionarPorCodigo = DOServicio
    Else
        Set ServiciosSeleccionarPorCodigo = Nothing
    End If
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function ServiciosFiltrar(ByVal oServicio As DOServicio, lDepartamentoHospital As Long, lnTipoEstado As sghTipoEstados) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Servicios
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set ServiciosFiltrar = oTabla.Filtrar(oServicio, lDepartamentoHospital, lnTipoEstado)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function ServiciosSeleccionarPorTipo(idTipoServicio As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Servicios
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set ServiciosSeleccionarPorTipo = oTabla.SeleccionarPorTipo(idTipoServicio)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function ServiciosSeleccionarPorTipoV2(idTipoServicio As Long, lnTipoEstados As sghTipoEstados) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Servicios
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set ServiciosSeleccionarPorTipoV2 = oTabla.SeleccionarPorTipoV2(idTipoServicio, lnTipoEstados)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'SCCQ 19/02/2020 Cambio7 Inicio
Function ServiciosSeleccionarCEDisponibles(IdEspecialidad As Long, lnHoraInicio As String, lnHoraFin As String, lnFecha As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Servicios
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set ServiciosSeleccionarCEDisponibles = oTabla.SeleccionarCEDisponibles(IdEspecialidad, lnHoraInicio, lnHoraFin, lnFecha)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'SCCQ 19/02/2020 Cambio7 Fin

Function ServiciosSeleccionarConsultoriosPorEspecialidadDEBB(IdEspecialidad As Long, lnTipoEstados As sghTipoEstados, oConexion As Connection) As ADODB.Recordset
Dim oTabla As New Servicios
    Set oTabla.Conexion = oConexion
    Set ServiciosSeleccionarConsultoriosPorEspecialidadDEBB = oTabla.ServiciosSeleccionarConsultoriosPorEspecialidadDEBB(IdEspecialidad, lnTipoEstados)
    ms_MensajeError = oTabla.MensajeError

    Set oTabla = Nothing

End Function


Function ServiciosSeleccionarPorTipoServicioDptoEspecialidad(idTipoServicio As Long, IdDepartamento As Long, IdEspecialidad As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Servicios
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set ServiciosSeleccionarPorTipoServicioDptoEspecialidad = oTabla.SeleccionarPorTipoServicioDptoEspecialidad(idTipoServicio, IdDepartamento, IdEspecialidad)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function ServiciosSeleccionarPorTipoServicioYDpto(idTipoServicio As Long, IdDepartamento As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Servicios
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set ServiciosSeleccionarPorTipoServicioYDpto = oTabla.SeleccionarPorTipoServicioYDpto(idTipoServicio, IdDepartamento)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function ServiciosObtenerConElMismoCodigo(oDOServicio As DOServicio) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Servicios
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set ServiciosObtenerConElMismoCodigo = oTabla.ObtenerConElMismoCodigo(oDOServicio)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function ServiciosObtenerConElMismoNombre(oDOServicio As DOServicio) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Servicios
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set ServiciosObtenerConElMismoNombre = oTabla.ObtenerConElMismoNombre(oDOServicio)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function DepartamentoSeleccionarPorTipoServicio(idTipoServicio As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New DepartamentosHospital
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set DepartamentoSeleccionarPorTipoServicio = oTabla.SeleccionarPorTipoServicio(idTipoServicio)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EspecialidadSeleccionarPorTipoServicioYDpto(idTipoServicio As Long, IdDepartamento As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Especialidades
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EspecialidadSeleccionarPorTipoServicioYDpto = oTabla.SeleccionarPorTipoServicioYDpto(idTipoServicio, IdDepartamento)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EspecialidadSeleccionarPorTipoServicioSql2000(idTipoServicio As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Especialidades
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EspecialidadSeleccionarPorTipoServicioSql2000 = oTabla.SeleccionarPorTipoServicio(idTipoServicio)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'----------------------------------------- Especialidades -------------------------------
Function EspecialidadesAgregar(oDOEspecialidad As DOEspecialidades, oDOEspecialidadCE As doEspecialidadCE, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNespecialidad As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oEspecialidad As New Especialidades
Dim oEspecialidadCE As New EspecialidadCE
    
    EspecialidadesAgregar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oEspecialidad.Conexion = oConexion
    Set oEspecialidadCE.Conexion = oConexion
    
    If oEspecialidad.Insertar(oDOEspecialidad) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOEspecialidad.IdUsuarioAuditoria, "A", oDOEspecialidad.IdEspecialidad, "Especialidades", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNespecialidad)          'ListBarItems.idListItem
        oDOEspecialidadCE.IdEspecialidad = oDOEspecialidad.IdEspecialidad
        If oEspecialidadCE.Insertar(oDOEspecialidadCE) Then
            EspecialidadesAgregar = True
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oEspecialidad = Nothing
    Set oEspecialidadCE = Nothing
End Function

Function EspecialidadesModificar(oDOEspecialidad As DOEspecialidades, oDOEspecialidadCE As doEspecialidadCE, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNespecialidad As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oEspecialidad As New Especialidades
Dim oEspecialidadCE As New EspecialidadCE
Dim oRsTmp1 As New Recordset, lnTiempoPromedio As Long
Dim oDOProgramacionMedica As New DOProgramacionMedica, oProgramacionMedica As New ProgramacionMedica
Dim lcBuscaParametro As New SIGHDatos.Parametros, ldHoy As Date
Dim mo_AdminProgramacionMedica As New SIGHNegocios.ReglasDeProgMedica

Dim lnIdProductoConsulta As Long, lnIdProductoInterconsulta As Long

Dim dHoraFin As Long
    EspecialidadesModificar = False
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    Set oEspecialidad.Conexion = oConexion
    Set oEspecialidadCE.Conexion = oConexion
    lnTiempoPromedio = oDOEspecialidadCE.TiempoPromedioAtencion
    lnIdProductoConsulta = oDOEspecialidadCE.IdProductoConsulta
    lnIdProductoInterconsulta = oDOEspecialidadCE.IdProductoInterconsulta
    If oEspecialidad.modificar(oDOEspecialidad) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOEspecialidad.IdUsuarioAuditoria, "M", oDOEspecialidad.IdEspecialidad, "Especialidades", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNespecialidad)          'ListBarItems.idListItem
        oDOEspecialidadCE.IdEspecialidad = oDOEspecialidad.IdEspecialidad
        If oEspecialidadCE.SeleccionarPorId(oDOEspecialidadCE) = False Then
           If oEspecialidadCE.Insertar(oDOEspecialidadCE) Then
           End If
        End If
        oDOEspecialidadCE.TiempoPromedioAtencion = lnTiempoPromedio
        oDOEspecialidadCE.IdProductoConsulta = lnIdProductoConsulta
        oDOEspecialidadCE.IdProductoInterconsulta = lnIdProductoInterconsulta
        If oEspecialidadCE.modificar(oDOEspecialidadCE) Then
            'Actualiza Tiempo Promedio Atencion en CONSULTORIOS PROGRAMADOS pero SIN CUPO DE PACIENTE
            Set oRsTmp1 = ProgramacionMedicaSinCitasXespecialidad(oDOEspecialidad.IdEspecialidad, oConexion)
            If oRsTmp1.RecordCount > 0 Then
               ldHoy = CDate(lcBuscaParametro.RetornaFechaServidorSQL)
               oDOProgramacionMedica.IdUsuarioAuditoria = oDOEspecialidad.IdUsuarioAuditoria
               Set oProgramacionMedica.Conexion = oConexion
               oRsTmp1.MoveFirst
               Do While Not oRsTmp1.EOF
                  If oRsTmp1.Fields!fecha >= ldHoy Then
                        oDOProgramacionMedica.idProgramacion = oRsTmp1!idProgramacion
                        If oProgramacionMedica.SeleccionarPorId(oDOProgramacionMedica) = True Then
                           dHoraFin = mo_AdminProgramacionMedica.ConvertirAMinutos(oDOProgramacionMedica.HoraInicio) + lnTiempoPromedio
                           If mo_AdminProgramacionMedica.ConvertirAHora(dHoraFin) > oDOProgramacionMedica.horaFin Then
                              oDOProgramacionMedica.horaFin = mo_AdminProgramacionMedica.ConvertirAHora(dHoraFin)
                           End If
                           oDOProgramacionMedica.TiempoPromedioAtencion = lnTiempoPromedio
                           If oProgramacionMedica.modificar(oDOProgramacionMedica) = True Then
                           End If
                        End If
                  End If
                  oRsTmp1.MoveNext
               Loop
            End If
            '
            EspecialidadesModificar = True
        End If
    End If
    
    
    oConexion.Close
    Set oConexion = Nothing
    Set oEspecialidad = Nothing
    Set oEspecialidadCE = Nothing
    Set oRsTmp1 = Nothing
    Set oDOProgramacionMedica = Nothing
    Set oProgramacionMedica = Nothing
End Function

Function EspecialidadesEliminar(oDOEspecialidad As DOEspecialidades, oDOEspecialidadCE As doEspecialidadCE, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNespecialidad As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oEspecialidad As New Especialidades
Dim oEspecialidadCE As New EspecialidadCE
    
    EspecialidadesEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    
    Set oEspecialidad.Conexion = oConexion
    Set oEspecialidadCE.Conexion = oConexion
    
    If oEspecialidadCE.Eliminar(oDOEspecialidadCE) Then
        If oEspecialidad.Eliminar(oDOEspecialidad) Then
            '
            Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOEspecialidad.IdUsuarioAuditoria, "E", oDOEspecialidad.IdEspecialidad, "Especialidades", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNespecialidad)          'ListBarItems.idListItem
            EspecialidadesEliminar = True
        End If
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oEspecialidad = Nothing
    Set oEspecialidadCE = Nothing
    
End Function

Function EspecialidadesFiltrar(ByVal oDoEspecialidades As DOEspecialidades) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Especialidades
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EspecialidadesFiltrar = oTabla.Filtrar(oDoEspecialidades)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'----------------------------------------- EspecialidadesCE -------------------------------
Function EspecialidadesCESeleccionarPorId(IdEspecialidadCE As Long) As doEspecialidadCE
Dim oConexion As New ADODB.Connection
Dim oTabla As New EspecialidadCE
Dim doEspecialidadCE As New doEspecialidadCE
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    doEspecialidadCE.IdEspecialidadCE = IdEspecialidadCE
    If oTabla.SeleccionarPorId(doEspecialidadCE) Then
        Set EspecialidadesCESeleccionarPorId = doEspecialidadCE
    Else
        Set EspecialidadesCESeleccionarPorId = Nothing
    End If
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EspecialidadesCESeleccionarPorIdEspecialidad(IdEspecialidad As Long) As doEspecialidadCE
Dim oConexion As New ADODB.Connection
Dim oTabla As New EspecialidadCE
Dim doEspecialidadCE As New doEspecialidadCE
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    doEspecialidadCE.IdEspecialidad = IdEspecialidad
    If oTabla.SeleccionarPorIdEspecialidad(doEspecialidadCE) Then
        Set EspecialidadesCESeleccionarPorIdEspecialidad = doEspecialidadCE
    Else
        Set EspecialidadesCESeleccionarPorIdEspecialidad = Nothing
    End If
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'09/08/2011
Function ServiciosSeleccionarPorTipoV2debb(idTipoServicio As Long, lcEspecialidadUsuario As String, _
                                          lnTipoEstado As sghTipoEstados) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
Dim lcSql As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
   'Creamos la cadena para el filtro
   lcSql = lcEspecialidadUsuario
   If lnTipoEstado <> sghFiltraAnuladosYactivos Then
      lcSql = lcSql & " and idEstado= " & lnTipoEstado
   End If
   If idTipoServicio = sghTipoServicio.sghEmergenciaConsultorios Then
      lcSql = lcSql & " and EsObservacionEmergencia=1"
   End If
   lcSql = Mid(lcSql, 5)
   
   If Len(lcSql) > 1000 Then
        MsgBox "Observación: El usuario labora en varias Áreas ", vbExclamation, "Reglas Servicios Hospitalarios"
        Exit Function
   End If
   
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosSeleccionarPorTipoV2debb"
        Set oParameter = .CreateParameter("@idTipoServicio", adInteger, adParamInput, 0, idTipoServicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 1000, lcSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   
'   oRecordset.Filter = lcSql
   Set ServiciosSeleccionarPorTipoV2debb = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   '
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function EspecialidadCESeleccionarPorIdServicio(lnIdServicio As Long) As ADODB.Recordset
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
        .CommandText = "EspecialidadCESeleccionarPorIdServicio"
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set EspecialidadCESeleccionarPorIdServicio = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   '
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ServiciosDevuelveNombre(lnIdServicio As Long) As String
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
       .CommandText = "ServiciosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
       Set oRecordset = .Execute
    End With
    If oRecordset.RecordCount > 0 Then
       ServiciosDevuelveNombre = oRecordset.Fields!nombre
    Else
       ServiciosDevuelveNombre = ""
    End If
    oRecordset.Close
    oConexion.Close
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oConexion = Nothing
    Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function EspecialidadCEseleccionarIdServicio(lnIdServicio As Long) As String


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
        .CommandText = "EspecialidadCEseleccionarIdServicio"
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   '
   EspecialidadCEseleccionarIdServicio = "0"
   If oRecordset.RecordCount > 0 Then
       EspecialidadCEseleccionarIdServicio = oRecordset.Fields!TiempoPromedioAtencion
   End If
   oRecordset.Close
   Set oRecordset = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
    
End Function


Function ProgramacionMedicaSinCitasXespecialidad(lnIdEspecialidad As Long, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProgramacionMedicaSinCitasXespecialidad"
        Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, lnIdEspecialidad): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ProgramacionMedicaSinCitasXespecialidad = oRecordset
   Set oRecordset = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
   '
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function EspecialidadesSeleccionarXcodigoHIS(lcCodigoServicioHIS As String, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EspecialidadesSeleccionarXcodigoHIS"
        Set oParameter = .CreateParameter("@CodigoServicioHIS", adVarChar, adParamInput, 6, lcCodigoServicioHIS): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set EspecialidadesSeleccionarXcodigoHIS = oRecordset
   Set oRecordset = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
   '
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'debb-19/04/2016
Function EspecialidadCESeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 900
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EspecialidadCESeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set EspecialidadCESeleccionarTodos = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   '
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'debb-19/04/2016
Function EspecialidadesSeleccionarXid(lnIdEspecialidad As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EspecialidadesSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, lnIdEspecialidad): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set EspecialidadesSeleccionarXid = oRecordset
   Set oRecordset = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
   '
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



Function ServiciosPuedeUsarFichaNacimientos(lIdServicio As Long, oConexion As Connection) As Boolean
    Dim oServicio As Servicios
    Dim oDOServicio As New DOServicio
    'Dim oConexion As New Connection
    'oConexion.CommandTimeout = 300
    'oConexion.CursorLocation = adUseClient
    'oConexion.Open sighentidades.CadenaConexion
    ServiciosPuedeUsarFichaNacimientos = False
    oDOServicio.IdServicio = lIdServicio
    Set oServicio = New Servicios
    Set oServicio.Conexion = oConexion
    If oServicio.SeleccionarPorId(oDOServicio) = True Then
        If oDOServicio.usaNacimiento = 1 Then
           ServiciosPuedeUsarFichaNacimientos = True
        End If
    End If
    'oConexion.Close
    'Set oConexion = Nothing
End Function
