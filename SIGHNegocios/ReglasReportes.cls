VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para REPORTES
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError  As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property

Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

'---------------------------------------------------------------------------------------------
'                           REPORTES DE FILIACION DE PACIENTE
'---------------------------------------------------------------------------------------------
Function ReporteAtencionesParaHistoriaClinica(idAtencion As Long) As ADODB.Recordset ' scrafet reporte 1
Dim oAtencion As New Atenciones
Dim oConexion As New ADODB.Connection

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
    Set oAtencion.Conexion = oConexion
    Set ReporteAtencionesParaHistoriaClinica = oAtencion.ReporteParaHistoriaClinica(idAtencion)
    oConexion.Close
    ms_MensajeError = oAtencion.MensajeError
    Set oConexion = Nothing
    Set oAtencion = Nothing
End Function
'<(Inicio) Añadido Por: WABG el: 13/05/2021-09:45:24 a.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-74>
'---------------------------------------------------------------------------------------------
'                           REPORTES DE FILIACION DE PACIENTE (REFERENCIA)
'---------------------------------------------------------------------------------------------
Function ObtenerDatosAdicionalesDeAtencionPorIdAtencion(idAtencion As Long) As ADODB.Recordset
Dim oAtencion As New Atenciones
Dim oConexion As New ADODB.Connection

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
    Set oAtencion.Conexion = oConexion
    Set ObtenerDatosAdicionalesDeAtencionPorIdAtencion = oAtencion.ObtenerDatosAdicionalesDeAtencionPorIdAtencion(idAtencion)
    oConexion.Close
    ms_MensajeError = oAtencion.MensajeError
    Set oConexion = Nothing
    Set oAtencion = Nothing
End Function

'---------------------------------------------------------------------------------------------
'                           REPORTES DE FILIACION DE PACIENTE (DIAGNOSTICO)
'---------------------------------------------------------------------------------------------
'</(Fin) Añadido Por: WABG el: 13/05/2021-09:45:24 a.m. en el Equipo: SISGALENPLUS-PC<CAMBIO-74>


'---------------------------------------------------------------------------------------------
'                           REPORTES DE PROGRAMACION MEDICA
'---------------------------------------------------------------------------------------------
'MODIFICADO POR FRANKLIN CACHAY 30/09/2013 - se cambio a store procedure
Function ReporteProgramacionMedicaPorMedico(lIdMedico As Long, daFechaInicio As Date, daFechaFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
'Dim ssql As String
Dim oConexion As New ADODB.Connection

  Set ReporteProgramacionMedicaPorMedico = Nothing
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
     With oCommand
'        .CommandType = adCmdText
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "ReporteProgramacionMedicaPorMedico"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, Format(daFechaInicio, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(daFechaFin, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lIdMedico): .Parameters.Append oParameter
'        .CommandText = ssql
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    oConexion.Close
    Set ReporteProgramacionMedicaPorMedico = oRecordset
    ms_MensajeError = ""
  'Set oRecordSet = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 30/09/2013 - se cambio a store procedure
Function FiltrarMedicosPorDepartamento(lIdDepartamento As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
'Dim ssql As String
Dim oConexion As New ADODB.Connection

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
        
        Set FiltrarMedicosPorDepartamento = Nothing

        With oCommand
'        .CommandType = adCmdText
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FiltrarMedicosPorDepartamento"
        Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, lIdDepartamento): .Parameters.Append oParameter
'        .CommandText = ssql
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    
    Set FiltrarMedicosPorDepartamento = oRecordset
    oConexion.Close
  Set oConexion = Nothing
   Set oCommand = Nothing
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'---------------------------------------------------------------------------------------------
'                           REPORTES DE ARCHIVO CLINICO
'---------------------------------------------------------------------------------------------

'MODIFICADO POR FRANKLIN CACHAY 01/09/2013 - se cambio a store procedure
Function FiltrarHistoriaSolicitadas(lIdResponsable As Long, daFechaRequeridaDesde As Date, daFechaRequeridaHasta As Date, _
                                   daFechaSolicitudDesde As Date, daFechaSolicitudHasta As Date, lIdTipoServicio As Long, _
                                   ml_IncluyeHistoriasQueSalieron As Boolean) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim sSql As String
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    sSql = IIf(ml_IncluyeHistoriasQueSalieron = True, "", " and IdMovimiento is null ")
    If daFechaRequeridaDesde <> 0 Then
        If daFechaRequeridaHasta = 0 Then
            sSql = sSql + " and FechaRequerida = '" & Format(daFechaRequeridaDesde, sighentidades.DevuelveFechaSoloFormato_DMY) & "'"
        Else
            sSql = sSql + " and FechaRequerida between CONVERT(DATETIME,'" & Format(daFechaRequeridaDesde, sighentidades.DevuelveFechaSoloFormato_DMY_HM) & "', 103)"
            sSql = sSql + " and CONVERT(DATETIME,'" & Format(daFechaRequeridaHasta, sighentidades.DevuelveFechaSoloFormato_DMY_HM) & "', 103)"
        End If
    End If
    If daFechaSolicitudDesde <> 0 Then
        If daFechaSolicitudHasta = 0 Then
            sSql = sSql + " and FechaSolicitud = '" & Format(daFechaSolicitudDesde, sighentidades.DevuelveFechaSoloFormato_DMY) & "'"
        Else
            sSql = sSql + " and FechaSolicitud between CONVERT(DATETIME,'" & Format(daFechaSolicitudDesde, sighentidades.DevuelveFechaSoloFormato_DMY) & "', 103)"
            sSql = sSql + " and CONVERT(DATETIME,'" & Format(daFechaSolicitudHasta, sighentidades.DevuelveFechaSoloFormato_DMY) & "', 103)"
        End If
    End If
    sSql = sSql + " Order By HistoriasSolicitadas.IdServicio,HistoriasSolicitadas.FechaRequerida, HistoriasSolicitadas.HoraRequerida,Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HistoriasSolicitadasFiltrar"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 3000, sSql): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdResponsable", adVarChar, adParamInput, 10, lIdResponsable): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdTipoServicio", adVarChar, adParamInput, 10, lIdTipoServicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcBuscaParametro231", adVarChar, adParamInput, 10, lcBuscaParametro.SeleccionaFilaParametro(231)): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set FiltrarHistoriaSolicitadas = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
    
End Function

'MODIFICADO POR FRANKLIN CACHAY 01/09/2013 - Se cambio a store procedure
Function CantidadHistoriaSolicitadasPorServicio(daFechaRequeridaDesde As Date, daFechaRequeridaHasta As Date, daFechaSolicitudDesde As Date, daFechaSolicitudHasta As Date, lIdTipoServicio As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
'Dim ssql As String
Dim sCommandText As String
Dim oConexion As New ADODB.Connection
oConexion.CursorLocation = adUseClient
oConexion.CommandTimeout = 300
oConexion.Open sighentidades.CadenaConexion
        
        Set CantidadHistoriaSolicitadasPorServicio = Nothing

     sCommandText = "RepCantidadHistoriaSolicitadasPorServicio"
     With oCommand
        .CommandType = adCmdStoredProc
        .CommandTimeout = 150
        Set .ActiveConnection = oConexion
        
        If daFechaRequeridaDesde <> 0 Then
            If daFechaRequeridaHasta = 0 Then
                sCommandText = "RepCantidadHistoriaSolicitadasPorServicioFechaRequerida"
                Set oParameter = .CreateParameter("@FechaRequeridaDesde", adDBTimeStamp, adParamInput, 0, daFechaRequeridaDesde): .Parameters.Append oParameter
            Else
                sCommandText = "RepCantidadHistoriaSolicitadasPorServicioFechaReqDesdeHasta"
                Set oParameter = .CreateParameter("@FechaRequeridaDesde", adDBTimeStamp, adParamInput, 0, daFechaRequeridaDesde): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@FechaRequeridaHasta", adDBTimeStamp, adParamInput, 0, daFechaRequeridaHasta): .Parameters.Append oParameter
            End If
        End If
        
        If daFechaSolicitudDesde <> 0 Then
            If daFechaSolicitudHasta = 0 Then
                If sCommandText = "RepCantidadHistoriaSolicitadasPorServicio" Then
                    sCommandText = "RepCantidadHistoriaSolicitadasPorServicioFechaSolicitud"
                Else
                    sCommandText = sCommandText & "FechaSolicitud" 'Nombre del procedimiento
                End If
                Set oParameter = .CreateParameter("@FechaSolicitudDesde", adDBTimeStamp, adParamInput, 0, daFechaSolicitudDesde): .Parameters.Append oParameter
            Else
                If sCommandText = "RepCantidadHistoriaSolicitadasPorServicio" Then
                    sCommandText = "RepCantidadHistoriaSolicitadasPorServicioFechaSolDesdeHasta"
                Else
                    sCommandText = sCommandText & "FechaSolDesdeHasta" 'Nombre del procedimiento
                End If
                Set oParameter = .CreateParameter("@FechaSolicitudDesde", adDBTimeStamp, adParamInput, 0, daFechaSolicitudDesde): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@FechaSolicitudHasta", adDBTimeStamp, adParamInput, 0, daFechaSolicitudHasta): .Parameters.Append oParameter
            End If
        End If
        .CommandText = sCommandText 'El nombre del procedimiento almacenado llamado es dinamico FCV 02/10/2013
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
     Set CantidadHistoriaSolicitadasPorServicio = oRecordset
     oConexion.Close
    
   ms_MensajeError = ""
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
    
End Function

'MODIFICADO POR FRANKLIN CACHAY 30/09/2013 - se cambio a store procedure
Function ObtenerCabeceraDeMovimiento(lIdGrupoMovimiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
'Dim ssql As String
Dim oConexion As New ADODB.Connection

oConexion.CursorLocation = adUseClient
oConexion.CommandTimeout = 300
oConexion.Open sighentidades.CadenaConexion
        

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ObtenerCabeceraDeMovimiento"
            Set oParameter = .CreateParameter("@IdGrupoMovimiento", adInteger, adParamInput, 0, lIdGrupoMovimiento): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
        End With
    
    Set ObtenerCabeceraDeMovimiento = oRecordset
    oConexion.Close
    
   ms_MensajeError = ""
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
    
End Function

'MODIFICADO POR FRANKLIN CACHAY 30/09/2013 - se cambio a store procedure
Function ObtenerHistoriasClinicasDeEsteMovimiento(lIdGrupoMovimiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
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
            .CommandText = "ObtenerHistoriasClinicasDeEsteMovimiento"
            Set oParameter = .CreateParameter("@IdGrupoMovimiento", adInteger, adParamInput, 0, lIdGrupoMovimiento): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
        End With
    
    Set ObtenerHistoriasClinicasDeEsteMovimiento = oRecordset
    oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
    
End Function

'***************daniel barrantes**************
'***************filtro de datos para el Reporte de Egresos detallado
'***************Hospitalizacion o Emergencia
'MODIFICADO POR FRANKLIN CACHAY 02/09/2013 - se cambio a store procedure
Function ReporteEgresosHospitalarios(lIdDepartamentoEgreso As Long, lIdEspecialidadEgreso As Long, _
                                     lIdServicioEgreso As Long, FechaEgrIni As Date, FechaEgrFin As Date, _
                                     lnTipoEspecialidad As Integer) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
'Dim ssql As String
'Dim sWhere As String
Dim oConexion As New ADODB.Connection

        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient

     Set ReporteEgresosHospitalarios = Nothing
     
       
        With oCommand
            '.CommandType = adCmdText
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            '.CommandText = ssql
            .CommandText = "ReporteEgresosHospitalarios"
            Set oParameter = .CreateParameter("@IdDepartamentoEgreso", adInteger, adParamInput, 0, lIdDepartamentoEgreso): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdServicioEgreso", adInteger, adParamInput, 0, lIdServicioEgreso): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdEspecialidadEgreso", adInteger, adParamInput, 0, lIdEspecialidadEgreso): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@TipoEspecialidad", adInteger, adParamInput, 0, lnTipoEspecialidad): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, FechaEgrIni): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, FechaEgrFin): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
        End With
    Set ReporteEgresosHospitalarios = oRecordset
    ms_MensajeError = ""
    
    'oRecordSet.Close
    oConexion.Close
    'Set oRecordSet = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing

Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


'***************daniel barrantes**************
'***************filtro de datos para el Reporte de Ingresos detallado
'***************Hospitalizacion o Emergencia
'MODIFICADO POR FRANKLIN CACHAY 02/10/2013 - se cambio a store procedure
Function ReporteIngresosHospitalarios(lIdDepartamentoIngreso As Long, lIdEspecialidadIngreso As Long, lIdServicioIngreso As Long, FechaIngIni As Date, FechaIngFin As Date, lnTipoEspecialidad As Integer) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
'Dim ssql As String
'Dim sWhere As String
Dim oConexion As New ADODB.Connection

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion

     Set ReporteIngresosHospitalarios = Nothing
     

       With oCommand
            '.CommandType = adCmdText
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            '.CommandText = ssql
            .CommandText = "ReporteIngresosHospitalarios"
            Set oParameter = .CreateParameter("@IdDepartamentoIngreso", adInteger, adParamInput, 0, lIdDepartamentoIngreso): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdServicioIngreso", adInteger, adParamInput, 0, lIdServicioIngreso): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdEspecialidadIngreso", adInteger, adParamInput, 0, lIdEspecialidadIngreso): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@TipoEspecialidad", adInteger, adParamInput, 0, lnTipoEspecialidad): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, FechaIngIni): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, FechaIngFin): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
        End With
    Set ReporteIngresosHospitalarios = oRecordset
    
    ms_MensajeError = ""
  
    'oRecordSet.Close
    oConexion.Close
    'Set oRecordSet = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function ReporteAtencionesDiagnosticosDeEgreso(idAtencion As Long) As ADODB.Recordset 'scrafet reporte 4
Dim oAtencion As New AtencionesDiagnosticos
Dim oConexion As New ADODB.Connection
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oAtencion.Conexion = oConexion
    Set ReporteAtencionesDiagnosticosDeEgreso = oAtencion.SeleccionarDiagnosticosDeEgreso(idAtencion)
    oConexion.Close
    
    ms_MensajeError = oAtencion.MensajeError
    
    
End Function

'---------------------------------------------------------------------------------------------
'                           REPORTES DE ARCHIVO CLINICO
'---------------------------------------------------------------------------------------------
'MODIFICADO POR FRANKLIN CACHAY 02/09/2013 - se cambio a store procedure
Function ReporteHistoriasSolicitadasCEPorMedico(lIdResponsable As Long, daFechaRequeridaDesde As Date, _
                            daFechaRequeridaHasta As Date, daFechaSolicitudDesde As Date, daFechaSolicitudHasta As Date, _
                            lIdTipoServicio As Long, ml_IncluyeHistoriasQueSalieron As Boolean) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
Dim sSql As String
   oConexion.CommandTimeout = 300
   oConexion.CursorLocation = adUseClient
   oConexion.Open sighentidades.CadenaConexion
   sSql = IIf(ml_IncluyeHistoriasQueSalieron = True, "", " and IdMovimiento is null ")
   If daFechaRequeridaDesde <> 0 Then
        If daFechaRequeridaHasta = 0 Then
            sSql = sSql + " and HistoriasSolicitadas.FechaRequerida = '" & Format(daFechaRequeridaDesde, sighentidades.DevuelveFechaSoloFormato_DMY) & "'"
        Else
            sSql = sSql + " and HistoriasSolicitadas.FechaRequerida between CONVERT(DATETIME,'" & Format(daFechaRequeridaDesde, sighentidades.DevuelveFechaSoloFormato_DMY) & "', 103)"
            sSql = sSql + " and CONVERT(DATETIME,'" & Format(daFechaRequeridaHasta, sighentidades.DevuelveFechaSoloFormato_DMY) & "', 103)"
        End If
   End If
   If daFechaSolicitudDesde <> 0 Then
        If daFechaSolicitudHasta = 0 Then
            sSql = sSql + " and HistoriasSolicitadas.FechaSolicitud = '" & Format(daFechaSolicitudDesde, sighentidades.DevuelveFechaSoloFormato_DMY) & "'"
        Else
            sSql = sSql + " and (HistoriasSolicitadas.FechaSolicitud between CONVERT(DATETIME,'" & Format(daFechaSolicitudDesde, sighentidades.DevuelveFechaSoloFormato_DMY) & "', 103)"
            sSql = sSql + " and CONVERT(DATETIME,'" & Format(daFechaSolicitudHasta, sighentidades.DevuelveFechaSoloFormato_DMY) & "', 103))"
        End If
   End If
   sSql = sSql + " Order By Empleados.ApellidoPaterno, Empleados.ApellidoMaterno, Empleados.Nombres, Citas.Fecha, Citas.HoraInicio"
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = oConexion
     .CommandText = "HistoriasSolicitadasFiltrarXmedico"
     Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 3000, sSql): .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@lIdResponsable", adVarChar, adParamInput, 10, lIdResponsable): .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@lIdTipoServicio", adVarChar, adParamInput, 10, lIdTipoServicio): .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@lcBuscaParametro231", adVarChar, adParamInput, 1, lcBuscaParametro.SeleccionaFilaParametro(231)): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
   Set oRecordset.ActiveConnection = Nothing
   Set ReporteHistoriasSolicitadasCEPorMedico = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
     
MsgBox Err.Description
End Function

'---------------------------------------------------------------------------------------------
'                           REPORTES DE ARCHIVO CLINICO
'---------------------------------------------------------------------------------------------
Function ReporteAtencionesProcedimientos(lIdAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection

    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
 
   Set ReporteAtencionesProcedimientos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = oConexion
     .CommandText = "FacturacionServicioProcedimientos"
       Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
   
   Set ReporteAtencionesProcedimientos = oRecordset
   Set oRecordset.ActiveConnection = Nothing
   
   oConexion.Close
  Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
End Function


'---------------------------------------------------------------------------------------------
'                           REPORTES DE CAJA
'---------------------------------------------------------------------------------------------
'MODIFICADO POR FRANKLIN CACHAY 30/09/2013 - se cambio a store procedure
Function ReportDiarioDeCaja(lIdGestionCaja As Long, sEstados As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection

oConexion.CursorLocation = adUseClient
oConexion.CommandTimeout = 300
oConexion.Open sighentidades.CadenaConexion
        

    With oCommand
        '.CommandType = adCmdText
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ReportDiarioDeCaja"
        Set oParameter = .CreateParameter("@Estados", adVarChar, adParamInput, 50, sEstados): .Parameters.Append oParameter

        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
  
   Set ReportDiarioDeCaja = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
    
End Function

'MODIFICADO POR FRANKLIN CACHAY 30/09/2013 - se cambio a store procedure
Function DatosGestionCaja(lIdGestionCaja As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
'Dim ssql As String
Dim oConexion As New ADODB.Connection

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion

     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DatosGestionCaja"
        Set oParameter = .CreateParameter("@IdGestionCaja", adInteger, adParamInput, 0, lIdGestionCaja): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    
    Set DatosGestionCaja = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 30/09/2013 - se cambio a store procedure
Function ReporteConsolidadodeServiciosSinExoneraciones(lIdGestionCaja As Long, sEstados As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
'Dim ssql As String
Dim oConexion As New ADODB.Connection

        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        
                            

      With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ReporteConsolidadodeServiciosSinExoneraciones"
        Set oParameter = .CreateParameter("@Estados", adVarChar, adParamInput, 50, sEstados): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    
    Set ReporteConsolidadodeServiciosSinExoneraciones = oRecordset
    oConexion.Close
  Set oConexion = Nothing
   Set oCommand = Nothing
   
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
    
End Function

'MODIFICADO POR FRANKLIN CACHAY 30/09/2013 - se cambio a store procedure
Function ReporteConsolidadodeServiciosExoneraciones(lIdGestionCaja As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
'Dim ssql As String
Dim oConexion As New ADODB.Connection

        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
                                                                    
                                                                    

     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ReporteConsolidadodeServiciosExoneraciones"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    
    Set ReporteConsolidadodeServiciosExoneraciones = oRecordset
    oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
    
End Function

Sub ExportarRecordSetAexcel(rsReporte As Recordset, lcTitulo As String, lcSubTitulo As String, _
                            lcPiePagina As String, lnHwnd As Long, Optional lbConfiguraAnchoDeColumna As Boolean, _
                            Optional lbImpresionHorizontal As Boolean, Optional oRsPie As Recordset)
'On Error Resume Next
                            
Dim mo_ReporteUtil As New ReporteUtil
Dim iFila As Long
Dim lnCampo As Long
Dim lnNroCamposRs As Integer, lnNroCamposRs1 As Long
Dim lbEsOpenOffice As Boolean
Dim lcSql As String
lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)

    If lbEsOpenOffice = True Then
        Dim ServiceManager As Object
        Dim Desktop As Object
        Dim Document As Object
        Dim Feuille As Object
        Dim Plage As Object
        Dim args()
        Dim Chemin As String
        Dim Fichier As String
        Dim lcArchivoExcel As String
        Dim PrintArea(0)
        Dim Style As Object
        Dim Border As Object
        'encabezado
        Dim PageStyles As Object
        Dim Sheet As Object
        Dim StyleFamilies As Object
        Dim DefPage As Object
        Dim Htext As Object
        Dim Hcontent As Object
        Dim ret As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If

        If rsReporte.RecordCount > 0 Then
       ' On Error Resume Next
        lnNroCamposRs = rsReporte.Fields.Count
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\HojaLibre.ods"
'            FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'            Chemin = "file:///" & App.Path & "\Plantillas\"
'            Chemin = Replace(Chemin, "\", "/")
'            Fichier = Chemin & "/OpenOffice.ods"
            '
            Fichier = Format(Time, "hhmmss") & ".ods"
            FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
            lcArchivoExcel = Fichier
            Chemin = "file:///" & App.Path & "\Plantillas\"
            Chemin = Replace(Chemin, "\", "/")
            Fichier = Chemin & "/" & lcArchivoExcel
            '
            Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
            Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
            Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
            Set Feuille = Document.getSheets().getByIndex(0)
            'Encabezado de Pagina
            mo_CabeceraReportes.CabeceraReportes Document, True
            ' Pone la ventana en primer plano, pasándole el Hwnd
            ret = SetForegroundWindow(lnHwnd)
        Else
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            'Crea nueva hoja
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HojaLibre.xls")
            oWorkBookPlantilla.Worksheets("Hoja_libre").COPY Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(0, 0).setFormula(lcTitulo)
            Call Feuille.getcellbyposition(0, 1).setFormula(lcSubTitulo)
        Else
            'Titulos
            oWorkSheet.Cells(1, 1).value = lcTitulo
            oWorkSheet.Cells(2, 1).value = lcSubTitulo
        End If
        'Titulo de Campos
        iFila = 5
        For lnCampo = 0 To lnNroCamposRs - 1
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(lnCampo + 0, iFila - 1).setFormula(rsReporte.Fields(lnCampo).Name)
                If lbConfiguraAnchoDeColumna = True Then
                    Call Feuille.getcellbyposition(lnCampo + 0, iFila - 1).Width(rsReporte.Fields(lnCampo).DefinedSize)
                End If
            Else
                oWorkSheet.Cells(iFila, lnCampo + 1).value = rsReporte.Fields(lnCampo).Name
                If lbConfiguraAnchoDeColumna = True Then
                    oWorkSheet.Cells(iFila, lnCampo + 1).ColumnWidth = rsReporte.Fields(lnCampo).DefinedSize
                End If
            End If
        Next
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName("A" & CStr(iFila) & ":" & BuscaNombreColumna(lnNroCamposRs) & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 1, iFila, lnNroCamposRs
        End If
        iFila = iFila + 1
        'Filtra los Datos
        
        rsReporte.MoveFirst
        Do While Not rsReporte.EOF
           For lnCampo = 0 To lnNroCamposRs - 1
                lcSql = ""
                If Not IsNull(rsReporte.Fields(lnCampo).value) Then
                   Select Case rsReporte.Fields(lnCampo).Type
                   Case adDate, adDBDate
                        lcSql = Format(rsReporte.Fields(lnCampo).value, "General Date")
                   Case adInteger, adNumeric
                        lcSql = Format(rsReporte.Fields(lnCampo).value, "#########")
                   Case adDecimal, adDouble
                        lcSql = Format(rsReporte.Fields(lnCampo).value, "#########.##")
                   Case adBoolean
                        lcSql = IIf(rsReporte.Fields(lnCampo).value = True, "Verdadero", "Falso")
                   Case Else
                        lcSql = rsReporte.Fields(lnCampo).value
                        If InStr("0123456789", Left(lcSql, 1)) > 0 Then
                           lcSql = "'" & rsReporte.Fields(lnCampo).value
                        End If
                   End Select
                   If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(lnCampo + 0, iFila - 1).setFormula(lcSql)
                   Else
                        oWorkSheet.Cells(iFila, lnCampo + 1).value = lcSql
'                        If Err.Number = 1004 Then
'                           iFila = iFila - 10
'                           Exit Do
'                        End If
                   End If
                End If
           Next
           iFila = iFila + 1
           rsReporte.MoveNext
        Loop
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName("A" & CStr(iFila) & ":" & BuscaNombreColumna(lnNroCamposRs) & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 1, iFila, lnNroCamposRs
        End If
        iFila = iFila + 1
        'debb-09/08/2016
        If oRsPie Is Nothing Then
        Else
        If oRsPie.RecordCount > 0 Then
            iFila = iFila - 1
            oRsPie.MoveFirst
            Do While Not oRsPie.EOF
               For lnCampo = 0 To lnNroCamposRs - 1
                    lcSql = ""
                    If Not IsNull(oRsPie.Fields(lnCampo).value) Then
                       Select Case rsReporte.Fields(lnCampo).Type
                       Case adDate, adDBDate
                            lcSql = Format(oRsPie.Fields(lnCampo).value, "General Date")
                       Case adInteger, adNumeric
                            lcSql = Format(oRsPie.Fields(lnCampo).value, "#########")
                       Case adDecimal, adDouble
                            lcSql = Format(oRsPie.Fields(lnCampo).value, "#########.##")
                       Case adBoolean
                            lcSql = IIf(oRsPie.Fields(lnCampo).value = True, "Verdadero", "Falso")
                       Case Else
                            lcSql = oRsPie.Fields(lnCampo).value
                            If InStr("0123456789", Left(lcSql, 1)) > 0 Then
                               lcSql = "'" & oRsPie.Fields(lnCampo).value
                            End If
                       End Select
                       If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(lnCampo + 0, iFila - 1).setFormula(lcSql)
                       Else
                            oWorkSheet.Cells(iFila, lnCampo + 1).value = lcSql
                       End If
                    End If
               Next
               iFila = iFila + 1
               oRsPie.MoveNext
            Loop
           
        End If
        End If
        If lbEsOpenOffice = True Then
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 1
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = lnNroCamposRs
            PrintArea(0).EndRow = iFila
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oWorkSheet.Cells(iFila, 1).value = lcPiePagina
            If oWorkSheet.PageSetup.PrintArea <> "" Then
               lnNroCamposRs1 = lnNroCamposRs
               oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, lnNroCamposRs1)
            End If
            If lbImpresionHorizontal = True Then
                oWorkSheet.PageSetup.Orientation = xlLandscape
            End If
            oExcel.Visible = True
            oWorkSheet.PrintPreview
        End If
        If lbEsOpenOffice = True Then
            'Liberar Memoria
            Set Plage = Nothing
            Set Feuille = Nothing
            Set Document = Nothing
            Set Desktop = Nothing
            Set ServiceManager = Nothing
            Set Style = Nothing
            Set Border = Nothing

            'encabezado de pagina
            Set PageStyles = Nothing
            Set Sheet = Nothing
            Set StyleFamilies = Nothing
            Set DefPage = Nothing
            Set Htext = Nothing
            Set Hcontent = Nothing
        Else
            'Liberar memoria
            Set oExcel = Nothing
            Set oWorkBookPlantilla = Nothing
            Set oWorkBook = Nothing
            Set oWorkSheet = Nothing
        End If
    End If
End Sub

Sub ExportarRecordSetAexcelFast(rs As Recordset, textoCabecera As String, lcSubTitulo As String, _
                            lcPiePagina As String, lnHwnd As Long, Optional lbConfiguraAnchoDeColumna As Boolean, _
                            Optional lbImpresionHorizontal As Boolean, Optional oRsPie As Recordset)

    Dim oExcel As Excel.Application
    Dim oLibro As Excel.Workbook
    Dim oHoja As Excel.Worksheet
    Dim iFila As Long, iCol As Integer, I As Integer, lnCampo As Integer, lcSql As String
    Dim cRangeCabecera As String
    Dim cRangeCuerpo As String
    
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim NOMBREEESS As String
    Dim DIRECCIONEESS As String
    Dim TELEFONO As String
        
    NOMBREEESS = lcBuscaParametro.SeleccionaFilaParametro(205)
    DIRECCIONEESS = lcBuscaParametro.SeleccionaFilaParametro(206)
    TELEFONO = lcBuscaParametro.SeleccionaFilaParametro(207)
    
    
    
    Set oExcel = New Excel.Application
    Set oLibro = oExcel.Workbooks.Add
    Set oHoja = oLibro.Worksheets(1)
    Screen.MousePointer = vbHourglass
    iFila = 3
    iCol = 1
    
    
    oHoja.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
    oHoja.PageSetup.CenterHeader = "&""Arial Narrow,Normal""&8" & NOMBREEESS & Chr(13) & DIRECCIONEESS & Chr(13) & "TELEFONO:  " & TELEFONO
    oHoja.PageSetup.RightHeader = "&""Arial Narrow,Normal""&8" & "&D" & Chr(13) & "&T" & vbCrLf & "Pagina " & "&P" & " de " & "&N"
     
    '--Cabecera
    
    With oLibro.ActiveSheet
    .Range("A1:BE1").Merge
    .Range("A1:BE1").value = textoCabecera
    .Range("A2:BE2").Merge
    .Range("A2:BE2").value = lcSubTitulo
    End With
    
    '--Cuerpo
    cRangeCuerpo = "A3:BE3"
    oExcel.Worksheets(1).Range(cRangeCuerpo).Select
    With (oExcel.Selection.Borders)
        .LineStyle = xlContinuous
    End With
    With (oExcel.Selection.Font)
        .Color = 0
        .Bold = 1
    End With
    
    
    If rs.RecordCount >= 1 Then
    
    rs.MoveFirst
    For I = 0 To rs.Fields.Count - 1
        oHoja.Cells(iFila, I + 1) = rs.Fields(I).Name
        'If lbConfiguraAnchoDeColumna = True Then
        '    oHoja.Cells(iFila, I + 1).ColumnWidth = rs.Fields(I).DefinedSize
        'End If
        
    Next
    
    iFila = iFila + 1
    With oHoja
        .Cells(iFila, iCol).CopyFromRecordset rs
        .Columns.AutoFit
    End With
    
    iFila = iFila + rs.RecordCount
    
    If lcPiePagina <> "" Then
        oHoja.Cells(iFila, iCol) = "________________________________________"
        iFila = iFila + 1
        oHoja.Cells(iFila, iCol) = lcPiePagina
        iFila = iFila + 1
    End If
    
    If lbImpresionHorizontal = True Then
       oHoja.PageSetup.Orientation = xlLandscape
    End If
    
    If oRsPie Is Nothing Then
    Else
        If oRsPie.RecordCount > 0 Then
            iFila = iFila - 1
            oRsPie.MoveFirst
            Do While Not oRsPie.EOF
               For lnCampo = 0 To rs.Fields.Count - 1
                    lcSql = ""
                    If Not IsNull(oRsPie.Fields(lnCampo).value) Then
                       Select Case rs.Fields(lnCampo).Type
                       Case adDate, adDBDate
                            lcSql = Format(oRsPie.Fields(lnCampo).value, "General Date")
                       Case adInteger, adNumeric
                            lcSql = Format(oRsPie.Fields(lnCampo).value, "#########")
                       Case adDecimal, adDouble
                            lcSql = Format(oRsPie.Fields(lnCampo).value, "#########.##")
                       Case adBoolean
                            lcSql = IIf(oRsPie.Fields(lnCampo).value = True, "Verdadero", "Falso")
                       Case Else
                            lcSql = oRsPie.Fields(lnCampo).value
                            If InStr("0123456789", Left(lcSql, 1)) > 0 Then
                               lcSql = "'" & oRsPie.Fields(lnCampo).value
                            End If
                       End Select
                       oHoja.Cells(iFila, lnCampo + 1).value = lcSql
                    End If
               Next
               iFila = iFila + 1
               oRsPie.MoveNext
            Loop
        End If
    End If
    
    
    oExcel.Visible = True
    Set oExcel = Nothing
    Set lcBuscaParametro = Nothing
    
    Screen.MousePointer = vbDefault
    
    Else
        MsgBox ("No Existe Informacion que Exportar"), vbInformation, "Aviso Importante"
    End If

End Sub
Sub ExportarRecordSetAexcelFast1(rs As Recordset, textoCabecera As String)

    Dim oExcel As Excel.Application
    Dim oLibro As Excel.Workbook
    Dim oHoja As Excel.Worksheet
    Dim iFila As Long, iCol As Integer, I As Integer
    Dim cRangeCabecera As String
    Dim cRangeCuerpo As String
    
    Set oExcel = New Excel.Application
    Set oLibro = oExcel.Workbooks.Add
    Set oHoja = oLibro.Worksheets(1)
    Screen.MousePointer = vbHourglass
    iFila = 2
    iCol = 1
    
    '--Cabecera
    
    With oLibro.ActiveSheet
    .Range("A1:BE1").Merge
    .Range("A1:BE1").value = textoCabecera
    End With
    
    '--Cuerpo
    cRangeCuerpo = "A2:BE2"
    oExcel.Worksheets(1).Range(cRangeCuerpo).Select
    With (oExcel.Selection.Borders)
        .LineStyle = xlContinuous
    End With
    With (oExcel.Selection.Font)
        .Color = 0
        .Bold = 1
    End With
    
    
    If rs.RecordCount >= 1 Then
    
    rs.MoveFirst
    For I = 0 To rs.Fields.Count - 1
        oHoja.Cells(iFila, I + 1) = rs.Fields(I).Name
    Next
    
    iFila = iFila + 1
    With oHoja
        .Cells(iFila, iCol).CopyFromRecordset rs
        .Columns.AutoFit
    End With
    
    oExcel.Visible = True
    Set oExcel = Nothing
    Screen.MousePointer = vbDefault
    
    Else
    MsgBox ("No Existe Informacion que Exportar"), vbInformation, "Aviso Importante"
    End If

End Sub



Function AtencionesSeleccionarPorFechaIngresoYespecialidadDelServicio(ldFechaIni As Date, ldFechaFin As Date, lnIdEspecialidadDelServicio As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesSeleccionarPorFechaIngresoYespecialidadDelServicio = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarPorFechaIngresoYespecialidadDelServicio"
        Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, lnIdEspecialidadDelServicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPorFechaIngresoYespecialidadDelServicio = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 30/09/2013 - se cambio a store procedure
Function AtencionesSeleccionarPorFechasDeEgresoAdministrativo(ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
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
      .CommandText = "AtencionesSeleccionarPorFechasDeEgresoAdministrativo"
      Set oParameter = .CreateParameter("@FechaEgresoIni", adDBTimeStamp, adParamInput, 0, Format(ldFechaEgresoIni, "dd/mm/yyyy") & " 00:00:01"): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@FechaEgresoFin", adDBTimeStamp, adParamInput, 0, Format(ldFechaEgresoFin, "dd/mm/yyyy") & " 23:59:59"): .Parameters.Append oParameter
      Set oRecordset = .Execute
      Set oRecordset.ActiveConnection = Nothing
    End With
    
    Set AtencionesSeleccionarPorFechasDeEgresoAdministrativo = oRecordset
    'oRecordSet.Close
    oConexion.Close
    'Set oRecordSet = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
  
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 30/09/2013 - se cambio a store procedure
Sub PreparaTablasDeReportes(ByRef mRs_Cab As Recordset, ByRef mRs_det As Recordset)
        Dim oCommand As New ADODB.Command
        Dim oConexion As New ADODB.Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "RepPreparaEliminaTablasReportes"
            .Execute
        End With
        Set oCommand = Nothing
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "RepPreparaConsultaTablasReportesCabecera"
            Set mRs_Cab = .Execute
            Set mRs_Cab.ActiveConnection = Nothing
        End With
        Set oCommand = Nothing
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "RepPreparaConsultaTablasReportesDetalle"
            Set mRs_det = .Execute
            Set mRs_det.ActiveConnection = Nothing
        End With
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 30/09/2013 - se cambio a store procedure
Function MuerteFetalSeleccionarPorAnioYMes(lnAnio As Long, lnMes As Integer, Cadena As String, Departamento As Long) As Long
  Dim oRsTmp As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  'Dim Consulta As String
  MuerteFetalSeleccionarPorAnioYMes = 0
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open Cadena
  

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "MuerteFetalSeleccionarPorAnioYMes"
    Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, lnAnio): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Mes", adChar, adParamInput, 2, Format(lnMes, "00")): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Departamento", adInteger, adParamInput, 0, Departamento): .Parameters.Append oParameter
    Set oRsTmp = .Execute
    Set oRsTmp.ActiveConnection = Nothing
  End With
  
  If oRsTmp.EOF = True And oRsTmp.BOF = True Then
    MuerteFetalSeleccionarPorAnioYMes = 0
  Else
    If oRsTmp.RecordCount = 0 Then
      MuerteFetalSeleccionarPorAnioYMes = 0
    Else
      MuerteFetalSeleccionarPorAnioYMes = oRsTmp.RecordCount '!Contador
    End If
  End If
  'oRsTmp.Close
  oConexion.Close
  'Set oRsTmp = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 01/10/2013 - se cambio a store procedure
Function MuerteNeoNatalSeleccionarPorAnioYMes(lnAnio As Long, lnMes As Integer, Cadena As String, Departamento As Long) As Long
  Dim oRsTmp As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  
  'Dim Consulta As String
  MuerteNeoNatalSeleccionarPorAnioYMes = 0
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open Cadena
  

  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "MuerteNeoNatalSeleccionarPorAnioYMes"
    Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, lnAnio): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Mes", adChar, adParamInput, 2, Format(lnMes, "00")): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Departamento", adInteger, adParamInput, 0, Departamento): .Parameters.Append oParameter
    Set oRsTmp = .Execute
    Set oRsTmp.ActiveConnection = Nothing
  End With
  Set oCommand = Nothing
  Set oParameter = Nothing
  
  If oRsTmp.EOF = True And oRsTmp.BOF = True Then
    MuerteNeoNatalSeleccionarPorAnioYMes = 0
  Else
    If oRsTmp.RecordCount = 0 Then
      MuerteNeoNatalSeleccionarPorAnioYMes = 0
    Else
      oRsTmp.MoveFirst
      Do While Not oRsTmp.EOF
        If oRsTmp!Diferencia <= 7 Then MuerteNeoNatalSeleccionarPorAnioYMes = MuerteNeoNatalSeleccionarPorAnioYMes + 1
        oRsTmp.MoveNext
      Loop
    End If
  End If
  'oRsTmp.Close
  oConexion.Close
  'Set oRsTmp = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 01/10/2013 - se cambio a store procedure
Function NacidosVivosSeleccionarPorAnioYMes(lnAnio As Long, lnMes As Integer, Cadena As String, Departamento As Long) As Long
  Dim oRsTmp As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  
'  Dim Consulta As String
  NacidosVivosSeleccionarPorAnioYMes = 0
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open Cadena
  
 
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "NacidosVivosSeleccionarPorAnioYMes"
    Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, lnAnio): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Mes", adChar, adParamInput, 2, Format(lnMes, "00")): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Departamento", adInteger, adParamInput, 0, Departamento): .Parameters.Append oParameter
    Set oRsTmp = .Execute
    Set oRsTmp.ActiveConnection = Nothing
  End With
  
  If oRsTmp.EOF = True And oRsTmp.BOF = True Then
    NacidosVivosSeleccionarPorAnioYMes = 0
  Else
    If oRsTmp.RecordCount = 0 Then
      NacidosVivosSeleccionarPorAnioYMes = 0
    Else
      NacidosVivosSeleccionarPorAnioYMes = oRsTmp.RecordCount '!Contador
    End If
  End If
  'oRsTmp.Close
  oConexion.Close
  'Set oRsTmp = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 01/10/2013 - se cambio a store procedure
Function RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes(lnAnio As Long, lnMes As Integer, Cadena As String, Departamento As Long) As Long
  Dim oRsTmp As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  
'  Dim Consulta As String
  RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes = 0
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open Cadena
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes"
    Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, lnAnio): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Mes", adChar, adParamInput, 2, Format(lnMes, "00")): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Departamento", adInteger, adParamInput, 0, Departamento): .Parameters.Append oParameter
    Set oRsTmp = .Execute
    Set oRsTmp.ActiveConnection = Nothing
  End With
  
  If oRsTmp.EOF = True And oRsTmp.BOF = True Then
    RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes = 0
  Else
    If oRsTmp.RecordCount = 0 Then
      RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes = 0
    Else
      RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes = oRsTmp.RecordCount '!Contador
    End If
  End If
  'oRsTmp.Close
  oConexion.Close
  'Set oRsTmp = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 30/09/2013 - se cambio a store procedure
Function BoletasDeServicioParaSOATxCuenta(lnIdCuentaAtencion As Long) As Recordset
    Dim oRecordset As New ADODB.Recordset
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
        .CommandText = "BoletasDeServicioParaSOATxCuenta"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set BoletasDeServicioParaSOATxCuenta = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 30/09/2013 - se cambio a store procedure
Function BoletasDeFarmaciaParaSOATxCuenta(lnIdCuentaAtencion As Long) As Recordset
    Dim oRecordset As New ADODB.Recordset
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
        .CommandText = "BoletasDeFarmaciaParaSOATxCuenta"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set BoletasDeFarmaciaParaSOATxCuenta = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 30/09/2013 - se cambio a store procedure
Function CajaComprobantesPagoParaTipoTarifa(lcNroSerie As String, lcNroDocumento As String, oConexion As Connection) As Recordset
'         Dim oRsCitaPagada1 As New Recordset
'         Dim lcSql As String
        Dim oRecordset As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
         
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CajaComprobantesPagoParaTipoTarifa"
            Set oParameter = .CreateParameter("@NroSerie", adChar, adParamInput, 4, lcNroSerie): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@NroDocumento", adChar, adParamInput, 7, lcNroDocumento): .Parameters.Append oParameter
            Set oRecordset = .Execute
        End With
        Set CajaComprobantesPagoParaTipoTarifa = oRecordset
        'Set oRecordSet = Nothing
        Set oCommand = Nothing
        
End Function

'MODIFICADO POR FRANKLIN CACHAY 02/09/2013 - se cambio a store procedure
Sub DesagregaCuentasItemsXreembolso(ByVal oRsItemsXCuenta1 As Recordset, lnIdComprobantePago As Long, _
                                    lnTotalIdProducto As Double, moConexion As Connection)
    'Dim lcSql As String
    Dim oRsItems As New Recordset
    Dim oRsCuentas As New Recordset
    Dim lbNuevo As Boolean
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lnImpTotal As Double, lnImpProrrateo As Double, lnTotal1 As Double
    On Error GoTo ErrCtas
    If oRsItemsXCuenta1.State = adStateOpen Then
         oRsItemsXCuenta1.Close
    End If
    With oRsItemsXCuenta1
       .Fields.Append "IdPartida", adInteger
       .Fields.Append "IdProducto", adInteger
       .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
       .Fields.Append "NombreProducto", adVarChar, 255, adFldIsNullable
       .Fields.Append "Precio", adDouble
       .Fields.Append "Cantidad", adInteger
       .Fields.Append "Importe", adDouble
       .Fields.Append "ImporteProrrateado", adDouble
       .LockType = adLockOptimistic
       .Open
    End With
    '
      
    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = moConexion
      .CommandTimeout = 150
      .CommandText = "RepDesagregaCuentasItemsXreembolso"
      Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
      Set oRsCuentas = .Execute
    End With
    Set oParameter = Nothing
    Set oCommand = Nothing
  
    lnImpTotal = 0
    If oRsCuentas.RecordCount > 0 Then
       oRsCuentas.MoveFirst
       Do While Not oRsCuentas.EOF
           
            With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = moConexion
              .CommandTimeout = 150
              .CommandText = "RepDesagregaCuentasItemsXreembolsoServicio"
              Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsCuentas.Fields!idCuentaAtencion): .Parameters.Append oParameter
              Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, oRsCuentas.Fields!idFuenteFinanciamiento): .Parameters.Append oParameter
              Set oRsItems = .Execute
            End With
            Set oParameter = Nothing
            Set oCommand = Nothing
            
            If oRsItems.RecordCount > 0 Then
               oRsItems.MoveFirst
               Do While Not oRsItems.EOF
                  lbNuevo = True
                  If oRsItemsXCuenta1.RecordCount > 0 Then
                     oRsItemsXCuenta1.MoveFirst
                     oRsItemsXCuenta1.Find "idProducto=" & oRsItems.Fields!idProducto
                     If Not oRsItemsXCuenta1.EOF Then
                        lbNuevo = False
                     End If
                  End If
                  If lbNuevo = True Then
                     oRsItemsXCuenta1.AddNew
                     oRsItemsXCuenta1.Fields!idPartida = oRsItems.Fields!IdTipoTarifa
                     oRsItemsXCuenta1.Fields!idProducto = oRsItems.Fields!idProducto
                     oRsItemsXCuenta1.Fields!Codigo = oRsItems.Fields!Codigo
                     oRsItemsXCuenta1.Fields!nombreProducto = oRsItems.Fields!nombre
                     oRsItemsXCuenta1.Fields!Precio = oRsItems.Fields!PrecioFinanciado
                  End If
                  oRsItemsXCuenta1.Fields!cantidad = oRsItemsXCuenta1.Fields!cantidad + oRsItems.Fields!CantidadFinanciada
                  oRsItemsXCuenta1.Fields!Importe = oRsItemsXCuenta1.Fields!Importe + oRsItems.Fields!TotalFinanciado
                  oRsItemsXCuenta1.Update
                  lnImpTotal = lnImpTotal + oRsItems.Fields!TotalFinanciado
                  oRsItems.MoveNext
               Loop
            End If
            oRsItems.Close
            oRsCuentas.MoveNext
       Loop
    End If
    oRsCuentas.Close
    Set oCommand = Nothing
    '***proceso de prorrateo
    If oRsItemsXCuenta1.RecordCount > 0 Then
       lnTotal1 = 0
       oRsItemsXCuenta1.MoveFirst
       Do While Not oRsItemsXCuenta1.EOF
          lnImpProrrateo = Round((lnTotalIdProducto * oRsItemsXCuenta1.Fields!Importe) / lnImpTotal, 2)
          lnTotal1 = lnTotal1 + lnImpProrrateo
          oRsItemsXCuenta1.Fields!ImporteProrrateado = lnImpProrrateo
          oRsItemsXCuenta1.Update
          oRsItemsXCuenta1.MoveNext
       Loop
       If lnTotal1 <> lnTotalIdProducto Then
          lnImpProrrateo = lnTotalIdProducto - lnTotal1
          oRsItemsXCuenta1.MoveFirst
          oRsItemsXCuenta1.Fields!ImporteProrrateado = oRsItemsXCuenta1.Fields!ImporteProrrateado + lnImpProrrateo
          oRsItemsXCuenta1.Update
       End If
    End If
    Exit Sub
ErrCtas:
    MsgBox Err.Description
    Resume
End Sub

'MODIFICADO POR FRANKLIN CACHAY 02/09/2013 - se cambio a store procedure
Sub DesagregaCuentasItemsXServHosp(ByVal oRsItemsXCuenta1 As Recordset, lnIdCuentaAtencion As Long, _
                                   lnTotalIdProducto As Double, lnIdServiciosHospitalarios As Long, moConexion As Connection)
    'Dim lcSql As String
    Dim oRsItems As New Recordset
    Dim oRsCuentas As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lbNuevo As Boolean
    Dim lnImpTotal As Double, lnImpProrrateo As Double, lnTotal1 As Double, lnConsumoServicio As Double
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    On Error GoTo ErrCtas1
    If oRsItemsXCuenta1.State = adStateOpen Then
         oRsItemsXCuenta1.Close
    End If
    With oRsItemsXCuenta1
       .Fields.Append "IdPartida", adInteger
       .Fields.Append "IdProducto", adInteger
       .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
       .Fields.Append "NombreProducto", adVarChar, 255, adFldIsNullable
       .Fields.Append "Precio", adDouble
       .Fields.Append "Cantidad", adInteger
       .Fields.Append "Importe", adDouble
       .Fields.Append "ImporteProrrateado", adDouble
       .LockType = adLockOptimistic
       .Open
    End With
    'Farmacia
    lnImpTotal = 0
    'ServicioHospitalario

    
    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = moConexion
      .CommandTimeout = 150
      .CommandText = "DesagregaCuentasItemsXServHosp"
      Set oParameter = .CreateParameter("@IdProductoCpt", adInteger, adParamInput, 0, lnIdServiciosHospitalarios): .Parameters.Append oParameter
      Set oRsItems = .Execute
    End With
    Set oParameter = Nothing
    Set oCommand = Nothing
    
    If oRsItems.RecordCount > 0 Then
        lnConsumoServicio = mo_ReglasFacturacion.RetornaConsumoPacienteServiciosConSeguroPorNroCuenta(lnIdCuentaAtencion)
        oRsItemsXCuenta1.AddNew
        oRsItemsXCuenta1.Fields!idPartida = oRsItems.Fields!IdTipoTarifa
        oRsItemsXCuenta1.Fields!idProducto = oRsItems.Fields!idProductoCPT
        oRsItemsXCuenta1.Fields!Codigo = oRsItems.Fields!Codigo
        oRsItemsXCuenta1.Fields!nombreProducto = oRsItems.Fields!nombre
        oRsItemsXCuenta1.Fields!Precio = 1
        oRsItemsXCuenta1.Fields!cantidad = 1
        oRsItemsXCuenta1.Fields!Importe = oRsItemsXCuenta1.Fields!Importe + lnConsumoServicio
        oRsItemsXCuenta1.Update
        lnImpTotal = lnImpTotal + lnConsumoServicio
    End If
    oRsItems.Close
    '***proceso de prorrateo
    If oRsItemsXCuenta1.RecordCount > 0 Then
       lnTotal1 = 0
       oRsItemsXCuenta1.MoveFirst
       Do While Not oRsItemsXCuenta1.EOF
If oRsItemsXCuenta1.Fields!idProducto = lnIdServiciosHospitalarios Then
lnImpProrrateo = 0
End If
          If oRsItemsXCuenta1.Fields!Importe > 0 Then
            lnImpProrrateo = Round((lnTotalIdProducto * oRsItemsXCuenta1.Fields!Importe) / lnImpTotal, 2)
            lnTotal1 = lnTotal1 + lnImpProrrateo
            oRsItemsXCuenta1.Fields!ImporteProrrateado = lnImpProrrateo
            oRsItemsXCuenta1.Update
          End If
          oRsItemsXCuenta1.MoveNext
       Loop
       If lnTotal1 <> lnTotalIdProducto Then
          lnImpProrrateo = lnTotalIdProducto - lnTotal1
          oRsItemsXCuenta1.MoveFirst
          oRsItemsXCuenta1.Fields!ImporteProrrateado = oRsItemsXCuenta1.Fields!ImporteProrrateado + lnImpProrrateo
          oRsItemsXCuenta1.Update
       End If
    End If
    Set mo_ReglasFacturacion = Nothing
    Exit Sub
ErrCtas1:
    MsgBox Err.Description
    Resume
End Sub




Function BuscaNombreColumna(lnNroCampo As Integer)
    Dim lcCampoTexto As String
    Dim lnCociente As Integer
    Dim lbResiduo As Integer
    If lnNroCampo <= 26 Then
        lcCampoTexto = BuscarLetraAlfabeto(lnNroCampo)
    Else
        lnCociente = IIf(Round(lnNroCampo / 26) < lnNroCampo / 26, Round(lnNroCampo / 26), Round(lnNroCampo / 26) - 1)
        lbResiduo = IIf(lnNroCampo Mod 26 = 0, 26, lnNroCampo Mod 26)
        lcCampoTexto = BuscarLetraAlfabeto(lnCociente) & BuscarLetraAlfabeto(lbResiduo)
    End If
    BuscaNombreColumna = lcCampoTexto
End Function

Function BuscarLetraAlfabeto(lnNumeroLetra As Integer)
    Dim lcLetraAlfabeto As String
    Select Case lnNumeroLetra
    Case 1
        lcLetraAlfabeto = "A"
    Case 2
        lcLetraAlfabeto = "B"
    Case 3
        lcLetraAlfabeto = "C"
    Case 4
        lcLetraAlfabeto = "D"
    Case 5
        lcLetraAlfabeto = "E"
    Case 6
        lcLetraAlfabeto = "F"
    Case 7
        lcLetraAlfabeto = "G"
    Case 8
        lcLetraAlfabeto = "H"
    Case 9
        lcLetraAlfabeto = "I"
    Case 10
        lcLetraAlfabeto = "J"
    Case 11
        lcLetraAlfabeto = "K"
    Case 12
        lcLetraAlfabeto = "L"
    Case 13
        lcLetraAlfabeto = "M"
    Case 14
        lcLetraAlfabeto = "N"
    Case 15
        lcLetraAlfabeto = "O"
    Case 16
        lcLetraAlfabeto = "P"
    Case 17
        lcLetraAlfabeto = "Q"
    Case 18
        lcLetraAlfabeto = "R"
    Case 19
        lcLetraAlfabeto = "S"
    Case 20
        lcLetraAlfabeto = "T"
    Case 21
        lcLetraAlfabeto = "U"
    Case 22
        lcLetraAlfabeto = "V"
    Case 23
        lcLetraAlfabeto = "W"
    Case 24
        lcLetraAlfabeto = "X"
    Case 25
        lcLetraAlfabeto = "Y"
    Case 26
        lcLetraAlfabeto = "Z"
    End Select
    BuscarLetraAlfabeto = lcLetraAlfabeto
End Function

Function GrabarImpresionFichaAtencion(idAtencion As Long) As Boolean
    Dim oAtencionDatosAdicionales As New AtencionesDatosAdicionales
    Dim oDOAtencionDatosAdicionales As New DoAtencionDatosAdicionales
    
    Dim oConexion As New ADODB.Connection

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    oDOAtencionDatosAdicionales.idAtencion = idAtencion
    oDOAtencionDatosAdicionales.seImprimioFicha = True
  
    Set oAtencionDatosAdicionales.Conexion = oConexion
    GrabarImpresionFichaAtencion = oAtencionDatosAdicionales.ModificarImprimioFicha(oDOAtencionDatosAdicionales)
    oConexion.Close
    ms_MensajeError = oAtencionDatosAdicionales.MensajeError
    Set oConexion = Nothing
End Function


'Yamill Palomino
Function FiltrarMedicosPorDepartamentoyEspecialidad(lIdDepartamento As Long, lIdEspecialidad As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
'Dim ssql As String
Dim oConexion As New ADODB.Connection

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
        
        Set FiltrarMedicosPorDepartamentoyEspecialidad = Nothing

        With oCommand
'        .CommandType = adCmdText
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FiltrarMedicosPorDepartamentoyEspecialidad"
        Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, lIdDepartamento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, lIdEspecialidad): .Parameters.Append oParameter
'        .CommandText = ssql
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    
    Set FiltrarMedicosPorDepartamentoyEspecialidad = oRecordset
    oConexion.Close
  Set oConexion = Nothing
   Set oCommand = Nothing
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


Function ReporteProgramacionMedicaPorMedicoyEspecialidad(lIdMedico As Long, daFechaInicio As Date, daFechaFin As Date, lIdEspecialidad As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
'Dim ssql As String
Dim oConexion As New ADODB.Connection

  Set ReporteProgramacionMedicaPorMedicoyEspecialidad = Nothing
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
     With oCommand
'        .CommandType = adCmdText
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "ReporteProgramacionMedicaPorMedicoyEspecialidad"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, Format(daFechaInicio, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(daFechaFin, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lIdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, lIdEspecialidad): .Parameters.Append oParameter

'        .CommandText = ssql
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    oConexion.Close
    Set ReporteProgramacionMedicaPorMedicoyEspecialidad = oRecordset
    ms_MensajeError = ""
  'Set oRecordSet = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'mgaray201411h
Function ReporteCREDSuplementoHierro(daFechaReporte As Date, lIdDepartamento As Long, lIdProvincia As Long, _
                lIdDistrito As Long, ml_EdadMesesInicial As Long, ml_EdadMesesFinal As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection

    Set ReporteCREDSuplementoHierro = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "PerinatalAtencionIndicadoresSuplementoHierro"
        Set oParameter = .CreateParameter("@FechaReporte", adDBTimeStamp, adParamInput, 0, Format(daFechaReporte, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, IIf(lIdDepartamento = 0, Null, lIdDepartamento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProvincia", adInteger, adParamInput, 0, IIf(lIdProvincia = 0, Null, lIdProvincia)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdDistritoDomicilio", adInteger, adParamInput, 0, IIf(lIdDistrito = 0, Null, lIdDistrito)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EdadInicialEnMeses", adInteger, adParamInput, 0, ml_EdadMesesInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EdadFinalEnMeses", adInteger, adParamInput, 0, ml_EdadMesesFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    oConexion.Close
    Set ReporteCREDSuplementoHierro = oRecordset
    ms_MensajeError = ""
    'Set oRecordSet = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
  
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


Function ReporteCREDLactanciaMaternaExclusiva(daFechaReporte As Date, lIdDepartamento As Long, lIdProvincia As Long, _
                lIdDistrito As Long, ml_EdadMesesInicial As Long, ml_EdadMesesFinal As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection

    Set ReporteCREDLactanciaMaternaExclusiva = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "PerinatalAtencionIndicadoresLactanciaMaternaExclusiva"
        Set oParameter = .CreateParameter("@FechaReporte", adDBTimeStamp, adParamInput, 0, Format(daFechaReporte, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, IIf(lIdDepartamento = 0, Null, lIdDepartamento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProvincia", adInteger, adParamInput, 0, IIf(lIdProvincia = 0, Null, lIdProvincia)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdDistritoDomicilio", adInteger, adParamInput, 0, IIf(lIdDistrito = 0, Null, lIdDistrito)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EdadInicialEnMeses", adInteger, adParamInput, 0, ml_EdadMesesInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EdadFinalEnMeses", adInteger, adParamInput, 0, ml_EdadMesesFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    oConexion.Close
    Set ReporteCREDLactanciaMaternaExclusiva = oRecordset
    ms_MensajeError = ""
    'Set oRecordSet = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
  
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function ReporteCREDSesionesDemostrativas(daFechaReporte As Date, lIdDepartamento As Long, lIdProvincia As Long, _
                lIdDistrito As Long, ml_EdadMesesInicial As Long, ml_EdadMesesFinal As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection

    Set ReporteCREDSesionesDemostrativas = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "PerinatalAtencionIndicadoresSesionesDemostrativas"
        Set oParameter = .CreateParameter("@FechaReporte", adDBTimeStamp, adParamInput, 0, Format(daFechaReporte, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, IIf(lIdDepartamento = 0, Null, lIdDepartamento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProvincia", adInteger, adParamInput, 0, IIf(lIdProvincia = 0, Null, lIdProvincia)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdDistritoDomicilio", adInteger, adParamInput, 0, IIf(lIdDistrito = 0, Null, lIdDistrito)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EdadInicialEnMeses", adInteger, adParamInput, 0, ml_EdadMesesInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EdadFinalEnMeses", adInteger, adParamInput, 0, ml_EdadMesesFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    oConexion.Close
    Set ReporteCREDSesionesDemostrativas = oRecordset
    ms_MensajeError = ""
    'Set oRecordSet = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
  
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function ReporteCREDVacunaRatavirusNeumococo(daFechaReporte As Date, lIdDepartamento As Long, lIdProvincia As Long, _
                lIdDistrito As Long, ml_EdadMesesInicial As Long, ml_EdadMesesFinal As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection

    Set ReporteCREDVacunaRatavirusNeumococo = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "PerinatalAtencionIndicadoresVacunasNeumococoRatavirus"
        Set oParameter = .CreateParameter("@FechaReporte", adDBTimeStamp, adParamInput, 0, Format(daFechaReporte, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtenInteItemPlan", adInteger, adParamInput, 0, sighItemPlanIntegral.Inmunizacion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, IIf(lIdDepartamento = 0, Null, lIdDepartamento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProvincia", adInteger, adParamInput, 0, IIf(lIdProvincia = 0, Null, lIdProvincia)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdDistritoDomicilio", adInteger, adParamInput, 0, IIf(lIdDistrito = 0, Null, lIdDistrito)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EdadInicialEnMeses", adInteger, adParamInput, 0, ml_EdadMesesInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EdadFinalEnMeses", adInteger, adParamInput, 0, ml_EdadMesesFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    oConexion.Close
    Set ReporteCREDVacunaRatavirusNeumococo = oRecordset
    ms_MensajeError = ""
    'Set oRecordSet = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
  
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function ReporteCREDCompleto(daFechaReporte As Date, lIdDepartamento As Long, lIdProvincia As Long, _
                lIdDistrito As Long, ml_EdadMesesInicial As Long, ml_EdadMesesFinal As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection

    Set ReporteCREDCompleto = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "PerinatalAtencionIndicadoresCREDCompleto"
        Set oParameter = .CreateParameter("@FechaReporte", adDBTimeStamp, adParamInput, 0, Format(daFechaReporte, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtenInteItemPlan", adInteger, adParamInput, 0, sighItemPlanIntegral.Desarrollo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, IIf(lIdDepartamento = 0, Null, lIdDepartamento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProvincia", adInteger, adParamInput, 0, IIf(lIdProvincia = 0, Null, lIdProvincia)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdDistritoDomicilio", adInteger, adParamInput, 0, IIf(lIdDistrito = 0, Null, lIdDistrito)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EdadInicialEnMeses", adInteger, adParamInput, 0, ml_EdadMesesInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EdadFinalEnMeses", adInteger, adParamInput, 0, ml_EdadMesesFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    oConexion.Close
    Set ReporteCREDCompleto = oRecordset
    ms_MensajeError = ""
    'Set oRecordSet = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
  
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function ReporteCREDIndicadoresTotalPaciente(daFechaReporte As Date, lIdDepartamento As Long, lIdProvincia As Long, _
        lIdDistrito As Long, ml_EdadMesesInicial As Long, ml_EdadMesesFinal As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection

    Set ReporteCREDIndicadoresTotalPaciente = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "PerinatalAtencionIndicadoresTotalPacientes"
        Set oParameter = .CreateParameter("@FechaReporte", adDBTimeStamp, adParamInput, 0, Format(daFechaReporte, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, IIf(lIdDepartamento = 0, Null, lIdDepartamento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProvincia", adInteger, adParamInput, 0, IIf(lIdProvincia = 0, Null, lIdProvincia)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdDistritoDomicilio", adInteger, adParamInput, 0, IIf(lIdDistrito = 0, Null, lIdDistrito)): .Parameters.Append oParameter
        
        Set oParameter = .CreateParameter("@EdadInicialEnMeses", adInteger, adParamInput, 0, ml_EdadMesesInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EdadFinalEnMeses", adInteger, adParamInput, 0, ml_EdadMesesFinal): .Parameters.Append oParameter
        
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    oConexion.Close
    Set ReporteCREDIndicadoresTotalPaciente = oRecordset
    ms_MensajeError = ""
    'Set oRecordSet = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
  
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function AtencionesCensoEstanciaHospitalariaPacientes() As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection

    Set AtencionesCensoEstanciaHospitalariaPacientes = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "AtencionesCensoEstanciaHospitalariaPacientes"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    oConexion.Close
    Set AtencionesCensoEstanciaHospitalariaPacientes = oRecordset
    ms_MensajeError = ""
    'Set oRecordSet = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
  
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


'debb-25/08/2015
Function ReporteIngresosHospitalariosXfuente(lnNroDiasEstancia As Long, lnIdFuenteFinanciamiento As Long, _
                                             FechaIngIni As Date, FechaIngFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
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
            .CommandText = "ReporteIngresosHospitalariosXfuente"
            Set oParameter = .CreateParameter("@FuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@DiasEstancia", adInteger, adParamInput, 0, lnNroDiasEstancia): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, FechaIngIni): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, FechaIngFin): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
        End With
    Set ReporteIngresosHospitalariosXfuente = oRecordset
    ms_MensajeError = ""
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


'debb-25/08/2015
Function SisFuaAtencionFuasPorCuenta(lnIdCuentaAtencion As Long, oConexionExterna As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
       With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexionExterna
            .CommandTimeout = 150
            .CommandText = "SisFuaAtencionFuasPorCuenta"
            Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
            Set oRecordset = .Execute
            
        End With
    Set SisFuaAtencionFuasPorCuenta = oRecordset
    ms_MensajeError = ""

    Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

