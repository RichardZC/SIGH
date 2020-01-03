VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistoriasSolicitadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla HistoriasSolicitadas
'        Programado por: Castro W
'        Fecha: Noviembre 2006
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property

'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla HistoriasSolicitadas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOHistoriaSolicitada) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "HistoriasSolicitadasAgregar"
           Set oParameter = .CreateParameter("@IdMotivo", adInteger, adParamInput, 0, IIf(oTabla.IdMotivo = 0, Null, oTabla.IdMotivo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraRequerida", adChar, adParamInput, 5, IIf(oTabla.HoraRequerida = "", Null, oTabla.HoraRequerida)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaRequerida", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fecharequerida = 0, Null, oTabla.fecharequerida)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraSolicitud", adChar, adParamInput, 5, IIf(oTabla.HoraSolicitud = "", Null, oTabla.HoraSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaSolicitud", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaSolicitud = 0, Null, oTabla.FechaSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, 0, oTabla.IdPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdHistoriaSolicitada", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoSolicita", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoSolicita = 0, Null, oTabla.IdEmpleadoSolicita)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMovimiento = 0, Null, oTabla.IdMovimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 150, IIf(oTabla.Observacion = "", Null, oTabla.Observacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           .Execute
           oTabla.IdHistoriaSolicitada = .Parameters("@IdHistoriaSolicitada")
   End With
 
   Insertar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla HistoriasSolicitadas
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOHistoriaSolicitada) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "HistoriasSolicitadasModificar"
           Set oParameter = .CreateParameter("@IdMotivo", adInteger, adParamInput, 0, IIf(oTabla.IdMotivo = 0, Null, oTabla.IdMotivo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraRequerida", adChar, adParamInput, 5, IIf(oTabla.HoraRequerida = "", Null, oTabla.HoraRequerida)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaRequerida", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fecharequerida = 0, Null, oTabla.fecharequerida)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraSolicitud", adChar, adParamInput, 5, IIf(oTabla.HoraSolicitud = "", Null, oTabla.HoraSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaSolicitud", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaSolicitud = 0, Null, oTabla.FechaSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdHistoriaSolicitada", adInteger, adParamInput, 0, IIf(oTabla.IdHistoriaSolicitada = 0, Null, oTabla.IdHistoriaSolicitada)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoSolicita", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoSolicita = 0, Null, oTabla.IdEmpleadoSolicita)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMovimiento = 0, Null, oTabla.IdMovimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 150, IIf(oTabla.Observacion = "", Null, oTabla.Observacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter

       .Execute
   End With
 
   Modificar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla HistoriasSolicitadas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOHistoriaSolicitada) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "HistoriasSolicitadasEliminar"
           Set oParameter = .CreateParameter("@IdHistoriaSolicitada", adInteger, adParamInput, 0, IIf(oTabla.IdHistoriaSolicitada = 0, Null, oTabla.IdHistoriaSolicitada)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   Eliminar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla HistoriasSolicitadas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOHistoriaSolicitada) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "HistoriasSolicitadasSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdHistoriaSolicitada", adInteger, adParamInput, 0, oTabla.IdHistoriaSolicitada): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdMotivo = IIf(IsNull(oRecordset!IdMotivo), 0, oRecordset!IdMotivo)
       oTabla.HoraRequerida = IIf(IsNull(oRecordset!HoraRequerida), "", oRecordset!HoraRequerida)
       oTabla.fecharequerida = IIf(IsNull(oRecordset!fecharequerida), 0, oRecordset!fecharequerida)
       oTabla.HoraSolicitud = IIf(IsNull(oRecordset!HoraSolicitud), "", oRecordset!HoraSolicitud)
       oTabla.FechaSolicitud = IIf(IsNull(oRecordset!FechaSolicitud), 0, oRecordset!FechaSolicitud)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.IdHistoriaSolicitada = IIf(IsNull(oRecordset!IdHistoriaSolicitada), 0, oRecordset!IdHistoriaSolicitada)
       oTabla.IdEmpleadoSolicita = IIf(IsNull(oRecordset!IdEmpleadoSolicita), 0, oRecordset!IdEmpleadoSolicita)
       oTabla.IdMovimiento = IIf(IsNull(oRecordset!IdMovimiento), 0, oRecordset!IdMovimiento)
       oTabla.Observacion = IIf(IsNull(oRecordset!Observacion), "", oRecordset!Observacion)
       oTabla.IdServicio = IIf(IsNull(oRecordset!IdServicio), 0, oRecordset!IdServicio)
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function Filtrar(ByVal oDOPaciente As DoPaciente, ByVal oDOHistoria As DOHistoriaSolicitada, ByVal oDOArchiveroServ As DOArchiveroServicio) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    If oDOArchiveroServ.IdEmpleado <> 0 Then
        sWhere = sWhere + " HistoriasSolicitadas.IdServicio in (select IdServicio from ArchiveroServicio "
        sWhere = sWhere + " where IdEmpleado = " & oDOArchiveroServ.IdEmpleado & ") and "
    End If

    If oDOPaciente.ApellidoPaterno <> "" Then
        sWhere = sWhere + " Pacientes.ApellidoPaterno like '" + oDOPaciente.ApellidoPaterno + "%' and "
    End If
    If oDOPaciente.ApellidoMaterno <> "" Then
         sWhere = sWhere + " Pacientes.ApellidoMaterno like '" + oDOPaciente.ApellidoMaterno + "%' and "
    End If
    If oDOPaciente.PrimerNombre <> "" Then
         sWhere = sWhere + " Pacientes.PrimerNombre like '" + oDOPaciente.PrimerNombre + "%' and "
    End If
    If oDOPaciente.NroHistoriaClinica <> 0 Then
         sWhere = sWhere + " Pacientes.NroHistoriaClinica ='" & oDOPaciente.NroHistoriaClinica & "' and "
    End If
    If oDOHistoria.FechaSolicitud <> 0 Then
         sWhere = sWhere + " HistoriasSolicitadas.FechaSolicitud = '" & oDOHistoria.FechaSolicitud & "' and "
    End If
    If oDOHistoria.fecharequerida <> 0 Then
         sWhere = sWhere + " HistoriasSolicitadas.FechaRequerida = '" & oDOHistoria.fecharequerida & "' and "
    End If
    If sWhere <> "" Then
         sSql = sSql + Left(sWhere, Len(sWhere) - 4)
    End If
    sSql = sSql + " order by convert(datetime, HistoriasSolicitadas.FechaSolicitud,103) desc, Pacientes.NroHistoriaClinica"
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HistoriasSolicitadasSegunFiltro"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set Filtrar = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarPorIdPaciente(ByVal lIdPaciente As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HistoriasSolicitadasXidPaciente"
        Set oParameter = .CreateParameter("@lIdPaciente", adInteger, adParamInput, 0, lIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorIdPaciente = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarPorIdMovimiento(ByVal lIdMovimiento As Long) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HistoriasSolicitadasXidMovimiento"
        Set oParameter = .CreateParameter("@lIdMovimiento", adInteger, adParamInput, 0, lIdMovimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorIdMovimiento = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ActualizarIdMovimiento(rsMovimientos As Recordset) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
   rsMovimientos.MoveFirst
   Do While Not rsMovimientos.EOF
        If Not IsNull(rsMovimientos!IdHistoriaSolicitada) Then
             With oCommand
                 .CommandType = adCmdStoredProc
                 Set .ActiveConnection = mo_Conexion
                 .CommandTimeout = 150
                 .CommandText = "HistoriasSolicitadasActualizarIdMovimiento"
                 Set oParameter = .CreateParameter("@IdHistoriaSolicitada", adInteger, adParamInput, 0, rsMovimientos!IdHistoriaSolicitada): .Parameters.Append oParameter
                 Set oParameter = .CreateParameter("@IdEstadoRegistro", adVarChar, adParamInput, 1, rsMovimientos!IdEstadoRegistro): .Parameters.Append oParameter
                 Set oParameter = .CreateParameter("@Seleccionar", adBoolean, adParamInput, 0, rsMovimientos!Seleccionar): .Parameters.Append oParameter
                 Set oParameter = .CreateParameter("@IdMovimientoHistoria", adInteger, adParamInput, 0, rsMovimientos!IdMovimientoHistoria): .Parameters.Append oParameter
                 .Execute
            End With
            Set oParameter = Nothing
            Set oCommand = Nothing
        End If
        rsMovimientos.MoveNext
   Loop
   ActualizarIdMovimiento = True
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function EliminarIdMovimiento(rsMovimientos As Recordset) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
   rsMovimientos.MoveFirst
   Do While Not rsMovimientos.EOF
        If Not IsNull(rsMovimientos!IdHistoriaSolicitada) Then
             With oCommand
                 .CommandType = adCmdStoredProc
                 Set .ActiveConnection = mo_Conexion
                 .CommandTimeout = 150
                 .CommandText = "HistoriasSolicitadasEliminarIdMovimiento"
                 Set oParameter = .CreateParameter("@IdHistoriaSolicitada", adInteger, adParamInput, 0, rsMovimientos!IdHistoriaSolicitada): .Parameters.Append oParameter
                 .Execute
             End With
             Set oParameter = Nothing
             Set oCommand = Nothing
        End If
        rsMovimientos.MoveNext
   Loop
   EliminarIdMovimiento = True
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function EliminarPorCita(oCita As DoCita) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HistoriasSolicitadasEliminarPorCita"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oCita.IdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, oCita.fecha): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@HoraInicio", adVarChar, adParamInput, 5, oCita.HoraInicio): .Parameters.Append oParameter
        .Execute
    End With
    EliminarPorCita = True
    Set oCommand = Nothing
    Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarPorArchivero(ByVal lIdEmpleado As Long, lIdMotivo As Long, lIdServiciodestino As Long, _
                                 sOperadorFecha As String, sFechaDesde As String, sFechaHasta As String, _
                                 lnNroHistoriaClinica As Long, esFechaRequer As Boolean, _
                                 lnIdArchivoClinico As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String, lcParametro231 As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
    lcParametro231 = lcBuscaParametro.SeleccionaFilaParametro(231)
    sSql = ""
    If lIdServiciodestino <> 0 Then
        sSql = sSql + " and ServicioDestino.IdServicio = " & lIdServiciodestino
    End If
    If sOperadorFecha <> "" Then
        If sFechaHasta = "" Then
            If esFechaRequer = True Then
               sSql = sSql + " and  (dbo.HistoriasSolicitadas.FechaRequerida " & sOperadorFecha & " CONVERT(DATETIME, '" & Format(sFechaDesde, "yyyy-mm-dd") & " 00:00:00', 102))"
            Else
               sSql = sSql + " and convert(char(10),FechaSolicitud,103) " & sOperadorFecha & "'" & sFechaDesde & "'"
            End If
        Else
            If esFechaRequer = True Then
               sSql = sSql + " and  (dbo.HistoriasSolicitadas.FechaRequerida >= CONVERT(DATETIME, '" & Format(sFechaDesde, "yyyy-mm-dd") & " 00:00:00', 102))"
               sSql = sSql + " and  (dbo.HistoriasSolicitadas.FechaRequerida <= CONVERT(DATETIME, '" & Format(sFechaHasta, "yyyy-mm-dd") & " 23:59:59', 102))"
            Else
               sSql = sSql + " and  FechaSolicitud Between (CONVERT(DATETIME,'" & sFechaDesde & "',103)) and (CONVERT(DATETIME,'" & sFechaHasta & "',103))"
            End If
        End If
    End If
    If lnNroHistoriaClinica > 0 Then
       sSql = sSql + " and Pacientes.NroHistoriaClinica = " & lnNroHistoriaClinica
    End If
    sSql = sSql + " order by Pacientes.NroHistoriaClinica,Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HistoriasSolicitadasPorFiltro"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1500, sSql): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdEmpleado", adVarChar, adParamInput, 10, Trim(Str(lIdEmpleado))): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdArchivoClinico", adVarChar, adParamInput, 10, Trim(Str(lnIdArchivoClinico))): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcParametro231", adVarChar, adParamInput, 1, lcParametro231): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set SeleccionarPorArchivero = oRecordset
    Set oCommand = Nothing
    Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function HistoriasSolicitadasSeleccionarDetallePorIdPaciente(lIdPaciente As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lIdServicio As Long
Dim oRecordset As New Recordset
 
   Set HistoriasSolicitadasSeleccionarDetallePorIdPaciente = Nothing
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "HistoriasSolicitadasSeleccionarDetallePorIdPaciente"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
 
    Set oRecordset.ActiveConnection = Nothing
    Set HistoriasSolicitadasSeleccionarDetallePorIdPaciente = oRecordset
   
    ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function HistoriasSolicitadasSeleccionarPorIdAtencion(lnIdAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lIdServicio As Long
Dim oRecordset As New Recordset
   Set HistoriasSolicitadasSeleccionarPorIdAtencion = Nothing
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "HistoriasSolicitadasSeleccionarPorIdAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
    Set oRecordset.ActiveConnection = Nothing
    Set HistoriasSolicitadasSeleccionarPorIdAtencion = oRecordset
    ms_MensajeError = ""
    Set oCommand = Nothing
    Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function EliminarIdMovimientoIdPaciente(rsMovimientos As Recordset) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
   rsMovimientos.MoveFirst
   Do While Not rsMovimientos.EOF
        If Not IsNull(rsMovimientos!IdHistoriaSolicitada) Then
             With oCommand
                 .CommandType = adCmdStoredProc
                 Set .ActiveConnection = mo_Conexion
                 .CommandTimeout = 150
                 .CommandText = "HistoriasSolicitadasEliminarIdMovimientoIdPaciente"
                 Set oParameter = .CreateParameter("@IdHistoriaSolicitada", adInteger, adParamInput, 0, rsMovimientos!IdHistoriaSolicitada): .Parameters.Append oParameter
                 Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, rsMovimientos!IdPaciente): .Parameters.Append oParameter
                 .Execute
             End With
             Set oParameter = Nothing
             Set oCommand = Nothing
        End If
        rsMovimientos.MoveNext
   Loop
   EliminarIdMovimientoIdPaciente = True
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

