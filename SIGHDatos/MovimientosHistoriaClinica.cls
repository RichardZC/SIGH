VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MovimientosHistoriaClinica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla MovimientosHistoriaClinica
'        Programado por: Barrantes D
'        Fecha: Mayo 2009
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
'   Descripción:    Inserta un registro a la tabla MovimientosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOMovimientoHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "MovimientosHistoriaClinicaAgregar"
           Set oParameter = .CreateParameter("@NroFolios", adInteger, adParamInput, 0, IIf(oTabla.NroFolios = 0, Null, oTabla.NroFolios)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioDestino", adInteger, adParamInput, 0, IIf(oTabla.IdServicioDestino = 0, Null, oTabla.IdServicioDestino)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdServicioOrigen = 0, Null, oTabla.IdServicioOrigen)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 100, IIf(oTabla.Observacion = "", Null, oTabla.Observacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMotivo", adInteger, adParamInput, 0, IIf(oTabla.idMotivo = 0, Null, oTabla.idMotivo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaMovimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaMovimiento = 0, Null, oTabla.FechaMovimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoRecepcion", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoRecepcion = 0, Null, oTabla.IdEmpleadoRecepcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoTransporte", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoTransporte = 0, Null, oTabla.IdEmpleadoTransporte)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoArchivo", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoArchivo = 0, Null, oTabla.IdEmpleadoArchivo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGrupoMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdGrupoMovimiento = 0, Null, oTabla.IdGrupoMovimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter

       .Execute
           oTabla.idMovimiento = .Parameters("@IdMovimiento")
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
'   Descripción:    modifica un registro a la tabla MovimientosHistoriaClinica
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOMovimientoHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "MovimientosHistoriaClinicaModificar"
           Set oParameter = .CreateParameter("@NroFolios", adInteger, adParamInput, 0, IIf(oTabla.NroFolios = 0, Null, oTabla.NroFolios)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioDestino", adInteger, adParamInput, 0, IIf(oTabla.IdServicioDestino = 0, Null, oTabla.IdServicioDestino)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdServicioOrigen = 0, Null, oTabla.IdServicioOrigen)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 100, IIf(oTabla.Observacion = "", Null, oTabla.Observacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMotivo", adInteger, adParamInput, 0, IIf(oTabla.idMotivo = 0, Null, oTabla.idMotivo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaMovimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaMovimiento = 0, Null, oTabla.FechaMovimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, IIf(oTabla.idMovimiento = 0, Null, oTabla.idMovimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoRecepcion", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoRecepcion = 0, Null, oTabla.IdEmpleadoRecepcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoTransporte", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoTransporte = 0, Null, oTabla.IdEmpleadoTransporte)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoArchivo", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoArchivo = 0, Null, oTabla.IdEmpleadoArchivo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGrupoMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdGrupoMovimiento = 0, Null, oTabla.IdGrupoMovimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria):  .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla MovimientosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOMovimientoHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "MovimientosHistoriaClinicaEliminar"
           Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, IIf(oTabla.idMovimiento = 0, Null, oTabla.idMovimiento)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla MovimientosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOMovimientoHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "MovimientosHistoriaClinicaSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, oTabla.idMovimiento): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.NroFolios = IIf(IsNull(oRecordset!NroFolios), 0, oRecordset!NroFolios)
       oTabla.IdServicioDestino = IIf(IsNull(oRecordset!IdServicioDestino), 0, oRecordset!IdServicioDestino)
       oTabla.IdServicioOrigen = IIf(IsNull(oRecordset!IdServicioOrigen), 0, oRecordset!IdServicioOrigen)
       oTabla.Observacion = IIf(IsNull(oRecordset!Observacion), "", oRecordset!Observacion)
       oTabla.idMotivo = IIf(IsNull(oRecordset!idMotivo), 0, oRecordset!idMotivo)
       oTabla.FechaMovimiento = IIf(IsNull(oRecordset!FechaMovimiento), 0, oRecordset!FechaMovimiento)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.idMovimiento = IIf(IsNull(oRecordset!idMovimiento), 0, oRecordset!idMovimiento)
       oTabla.IdEmpleadoRecepcion = IIf(IsNull(oRecordset!IdEmpleadoRecepcion), 0, oRecordset!IdEmpleadoRecepcion)
       oTabla.IdEmpleadoTransporte = IIf(IsNull(oRecordset!IdEmpleadoTransporte), 0, oRecordset!IdEmpleadoTransporte)
       oTabla.IdEmpleadoArchivo = IIf(IsNull(oRecordset!IdEmpleadoArchivo), 0, oRecordset!IdEmpleadoArchivo)
       oTabla.IdGrupoMovimiento = IIf(IsNull(oRecordset!IdGrupoMovimiento), 0, oRecordset!IdGrupoMovimiento)
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

Function AgregarVarios(oMovimiento As DOMovimientoHistoriaClinica, rsMovimientos As Recordset) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
'Dim oDOMovimiento As New DOMovimientoHistoriaClinica
Dim oParametros As New Parametros
Dim IdGrupoMovimiento As Long

   AgregarVarios = False
   
   Set oParametros.Conexion = mo_Conexion
   IdGrupoMovimiento = oParametros.GenerarIdGrupoMovimiento()
   If IdGrupoMovimiento <= 0 Then
        MsgBox oParametros.MensajeError, vbInformation, "Movimientos"
        Exit Function
   End If
   
   rsMovimientos.MoveFirst
   
    Do While Not rsMovimientos.EOF
        If rsMovimientos!Seleccionar Then
            oMovimiento.IdGrupoMovimiento = IdGrupoMovimiento
            oMovimiento.IdHistoriaSolicitada = rsMovimientos!IdHistoriaSolicitada
            oMovimiento.IdPaciente = rsMovimientos!IdPaciente
            oMovimiento.IdServicioDestino = rsMovimientos!IdServicioDestino
            oMovimiento.IdServicioOrigen = IIf(IsNull(rsMovimientos!IdServicioOrigen), 0, rsMovimientos!IdServicioOrigen)
            oMovimiento.NroFolios = rsMovimientos!NroFolios
            oMovimiento.idAtencion = rsMovimientos!idAtencion
            oMovimiento.Observacion = IIf(IsNull(rsMovimientos!Observaciones), "", rsMovimientos!Observaciones)  'debb-03/11/2015
            If Not Insertar(oMovimiento) Then
                Exit Function
            End If
            rsMovimientos!IdMovimientoHistoria = oMovimiento.idMovimiento
        End If
        rsMovimientos.MoveNext
    Loop
    
   AgregarVarios = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function ModificarVarios(oMovimiento As DOMovimientoHistoriaClinica, rsMovimientos As Recordset) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oDOMovimiento As New DOMovimientoHistoriaClinica
 
   ModificarVarios = False

    rsMovimientos.MoveFirst
    Do While Not rsMovimientos.EOF
        oMovimiento.idMovimiento = rsMovimientos!IdMovimientoHistoria
        oMovimiento.IdHistoriaSolicitada = IIf(IsNull(rsMovimientos!IdHistoriaSolicitada), 0, rsMovimientos!IdHistoriaSolicitada)
        oMovimiento.IdPaciente = rsMovimientos!IdPaciente
        oMovimiento.IdServicioDestino = rsMovimientos!IdServicioDestino
        oMovimiento.IdServicioOrigen = IIf(IsNull(rsMovimientos!IdServicioOrigen), 0, rsMovimientos!IdServicioOrigen)
        oMovimiento.NroFolios = IIf(IsNull(rsMovimientos!NroFolios), 0, rsMovimientos!NroFolios)
        oMovimiento.idAtencion = rsMovimientos!idAtencion
        Select Case rsMovimientos!IdEstadoRegistro
        Case "A"
            If rsMovimientos!Seleccionar Then
                If Not Insertar(oMovimiento) Then
                    Exit Function
                End If
            End If
        Case "M"
            If rsMovimientos!Seleccionar Then
                If Not Modificar(oMovimiento) Then
                    Exit Function
                End If
            Else
                If Not Eliminar(oMovimiento) Then
                    Exit Function
                End If
            End If
        Case "E"
        End Select
        rsMovimientos.MoveNext
    Loop
        
   ModificarVarios = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function EliminarVarios(rsMovimientos As Recordset) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oDOMovimiento As New DOMovimientoHistoriaClinica
 
   EliminarVarios = False
   
        rsMovimientos.MoveFirst
        Do While Not rsMovimientos.EOF
            oDOMovimiento.idMovimiento = rsMovimientos!IdMovimientoHistoria
            oDOMovimiento.IdPaciente = rsMovimientos!IdPaciente
            'If Not Eliminar(oDOMovimiento) Then
            If Not MovimientosHistoriaClinicaEliminarMovIdPaciente(oDOMovimiento) Then
                Exit Function
            End If
            rsMovimientos.MoveNext
         Loop
    
   EliminarVarios = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function Filtrar(ByVal oDOPaciente As DoPaciente, ByVal oDOHistoria As DOMovimientoHistoriaClinica, _
                 lcApellidoVacio As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    If oDOPaciente.ApellidoPaterno <> "" Then
        If oDOPaciente.ApellidoPaterno = lcApellidoVacio Then
           sWhere = sWhere + " Pacientes.ApellidoPaterno ='" + oDOPaciente.ApellidoPaterno + "' and "
        Else
           sWhere = sWhere + " Pacientes.ApellidoPaterno like '" + oDOPaciente.ApellidoPaterno + "%' and "
        End If
    End If
    If oDOPaciente.ApellidoMaterno <> "" Then
         If oDOPaciente.ApellidoMaterno = lcApellidoVacio Then
            sWhere = sWhere + " Pacientes.ApellidoMaterno = '" + oDOPaciente.ApellidoMaterno + "' and "
         Else
            sWhere = sWhere + " Pacientes.ApellidoMaterno like '" + oDOPaciente.ApellidoMaterno + "%' and "
         End If
    End If
    If oDOPaciente.PrimerNombre <> "" Then
         sWhere = sWhere + " Pacientes.PrimerNombre like '" + oDOPaciente.PrimerNombre + "%' and "
    End If
    If oDOPaciente.NroHistoriaClinica <> 0 Then
         sWhere = sWhere + " Pacientes.NroHistoriaClinica ='" & oDOPaciente.NroHistoriaClinica & "' and "
    End If
    If oDOHistoria.FechaMovimiento <> 0 Then
         sWhere = sWhere + " convert(char(10), MovimientosHistoriaClinica.FechaMovimiento, 103) = '" & oDOHistoria.FechaMovimiento & "' and "
    End If
    If oDOPaciente.FichaFamiliar <> "" Then
       sWhere = sWhere + " dbo.Pacientes.FichaFamiliar like '%" & oDOPaciente.FichaFamiliar & "%' and "
    End If
    If sWhere <> "" Then
         sSql = sSql + " where " + Left(sWhere, Len(sWhere) - 4)
    End If
    'sSql = sSql + " order by MovimientosHistoriaClinica.FechaMovimiento desc,dbo.MovimientosHistoriaClinica.IdMovimiento desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
    sSql = sSql + " order by dbo.MovimientosHistoriaClinica.IdMovimiento desc"
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "MovimientosHistoriaClinicaFiltrar"
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

Function UltimaUbicacionDeHistoria(lIdPaciente As Long) As Long
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lIdServicio As Long
 
   UltimaUbicacionDeHistoria = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "MovimientosHistoriaClinicaUltimaUbicacion"
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(lIdPaciente = 0, Null, lIdPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamOutput, 0): .Parameters.Append oParameter
            .Execute
           lIdServicio = .Parameters("@IdServicio")
   End With
 
   UltimaUbicacionDeHistoria = lIdServicio
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function MovimientosHistoriasClinicasPorIdGrupo(lIdGrupo As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lIdServicio As Long
Dim oRecordset As New Recordset
 
   Set MovimientosHistoriasClinicasPorIdGrupo = Nothing
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "MovimientosHistoriasClinicasPorIdGrupo"
            Set oParameter = .CreateParameter("@IdGrupo", adInteger, adParamInput, 0, lIdGrupo): .Parameters.Append oParameter
            Set oRecordset = .Execute
   End With
 
    Set oRecordset.ActiveConnection = Nothing
    Set MovimientosHistoriasClinicasPorIdGrupo = oRecordset
    
    ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function




Function MovimientosHistoriasClinicasDetallePorIdPaciente(lIdPaciente As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lIdServicio As Long
Dim oRecordset As New Recordset

    Set MovimientosHistoriasClinicasDetallePorIdPaciente = Nothing
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "MovimientosHistoriasClinicasDetallePorIdPaciente"
       Set oParameter = .CreateParameter("@IdGrupo", adInteger, adParamInput, 0, lIdPaciente): .Parameters.Append oParameter
       Set oRecordset = .Execute
    End With
 
    Set oRecordset.ActiveConnection = Nothing
    Set MovimientosHistoriasClinicasDetallePorIdPaciente = oRecordset
    
    ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function EsUltimoMovimiento(lIdPaciente As Long, lIdMovimiento As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   EsUltimoMovimiento = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "MovimientosHistoriaClinicaEsUltimoMovimiento"
            Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(lIdPaciente = 0, Null, lIdPaciente)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, IIf(lIdMovimiento = 0, Null, lIdMovimiento)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@EsUltimo", adInteger, adParamOutput, 0): .Parameters.Append oParameter
            .Execute
            EsUltimoMovimiento = .Parameters("@EsUltimo")
   End With
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function MovimientosHistoriasClinicasParaDevolverPorNroHistoria(lIdServiciodestino As Long, sFechaMovimiento As String, _
                            lIdUsuario As Long, lnNroHistoriaClinica As Long, lnIdArchivoClinico As Long, _
                            lnIdPacienteSolo As Long) As Recordset
 On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lIdServicio As Long
Dim oRecordset As New Recordset
Dim rsPaciente As New Recordset
Dim rsMaxMov As New Recordset
Dim sSql As String
Dim lcParametro231 As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "tempMovimientosEliminarXidUsuario"
        Set oParameter = .CreateParameter("@lIdUsuario", adInteger, adParamInput, 0, lIdUsuario): .Parameters.Append oParameter
        .Execute
    End With
    Dim strWhere As String
    If lnIdPacienteSolo = 0 Then
        strWhere = ""
    End If
    '
    'rsPaciente.Open sSql, mo_Conexion
    Set oCommand = Nothing
    Set oParameter = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "MovimientosHistoriaClinicaPorFiltro"
        Set oParameter = .CreateParameter("@sFechaMovimiento", adDBTimeStamp, adParamInput, 0, CDate(sFechaMovimiento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdServiciodestino", adInteger, adParamInput, 0, lIdServiciodestino): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdPacienteSolo", adInteger, adParamInput, 0, lnIdPacienteSolo): .Parameters.Append oParameter
        Set rsPaciente = .Execute
    End With
    '
    If rsPaciente.RecordCount > 0 Then
        Do While Not rsPaciente.EOF
            '
            Set oCommand = Nothing
            Set oParameter = Nothing
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = mo_Conexion
                .CommandTimeout = 150
                .CommandText = "MovimientosHistoriaClinicaXidPaciente"
                Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, rsPaciente!IdPaciente): .Parameters.Append oParameter
                Set rsMaxMov = .Execute
            End With
            '
            If Not (rsMaxMov.EOF And rsMaxMov.BOF) Then
                '
                Set oCommand = Nothing
                Set oParameter = Nothing
                With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = mo_Conexion
                    .CommandTimeout = 150
                    .CommandText = "tempMovimientosInsertar"
                    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, rsMaxMov!idMovimiento): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@lIdUsuario", adInteger, adParamInput, 0, lIdUsuario): .Parameters.Append oParameter
                    .Execute
                End With
                '
            End If
            rsMaxMov.Close
            
            rsPaciente.MoveNext
        Loop
    End If
    rsPaciente.Close
    lcParametro231 = lcBuscaParametro.SeleccionaFilaParametro(231)
    sSql = ""
    
    If lcParametro231 = "S" Then
        '****debb2009-version HRA
        If lnNroHistoriaClinica > 0 Then
           sSql = sSql + " and Pacientes.NroHistoriaClinica=" & lnNroHistoriaClinica
        End If
        If lIdServiciodestino > 0 Then
           sSql = sSql + " and dbo.MovimientosHistoriaClinica.IdServicioDestino = " & lIdServiciodestino
        End If
        sSql = sSql + " order by dbo.Pacientes.NroHistoriaClinica, dbo.MovimientosHistoriaClinica.fechaMovimiento DESC"
     Else
        '****debb2009-version HBT
        If lnNroHistoriaClinica > 0 Then
           sSql = sSql + " and Pacientes.NroHistoriaClinica=" & lnNroHistoriaClinica
        End If
        sSql = sSql + " order by dbo.Pacientes.NroHistoriaClinica"
     End If
    Set oCommand = Nothing
    Set oParameter = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "MovimientosHistoriaClinicaParaDevolver2"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdArchivoClinico", adVarChar, adParamInput, 20, lnIdArchivoClinico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcParametro231", adVarChar, adParamInput, 1, lcParametro231): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdUsuario", adVarChar, adParamInput, 20, lIdUsuario): .Parameters.Append oParameter
        Set rsMaxMov = .Execute
    End With
    Set MovimientosHistoriasClinicasParaDevolverPorNroHistoria = rsMaxMov
    ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function


Function MovimientosHistoriasClinicasParaDevolverPorServicio(lIdServiciodestino As Long, sFechaMovimiento As String, _
                            lIdUsuario As Long, lnNroHistoriaClinica As Long, lnIdArchivoClinico As Long, _
                            lnIdPacienteSolo As Long, FechaTipoFiltro As Integer, operador As String, _
                            fechaInicio As String, fechaFin As String, fechaFinRequerida As Boolean) As Recordset
 On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lIdServicio As Long
Dim oRecordset As New Recordset
Dim rsPaciente As New Recordset
Dim rsMaxMov As New Recordset
Dim sSql As String
Dim lcParametro231 As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "tempMovimientosEliminarXidUsuario"
        Set oParameter = .CreateParameter("@lIdUsuario", adInteger, adParamInput, 0, lIdUsuario): .Parameters.Append oParameter
        .Execute
    End With
    Dim strWhere As String
    If lnIdPacienteSolo = 0 Then
        strWhere = ""
    End If
    '
    'rsPaciente.Open sSql, mo_Conexion
    Set oCommand = Nothing
    Set oParameter = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "MovimientosHistoriaClinicaPorFiltro"
        Set oParameter = .CreateParameter("@sFechaMovimiento", adDBTimeStamp, adParamInput, 0, CDate(sFechaMovimiento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdServiciodestino", adInteger, adParamInput, 0, lIdServiciodestino): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdPacienteSolo", adInteger, adParamInput, 0, lnIdPacienteSolo): .Parameters.Append oParameter
        Set rsPaciente = .Execute
    End With
    '
    If rsPaciente.RecordCount > 0 Then
        Do While Not rsPaciente.EOF
            '
            Set oCommand = Nothing
            Set oParameter = Nothing
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = mo_Conexion
                .CommandTimeout = 150
                .CommandText = "MovimientosHistoriaClinicaXidPaciente"
                Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, rsPaciente!IdPaciente): .Parameters.Append oParameter
                Set rsMaxMov = .Execute
            End With
            '
            If Not (rsMaxMov.EOF And rsMaxMov.BOF) Then
                '
                Set oCommand = Nothing
                Set oParameter = Nothing
                With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = mo_Conexion
                    .CommandTimeout = 150
                    .CommandText = "tempMovimientosInsertar"
                    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, rsMaxMov!idMovimiento): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@lIdUsuario", adInteger, adParamInput, 0, lIdUsuario): .Parameters.Append oParameter
                    .Execute
                End With
                '
            End If
            rsMaxMov.Close
            
            rsPaciente.MoveNext
        Loop
    End If
    rsPaciente.Close
    lcParametro231 = lcBuscaParametro.SeleccionaFilaParametro(231)
    sSql = ""
    
'   FechaTipoFiltro As Integer, operador As String, _
'   fechaInicio As String, fechaFin As String
    '<<<< YAMILL
    Dim lcTipoFiltroFecha As String
    Select Case FechaTipoFiltro
        Case 0
            lcTipoFiltroFecha = "DATEADD(dd, 0,DATEDIFF(dd,0,dbo.HistoriasSolicitadas.FechaSolicitud))"
        Case 1
            lcTipoFiltroFecha = "DATEADD(dd, 0,DATEDIFF(dd,0,dbo.HistoriasSolicitadas.FechaRequerida))"
        Case 2
            lcTipoFiltroFecha = "DATEADD(dd, 0,DATEDIFF(dd,0,dbo.MovimientosHistoriaClinica.FechaMovimiento))"
    End Select
    
    If operador <> "" Then
        If fechaFinRequerida = False Then
            sSql = sSql + " and  ( " & lcTipoFiltroFecha & " " & operador & " CONVERT(DATETIME, '" & Format(fechaInicio, "yyyy-mm-dd") & " 00:00:00', 102))"
        Else
            'sSql = sSql + " and  " & lcTipoFiltroFecha & " Between (CONVERT(DATETIME,'" & fechaInicio & "',103)) and (CONVERT(DATETIME,'" & fechaFin & "',103))"
            
            sSql = sSql + " and  (" & lcTipoFiltroFecha & " >= CONVERT(DATETIME, '" & Format(fechaInicio, "yyyy-mm-dd") & " 00:00:00', 102))"
            sSql = sSql + " and  (" & lcTipoFiltroFecha & " <= CONVERT(DATETIME, '" & Format(fechaFin, "yyyy-mm-dd") & " 23:59:59', 102))"
        End If
    End If
    
'    If lnNroHistoriaClinica > 0 Then
'       sSql = sSql + " and Pacientes.NroHistoriaClinica = " & lnNroHistoriaClinica
'    End If
 '   sSql = sSql + " order by Pacientes.NroHistoriaClinica,Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
    '<<<<<<
   
   
    If lcParametro231 = "S" Then
        '****debb2009-version HRA
        If lnNroHistoriaClinica > 0 Then
           sSql = sSql + " and Pacientes.NroHistoriaClinica=" & lnNroHistoriaClinica
        End If
        If lIdServiciodestino > 0 Then
           sSql = sSql + " and dbo.MovimientosHistoriaClinica.IdServicioDestino = " & lIdServiciodestino
        End If
        sSql = sSql + " order by dbo.Pacientes.NroHistoriaClinica, dbo.MovimientosHistoriaClinica.fechaMovimiento DESC"
     Else
        '****debb2009-version HBT
        If lnNroHistoriaClinica > 0 Then
           sSql = sSql + " and Pacientes.NroHistoriaClinica=" & lnNroHistoriaClinica
        End If
        sSql = sSql + " order by dbo.Pacientes.NroHistoriaClinica"
     End If
     
     
    Set oCommand = Nothing
    Set oParameter = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "MovimientosHistoriaClinicaParaDevolver3"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdArchivoClinico", adVarChar, adParamInput, 20, lnIdArchivoClinico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcParametro231", adVarChar, adParamInput, 1, lcParametro231): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdUsuario", adVarChar, adParamInput, 20, lIdUsuario): .Parameters.Append oParameter
        Set rsMaxMov = .Execute
    End With
    Set MovimientosHistoriasClinicasParaDevolverPorServicio = rsMaxMov
    ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function


Function MovimientosHistoriasClinicasParaDevolver2(lIdServiciodestino As Long, sFechaMovimiento As String, _
                            lIdUsuario As Long, lnNroHistoriaClinica As Long, lnIdArchivoClinico As Long, _
                            lnIdPacienteSolo As Long) As Recordset
 On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lIdServicio As Long
Dim oRecordset As New Recordset
Dim rsPaciente As New Recordset
Dim rsMaxMov As New Recordset
Dim sSql As String
Dim lcParametro231 As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "tempMovimientosEliminarXidUsuario"
        Set oParameter = .CreateParameter("@lIdUsuario", adInteger, adParamInput, 0, lIdUsuario): .Parameters.Append oParameter
        .Execute
    End With
    '
    'rsPaciente.Open sSql, mo_Conexion
    Set oCommand = Nothing
    Set oParameter = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "MovimientosHistoriaClinicaPorFiltro"
        Set oParameter = .CreateParameter("@sFechaMovimiento", adDBTimeStamp, adParamInput, 0, CDate(sFechaMovimiento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdServiciodestino", adInteger, adParamInput, 0, lIdServiciodestino): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdPacienteSolo", adInteger, adParamInput, 0, lnIdPacienteSolo): .Parameters.Append oParameter
        Set rsPaciente = .Execute
    End With
    
    '
    If rsPaciente.RecordCount > 0 Then
        Do While Not rsPaciente.EOF
            '
            Set oCommand = Nothing
            Set oParameter = Nothing
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = mo_Conexion
                .CommandTimeout = 150
                .CommandText = "MovimientosHistoriaClinicaXidPaciente"
                Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, rsPaciente!IdPaciente): .Parameters.Append oParameter
                Set rsMaxMov = .Execute
            End With
            '
            If Not (rsMaxMov.EOF And rsMaxMov.BOF) Then
                '
                Set oCommand = Nothing
                Set oParameter = Nothing
                With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = mo_Conexion
                    .CommandTimeout = 150
                    .CommandText = "tempMovimientosInsertar"
                    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, rsMaxMov!idMovimiento): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@lIdUsuario", adInteger, adParamInput, 0, lIdUsuario): .Parameters.Append oParameter
                    .Execute
                End With
                '
            End If
            rsMaxMov.Close
            
            rsPaciente.MoveNext
        Loop
    End If
    rsPaciente.Close
    lcParametro231 = lcBuscaParametro.SeleccionaFilaParametro(231)
    sSql = ""
    If lcParametro231 = "S" Then
        '****debb2009-version HRA
        If lnNroHistoriaClinica > 0 Then
           sSql = sSql + " and Pacientes.NroHistoriaClinica=" & lnNroHistoriaClinica
        End If
        If lIdServiciodestino > 0 Then
           sSql = sSql + " and dbo.MovimientosHistoriaClinica.IdServicioDestino = " & lIdServiciodestino
        End If
        sSql = sSql + " order by dbo.Pacientes.NroHistoriaClinica, dbo.MovimientosHistoriaClinica.fechaMovimiento DESC"
     Else
        '****debb2009-version HBT
        If lnNroHistoriaClinica > 0 Then
           sSql = sSql + " and Pacientes.NroHistoriaClinica=" & lnNroHistoriaClinica
        End If
        sSql = sSql + " order by dbo.Pacientes.NroHistoriaClinica"
     End If
    Set oCommand = Nothing
    Set oParameter = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "MovimientosHistoriaClinicaParaDevolver2"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdArchivoClinico", adVarChar, adParamInput, 20, lnIdArchivoClinico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcParametro231", adVarChar, adParamInput, 1, lcParametro231): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdUsuario", adVarChar, adParamInput, 20, lIdUsuario): .Parameters.Append oParameter
        Set rsMaxMov = .Execute
    End With
    Set MovimientosHistoriasClinicasParaDevolver2 = rsMaxMov
    ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function


Function MovimientosHistoriasClinicasParaDevolver(lIdServiciodestino As Long, sFechaMovimiento As String, lIdUsuario As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lIdServicio As Long
Dim oRecordset As New Recordset

   Set MovimientosHistoriasClinicasParaDevolver = Nothing
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "MovimientosHistoriasClinicasParaDevolver"
            Set oParameter = .CreateParameter("@IdServiciodestino", adInteger, adParamInput, 0, lIdServiciodestino): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaMovimiento", adChar, adParamInput, 10, sFechaMovimiento): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdUsuario", adInteger, adParamInput, 0, lIdUsuario): .Parameters.Append oParameter
            Set oRecordset = .Execute
   End With
 
    Set oRecordset.ActiveConnection = Nothing
    Set MovimientosHistoriasClinicasParaDevolver = oRecordset
    
    ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function MovimientosHistoriaClinicaEliminarMovIdPaciente(ByVal oTabla As DOMovimientoHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   MovimientosHistoriaClinicaEliminarMovIdPaciente = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "MovimientosHistoriaClinicaEliminarMovIdPaciente"
       Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, IIf(oTabla.idMovimiento = 0, Null, oTabla.idMovimiento)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente)
       .Parameters.Append oParameter
       .Execute
   End With
 
   MovimientosHistoriaClinicaEliminarMovIdPaciente = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'SCCQ 26/08/2020 Cambio27 Inicio
Function SeleccionarHCSinDevolver(horas As Integer) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarHCSinDevolver = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "usp_HCSinDevolver"
     Set oParameter = .CreateParameter("@horas", adInteger, adParamInput, 0, horas): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarHCSinDevolver = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'SCCQ 26/08/2020 Cambio27 Fin
