VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MovimientoFormatosHC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla MovimientosFormatosHistoriaClinica
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
'   Descripción:    Inserta un registro a la tabla MovimientosFormatosHistoriaClinica
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOMovimientoFormatoHC) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "MovimientosFormatosHistoriaClinicaAgregar"
   Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaMovimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaMovimiento = 0, Null, oTabla.FechaMovimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdMotivo", adInteger, adParamInput, 0, IIf(oTabla.IdMotivo = 0, Null, oTabla.IdMotivo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 100, IIf(oTabla.Observacion = "", Null, oTabla.Observacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicioOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdServicioOrigen = 0, Null, oTabla.IdServicioOrigen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicioDestino", adInteger, adParamInput, 0, IIf(oTabla.IdServicioDestino = 0, Null, oTabla.IdServicioDestino))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroFolios", adInteger, adParamInput, 0, IIf(oTabla.NroFolios = 0, Null, oTabla.NroFolios))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEmpleadoArchivo", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoArchivo = 0, Null, oTabla.IdEmpleadoArchivo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEmpleadoTransporte", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoTransporte = 0, Null, oTabla.IdEmpleadoTransporte))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEmpleadoRecepcion", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoRecepcion = 0, Null, oTabla.IdEmpleadoRecepcion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdGrupoMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdGrupoMovimiento = 0, Null, oTabla.IdGrupoMovimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdMovimiento = .Parameters("@IdMovimiento")
End With
 
Insertar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla MovimientosFormatosHistoriaClinica
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOMovimientoFormatoHC) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "MovimientosFormatosHistoriaClinicaModificar"
   Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMovimiento = 0, Null, oTabla.IdMovimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaMovimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaMovimiento = 0, Null, oTabla.FechaMovimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdMotivo", adInteger, adParamInput, 0, IIf(oTabla.IdMotivo = 0, Null, oTabla.IdMotivo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 100, IIf(oTabla.Observacion = "", Null, oTabla.Observacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicioOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdServicioOrigen = 0, Null, oTabla.IdServicioOrigen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicioDestino", adInteger, adParamInput, 0, IIf(oTabla.IdServicioDestino = 0, Null, oTabla.IdServicioDestino))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroFolios", adInteger, adParamInput, 0, IIf(oTabla.NroFolios = 0, Null, oTabla.NroFolios))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEmpleadoArchivo", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoArchivo = 0, Null, oTabla.IdEmpleadoArchivo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEmpleadoTransporte", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoTransporte = 0, Null, oTabla.IdEmpleadoTransporte))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEmpleadoRecepcion", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoRecepcion = 0, Null, oTabla.IdEmpleadoRecepcion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdGrupoMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdGrupoMovimiento = 0, Null, oTabla.IdGrupoMovimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Modificar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla MovimientosFormatosHistoriaClinica
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOMovimientoFormatoHC) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "MovimientosFormatosHistoriaClinicaEliminar"
   Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, oTabla.IdMovimiento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Eliminar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla MovimientosFormatosHistoriaClinica
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOMovimientoFormatoHC) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "MovimientosFormatosHistoriaClinicaSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, oTabla.IdMovimiento)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdMovimiento = IIf(IsNull(oRecordset!IdMovimiento), 0, oRecordset!IdMovimiento)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.FechaMovimiento = IIf(IsNull(oRecordset!FechaMovimiento), 0, oRecordset!FechaMovimiento)
       oTabla.IdMotivo = IIf(IsNull(oRecordset!IdMotivo), 0, oRecordset!IdMotivo)
       oTabla.Observacion = IIf(IsNull(oRecordset!Observacion), "", oRecordset!Observacion)
       oTabla.IdServicioOrigen = IIf(IsNull(oRecordset!IdServicioOrigen), 0, oRecordset!IdServicioOrigen)
       oTabla.IdServicioDestino = IIf(IsNull(oRecordset!IdServicioDestino), 0, oRecordset!IdServicioDestino)
       oTabla.NroFolios = IIf(IsNull(oRecordset!NroFolios), 0, oRecordset!NroFolios)
       oTabla.IdEmpleadoArchivo = IIf(IsNull(oRecordset!IdEmpleadoArchivo), 0, oRecordset!IdEmpleadoArchivo)
       oTabla.IdEmpleadoTransporte = IIf(IsNull(oRecordset!IdEmpleadoTransporte), 0, oRecordset!IdEmpleadoTransporte)
       oTabla.IdEmpleadoRecepcion = IIf(IsNull(oRecordset!IdEmpleadoRecepcion), 0, oRecordset!IdEmpleadoRecepcion)
       oTabla.IdGrupoMovimiento = IIf(IsNull(oRecordset!IdGrupoMovimiento), 0, oRecordset!IdGrupoMovimiento)
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOMovimientoFormatoHC)
   oTabla.IdMovimiento = 0
   oTabla.IdPaciente = 0
   oTabla.FechaMovimiento = Now
   oTabla.IdMotivo = 0
   oTabla.Observacion = ""
   oTabla.IdServicioOrigen = 0
   oTabla.IdServicioDestino = 0
   oTabla.NroFolios = 0
   oTabla.IdEmpleadoArchivo = 0
   oTabla.IdEmpleadoTransporte = 0
   oTabla.IdEmpleadoRecepcion = 0
   oTabla.IdGrupoMovimiento = 0
   oTabla.idAtencion = 0
End Sub

Function AgregarVarios(oMovimiento As DOMovimientoFormatoHC, rsMovimientos As Recordset) As Boolean
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
            If Not Insertar(oMovimiento) Then
                Exit Function
            End If
            rsMovimientos!IdMovimientoHistoria = oMovimiento.IdMovimiento
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

Function ModificarVarios(oMovimiento As DOMovimientoFormatoHC, rsMovimientos As Recordset) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oDOMovimiento As New DOMovimientoFormatoHC
    ModificarVarios = False
    rsMovimientos.MoveFirst
    Do While Not rsMovimientos.EOF
        oMovimiento.IdMovimiento = rsMovimientos!IdMovimientoHistoria
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
Dim oDOMovimiento As New DOMovimientoFormatoHC
    EliminarVarios = False
        rsMovimientos.MoveFirst
        Do While Not rsMovimientos.EOF
            oDOMovimiento.IdMovimiento = rsMovimientos!IdMovimientoHistoria
            If Not Eliminar(oDOMovimiento) Then
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

Function Filtrar(ByVal oDOPaciente As DoPaciente, ByVal oDOHistoria As DOMovimientoHistoriaClinica) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
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
    If oDOHistoria.FechaMovimiento <> 0 Then
         sWhere = sWhere + " convert(char(10), MovimientosFormatosHistoriaClinica.FechaMovimiento, 103) = '" & oDOHistoria.FechaMovimiento & "' and "
    End If
    If sWhere <> "" Then
         sSql = sSql + " where " + Left(sWhere, Len(sWhere) - 4)
    End If
    sSql = sSql + " order by MovimientosFormatosHistoriaClinica.FechaMovimiento desc,dbo.MovimientosFormatosHistoriaClinica.IdMovimiento desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "MovimientosFormatosHistoriaClinicaFiltrar"
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


Function MovimientosFormatosHCPorIdGrupo(lIdGrupo As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lIdServicio As Long
Dim oRecordset As New Recordset
 
   Set MovimientosFormatosHCPorIdGrupo = Nothing
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "MovimientosFormatosHCPorIdGrupo"
            Set oParameter = .CreateParameter("@IdGrupo", adInteger, adParamInput, 0, lIdGrupo): .Parameters.Append oParameter
            Set oRecordset = .Execute
   End With
 
    Set oRecordset.ActiveConnection = Nothing
    Set MovimientosFormatosHCPorIdGrupo = oRecordset
    
    ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


