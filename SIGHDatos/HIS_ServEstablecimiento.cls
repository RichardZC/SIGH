VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HIS_ServEstablecimiento"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla HIS_ServEstablecimiento
'        Programado por: Cachay F
'        Fecha: Noviembre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ml_IdHisServEstablecimiento As Long
Dim ml_IdEstablecimiento As Long
Dim ml_IdServicio As Long

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
'   Descripción:    Inserta un registro a la tabla HIS_ServEstablecimiento
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOHIS_ServEstablecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_ServEstablecimientoAgregar"
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
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
'   Descripción:    modifica un registro a la tabla HIS_ServEstablecimiento
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOHIS_ServEstablecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_ServEstablecimientoModificar"
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio))
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
'   Descripción:    elimia o borra un registro a la tabla HIS_ServEstablecimiento
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOHIS_ServEstablecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_ServEstablecimientoEliminar"
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio))
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

Sub ActualizaTabla(ByRef oTabla As DOHIS_ServEstablecimiento, oRecordset As Recordset)
       oTabla.IdEstablecimiento = IIf(IsNull(oRecordset!IdEstablecimiento), 0, oRecordset!IdEstablecimiento)
       oTabla.IdServicio = IIf(IsNull(oRecordset!IdServicio), 0, oRecordset!IdServicio)
End Sub


'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla HIS_ServEstablecimiento
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOHIS_ServEstablecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "HIS_ServEstablecimientoSeleccionarPorId"
    Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio))
    .Parameters.Append oParameter
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       ActualizaTabla oTabla, oRecordset
'       oTabla.IdHisServEstablecimiento = IIf(IsNull(oRecordset!IdHisServEstablecimiento), 0, oRecordset!IdHisServEstablecimiento)
'       oTabla.IdEstablecimiento = IIf(IsNull(oRecordset!IdEstablecimiento), 0, oRecordset!IdEstablecimiento)
'       oTabla.IdServicio = IIf(IsNull(oRecordset!IdServicio), 0, oRecordset!IdServicio)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOHIS_ServEstablecimiento)
   oTabla.IdEstablecimiento = 0
   oTabla.IdServicio = 0
End Sub

Public Function ObtenerListaEstablecimientosMR() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
Dim mo_TablaHISPaciente As New DOHIS_Paciente
   
     With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = mo_Conexion
         .CommandTimeout = 150
         .CommandText = "HIS_SeleccionarEstablecimientosMR"
         Set oRecordset = .Execute
         Set oRecordset.ActiveConnection = Nothing
    End With
   Set ObtenerListaEstablecimientosMR = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Public Function ObtenerListaServiciosCentroMR() As Recordset
'On Error GoTo ManejadorDeError
'Dim oRcs_ListarServiciosDeEstablecimientosMR As New ADODB.Recordset
'Dim oCommand As New ADODB.Command
'Dim sSql As String
'
'sSql = "select SE.IdHisServEstablecimiento, SE.IdEstablecimiento, SE.IdServicio, S.nombre, 0 IdEstado"
'sSql = sSql & " from HIS_ServEstablecimiento SE inner join Servicios S"
'sSql = sSql & " on SE.IdServicio = S.IdServicio"
'
'With oCommand
'    .CommandType = adCmdText
'    Set .ActiveConnection = mo_Conexion
'    .CommandText = sSql
'    Set oRcs_ListarServiciosDeEstablecimientosMR = .Execute
'    Set oRcs_ListarServiciosDeEstablecimientosMR.ActiveConnection = Nothing
'End With
'
'Set ObtenerListaServiciosCentroMR = oRcs_ListarServiciosDeEstablecimientosMR
'On Error GoTo 0
'Exit Function
'
'ManejadorDeError:
'   ms_MensajeError = Err.Number & " " + Err.Description
Stop
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
Dim mo_TablaHISPaciente As New DOHIS_Paciente
    sSql = ""
    sWhere = ""
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_ServEstablecimientoListaServiciosCentroMR"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerListaServiciosCentroMR = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



Function ListaServiciosPorEstablecimiento(ml_IdEstablecimientoActual As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
Dim mo_TablaHISPaciente As New DOHIS_Paciente
    sSql = ""
    sWhere = ""
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_ServEstablecimientoPorEstablecimiento"
        Set oParameter = .CreateParameter("@ml_IdEstablecimientoActual", adInteger, adParamInput, 0, ml_IdEstablecimientoActual): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ListaServiciosPorEstablecimiento = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Public Function ObtenerDatosIdServicioPorIdServEstablecimiento(IdServEstablec As Long) As Long
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
Dim mo_TablaHISPaciente As New DOHIS_Paciente
    sSql = ""
    sWhere = ""
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_ServEstablecimientoIdServEstablecimiento"
        Set oParameter = .CreateParameter("@IdServEstablec", adInteger, adParamInput, 0, IdServEstablec): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   ObtenerDatosIdServicioPorIdServEstablecimiento = 0
   If oRecordset.RecordCount > 0 Then
      ObtenerDatosIdServicioPorIdServEstablecimiento = CLng(oRecordset!IdServicio)
   End If
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Public Function ObtenerDatosIdServEstablecimiento(oTabla As DOHIS_ServEstablecimiento) As DOHIS_ServEstablecimiento
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
Dim mo_TablaHISPaciente As New DOHIS_Paciente
    sSql = ""
    sWhere = ""
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_ServEstablecimientoIdEstablecimientoIdServicio"
        Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, oTabla.IdEstablecimiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, oTabla.IdServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   
   If oRecordset.RecordCount > 0 Then
      ActualizaTabla oTabla, oRecordset
      Set ObtenerDatosIdServEstablecimiento = oTabla
   Else
      Set ObtenerDatosIdServEstablecimiento = Nothing
   End If
   
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Public Function ListarEspecialidadesPorEstablecimiento(IdEstablecimiento As Long) As Recordset
'On Error GoTo ListarEspecialidadesPorEstablecimiento_Error
'Dim oRcs_ListarEspecialidadesPorEstablecimiento As New ADODB.Recordset
'Dim oCommand As New ADODB.Command
'Dim sSql As String
'
'sSql = "SELECT IdEspecialidad, Nombre FROM Especialidades"
'sSql = sSql & " WHERE IdEspecialidad IN(select DISTINCT IdEspecialidad from servicios"
'sSql = sSql & " where IdServicio IN (SELECT DISTINCT IdServicio"
'sSql = sSql & " From HIS_ServEstablecimiento"
'sSql = sSql & " WHERE IdEstablecimiento =" & IdEstablecimiento & "))"
'
'With oCommand
'    .CommandType = adCmdText
'    Set .ActiveConnection = mo_Conexion
'    .CommandText = sSql
'    Set oRcs_ListarEspecialidadesPorEstablecimiento = .Execute
'    Set oRcs_ListarEspecialidadesPorEstablecimiento.ActiveConnection = Nothing
'End With
'
'Set ListarEspecialidadesPorEstablecimiento = oRcs_ListarEspecialidadesPorEstablecimiento
'On Error GoTo 0
'Exit Function
'
'ListarEspecialidadesPorEstablecimiento_Error:
'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListarEspecialidadesDeEstablecimientosMR of Módulo de clase HIS_ServiciosPorEstablec"
Stop
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
Dim mo_TablaHISPaciente As New DOHIS_Paciente
    sSql = ""
    sWhere = ""
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_ServEstablecimientoXidEstablecimiento"
        Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IdEstablecimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ListarEspecialidadesPorEstablecimiento = oRecordset
   
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Public Function ListarEspecialidadesEstablecimientosExternos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
Dim mo_TablaHISPaciente As New DOHIS_Paciente
    sSql = ""
    sWhere = ""
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "EspecialidadesEstablecimientosExternos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ListarEspecialidadesEstablecimientosExternos = oRecordset
   
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

