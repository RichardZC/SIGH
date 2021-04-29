VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturacionServicioFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FacturacionServicioFinanciamientos
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String

Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property


'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla FacturacionServicioFinanciamientos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoFacturacionServicioFin) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionServicioFinanciamientosAgregar"
   Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, oTabla.idFuenteFinanciamiento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadFinanciada", adInteger, adParamInput, 0, oTabla.CantidadFinanciada)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrecioFinanciado", adCurrency, adParamInput, 0, oTabla.PrecioFinanciado)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TotalFinanciado", adCurrency, adParamInput, 0, oTabla.TotalFinanciado)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaAutoriza", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaAutoriza = 0, Null, oTabla.FechaAutoriza))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAutoriza", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutoriza = 0, Null, oTabla.IdUsuarioAutoriza))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
    'SCCQ 26-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
    'HRA 10/12/2020 Cambio 46 Inicio
'   Set oParameter = .CreateParameter("@FuaCodigoPrestacion", adVarChar, adParamInput, 3, IIf(oTabla.CodPres = "", Null, oTabla.CodPres))
'   .Parameters.Append oParameter
   'HRA 10/12/2020 Cambio 46 Fin
    'SCCQ 26-04-2021 Cambio 64 Fin (Reversion Cambio 46)
   
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
'   Descripción:    modifica un registro a la tabla FacturacionServicioFinanciamientos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoFacturacionServicioFin) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionServicioFinanciamientosModificar"
   Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, oTabla.idFuenteFinanciamiento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadFinanciada", adInteger, adParamInput, 0, oTabla.CantidadFinanciada)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrecioFinanciado", adCurrency, adParamInput, 0, oTabla.PrecioFinanciado)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TotalFinanciado", adCurrency, adParamInput, 0, oTabla.TotalFinanciado)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaAutoriza", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaAutoriza = 0, Null, oTabla.FechaAutoriza))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAutoriza", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutoriza = 0, Null, oTabla.IdUsuarioAutoriza))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoFacturacion", adInteger, adParamInput, 0, oTabla.IdEstadoFacturacion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
       .Execute
    'SCCQ 29/04/2021 Cambio 64 Inicio (Reversion Cambio 46)
    'RHA 17/11/2020 Cambio 46 Inicio
'    Set oParameter = .CreateParameter("@FuaCodigoPrestacion", adVarChar, adParamInput, 3, IIf(oTabla.CodPres = "", Null, oTabla.CodPres))
'   .Parameters.Append oParameter
    'RHA 17/11/2020 Cambio 46 Fin
     'SCCQ 29/04/2021 Cambio 64 Fin (Reversion Cambio 46)
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
'   Descripción:    elimia o borra un registro a la tabla FacturacionServicioFinanciamientos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoFacturacionServicioFin) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionServicioFinanciamientosEliminar"
   Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, oTabla.IdOrden)
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
'   Descripción:    Selecciona un unico registro de la tabla FacturacionServicioFinanciamientos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoFacturacionServicioFin) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FacturacionServicioFinanciamientosSeleccionarPorId"
     Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, oTabla.IdOrden)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdOrden = IIf(IsNull(oRecordset!IdOrden), 0, oRecordset!IdOrden)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.idTipoFinanciamiento = IIf(IsNull(oRecordset!idTipoFinanciamiento), 0, oRecordset!idTipoFinanciamiento)
       oTabla.idFuenteFinanciamiento = IIf(IsNull(oRecordset!idFuenteFinanciamiento), 0, oRecordset!idFuenteFinanciamiento)
       oTabla.CantidadFinanciada = IIf(IsNull(oRecordset!CantidadFinanciada), 0, oRecordset!CantidadFinanciada)
       oTabla.PrecioFinanciado = oRecordset!PrecioFinanciado
       oTabla.TotalFinanciado = oRecordset!TotalFinanciado
       oTabla.FechaAutoriza = IIf(IsNull(oRecordset!FechaAutoriza), 0, oRecordset!FechaAutoriza)
       oTabla.IdUsuarioAutoriza = IIf(IsNull(oRecordset!IdUsuarioAutoriza), 0, oRecordset!IdUsuarioAutoriza)
       oTabla.IdEstadoFacturacion = IIf(IsNull(oRecordset!IdEstadoFacturacion), 0, oRecordset!IdEstadoFacturacion)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoFacturacionServicioFin)
   oTabla.IdOrden = 0
   oTabla.idProducto = 0
   oTabla.idTipoFinanciamiento = 0
   oTabla.idFuenteFinanciamiento = 0
   oTabla.CantidadFinanciada = 0
   oTabla.PrecioFinanciado = 0
   oTabla.TotalFinanciado = 0
   oTabla.FechaAutoriza = Now
   oTabla.IdUsuarioAutoriza = 0
   oTabla.IdEstadoFacturacion = 0
End Sub
Function SeleccionarPorIdOrdenIdProducto(lnIdOrden, lnIdProducto As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set SeleccionarPorIdOrdenIdProducto = Nothing
    ms_MensajeError = ""
    oConexion.Open SIGHENTIDADES.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioFinanciamientosSeleccionarPorIdOrdenIdProducto"
        Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorIdOrdenIdProducto = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function EliminarPorIdTipoFinanciamiento(ByVal oTabla As DoFacturacionServicioFin) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
EliminarPorIdTipoFinanciamiento = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionServicioFinanciamientosEliminarPorIdTipoFinanciamiento"
   Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, oTabla.IdOrden)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, oTabla.idTipoFinanciamiento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
EliminarPorIdTipoFinanciamiento = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function ActualizaIdEstadoFacturacion(lnIdOrden As Long, lnNuevoIdEstadoFacturacion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioFinanciamientosActualizaIdEstadoFacturacion"
        Set oParameter = .CreateParameter("@lnNuevoIdEstadoFacturacion", adInteger, adParamInput, 0, lnNuevoIdEstadoFacturacion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        .Execute
   End With
   ActualizaIdEstadoFacturacion = True
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'*****lcSoloActualizarIdEstadoFacturacion="/15/1"
'*****tambien actualiza FactOrdenServicio.IdEstadoFacturacion, solo para los que tienen algun tipo de seguros
Function ActualizaIdEstadoFacturacionPorCuenta(lnIdCuentaAtencion As Long, lcSoloActualizarIdEstadoFacturacion As String, lnNuevoIdEstadoFacturacion As Long) As Boolean
    ActualizaIdEstadoFacturacionPorCuenta = False
    Dim oRsActualiza As New Recordset
    Dim oRsActualizaFactOrdenServicio As New Recordset
    Dim oDoFacturacionServicioFin As New DoFacturacionServicioFin, oFacturacionServicioFin As New FacturacionServicioFin
    Dim oDoFactOrdenServ As New DoFactOrdenServ, oFactOrdenServicio As New FactOrdenServicio
    Dim lcSql As String
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioXidCuentaAtencion"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRsActualizaFactOrdenServicio = .Execute
        Set oRsActualizaFactOrdenServicio.ActiveConnection = Nothing
    End With
    If oRsActualizaFactOrdenServicio.RecordCount > 0 Then
       'mo_Conexion.BeginTrans
       Set oFacturacionServicioFin.Conexion = mo_Conexion
       Set oFactOrdenServicio.Conexion = mo_Conexion
       oRsActualizaFactOrdenServicio.MoveFirst
       Do While Not oRsActualizaFactOrdenServicio.EOF
          If InStr(lcSoloActualizarIdEstadoFacturacion, Trim(Str(oRsActualizaFactOrdenServicio.Fields!IdEstadoFacturacion))) > 0 Then
                If oFacturacionServicioFin.ActualizaIdEstadoFacturacionXidOrden(oRsActualizaFactOrdenServicio.Fields!IdOrden, lnNuevoIdEstadoFacturacion) = True Then
                      oDoFactOrdenServ.IdOrden = oRsActualizaFactOrdenServicio.Fields!IdOrden
                      If oFactOrdenServicio.SeleccionarPorId(oDoFactOrdenServ) = True Then
                         oDoFactOrdenServ.IdEstadoFacturacion = lnNuevoIdEstadoFacturacion
                         If oFactOrdenServicio.Modificar(oDoFactOrdenServ) = True Then
                         End If
                      End If
                End If
          End If
          oRsActualizaFactOrdenServicio.MoveNext
       Loop
       'mo_Conexion.CommitTrans
    End If
    oRsActualizaFactOrdenServicio.Close
    ms_MensajeError = ""
    Set oCommand = Nothing
    
    Set oRsActualiza = Nothing
    Set oRsActualizaFactOrdenServicio = Nothing
    Set oDoFacturacionServicioFin = Nothing
    Set oFacturacionServicioFin = Nothing
    Set oDoFactOrdenServ = Nothing
    Set oFactOrdenServicio = Nothing
    
    ActualizaIdEstadoFacturacionPorCuenta = True
    Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function ActualizaIdEstadoFacturacionXidOrden(lnIdOrden As Long, lnIdEstadoFacturacion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioFinanciamientosActualizaIdEstadoFact"
        Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idEstadoFacturacion", adInteger, adParamInput, 0, lnIdEstadoFacturacion): .Parameters.Append oParameter
        .Execute
   End With
   ActualizaIdEstadoFacturacionXidOrden = True
   Set oCommand = Nothing
   Set oRecordset = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description

End Function



