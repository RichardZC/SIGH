VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturacionBienesFinanciam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FacturacionBienesFinanciamientos
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
'   Descripción:    Inserta un registro a la tabla FacturacionBienesFinanciamientos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoFacturacionBienesFinanciam) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionBienesFinanciamientosAgregar"
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
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
   .Execute
   Set .ActiveConnection = Nothing
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
'   Descripción:    modifica un registro a la tabla FacturacionBienesFinanciamientos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoFacturacionBienesFinanciam) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionBienesFinanciamientosModificar"
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
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
'   Descripción:    elimia o borra un registro a la tabla FacturacionBienesFinanciamientos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoFacturacionBienesFinanciam) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionBienesFinanciamientosEliminar"
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, oTabla.movNumero)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   Set .ActiveConnection = Nothing
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
'   Descripción:    Selecciona un unico registro de la tabla FacturacionBienesFinanciamientos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoFacturacionBienesFinanciam) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FacturacionBienesFinanciamientosSeleccionarPorId"
     Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, oTabla.movNumero)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.movNumero = IIf(IsNull(oRecordset!movNumero), "", oRecordset!movNumero)
       oTabla.movTipo = IIf(IsNull(oRecordset!movTipo), "", oRecordset!movTipo)
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
Sub SetDefaults(ByVal oTabla As DoFacturacionBienesFinanciam)
   oTabla.movNumero = ""
   oTabla.movTipo = ""
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

Function SeleccionarPorIdProducto(LcMovNumero As String, LcMovTipo As String, lnIdProducto As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set SeleccionarPorIdProducto = Nothing
    ms_MensajeError = ""
    oConexion.Open SIGHENTIDADES.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesFinanciamientosSeleccionarPorIdProducto"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, LcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, LcMovTipo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorIdProducto = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function EliminarPorTipoFinanciamiento(ByVal oTabla As DoFacturacionBienesFinanciam) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
EliminarPorTipoFinanciamiento = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionBienesFinanciamientosEliminarPorIdTipoFinanciamiento"
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, oTabla.movNumero)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, oTabla.idTipoFinanciamiento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   Set .ActiveConnection = Nothing
End With
 
EliminarPorTipoFinanciamiento = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


Function ActualizaIdEstadoFacturacion(LcMovNumero As String, LcMovTipo As String, lnNuevoIdEstadoFacturacion As Long) As Boolean

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesFinanciamientosActualizaIdEstadoFacturacion"
        Set oParameter = .CreateParameter("@LcMovTipo", adVarChar, adParamInput, 1, LcMovTipo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@LcMovNumero", adVarChar, adParamInput, 9, LcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnNuevoIdEstadoFacturacion", adInteger, adParamInput, 0, lnNuevoIdEstadoFacturacion): .Parameters.Append oParameter
        .Execute
   End With
   Set oRecordset = Nothing
   Set oCommand = Nothing
   ActualizaIdEstadoFacturacion = True
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ActualizaIdEstadoFacturacionPorCuenta(lnIdCuentaAtencion As Long, lnNuevoIdEstadoFacturacion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRsTmp1 As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRsActualiza As New Recordset
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoVentasPorCuenta"
        Set oParameter = .CreateParameter("@lnIdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          Set oCommand = Nothing
          Set oParameter = Nothing
          If oRsActualiza.State = 1 Then oRsActualiza.Close
          With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = mo_Conexion
             .CommandTimeout = 150
             .CommandText = "FacturacionBienesFinanciamientosActualizaIdEstadoFacturacionPorCuenta"
             Set oParameter = .CreateParameter("@lnNuevoIdEstadoFacturacion", adInteger, adParamInput, 0, lnNuevoIdEstadoFacturacion): .Parameters.Append oParameter
             Set oParameter = .CreateParameter("@movTipo", adVarChar, adParamInput, 1, oRsTmp1.Fields!movTipo): .Parameters.Append oParameter
             Set oParameter = .CreateParameter("@movNumero", adVarChar, adParamInput, 9, oRsTmp1.Fields!movNumero): .Parameters.Append oParameter
             .Execute
          End With
          oRsTmp1.MoveNext
       Loop
   End If
   Set oCommand = Nothing
   Set oRsActualiza = Nothing
   Set oRsTmp1 = Nothing
   ActualizaIdEstadoFacturacionPorCuenta = True
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function







