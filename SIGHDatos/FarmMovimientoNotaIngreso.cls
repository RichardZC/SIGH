VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FarmMovimientoNotaIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FarmMovimientoNotaIngreso
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
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
'   Descripción:    Inserta un registro a la tabla farmMovimientoNotaIngreso
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOfarmMovimientoNotaIngreso) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmMovimientoNotaIngresoAgregar"
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoFechaRecepcion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.DocumentoFechaRecepcion = 0, Null, oTabla.DocumentoFechaRecepcion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OrigenIdTipo", adInteger, adParamInput, 0, IIf(oTabla.OrigenIdTipo = 0, Null, oTabla.OrigenIdTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OrigenNumero", adVarChar, adParamInput, 20, IIf(oTabla.OrigenNumero = "", Null, oTabla.OrigenNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OrigenFecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.OrigenFecha = 0, Null, oTabla.OrigenFecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProveedor", adInteger, adParamInput, 0, IIf(oTabla.idProveedor = 0, Null, oTabla.idProveedor))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoCompra", adInteger, adParamInput, 0, IIf(oTabla.idTipoCompra = 0, Null, oTabla.idTipoCompra))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoProceso", adInteger, adParamInput, 0, IIf(oTabla.idTipoProceso = 0, Null, oTabla.idTipoProceso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NumeroProceso", adVarChar, adParamInput, 50, IIf(oTabla.NumeroProceso = "", Null, oTabla.NumeroProceso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaModificacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaModificacion = 0, Null, oTabla.FechaModificacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuarioModifica", adInteger, adParamInput, 0, IIf(oTabla.idUsuarioModifica = 0, Null, oTabla.idUsuarioModifica))
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
'   Descripción:    modifica un registro a la tabla farmMovimientoNotaIngreso
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOfarmMovimientoNotaIngreso) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmMovimientoNotaIngresoModificar"
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoFechaRecepcion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.DocumentoFechaRecepcion = 0, Null, oTabla.DocumentoFechaRecepcion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OrigenIdTipo", adInteger, adParamInput, 0, IIf(oTabla.OrigenIdTipo = 0, Null, oTabla.OrigenIdTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OrigenNumero", adVarChar, adParamInput, 20, IIf(oTabla.OrigenNumero = "", Null, oTabla.OrigenNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OrigenFecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.OrigenFecha = 0, Null, oTabla.OrigenFecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProveedor", adInteger, adParamInput, 0, IIf(oTabla.idProveedor = 0, Null, oTabla.idProveedor))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoCompra", adInteger, adParamInput, 0, IIf(oTabla.idTipoCompra = 0, Null, oTabla.idTipoCompra))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoProceso", adInteger, adParamInput, 0, IIf(oTabla.idTipoProceso = 0, Null, oTabla.idTipoProceso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NumeroProceso", adVarChar, adParamInput, 50, IIf(oTabla.NumeroProceso = "", Null, oTabla.NumeroProceso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaModificacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaModificacion = 0, Null, oTabla.FechaModificacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuarioModifica", adInteger, adParamInput, 0, IIf(oTabla.idUsuarioModifica = 0, Null, oTabla.idUsuarioModifica))
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
'   Descripción:    elimia o borra un registro a la tabla farmMovimientoNotaIngreso
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOfarmMovimientoNotaIngreso) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmMovimientoNotaIngresoEliminar"
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, oTabla.movNumero)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, oTabla.movTipo)
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
'   Descripción:    Selecciona un unico registro de la tabla farmMovimientoNotaIngreso
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOfarmMovimientoNotaIngreso) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "farmMovimientoNotaIngresoSeleccionarPorId"
     Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, oTabla.movNumero)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, oTabla.movTipo)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.movNumero = IIf(IsNull(oRecordset!movNumero), "", oRecordset!movNumero)
       oTabla.movTipo = IIf(IsNull(oRecordset!movTipo), "", oRecordset!movTipo)
       oTabla.DocumentoFechaRecepcion = IIf(IsNull(oRecordset!DocumentoFechaRecepcion), 0, oRecordset!DocumentoFechaRecepcion)
       oTabla.OrigenIdTipo = IIf(IsNull(oRecordset!OrigenIdTipo), 0, oRecordset!OrigenIdTipo)
       oTabla.OrigenNumero = IIf(IsNull(oRecordset!OrigenNumero), "", oRecordset!OrigenNumero)
       oTabla.OrigenFecha = IIf(IsNull(oRecordset!OrigenFecha), 0, oRecordset!OrigenFecha)
       oTabla.idProveedor = IIf(IsNull(oRecordset!idProveedor), 0, oRecordset!idProveedor)
       oTabla.idTipoCompra = IIf(IsNull(oRecordset!idTipoCompra), 0, oRecordset!idTipoCompra)
       oTabla.idTipoProceso = IIf(IsNull(oRecordset!idTipoProceso), 0, oRecordset!idTipoProceso)
       oTabla.NumeroProceso = IIf(IsNull(oRecordset!NumeroProceso), "", oRecordset!NumeroProceso)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.idFuenteFinanciamiento = IIf(IsNull(oRecordset!idFuenteFinanciamiento), 0, oRecordset!idFuenteFinanciamiento)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.FechaModificacion = IIf(IsNull(oRecordset!FechaModificacion), 0, oRecordset!FechaModificacion)
       oTabla.idUsuarioModifica = IIf(IsNull(oRecordset!idUsuarioModifica), 0, oRecordset!idUsuarioModifica)
  
   End If
   oRecordset.Close
  SeleccionarPorId = True
   ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOfarmMovimientoNotaIngreso)
   oTabla.movNumero = ""
   oTabla.movTipo = ""
   oTabla.DocumentoFechaRecepcion = Now
   oTabla.OrigenIdTipo = 0
   oTabla.OrigenNumero = ""
   oTabla.OrigenFecha = Now
   oTabla.idProveedor = 0
   oTabla.idTipoCompra = 0
   oTabla.idTipoProceso = 0
   oTabla.NumeroProceso = ""
   oTabla.IdPaciente = 0
   oTabla.IdComprobantePago = 0
   oTabla.FechaModificacion = Now
   oTabla.idUsuarioModifica = 0
End Sub


