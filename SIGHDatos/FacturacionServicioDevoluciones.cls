VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturacionServicioDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FacturacionServicioDevoluciones
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
'   Descripción:    Inserta un registro a la tabla FacturacionServicioDevoluciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOFacturacionServicioDev) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionServicioDevolucionesAgregar"
   Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadAdevolver", adInteger, adParamInput, 0, IIf(oTabla.CantidadAdevolver = 0, Null, oTabla.CantidadAdevolver))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoDevolucion", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoDevolucion = 0, Null, oTabla.IdEstadoDevolucion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaAutoriza", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaAutoriza = 0, Null, oTabla.FechaAutoriza))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAutoriza", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutoriza = 0, Null, oTabla.IdUsuarioAutoriza))
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
'   Descripción:    modifica un registro a la tabla FacturacionServicioDevoluciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFacturacionServicioDev) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionServicioDevolucionesModificar"
   Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadAdevolver", adInteger, adParamInput, 0, IIf(oTabla.CantidadAdevolver = 0, Null, oTabla.CantidadAdevolver))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoDevolucion", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoDevolucion = 0, Null, oTabla.IdEstadoDevolucion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaAutoriza", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaAutoriza = 0, Null, oTabla.FechaAutoriza))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAutoriza", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutoriza = 0, Null, oTabla.IdUsuarioAutoriza))
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
'   Descripción:    elimia o borra un registro a la tabla FacturacionServicioDevoluciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFacturacionServicioDev) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionServicioDevolucionesEliminar"
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
'   Descripción:    Selecciona un unico registro de la tabla FacturacionServicioDevoluciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFacturacionServicioDev) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FacturacionServicioDevolucionesSeleccionarPorId"
     Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, oTabla.IdOrden)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdOrden = IIf(IsNull(oRecordset!IdOrden), 0, oRecordset!IdOrden)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.CantidadAdevolver = IIf(IsNull(oRecordset!CantidadAdevolver), 0, oRecordset!CantidadAdevolver)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.IdEstadoDevolucion = IIf(IsNull(oRecordset!IdEstadoDevolucion), 0, oRecordset!IdEstadoDevolucion)
       oTabla.FechaAutoriza = IIf(IsNull(oRecordset!FechaAutoriza), 0, oRecordset!FechaAutoriza)
       oTabla.IdUsuarioAutoriza = IIf(IsNull(oRecordset!IdUsuarioAutoriza), 0, oRecordset!IdUsuarioAutoriza)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOFacturacionServicioDev)
   oTabla.IdOrden = 0
   oTabla.idProducto = 0
   oTabla.CantidadAdevolver = 0
   oTabla.IdComprobantePago = 0
   oTabla.IdEstadoDevolucion = 0
   oTabla.FechaAutoriza = Now
   oTabla.IdUsuarioAutoriza = 0
End Sub
Function SeleccionarPorIdOrdenIdProducto(lnIdOrden As Long, lnIdProducto As Long) As Recordset
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
        .CommandText = "FacturacionServicioDevolucionesSeleccionarPorIdOrdenIdProducto"
        Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorIdOrdenIdProducto = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

