VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturacionServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: FacturacionServicios
'        Autor: William Castro Grijalva
'        Fecha: 04/08/2006 12:17:33 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------
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
'   Descripción:    Inserta un registro a la tabla FacturacionServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOFacturacionServicios) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionServiciosAgregar"
           Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, IIf(oTabla.PrecioUnitario = 0, Null, oTabla.PrecioUnitario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoFacturacion", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoFacturacion = 0, Null, oTabla.IdEstadoFacturacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.IdFuenteFinanciamiento = 0, Null, oTabla.IdFuenteFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.IdTipoFinanciamiento = 0, Null, oTabla.IdTipoFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPartidaPresupuestal", adInteger, adParamInput, 0, IIf(oTabla.IdPartidaPresupuestal = 0, Null, oTabla.IdPartidaPresupuestal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCentroCosto", adInteger, adParamInput, 0, IIf(oTabla.IdCentroCosto = 0, Null, oTabla.IdCentroCosto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioInternamiento", adInteger, adParamInput, 0, IIf(oTabla.IdServicioInternamiento = 0, Null, oTabla.IdServicioInternamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalPorPagar", adCurrency, adParamInput, 0, IIf(oTabla.TotalPorPagar = 0, Null, oTabla.TotalPorPagar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdAtencion = 0, Null, oTabla.IdAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFacturacionServicio", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioModifica", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioModifica = 0, Null, oTabla.IdUsuarioModifica)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 4, IIf(oTabla.IdPuntoCarga = 0, Null, oTabla.IdPuntoCarga)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaModificacion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaModificacion = 0, Null, oTabla.FechaModificacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 4, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaSeguro", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioAutorizaSeguro = 0, Null, oTabla.IdUsuarioAutorizaSeguro)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaPendiente", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioAutorizaPendiente = 0, Null, oTabla.IdUsuarioAutorizaPendiente)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaDevolucion", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioAutorizaDevolucion = 0, Null, oTabla.IdUsuarioAutorizaDevolucion)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 4, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaSeguro", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaSeguro = 0, Null, oTabla.FechaAutorizaSeguro)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaPendiente", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaPendiente = 0, Null, oTabla.FechaAutorizaPendiente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaDevolucion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaDevolucion = 0, Null, oTabla.FechaAutorizaDevolucion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCajero", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaCajero = 0, Null, oTabla.FechaCajero)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePagoDevolucion", adInteger, adParamInput, 4, IIf(oTabla.IdComprobantePagoDevolucion = 0, Null, oTabla.IdComprobantePagoDevolucion)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Cantidad", adDecimal, adParamInput, 5, IIf(oTabla.Cantidad = 0, Null, oTabla.Cantidad)): oParameter.Precision = 10: oParameter.NumericScale = 2: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioCrea", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioCrea = 0, Null, oTabla.IdUsuarioCrea)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
           
           Set oParameter = .CreateParameter("@ImporteSIS", adCurrency, adParamInput, 0, IIf(oTabla.ImporteSIS = 0, Null, oTabla.ImporteSIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImporteEXO", adCurrency, adParamInput, 0, IIf(oTabla.ImporteEXO = 0, Null, oTabla.ImporteEXO)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaEXO", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioAutorizaEXO = 0, Null, oTabla.IdUsuarioAutorizaEXO)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaEXO", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaEXO = 0, Null, oTabla.FechaAutorizaEXO)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProductoLocal", adInteger, adParamInput, 4, IIf(oTabla.IdProductoLocal = 0, Null, oTabla.IdProductoLocal)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImporteSOAT", adCurrency, adParamInput, 0, IIf(oTabla.ImporteSOAT = 0, Null, oTabla.ImporteSOAT)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalPagar", adCurrency, adParamInput, 0, IIf(oTabla.TotalPagar = 0, Null, oTabla.TotalPagar)): .Parameters.Append oParameter

       .Execute
           oTabla.IdFacturacionServicio = .Parameters("@IdFacturacionServicio")
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
'   Descripción:    modifica un registro a la tabla FacturacionServicios
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFacturacionServicios) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionServiciosModificar"
           Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, IIf(oTabla.PrecioUnitario = 0, Null, oTabla.PrecioUnitario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoFacturacion", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoFacturacion = 0, Null, oTabla.IdEstadoFacturacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.IdFuenteFinanciamiento = 0, Null, oTabla.IdFuenteFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.IdTipoFinanciamiento = 0, Null, oTabla.IdTipoFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPartidaPresupuestal", adInteger, adParamInput, 0, IIf(oTabla.IdPartidaPresupuestal = 0, Null, oTabla.IdPartidaPresupuestal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCentroCosto", adInteger, adParamInput, 0, IIf(oTabla.IdCentroCosto = 0, Null, oTabla.IdCentroCosto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioInternamiento", adInteger, adParamInput, 0, IIf(oTabla.IdServicioInternamiento = 0, Null, oTabla.IdServicioInternamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalPorPagar", adCurrency, adParamInput, 0, IIf(oTabla.TotalPorPagar = 0, Null, oTabla.TotalPorPagar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdAtencion = 0, Null, oTabla.IdAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFacturacionServicio", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionServicio = 0, Null, oTabla.IdFacturacionServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioModifica", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioModifica = 0, Null, oTabla.IdUsuarioModifica)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 4, IIf(oTabla.IdPuntoCarga = 0, Null, oTabla.IdPuntoCarga)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaModificacion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaModificacion = 0, Null, oTabla.FechaModificacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 4, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaSeguro", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioAutorizaSeguro = 0, Null, oTabla.IdUsuarioAutorizaSeguro)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaPendiente", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioAutorizaPendiente = 0, Null, oTabla.IdUsuarioAutorizaPendiente)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaDevolucion", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioAutorizaDevolucion = 0, Null, oTabla.IdUsuarioAutorizaDevolucion)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 4, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaSeguro", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaSeguro = 0, Null, oTabla.FechaAutorizaSeguro)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaPendiente", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaPendiente = 0, Null, oTabla.FechaAutorizaPendiente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaDevolucion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaDevolucion = 0, Null, oTabla.FechaAutorizaDevolucion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCajero", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaCajero = 0, Null, oTabla.FechaCajero)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePagoDevolucion", adInteger, adParamInput, 4, IIf(oTabla.IdComprobantePagoDevolucion = 0, Null, oTabla.IdComprobantePagoDevolucion)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Cantidad", adDecimal, adParamInput, 5, IIf(oTabla.Cantidad = 0, Null, oTabla.Cantidad)): oParameter.Precision = 10: oParameter.NumericScale = 2: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioCrea", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioCrea = 0, Null, oTabla.IdUsuarioCrea)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
       
           Set oParameter = .CreateParameter("@ImporteSIS", adCurrency, adParamInput, 0, IIf(oTabla.ImporteSIS = 0, Null, oTabla.ImporteSIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImporteEXO", adCurrency, adParamInput, 0, IIf(oTabla.ImporteEXO = 0, Null, oTabla.ImporteEXO)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaEXO", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioAutorizaEXO = 0, Null, oTabla.IdUsuarioAutorizaEXO)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaEXO", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaEXO = 0, Null, oTabla.FechaAutorizaEXO)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProductoLocal", adInteger, adParamInput, 4, IIf(oTabla.IdProductoLocal = 0, Null, oTabla.IdProductoLocal)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImporteSOAT", adCurrency, adParamInput, 0, IIf(oTabla.ImporteSOAT = 0, Null, oTabla.ImporteSOAT)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalPagar", adCurrency, adParamInput, 0, IIf(oTabla.TotalPagar = 0, Null, oTabla.TotalPagar)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla FacturacionServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFacturacionServicios) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionServiciosEliminar"
           Set oParameter = .CreateParameter("@IdFacturacionServicio", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionServicio = 0, Null, oTabla.IdFacturacionServicio)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla FacturacionServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFacturacionServicios) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FacturacionServiciosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdFacturacionServicio", adInteger, adParamInput, 0, oTabla.IdFacturacionServicio): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.PrecioUnitario = IIf(IsNull(oRecordset!PrecioUnitario), 0, oRecordset!PrecioUnitario)
       oTabla.IdEstadoFacturacion = IIf(IsNull(oRecordset!IdEstadoFacturacion), 0, oRecordset!IdEstadoFacturacion)
       oTabla.IdFuenteFinanciamiento = IIf(IsNull(oRecordset!IdFuenteFinanciamiento), 0, oRecordset!IdFuenteFinanciamiento)
       oTabla.IdTipoFinanciamiento = IIf(IsNull(oRecordset!IdTipoFinanciamiento), 0, oRecordset!IdTipoFinanciamiento)
       oTabla.IdPartidaPresupuestal = IIf(IsNull(oRecordset!IdPartidaPresupuestal), 0, oRecordset!IdPartidaPresupuestal)
       oTabla.IdCentroCosto = IIf(IsNull(oRecordset!IdCentroCosto), 0, oRecordset!IdCentroCosto)
       oTabla.IdServicioInternamiento = IIf(IsNull(oRecordset!IdServicioInternamiento), 0, oRecordset!IdServicioInternamiento)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.TotalPorPagar = IIf(IsNull(oRecordset!TotalPorPagar), 0, oRecordset!TotalPorPagar)
       oTabla.IdAtencion = IIf(IsNull(oRecordset!IdAtencion), 0, oRecordset!IdAtencion)
       oTabla.IdFacturacionServicio = IIf(IsNull(oRecordset!IdFacturacionServicio), 0, oRecordset!IdFacturacionServicio)
       oTabla.IdUsuarioModifica = IIf(IsNull(oRecordset!IdUsuarioModifica), 0, oRecordset!IdUsuarioModifica)
       oTabla.IdPuntoCarga = IIf(IsNull(oRecordset!IdPuntoCarga), 0, oRecordset!IdPuntoCarga)
       oTabla.FechaModificacion = IIf(IsNull(oRecordset!FechaModificacion), 0, oRecordset!FechaModificacion)
       oTabla.IdOrden = IIf(IsNull(oRecordset!IdOrden), 0, oRecordset!IdOrden)
       oTabla.IdUsuarioAutorizaSeguro = IIf(IsNull(oRecordset!IdUsuarioAutorizaSeguro), 0, oRecordset!IdUsuarioAutorizaSeguro)
       oTabla.IdUsuarioAutorizaPendiente = IIf(IsNull(oRecordset!IdUsuarioAutorizaPendiente), 0, oRecordset!IdUsuarioAutorizaPendiente)
       oTabla.IdUsuarioAutorizaDevolucion = IIf(IsNull(oRecordset!IdUsuarioAutorizaDevolucion), 0, oRecordset!IdUsuarioAutorizaDevolucion)
       oTabla.IdCajero = IIf(IsNull(oRecordset!IdCajero), 0, oRecordset!IdCajero)
       oTabla.FechaAutorizaSeguro = IIf(IsNull(oRecordset!FechaAutorizaSeguro), 0, oRecordset!FechaAutorizaSeguro)
       oTabla.FechaAutorizaPendiente = IIf(IsNull(oRecordset!FechaAutorizaPendiente), 0, oRecordset!FechaAutorizaPendiente)
       oTabla.FechaAutorizaDevolucion = IIf(IsNull(oRecordset!FechaAutorizaDevolucion), 0, oRecordset!FechaAutorizaDevolucion)
       oTabla.FechaCajero = IIf(IsNull(oRecordset!FechaCajero), 0, oRecordset!FechaCajero)
       oTabla.IdComprobantePagoDevolucion = IIf(IsNull(oRecordset!IdComprobantePagoDevolucion), 0, oRecordset!IdComprobantePagoDevolucion)
       oTabla.Cantidad = IIf(IsNull(oRecordset!Cantidad), 0, oRecordset!Cantidad)
       oTabla.IdUsuarioCrea = IIf(IsNull(oRecordset!IdUsuarioCrea), 0, oRecordset!IdUsuarioCrea)
       
       oTabla.ImporteSIS = IIf(IsNull(oRecordset!ImporteSIS), 0, oRecordset!ImporteSIS)
       oTabla.IdUsuarioAutorizaEXO = IIf(IsNull(oRecordset!IdUsuarioAutorizaEXO), 0, oRecordset!IdUsuarioAutorizaEXO)
       oTabla.FechaAutorizaEXO = IIf(IsNull(oRecordset!FechaAutorizaEXO), 0, oRecordset!FechaAutorizaEXO)
       oTabla.IdProductoLocal = IIf(IsNull(oRecordset!IdProductoLocal), 0, oRecordset!IdProductoLocal)
       oTabla.ImporteEXO = IIf(IsNull(oRecordset!ImporteEXO), 0, oRecordset!ImporteEXO)
       oTabla.ImporteSOAT = IIf(IsNull(oRecordset!ImporteSOAT), 0, oRecordset!ImporteSOAT)
       oTabla.TotalPagar = IIf(IsNull(oRecordset!TotalPagar), 0, oRecordset!TotalPagar)
   
       oTabla.CantidadSIS = IIf(IsNull(oRecordset!CantidadSIS), 0, oRecordset!CantidadSIS)
       oTabla.CantidadSOAT = IIf(IsNull(oRecordset!CantidadSOAT), 0, oRecordset!CantidadSOAT)
       oTabla.CantidadPagar = IIf(IsNull(oRecordset!CantidadPagar), 0, oRecordset!CantidadPagar)
       oTabla.PrecioSIS = IIf(IsNull(oRecordset!PrecioSIS), 0, oRecordset!PrecioSIS)
       oTabla.precioSOAT = IIf(IsNull(oRecordset!precioSOAT), 0, oRecordset!precioSOAT)
   
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


'------------------------------------------------------------------------------------
'   Función:        InsertarServicioSP
'   Descripción:    Inserta un registro a la tabla FacturacionServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function InsertarServicioSP(ByVal oTabla As DOFacturacionServicios, ByRef sResultado As String) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   InsertarServicioSP = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionServiciosAgregarV2"
        Set oParameter = .CreateParameter("@IdFacturacionServicio", adInteger, adParamOutput, 0): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@result", adVarChar, adParamOutput, 500): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
        'Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, IIf(oTabla.Cantidad = 0, 0, oTabla.Cantidad)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Cantidad", adDecimal, adParamInput, 5, IIf(oTabla.Cantidad = 0, 0, oTabla.Cantidad)): oParameter.Precision = 10: oParameter.NumericScale = 2: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdEstadoFacturacion", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoFacturacion = 0, Null, oTabla.IdEstadoFacturacion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.IdFuenteFinanciamiento = 0, Null, oTabla.IdFuenteFinanciamiento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.IdTipoFinanciamiento = 0, Null, oTabla.IdTipoFinanciamiento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdServicioInternamiento", adInteger, adParamInput, 0, IIf(oTabla.IdServicioInternamiento = 0, Null, oTabla.IdServicioInternamiento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 4, IIf(oTabla.IdPuntoCarga = 0, Null, oTabla.IdPuntoCarga)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 4, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAutorizaSeguro", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioAutorizaSeguro = 0, Null, oTabla.IdUsuarioAutorizaSeguro)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAutorizaPendiente", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioAutorizaPendiente = 0, Null, oTabla.IdUsuarioAutorizaPendiente)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAutorizaDevolucion", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioAutorizaDevolucion = 0, Null, oTabla.IdUsuarioAutorizaDevolucion)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 4, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaAutorizaSeguro", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaSeguro = 0, Null, oTabla.FechaAutorizaSeguro)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaAutorizaPendiente", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaPendiente = 0, Null, oTabla.FechaAutorizaPendiente)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaAutorizaDevolucion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaDevolucion = 0, Null, oTabla.FechaAutorizaDevolucion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaCajero", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaCajero = 0, Null, oTabla.FechaCajero)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdAtencion = 0, Null, oTabla.IdAtencion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAuditoria = 0, Null, oTabla.IdUsuarioAuditoria)): .Parameters.Append oParameter
       .Execute
        If .Parameters("@result") <> "" Then
            sResultado = .Parameters("@result")
            Exit Function
        End If
        
        oTabla.IdFacturacionServicio = .Parameters("@IdFacturacionServicio")
   End With
 
   InsertarServicioSP = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "InsertarServicioSP"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        ModificarServicioSP
'   Descripción:    Inserta un registro a la tabla FacturacionServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function ModificarServicioSP(ByVal oTabla As DOFacturacionServicios, ByRef sResultado As String) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   ModificarServicioSP = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionServiciosModificarV2"
        Set oParameter = .CreateParameter("@IdFacturacionServicio", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionServicio = 0, Null, oTabla.IdFacturacionServicio)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@result", adVarChar, adParamOutput, 500): .Parameters.Append oParameter
        'Set oParameter = .CreateParameter("@IdUsuarioModifica", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioModifica = 0, Null, oTabla.IdUsuarioModifica)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
        'Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, IIf(oTabla.Cantidad = 0, Null, oTabla.Cantidad)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Cantidad", adDecimal, adParamInput, 5, IIf(oTabla.Cantidad = 0, Null, oTabla.Cantidad)): oParameter.Precision = 10: oParameter.NumericScale = 2: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdEstadoFacturacion", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoFacturacion = 0, Null, oTabla.IdEstadoFacturacion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.IdFuenteFinanciamiento = 0, Null, oTabla.IdFuenteFinanciamiento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.IdTipoFinanciamiento = 0, Null, oTabla.IdTipoFinanciamiento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdServicioInternamiento", adInteger, adParamInput, 0, IIf(oTabla.IdServicioInternamiento = 0, Null, oTabla.IdServicioInternamiento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
        'Set oParameter = .CreateParameter("@TotalPorPagar", adCurrency, adParamInput, 0, IIf(oTabla.TotalPorPagar = 0, Null, oTabla.TotalPorPagar)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 4, IIf(oTabla.IdPuntoCarga = 0, Null, oTabla.IdPuntoCarga)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        'Set oParameter = .CreateParameter("@FechaModificacion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaModificacion = 0, Null, oTabla.FechaModificacion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 4, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAutorizaSeguro", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioAutorizaSeguro = 0, Null, oTabla.IdUsuarioAutorizaSeguro)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAutorizaPendiente", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioAutorizaPendiente = 0, Null, oTabla.IdUsuarioAutorizaPendiente)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAutorizaDevolucion", adInteger, adParamInput, 4, IIf(oTabla.IdUsuarioAutorizaDevolucion = 0, Null, oTabla.IdUsuarioAutorizaDevolucion)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 4, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaAutorizaSeguro", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaSeguro = 0, Null, oTabla.FechaAutorizaSeguro)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaAutorizaPendiente", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaPendiente = 0, Null, oTabla.FechaAutorizaPendiente)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaAutorizaDevolucion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaDevolucion = 0, Null, oTabla.FechaAutorizaDevolucion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaCajero", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaCajero = 0, Null, oTabla.FechaCajero)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdAtencion = 0, Null, oTabla.IdAtencion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdComprobantePagoDevolucion", adInteger, adParamInput, 4, IIf(oTabla.IdComprobantePagoDevolucion = 0, Null, oTabla.IdComprobantePagoDevolucion)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAuditoria = 0, Null, oTabla.IdUsuarioAuditoria)): .Parameters.Append oParameter
       
       .Execute
        If .Parameters("@result") <> "" Then
            sResultado = .Parameters("@result")
            Exit Function
        End If
        
        'oTabla.IdFacturacionServicio = .Parameters("@IdFacturacionServicio")
   End With
 
   ModificarServicioSP = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "ModificarServicioSP"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        EliminarServicioSP
'   Descripción:    Inserta un registro a la tabla FacturacionServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function EliminarServicioSP(ByVal oTabla As DOFacturacionServicios, ByRef sResultado As String) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   EliminarServicioSP = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionServiciosEliminarServicio"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdAtencion = 0, Null, oTabla.IdAtencion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
 
   EliminarServicioSP = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "EliminarServicioSP"
Exit Function
End Function

Function ModificarComoFacturado(oTabla As DOFacturacionServicios) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim SQL As String
 
   ModificarComoFacturado = False
   
'   SQL = "Update FacturacionServicios set IdEstadoFacturacion = " & oTabla.IdEstadoFacturacion _
'            & " ,IdComprobanteDetalleServicio = " & oTabla.IdComprobanteDetalleServicio _
'            & " where IdFacturacionServicio  = " & oTabla.IdFacturacionServicio

    SQL = "Update FacturacionServicios set " _
            & " IdEstadoFacturacion = Case " _
            & " when SubTotalPorPagar - (SubTotalPagado + " & Str(oTabla.TotalPorPagar) & ")<= 0  then " & sghEstadoFacturacion.sghPagado _
            & " Else " & sghEstadoFacturacion.sghPendientePago & " end  " _
            & " ,SubTotalPagadoACuenta = Case " _
            & " when SubTotalPorPagar - (SubTotalPagado + " & Str(oTabla.TotalPorPagar) & ") > 0  then SubTotalPagadoACuenta + " & Str(oTabla.TotalPorPagar) _
            & " else SubTotalPagadoACuenta end " _
            & " ,SubTotalPagado = round(SubTotalPagado + " & Str(oTabla.TotalPorPagar) & ",2) " _
            & " ,SubTotalPorPagar = round(SubTotalPorPagar - " & Str(oTabla.TotalPorPagar) & ",2) " _
            & " where IdFacturacionServicio  = " & oTabla.IdFacturacionServicio
   
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = SQL
       .Execute
   End With
 
   ModificarComoFacturado = True
   ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "ModificarComoFacturado"
Exit Function

End Function

'Function SeleccionarParaExoneracion(lIdCuentaAtencion As Long) As Recordset
'On Error GoTo ManejadorDeError
'Dim oRecordset As New ADODB.Recordset
'Dim oCommand As New ADODB.Command
'Dim sSQL As String
'
'    Set SeleccionarParaExoneracion = Nothing
'    sSQL = "Select fs.IdFacturacionServicio,fs.IdProducto,fs.PrecioUnitario,fs.Cantidad,fs.SubTotalPorPagar " _
'            & " ,fs.SubTotalExonerado,fs.SubTotalPagadoACuenta,fs.IdEstadoAtencion,fs.IdEstadoFacturacion" _
'            & " ,cs.Codigo CodProducto,cs.Nombre NombreServicio,ord.NroOrden,ord.FechaOrden  from" _
'            & " (FacturacionServicios fs  join CatalogoServicios cs on fs.IdProducto = cs.IdProducto)" _
'            & " left join (Select ap.IdCuentaAtencion,ap.NroOrden,ap.FechaOrden, ad.IdFacturacionServicio" _
'            & " from AtencionProcedimientoDetalle ad, AtencionProcedimientos ap " _
'            & " Where ad.IdAtencionProcedimiento = ap.IdAtencionProcedimiento and ap.IdCuentaAtencion = " & lIdCuentaAtencion _
'            & " Union All " _
'            & " Select ap.IdCuentaAtencion,ap.OrdenNro,ap.FechaOrden, ad.IdFacturacionServicio " _
'            & " from AtencionApoyoDiagDetalle ad, AtencionApoyoDiagnostico ap " _
'            & " Where ad.IdAtencionApoyoDx = ap.IdAtencionApoyoDx and ap.IdCuentaAtencion = " & lIdCuentaAtencion _
'            & " ) ord on fs.IdFacturacionServicio = ord.IdFacturacionServicio " _
'            & " Where fs.IdCuentaAtencion = " & lIdCuentaAtencion
'
'    With oCommand
'        .CommandType = adCmdText
'        Set .ActiveConnection = mo_Conexion
'        .CommandText = sSQL
'        Set oRecordset = .Execute
'    End With
'    Set oRecordset.ActiveConnection = Nothing
'
'    Set SeleccionarParaExoneracion = oRecordset
'
'    ms_MensajeError = ""
'
'
'
'Exit Function
'ManejadorDeError:
'   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
'Exit Function
'
'End Function

Function SeleccionarParaPendientePago(lIdCuentaAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sSQL1 As String
Dim sSQL2 As String

     Set SeleccionarParaPendientePago = Nothing
     
    sSQL1 = "select CategoriasProducto.IdCategoriaProducto, CategoriasProducto.Descripcion," + _
    " sum(FacturacionServicios.SubTotalPorPagar) as SubTotalPorPagar," + _
    " sum(FacturacionServicios.SubTotalPendientePago) as SubTotalPendientePago " + _
    " from " + _
    " FacturacionServicios left join CategoriasProducto" + _
    " on FacturacionServicios.IdCategoriaProducto = CategoriasProducto.IdCategoriaProducto" + _
    " Where FacturacionServicios.IdCuentaAtencion = " & lIdCuentaAtencion & _
    " group by CategoriasProducto.IdCategoriaProducto, CategoriasProducto.Descripcion" + _
    " union " + _
    "select 0 , 'Total' ," + _
    " sum(FacturacionServicios.SubTotalPorPagar) as SubTotalPorPagar," + _
    " sum(FacturacionServicios.SubTotalPendientePago) as SubTotalPendientePago " + _
    " from " + _
    " FacturacionServicios " + _
    " Where FacturacionServicios.IdCuentaAtencion = " & lIdCuentaAtencion
    
    sSQL2 = "select " + _
    " FacturacionServicios.IdCategoriaProducto ,  FactCatalogoServicios.Nombre, " + _
    " FacturacionServicios.PrecioUnitario ,  FacturacionServicios.Cantidad, " + _
    " FacturacionServicios.SubTotalPorPagar as SubTotalPorPagar, " + _
    " FacturacionServicios.SubTotalPendientePago as SubTotalPendientePago,  " + _
    " FacturacionServicios.IdEstadoFacturacion as IdEstadoFacturacion," + _
    " FacturacionServicios.IdEmpleadoAutorizaPendiente as IdEmpleadoAutorizaPendiente,  " + _
    " FacturacionServicios.FechaAutorizaPendiente as FechaAutorizaPendiente, " + _
    " FacturacionServicios.IdEmpleadoModifica as IdEmpleadoModifica,  " + _
    " FacturacionServicios.FechaModificacion as FechaModificacion " + _
    " from " + _
    " ((FacturacionServicios  left join FactCatalogoServicios on FacturacionServicios.IdProducto = FactCatalogoServicios.IdProducto ) " + _
    " left join Empleados on FacturacionServicios.IdEmpleadoAutorizaPendiente = Empleados.IdEmpleado) " + _
    " where FacturacionServicios.IdCuentaAtencion = " & lIdCuentaAtencion
    
    sSql = "SHAPE {" + sSQL1 + "} " + _
            "APPEND ({" + sSQL2 + "} " + _
            " AS title_chap RELATE IdCategoriaProducto TO IdCategoriaProducto) "


    oRecordset.Open sSql, mo_Conexion, adOpenDynamic, adLockOptimistic
    Set oRecordset.ActiveConnection = Nothing
    
    Set SeleccionarParaPendientePago = oRecordset
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
      Trace Err.Number & " " + Err.Description, "SeleccionarParaPendientePago"
Exit Function

End Function

Function SeleccionarParaEstadoCuenta(lIdCuentaAtencion As Long, TipoServicioARecuperar As sghTipoFacturacionServicio) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String

    Set SeleccionarParaEstadoCuenta = Nothing
    Select Case TipoServicioARecuperar
        Case sghTipoFacturacionServicio.sghFacturacionServicioPorEstancia
            sSql = "Select fs.IdFacturacionServicio,fs.IdProducto,fs.PrecioUnitario,fs.Cantidad,fs.SubTotalPorPagar " _
                    & " ,fs.SubTotalExonerado,fs.SubTotalPagadoACuenta,fs.IdEstadoAtencion,fs.IdEstadoFacturacion" _
                    & " ,cs.Codigo CodProducto,cs.Nombre + ' (* Estancia)' NombreServicio,ord.NroOrden,ord.FechaOrden  from" _
                    & " (FacturacionServicios fs  join Servicios cs on fs.IdServicioFactura = cs.IdServicio)" _
                    & " left join (Select ap.IdCuentaAtencion,ap.NroOrden,ap.FechaOrden, ad.IdFacturacionServicio" _
                    & " from AtencionProcedimientoDetalle ad, AtencionProcedimientos ap " _
                    & " Where ad.IdAtencionProcedimiento = ap.IdAtencionProcedimiento and ap.IdCuentaAtencion = " & lIdCuentaAtencion & " ) ord " _
                    & " on fs.IdFacturacionServicio = ord.IdFacturacionServicio " _
                    & " Where fs.IdCuentaAtencion = " & lIdCuentaAtencion _
                    & " and isnull(fs.ServicioPorEstancia,0) = 1 "
        
        Case sghTipoFacturacionServicio.sghFacturacionServicioPorProcedimiento
            sSql = "Select fs.IdFacturacionServicio,fs.IdProducto,fs.PrecioUnitario,fs.Cantidad,fs.SubTotalPorPagar " _
                    & " ,fs.SubTotalExonerado,fs.SubTotalPagadoACuenta,fs.IdEstadoAtencion,fs.IdEstadoFacturacion" _
                    & " ,cs.Codigo CodProducto,cs.Nombre NombreServicio,ord.NroOrden,ord.FechaOrden  from" _
                    & " (FacturacionServicios fs  join FactCatalogoServicios cs on fs.IdProducto = cs.IdProducto)" _
                    & " left join (Select ap.IdCuentaAtencion,ap.NroOrden,ap.FechaOrden, ad.IdFacturacionServicio" _
                    & " from AtencionProcedimientoDetalle ad, AtencionProcedimientos ap " _
                    & " Where ad.IdAtencionProcedimiento = ap.IdAtencionProcedimiento and ap.IdCuentaAtencion = " & lIdCuentaAtencion _
                    & " Union All " _
                    & " Select ap.IdCuentaAtencion,ap.OrdenNro,ap.FechaOrden, ad.IdFacturacionServicio " _
                    & " from AtencionApoyoDiagDetalle ad, AtencionApoyoDiagnostico ap " _
                    & " Where ad.IdAtencionApoyoDx = ap.IdAtencionApoyoDx and ap.IdCuentaAtencion = " & lIdCuentaAtencion _
                    & " ) ord on fs.IdFacturacionServicio = ord.IdFacturacionServicio " _
                    & " Where fs.IdCuentaAtencion = " & lIdCuentaAtencion _
                    & " and isnull(fs.ServicioPorEstancia,0) <> 1 "
        
        
        Case sghTipoFacturacionServicio.sghFacturacionServicioTotal
            sSql = "Select fs.IdFacturacionServicio,fs.IdProducto,fs.PrecioUnitario,fs.Cantidad,fs.SubTotalPorPagar " _
                    & " ,fs.SubTotalExonerado,fs.SubTotalPagadoACuenta,fs.IdEstadoAtencion,fs.IdEstadoFacturacion" _
                    & " ,cs.Codigo CodProducto,cs.Nombre NombreServicio,ord.NroOrden,ord.FechaOrden  from" _
                    & " (FacturacionServicios fs  join FactCatalogoServicios cs on fs.IdProducto = cs.IdProducto)" _
                    & " left join (Select ap.IdCuentaAtencion,ap.NroOrden,ap.FechaOrden, ad.IdFacturacionServicio" _
                    & " from AtencionProcedimientoDetalle ad, AtencionProcedimientos ap " _
                    & " Where ad.IdAtencionProcedimiento = ap.IdAtencionProcedimiento and ap.IdCuentaAtencion = " & lIdCuentaAtencion _
                    & " Union All " _
                    & " Select ap.IdCuentaAtencion,ap.OrdenNro,ap.FechaOrden, ad.IdFacturacionServicio " _
                    & " from AtencionApoyoDiagDetalle ad, AtencionApoyoDiagnostico ap " _
                    & " Where ad.IdAtencionApoyoDx = ap.IdAtencionApoyoDx and ap.IdCuentaAtencion = " & lIdCuentaAtencion _
                    & " ) ord on fs.IdFacturacionServicio = ord.IdFacturacionServicio " _
                    & " Where fs.IdCuentaAtencion = " & lIdCuentaAtencion _
                    & " and isnull(fs.ServicioPorEstancia,0) <> 1 " _
                    & " UNION ALL Select fs.IdFacturacionServicio,fs.IdProducto,fs.PrecioUnitario,fs.Cantidad,fs.SubTotalPorPagar " _
                    & " ,fs.SubTotalExonerado,fs.SubTotalPagadoACuenta,fs.IdEstadoAtencion,fs.IdEstadoFacturacion" _
                    & " ,cs.Codigo CodProducto,cs.Nombre + ' (* Estancia)' NombreServicio,ord.NroOrden,ord.FechaOrden  from" _
                    & " (FacturacionServicios fs  join Servicios cs on fs.IdServicioFactura = cs.IdServicio)" _
                    & " left join (Select ap.IdCuentaAtencion,ap.NroOrden,ap.FechaOrden, ad.IdFacturacionServicio" _
                    & " from AtencionProcedimientoDetalle ad, AtencionProcedimientos ap " _
                    & " Where ad.IdAtencionProcedimiento = ap.IdAtencionProcedimiento and ap.IdCuentaAtencion = " & lIdCuentaAtencion & " ) ord " _
                    & " on fs.IdFacturacionServicio = ord.IdFacturacionServicio " _
                    & " Where fs.IdCuentaAtencion = " & lIdCuentaAtencion _
                    & " and isnull(fs.ServicioPorEstancia,0) = 1 "
    End Select
    With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandText = sSql
        Set oRecordset = .Execute
    End With
    Set oRecordset.ActiveConnection = Nothing
     
    Set SeleccionarParaEstadoCuenta = oRecordset
    
    ms_MensajeError = ""
     
    
   
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "SeleccionarParaEstadoCuenta"
Exit Function

End Function

Function SeleccionarParaCaja(lIdCuentaAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sSQL1 As String
Dim sSQL2 As String

     Set SeleccionarParaCaja = Nothing
     
    sSQL1 = "select CategoriasProducto.IdCategoriaProducto, CategoriasProducto.Descripcion," + _
    " sum(FacturacionServicios.SubTotalPorPagar) as SubTotalPorPagar," + _
    " sum(FacturacionServicios.SubTotalPagadoACuenta) as SubTotalPagadoACuenta,  " + _
    " sum(FacturacionServicios.SubTotalPendientePago) as SubTotalPendientePago,  " + _
    " sum(FacturacionServicios.SubTotalPagado) as SubTotalPagado  " + _
    " from " + _
    " FacturacionServicios left join CategoriasProducto" + _
    " on FacturacionServicios.IdCategoriaProducto = CategoriasProducto.IdCategoriaProducto" + _
    " Where FacturacionServicios.IdCuentaAtencion = " & lIdCuentaAtencion & _
    " group by CategoriasProducto.IdCategoriaProducto, CategoriasProducto.Descripcion" + _
    " union " + _
    "select 0 , 'Total' ," + _
    " sum(FacturacionServicios.SubTotalPorPagar) as SubTotalPorPagar," + _
    " sum(FacturacionServicios.SubTotalPendientePago) as SubTotalPendientePago,  " + _
    " sum(FacturacionServicios.SubTotalPagadoACuenta) as SubTotalPagadoACuenta,  " + _
    " sum(FacturacionServicios.SubTotalPagado) as SubTotalPagado  " + _
    " from " + _
    " FacturacionServicios " + _
    " Where FacturacionServicios.IdCuentaAtencion = " & lIdCuentaAtencion & _
    " and FacturacionServicios.IdEstadoFacturacion in (1,2,7)"
    
    sSQL2 = "select " + _
    " FacturacionServicios.IdCategoriaProducto ,  FactCatalogoServicios.Nombre, " + _
    " FacturacionServicios.IdProducto ,  FacturacionServicios.IdFacturacionServicio, " + _
    " FacturacionServicios.PrecioUnitario ,  FacturacionServicios.Cantidad, " + _
    " FacturacionServicios.SubTotalPorPagar as SubTotalPorPagar, " + _
    " FacturacionServicios.SubTotalPagadoACuenta as SubTotalPagadoACuenta,  " + _
    " FacturacionServicios.SubTotalPendientePago as SubTotalPendientePago,  " + _
    " FacturacionServicios.SubTotalPagado as SubTotalPagado,  " + _
    " FacturacionServicios.IdEstadoFacturacion as IdEstadoFacturacion," + _
    " FacturacionServicios.IdEmpleadoModifica as IdEmpleadoModifica,  " + _
    " FacturacionServicios.FechaModificacion as FechaModificacion " + _
    " from " + _
    " ((FacturacionServicios  left join FactCatalogoServicios on FacturacionServicios.IdProducto = FactCatalogoServicios.IdProducto ) " + _
    " left join Empleados on FacturacionServicios.IdEmpleadoAutorizaPendiente = Empleados.IdEmpleado) " + _
    " where FacturacionServicios.IdCuentaAtencion = " & lIdCuentaAtencion & _
    " and FacturacionServicios.IdEstadoFacturacion in (1,2,7)"
    
    sSql = "SHAPE {" + sSQL1 + "} " + _
            "APPEND ({" + sSQL2 + "} " + _
            " AS title_chap RELATE IdCategoriaProducto TO IdCategoriaProducto) "


    oRecordset.Open sSql, mo_Conexion, adOpenDynamic, adLockOptimistic
    Set oRecordset.ActiveConnection = Nothing
    
    Set SeleccionarParaCaja = oRecordset
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "SeleccionarParaCaja"
Exit Function

End Function

Function EliminarServiciosDeLaCuenta(lIdAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
 
   EliminarServiciosDeLaCuenta = False
   
    sSql = "delete from FacturacionServicios where IdAtencion = " & lIdAtencion
   
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSql
       .Execute
   End With
 
   EliminarServiciosDeLaCuenta = True
   ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "EliminarServiciosDeLaCuenta"
Exit Function

End Function

Function ObtenerDatosDeServicioFacturado(lIdAtencion As Long, lIdServicio As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oRecordset As New Recordset
   
   Set ObtenerDatosDeServicioFacturado = Nothing
   
    sSql = "select IdProducto, IdEstadoFacturacion from FacturacionServicios where IdAtencion = " & lIdAtencion & " and IdProducto = " & lIdServicio
   
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSql
       Set oRecordset = .Execute
   End With
 
    Set ObtenerDatosDeServicioFacturado = oRecordset
 
    Set oRecordset.ActiveConnection = Nothing
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
        Trace Err.Number & " " + Err.Description, "ObtenerDatosDeServicioFacturado"
   
Exit Function

End Function
Function ObtieneTipoConsulta(lIdAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oRecordset As New Recordset
   
    Set ObtieneTipoConsulta = Nothing
   
    sSql = "select IdProducto, IdEstadoFacturacion from FacturacionServicios where lIdAtencion = " & lIdAtencion & " and IdProducto in (select ValorInt from Parametros where Tipo = 'TipoConsulta' and Codigo = 'CE') "
   
    With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandText = sSql
        Set oRecordset = .Execute
    End With
 
    Set ObtieneTipoConsulta = oRecordset
    
    Set oRecordset.ActiveConnection = Nothing
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
    Trace Err.Number & " " + Err.Description, "ObtieneTipoConsulta"
Exit Function

End Function

Function EliminarPorIdAtencion(lIdAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
 
   EliminarPorIdAtencion = False
   
    sSql = "delete from FacturacionServicios where IdAtencion = " & lIdAtencion
   
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSql
       .Execute
   End With
 
   EliminarPorIdAtencion = True
   ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
    Trace Err.Number & " " + Err.Description, "EliminarPorIdAtencion"
End Function

Function ActualizarExoneracion(oTabla As DOFacturacionServicios) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
 
   ActualizarExoneracion = False
   
    sSql = "Update FacturacionServicios set SubTotalExonerado = " & Str(oTabla.TotalPorPagar) _
            & " ,SubTotalPorPagar = (Cantidad*PrecioUnitario - (" & Str(oTabla.TotalPorPagar) & "+ SubTotalPagado))" _
            & " ,IdEmpleadoExonera = " & oTabla.IdUsuarioAuditoria _
            & " where IdFacturacionServicio = " & oTabla.IdFacturacionServicio
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSql
       .Execute
   End With
 
   ActualizarExoneracion = True
   ms_MensajeError = ""
Exit Function
ManejadorDeError:
    Trace Err.Number & " " + Err.Description, "ActualizarExoneracion"
Exit Function

End Function

Function AtencionTieneServiciosFacturados(lIdAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   AtencionTieneServiciosFacturados = False
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtencionCantidadServiciosFacturados"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamOutput, 0): .Parameters.Append oParameter
        .Execute
        AtencionTieneServiciosFacturados = .Parameters("@Cantidad").Value > 0
    End With
    ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
       Trace Err.Number & " " + Err.Description, "AtencionTieneServiciosFacturados"
Exit Function
End Function


'EFGL

Function SeleccionarPorCuentaAtencionParaEstadoCuenta(IdCuentaAtencion As Long, IdPuntosDeCarga As Long, lEstadosFacturacion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oRecordset As New Recordset
   
    Set SeleccionarPorCuentaAtencionParaEstadoCuenta = Nothing
   
    sSql = " select a.idfacturacionservicio, a.idProducto,b.codigo as Codigo,b.nombre as Descripcion"
    sSql = sSql & " ,c.idTipoFinanciamiento,'' as Poliza, c.descripcion as TipoFinanciamiento,Cantidad,"
    sSql = sSql & " PrecioUnitario, Totalporpagar as SubTotal,d.idEstadoFacturacion,d.descripcion as Estado, a.IdPuntoCarga, a.IdAtencion, "
    sSql = sSql & " FechaAutorizaPendiente, FechaAutorizaSeguro, IdEmpAutorizaPendiente, IdEmpAutorizaSeguro, IdFuenteFinanciamiento, IdServicioInternamiento, 0 as IdUsuarioAuditoria "
    sSql = sSql & " From facturacionservicios a, factcatalogoservicios b,tiposFinanciamiento c, estadosfacturacion d"
    sSql = sSql & " where a.idproducto = b.idproducto and c.idTipoFinanciamiento = a.idTipoFinanciamiento"
    sSql = sSql & " and a.idEstadoFacturacion = d.idEstadoFacturacion"
    sSql = sSql & " and idatencion in( select idatencion from atenciones"
    sSql = sSql & " where idcuentaatencion = " & IdCuentaAtencion & ")"
    
    If IdPuntosDeCarga > 0 Then
        sSql = sSql & " and a.IdPuntoCarga = " & IdPuntosDeCarga
    End If
    If lEstadosFacturacion <> 0 Then
        sSql = sSql & " and d.IdEstadoFacturacion  = " & lEstadosFacturacion & " "
    End If
   
'    With oCommand
'        .CommandType = adCmdText
'        Set .ActiveConnection = mo_Conexion
'        .CommandText = sSQL
'        Set oRecordset = .Execute
'    End With
oRecordset.Open sSql, mo_Conexion, adOpenKeyset, adLockOptimistic
 
 Set oRecordset.ActiveConnection = Nothing
 
    Set SeleccionarPorCuentaAtencionParaEstadoCuenta = oRecordset
    
    
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
    Trace Err.Number & " " + Err.Description, "SeleccionarPorCuentaAtencionParaEstadoCuenta"
Exit Function

End Function
Function SeleccionarPorCodigoDEBB(codigo As String, lIdTipoFinanciamiento As Long, lIdPuntoCarga As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oRecordset As New Recordset
    sSql = "select FactCatalogoServicios.IdProducto, " + _
            "FactCatalogoServicios.Codigo, FactCatalogoServicios.Nombre as NombreProducto,  " + _
            "FactCatalogoServiciosHosp.PrecioUnitario , FactCatalogoServiciosHosp.Activo " + _
            " From " + _
            " (FactCatalogoServicios left join FactCatalogoServiciosHosp " + _
            " on FactCatalogoServicios.IdProducto = FactCatalogoServiciosHosp.IdProducto) " + _
            " left join FactPuntosCargaServicio " + _
            " on FactCatalogoServicios.IdServicioSubGrupo = FactPuntosCargaServicio.IdServicioSubGrupo " + _
            " where FactCatalogoServiciosHosp.Activo = 1 "
    If lIdPuntoCarga <> 1 And lIdPuntoCarga <> 99 Then             '<> 0 Then
        sSql = sSql + " and FactPuntosCargaServicio.IdPuntoCarga = " & lIdPuntoCarga
    End If
    
    If lIdTipoFinanciamiento <> 0 Then
        sSql = sSql + " and FactCatalogoServiciosHosp.IdTipoFinanciamiento = " & lIdTipoFinanciamiento
    End If
    
    If codigo <> "" Then
        sSql = sSql + " and FactCatalogoServicios.Codigo = '" & codigo & "'"
    End If
    oRecordset.Open sSql, mo_Conexion, adOpenDynamic, adLockOptimistic
 
    Set oRecordset.ActiveConnection = Nothing
 
    Set SeleccionarPorCodigoDEBB = oRecordset
    
    
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
    Trace Err.Number & " " + Err.Description, "SeleccionarPorCodigo"
    
Exit Function


End Function

Function SeleccionarPorCodigo(codigo As String, IdTipoFinanciamiento As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oRecordset As New Recordset
    Set SeleccionarPorCodigo = Nothing
    sSql = "select FactCatalogoServiciosHosp.IdProducto, "
    sSql = sSql + " FactCatalogoServicios.Nombre as NombreProducto, FactCatalogoServiciosHosp.PrecioUnitario"
    sSql = sSql + " From FactCatalogoServiciosHosp"
    sSql = sSql + " left join FactCatalogoServicios"
    sSql = sSql + " on FactCatalogoServiciosHosp.IdProducto = FactCatalogoServicios.IdProducto"
    sSql = sSql + " Where"
    sSql = sSql + " FactCatalogoServiciosHosp.IdTipoFinanciamiento = " & IdTipoFinanciamiento
    sSql = sSql + " and FactCatalogoServicios.Codigo = '" & codigo & "'"

    oRecordset.Open sSql, mo_Conexion, adOpenDynamic, adLockOptimistic
 
    Set oRecordset.ActiveConnection = Nothing
 
    Set SeleccionarPorCodigo = oRecordset
    
    
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
    Trace Err.Number & " " + Err.Description, "SeleccionarPorCodigo"
    
Exit Function

End Function




Function SeleccionarPorOrdenAtencion(IdOrden As Long, sEstadosFacturacion As String, sTiposFinanciamiento As String) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oRecordset As New Recordset
   
    Set SeleccionarPorOrdenAtencion = Nothing
   
    sSql = "select FacturacionServicios.IdFacturacionServicio as IdFacturacionProducto, FacturacionServicios.IdProducto,"
    sSql = sSql + " FacturacionServicios.Cantidad, FacturacionServicios.PrecioUnitario, FacturacionServicios.TotalPorPagar,"
    sSql = sSql + " FacturacionServicios.IdEstadoFacturacion, FacturacionServicios.IdPuntoCarga, FacturacionServicios.IdAtencion,"
    sSql = sSql + " FacturacionServicios.IdProducto, FacturacionServicios.IdUsuarioAutorizaSeguro, FacturacionServicios.FechaAutorizaSeguro,"
    sSql = sSql + " FacturacionServicios.IdUsuarioAutorizaPendiente, FacturacionServicios.FechaAutorizaPendiente, FacturacionServicios.IdServicioInternamiento,"
    sSql = sSql + " FacturacionServicios.IdUsuarioAutorizaDevolucion, FacturacionServicios.FechaAutorizaDevolucion, FacturacionServicios.IdComprobantePago, FacturacionServicios.IdComprobantePagoDevolucion, "
    sSql = sSql + " FactCatalogoServicios.codigo , case when Atenciones.IdTipoServicio = 1 then FactCatalogoServicios.Nombre + ' - ' + Servicios.Nombre else FactCatalogoServicios.Nombre end  as NombreProducto, FacturacionServicios.IdTipoFinanciamiento, FacturacionServicios.IdFuenteFinanciamiento, FacturacionServicios.IdOrden, FacturacionServicios.IdCajero, FacturacionServicios.FechaCajero,"
    sSql = sSql + " FacturacionServicios.CantidadPagar, FacturacionServicios.TotalPagar"
    sSql = sSql + " from ((FacturacionServicios"
    sSql = sSql + " left join FactCatalogoServicios on FacturacionServicios.IdProducto = FactCatalogoServicios.IdProducto)"
    sSql = sSql + " left join Atenciones on FacturacionServicios.IdAtencion = Atenciones.IdAtencion)"
    sSql = sSql + " left join Servicios on Servicios.IdServicio = Atenciones.IdServicioIngreso"
    sSql = sSql + " Where FacturacionServicios.IdOrden = " & IdOrden
    
    If sEstadosFacturacion <> "" Then
        sSql = sSql & " and FacturacionServicios.IdEstadoFacturacion  in (" & sEstadosFacturacion & ") "
    End If
   
    If sTiposFinanciamiento <> "" Then
        sSql = sSql & " and FacturacionServicios.IdTipoFinanciamiento in (" & sTiposFinanciamiento & ") "
    End If
   
    oRecordset.Open sSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    Set oRecordset.ActiveConnection = Nothing
    Set SeleccionarPorOrdenAtencion = oRecordset

    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
    Trace Err.Number & " " + Err.Description, "SeleccionarPorOrdenAtencion"
Exit Function

End Function

Function SeleccionarPorCuentaAtencion(IdCuentaAtencion As Long, sEstadosFacturacion As String, sTiposFinanciamiento As String, lIdAgrupador As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oRecordset As New Recordset
   
    Set SeleccionarPorCuentaAtencion = Nothing
   
    sSql = "select FacturacionServicios.IdFacturacionServicio as IdFacturacionProducto, FacturacionServicios.IdProducto,"
    sSql = sSql + " FactCatalogoServicios.codigo , case when Atenciones.IdTipoServicio = 1 then FactCatalogoServicios.Nombre + ' - ' + Servicios2.Nombre else FactCatalogoServicios.Nombre end  as NombreProducto, FacturacionServicios.Cantidad, FacturacionServicios.PrecioUnitario, FacturacionServicios.TotalPorPagar,"
    sSql = sSql + " FacturacionServicios.IdEstadoFacturacion, FacturacionServicios.IdPuntoCarga, FacturacionServicios.IdServicioInternamiento, Servicios.Nombre as NombreServicio, FacturacionServicios.IdAtencion, FacturacionServicios.IdOrden,"
    sSql = sSql + " FacturacionServicios.IdUsuarioAutorizaSeguro, FacturacionServicios.FechaAutorizaSeguro,"
    sSql = sSql + " FacturacionServicios.IdUsuarioAutorizaPendiente, FacturacionServicios.FechaAutorizaPendiente, "
    sSql = sSql + " FacturacionServicios.IdUsuarioAutorizaDevolucion, FacturacionServicios.FechaAutorizaDevolucion, FacturacionServicios.IdComprobantePago, (CajaComprobantesPago.NroSerie + ' - ' + CajaComprobantesPago.NroDocumento) as NroComprobante, FacturacionServicios.IdComprobantePagoDevolucion, "
    sSql = sSql + " FacturacionServicios.IdTipoFinanciamiento, FacturacionServicios.IdFuenteFinanciamiento"
    sSql = sSql + " from (((((FacturacionServicios"
    sSql = sSql + " left join FactCatalogoServicios on FacturacionServicios.IdProducto = FactCatalogoServicios.IdProducto)"
    sSql = sSql + " left join Atenciones on FacturacionServicios.IdAtencion = Atenciones.IdAtencion)"
    sSql = sSql + " left join FacturacionCuentasAtencion on FacturacionCuentasAtencion.IdCuentaAtencion = Atenciones.IdCuentaAtencion)"
    sSql = sSql + " left join Servicios on FacturacionServicios.IdServicioInternamiento = Servicios.IdServicio)"
    sSql = sSql + " left join CajaComprobantesPago on FacturacionServicios.IdComprobantePago = CajaComprobantesPago.IdComprobantePago)"
    sSql = sSql + " left join Servicios as Servicios2 on Servicios2.IdServicio = Atenciones.IdServicioIngreso"
    
    sSql = sSql + " Where FacturacionCuentasAtencion.IdCuentaAtencion = " & IdCuentaAtencion
    
    If sEstadosFacturacion <> "" Then
        sSql = sSql & " and (FacturacionServicios.IdEstadoFacturacion  in (" & sEstadosFacturacion & ") "
        sSql = sSql & " or (FacturacionServicios.IdEstadoFacturacion = 4 and FacturacionServicios.IdProducto = 4691))"
    End If
   
    If sTiposFinanciamiento <> "" Then
        sSql = sSql & " and FacturacionServicios.IdTipoFinanciamiento in (" & sTiposFinanciamiento & ") "
    End If
   
    Dim sSQL1 As String
    Dim sSQL2 As String
    
    Select Case lIdAgrupador
    Case 1
    Case 2
        sSQL1 = "select IdOrden from (FactOrdenesServicio"
        sSQL1 = sSQL1 + " left join Atenciones on FactOrdenesServicio.IdAtencion = Atenciones.IdAtencion)"
        sSQL1 = sSQL1 + " left join FacturacionCuentasAtencion on FacturacionCuentasAtencion.IdCuentaAtencion = Atenciones.IdCuentaAtencion"
        sSQL1 = sSQL1 + " Where FacturacionCuentasAtencion.IdCuentaAtencion = " & IdCuentaAtencion
        
        sSQL2 = sSql
        sSql = "SHAPE {" + sSQL1 + "} " + _
                "APPEND ({" + sSQL2 + "} " + _
                " AS title_chap RELATE IdOrden TO IdOrden) "
    Case 3
        sSQL1 = "select distinct TiposFinanciamiento.IdTipoFinanciamiento, TiposFinanciamiento.Descripcion from ((FacturacionServicios"
        sSQL1 = sSQL1 + " left join Atenciones on FacturacionServicios.IdAtencion = Atenciones.IdAtencion)"
        sSQL1 = sSQL1 + " left join FacturacionCuentasAtencion on FacturacionCuentasAtencion.IdCuentaAtencion = Atenciones.IdCuentaAtencion)"
        sSQL1 = sSQL1 + " left join TiposFinanciamiento on TiposFinanciamiento.IdTipoFinanciamiento = FacturacionServicios.IdTipoFinanciamiento"
        sSQL1 = sSQL1 + " Where FacturacionCuentasAtencion.IdCuentaAtencion = " & IdCuentaAtencion
        
        sSQL2 = sSql
        sSql = "SHAPE {" + sSQL1 + "} " + _
                "APPEND ({" + sSQL2 + "} " + _
                " AS title_chap RELATE IdTipoFinanciamiento TO IdTipoFinanciamiento) "
    Case 4
        sSQL1 = "select Atenciones.IdAtencion from Atenciones"
        sSQL1 = sSQL1 + " left join FacturacionCuentasAtencion on FacturacionCuentasAtencion.IdCuentaAtencion = Atenciones.IdCuentaAtencion"
        sSQL1 = sSQL1 + " Where FacturacionCuentasAtencion.IdCuentaAtencion = " & IdCuentaAtencion
        
        sSQL2 = sSql
        sSql = "SHAPE {" + sSQL1 + "} " + _
                "APPEND ({" + sSQL2 + "} " + _
                " AS title_chap RELATE IdAtencion TO IdAtencion) "
    Case 5
        sSQL1 = "select distinct FactPuntosCarga.IdPuntoCarga, FactPuntosCarga.Descripcion from ((FacturacionServicios"
        sSQL1 = sSQL1 + " left join Atenciones on FacturacionServicios.IdAtencion = Atenciones.IdAtencion)"
        sSQL1 = sSQL1 + " left join FacturacionCuentasAtencion on FacturacionCuentasAtencion.IdCuentaAtencion = Atenciones.IdCuentaAtencion)"
        sSQL1 = sSQL1 + " left join FactPuntosCarga on FactPuntosCarga.IdPuntoCarga = FacturacionServicios.IdPuntoCarga"
        sSQL1 = sSQL1 + " Where FacturacionCuentasAtencion.IdCuentaAtencion = " & IdCuentaAtencion
        
        sSQL2 = sSql
        sSql = "SHAPE {" + sSQL1 + "} " + _
                "APPEND ({" + sSQL2 + "} " + _
                " AS title_chap RELATE IdPuntoCarga TO IdPuntoCarga) "
    
    End Select
   
    
    oRecordset.Open sSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    
    Set oRecordset.ActiveConnection = Nothing
    Set SeleccionarPorCuentaAtencion = oRecordset

    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
    Trace Err.Number & " " + Err.Description, "SeleccionarPorCuentaAtencion"
Exit Function

End Function

Function SeleccionarDevolucionPorIdComprobante(lIdComprobanteDevolucion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oRecordset As New Recordset
   
    Set SeleccionarDevolucionPorIdComprobante = Nothing
   
    sSql = "select FacturacionServicios.IdFacturacionServicio as IdFacturacionProducto, FacturacionServicios.IdProducto,"
    sSql = sSql + " FacturacionServicios.Cantidad, FacturacionServicios.PrecioUnitario, FacturacionServicios.TotalPorPagar,"
    sSql = sSql + " FacturacionServicios.IdEstadoFacturacion, FacturacionServicios.IdPuntoCarga, FacturacionServicios.IdAtencion,"
    sSql = sSql + " FacturacionServicios.IdProducto, FacturacionServicios.IdUsuarioAutorizaSeguro, FacturacionServicios.FechaAutorizaSeguro,"
    sSql = sSql + " FacturacionServicios.IdUsuarioAutorizaPendiente, FacturacionServicios.FechaAutorizaPendiente, FacturacionServicios.IdServicioInternamiento,"
    sSql = sSql + " FacturacionServicios.IdUsuarioAutorizaDevolucion, FacturacionServicios.FechaAutorizaDevolucion, FacturacionServicios.IdComprobantePago, FacturacionServicios.IdComprobantePagoDevolucion, "
    sSql = sSql + " FactCatalogoServicios.codigo , FactCatalogoServicios.Nombre as NombreProducto, FacturacionServicios.IdTipoFinanciamiento, FacturacionServicios.IdFuenteFinanciamiento, FacturacionServicios.IdOrden, FacturacionServicios.IdCajero, FacturacionServicios.FechaCajero"
    sSql = sSql + " from (FacturacionServicios"
    sSql = sSql + " left join FactCatalogoServicios on FacturacionServicios.IdProducto = FactCatalogoServicios.IdProducto)"
    sSql = sSql + " left join Atenciones on FacturacionServicios.IdAtencion = Atenciones.IdAtencion"
    sSql = sSql + " Where FacturacionServicios.IdComprobantePagoDevolucion = " & lIdComprobanteDevolucion
    sSql = sSql & " and FacturacionServicios.IdEstadoFacturacion  = 6"
      
    oRecordset.Open sSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    Set oRecordset.ActiveConnection = Nothing
    Set SeleccionarDevolucionPorIdComprobante = oRecordset

    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
    Trace Err.Number & " " + Err.Description, "SeleccionarDevolucionPorIdComprobante"
Exit Function

End Function


Sub Trace(sMensaje As String, NombreDeMetodo As String)
    MsgBox "Mensaje: " & sMensaje + Chr(13) + "Metodo:" & NombreDeMetodo + Chr(13) + "Clase de Datos: FacturacionServicios", vbInformation, "FacturacionServicios"
End Sub

Function SeleccionarPorCuentaAtencionDEBB(IdCuentaAtencion As Long, sEstadosFacturacion As String, sTiposFinanciamiento As String, lIdAgrupador As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oRecordset As New Recordset
   
    Set SeleccionarPorCuentaAtencionDEBB = Nothing
   
    sSql = "select FacturacionServicios.IdFacturacionServicio as IdFacturacionProducto, FacturacionServicios.IdProducto,"
    sSql = sSql + " FactCatalogoServicios.codigo , case when Atenciones.IdTipoServicio = 1 then FactCatalogoServicios.Nombre + ' - ' + Servicios2.Nombre else FactCatalogoServicios.Nombre end  as NombreProducto, "
    sSql = sSql + " FacturacionServicios.CantidadPagar , FacturacionServicios.PrecioUnitario, FacturacionServicios.TotalPagar, "
    sSql = sSql + " FacturacionServicios.cantidadSIS, FacturacionServicios.precioSIS, FacturacionServicios.ImporteSIS,"
    sSql = sSql + " FacturacionServicios.cantidadSOAT, FacturacionServicios.precioSOAT, FacturacionServicios.ImporteSOAT,"
    sSql = sSql + " FacturacionServicios.ImporteEXO,"
    sSql = sSql + " FacturacionServicios.IdEstadoFacturacion, FacturacionServicios.IdPuntoCarga, FacturacionServicios.IdServicioInternamiento, Servicios.Nombre as NombreServicio, FacturacionServicios.IdAtencion, FacturacionServicios.IdOrden,"
    sSql = sSql + " FacturacionServicios.IdUsuarioAutorizaSeguro, FacturacionServicios.FechaAutorizaSeguro,"
    sSql = sSql + " FacturacionServicios.IdUsuarioAutorizaPendiente, FacturacionServicios.FechaAutorizaPendiente, "
    sSql = sSql + " FacturacionServicios.IdUsuarioAutorizaDevolucion, FacturacionServicios.FechaAutorizaDevolucion, FacturacionServicios.IdComprobantePago, (CajaComprobantesPago.NroSerie + ' - ' + CajaComprobantesPago.NroDocumento) as NroComprobante, FacturacionServicios.IdComprobantePagoDevolucion, "
    sSql = sSql + " FacturacionServicios.IdTipoFinanciamiento, FacturacionServicios.IdFuenteFinanciamiento,"
    sSql = sSql + " FacturacionServicios.cantidad,  FacturacionServicios.TotalPorPagar"
    sSql = sSql + " from (((((FacturacionServicios"
    sSql = sSql + " left join FactCatalogoServicios on FacturacionServicios.IdProducto = FactCatalogoServicios.IdProducto)"
    sSql = sSql + " left join Atenciones on FacturacionServicios.IdAtencion = Atenciones.IdAtencion)"
    sSql = sSql + " left join FacturacionCuentasAtencion on FacturacionCuentasAtencion.IdCuentaAtencion = Atenciones.IdCuentaAtencion)"
    sSql = sSql + " left join Servicios on FacturacionServicios.IdServicioInternamiento = Servicios.IdServicio)"
    sSql = sSql + " left join CajaComprobantesPago on FacturacionServicios.IdComprobantePago = CajaComprobantesPago.IdComprobantePago)"
    sSql = sSql + " left join Servicios as Servicios2 on Servicios2.IdServicio = Atenciones.IdServicioIngreso"
    
    sSql = sSql + " Where FacturacionCuentasAtencion.IdCuentaAtencion = " & IdCuentaAtencion
    
    If sEstadosFacturacion <> "" Then
        sSql = sSql & " and (FacturacionServicios.IdEstadoFacturacion  in (" & sEstadosFacturacion & ") "
        sSql = sSql & " or (FacturacionServicios.IdEstadoFacturacion = 4 and FacturacionServicios.IdProducto = 4691))"
    End If
   
    If sTiposFinanciamiento <> "" Then
        sSql = sSql & " and FacturacionServicios.IdTipoFinanciamiento in (" & sTiposFinanciamiento & ") "
    End If
   
    Dim sSQL1 As String
    Dim sSQL2 As String
    
    Select Case lIdAgrupador
    Case 1
    Case 2
        sSQL1 = "select IdOrden from (FactOrdenesServicio"
        sSQL1 = sSQL1 + " left join Atenciones on FactOrdenesServicio.IdAtencion = Atenciones.IdAtencion)"
        sSQL1 = sSQL1 + " left join FacturacionCuentasAtencion on FacturacionCuentasAtencion.IdCuentaAtencion = Atenciones.IdCuentaAtencion"
        sSQL1 = sSQL1 + " Where FacturacionCuentasAtencion.IdCuentaAtencion = " & IdCuentaAtencion
        
        sSQL2 = sSql
        sSql = "SHAPE {" + sSQL1 + "} " + _
                "APPEND ({" + sSQL2 + "} " + _
                " AS title_chap RELATE IdOrden TO IdOrden) "
    Case 3
        sSQL1 = "select distinct TiposFinanciamiento.IdTipoFinanciamiento, TiposFinanciamiento.Descripcion from ((FacturacionServicios"
        sSQL1 = sSQL1 + " left join Atenciones on FacturacionServicios.IdAtencion = Atenciones.IdAtencion)"
        sSQL1 = sSQL1 + " left join FacturacionCuentasAtencion on FacturacionCuentasAtencion.IdCuentaAtencion = Atenciones.IdCuentaAtencion)"
        sSQL1 = sSQL1 + " left join TiposFinanciamiento on TiposFinanciamiento.IdTipoFinanciamiento = FacturacionServicios.IdTipoFinanciamiento"
        sSQL1 = sSQL1 + " Where FacturacionCuentasAtencion.IdCuentaAtencion = " & IdCuentaAtencion
        
        sSQL2 = sSql
        sSql = "SHAPE {" + sSQL1 + "} " + _
                "APPEND ({" + sSQL2 + "} " + _
                " AS title_chap RELATE IdTipoFinanciamiento TO IdTipoFinanciamiento) "
    Case 4
        sSQL1 = "select Atenciones.IdAtencion from Atenciones"
        sSQL1 = sSQL1 + " left join FacturacionCuentasAtencion on FacturacionCuentasAtencion.IdCuentaAtencion = Atenciones.IdCuentaAtencion"
        sSQL1 = sSQL1 + " Where FacturacionCuentasAtencion.IdCuentaAtencion = " & IdCuentaAtencion
        
        sSQL2 = sSql
        sSql = "SHAPE {" + sSQL1 + "} " + _
                "APPEND ({" + sSQL2 + "} " + _
                " AS title_chap RELATE IdAtencion TO IdAtencion) "
    Case 5
        sSQL1 = "select distinct FactPuntosCarga.IdPuntoCarga, FactPuntosCarga.Descripcion from ((FacturacionServicios"
        sSQL1 = sSQL1 + " left join Atenciones on FacturacionServicios.IdAtencion = Atenciones.IdAtencion)"
        sSQL1 = sSQL1 + " left join FacturacionCuentasAtencion on FacturacionCuentasAtencion.IdCuentaAtencion = Atenciones.IdCuentaAtencion)"
        sSQL1 = sSQL1 + " left join FactPuntosCarga on FactPuntosCarga.IdPuntoCarga = FacturacionServicios.IdPuntoCarga"
        sSQL1 = sSQL1 + " Where FacturacionCuentasAtencion.IdCuentaAtencion = " & IdCuentaAtencion
        
        sSQL2 = sSql
        sSql = "SHAPE {" + sSQL1 + "} " + _
                "APPEND ({" + sSQL2 + "} " + _
                " AS title_chap RELATE IdPuntoCarga TO IdPuntoCarga) "
    
    End Select
   
    
    oRecordset.Open sSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    
    Set oRecordset.ActiveConnection = Nothing
    Set SeleccionarPorCuentaAtencionDEBB = oRecordset

    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
    Trace Err.Number & " " + Err.Description, "SeleccionarPorCuentaAtencion"
Exit Function

End Function


Function SeleccionarPorServicioYAtencion(lnIdProducto As Long, lnIdAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oRecordset As New Recordset
   
    Set SeleccionarPorServicioYAtencion = Nothing
    sSql = "select * from FacturacionServicios where IdAtencion=" & lnIdAtencion & " and idProducto=" & lnIdProducto
    oRecordset.Open sSql, mo_Conexion, adOpenKeyset, adLockOptimistic
'    Set oRecordset.ActiveConnection = Nothing
    Set SeleccionarPorServicioYAtencion = oRecordset
    ms_MensajeError = ""

Exit Function
ManejadorDeError:
    Trace Err.Number & " " + Err.Description, "SeleccionarPorCodigo"
    
Exit Function

End Function

