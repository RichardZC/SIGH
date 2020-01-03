VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturacionBienesInsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FacturacionBienesInsumos
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
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
'   Descripción:    Inserta un registro a la tabla FacturacionBienesInsumos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOFacturacionBienesInsumos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionBienesInsumosAgregar"
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaPendiente", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutorizaPendiente = 0, Null, oTabla.IdUsuarioAutorizaPendiente)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, IIf(oTabla.PrecioUnitario = 0, Null, oTabla.PrecioUnitario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, IIf(oTabla.Cantidad = 0, Null, oTabla.Cantidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoFacturacion", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoFacturacion = 0, Null, oTabla.IdEstadoFacturacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPartidaPresupuestal", adInteger, adParamInput, 0, IIf(oTabla.IdPartidaPresupuestal = 0, Null, oTabla.IdPartidaPresupuestal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCentroCosto", adInteger, adParamInput, 0, IIf(oTabla.IdCentroCosto = 0, Null, oTabla.IdCentroCosto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalPorPagar", adCurrency, adParamInput, 0, IIf(oTabla.TotalPorPagar = 0, Null, oTabla.TotalPorPagar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, IIf(oTabla.IdReceta = 0, Null, oTabla.IdReceta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFacturacionBienes", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaSeguro", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutorizaSeguro = 0, Null, oTabla.IdUsuarioAutorizaSeguro)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaDevolucion", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutorizaDevolucion = 0, Null, oTabla.IdUsuarioAutorizaDevolucion)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaSeguro", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaSeguro = 0, Null, oTabla.FechaAutorizaSeguro)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaPendiente", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaPendiente = 0, Null, oTabla.FechaAutorizaPendiente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaDevolucion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaDevolucion = 0, Null, oTabla.FechaAutorizaDevolucion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCajero", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaCajero = 0, Null, oTabla.FechaCajero)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePagoDevolucion", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePagoDevolucion = 0, Null, oTabla.IdComprobantePagoDevolucion)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.IdPuntoCarga = 0, Null, oTabla.IdPuntoCarga)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaModificacion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaModificacion = 0, Null, oTabla.FechaModificacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioCrea", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioCrea = 0, Null, oTabla.IdUsuarioCrea)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdusuarioModifica", adInteger, adParamInput, 0, IIf(oTabla.idUsuarioModifica = 0, Null, oTabla.idUsuarioModifica)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter

'***************daniel barrantes**************
'***************Nuevos campos
'***************
           Set oParameter = .CreateParameter("@ImporteSIS", adCurrency, adParamInput, 0, IIf(oTabla.ImporteSIS = 0, Null, oTabla.ImporteSIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImporteEXO", adCurrency, adParamInput, 0, IIf(oTabla.ImporteEXO = 0, Null, oTabla.ImporteEXO)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaEXO", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutorizaEXO = 0, Null, oTabla.IdUsuarioAutorizaEXO)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaEXO", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaEXO = 0, Null, oTabla.FechaAutorizaEXO)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImporteSOAT", adCurrency, adParamInput, 0, IIf(oTabla.ImporteSOAT = 0, Null, oTabla.ImporteSOAT)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalPagar", adCurrency, adParamInput, 0, IIf(oTabla.TotalPagar = 0, Null, oTabla.TotalPagar)): .Parameters.Append oParameter
           
           Set oParameter = .CreateParameter("@cantidadSIS", adInteger, adParamInput, 0, IIf(oTabla.cantidadSIS = 0, Null, oTabla.cantidadSIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@precioSIS", adCurrency, adParamInput, 0, IIf(oTabla.PrecioSIS = 0, Null, oTabla.PrecioSIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@cantidadSOAT", adInteger, adParamInput, 0, IIf(oTabla.cantidadSOAT = 0, Null, oTabla.cantidadSOAT)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@precioSOAT", adCurrency, adParamInput, 0, IIf(oTabla.PrecioSOAT = 0, Null, oTabla.PrecioSOAT)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@cantidadPagar", adInteger, adParamInput, 0, IIf(oTabla.CantidadPagar = 0, Null, oTabla.CantidadPagar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@cantidadDev", adInteger, adParamInput, 0, IIf(oTabla.cantidadDev = 0, Null, oTabla.cantidadDev)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaConv", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaConv = 0, Null, oTabla.FechaAutorizaConv)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@cantidadConv", adInteger, adParamInput, 0, IIf(oTabla.cantidadConv = 0, Null, oTabla.cantidadConv)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@precConv", adCurrency, adParamInput, 0, IIf(oTabla.precConv = 0, Null, oTabla.precConv)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImporteConv", adCurrency, adParamInput, 0, IIf(oTabla.ImporteConv = 0, Null, oTabla.ImporteConv)): .Parameters.Append oParameter
       
       
       .Execute
           oTabla.IdFacturacionBienes = .Parameters("@IdFacturacionBienes")
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
'   Descripción:    modifica un registro a la tabla FacturacionBienesInsumos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFacturacionBienesInsumos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionBienesInsumosModificar"
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaPendiente", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutorizaPendiente = 0, Null, oTabla.IdUsuarioAutorizaPendiente)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, IIf(oTabla.PrecioUnitario = 0, Null, oTabla.PrecioUnitario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, IIf(oTabla.Cantidad = 0, Null, oTabla.Cantidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoFacturacion", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoFacturacion = 0, Null, oTabla.IdEstadoFacturacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPartidaPresupuestal", adInteger, adParamInput, 0, IIf(oTabla.IdPartidaPresupuestal = 0, Null, oTabla.IdPartidaPresupuestal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCentroCosto", adInteger, adParamInput, 0, IIf(oTabla.IdCentroCosto = 0, Null, oTabla.IdCentroCosto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalPorPagar", adCurrency, adParamInput, 0, IIf(oTabla.TotalPorPagar = 0, Null, oTabla.TotalPorPagar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, IIf(oTabla.IdReceta = 0, Null, oTabla.IdReceta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFacturacionBienes", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionBienes = 0, Null, oTabla.IdFacturacionBienes)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaSeguro", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutorizaSeguro = 0, Null, oTabla.IdUsuarioAutorizaSeguro)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaDevolucion", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutorizaDevolucion = 0, Null, oTabla.IdUsuarioAutorizaDevolucion)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaSeguro", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaSeguro = 0, Null, oTabla.FechaAutorizaSeguro)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaPendiente", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaPendiente = 0, Null, oTabla.FechaAutorizaPendiente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaDevolucion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaDevolucion = 0, Null, oTabla.FechaAutorizaDevolucion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCajero", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaCajero = 0, Null, oTabla.FechaCajero)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePagoDevolucion", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePagoDevolucion = 0, Null, oTabla.IdComprobantePagoDevolucion)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.IdPuntoCarga = 0, Null, oTabla.IdPuntoCarga)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaModificacion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaModificacion = 0, Null, oTabla.FechaModificacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioCrea", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioCrea = 0, Null, oTabla.IdUsuarioCrea)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdusuarioModifica", adInteger, adParamInput, 0, IIf(oTabla.idUsuarioModifica = 0, Null, oTabla.idUsuarioModifica)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       
'***************daniel barrantes**************
'***************nuevos campos
'***************
           Set oParameter = .CreateParameter("@ImporteSIS", adCurrency, adParamInput, 0, IIf(oTabla.ImporteSIS = 0, Null, oTabla.ImporteSIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImporteEXO", adCurrency, adParamInput, 0, IIf(oTabla.ImporteEXO = 0, Null, oTabla.ImporteEXO)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaEXO", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutorizaEXO = 0, Null, oTabla.IdUsuarioAutorizaEXO)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaEXO", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaEXO = 0, Null, oTabla.FechaAutorizaEXO)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImporteSOAT", adCurrency, adParamInput, 0, IIf(oTabla.ImporteSOAT = 0, Null, oTabla.ImporteSOAT)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalPagar", adCurrency, adParamInput, 0, IIf(oTabla.TotalPagar = 0, Null, oTabla.TotalPagar)): .Parameters.Append oParameter
           
           Set oParameter = .CreateParameter("@cantidadSIS", adInteger, adParamInput, 0, IIf(oTabla.cantidadSIS = 0, Null, oTabla.cantidadSIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@precioSIS", adCurrency, adParamInput, 0, IIf(oTabla.PrecioSIS = 0, Null, oTabla.PrecioSIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@cantidadSOAT", adInteger, adParamInput, 0, IIf(oTabla.cantidadSOAT = 0, Null, oTabla.cantidadSOAT)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@precioSOAT", adCurrency, adParamInput, 0, IIf(oTabla.PrecioSOAT = 0, Null, oTabla.PrecioSOAT)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@cantidadPagar", adInteger, adParamInput, 0, IIf(oTabla.CantidadPagar = 0, Null, oTabla.CantidadPagar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@cantidadDev", adInteger, adParamInput, 0, IIf(oTabla.cantidadDev = 0, Null, oTabla.cantidadDev)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaConv", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaConv = 0, Null, oTabla.FechaAutorizaConv)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@cantidadConv", adInteger, adParamInput, 0, IIf(oTabla.cantidadConv = 0, Null, oTabla.cantidadConv)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@precConv", adCurrency, adParamInput, 0, IIf(oTabla.precConv = 0, Null, oTabla.precConv)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImporteConv", adCurrency, adParamInput, 0, IIf(oTabla.ImporteConv = 0, Null, oTabla.ImporteConv)): .Parameters.Append oParameter
       
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
'   Descripción:    elimia o borra un registro a la tabla FacturacionBienesInsumos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFacturacionBienesInsumos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionBienesInsumosEliminar"
           Set oParameter = .CreateParameter("@IdFacturacionBienes", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionBienes = 0, Null, oTabla.IdFacturacionBienes)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla FacturacionBienesInsumos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFacturacionBienesInsumos) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FacturacionBienesInsumosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdFacturacionBienes", adInteger, adParamInput, 0, oTabla.IdFacturacionBienes): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdUsuarioAutorizaPendiente = IIf(IsNull(oRecordset!IdUsuarioAutorizaPendiente), 0, oRecordset!IdUsuarioAutorizaPendiente)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.PrecioUnitario = IIf(IsNull(oRecordset!PrecioUnitario), 0, oRecordset!PrecioUnitario)
       oTabla.Cantidad = IIf(IsNull(oRecordset!Cantidad), 0, oRecordset!Cantidad)
       oTabla.IdEstadoFacturacion = IIf(IsNull(oRecordset!IdEstadoFacturacion), 0, oRecordset!IdEstadoFacturacion)
       oTabla.IdPartidaPresupuestal = IIf(IsNull(oRecordset!IdPartidaPresupuestal), 0, oRecordset!IdPartidaPresupuestal)
       oTabla.IdCentroCosto = IIf(IsNull(oRecordset!IdCentroCosto), 0, oRecordset!IdCentroCosto)
       oTabla.idTipoFinanciamiento = IIf(IsNull(oRecordset!idTipoFinanciamiento), 0, oRecordset!idTipoFinanciamiento)
       oTabla.idFuenteFinanciamiento = IIf(IsNull(oRecordset!idFuenteFinanciamiento), 0, oRecordset!idFuenteFinanciamiento)
       oTabla.TotalPorPagar = IIf(IsNull(oRecordset!TotalPorPagar), 0, oRecordset!TotalPorPagar)
       oTabla.IdReceta = IIf(IsNull(oRecordset!IdReceta), 0, oRecordset!IdReceta)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.IdFacturacionBienes = IIf(IsNull(oRecordset!IdFacturacionBienes), 0, oRecordset!IdFacturacionBienes)
       oTabla.IdUsuarioAutorizaSeguro = IIf(IsNull(oRecordset!IdUsuarioAutorizaSeguro), 0, oRecordset!IdUsuarioAutorizaSeguro)
       oTabla.IdOrden = IIf(IsNull(oRecordset!IdOrden), 0, oRecordset!IdOrden)
       oTabla.IdUsuarioAutorizaDevolucion = IIf(IsNull(oRecordset!IdUsuarioAutorizaDevolucion), 0, oRecordset!IdUsuarioAutorizaDevolucion)
       oTabla.IdCajero = IIf(IsNull(oRecordset!IdCajero), 0, oRecordset!IdCajero)
       oTabla.FechaAutorizaSeguro = IIf(IsNull(oRecordset!FechaAutorizaSeguro), 0, oRecordset!FechaAutorizaSeguro)
       oTabla.FechaAutorizaPendiente = IIf(IsNull(oRecordset!FechaAutorizaPendiente), 0, oRecordset!FechaAutorizaPendiente)
       oTabla.FechaAutorizaDevolucion = IIf(IsNull(oRecordset!FechaAutorizaDevolucion), 0, oRecordset!FechaAutorizaDevolucion)
       oTabla.FechaCajero = IIf(IsNull(oRecordset!FechaCajero), 0, oRecordset!FechaCajero)
       oTabla.IdComprobantePagoDevolucion = IIf(IsNull(oRecordset!IdComprobantePagoDevolucion), 0, oRecordset!IdComprobantePagoDevolucion)
       oTabla.IdPuntoCarga = IIf(IsNull(oRecordset!IdPuntoCarga), 0, oRecordset!IdPuntoCarga)
       oTabla.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
       oTabla.FechaModificacion = IIf(IsNull(oRecordset!FechaModificacion), 0, oRecordset!FechaModificacion)
       oTabla.IdUsuarioCrea = IIf(IsNull(oRecordset!IdUsuarioCrea), 0, oRecordset!IdUsuarioCrea)
       oTabla.idUsuarioModifica = IIf(IsNull(oRecordset!idUsuarioModifica), 0, oRecordset!idUsuarioModifica)
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       
'***************daniel barrantes**************
'***************nuevos campos
'***************
       oTabla.ImporteSIS = IIf(IsNull(oRecordset!ImporteSIS), 0, oRecordset!ImporteSIS)
       oTabla.IdUsuarioAutorizaEXO = IIf(IsNull(oRecordset!IdUsuarioAutorizaEXO), 0, oRecordset!IdUsuarioAutorizaEXO)
       oTabla.FechaAutorizaEXO = IIf(IsNull(oRecordset!FechaAutorizaEXO), 0, oRecordset!FechaAutorizaEXO)
       oTabla.ImporteEXO = IIf(IsNull(oRecordset!ImporteEXO), 0, oRecordset!ImporteEXO)
       oTabla.ImporteSOAT = IIf(IsNull(oRecordset!ImporteSOAT), 0, oRecordset!ImporteSOAT)
       oTabla.TotalPagar = IIf(IsNull(oRecordset!TotalPagar), 0, oRecordset!TotalPagar)
       oTabla.cantidadSIS = IIf(IsNull(oRecordset!cantidadSIS), 0, oRecordset!cantidadSIS)
       oTabla.cantidadSOAT = IIf(IsNull(oRecordset!cantidadSOAT), 0, oRecordset!cantidadSOAT)
       oTabla.CantidadPagar = IIf(IsNull(oRecordset!CantidadPagar), 0, oRecordset!CantidadPagar)
       oTabla.PrecioSIS = IIf(IsNull(oRecordset!PrecioSIS), 0, oRecordset!PrecioSIS)
       oTabla.PrecioSOAT = IIf(IsNull(oRecordset!PrecioSOAT), 0, oRecordset!PrecioSOAT)
       oTabla.cantidadDev = IIf(IsNull(oRecordset!cantidadDev), 0, oRecordset!cantidadDev)
       oTabla.FechaAutorizaConv = IIf(IsNull(oRecordset!FechaAutorizaConv), 0, oRecordset!FechaAutorizaConv)
       oTabla.cantidadConv = IIf(IsNull(oRecordset!cantidadConv), 0, oRecordset!cantidadConv)
       oTabla.precConv = IIf(IsNull(oRecordset!precConv), 0, oRecordset!precConv)
       oTabla.ImporteConv = IIf(IsNull(oRecordset!ImporteConv), 0, oRecordset!ImporteConv)
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









Function SeleccionarParaPendientePago(lIdCuentaAtencion As Long) As Recordset


End Function


Function SeleccionarParaEstadoCuenta(lIdCuentaAtencion As Long) As Recordset


End Function

Function SeleccionarParaCaja(lIdCuentaAtencion As Long) As Recordset


End Function


Function EliminarPorIdAtencion(lIdAtencion As Long) As Boolean

End Function



'EFGL

Function SeleccionarPorCuentaAtencionParaEstadoCuenta(IdCuentaAtencion As Long, IdPuntosDeCarga As Long, lEstadosFacturacion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oRecordset As New Recordset
   
    Set SeleccionarPorCuentaAtencionParaEstadoCuenta = Nothing
   
    sSql = " select a.idfacturacionbienes as idfacturacionbienes,a.idProducto, b.codigo as Codigo,b.nombre as Descripcion"
    sSql = sSql & " ,a.idTipoFinanciamiento,'' as Poliza,c.descripcion as TipoFinanciamiento,Cantidad,"
    sSql = sSql & " PrecioUnitario , Totalporpagar as SubTotal,a.idEstadoFacturacion,d.descripcion as Estado, a.IdPuntoCarga, a.IdAtencion, "
    sSql = sSql & " FechaAutorizaPendiente, FechaAutorizaSeguro, IdEmpAutorizaPendiente, IdEmpAutorizaSeguro, IdFuenteFinanciamiento, 0 as IdUsuarioAuditoria "
    sSql = sSql & " From facturacionbienesinsumos a, factcatalogobienesinsumos b,tiposFinanciamiento c, estadosfacturacion d"
    sSql = sSql & " where a.idproducto = b.idproducto and c.idTipoFinanciamiento = a.idTipoFinanciamiento"
    sSql = sSql & " and a.idEstadoFacturacion = d.idEstadoFacturacion"
    sSql = sSql & " and idatencion in( select idatencion from atenciones"
    sSql = sSql & " where idcuentaatencion = " & IdCuentaAtencion & ")"
    
    If IdPuntosDeCarga > 0 Then
        sSql = sSql & " and a.IdPuntoCarga = " & IdPuntosDeCarga
    End If
    
    If lEstadosFacturacion <> 0 Then
        sSql = sSql & " and d.IdEstadoFacturacion = " & lEstadosFacturacion & " "
    End If
    
    oRecordset.Open sSql, mo_Conexion, adOpenDynamic, adLockOptimistic
    
    Set oRecordset.ActiveConnection = Nothing
 
    Set SeleccionarPorCuentaAtencionParaEstadoCuenta = oRecordset
    
    
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function

'***************daniel barrantes**************
'***************Filtro del Catalogo de Medicamentos Farmacia
'***************
Function SeleccionarPorCodigoDEBB(codigo As String, lIdTipoFinanciamiento As Long, lIdPuntoCarga As Long, _
                                  Optional lbSoloItemsDePaquete As Boolean) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    If lIdPuntoCarga <> 1 And lIdPuntoCarga <> 99 Then             '<> 0 Then
        sSql = sSql + " and FactPuntosCargaBienesInsumos.IdPuntoCarga = " & lIdPuntoCarga
    End If
    If lIdTipoFinanciamiento <> 0 Then
        sSql = sSql + " and FactCatalogoBienesInsumosHosp.IdTipoFinanciamiento = " & lIdTipoFinanciamiento
    End If
    If codigo <> "" Then
        sSql = sSql + " and FactCatalogoBienesInsumos.Codigo = '" & codigo & "'"
    End If
    If lbSoloItemsDePaquete = True Then
        sSql = sSql + " and FactCatalogoBienesInsumos.esPaquete=1"
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoBienesInsumosSeleccionarPorCodigo"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorCodigoDEBB = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description

End Function

Function SeleccionarPorCodigo(codigo As String, idTipoFinanciamiento As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoBienesInsumosXcodigoYtipofinanciamiento"
        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, idTipoFinanciamiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 7, codigo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorCodigo = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description

End Function

'------------------------------------------------------------------------------------
'   Función:        InsertarBienSP
'   Descripción:    Inserta un registro a la tabla FacturacionServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function InsertarBienInsumoSP(ByVal oTabla As DOFacturacionBienesInsumos, ByRef sResultado As String) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   InsertarBienInsumoSP = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionBienesAgregarV2"
        Set oParameter = .CreateParameter("@IdFacturacionBienes", adInteger, adParamOutput, 0): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@result", adVarChar, adParamOutput, 500): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, IIf(oTabla.Cantidad = 0, 0, oTabla.Cantidad)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdEstadoFacturacion", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoFacturacion = 0, Null, oTabla.IdEstadoFacturacion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento)): .Parameters.Append oParameter
        'Set oParameter = .CreateParameter("@IdServicioInternamiento", adInteger, adParamInput, 0, IIf(oTabla.IdServicioInternamiento = 0, Null, oTabla.IdServicioInternamiento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.IdPuntoCarga = 0, Null, oTabla.IdPuntoCarga)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAutorizaSeguro", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutorizaSeguro = 0, Null, oTabla.IdUsuarioAutorizaSeguro)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAutorizaPendiente", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutorizaPendiente = 0, Null, oTabla.IdUsuarioAutorizaPendiente)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAutorizaDevolucion", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutorizaDevolucion = 0, Null, oTabla.IdUsuarioAutorizaDevolucion)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaAutorizaSeguro", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaSeguro = 0, Null, oTabla.FechaAutorizaSeguro)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaAutorizaPendiente", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaPendiente = 0, Null, oTabla.FechaAutorizaPendiente)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaAutorizaDevolucion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaDevolucion = 0, Null, oTabla.FechaAutorizaDevolucion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaCajero", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaCajero = 0, Null, oTabla.FechaCajero)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAuditoria = 0, Null, oTabla.IdUsuarioAuditoria)): .Parameters.Append oParameter
'***************daniel barrantes**************
'***************Nuevos campos
'***************
           Set oParameter = .CreateParameter("@ImporteSIS", adCurrency, adParamInput, 0, IIf(oTabla.ImporteSIS = 0, Null, oTabla.ImporteSIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImporteEXO", adCurrency, adParamInput, 0, IIf(oTabla.ImporteEXO = 0, Null, oTabla.ImporteEXO)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaEXO", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutorizaEXO = 0, Null, oTabla.IdUsuarioAutorizaEXO)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaEXO", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaEXO = 0, Null, oTabla.FechaAutorizaEXO)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImporteSOAT", adCurrency, adParamInput, 0, IIf(oTabla.ImporteSOAT = 0, Null, oTabla.ImporteSOAT)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalPagar", adCurrency, adParamInput, 0, IIf(oTabla.TotalPagar = 0, Null, oTabla.TotalPagar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@cantidadSIS", adInteger, adParamInput, 0, IIf(oTabla.cantidadSIS = 0, 0, oTabla.cantidadSIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@precioSIS", adCurrency, adParamInput, 0, IIf(oTabla.PrecioSIS = 0, Null, oTabla.PrecioSIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@cantidadSOAT", adInteger, adParamInput, 0, IIf(oTabla.cantidadSOAT = 0, 0, oTabla.cantidadSOAT)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@precioSOAT", adCurrency, adParamInput, 0, IIf(oTabla.PrecioSOAT = 0, Null, oTabla.PrecioSOAT)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@cantidadPagar", adInteger, adParamInput, 0, IIf(oTabla.CantidadPagar = 0, Null, oTabla.CantidadPagar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@cantidadDev", adInteger, adParamInput, 0, IIf(oTabla.cantidadDev = 0, 0, oTabla.cantidadDev)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaConv", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaConv = 0, Null, oTabla.FechaAutorizaConv)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@cantidadConv", adInteger, adParamInput, 0, IIf(oTabla.cantidadConv = 0, 0, oTabla.cantidadConv)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@precConv", adCurrency, adParamInput, 0, IIf(oTabla.precConv = 0, Null, oTabla.precConv)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImporteConv", adCurrency, adParamInput, 0, IIf(oTabla.ImporteConv = 0, Null, oTabla.ImporteConv)): .Parameters.Append oParameter
        
        
       .Execute
        If .Parameters("@result") <> "" Then
            sResultado = .Parameters("@result")
            Exit Function
        End If
        oTabla.IdFacturacionBienes = .Parameters("@IdFacturacionBienes")
        
   End With
 
   InsertarBienInsumoSP = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "InsertarBienInsumoSP"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        ModificarBienInsumoSP
'   Descripción:    Inserta un registro a la tabla FacturacionServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function ModificarBienInsumoSP(ByVal oTabla As DOFacturacionBienesInsumos, ByRef sResultado As String) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   ModificarBienInsumoSP = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionBienesModificarV2"
        Set oParameter = .CreateParameter("@IdFacturacionBienes", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionBienes = 0, Null, oTabla.IdFacturacionBienes)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@result", adVarChar, adParamOutput, 500): .Parameters.Append oParameter
        'Set oParameter = .CreateParameter("@IdUsuarioModifica", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioModifica = 0, Null, oTabla.IdUsuarioModifica)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, IIf(oTabla.Cantidad = 0, 0, oTabla.Cantidad)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdEstadoFacturacion", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoFacturacion = 0, Null, oTabla.IdEstadoFacturacion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento)): .Parameters.Append oParameter
        'Set oParameter = .CreateParameter("@IdServicioInternamiento", adInteger, adParamInput, 0, IIf(oTabla.IdServicioInternamiento = 0, Null, oTabla.IdServicioInternamiento)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
        'Set oParameter = .CreateParameter("@TotalPorPagar", adCurrency, adParamInput, 0, IIf(oTabla.TotalPorPagar = 0, Null, oTabla.TotalPorPagar)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.IdPuntoCarga = 0, Null, oTabla.IdPuntoCarga)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        'Set oParameter = .CreateParameter("@FechaModificacion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaModificacion = 0, Null, oTabla.FechaModificacion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAutorizaSeguro", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutorizaSeguro = 0, Null, oTabla.IdUsuarioAutorizaSeguro)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAutorizaPendiente", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutorizaPendiente = 0, Null, oTabla.IdUsuarioAutorizaPendiente)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAutorizaDevolucion", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutorizaDevolucion = 0, Null, oTabla.IdUsuarioAutorizaDevolucion)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaAutorizaSeguro", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaSeguro = 0, Null, oTabla.FechaAutorizaSeguro)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaAutorizaPendiente", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaPendiente = 0, Null, oTabla.FechaAutorizaPendiente)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaAutorizaDevolucion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaDevolucion = 0, Null, oTabla.FechaAutorizaDevolucion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaCajero", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaCajero = 0, Null, oTabla.FechaCajero)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdComprobantePagoDevolucion", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePagoDevolucion = 0, Null, oTabla.IdComprobantePagoDevolucion)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAuditoria = 0, Null, oTabla.IdUsuarioAuditoria)): .Parameters.Append oParameter
'***************daniel barrantes**************
'***************Nuevos campos
'***************
           Set oParameter = .CreateParameter("@ImporteSIS", adCurrency, adParamInput, 0, IIf(oTabla.ImporteSIS = 0, Null, oTabla.ImporteSIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImporteEXO", adCurrency, adParamInput, 0, IIf(oTabla.ImporteEXO = 0, Null, oTabla.ImporteEXO)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAutorizaEXO", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutorizaEXO = 0, Null, oTabla.IdUsuarioAutorizaEXO)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaEXO", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaEXO = 0, Null, oTabla.FechaAutorizaEXO)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImporteSOAT", adCurrency, adParamInput, 0, IIf(oTabla.ImporteSOAT = 0, Null, oTabla.ImporteSOAT)): .Parameters.Append oParameter
          
           Set oParameter = .CreateParameter("@cantidadSIS", adInteger, adParamInput, 0, IIf(oTabla.cantidadSIS = 0, 0, oTabla.cantidadSIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@precioSIS", adCurrency, adParamInput, 0, IIf(oTabla.PrecioSIS = 0, Null, oTabla.PrecioSIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@cantidadSOAT", adInteger, adParamInput, 0, IIf(oTabla.cantidadSOAT = 0, 0, oTabla.cantidadSOAT)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@precioSOAT", adCurrency, adParamInput, 0, IIf(oTabla.PrecioSOAT = 0, Null, oTabla.PrecioSOAT)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@cantidadDev", adInteger, adParamInput, 0, IIf(oTabla.cantidadDev = 0, 0, oTabla.cantidadDev)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAutorizaConv", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaAutorizaConv = 0, Null, oTabla.FechaAutorizaConv)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@cantidadConv", adInteger, adParamInput, 0, IIf(oTabla.cantidadConv = 0, 0, oTabla.cantidadConv)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@precConv", adCurrency, adParamInput, 0, IIf(oTabla.precConv = 0, Null, oTabla.precConv)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImporteConv", adCurrency, adParamInput, 0, IIf(oTabla.ImporteConv = 0, Null, oTabla.ImporteConv)): .Parameters.Append oParameter
        
        
       
       .Execute
        If .Parameters("@result") <> "" Then
            sResultado = .Parameters("@result")
            Exit Function
        End If
        
        'oTabla.IdFacturacionServicio = .Parameters("@IdFacturacionServicio")
   End With
 
   ModificarBienInsumoSP = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "ModificarBienInsumoSP"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        EliminarBienSP
'   Descripción:    Inserta un registro a la tabla FacturacionServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function EliminarBienSP(ByVal oTabla As DOFacturacionBienesInsumos, ByRef sResultado As String) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   EliminarBienSP = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionBienesEliminarBien"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
 
   EliminarBienSP = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function SeleccionarPorOrdenAtencion(IdOrden As Long, sEstadosFacturacion As String, sTiposFinanciamiento As String) As Recordset


End Function

Function SeleccionarPorCuentaAtencion(IdCuentaAtencion As Long, sEstadosFacturacion As String, sTiposFinanciamiento As String, lIdAgrupador As Long) As Recordset

End Function


'***************daniel barrantes**************
'***************Filtrar Medicamentos de una Cuenta de  Atencion
'***************
Function SeleccionarPorCuentaAtencionDEBB(IdCuentaAtencion As Long, sEstadosFacturacion As String, sTiposFinanciamiento As String, lIdAgrupador As Long) As Recordset

End Function



Function SeleccionarDevolucionPorIdComprobante(lIdComprobanteDevolucion As Long) As Recordset


End Function

Sub Trace(sMensaje As String, NombreDeMetodo As String)
    MsgBox "Mensaje: " & sMensaje + Chr(13) + "Metodo:" & NombreDeMetodo + Chr(13) + "Clase de Datos: FacturacionServicios", vbInformation, "FacturacionBienesInsumos"
End Sub


