VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CajaCompDetalleInsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'WCG-05 [Todo el archivo]

Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: CajaComprobantesDetalleBienesInsumos
'        Autor: William Castro Grijalva
'        Fecha: 10/08/2005 07:43:00
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
'   Descripción:    Inserta un registro a la tabla CajaComprobantesDetalleBienesInsumos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCajaCompDetalleInsumos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesDetalleBienesInsumosAgregar"
           Set oParameter = .CreateParameter("@SubTotalPagado", adCurrency, adParamInput, 0, oTabla.SubTotalPagado): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, oTabla.Cantidad): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, oTabla.PrecioUnitario): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobanteDetalleBienes", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFacturacionBienes", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionBienes = 0, Null, oTabla.IdFacturacionBienes)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsPagoACuenta", adInteger, adParamInput, 0, oTabla.EsPagoACuenta): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.IdProducto = 0, Null, oTabla.IdProducto)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdComprobanteDetalleBienes = .Parameters("@IdComprobanteDetalleBienes")
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
'   Descripción:    modifica un registro a la tabla CajaComprobantesDetalleBienesInsumos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCajaCompDetalleInsumos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesDetalleBienesInsumosModificar"
           Set oParameter = .CreateParameter("@SubTotalPagado", adCurrency, adParamInput, 0, oTabla.SubTotalPagado): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, oTabla.Cantidad): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, oTabla.PrecioUnitario): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobanteDetalleBienes", adInteger, adParamInput, 0, IIf(oTabla.IdComprobanteDetalleBienes = 0, Null, oTabla.IdComprobanteDetalleBienes)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFacturacionBienes", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionBienes = 0, Null, oTabla.IdFacturacionBienes)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsPagoACuenta", adInteger, adParamInput, 0, oTabla.EsPagoACuenta): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.IdProducto = 0, Null, oTabla.IdProducto)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CajaComprobantesDetalleBienesInsumos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCajaCompDetalleInsumos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesDetalleBienesInsumosEliminar"
           Set oParameter = .CreateParameter("@IdComprobanteDetalleBienes", adInteger, adParamInput, 0, IIf(oTabla.IdComprobanteDetalleBienes = 0, Null, oTabla.IdComprobanteDetalleBienes)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla CajaComprobantesDetalleBienesInsumos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCajaCompDetalleInsumos) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CajaComprobantesDetalleBienesInsumosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdComprobanteDetalleBienes", adInteger, adParamInput, 0, oTabla.IdComprobanteDetalleBienes): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.SubTotalPagado = oRecordset!SubTotalPagado
       oTabla.Cantidad = oRecordset!Cantidad
       oTabla.PrecioUnitario = oRecordset!PrecioUnitario
       oTabla.IdComprobanteDetalleBienes = IIf(IsNull(oRecordset!IdComprobanteDetalleBienes), 0, oRecordset!IdComprobanteDetalleBienes)
       oTabla.IdFacturacionBienes = IIf(IsNull(oRecordset!IdFacturacionBienes), 0, oRecordset!IdFacturacionBienes)
       oTabla.EsPagoACuenta = oRecordset!EsPagoACuenta
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.IdProducto = IIf(IsNull(oRecordset!IdProducto), 0, oRecordset!IdProducto)
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


Function EliminarPorComprobante(ByVal oTabla As DOCajaComprobantesPago, IdEstadoFacturacionPendiente As Integer) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim SQL As String
 
   EliminarPorComprobante = False
   
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
   End With
   'Eliminamos las relaciones que tiene con los items de la cuenta
   
    SQL = "Update FacturacionBienesInsumos set SubTotalPorPagar = SubTotalPorPagar + ds.SubTotalPagado " _
            & " ,SubTotalPagadoACuenta = case ds.EsPagoACuenta when 1 then fs.SubTotalPagadoACuenta - ds.SubTotalPagado else fs.SubTotalPagadoACuenta end " _
            & " ,SubTotalPagado = fs.SubTotalPagado - ds.SubTotalPagado " _
            & " ,IdEstadoFacturacion = " & sghEstadoFacturacion.sghPendientePago _
            & " from  FacturacionBienesInsumos fs, CajaComprobantesDetalleBienesInsumos ds " _
            & " Where fs.IdFacturacionBienes = ds.IdFacturacionBienes " _
            & " and IdComprobantePago = " & oTabla.IdComprobantePago
   
   oCommand.CommandText = SQL
   oCommand.Execute
      
   'SQL = "Delete from CajaComprobantesDetalleBienesInsumos where IdComprobantePago = " & oTabla.IdComprobantePago
   SQL = "Update CajaComprobantesDetalleBienesInsumos set IdFacturacionBienes = null where IdComprobantePago = " & oTabla.IdComprobantePago
   oCommand.CommandText = SQL
   oCommand.Execute
 
   EliminarPorComprobante = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


