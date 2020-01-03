VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CajaComprobantesDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'WCG Ini 02/06/2005 [Todo el archivo]

Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: CajaComprobantesDetalle
'        Autor: William Castro Grijalva
'        Fecha: 03/06/2005 10:24:58
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
Function Insertar(ByVal oTabla As DOCajaComprobantesDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesDetalleBienesInsumosAgregar"
           Set oParameter = .CreateParameter("@SubTotalPagado", adCurrency, adParamInput, 0, IIf(oTabla.SubTotalPagado = "", Null, oTabla.SubTotalPagado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Cantidad", adCurrency, adParamInput, 0, IIf(oTabla.Cantidad = "", Null, oTabla.Cantidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, IIf(oTabla.PrecioUnitario = "", Null, oTabla.PrecioUnitario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobanteDetalleBienes", adInteger, adParamInput, 0, IIf(oTabla.IdComprobanteDetalle = "", Null, oTabla.IdComprobanteDetalle)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = "", Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.IdProducto = "", Null, oTabla.IdProducto)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
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
'   Descripción:    modifica un registro a la tabla CajaComprobantesDetalleBienesInsumos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCajaComprobantesDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesDetalleBienesInsumosModificar"
           Set oParameter = .CreateParameter("@SubTotalPagado", adCurrency, adParamInput, 0, IIf(oTabla.SubTotalPagado = "", Null, oTabla.SubTotalPagado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Cantidad", adCurrency, adParamInput, 0, IIf(oTabla.Cantidad = "", Null, oTabla.Cantidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, IIf(oTabla.PrecioUnitario = "", Null, oTabla.PrecioUnitario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobanteDetalleBienes", adInteger, adParamInput, 0, IIf(oTabla.IdComprobanteDetalle = "", Null, oTabla.IdComprobanteDetalle)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = "", Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.IdProducto = "", Null, oTabla.IdProducto)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CajaComprobantesDetalleBienesInsumos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCajaComprobantesDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesDetalleBienesInsumosEliminar"
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
Function ComprobantesDetalle(ByVal oTabla As DOCajaComprobantesDetalle) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim SQL As String
    
    SQL = "Select d.IdComprobanteDetalleServicio IdComprobanteDetalle," _
            & " d.PrecioUnitario,d.Cantidad,d.SubTotalPagado,c.Codigo,d.IdProducto " _
            & " ," & SIGHComun.sghDetalleComprobanteServicios & " as TipoDetalle " _
            & " ,c.Nombre as Producto,fs.IdFacturacionServicio as IdFacturacionDetalle " _
            & " ,isnull(fs.SubTotalExonerado,0) as SubTotalExonerado " _
            & " ,isnull(fs.SubTotalPagadoACuenta,0) as SubTotalPagadoACuenta " _
            & " from (CajaComprobantesDetalleServicios d inner join FactCatalogoServicios c " _
            & " on d.IdProducto = c.IdProducto) left join FacturacionServicios fs " _
            & " on d.IdFacturacionServicio = fs.IdFacturacionServicio " _
            & " where d.IdComprobantePago = " & oTabla.IdComprobantePago _
            & " UNION ALL " _
            & " Select d.IdComprobanteDetalleBienes IdComprobanteDetalle," _
            & " d.PrecioUnitario,d.Cantidad,d.SubTotalPagado,c.Codigo,d.IdProducto " _
            & " ," & SIGHComun.sghDetalleComprobanteInsumos & " as TipoDetalle " _
            & " ,c.Nombre as Producto, fb.IdFacturacionBienes as IdFacturacionDetalle " _
            & " ,isnull(fb.SubTotalExonerado,0) as SubTotalExonerado " _
            & " ,isnull(fb.SubTotalPagadoACuenta,0) as SubTotalPagadoACuenta " _
            & " from (CajaComprobantesDetalleBienesInsumos  d inner join CatalogoBienesInsumos c " _
            & " on d.IdProducto = c.IdProducto) left join FacturacionBienesInsumos fb " _
            & " on d.IdFacturacionBienes = fb.IdFacturacionBienes " _
            & " where IdComprobantePago = " & oTabla.IdComprobantePago

   Set ComprobantesDetalle = Nothing
   With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandText = SQL
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set ComprobantesDetalle = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

