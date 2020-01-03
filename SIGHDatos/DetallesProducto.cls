VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DetallesProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DetalleProductos
'        Autor: William Castro Grijalva
'        Fecha: 01/09/2004 11:51:35 p.m.
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
'   Descripción:    Inserta un registro a la tabla DetalleProductos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DODetalleProducto) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "DetalleProductosAgregar"
           Set oParameter = .CreateParameter("@IdMotivoNoAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdMotivoNoAtencion = 0, Null, oTabla.IdMotivoNoAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoProducto", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoProducto = 0, Null, oTabla.IdEstadoProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDocumentoDetalle", adInteger, adParamInput, 0, IIf(oTabla.IdDocumentoDetalle = 0, Null, oTabla.IdDocumentoDetalle)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, IIf(oTabla.PrecioUnitario = 0, Null, oTabla.PrecioUnitario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Cantidad", adCurrency, adParamInput, 0, IIf(oTabla.Cantidad = 0, Null, oTabla.Cantidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioTotal", adCurrency, adParamInput, 0, IIf(oTabla.PrecioTotal = 0, Null, oTabla.PrecioTotal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.IdProducto = 0, Null, oTabla.IdProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDetalleProducto", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CubiertoPorSeguro", adBoolean, adParamInput, 0, IIf(oTabla.CubiertoPorSeguro = "", Null, oTabla.CubiertoPorSeguro)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdDetalleProducto = .Parameters("@IdDetalleProducto")
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
'   Descripción:    modifica un registro a la tabla DetalleProductos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DODetalleProducto) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "DetalleProductosModificar"
           Set oParameter = .CreateParameter("@IdMotivoNoAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdMotivoNoAtencion = 0, Null, oTabla.IdMotivoNoAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoProducto", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoProducto = 0, Null, oTabla.IdEstadoProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDocumentoDetalle", adInteger, adParamInput, 0, IIf(oTabla.IdDocumentoDetalle = 0, Null, oTabla.IdDocumentoDetalle)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, IIf(oTabla.PrecioUnitario = 0, Null, oTabla.PrecioUnitario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Cantidad", adCurrency, adParamInput, 0, IIf(oTabla.Cantidad = 0, Null, oTabla.Cantidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioTotal", adCurrency, adParamInput, 0, IIf(oTabla.PrecioTotal = 0, Null, oTabla.PrecioTotal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.IdProducto = 0, Null, oTabla.IdProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDetalleProducto", adInteger, adParamInput, 0, IIf(oTabla.IdDetalleProducto = 0, Null, oTabla.IdDetalleProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CubiertoPorSeguro", adBoolean, adParamInput, 0, IIf(oTabla.CubiertoPorSeguro = "", Null, oTabla.CubiertoPorSeguro)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla DetalleProductos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DODetalleProducto) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "DetalleProductosEliminar"
           Set oParameter = .CreateParameter("@IdDetalleProducto", adInteger, adParamInput, 0, IIf(oTabla.IdDetalleProducto = 0, Null, oTabla.IdDetalleProducto)): .Parameters.Append oParameter
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

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla DetalleProductos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DODetalleProducto) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "DetalleProductosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdDetalleProducto", adInteger, adParamInput, 0, oTabla.IdDetalleProducto): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdMotivoNoAtencion = IIf(IsNull(oRecordset!IdMotivoNoAtencion), 0, oRecordset!IdMotivoNoAtencion)
       oTabla.IdEstadoProducto = IIf(IsNull(oRecordset!IdEstadoProducto), 0, oRecordset!IdEstadoProducto)
       oTabla.IdDocumentoDetalle = IIf(IsNull(oRecordset!IdDocumentoDetalle), 0, oRecordset!IdDocumentoDetalle)
       oTabla.PrecioUnitario = IIf(IsNull(oRecordset!PrecioUnitario), 0, oRecordset!PrecioUnitario)
       oTabla.Cantidad = IIf(IsNull(oRecordset!Cantidad), 0, oRecordset!Cantidad)
       oTabla.PrecioTotal = IIf(IsNull(oRecordset!PrecioTotal), 0, oRecordset!PrecioTotal)
       oTabla.IdProducto = IIf(IsNull(oRecordset!IdProducto), 0, oRecordset!IdProducto)
       oTabla.IdDetalleProducto = IIf(IsNull(oRecordset!IdDetalleProducto), 0, oRecordset!IdDetalleProducto)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.CubiertoPorSeguro = IIf(IsNull(oRecordset!CubiertoPorSeguro), "", oRecordset!CubiertoPorSeguro)
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


