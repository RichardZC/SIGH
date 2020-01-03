VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CajaComprobantesPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: CajaComprobantesPago
'        Autor: William Castro Grijalva
'        Fecha: 27/08/2006 09:54:23 p.m.
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
'   Descripción:    Inserta un registro a la tabla CajaComprobantesPago
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCajaComprobantesPago) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesPagoAgregar"
           Set oParameter = .CreateParameter("@TipoCambio", adCurrency, adParamInput, 0, IIf(oTabla.TipoCambio = 0, Null, oTabla.TipoCambio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroSerie", adChar, adParamInput, 3, IIf(oTabla.NroSerie = "", Null, oTabla.NroSerie)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroDocumento", adChar, adParamInput, 7, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@RazonSocial", adVarChar, adParamInput, 50, IIf(oTabla.RazonSocial = "", Null, oTabla.RazonSocial)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@RUC", adChar, adParamInput, 11, IIf(oTabla.RUC = "", Null, oTabla.RUC)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SubTotal", adCurrency, adParamInput, 0, IIf(oTabla.SubTotal = 0, Null, oTabla.SubTotal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IGV", adCurrency, adParamInput, 0, IIf(oTabla.IGV = 0, Null, oTabla.IGV)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCobranza", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCobranza = 0, Null, oTabla.FechaCobranza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoOrden", adInteger, adParamInput, 4, IIf(oTabla.IdTipoOrden = 0, Null, oTabla.IdTipoOrden)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observaciones", adVarChar, adParamInput, 100, IIf(oTabla.Observaciones = "", Null, oTabla.Observaciones)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoComprobante", adInteger, adParamInput, 0, IIf(oTabla.IdTipoComprobante = 0, Null, oTabla.IdTipoComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoComprobante", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoComprobante = 0, Null, oTabla.IdEstadoComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGestionCaja", adInteger, adParamInput, 4, IIf(oTabla.IdGestionCaja = 0, Null, oTabla.IdGestionCaja)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoPago", adInteger, adParamInput, 4, IIf(oTabla.IdTipoPago = 0, Null, oTabla.IdTipoPago)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, IIf(oTabla.total = 0, Null, oTabla.total)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdComprobantePago = .Parameters("@IdComprobantePago")
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
'   Descripción:    modifica un registro a la tabla CajaComprobantesPago
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCajaComprobantesPago) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesPagoModificar"
           Set oParameter = .CreateParameter("@TipoCambio", adCurrency, adParamInput, 0, IIf(oTabla.TipoCambio = 0, Null, oTabla.TipoCambio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroSerie", adChar, adParamInput, 3, IIf(oTabla.NroSerie = "", Null, oTabla.NroSerie)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroDocumento", adChar, adParamInput, 7, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@RazonSocial", adVarChar, adParamInput, 50, IIf(oTabla.RazonSocial = "", Null, oTabla.RazonSocial)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@RUC", adChar, adParamInput, 11, IIf(oTabla.RUC = "", Null, oTabla.RUC)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SubTotal", adCurrency, adParamInput, 0, IIf(oTabla.SubTotal = 0, Null, oTabla.SubTotal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IGV", adCurrency, adParamInput, 0, IIf(oTabla.IGV = 0, Null, oTabla.IGV)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCobranza", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCobranza = 0, Null, oTabla.FechaCobranza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoOrden", adInteger, adParamInput, 4, IIf(oTabla.IdTipoOrden = 0, Null, oTabla.IdTipoOrden)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observaciones", adVarChar, adParamInput, 100, IIf(oTabla.Observaciones = "", Null, oTabla.Observaciones)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoComprobante", adInteger, adParamInput, 0, IIf(oTabla.IdTipoComprobante = 0, Null, oTabla.IdTipoComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoComprobante", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoComprobante = 0, Null, oTabla.IdEstadoComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGestionCaja", adInteger, adParamInput, 4, IIf(oTabla.IdGestionCaja = 0, Null, oTabla.IdGestionCaja)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoPago", adInteger, adParamInput, 4, IIf(oTabla.IdTipoPago = 0, Null, oTabla.IdTipoPago)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, IIf(oTabla.total = 0, Null, oTabla.total)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CajaComprobantesPago
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCajaComprobantesPago) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesPagoEliminar"
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla CajaComprobantesPago
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCajaComprobantesPago) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CajaComprobantesPagoSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, oTabla.IdComprobantePago): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.TipoCambio = IIf(IsNull(oRecordset!TipoCambio), 0, oRecordset!TipoCambio)
       oTabla.NroSerie = IIf(IsNull(oRecordset!NroSerie), "", oRecordset!NroSerie)
       oTabla.NroDocumento = IIf(IsNull(oRecordset!NroDocumento), "", oRecordset!NroDocumento)
       oTabla.RazonSocial = IIf(IsNull(oRecordset!RazonSocial), "", oRecordset!RazonSocial)
       oTabla.RUC = IIf(IsNull(oRecordset!RUC), "", oRecordset!RUC)
       oTabla.SubTotal = IIf(IsNull(oRecordset!SubTotal), 0, oRecordset!SubTotal)
       oTabla.IGV = IIf(IsNull(oRecordset!IGV), 0, oRecordset!IGV)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.FechaCobranza = IIf(IsNull(oRecordset!FechaCobranza), 0, oRecordset!FechaCobranza)
       oTabla.IdTipoOrden = IIf(IsNull(oRecordset!IdTipoOrden), 0, oRecordset!IdTipoOrden)
       oTabla.Observaciones = IIf(IsNull(oRecordset!Observaciones), "", oRecordset!Observaciones)
       oTabla.IdTipoComprobante = IIf(IsNull(oRecordset!IdTipoComprobante), 0, oRecordset!IdTipoComprobante)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.IdEstadoComprobante = IIf(IsNull(oRecordset!IdEstadoComprobante), 0, oRecordset!IdEstadoComprobante)
       oTabla.IdGestionCaja = IIf(IsNull(oRecordset!IdGestionCaja), 0, oRecordset!IdGestionCaja)
       oTabla.IdTipoPago = IIf(IsNull(oRecordset!IdTipoPago), 0, oRecordset!IdTipoPago)
       oTabla.total = IIf(IsNull(oRecordset!total), 0, oRecordset!total)
       oTabla.dctos = IIf(IsNull(oRecordset!total), 0, oRecordset!dctos)
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

Function SeleccionarPorCuentaAtencion(ByVal oTabla As DOCajaComprobantesPago) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorCuentaAtencion = False
   With oCommand
     .CommandType = adCmdText
     Set .ActiveConnection = mo_Conexion
     .CommandText = "Select * from CajaComprobantesPago where IdCuentaAtencion = " & oTabla.IdCuentaAtencion
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorCuentaAtencion = True
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.IdTipoComprobante = IIf(IsNull(oRecordset!IdTipoComprobante), 0, oRecordset!IdTipoComprobante)
       oTabla.IdGestionCaja = IIf(IsNull(oRecordset!IdGestionCaja), 0, oRecordset!IdGestionCaja)
       oTabla.Observaciones = Trim(IIf(IsNull(oRecordset!Observaciones), "", oRecordset!Observaciones))
       oTabla.FechaCobranza = IIf(IsNull(oRecordset!FechaCobranza), "", oRecordset!FechaCobranza)
       oTabla.total = IIf(IsNull(oRecordset!total), 0, oRecordset!total)
       oTabla.IGV = IIf(IsNull(oRecordset!IGV), 0, oRecordset!IGV)
       oTabla.SubTotal = IIf(IsNull(oRecordset!SubTotal), 0, oRecordset!SubTotal)
       oTabla.RUC = Trim(IIf(IsNull(oRecordset!RUC), "", oRecordset!RUC))
       oTabla.RazonSocial = Trim(IIf(IsNull(oRecordset!RazonSocial), "", oRecordset!RazonSocial))
       oTabla.NroDocumento = Trim(IIf(IsNull(oRecordset!NroDocumento), "", oRecordset!NroDocumento))
       oTabla.NroSerie = Trim(IIf(IsNull(oRecordset!NroSerie), "", oRecordset!NroSerie))
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.IdTipoOrden = IIf(IsNull(oRecordset!IdTipoOrden), 0, oRecordset!IdTipoOrden)
       oTabla.IdTipoPago = IIf(IsNull(oRecordset!IdTipoPago), 0, oRecordset!IdTipoPago)
       oTabla.dctos = IIf(IsNull(oRecordset!total), 0, oRecordset!dctos)
   Else
       SeleccionarPorCuentaAtencion = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function AnularComprobanteYOrdenServicio(lIdComprobantePago As Long, lIdUsuario As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   AnularComprobanteYOrdenServicio = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesAnularComprobanteYOrdenServicio"
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lIdComprobantePago): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, lIdUsuario)
       .Parameters.Append oParameter
       .Execute
   End With
 
   AnularComprobanteYOrdenServicio = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "AnularComprobanteYOrdenServicio"
Exit Function
End Function

Function AnularComprobanteYOrdenBienInsumo(lIdComprobantePago As Long, lIdUsuario As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   AnularComprobanteYOrdenBienInsumo = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesAnularComprobanteYOrdenBienInsumo"
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lIdComprobantePago): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, lIdUsuario)
       .Parameters.Append oParameter
       .Execute
   End With
 
   AnularComprobanteYOrdenBienInsumo = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "AnularComprobanteYOrdenBienInsumo"
Exit Function
End Function
Function DevolverComprobanteYOrdenServicio(lIdComprobantePago As Long, lIdComprobanteDeDevolucion As Long, lIdUsuario As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   DevolverComprobanteYOrdenServicio = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesDevolverComprobanteYOrdenServicio"
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lIdComprobantePago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdComprobantePagoDevolucion", adInteger, adParamInput, 0, lIdComprobanteDeDevolucion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, lIdUsuario): .Parameters.Append oParameter
       .Execute
   End With
 
   DevolverComprobanteYOrdenServicio = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "DevolverComprobanteYOrdenServicio"
Exit Function
End Function

Function DevolverComprobanteYOrdenBienInsumo(lIdComprobantePago As Long, lIdComprobanteDeDevolucion As Long, lIdUsuario As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   DevolverComprobanteYOrdenBienInsumo = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesDevolverComprobanteYOrdenBienInsumo"
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lIdComprobantePago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdComprobantePagoDevolucion", adInteger, adParamInput, 0, lIdComprobanteDeDevolucion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, lIdUsuario): .Parameters.Append oParameter
       .Execute
   End With
 
   DevolverComprobanteYOrdenBienInsumo = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "DevolverComprobanteYOrdenBienInsumo"
Exit Function
End Function


Sub Trace(sMensaje As String, NombreDeMetodo As String)
    MsgBox "Mensaje: " & sMensaje + Chr(13) + "Metodo:" & NombreDeMetodo + Chr(13) + "Clase de Datos: CajaComprobantesPago", vbInformation, "Atenciones"
End Sub

