VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CajaFormaPagoComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'WCG Ini 02/06/2005 [Todo el archivo]

Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: CajaFormaPagoComprobante
'        Autor: William Castro Grijalva
'        Fecha: 03/06/2005 10:25:01
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
'   Descripción:    Inserta un registro a la tabla CajaFormaPagoComprobante
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCajaFormaPagoComprobante) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaFormaPagoComprobanteAgregar"
           Set oParameter = .CreateParameter("@IdTipoFormaPago", adInteger, adParamInput, 0, IIf(oTabla.IdTipoFormaPago = 0, Null, oTabla.IdTipoFormaPago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoMoneda", adInteger, adParamInput, 0, IIf(oTabla.IdTipoMoneda = 0, Null, oTabla.IdTipoMoneda)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Importe", adCurrency, adParamInput, 0, oTabla.Importe): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalSoles", adCurrency, adParamInput, 0, oTabla.TotalSoles): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TipoCambio", adCurrency, adParamInput, 0, oTabla.TipoCambio): .Parameters.Append oParameter
           
           Set oParameter = .CreateParameter("@IdFormaPago", adInteger, adParamOutput, 0): .Parameters.Append oParameter

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
'   Descripción:    modifica un registro a la tabla CajaFormaPagoComprobante
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCajaFormaPagoComprobante) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaFormaPagoComprobanteModificar"
           Set oParameter = .CreateParameter("@TipoCambio", adCurrency, adParamInput, 0, IIf(oTabla.TipoCambio = "", Null, oTabla.TipoCambio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalSoles", adCurrency, adParamInput, 0, IIf(oTabla.TotalSoles = "", Null, oTabla.TotalSoles)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalDolares", adCurrency, adParamInput, 0, IIf(oTabla.Importe = "", Null, oTabla.Importe)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = "", Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoMoneda", adInteger, adParamInput, 0, IIf(oTabla.IdTipoMoneda = "", Null, oTabla.IdTipoMoneda)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoFormaPago", adInteger, adParamInput, 0, IIf(oTabla.IdTipoFormaPago = "", Null, oTabla.IdTipoFormaPago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFormaPago", adInteger, adParamInput, 0, IIf(oTabla.IdFormaPago = "", Null, oTabla.IdFormaPago)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CajaFormaPagoComprobante
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCajaFormaPagoComprobante) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim SQL As String

    SQL = "Delete from CajaFormaPagoComprobante where IdFormaPago = " & oTabla.IdFormaPago

   Eliminar = False
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = SQL
       .Execute
   End With
 
   Eliminar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Function EliminarPorComprobante(ByVal oTabla As DOCajaComprobantesPago) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim SQL As String

    SQL = "Delete from CajaFormaPagoComprobante where IdComprobantePago = " & oTabla.IdComprobantePago

   EliminarPorComprobante = False
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = SQL
       .Execute
   End With
 
   EliminarPorComprobante = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla CajaFormaPagoComprobante
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCajaFormaPagoComprobante) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdText
     Set .ActiveConnection = mo_Conexion
     .CommandText = "Select * from CajaFormaPagoComprobante where id = " + oTabla.IdComprobantePago
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.TipoCambio = IIf(IsNull(oRecordset!TipoCambio), "", oRecordset!TipoCambio)
       oTabla.TotalSoles = IIf(IsNull(oRecordset!TotalSoles), "", oRecordset!TotalSoles)
       oTabla.Importe = IIf(IsNull(oRecordset!TotalDolares), "", oRecordset!TotalDolares)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), "", oRecordset!IdComprobantePago)
       oTabla.IdTipoMoneda = IIf(IsNull(oRecordset!IdTipoMoneda), "", oRecordset!IdTipoMoneda)
       oTabla.IdTipoFormaPago = IIf(IsNull(oRecordset!IdTipoFormaPago), "", oRecordset!IdTipoFormaPago)
       oTabla.IdFormaPago = IIf(IsNull(oRecordset!IdFormaPago), "", oRecordset!IdFormaPago)
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


Function SeleccionarPorComprobante(oCajaFormaPago As DOCajaFormaPagoComprobante) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter

Dim sSQL As String

    sSQL = "Select * from CajaFormaPagoComprobante" _
        & " where IdComprobantePago = " & oCajaFormaPago.IdComprobantePago
 
   Set SeleccionarPorComprobante = Nothing
   With oCommand
     .CommandType = adCmdText
     Set .ActiveConnection = mo_Conexion
     .CommandText = sSQL
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarPorComprobante = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function




