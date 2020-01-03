VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CajaNroDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla CajaNroDocumento
'        Programado por: Castro W
'        Fecha: Agosto 2006
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
'   Descripción:    Inserta un registro a la tabla CajaNroDocumento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCajaNroDocumento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaNroDocumentoAgregar"
           Set oParameter = .CreateParameter("@IdTipoComprobante", adInteger, adParamInput, 0, IIf(oTabla.IdTipoComprobante = 0, Null, oTabla.IdTipoComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroDocumento", adChar, adParamInput, 12, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, IIf(oTabla.NroSerie = "", Null, oTabla.NroSerie)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroDocumentoFinal", adChar, adParamInput, 12, IIf(oTabla.NroDocumentoFinal = "", Null, oTabla.NroDocumentoFinal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IIf(oTabla.IdTipoComprobante = 0, Null, oTabla.IdCaja)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroDocumentoInicial", adChar, adParamInput, 12, IIf(oTabla.NroDocumentoInicial = "", Null, oTabla.NroDocumentoInicial)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FacturaSinIGV", adBoolean, adParamInput, 0, oTabla.FacturaSinIGV): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdCaja = .Parameters("@IdCaja")
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
'   Descripción:    modifica un registro a la tabla CajaNroDocumento
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCajaNroDocumento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   oTabla.NroDocumentoInicial = "000000"
   oTabla.NroDocumentoFinal = "9999999"
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaNroDocumentoModificar"
           Set oParameter = .CreateParameter("@IdTipoComprobante", adInteger, adParamInput, 0, IIf(oTabla.IdTipoComprobante = 0, Null, oTabla.IdTipoComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroDocumento", adChar, adParamInput, 12, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, IIf(oTabla.NroSerie = "", Null, oTabla.NroSerie)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroDocumentoFinal", adChar, adParamInput, 12, IIf(oTabla.NroDocumentoFinal = "", Null, oTabla.NroDocumentoFinal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IIf(oTabla.IdCaja = 0, Null, oTabla.IdCaja)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroDocumentoInicial", adChar, adParamInput, 12, IIf(oTabla.NroDocumentoInicial = "", Null, oTabla.NroDocumentoInicial)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FacturaSinIGV", adBoolean, adParamInput, 0, oTabla.FacturaSinIGV): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CajaNroDocumento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCajaNroDocumento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaNroDocumentoEliminar"
           Set oParameter = .CreateParameter("@IdTipoComprobante", adInteger, adParamInput, 0, IIf(oTabla.IdTipoComprobante = 0, Null, oTabla.IdTipoComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IIf(oTabla.IdCaja = 0, Null, oTabla.IdCaja)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla CajaNroDocumento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCajaNroDocumento) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oCajaCaja As New CajaCaja
Dim oDOCajaCaja As New DOCajaCaja
Dim oCajaNroDocumento As New CajaNroDocumento
Dim oDOCajaNroDocumento As New DOCajaNroDocumento
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CajaNroDocumentoSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdTipoComprobante", adInteger, adParamInput, 0, oTabla.IdTipoComprobante): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, oTabla.IdCaja): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       cargaTabla oTabla, oRecordset
   Else
       oDOCajaCaja.IdCaja = oTabla.IdCaja
       oDOCajaCaja.IdUsuarioAuditoria = SIGHENTIDADES.Usuario
       Set oCajaCaja.Conexion = mo_Conexion
       If oCajaCaja.SeleccionarPorId(oDOCajaCaja) = True Then
          Set oCommand = Nothing
          With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = mo_Conexion
              .CommandText = "FarmaciaUltimoTicket"
              Set oRecordset = .Execute
          End With
          If oRecordset.RecordCount = 0 Then
             oTabla.NroDocumento = "0000000"
             oTabla.NroSerie = "007"
          Else
             oTabla.NroSerie = Left(oRecordset!DocumentoNumero, InStr(oRecordset!DocumentoNumero, "-") - 1)
             oTabla.NroDocumento = Trim(Mid(oRecordset!DocumentoNumero, InStr(oRecordset!DocumentoNumero, "-") + 1, 20))
          End If
          'oTabla.IdCaja
          'oTabla.IdTipoComprobante
          oTabla.IdUsuarioAuditoria = SIGHENTIDADES.Usuario
          oTabla.NroDocumentoInicial = "0000001"
          oTabla.NroDocumentoFinal = "9999999"
          Set oCajaNroDocumento.Conexion = mo_Conexion
          If oCajaNroDocumento.Insertar(oTabla) = True Then
             SeleccionarPorId = True
          Else
             oTabla.NroDocumento = ""
             oTabla.NroSerie = ""
          End If
       Else
          MsgBox "No existe la CAJA con ID=" & oTabla.IdCaja & " chequee parametro 225" & Chr(13) & "asigne al Parametro und IdCaja existente"
       End If
   End If
   oRecordset.Close
   ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub cargaTabla(ByRef oTabla As DOCajaNroDocumento, oRecordset As Recordset)
       oTabla.IdCaja = IIf(IsNull(oRecordset!IdCaja), 0, oRecordset!IdCaja)
       oTabla.IdTipoComprobante = IIf(IsNull(oRecordset!IdTipoComprobante), 0, oRecordset!IdTipoComprobante)
       oTabla.NroDocumento = IIf(IsNull(oRecordset!NroDocumento), 0, oRecordset!NroDocumento)
       oTabla.NroSerie = IIf(IsNull(oRecordset!NroSerie), 0, oRecordset!NroSerie)
       oTabla.NroDocumentoInicial = IIf(IsNull(oRecordset!NroDocumentoInicial), 0, oRecordset!NroDocumentoInicial)
       oTabla.NroDocumentoFinal = IIf(IsNull(oRecordset!NroDocumentoFinal), 0, oRecordset!NroDocumentoFinal)
       oTabla.FacturaSinIGV = IIf(IsNull(oRecordset!FacturaSinIGV), False, oRecordset!FacturaSinIGV)
End Sub

Function EliminarPorCaja(ByVal lIdCaja As Integer) As Boolean

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CajaNroDocumentoEliminarXidCaja"
        Set oParameter = .CreateParameter("@lIdCaja", adInteger, adParamInput, 0, lIdCaja): .Parameters.Append oParameter
        .Execute
   End With
   EliminarPorCaja = True
   Set oCommand = Nothing
   Set oRecordset = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarPorIdCaja(ByVal lIdCaja As Integer) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CajaNroDocumentoSeleccionarPorIdCaja"
        Set oParameter = .CreateParameter("@lIdCaja", adInteger, adParamInput, 0, lIdCaja): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorIdCaja = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



Function SeleccionarPorIdCajaYTipoComprobante(ByVal lIdCaja As Long, tipoComprobante As Long) As DOCajaNroDocumento

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oTabla As New DOCajaNroDocumento
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CajaNroDocumentoSeleccionarPorIdCajaYTipoComprobante"
        Set oParameter = .CreateParameter("@tipoComprobante", adInteger, adParamInput, 0, tipoComprobante): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdCaja", adInteger, adParamInput, 0, lIdCaja): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorIdCajaYTipoComprobante = Nothing
   If oRecordset.RecordCount > 0 Then
      cargaTabla oTabla, oRecordset
      Set SeleccionarPorIdCajaYTipoComprobante = oTabla
   End If
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function




Function AumentarPorIdCajaYTipoComprobante(NuevoNroDocumento As String, IdTipoComprobante As Integer, IdCaja As Integer) As Boolean
On Error GoTo ManejadorDeError
Dim oTabla As New DOCajaNroDocumento
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String

    
  sSql = " Update CajaNroDocumento " & _
            " Set nroDocumento = '" & NuevoNroDocumento & "' " & _
            " Where IdTipoComprobante = " & IdTipoComprobante & _
            " and IdCaja = " & IdCaja
            
   AumentarPorIdCajaYTipoComprobante = False
   With oCommand
     .CommandType = adCmdText
     Set .ActiveConnection = mo_Conexion
     .CommandText = sSql
     Set oRecordset = .Execute
   End With
   
   AumentarPorIdCajaYTipoComprobante = True

   'oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

