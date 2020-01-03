VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotaCreditoDebitoTipoNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: NotaCreditoDebitoTipoNota
'        Autor: Daniel Enrique Barrantes
'        Fecha: 04/06/2015 12:44:13
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla NotaCreditoDebitoTipoNota
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DONotaCreditoDebitoTipoNota) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "NotaCreditoDebitoTipoNotaAgregar"
   Set oParameter = .CreateParameter("@IdTipoNota", adWChar, adParamInput, 10, oTabla.IdTipoNota)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoNota", adVarChar, adParamInput, 50, IIf(oTabla.NroSerie = "", Null, oTabla.TipoNota))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroSerie", adChar, adParamInput, 4, IIf(oTabla.NroSerie = "", Null, oTabla.NroSerie))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDocumento", adChar, adParamInput, 12, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDocumentoInicial", adChar, adParamInput, 12, IIf(oTabla.NroDocumentoInicial = "", Null, oTabla.NroDocumentoInicial))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDocumentoFinal", adChar, adParamInput, 12, IIf(oTabla.NroDocumentoFinal = "", Null, oTabla.NroDocumentoFinal))
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
'   Descripción:    modifica un registro a la tabla NotaCreditoDebitoTipoNota
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DONotaCreditoDebitoTipoNota) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "NotaCreditoDebitoTipoNotaModificar"
   Set oParameter = .CreateParameter("@IdTipoNota", adWChar, adParamInput, 10, oTabla.IdTipoNota)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoNota", adVarChar, adParamInput, 50, IIf(oTabla.NroSerie = "", Null, oTabla.TipoNota))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, IIf(oTabla.NroSerie = "", Null, oTabla.NroSerie))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDocumento", adChar, adParamInput, 12, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDocumentoInicial", adChar, adParamInput, 12, IIf(oTabla.NroDocumentoInicial = "", Null, oTabla.NroDocumentoInicial))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDocumentoFinal", adChar, adParamInput, 12, IIf(oTabla.NroDocumentoFinal = "", Null, oTabla.NroDocumentoFinal))
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
'   Descripción:    elimia o borra un registro a la tabla NotaCreditoDebitoTipoNota
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DONotaCreditoDebitoTipoNota) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "NotaCreditoDebitoTipoNotaEliminar"
   Set oParameter = .CreateParameter("@IdTipoNota", adWChar, adParamInput, 10, oTabla.IdTipoNota)
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
'   Descripción:    Selecciona un unico registro de la tabla NotaCreditoDebitoTipoNota
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
'sunat
Function SeleccionarPorIdTipoNota(ByVal oTabla As DONotaCreditoDebitoTipoNota) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorIdTipoNota = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "NotaCreditoDebitoTipoNotaSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdTipoNota", adWChar, adParamInput, 10, oTabla.IdTipoNota)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@TipoCompNota", adWChar, adParamInput, 10, oTabla.TipoNota)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdTipoNota = oRecordset!IdTipoNota
       oTabla.TipoNota = IIf(IsNull(oRecordset!TipoNota), "", oRecordset!TipoNota)
       oTabla.NroSerie = IIf(IsNull(oRecordset!NroSerie), "", oRecordset!NroSerie)
       oTabla.NroDocumento = IIf(IsNull(oRecordset!NroDocumento), "", oRecordset!NroDocumento)
       oTabla.NroDocumentoInicial = IIf(IsNull(oRecordset!NroDocumentoInicial), "", oRecordset!NroDocumentoInicial)
       oTabla.NroDocumentoFinal = IIf(IsNull(oRecordset!NroDocumentoFinal), "", oRecordset!NroDocumentoFinal)
   SeleccionarPorIdTipoNota = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub SetDefaults(ByVal oTabla As DONotaCreditoDebitoTipoNota)
   oTabla.IdTipoNota = ""
   oTabla.TipoNota = ""
   oTabla.NroSerie = ""
   oTabla.NroDocumento = ""
   oTabla.NroDocumentoInicial = ""
   oTabla.NroDocumentoFinal = ""
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: NotaCreditoDebitoTipoNota
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: NotaCreditoDebitoTipoNota
'        Autor:   DAniel Barrantes B.
'        Fecha: 04/06/2015 12:44:13
'------------------------------------------------------------------------------------
 

