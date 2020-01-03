VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "His_FactCatalogoServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla HIS_FACTCATALOGOSERVICIOS
'        Programado por: Cachay F
'        Fecha: Noviembre 2013
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
'   Descripción:    Inserta un registro a la tabla HIS_FACTCATALOGOSERVICIOS
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOHIS_FACTCATALOGOSERVICIOS) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_FACTCATALOGOSERVICIOSAgregar"
   Set oParameter = .CreateParameter("@IdDiagCpt", adInteger, adParamInput, 0, IIf(oTabla.IdDiagCpt = 0, Null, oTabla.IdDiagCpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoDiagCpt", adVarChar, adParamInput, 20, IIf(oTabla.codigodiagcpt = "", Null, oTabla.codigodiagcpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DescripcionDiagCpt", adVarChar, adParamInput, 1000, IIf(oTabla.descripciondiagcpt = "", Null, oTabla.descripciondiagcpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsCpt", adVarChar, adParamInput, 1, IIf(oTabla.EsCpt = "", Null, oTabla.EsCpt))
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
'   Descripción:    modifica un registro a la tabla HIS_FACTCATALOGOSERVICIOS
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOHIS_FACTCATALOGOSERVICIOS) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_FACTCATALOGOSERVICIOSModificar"
   Set oParameter = .CreateParameter("@IdDiagCpt", adInteger, adParamInput, 0, IIf(oTabla.IdDiagCpt = 0, Null, oTabla.IdDiagCpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoDiagCpt", adVarChar, adParamInput, 20, IIf(oTabla.codigodiagcpt = "", Null, oTabla.codigodiagcpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DescripcionDiagCpt", adVarChar, adParamInput, 1000, IIf(oTabla.descripciondiagcpt = "", Null, oTabla.descripciondiagcpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsCpt", adVarChar, adParamInput, 1, IIf(oTabla.EsCpt = "", Null, oTabla.EsCpt))
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
'   Descripción:    elimia o borra un registro a la tabla HIS_FACTCATALOGOSERVICIOS
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOHIS_FACTCATALOGOSERVICIOS) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_FACTCATALOGOSERVICIOSEliminar"
   Set oParameter = .CreateParameter("@IdDiagCpt", adInteger, adParamInput, 0, oTabla.IdDiagCpt)
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
'   Descripción:    Selecciona un unico registro de la tabla HIS_FACTCATALOGOSERVICIOS
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOHIS_FACTCATALOGOSERVICIOS) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter

Dim oConexion As New ADODB.Connection
oConexion.Open SIGHENTIDADES.CadenaConexion
oConexion.CursorLocation = adUseClient
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = oConexion
     .CommandText = "HIS_FACTCATALOGOSERVICIOSSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdDiagCpt", adInteger, adParamInput, 0, oTabla.IdDiagCpt)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdDiagCpt = IIf(IsNull(oRecordset!IdDiagCpt), 0, oRecordset!IdDiagCpt)
       oTabla.codigodiagcpt = IIf(IsNull(oRecordset!codigodiagcpt), "", oRecordset!codigodiagcpt)
       oTabla.descripciondiagcpt = IIf(IsNull(oRecordset!descripciondiagcpt), "", oRecordset!descripciondiagcpt)
       oTabla.EsCpt = IIf(IsNull(oRecordset!EsCpt), "", oRecordset!EsCpt)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
   oConexion.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOHIS_FACTCATALOGOSERVICIOS)
   oTabla.IdDiagCpt = 0
   oTabla.codigodiagcpt = ""
   oTabla.descripciondiagcpt = ""
   oTabla.EsCpt = ""
End Sub


