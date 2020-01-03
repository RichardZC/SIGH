VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FactReembolsosDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FactReembolsosDocumentos
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

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
'   Descripción:    Inserta un registro a la tabla FactReembolsosDocumentos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOFactReembolsosDcto) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter

Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FactReembolsosDocumentosAgregar"
   Set oParameter = .CreateParameter("@idFactReembolso", adInteger, adParamInput, 0, IIf(oTabla.idFactReembolso = 0, Null, oTabla.idFactReembolso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, IIf(oTabla.NroSerie = "", Null, oTabla.NroSerie))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MotivoAnulacion", adVarChar, adParamInput, 150, IIf(oTabla.MotivoAnulacion = "", Null, oTabla.MotivoAnulacion))
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
Resume
End Function

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla FactReembolsosDocumentos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFactReembolsosDcto) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FactReembolsosDocumentosModificar"
   Set oParameter = .CreateParameter("@idFactReembolso", adInteger, adParamInput, 0, IIf(oTabla.idFactReembolso = 0, Null, oTabla.idFactReembolso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, IIf(oTabla.NroSerie = "", Null, oTabla.NroSerie))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MotivoAnulacion", adVarChar, adParamInput, 150, IIf(oTabla.MotivoAnulacion = "", Null, oTabla.MotivoAnulacion))
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
'   Descripción:    elimia o borra un registro a la tabla FactReembolsosDocumentos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFactReembolsosDcto) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FactReembolsosDocumentosEliminar"
   Set oParameter = .CreateParameter("@idFactReembolso", adInteger, adParamInput, 0, oTabla.idFactReembolso)
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
'   Descripción:    Selecciona un unico registro de la tabla FactReembolsosDocumentos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFactReembolsosDcto) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FactReembolsosDocumentosSeleccionarPorId"
     Set oParameter = .CreateParameter("@idFactReembolso", adInteger, adParamInput, 0, oTabla.idFactReembolso)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idFactReembolso = IIf(IsNull(oRecordset!idFactReembolso), 0, oRecordset!idFactReembolso)
       oTabla.NroSerie = IIf(IsNull(oRecordset!NroSerie), "", oRecordset!NroSerie)
       oTabla.NroDocumento = IIf(IsNull(oRecordset!NroDocumento), "", oRecordset!NroDocumento)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.MotivoAnulacion = IIf(IsNull(oRecordset!MotivoAnulacion), "", oRecordset!MotivoAnulacion)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOFactReembolsosDcto)
   oTabla.idFactReembolso = 0
   oTabla.NroSerie = ""
   oTabla.NroDocumento = ""
   oTabla.MotivoAnulacion = ""
End Sub


