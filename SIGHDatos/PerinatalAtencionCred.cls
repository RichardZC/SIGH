VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerinatalAtencionCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla PerinatalAtencionCred
'        Programado por: Barrantes D
'        Fecha: Agosto 2011
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
'   Descripción:    Inserta un registro a la tabla PerinatalAtencionCred
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoPerinatalAtencionCred) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PerinatalAtencionCredAgregar"
   Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamInput, 0, IIf(oTabla.idPerinatalAtencion = 0, Null, oTabla.idPerinatalAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadEnAnios", adVarChar, adParamInput, 10, IIf(oTabla.EdadEnAnios = "", Null, oTabla.EdadEnAnios))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CredNumero", adInteger, adParamInput, 0, IIf(oTabla.CredNumero = 0, Null, oTabla.CredNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CredCheck", adVarChar, adParamInput, 1, IIf(oTabla.CredCheck = "", Null, oTabla.CredCheck))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
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
'   Descripción:    modifica un registro a la tabla PerinatalAtencionCred
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoPerinatalAtencionCred) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PerinatalAtencionCredModificar"
   Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamInput, 0, IIf(oTabla.idPerinatalAtencion = 0, Null, oTabla.idPerinatalAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadEnAnios", adVarChar, adParamInput, 10, IIf(oTabla.EdadEnAnios = "", Null, oTabla.EdadEnAnios))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CredNumero", adInteger, adParamInput, 0, IIf(oTabla.CredNumero = 0, Null, oTabla.CredNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CredCheck", adVarChar, adParamInput, 1, IIf(oTabla.CredCheck = "", Null, oTabla.CredCheck))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
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
'   Descripción:    elimia o borra un registro a la tabla PerinatalAtencionCred
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoPerinatalAtencionCred) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PerinatalAtencionCredEliminar"
   Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamInput, 0, oTabla.idPerinatalAtencion)
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
'   Descripción:    Selecciona un unico registro de la tabla PerinatalAtencionCred
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoPerinatalAtencionCred) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "PerinatalAtencionCredSeleccionarPorId"
     Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamInput, 0, oTabla.idPerinatalAtencion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idPerinatalAtencion = IIf(IsNull(oRecordset!idPerinatalAtencion), 0, oRecordset!idPerinatalAtencion)
       oTabla.EdadEnAnios = IIf(IsNull(oRecordset!EdadEnAnios), "", oRecordset!EdadEnAnios)
       oTabla.CredNumero = IIf(IsNull(oRecordset!CredNumero), 0, oRecordset!CredNumero)
       oTabla.CredCheck = IIf(IsNull(oRecordset!CredCheck), "", oRecordset!CredCheck)
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoPerinatalAtencionCred)
   oTabla.idPerinatalAtencion = 0
   oTabla.EdadEnAnios = ""
   oTabla.CredNumero = 0
   oTabla.CredCheck = ""
   oTabla.idAtencion = 0
End Sub

Function PerinatalAtencionCredSeleccionarPorIdPaciente(ml_idPaciente As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "PerinatalAtencionCredSeleccionarPorIdPaciente"
        Set oParameter = .CreateParameter("@ml_idPaciente", adInteger, adParamInput, 0, ml_idPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PerinatalAtencionCredSeleccionarPorIdPaciente = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function EliminarXatencion(ByVal oTabla As DoPerinatalAtencionCred) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
EliminarXatencion = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PerinatalAtencionCredEliminarXidAtencion"
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, oTabla.idAtencion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
EliminarXatencion = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function



