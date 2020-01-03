VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sms_mensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: sms_mensajes
'        Autor: Daniel Enrique Barrantes
'        Fecha: 07/06/2018 10:07:46 a.m.
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
'   Descripción:    Inserta un registro a la tabla sms_mensajes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOsms_mensajes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "sms_mensajesAgregar"
   Set oParameter = .CreateParameter("@idMensaje", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@celular", adVarChar, adParamInput, 20, IIf(oTabla.celular = "", Null, oTabla.celular))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.idCuentaAtencion = 0, Null, oTabla.idCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Mensaje", adVarChar, adParamInput, 140, IIf(oTabla.Mensaje = "", Null, oTabla.Mensaje))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.idPaciente = 0, Null, oTabla.idPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaEnvio", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEnvio = 0, Null, oTabla.FechaEnvio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@opcion", adVarChar, adParamInput, 50, IIf(oTabla.opcion = "", Null, oTabla.opcion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoEnvio", adInteger, adParamInput, 0, IIf(oTabla.idEstadoEnvio = 0, Null, oTabla.idEstadoEnvio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.idMensaje = .Parameters("@idMensaje")
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
'   Descripción:    modifica un registro a la tabla sms_mensajes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOsms_mensajes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "sms_mensajesModificar"
   Set oParameter = .CreateParameter("@idMensaje", adInteger, adParamInput, 0, IIf(oTabla.idMensaje = 0, Null, oTabla.idMensaje))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@celular", adVarChar, adParamInput, 20, IIf(oTabla.celular = "", Null, oTabla.celular))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.idCuentaAtencion = 0, Null, oTabla.idCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Mensaje", adVarChar, adParamInput, 140, IIf(oTabla.Mensaje = "", Null, oTabla.Mensaje))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.idPaciente = 0, Null, oTabla.idPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaEnvio", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEnvio = 0, Null, oTabla.FechaEnvio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@opcion", adVarChar, adParamInput, 50, IIf(oTabla.opcion = "", Null, oTabla.opcion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoEnvio", adInteger, adParamInput, 0, IIf(oTabla.idEstadoEnvio = 0, Null, oTabla.idEstadoEnvio))
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
'   Descripción:    elimia o borra un registro a la tabla sms_mensajes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOsms_mensajes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "sms_mensajesEliminar"
   Set oParameter = .CreateParameter("@idMensaje", adInteger, adParamInput, 0, oTabla.idMensaje)
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
'   Descripción:    Selecciona un unico registro de la tabla sms_mensajes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOsms_mensajes) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "sms_mensajesSeleccionarPorId"
     Set oParameter = .CreateParameter("@idMensaje", adInteger, adParamInput, 0, oTabla.idMensaje)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idMensaje = IIf(IsNull(oRecordset!idMensaje), 0, oRecordset!idMensaje)
       oTabla.celular = IIf(IsNull(oRecordset!celular), "", oRecordset!celular)
       oTabla.idCuentaAtencion = IIf(IsNull(oRecordset!idCuentaAtencion), 0, oRecordset!idCuentaAtencion)
       oTabla.Mensaje = IIf(IsNull(oRecordset!Mensaje), "", oRecordset!Mensaje)
       oTabla.idPaciente = IIf(IsNull(oRecordset!idPaciente), 0, oRecordset!idPaciente)
       oTabla.FechaEnvio = IIf(IsNull(oRecordset!FechaEnvio), 0, oRecordset!FechaEnvio)
       oTabla.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
       oTabla.opcion = IIf(IsNull(oRecordset!opcion), "", oRecordset!opcion)
       oTabla.idEstadoEnvio = IIf(IsNull(oRecordset!idEstadoEnvio), 0, oRecordset!idEstadoEnvio)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOsms_mensajes)
   oTabla.idMensaje = 0
   oTabla.celular = ""
   oTabla.idCuentaAtencion = 0
   oTabla.Mensaje = ""
   oTabla.idPaciente = 0
   oTabla.FechaEnvio = Now
   oTabla.FechaCreacion = Now
   oTabla.opcion = ""
   oTabla.idEstadoEnvio = 0
End Sub

Function EliminarMenoresAfecha(ByVal oTabla As DOsms_mensajes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
EliminarMenoresAfecha = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "sms_mensajesEliminarMenoresAfecha"
   Set oParameter = .CreateParameter("@ldFecha", adDBTimeStamp, adParamInput, 0, oTabla.FechaCreacion)
   .Parameters.Append oParameter
   .Execute
End With
 
EliminarMenoresAfecha = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: sms_mensajes
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: sms_mensajes
'        Autor:   DAniel Barrantes B.
'        Fecha: 07/06/2018 10:07:47 a.m.
'------------------------------------------------------------------------------------
 

