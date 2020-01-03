VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuentasEpisodioAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: CuentasEpisodioAtencion
'        Autor: William Castro Grijalva
'        Fecha: 04/04/2005 10:11:52 p.m.
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
'   Descripción:    Inserta un registro a la tabla CuentasEpisodioAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCuentasEpisodioAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CuentasEpisodioAtencionAgregar"
           Set oParameter = .CreateParameter("@FechaAlta", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaAlta = 0, Null, oTabla.FechaAlta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaIngreso", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaIngreso = 0, Null, oTabla.FechaIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEpisodioAtencion", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdEpisodioAtencion = .Parameters("@IdEpisodioAtencion")
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
'   Descripción:    modifica un registro a la tabla CuentasEpisodioAtencion
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCuentasEpisodioAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CuentasEpisodioAtencionModificar"
           Set oParameter = .CreateParameter("@FechaAlta", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaAlta = 0, Null, oTabla.FechaAlta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaIngreso", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaIngreso = 0, Null, oTabla.FechaIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEpisodioAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdEpisodioAtencion = 0, Null, oTabla.IdEpisodioAtencion)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CuentasEpisodioAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCuentasEpisodioAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CuentasEpisodioAtencionEliminar"
           Set oParameter = .CreateParameter("@IdEpisodioAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdEpisodioAtencion = 0, Null, oTabla.IdEpisodioAtencion)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla CuentasEpisodioAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCuentasEpisodioAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CuentasEpisodioAtencionSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdEpisodioAtencion", adInteger, adParamInput, 0, oTabla.IdEpisodioAtencion): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.FechaAlta = IIf(IsNull(oRecordset!FechaAlta), 0, oRecordset!FechaAlta)
       oTabla.FechaIngreso = IIf(IsNull(oRecordset!FechaIngreso), 0, oRecordset!FechaIngreso)
       oTabla.IdEpisodioAtencion = IIf(IsNull(oRecordset!IdEpisodioAtencion), 0, oRecordset!IdEpisodioAtencion)
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


Function EliminarPorIdCuentaAtencion(lIdCuentaAtencion As Long)
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSQL As String
 
   EliminarPorIdCuentaAtencion = False
   
   'elimina los relaciones de cuenta de atencion
   sSQL = "delete from CuentasAtencionRelacionadas where "
   sSQL = sSQL + " IdCuentaAtencionPadre = " & lIdCuentaAtencion & ""
   sSQL = sSQL + " or IdCuentaAtencionHijo = " & lIdCuentaAtencion & ""
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSQL
       .Execute
   End With
   
   EliminarPorIdCuentaAtencion = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function
