VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CajaCajeroSupervisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: CajaCajeroSupervisor
'        Autor: William Castro Grijalva
'        Fecha: 16/08/2005 11:06:17
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
'   Descripción:    Inserta un registro a la tabla CajaCajeroSupervisor
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCajaCajeroSupervisor) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaCajeroSupervisorAgregar"
           Set oParameter = .CreateParameter("@IdSupervisor", adInteger, adParamInput, 0, IIf(oTabla.IdSupervisor = 0, Null, oTabla.IdSupervisor)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCajeroSupervisor", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdCajeroSupervisor = .Parameters("@IdCajeroSupervisor")
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
'   Descripción:    modifica un registro a la tabla CajaCajeroSupervisor
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCajaCajeroSupervisor) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaCajeroSupervisorModificar"
           Set oParameter = .CreateParameter("@IdSupervisor", adInteger, adParamInput, 0, IIf(oTabla.IdSupervisor = 0, Null, oTabla.IdSupervisor)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCajeroSupervisor", adInteger, adParamInput, 0, IIf(oTabla.IdCajeroSupervisor = 0, Null, oTabla.IdCajeroSupervisor)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CajaCajeroSupervisor
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCajaCajeroSupervisor) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaCajeroSupervisorEliminar"
           Set oParameter = .CreateParameter("@IdCajeroSupervisor", adInteger, adParamInput, 0, IIf(oTabla.IdCajeroSupervisor = 0, Null, oTabla.IdCajeroSupervisor)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla CajaCajeroSupervisor
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCajaCajeroSupervisor) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CajaCajeroSupervisorSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdCajeroSupervisor", adInteger, adParamInput, 0, oTabla.IdCajeroSupervisor): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdSupervisor = IIf(IsNull(oRecordset!IdSupervisor), 0, oRecordset!IdSupervisor)
       oTabla.IdTurno = IIf(IsNull(oRecordset!IdTurno), 0, oRecordset!IdTurno)
       oTabla.IdCajero = IIf(IsNull(oRecordset!IdCajero), 0, oRecordset!IdCajero)
       oTabla.IdCajeroSupervisor = IIf(IsNull(oRecordset!IdCajeroSupervisor), 0, oRecordset!IdCajeroSupervisor)
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

Function SeleccionarPorIdCajero(ByVal lIdCajero As Integer) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String

    sSQL = "Select *,null as idSupervisor from CajaCajeroSupervisor " _
            & " Where IdCajero = " & lIdCajero _
            & " Union All " _
            & " Select 0," & lIdCajero & " ,IdTurno,null as idSupervisor from CajaTurno " _
            & " where IdTurno not in (Select IdTurno " _
            & " from CajaCajeroSupervisor where IdCajero = " & lIdCajero & " )"
 
   Set SeleccionarPorIdCajero = Nothing
   With oCommand
     .CommandType = adCmdText
     Set .ActiveConnection = mo_Conexion
     .CommandText = sSQL
     Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorIdCajero = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function EliminarPorCajero(ByVal lIdCajero As Integer) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   EliminarPorCajero = False
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = "Delete from CajaCajeroSupervisor where IdCajero = " & lIdCajero
       .Execute
   End With
   EliminarPorCajero = True
   ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

