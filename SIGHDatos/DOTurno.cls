VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: Turnos
'        Autor: William Castro Grijalva
'        Fecha: 06/05/2004 09:43:15 a.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOTurno) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "TurnosAgregar"
           Set oParameter = .CreateParameter("@Codigo", adChar, adParamInput, 3, oTabla.Codigo): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, oTabla.IdTipoServicio): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraFin", adDBTimeStamp, adParamInput, 0, oTabla.HoraFin): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraInicio", adDBTimeStamp, adParamInput, 0, oTabla.HoraInicio): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 50, oTabla.Descripcion): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdTurno = .Parameters("@IdTurno")
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
'   Descripción:    modifica un registro a la tabla Turnos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOTurno) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "TurnosModificar"
       Set oParameter = .CreateParameter("@Codigo", adChar, adParamInput, 3, oTabla.Codigo): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, oTabla.IdTipoServicio): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@HoraFin", adDBTimeStamp, adParamInput, 0, oTabla.HoraFin): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@HoraInicio", adDBTimeStamp, adParamInput, 0, oTabla.HoraInicio): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 50, oTabla.Descripcion): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, oTabla.IdTurno): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOTurno) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "TurnosEliminar"
       Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, oTabla.IdTurno): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorPrimaryKey(ByVal oTabla As DOTurnos) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorPrimaryKey = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TurnosSeleccionarPorPrimaryKey"
       Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, oTabla.IdTurno): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorPrimaryKey = True
       oTabla.Codigo = oRecordset("Codigo")
       oTabla.IdTipoServicio = oRecordset("IdTipoServicio")
       oTabla.HoraFin = oRecordset("HoraFin")
       oTabla.HoraInicio = oRecordset("HoraInicio")
       oTabla.Descripcion = oRecordset("Descripcion")
       oTabla.IdTurno = oRecordset("IdTurno")
   Else
       SeleccionarPorPrimaryKey = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


