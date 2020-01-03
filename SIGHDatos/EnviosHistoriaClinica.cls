VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnviosHistoriaClinica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: EnviosHistoriaClinica
'        Autor: William Castro Grijalva
'        Fecha: 04/09/2004 12:02:55 a.m.
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
'   Descripción:    Inserta un registro a la tabla EnviosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOEnvioHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EnviosHistoriaClinicaAgregar"
           Set oParameter = .CreateParameter("@HoraPrestamoReal", adChar, adParamInput, 5, IIf(oTabla.HoraPrestamoReal = "", Null, oTabla.HoraPrestamoReal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaPrestamoReal", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaPrestamoReal = 0, Null, oTabla.FechaPrestamoReal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdResponsableRecepcion", adInteger, adParamInput, 4, IIf(oTabla.IdResponsableRecepcion = 0, Null, oTabla.IdResponsableRecepcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdResponsableEnvio", adInteger, adParamInput, 4, IIf(oTabla.IdResponsableEnvio = 0, Null, oTabla.IdResponsableEnvio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEnvio", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdEnvio = .Parameters("@IdEnvio")
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
'   Descripción:    modifica un registro a la tabla EnviosHistoriaClinica
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOEnvioHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EnviosHistoriaClinicaModificar"
           Set oParameter = .CreateParameter("@HoraPrestamoReal", adChar, adParamInput, 5, IIf(oTabla.HoraPrestamoReal = "", Null, oTabla.HoraPrestamoReal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaPrestamoReal", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaPrestamoReal = 0, Null, oTabla.FechaPrestamoReal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdResponsableRecepcion", adInteger, adParamInput, 4, IIf(oTabla.IdResponsableRecepcion = 0, Null, oTabla.IdResponsableRecepcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdResponsableEnvio", adInteger, adParamInput, 4, IIf(oTabla.IdResponsableEnvio = 0, Null, oTabla.IdResponsableEnvio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEnvio", adInteger, adParamInput, 4, IIf(oTabla.IdEnvio = 0, Null, oTabla.IdEnvio)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla EnviosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOEnvioHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EnviosHistoriaClinicaEliminar"
           Set oParameter = .CreateParameter("@IdEnvio", adInteger, adParamInput, 4, IIf(oTabla.IdEnvio = 0, Null, oTabla.IdEnvio)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla EnviosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOEnvioHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EnviosHistoriaClinicaSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdEnvio", adInteger, adParamInput, 4, oTabla.IdEnvio): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.HoraPrestamoReal = IIf(IsNull(oRecordset!HoraPrestamoReal), "", oRecordset!HoraPrestamoReal)
       oTabla.FechaPrestamoReal = IIf(IsNull(oRecordset!FechaPrestamoReal), 0, oRecordset!FechaPrestamoReal)
       oTabla.IdResponsableRecepcion = IIf(IsNull(oRecordset!IdResponsableRecepcion), 0, oRecordset!IdResponsableRecepcion)
       oTabla.IdResponsableEnvio = IIf(IsNull(oRecordset!IdResponsableEnvio), 0, oRecordset!IdResponsableEnvio)
       oTabla.IdEnvio = IIf(IsNull(oRecordset!IdEnvio), 0, oRecordset!IdEnvio)
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


