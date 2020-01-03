VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InterconsultaAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: InterconsultaAtencion
'        Autor: William Castro Grijalva
'        Fecha: 31/10/2004 09:06:11 a.m.
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
'   Descripción:    Inserta un registro a la tabla InterconsultaAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOInterconsultaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "InterconsultaAtencionAgregar"
           Set oParameter = .CreateParameter("@HoraSolicitud", adChar, adParamInput, 5, IIf(oTabla.HoraSolicitud = "", Null, oTabla.HoraSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaSolicitud", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaSolicitud = 0, Null, oTabla.FechaSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraRealizacion", adChar, adParamInput, 5, IIf(oTabla.HoraRealizacion = "", Null, oTabla.HoraRealizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDetalleProducto", adInteger, adParamInput, 0, IIf(oTabla.IdDetalleProducto = 0, Null, oTabla.IdDetalleProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaRealizacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaRealizacion = 0, Null, oTabla.FechaRealizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoRealiza", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoRealiza = 0, Null, oTabla.IdMedicoRealiza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoSolicita", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoSolicita = 0, Null, oTabla.IdMedicoSolicita)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdInterconsulta", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdInterconsulta = .Parameters("@IdInterconsulta")
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
'   Descripción:    modifica un registro a la tabla InterconsultaAtencion
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOInterconsultaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "InterconsultaAtencionModificar"
           Set oParameter = .CreateParameter("@HoraSolicitud", adChar, adParamInput, 5, IIf(oTabla.HoraSolicitud = "", Null, oTabla.HoraSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaSolicitud", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaSolicitud = 0, Null, oTabla.FechaSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraRealizacion", adChar, adParamInput, 5, IIf(oTabla.HoraRealizacion = "", Null, oTabla.HoraRealizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDetalleProducto", adInteger, adParamInput, 0, IIf(oTabla.IdDetalleProducto = 0, Null, oTabla.IdDetalleProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaRealizacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaRealizacion = 0, Null, oTabla.FechaRealizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoRealiza", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoRealiza = 0, Null, oTabla.IdMedicoRealiza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoSolicita", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoSolicita = 0, Null, oTabla.IdMedicoSolicita)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdInterconsulta", adInteger, adParamInput, 0, IIf(oTabla.IdInterconsulta = 0, Null, oTabla.IdInterconsulta)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla InterconsultaAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOInterconsultaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "InterconsultaAtencionEliminar"
           Set oParameter = .CreateParameter("@IdInterconsulta", adInteger, adParamInput, 0, IIf(oTabla.IdInterconsulta = 0, Null, oTabla.IdInterconsulta)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla InterconsultaAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOInterconsultaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "InterconsultaAtencionSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdInterconsulta", adInteger, adParamInput, 0, oTabla.IdInterconsulta): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.HoraSolicitud = IIf(IsNull(oRecordset!HoraSolicitud), "", oRecordset!HoraSolicitud)
       oTabla.FechaSolicitud = IIf(IsNull(oRecordset!FechaSolicitud), 0, oRecordset!FechaSolicitud)
       oTabla.HoraRealizacion = IIf(IsNull(oRecordset!HoraRealizacion), "", oRecordset!HoraRealizacion)
       oTabla.IdDetalleProducto = IIf(IsNull(oRecordset!IdDetalleProducto), 0, oRecordset!IdDetalleProducto)
       oTabla.IdDiagnostico = IIf(IsNull(oRecordset!IdDiagnostico), 0, oRecordset!IdDiagnostico)
       oTabla.FechaRealizacion = IIf(IsNull(oRecordset!FechaRealizacion), 0, oRecordset!FechaRealizacion)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.IdMedicoRealiza = IIf(IsNull(oRecordset!IdMedicoRealiza), 0, oRecordset!IdMedicoRealiza)
       oTabla.IdMedicoSolicita = IIf(IsNull(oRecordset!IdMedicoSolicita), 0, oRecordset!IdMedicoSolicita)
       oTabla.IdInterconsulta = IIf(IsNull(oRecordset!IdInterconsulta), 0, oRecordset!IdInterconsulta)
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


Function SeleccionarPorCuentaAtencion(lIdCuentaAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorCuentaAtencion = Nothing
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "InterconsultaAtencionSeleccionarPorCuentaAtencion"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 4, lIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorCuentaAtencion = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function ActualizarInterconsultas(oInterconsultas As Collection, lIdCuentaAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSQL As String
Dim oDOInterconsultaAtencion As New DOInterconsultaAtencion
 
   ActualizarInterconsultas = False
   
   'elimina los complicaciones
   sSQL = "delete from InterconsultaAtencion where IdCuentaAtencion = " & lIdCuentaAtencion
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSQL
       .Execute
   End With
   
   If Not oInterconsultas Is Nothing Then
        For Each oDOInterconsultaAtencion In oInterconsultas
            oDOInterconsultaAtencion.IdCuentaAtencion = lIdCuentaAtencion
            If Not Insertar(oDOInterconsultaAtencion) Then
                Exit Function
            End If
         Next
    End If
    
   ActualizarInterconsultas = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function EliminarInterconsultas(lIdCuentaAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSQL As String
 
   EliminarInterconsultas = False
   
   'elimina los diagnosticos
   sSQL = "delete from InterconsultaAtencion where IdCuentaAtencion = " & lIdCuentaAtencion
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSQL
       .Execute
   End With
   
   EliminarInterconsultas = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

