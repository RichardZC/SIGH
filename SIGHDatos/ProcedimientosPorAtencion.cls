VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionesProcedimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: AtencionesProcedimientos
'        Autor: William Castro Grijalva
'        Fecha: 17/03/2005 12:17:16 a.m.
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
'   Descripción:    Inserta un registro a la tabla AtencionesProcedimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOAtencionProcedimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesProcedimientosAgregar"
           Set oParameter = .CreateParameter("@NroOrden", adChar, adParamInput, 10, IIf(oTabla.NroOrden = "", Null, oTabla.NroOrden)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoRealiza", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoRealiza = 0, Null, oTabla.IdMedicoRealiza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioRealiza", adInteger, adParamInput, 0, IIf(oTabla.IdServicioRealiza = 0, Null, oTabla.IdServicioRealiza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDetalleProducto", adInteger, adParamInput, 0, IIf(oTabla.IdDetalleProducto = 0, Null, oTabla.IdDetalleProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraRealizacion", adChar, adParamInput, 5, IIf(oTabla.HoraRealizacion = "", Null, oTabla.HoraRealizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaRealizacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaRealizacion = 0, Null, oTabla.FechaRealizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProcedimiento", adInteger, adParamInput, 0, IIf(oTabla.IdProcedimiento = 0, Null, oTabla.IdProcedimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionProcedimiento", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdAtencionProcedimiento = .Parameters("@IdAtencionProcedimiento")
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
'   Descripción:    modifica un registro a la tabla AtencionesProcedimientos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencionProcedimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesProcedimientosModificar"
           Set oParameter = .CreateParameter("@NroOrden", adChar, adParamInput, 10, IIf(oTabla.NroOrden = "", Null, oTabla.NroOrden)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoRealiza", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoRealiza = 0, Null, oTabla.IdMedicoRealiza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioRealiza", adInteger, adParamInput, 0, IIf(oTabla.IdServicioRealiza = 0, Null, oTabla.IdServicioRealiza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDetalleProducto", adInteger, adParamInput, 0, IIf(oTabla.IdDetalleProducto = 0, Null, oTabla.IdDetalleProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraRealizacion", adChar, adParamInput, 5, IIf(oTabla.HoraRealizacion = "", Null, oTabla.HoraRealizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaRealizacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaRealizacion = 0, Null, oTabla.FechaRealizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProcedimiento", adInteger, adParamInput, 0, IIf(oTabla.IdProcedimiento = 0, Null, oTabla.IdProcedimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionProcedimiento", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionProcedimiento = 0, Null, oTabla.IdAtencionProcedimiento)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla AtencionesProcedimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencionProcedimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesProcedimientosEliminar"
           Set oParameter = .CreateParameter("@IdAtencionProcedimiento", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionProcedimiento = 0, Null, oTabla.IdAtencionProcedimiento)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla AtencionesProcedimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtencionProcedimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionesProcedimientosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdAtencionProcedimiento", adInteger, adParamInput, 0, oTabla.IdAtencionProcedimiento): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.NroOrden = IIf(IsNull(oRecordset!NroOrden), "", oRecordset!NroOrden)
       oTabla.IdMedicoRealiza = IIf(IsNull(oRecordset!IdMedicoRealiza), 0, oRecordset!IdMedicoRealiza)
       oTabla.IdServicioRealiza = IIf(IsNull(oRecordset!IdServicioRealiza), 0, oRecordset!IdServicioRealiza)
       oTabla.IdDetalleProducto = IIf(IsNull(oRecordset!IdDetalleProducto), 0, oRecordset!IdDetalleProducto)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.HoraRealizacion = IIf(IsNull(oRecordset!HoraRealizacion), "", oRecordset!HoraRealizacion)
       oTabla.FechaRealizacion = IIf(IsNull(oRecordset!FechaRealizacion), 0, oRecordset!FechaRealizacion)
       oTabla.IdProcedimiento = IIf(IsNull(oRecordset!IdProcedimiento), 0, oRecordset!IdProcedimiento)
       oTabla.IdAtencionProcedimiento = IIf(IsNull(oRecordset!IdAtencionProcedimiento), 0, oRecordset!IdAtencionProcedimiento)
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

Function SeleccionarPorCuentaAtencionSinIterconsultas(lIdCuentaAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorCuentaAtencionSinIterconsultas = Nothing
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtencionesProcedimientosSeleccionarPorCuentaAtencionSinInterconsulta"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 4, lIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorCuentaAtencionSinIterconsultas = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Function SeleccionarPorCuentaAtencionEIterconsulta(lIdCuentaAtencion As Long, lIdInterconsulta As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorCuentaAtencionEIterconsulta = Nothing
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtencionesProcedimientosSeleccionarPorCuentaAtencionEInterconsulta"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 4, lIdCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdInterconsulta", adInteger, adParamInput, 4, lIdInterconsulta): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorCuentaAtencionEIterconsulta = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function ActualizarProcedimientosAtencion(oProcedimientos As Collection, lIdCuentaAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSQL As String
Dim oDOAtencionProcedimiento As New DOAtencionProcedimiento
 
   ActualizarProcedimientosAtencion = False
   
   'elimina los complicaciones
   sSQL = "delete from AtencionesProcedimientos where IdCuentaAtencion = " & lIdCuentaAtencion
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSQL
       .Execute
   End With
   
   If Not oProcedimientos Is Nothing Then
        For Each oDOAtencionProcedimiento In oProcedimientos
            oDOAtencionProcedimiento.IdCuentaAtencion = lIdCuentaAtencion
            If Not Insertar(oDOAtencionProcedimiento) Then
                Exit Function
            End If
         Next
    End If
    
   ActualizarProcedimientosAtencion = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Function ActualizarProcedimientosInterconsultas(oProcedimientos As Collection, lIdCuentaAtencion As Long, lIdInterconsulta As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSQL As String
Dim oDOAtencionProcedimiento As New DOAtencionProcedimiento
 
   ActualizarProcedimientosInterconsultas = False
   
'   ''elimina solo los proc de interconsultas
'   sSQL = "delete from AtencionesProcedimientos where IdCuentaAtencion = " & lIdCuentaAtencion & " and IdInterconsulta = " & lIdInterconsulta
'   With oCommand
'       .CommandType = adCmdText
'       Set .ActiveConnection = mo_Conexion
'       .CommandText = sSQL
'       .Execute
'   End With
'
'   If Not oProcedimientos Is Nothing Then
'        For Each oDOAtencionProcedimiento In oProcedimientos
'            oDOAtencionProcedimiento.IdCuentaAtencion = lIdCuentaAtencion
'            oDOAtencionProcedimiento.IdInterconsulta = lIdInterconsulta
'            If Not Insertar(oDOAtencionProcedimiento) Then
'                Exit Function
'            End If
'         Next
'    End If
    
   ActualizarProcedimientosInterconsultas = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function EliminarProcedimientosPorCuentaAtencion(lIdCuentaAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSQL As String
 
   EliminarProcedimientosPorCuentaAtencion = False
   
   'elimina los diagnosticos
   sSQL = "delete from AtencionesProcedimientos where IdCuentaAtencion = " & lIdCuentaAtencion
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSQL
       .Execute
   End With
   
   EliminarProcedimientosPorCuentaAtencion = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

