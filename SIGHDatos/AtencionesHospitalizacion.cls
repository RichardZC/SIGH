VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionesHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: AtencionesHospitalizacion
'        Autor: William Castro Grijalva
'        Fecha: 11/10/2004 09:27:29 p.m.
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
'   Descripción:    Inserta un registro a la tabla AtencionesHospitalizacion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOAtencionHospitalizacion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesHospitalizacionAgregar"
           Set oParameter = .CreateParameter("@IdCamaEgreso", adInteger, adParamInput, 0, IIf(oTabla.IdCamaEgreso = 0, Null, oTabla.IdCamaEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCamaIngreso", adInteger, adParamInput, 0, IIf(oTabla.IdCamaIngreso = 0, Null, oTabla.IdCamaIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TieneNecropsia", adBoolean, adParamInput, 0, IIf(oTabla.TieneNecropsia = 0, Null, oTabla.TieneNecropsia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HuboInfeccionIntraHospitalaria", adBoolean, adParamInput, 0, IIf(oTabla.HuboInfeccionIntraHospitalaria = 0, Null, oTabla.HuboInfeccionIntraHospitalaria)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioEgreso", adInteger, adParamInput, 0, IIf(oTabla.IdServicioEgreso = 0, Null, oTabla.IdServicioEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoAlta", adInteger, adParamInput, 0, IIf(oTabla.IdTipoAlta = 0, Null, oTabla.IdTipoAlta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCondicionAlta", adInteger, adParamInput, 0, IIf(oTabla.IdCondicionAlta = 0, Null, oTabla.IdCondicionAlta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdAtencion = 0, Null, oTabla.IdAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionHospitalizacion", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraEgresoAdministrativo", adChar, adParamInput, 5, IIf(oTabla.HoraEgresoAdministrativo = "", Null, oTabla.HoraEgresoAdministrativo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaEgresoAdministrativo", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEgresoAdministrativo = 0, Null, oTabla.FechaEgresoAdministrativo)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdAtencionHospitalizacion = .Parameters("@IdAtencionHospitalizacion")
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
'   Descripción:    modifica un registro a la tabla AtencionesHospitalizacion
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencionHospitalizacion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesHospitalizacionModificar"
           Set oParameter = .CreateParameter("@IdCamaEgreso", adInteger, adParamInput, 0, IIf(oTabla.IdCamaEgreso = 0, Null, oTabla.IdCamaEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCamaIngreso", adInteger, adParamInput, 0, IIf(oTabla.IdCamaIngreso = 0, Null, oTabla.IdCamaIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TieneNecropsia", adBoolean, adParamInput, 0, IIf(oTabla.TieneNecropsia = 0, Null, oTabla.TieneNecropsia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HuboInfeccionIntraHospitalaria", adBoolean, adParamInput, 0, IIf(oTabla.HuboInfeccionIntraHospitalaria = 0, Null, oTabla.HuboInfeccionIntraHospitalaria)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioEgreso", adInteger, adParamInput, 0, IIf(oTabla.IdServicioEgreso = 0, Null, oTabla.IdServicioEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoAlta", adInteger, adParamInput, 0, IIf(oTabla.IdTipoAlta = 0, Null, oTabla.IdTipoAlta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCondicionAlta", adInteger, adParamInput, 0, IIf(oTabla.IdCondicionAlta = 0, Null, oTabla.IdCondicionAlta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdAtencion = 0, Null, oTabla.IdAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionHospitalizacion", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionHospitalizacion = 0, Null, oTabla.IdAtencionHospitalizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraEgresoAdministrativo", adChar, adParamInput, 5, IIf(oTabla.HoraEgresoAdministrativo = "", Null, oTabla.HoraEgresoAdministrativo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaEgresoAdministrativo", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEgresoAdministrativo = 0, Null, oTabla.FechaEgresoAdministrativo)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla AtencionesHospitalizacion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencionHospitalizacion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesHospitalizacionEliminar"
           Set oParameter = .CreateParameter("@IdAtencionHospitalizacion", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionHospitalizacion = 0, Null, oTabla.IdAtencionHospitalizacion)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla AtencionesHospitalizacion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtencionHospitalizacion) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionesHospitalizacionSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdAtencionHospitalizacion", adInteger, adParamInput, 0, oTabla.IdAtencionHospitalizacion): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdCamaEgreso = IIf(IsNull(oRecordset!IdCamaEgreso), 0, oRecordset!IdCamaEgreso)
       oTabla.IdCamaIngreso = IIf(IsNull(oRecordset!IdCamaIngreso), 0, oRecordset!IdCamaIngreso)
       oTabla.TieneNecropsia = IIf(IsNull(oRecordset!TieneNecropsia), 0, oRecordset!TieneNecropsia)
       oTabla.HuboInfeccionIntraHospitalaria = IIf(IsNull(oRecordset!HuboInfeccionIntraHospitalaria), 0, oRecordset!HuboInfeccionIntraHospitalaria)
       oTabla.IdServicioEgreso = IIf(IsNull(oRecordset!IdServicioEgreso), 0, oRecordset!IdServicioEgreso)
       oTabla.IdTipoAlta = IIf(IsNull(oRecordset!IdTipoAlta), 0, oRecordset!IdTipoAlta)
       oTabla.IdCondicionAlta = IIf(IsNull(oRecordset!IdCondicionAlta), 0, oRecordset!IdCondicionAlta)
       oTabla.IdAtencion = IIf(IsNull(oRecordset!IdAtencion), 0, oRecordset!IdAtencion)
       oTabla.IdAtencionHospitalizacion = IIf(IsNull(oRecordset!IdAtencionHospitalizacion), 0, oRecordset!IdAtencionHospitalizacion)
       oTabla.HoraEgresoAdministrativo = IIf(IsNull(oRecordset!HoraEgresoAdministrativo), "", oRecordset!HoraEgresoAdministrativo)
       oTabla.FechaEgresoAdministrativo = IIf(IsNull(oRecordset!FechaEgresoAdministrativo), 0, oRecordset!FechaEgresoAdministrativo)
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


Function SeleccionarIdPorIdAtencion(lIdAtencion As Long) As Long
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarIdPorIdAtencion = 0
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesHospitalizacionSeleccionarIdPorIdAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencionHospitalizacion", adInteger, adParamOutput, 0): .Parameters.Append oParameter
       .Execute
       SeleccionarIdPorIdAtencion = .Parameters("@IdAtencionHospitalizacion")
   End With
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
