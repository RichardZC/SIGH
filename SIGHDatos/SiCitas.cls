VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SiCitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: SiCitas
'        Autor: Daniel Enrique Barrantes
'        Fecha: 13/11/2018 09:18:59 a.m.
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
'   Descripción:    Inserta un registro a la tabla SiCitas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoSiCitas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SiCitasAgregar"
   Set oParameter = .CreateParameter("@idCitaSI", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.IdPuntoCarga = 0, Null, oTabla.IdPuntoCarga))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Fecha = 0, Null, oTabla.Fecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@horaInicio", adVarChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraFinal", adVarChar, adParamInput, 5, IIf(oTabla.HoraFinal = "", Null, oTabla.HoraFinal))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Paciente", adVarChar, adParamInput, 100, IIf(oTabla.Paciente = "", Null, oTabla.Paciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.idTipoSexo = 0, Null, oTabla.idTipoSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idMedico", adInteger, adParamInput, 0, IIf(oTabla.IdMedico = 0, Null, oTabla.IdMedico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idMovimiento", adInteger, adParamInput, 0, IIf(oTabla.idMovimiento = 0, Null, oTabla.idMovimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, IIf(oTabla.IdEstado = 0, Null, oTabla.IdEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, IIf(oTabla.IdReceta = 0, Null, oTabla.IdReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdSala", adInteger, adParamInput, 0, oTabla.IdSala)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@llaveTicket", adVarChar, adParamInput, 20, IIf(oTabla.llaveTicket = "", Null, oTabla.llaveTicket))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cupo", adInteger, adParamInput, 0, oTabla.cupo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Telefono", adVarChar, adParamInput, 10, IIf(oTabla.Telefono = "", Null, oTabla.Telefono))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Direccion", adVarChar, adParamInput, 100, IIf(oTabla.Direccion = "", Null, oTabla.Direccion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProgramacion", adInteger, adParamInput, 0, oTabla.idProgramacion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdResponsable", adInteger, adParamInput, 0, oTabla.IdResponsable)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.idCitaSI = .Parameters("@idCitaSI")
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
'   Descripción:    modifica un registro a la tabla SiCitas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoSiCitas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SiCitasModificar"
   Set oParameter = .CreateParameter("@idCitaSI", adInteger, adParamInput, 0, IIf(oTabla.idCitaSI = 0, Null, oTabla.idCitaSI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.IdPuntoCarga = 0, Null, oTabla.IdPuntoCarga))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Fecha = 0, Null, oTabla.Fecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@horaInicio", adVarChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraFinal", adVarChar, adParamInput, 5, IIf(oTabla.HoraFinal = "", Null, oTabla.HoraFinal))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Paciente", adVarChar, adParamInput, 100, IIf(oTabla.Paciente = "", Null, oTabla.Paciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.idTipoSexo = 0, Null, oTabla.idTipoSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idMedico", adInteger, adParamInput, 0, IIf(oTabla.IdMedico = 0, Null, oTabla.IdMedico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idMovimiento", adInteger, adParamInput, 0, IIf(oTabla.idMovimiento = 0, Null, oTabla.idMovimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, IIf(oTabla.IdEstado = 0, Null, oTabla.IdEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, IIf(oTabla.IdReceta = 0, Null, oTabla.IdReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdSala", adInteger, adParamInput, 0, oTabla.IdSala)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@llaveTicket", adVarChar, adParamInput, 20, IIf(oTabla.llaveTicket = "", Null, oTabla.llaveTicket))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cupo", adInteger, adParamInput, 0, oTabla.cupo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Telefono", adVarChar, adParamInput, 10, IIf(oTabla.Telefono = "", Null, oTabla.Telefono))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Direccion", adVarChar, adParamInput, 100, IIf(oTabla.Direccion = "", Null, oTabla.Direccion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProgramacion", adInteger, adParamInput, 0, oTabla.idProgramacion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdResponsable", adInteger, adParamInput, 0, oTabla.IdResponsable)
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
'   Descripción:    elimia o borra un registro a la tabla SiCitas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoSiCitas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SiCitasEliminar"
   Set oParameter = .CreateParameter("@idCitaSI", adInteger, adParamInput, 0, oTabla.idCitaSI)
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

Sub CargaTabla(ByRef oTabla As DoSiCitas, oRecordset As Recordset)
       oTabla.idCitaSI = IIf(IsNull(oRecordset!idCitaSI), 0, oRecordset!idCitaSI)
       oTabla.IdPuntoCarga = IIf(IsNull(oRecordset!IdPuntoCarga), 0, oRecordset!IdPuntoCarga)
       oTabla.Fecha = IIf(IsNull(oRecordset!Fecha), 0, oRecordset!Fecha)
       oTabla.HoraInicio = IIf(IsNull(oRecordset!HoraInicio), "", oRecordset!HoraInicio)
       oTabla.HoraFinal = IIf(IsNull(oRecordset!HoraFinal), "", oRecordset!HoraFinal)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.Paciente = IIf(IsNull(oRecordset!Paciente), "", oRecordset!Paciente)
       oTabla.idTipoSexo = IIf(IsNull(oRecordset!idTipoSexo), 0, oRecordset!idTipoSexo)
       oTabla.FechaNacimiento = IIf(IsNull(oRecordset!FechaNacimiento), 0, oRecordset!FechaNacimiento)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.IdMedico = IIf(IsNull(oRecordset!IdMedico), 0, oRecordset!IdMedico)
       oTabla.idMovimiento = IIf(IsNull(oRecordset!idMovimiento), 0, oRecordset!idMovimiento)
       oTabla.IdUsuario = IIf(IsNull(oRecordset!IdUsuario), 0, oRecordset!IdUsuario)
       oTabla.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
       oTabla.IdReceta = IIf(IsNull(oRecordset!IdReceta), 0, oRecordset!IdReceta)
       oTabla.IdEstado = IIf(IsNull(oRecordset!IdEstado), 0, oRecordset!IdEstado)
       oTabla.IdSala = IIf(IsNull(oRecordset!IdSala), 0, oRecordset!IdSala)
       oTabla.llaveTicket = IIf(IsNull(oRecordset!llaveTicket), "", oRecordset!llaveTicket)
       oTabla.cupo = IIf(IsNull(oRecordset!cupo), 0, oRecordset!cupo)
       oTabla.Direccion = IIf(IsNull(oRecordset!Direccion), "", oRecordset!Direccion)
       oTabla.Telefono = IIf(IsNull(oRecordset!Telefono), "", oRecordset!Telefono)
       oTabla.idProgramacion = IIf(IsNull(oRecordset!idProgramacion), 0, oRecordset!idProgramacion)
       oTabla.IdResponsable = IIf(IsNull(oRecordset!IdResponsable), 0, oRecordset!IdResponsable)

End Sub

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla SiCitas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoSiCitas) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "SiCitasSeleccionarPorId"
     Set oParameter = .CreateParameter("@idCitaSI", adInteger, adParamInput, 0, oTabla.idCitaSI)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
   CargaTabla oTabla, oRecordset
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoSiCitas)
   oTabla.idCitaSI = 0
   oTabla.IdPuntoCarga = 0
   oTabla.Fecha = Now
   oTabla.HoraInicio = ""
   oTabla.HoraFinal = ""
   oTabla.IdCuentaAtencion = 0
   oTabla.IdComprobantePago = 0
   oTabla.Paciente = ""
   oTabla.idTipoSexo = 0
   oTabla.FechaNacimiento = Now
   oTabla.IdPaciente = 0
   oTabla.IdMedico = 0
   oTabla.idMovimiento = 0
   oTabla.IdUsuario = 0
   oTabla.FechaCreacion = Now
   oTabla.IdEstado = 0
End Sub

Function SiCitasSeleccionarPorIdMovimiento(ByVal oTabla As DoSiCitas) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SiCitasSeleccionarPorIdMovimiento = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "SiCitasSeleccionarPorIdMovimiento"
     Set oParameter = .CreateParameter("@idMovimiento", adInteger, adParamInput, 0, oTabla.idMovimiento)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
   CargaTabla oTabla, oRecordset
   SiCitasSeleccionarPorIdMovimiento = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: SiCitas
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: SiCitas
'        Autor:   DAniel Barrantes B.
'        Fecha: 13/11/2018 09:19:00 a.m.
'------------------------------------------------------------------------------------
 

