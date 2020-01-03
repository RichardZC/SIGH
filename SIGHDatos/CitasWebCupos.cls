VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CitasWebCupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla CitasWebCupos
'        Programado por: Barrantes D
'        Fecha: Agosto 2013
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
'   Descripción:    Inserta un registro a la tabla CitasWebCupos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOCitasWebCupos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "CitasWebCuposAgregar"
   Set oParameter = .CreateParameter("@idWeb", adInteger, adParamInput, 0, IIf(oTabla.idWeb = 0, Null, oTabla.idWeb))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FECHA = 0, Null, oTabla.FECHA))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idMedico", adInteger, adParamInput, 0, IIf(oTabla.IdMedico = 0, Null, oTabla.IdMedico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraInicio", adVarChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraFinal", adVarChar, adParamInput, 5, IIf(oTabla.HoraFinal = "", Null, oTabla.HoraFinal))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoCitaWeb", adInteger, adParamInput, 0, IIf(oTabla.idEstadoCitaWeb = 0, Null, oTabla.idEstadoCitaWeb))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCitaBloqueada", adInteger, adParamInput, 0, IIf(oTabla.idCitaBloqueada = 0, Null, oTabla.idCitaBloqueada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DNI", adVarChar, adParamInput, 8, IIf(oTabla.DNI = "", Null, oTabla.DNI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, IIf(oTabla.ApellidoPaterno = "", Null, oTabla.ApellidoPaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, IIf(oTabla.ApellidoMaterno = "", Null, oTabla.ApellidoMaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, IIf(oTabla.PrimerNombre = "", Null, oTabla.PrimerNombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SegundoNombre", adVarChar, adParamInput, 40, IIf(oTabla.SegundoNombre = "", Null, oTabla.SegundoNombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSexo = 0, Null, oTabla.IdTipoSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Ubigeo", adInteger, adParamInput, 0, IIf(oTabla.Ubigeo = 0, Null, oTabla.Ubigeo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaConfirmacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaConfirmacion = 0, Null, oTabla.FechaConfirmacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraConfirmacion", adVarChar, adParamInput, 5, IIf(oTabla.HoraConfirmacion = "", Null, oTabla.HoraConfirmacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Email", adVarChar, adParamInput, 50, IIf(oTabla.Email = "", Null, oTabla.Email))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Telefono", adVarChar, adParamInput, 10, IIf(oTabla.Telefono = "", Null, oTabla.Telefono))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   
   Set oParameter = .CreateParameter("@refIdEess", adInteger, adParamInput, 0, IIf(oTabla.refIdEess = 0, Null, oTabla.refIdEess))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@refNumero", adVarChar, adParamInput, 20, IIf(oTabla.refNumero = "", Null, oTabla.refNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@refUPS", adVarChar, adParamInput, 7, IIf(oTabla.refUPS = "", Null, oTabla.refUPS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@refidDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.refidDiagnostico = 0, Null, oTabla.refidDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@refMedicoCOLEGIATURA", adVarChar, adParamInput, 6, IIf(oTabla.refMedicoCOLEGIATURA = "", Null, oTabla.refMedicoCOLEGIATURA))
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
'   Descripción:    modifica un registro a la tabla CitasWebCupos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCitasWebCupos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "CitasWebCuposModificar"
   Set oParameter = .CreateParameter("@idWeb", adInteger, adParamInput, 0, IIf(oTabla.idWeb = 0, Null, oTabla.idWeb))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FECHA = 0, Null, oTabla.FECHA))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idMedico", adInteger, adParamInput, 0, IIf(oTabla.IdMedico = 0, Null, oTabla.IdMedico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraInicio", adVarChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraFinal", adVarChar, adParamInput, 5, IIf(oTabla.HoraFinal = "", Null, oTabla.HoraFinal))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoCitaWeb", adInteger, adParamInput, 0, IIf(oTabla.idEstadoCitaWeb = 0, Null, oTabla.idEstadoCitaWeb))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCitaBloqueada", adInteger, adParamInput, 0, IIf(oTabla.idCitaBloqueada = 0, Null, oTabla.idCitaBloqueada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DNI", adVarChar, adParamInput, 8, IIf(oTabla.DNI = "", Null, oTabla.DNI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, IIf(oTabla.ApellidoPaterno = "", Null, oTabla.ApellidoPaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, IIf(oTabla.ApellidoMaterno = "", Null, oTabla.ApellidoMaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, IIf(oTabla.PrimerNombre = "", Null, oTabla.PrimerNombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SegundoNombre", adVarChar, adParamInput, 40, IIf(oTabla.SegundoNombre = "", Null, oTabla.SegundoNombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSexo = 0, Null, oTabla.IdTipoSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Ubigeo", adInteger, adParamInput, 0, IIf(oTabla.Ubigeo = 0, Null, oTabla.Ubigeo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaConfirmacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaConfirmacion = 0, Null, oTabla.FechaConfirmacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraConfirmacion", adVarChar, adParamInput, 5, IIf(oTabla.HoraConfirmacion = "", Null, oTabla.HoraConfirmacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Email", adVarChar, adParamInput, 50, IIf(oTabla.Email = "", Null, oTabla.Email))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Telefono", adVarChar, adParamInput, 10, IIf(oTabla.Telefono = "", Null, oTabla.Telefono))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   
   Set oParameter = .CreateParameter("@refIdEess", adInteger, adParamInput, 0, IIf(oTabla.refIdEess = 0, Null, oTabla.refIdEess))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@refNumero", adVarChar, adParamInput, 20, IIf(oTabla.refNumero = "", Null, oTabla.refNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@refUPS", adVarChar, adParamInput, 7, IIf(oTabla.refUPS = "", Null, oTabla.refUPS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@refidDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.refidDiagnostico = 0, Null, oTabla.refidDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@refMedicoCOLEGIATURA", adVarChar, adParamInput, 6, IIf(oTabla.refMedicoCOLEGIATURA = "", Null, oTabla.refMedicoCOLEGIATURA))
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



Function ModificarDcitas(ByVal oTabla As DOCitasWebCupos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
ModificarDcitas = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "CitasWebCuposModificarDcitas"
   Set oParameter = .CreateParameter("@idWeb", adInteger, adParamInput, 0, IIf(oTabla.idWeb = 0, Null, oTabla.idWeb))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FECHA = 0, Null, oTabla.FECHA))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idMedico", adInteger, adParamInput, 0, IIf(oTabla.IdMedico = 0, Null, oTabla.IdMedico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraInicio", adVarChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraFinal", adVarChar, adParamInput, 5, IIf(oTabla.HoraFinal = "", Null, oTabla.HoraFinal))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoCitaWeb", adInteger, adParamInput, 0, IIf(oTabla.idEstadoCitaWeb = 0, Null, oTabla.idEstadoCitaWeb))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCitaBloqueada", adInteger, adParamInput, 0, IIf(oTabla.idCitaBloqueada = 0, Null, oTabla.idCitaBloqueada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DNI", adVarChar, adParamInput, 8, IIf(oTabla.DNI = "", Null, oTabla.DNI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, IIf(oTabla.ApellidoPaterno = "", Null, oTabla.ApellidoPaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, IIf(oTabla.ApellidoMaterno = "", Null, oTabla.ApellidoMaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, IIf(oTabla.PrimerNombre = "", Null, oTabla.PrimerNombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SegundoNombre", adVarChar, adParamInput, 40, IIf(oTabla.SegundoNombre = "", Null, oTabla.SegundoNombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSexo = 0, Null, oTabla.IdTipoSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Ubigeo", adInteger, adParamInput, 0, IIf(oTabla.Ubigeo = 0, Null, oTabla.Ubigeo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaConfirmacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaConfirmacion = 0, Null, oTabla.FechaConfirmacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraConfirmacion", adVarChar, adParamInput, 5, IIf(oTabla.HoraConfirmacion = "", Null, oTabla.HoraConfirmacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Email", adVarChar, adParamInput, 50, IIf(oTabla.Email = "", Null, oTabla.Email))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Telefono", adVarChar, adParamInput, 10, IIf(oTabla.Telefono = "", Null, oTabla.Telefono))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   
   Set oParameter = .CreateParameter("@refIdEess", adInteger, adParamInput, 0, IIf(oTabla.refIdEess = 0, Null, oTabla.refIdEess))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@refNumero", adVarChar, adParamInput, 20, IIf(oTabla.refNumero = "", Null, oTabla.refNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@refUPS", adVarChar, adParamInput, 7, IIf(oTabla.refUPS = "", Null, oTabla.refUPS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@refidDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.refidDiagnostico = 0, Null, oTabla.refidDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@refMedicoCOLEGIATURA", adVarChar, adParamInput, 6, IIf(oTabla.refMedicoCOLEGIATURA = "", Null, oTabla.refMedicoCOLEGIATURA))
   .Parameters.Append oParameter
   
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
ModificarDcitas = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function



'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla CitasWebCupos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCitasWebCupos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "CitasWebCuposEliminar"
   Set oParameter = .CreateParameter("@idWeb", adInteger, adParamInput, 0, oTabla.idWeb)
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

Sub CargaTabla(ByRef oTabla As DOCitasWebCupos, oRecordset As Recordset)
       oTabla.idWeb = IIf(IsNull(oRecordset!idWeb), 0, oRecordset!idWeb)
       oTabla.FECHA = IIf(IsNull(oRecordset!FECHA), 0, oRecordset!FECHA)
       oTabla.IdServicio = IIf(IsNull(oRecordset!IdServicio), 0, oRecordset!IdServicio)
       oTabla.IdMedico = IIf(IsNull(oRecordset!IdMedico), 0, oRecordset!IdMedico)
       oTabla.HoraInicio = IIf(IsNull(oRecordset!HoraInicio), "", oRecordset!HoraInicio)
       oTabla.HoraFinal = IIf(IsNull(oRecordset!HoraFinal), "", oRecordset!HoraFinal)
       oTabla.idEstadoCitaWeb = IIf(IsNull(oRecordset!idEstadoCitaWeb), 0, oRecordset!idEstadoCitaWeb)
       oTabla.idCitaBloqueada = IIf(IsNull(oRecordset!idCitaBloqueada), 0, oRecordset!idCitaBloqueada)
       oTabla.DNI = IIf(IsNull(oRecordset!DNI), "", oRecordset!DNI)
       oTabla.ApellidoPaterno = IIf(IsNull(oRecordset!ApellidoPaterno), "", oRecordset!ApellidoPaterno)
       oTabla.ApellidoMaterno = IIf(IsNull(oRecordset!ApellidoMaterno), "", oRecordset!ApellidoMaterno)
       oTabla.PrimerNombre = IIf(IsNull(oRecordset!PrimerNombre), "", oRecordset!PrimerNombre)
       oTabla.SegundoNombre = IIf(IsNull(oRecordset!SegundoNombre), "", oRecordset!SegundoNombre)
       oTabla.IdTipoSexo = IIf(IsNull(oRecordset!IdTipoSexo), 0, oRecordset!IdTipoSexo)
       oTabla.FechaNacimiento = IIf(IsNull(oRecordset!FechaNacimiento), 0, oRecordset!FechaNacimiento)
       oTabla.Ubigeo = IIf(IsNull(oRecordset!Ubigeo), 0, oRecordset!Ubigeo)
       oTabla.FechaConfirmacion = IIf(IsNull(oRecordset!FechaConfirmacion), 0, oRecordset!FechaConfirmacion)
       oTabla.HoraConfirmacion = IIf(IsNull(oRecordset!HoraConfirmacion), "", oRecordset!HoraConfirmacion)
       oTabla.idFuenteFinanciamiento = IIf(IsNull(oRecordset!idFuenteFinanciamiento), 0, oRecordset!idFuenteFinanciamiento)
       oTabla.IdTurno = IIf(IsNull(oRecordset!IdTurno), 0, oRecordset!IdTurno)
       oTabla.Email = IIf(IsNull(oRecordset!Email), "", oRecordset!Email)
       oTabla.Telefono = IIf(IsNull(oRecordset!Telefono), "", oRecordset!Telefono)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.refIdEess = IIf(IsNull(oRecordset!refIdEess), 0, oRecordset!refIdEess)
       oTabla.refNumero = IIf(IsNull(oRecordset!refNumero), "", oRecordset!refNumero)
       oTabla.refUPS = IIf(IsNull(oRecordset!refUPS), "", oRecordset!refUPS)
       oTabla.refidDiagnostico = IIf(IsNull(oRecordset!refidDiagnostico), 0, oRecordset!refidDiagnostico)
       oTabla.refMedicoCOLEGIATURA = IIf(IsNull(oRecordset!refMedicoCOLEGIATURA), "", oRecordset!refMedicoCOLEGIATURA)

End Sub

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla CitasWebCupos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCitasWebCupos) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CitasWebCuposSeleccionarPorId"
     Set oParameter = .CreateParameter("@idWeb", adInteger, adParamInput, 0, oTabla.idWeb)
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
Sub SetDefaults(ByVal oTabla As DOCitasWebCupos)
   oTabla.idWeb = 0
   oTabla.FECHA = Now
   oTabla.IdServicio = 0
   oTabla.IdMedico = 0
   oTabla.HoraInicio = ""
   oTabla.HoraFinal = ""
   oTabla.idEstadoCitaWeb = 0
   oTabla.idCitaBloqueada = 0
   oTabla.DNI = ""
   oTabla.ApellidoPaterno = ""
   oTabla.ApellidoMaterno = ""
   oTabla.PrimerNombre = ""
   oTabla.SegundoNombre = ""
   oTabla.IdTipoSexo = 0
   oTabla.FechaNacimiento = Now
   oTabla.Ubigeo = 0
   oTabla.FechaConfirmacion = Now
   oTabla.HoraConfirmacion = ""
   oTabla.idFuenteFinanciamiento = 0
   oTabla.IdTurno = 0
   oTabla.Email = ""
   oTabla.Telefono = ""
End Sub
 
Function SeleccionarPorIdCitaBloqueada(lnIdCitaBloqueada As Long) As DOCitasWebCupos
        On Error GoTo ErrCWE
        Dim oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oTabla As New DOCitasWebCupos
        Set SeleccionarPorIdCitaBloqueada = Nothing
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = mo_Conexion
            .CommandTimeout = 150
            .CommandText = "CitasWebCuposSeleccionarPorIdCitaBloqueada"
            Set oParameter = .CreateParameter("@IdCitaBloqueada", adInteger, adParamInput, 0, lnIdCitaBloqueada): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
        End With
        If oRsTmp1.RecordCount > 0 Then
           CargaTabla oTabla, oRsTmp1
           Set SeleccionarPorIdCitaBloqueada = oTabla
        End If
        Set oCommand = Nothing
        Exit Function
ErrCWE:
       ms_MensajeError = Err.Description
End Function


Function EliminarPorIdCitaBloqueada(ByVal oTabla As DOCitasWebCupos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
EliminarPorIdCitaBloqueada = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "CitasWebCuposEliminarPorIdCitaBloqueada"
   Set oParameter = .CreateParameter("@IdCitaBloqueada", adInteger, adParamInput, 0, oTabla.idCitaBloqueada)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
EliminarPorIdCitaBloqueada = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
