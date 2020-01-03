VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SiProgramacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: SiProgramacion
'        Autor: Daniel Enrique Barrantes
'        Fecha: 03/01/2019 05:33:02 a.m.
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
'   Descripción:    Inserta un registro a la tabla SiProgramacion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOSiProgramacion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SiProgramacionAgregar"
   Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdResponsable", adInteger, adParamInput, 0, IIf(oTabla.IdResponsable = 0, Null, oTabla.IdResponsable))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Fecha = 0, Null, oTabla.Fecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraInicio", adChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraFin", adChar, adParamInput, 5, IIf(oTabla.HoraFin = "", Null, oTabla.HoraFin))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdSala", adInteger, adParamInput, 0, IIf(oTabla.IdSala = 0, Null, oTabla.IdSala))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaReg", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaReg = 0, Null, oTabla.FechaReg))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TiempoPromedioAtencion", adInteger, adParamInput, 0, IIf(oTabla.TiempoPromedioAtencion = 0, Null, oTabla.TiempoPromedioAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdProgramacion = .Parameters("@IdProgramacion")
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
'   Descripción:    modifica un registro a la tabla SiProgramacion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOSiProgramacion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SiProgramacionModificar"
   Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamInput, 0, IIf(oTabla.IdProgramacion = 0, Null, oTabla.IdProgramacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdResponsable", adInteger, adParamInput, 0, IIf(oTabla.IdResponsable = 0, Null, oTabla.IdResponsable))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Fecha = 0, Null, oTabla.Fecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraInicio", adChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraFin", adChar, adParamInput, 5, IIf(oTabla.HoraFin = "", Null, oTabla.HoraFin))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdSala", adInteger, adParamInput, 0, IIf(oTabla.IdSala = 0, Null, oTabla.IdSala))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaReg", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaReg = 0, Null, oTabla.FechaReg))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TiempoPromedioAtencion", adInteger, adParamInput, 0, IIf(oTabla.TiempoPromedioAtencion = 0, Null, oTabla.TiempoPromedioAtencion))
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
'   Descripción:    elimia o borra un registro a la tabla SiProgramacion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOSiProgramacion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SiProgramacionEliminar"
   Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamInput, 0, oTabla.IdProgramacion)
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
'   Descripción:    Selecciona un unico registro de la tabla SiProgramacion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOSiProgramacion) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "SiProgramacionSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamInput, 0, oTabla.IdProgramacion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdProgramacion = IIf(IsNull(oRecordset!IdProgramacion), 0, oRecordset!IdProgramacion)
       oTabla.IdResponsable = IIf(IsNull(oRecordset!IdResponsable), 0, oRecordset!IdResponsable)
       oTabla.Fecha = IIf(IsNull(oRecordset!Fecha), 0, oRecordset!Fecha)
       oTabla.HoraInicio = IIf(IsNull(oRecordset!HoraInicio), "", oRecordset!HoraInicio)
       oTabla.HoraFin = IIf(IsNull(oRecordset!HoraFin), "", oRecordset!HoraFin)
       oTabla.IdTurno = IIf(IsNull(oRecordset!IdTurno), 0, oRecordset!IdTurno)
       oTabla.IdSala = IIf(IsNull(oRecordset!IdSala), 0, oRecordset!IdSala)
       oTabla.FechaReg = IIf(IsNull(oRecordset!FechaReg), 0, oRecordset!FechaReg)
       oTabla.TiempoPromedioAtencion = IIf(IsNull(oRecordset!TiempoPromedioAtencion), 0, oRecordset!TiempoPromedioAtencion)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOSiProgramacion)
   oTabla.IdProgramacion = 0
   oTabla.IdResponsable = 0
   oTabla.Fecha = Now
   oTabla.HoraInicio = ""
   oTabla.HoraFin = ""
   oTabla.IdTurno = 0
   oTabla.IdSala = 0
   oTabla.FechaReg = Now
   oTabla.TiempoPromedioAtencion = 0
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: SiProgramacion
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: SiProgramacion
'        Autor:   DAniel Barrantes B.
'        Fecha: 03/01/2019 05:33:02 a.m.
'------------------------------------------------------------------------------------
 

