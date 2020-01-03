VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Citas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla Citas
'        Programado por: Castro W
'        Fecha: Agosto 2006
'
'------------------------------------------------------------------------------------
Option Explicit
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
'   Descripción:    Inserta un registro a la tabla Citas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DoCita) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CitasAgregar"
           Set oParameter = .CreateParameter("@HoraSolicitud", adChar, adParamInput, 5, IIf(oTabla.HoraSolicitud = "", Null, oTabla.HoraSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaSolicitud", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaSolicitud = 0, Null, oTabla.FechaSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamInput, 0, IIf(oTabla.IdProgramacion = 0, Null, oTabla.IdProgramacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraFin", adChar, adParamInput, 5, IIf(oTabla.HoraFin = "", Null, oTabla.HoraFin)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraInicio", adChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCita", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 4, IIf(oTabla.Fecha = 0, Null, oTabla.Fecha)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoCita", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoCita = 0, Null, oTabla.IdEstadoCita)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IIf(oTabla.IdMedico = 0, Null, oTabla.IdMedico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidad = 0, Null, oTabla.IdEspecialidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsCitaAdicional", adBoolean, adParamInput, 0, IIf(oTabla.EsCitaAdicional = True, 1, 0)): .Parameters.Append oParameter
           
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdCita = .Parameters("@IdCita")
   End With
 
   Insertar = True
   ms_MensajeError = ""
   GrabaHistorico oTabla
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


Sub GrabaHistorico(oTabla As DoCita)
    On Error GoTo ErrHistorico
    Dim lcSql As String
    Dim oRsParametros As New Recordset
    Dim oRsMDB As New Recordset
    Dim oConexionMDB As New Connection
    lcSql = "select ValorTexto from Parametros where idparametro=581"
    oRsParametros.Open lcSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    oConexionMDB.CommandTimeout = 900
    oConexionMDB.CursorLocation = adUseClient
    oConexionMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & Trim(oRsParametros!valorTexto) & "\parametros.mdb"
    lcSql = "INSERT INTO Citas" & _
"           (IdCita,Fecha" & _
"           ,HoraInicio,HoraFin" & _
"           ,IdPaciente,IdEstadoCita" & _
"           ,IdAtencion,IdMedico" & _
"           ,IdEspecialidad,IdServicio" & _
"           ,IdProgramacion,IdProducto" & _
"           ,FechaSolicitud,HoraSolicitud" & _
"           ,EsCitaAdicional)" & _
"     Values" & _
"           (" & oTabla.IdCita & ",'" & oTabla.Fecha & "'" & _
"           ,'" & oTabla.HoraInicio & "','" & oTabla.HoraFin & "'" & _
"           ," & oTabla.IdPaciente & "," & oTabla.IdEstadoCita & _
"           ," & oTabla.idAtencion & "," & oTabla.IdMedico & _
"           ," & oTabla.IdEspecialidad & "," & oTabla.IdServicio & _
"           ," & oTabla.IdProgramacion & "," & oTabla.idProducto & _
"           ,'" & oTabla.FechaSolicitud & "','" & oTabla.HoraSolicitud & "'" & _
"           ," & IIf(oTabla.EsCitaAdicional = True, "1", "0") & ")"
    oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    oConexionMDB.Close
ErrHistorico:
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing
End Sub

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla Citas
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoCita) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CitasModificar"
           Set oParameter = .CreateParameter("@HoraSolicitud", adChar, adParamInput, 5, IIf(oTabla.HoraSolicitud = "", Null, oTabla.HoraSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaSolicitud", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaSolicitud = 0, Null, oTabla.FechaSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamInput, 0, IIf(oTabla.IdProgramacion = 0, Null, oTabla.IdProgramacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraFin", adChar, adParamInput, 5, IIf(oTabla.HoraFin = "", Null, oTabla.HoraFin)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraInicio", adChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCita", adInteger, adParamInput, 0, IIf(oTabla.IdCita = 0, Null, oTabla.IdCita)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 4, IIf(oTabla.Fecha = 0, Null, oTabla.Fecha)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoCita", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoCita = 0, Null, oTabla.IdEstadoCita)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IIf(oTabla.IdMedico = 0, Null, oTabla.IdMedico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidad = 0, Null, oTabla.IdEspecialidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsCitaAdicional", adBoolean, adParamInput, 0, IIf(oTabla.EsCitaAdicional = True, 1, 0)): .Parameters.Append oParameter
       
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
'   Descripción:    elimia o borra un registro a la tabla Citas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoCita) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CitasEliminar"
           Set oParameter = .CreateParameter("@IdCita", adInteger, adParamInput, 0, IIf(oTabla.IdCita = 0, Null, oTabla.IdCita)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla Citas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoCita) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CitasSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdCita", adInteger, adParamInput, 0, oTabla.IdCita): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.HoraSolicitud = IIf(IsNull(oRecordset!HoraSolicitud), "", oRecordset!HoraSolicitud)
       oTabla.FechaSolicitud = IIf(IsNull(oRecordset!FechaSolicitud), 0, oRecordset!FechaSolicitud)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.IdProgramacion = IIf(IsNull(oRecordset!IdProgramacion), 0, oRecordset!IdProgramacion)
       oTabla.IdServicio = IIf(IsNull(oRecordset!IdServicio), 0, oRecordset!IdServicio)
       oTabla.HoraFin = IIf(IsNull(oRecordset!HoraFin), "", oRecordset!HoraFin)
       oTabla.HoraInicio = IIf(IsNull(oRecordset!HoraInicio), "", oRecordset!HoraInicio)
       oTabla.IdCita = IIf(IsNull(oRecordset!IdCita), 0, oRecordset!IdCita)
       oTabla.Fecha = IIf(IsNull(oRecordset!Fecha), 0, oRecordset!Fecha)
       oTabla.IdEstadoCita = IIf(IsNull(oRecordset!IdEstadoCita), 0, oRecordset!IdEstadoCita)
       oTabla.IdMedico = IIf(IsNull(oRecordset!IdMedico), 0, oRecordset!IdMedico)
       oTabla.IdEspecialidad = IIf(IsNull(oRecordset!IdEspecialidad), 0, oRecordset!IdEspecialidad)
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.EsCitaAdicional = IIf(IsNull(oRecordset!EsCitaAdicional), 0, oRecordset!EsCitaAdicional)
       'En sighDatos.Atenciones, tambien se carga los mismos datos, si hay algun cambio aquí también se
       'hará en sighDatos.Atenciones
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

Function SeleccionarPorMedicoYFecha(ByVal lIdMedico As Long, daFecha As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorMedicoYFecha = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "CitasSeleccionarPorMedicoYFecha"
        Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lIdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(daFecha = 0, Null, daFecha)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorMedicoYFecha = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Function SeleccionarBloqueadasPorMedicoYFecha(ByVal lIdMedico As Long, daFecha As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarBloqueadasPorMedicoYFecha = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "CitasSeleccionarBloqueadasPorMedicoYFecha"
        Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lIdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(daFecha = 0, Null, daFecha)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarBloqueadasPorMedicoYFecha = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarPorDisponiblesPorMedicoYFecha(ByVal lIdMedico As Long, daFecha As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorDisponiblesPorMedicoYFecha = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "CitasSeleccionarDisponiblesPorMedicoYFecha"
        Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lIdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(daFecha = 0, Null, daFecha)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorDisponiblesPorMedicoYFecha = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Function SeleccionarPorDisponiblesPorMedicoEspecialidadYFecha(ByVal lIdMedico As Long, lIdEspecialidad As Long, daFecha As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorDisponiblesPorMedicoEspecialidadYFecha = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "CitasSeleccionarDisponiblesPorMedicoEspecialidadYFecha"
        Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lIdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, lIdEspecialidad): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(daFecha = 0, Null, daFecha)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorDisponiblesPorMedicoEspecialidadYFecha = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Function SeleccionarPacientesPorMedicoYFecha(ByVal lIdMedico As Long, daFecha As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPacientesPorMedicoYFecha = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "CitasSeleccionarPacientePorMedicoYFecha"
        Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lIdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(daFecha = 0, Null, daFecha)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPacientesPorMedicoYFecha = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarPorPacienteYFecha(ByVal lIdPaciente As Long, daFecha As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorPacienteYFecha = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "CitasSeleccionarPorPacienteYFecha"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lIdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(daFecha = 0, Null, daFecha)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorPacienteYFecha = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function ActualizarCitaPagada(lIdAtencion As Long, lIdUsuarioAuditoria As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   ActualizarCitaPagada = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CitasActualizarCitaPagada"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, lIdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
 
   ActualizarCitaPagada = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Function ActualizarCitaPagadaDEBB(lIdAtencion As Long, lIdUsuarioAuditoria As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   ActualizarCitaPagadaDEBB = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CitasActualizarCitaPagadaDEBB"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, lIdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
 
   ActualizarCitaPagadaDEBB = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


