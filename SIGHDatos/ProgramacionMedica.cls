VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProgramacionMedica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla ProgramacionMedica
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
'   Descripción:    Inserta un registro a la tabla ProgramacionMedica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOProgramacionMedica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ProgramacionMedicaAgregar"
           Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidad = 0, Null, oTabla.IdEspecialidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Color", adInteger, adParamInput, 0, IIf(oTabla.Color = 0, Null, oTabla.Color)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IIf(oTabla.IdMedico = 0, Null, oTabla.IdMedico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, IIf(oTabla.IdDepartamento = 0, Null, oTabla.IdDepartamento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, IIf(oTabla.IdTipoServicio = 0, Null, oTabla.IdTipoServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Fecha = 0, Null, oTabla.Fecha)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraInicio", adChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraFin", adChar, adParamInput, 5, IIf(oTabla.HoraFin = "", Null, oTabla.HoraFin)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 100, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoProgramacion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoProgramacion = 0, Null, oTabla.IdTipoProgramacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaReg", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaReg = 0, Null, oTabla.FechaReg)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TiempoPromedioAtencion", adInteger, adParamInput, 0, IIf(oTabla.TiempoPromedioAtencion = 0, Null, oTabla.TiempoPromedioAtencion)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdProgramacion = .Parameters("@IdProgramacion")
   End With
 
   Insertar = True
   ms_MensajeError = ""
   GrabaHistorico oTabla
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Sub GrabaHistorico(oTabla As DOProgramacionMedica)
    On Error GoTo ErrHistorico
    Dim lcSql As String
    Dim oRsParametros As New Recordset
    Dim oRsMDB As New Recordset
    Dim oConexionMDB As New Connection
    lcSql = "select ValorTexto from Parametros where idparametro=581"
    oRsParametros.Open lcSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    oConexionMDB.CommandTimeout = 900
    oConexionMDB.CursorLocation = adUseClient
    oConexionMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & Trim(oRsParametros!ValorTexto) & "\parametros.mdb"
    lcSql = "INSERT INTO ProgramacionMedica" & _
"           (idProgramacion,IdMedico" & _
"           ,IdDepartamento" & _
"           ,Fecha" & _
"           ,HoraInicio" & _
"           ,HoraFin" & _
"           ,IdTipoProgramacion" & _
"           ,Descripcion" & _
"           ,IdTurno" & _
"           ,IdEspecialidad" & _
"           ,IdServicio" & _
"           ,IdTipoServicio" & _
"           ,FechaReg" & _
"           ,TiempoPromedioAtencion)" & _
"     Values ("
lcSql = lcSql & oTabla.IdProgramacion & ", " & oTabla.IdMedico & _
     "           ," & oTabla.IdDepartamento & _
     "           ,'" & oTabla.Fecha & _
     "           ','" & oTabla.HoraInicio & _
     "           ','" & oTabla.HoraFin & _
     "           '," & oTabla.IdTipoProgramacion & _
     "           ,'" & oTabla.Descripcion & _
     "           '," & oTabla.IdTurno & _
     "           ," & oTabla.IdEspecialidad & _
     "           ," & oTabla.IdServicio & _
     "           ," & oTabla.IdTipoServicio & _
     "           ,'" & oTabla.FechaReg & _
     "           '," & oTabla.TiempoPromedioAtencion & _
     ")"
   oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    oConexionMDB.Close
ErrHistorico:
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing
End Sub

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla ProgramacionMedica
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOProgramacionMedica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ProgramacionMedicaModificar"
           Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidad = 0, Null, oTabla.IdEspecialidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Color", adInteger, adParamInput, 0, IIf(oTabla.Color = 0, Null, oTabla.Color)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamInput, 0, IIf(oTabla.IdProgramacion = 0, Null, oTabla.IdProgramacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IIf(oTabla.IdMedico = 0, Null, oTabla.IdMedico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, IIf(oTabla.IdDepartamento = 0, Null, oTabla.IdDepartamento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, IIf(oTabla.IdTipoServicio = 0, Null, oTabla.IdTipoServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Fecha = 0, Null, oTabla.Fecha)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraInicio", adChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraFin", adChar, adParamInput, 5, IIf(oTabla.HoraFin = "", Null, oTabla.HoraFin)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 100, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoProgramacion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoProgramacion = 0, Null, oTabla.IdTipoProgramacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TiempoPromedioAtencion", adInteger, adParamInput, 0, IIf(oTabla.TiempoPromedioAtencion = 0, Null, oTabla.TiempoPromedioAtencion)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   Modificar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla ProgramacionMedica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOProgramacionMedica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ProgramacionMedicaEliminar"
           Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamInput, 0, IIf(oTabla.IdProgramacion = 0, Null, oTabla.IdProgramacion)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   Eliminar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla ProgramacionMedica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOProgramacionMedica) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ProgramacionMedicaSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamInput, 0, oTabla.IdProgramacion): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdEspecialidad = IIf(IsNull(oRecordset!IdEspecialidad), 0, oRecordset!IdEspecialidad)
       oTabla.IdTurno = IIf(IsNull(oRecordset!IdTurno), 0, oRecordset!IdTurno)
       oTabla.Color = IIf(IsNull(oRecordset!Color), 0, oRecordset!Color)
       oTabla.IdServicio = IIf(IsNull(oRecordset!IdServicio), 0, oRecordset!IdServicio)
       oTabla.IdProgramacion = IIf(IsNull(oRecordset!IdProgramacion), 0, oRecordset!IdProgramacion)
       oTabla.IdMedico = IIf(IsNull(oRecordset!IdMedico), 0, oRecordset!IdMedico)
       oTabla.IdDepartamento = IIf(IsNull(oRecordset!IdDepartamento), 0, oRecordset!IdDepartamento)
       oTabla.IdTipoServicio = IIf(IsNull(oRecordset!IdTipoServicio), 0, oRecordset!IdTipoServicio)
       oTabla.Fecha = IIf(IsNull(oRecordset!Fecha), 0, oRecordset!Fecha)
       oTabla.HoraInicio = IIf(IsNull(oRecordset!HoraInicio), "", oRecordset!HoraInicio)
       oTabla.HoraFin = IIf(IsNull(oRecordset!HoraFin), "", oRecordset!HoraFin)
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
       oTabla.IdTipoProgramacion = IIf(IsNull(oRecordset!IdTipoProgramacion), 0, oRecordset!IdTipoProgramacion)
       oTabla.FechaReg = IIf(IsNull(oRecordset!FechaReg), 0, oRecordset!FechaReg)
       oTabla.TiempoPromedioAtencion = IIf(IsNull(oRecordset!TiempoPromedioAtencion), 0, oRecordset!TiempoPromedioAtencion)
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function EliminarPorMedicoYMes(ByVal lIdMedico As Long, iMes As Integer, iAnio As Integer, lIdUsuario As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   EliminarPorMedicoYMes = False
   With oCommand
       .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
       .CommandText = "ProgramacionMedicaEliminarPorMedicoYMes"
        Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lIdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, iMes): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, iAnio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, lIdUsuario): .Parameters.Append oParameter
       .Execute
   End With
 
   EliminarPorMedicoYMes = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarPorMedicoYMes(ByVal lIdMedico As Long, iMes As Integer, iAnio As Integer) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorMedicoYMes = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "ProgramacionMedicaSeleccionarPorMedicoYMes"
        Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lIdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, iMes): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, iAnio): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorMedicoYMes = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Function SeleccionarDiasDeCEProgPorMedicoYMes(ByVal lIdMedico As Long, iMes As Integer, iAnio As Integer) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ProgramacionMedicaXMedicoMesAnio"
        Set oParameter = .CreateParameter("@lIdMedico", adInteger, adParamInput, 0, lIdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@iMes", adInteger, adParamInput, 0, iMes): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@iAnio", adInteger, adParamInput, 0, iAnio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarDiasDeCEProgPorMedicoYMes = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarPorMedico(lIdMedico As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ProgramacionMedicaXmedico"
        Set oParameter = .CreateParameter("@lIdMedico", adInteger, adParamInput, 0, lIdMedico): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorMedico = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
'mgaray20141021
Function ModificarHoraFinPorCitaAdicional(ByVal oTabla As DOProgramacionMedica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   ModificarHoraFinPorCitaAdicional = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ProgramacionMedicaModificarHoraPorCitaAdicional"
           Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamInput, 0, IIf(oTabla.IdProgramacion = 0, Null, oTabla.IdProgramacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraFin", adChar, adParamInput, 5, IIf(oTabla.HoraFin = "", Null, oTabla.HoraFin)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   ModificarHoraFinPorCitaAdicional = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

