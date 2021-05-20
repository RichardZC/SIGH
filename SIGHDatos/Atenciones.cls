VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Atenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla Atenciones
'        Programado por: Barrantes D
'        Fecha: Diciembre 2012
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
'   Descripción:    Inserta un registro a la tabla Atenciones
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesAgregar"
       

       
           Set oParameter = .CreateParameter("@HoraIngreso", adChar, adParamInput, 5, IIf(oTabla.HoraIngreso = "", Null, oTabla.HoraIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaIngreso", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaIngreso = 0, Null, oTabla.FechaIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, IIf(oTabla.IdTipoServicio = 0, Null, oTabla.IdTipoServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoCondicionALEstab", adInteger, adParamInput, 0, IIf(oTabla.IdTipoCondicionALEstab = 0, Null, oTabla.IdTipoCondicionALEstab)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaEgresoAdministrativo", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEgresoAdministrativo = 0, Null, oTabla.FechaEgresoAdministrativo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCamaEgreso", adInteger, adParamInput, 0, IIf(oTabla.IdCamaEgreso = 0, Null, oTabla.IdCamaEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCamaIngreso", adInteger, adParamInput, 0, IIf(oTabla.IdCamaIngreso = 0, Null, oTabla.IdCamaIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioEgreso", adInteger, adParamInput, 0, IIf(oTabla.IdServicioEgreso = 0, Null, oTabla.IdServicioEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoAlta", adInteger, adParamInput, 0, IIf(oTabla.IdTipoAlta = 0, Null, oTabla.IdTipoAlta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCondicionAlta", adInteger, adParamInput, 0, IIf(oTabla.IdCondicionAlta = 0, Null, oTabla.IdCondicionAlta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoEdad", adInteger, adParamInput, 0, IIf(oTabla.IdTipoEdad = 0, Null, oTabla.IdTipoEdad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdOrigenAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdOrigenAtencion = 0, Null, oTabla.IdOrigenAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDestinoAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdDestinoAtencion = 0, Null, oTabla.IdDestinoAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraEgresoAdministrativo", adChar, adParamInput, 5, IIf(oTabla.HoraEgresoAdministrativo = "", Null, oTabla.HoraEgresoAdministrativo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoCondicionAlServicio", adInteger, adParamInput, 0, IIf(oTabla.IdTipoCondicionAlServicio = 0, Null, oTabla.IdTipoCondicionAlServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraEgreso", adChar, adParamInput, 5, IIf(oTabla.HoraEgreso = "", Null, oTabla.HoraEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaEgreso", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEgreso = 0, Null, oTabla.FechaEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoEgreso", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoEgreso = 0, Null, oTabla.IdMedicoEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Edad", adInteger, adParamInput, 0, IIf(oTabla.Edad = 0, Null, oTabla.Edad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEspecialidadMedico", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidadMedico = 0, Null, oTabla.IdEspecialidadMedico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoIngreso", adInteger, adParamInput, 0, oTabla.IdMedicoIngreso): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioIngreso", adInteger, adParamInput, 0, IIf(oTabla.IdServicioIngreso = 0, Null, oTabla.IdServicioIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoGravedad", adInteger, adParamInput, 0, IIf(oTabla.IdTipoGravedad = 0, Null, oTabla.IdTipoGravedad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idFormaPago", adInteger, adParamInput, 0, IIf(oTabla.idFormaPago = 0, Null, oTabla.idFormaPago)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idEstadoAtencion", adInteger, adParamInput, 0, oTabla.idEstadoAtencion): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsPacienteExterno", adBoolean, adParamInput, 0, IIf(oTabla.EsPacienteExterno = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idSunasaPacienteHistorico", adInteger, adParamInput, 0, IIf(oTabla.idSunasaPacienteHistorico = 0, Null, oTabla.idSunasaPacienteHistorico)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraInicioAtencion", adChar, adParamInput, 5, IIf(oTabla.HoraInicioAtencion = "", Null, oTabla.HoraInicioAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EpsPorcentaje", adCurrency, adParamInput, 0, oTabla.EpsPorcentaje): .Parameters.Append oParameter
           
       .Execute
           oTabla.idAtencion = .Parameters("@IdAtencion")
   End With
 
   Insertar = True
   ms_MensajeError = ""
   GrabaHistorico oTabla
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "Insertar"
   
Exit Function
End Function

Sub GrabaHistorico(oTabla As DOAtencion)
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
    lcSql = "INSERT INTO Atenciones" & _
"           (IdPaciente,Edad,FechaIngreso,HoraIngreso" & _
"           ,IdDestinoAtencion,IdTipoCondicionAlServicio,IdTipoCondicionALEstab,IdServicioIngreso" & _
"           ,IdMedicoIngreso,IdEspecialidadMedico,IdMedicoEgreso,FechaEgreso" & _
"           ,HoraEgreso,IdOrigenAtencion,FechaEgresoAdministrativo,HoraEgresoAdministrativo" & _
"           ,IdCondicionAlta,IdTipoAlta,IdServicioEgreso" & _
"           ,IdCamaIngreso,IdCamaEgreso,IdTipoGravedad" & _
"           ,IdTipoEdad,IdCuentaAtencion,IdTipoServicio" & _
"           ,IdFormaPago,idFuenteFinanciamiento" & _
"           ,idEstadoAtencion,EsPacienteExterno,idSunasaPacienteHistorico,idAtencion)" & _
"     Values(" & _
"            " & oTabla.IdPaciente & "," & oTabla.Edad & ",'" & oTabla.FechaIngreso & "','" & oTabla.HoraIngreso & "'" & _
"            ," & oTabla.IdDestinoAtencion & "," & oTabla.IdTipoCondicionAlServicio & "," & oTabla.IdTipoCondicionALEstab & "," & oTabla.IdServicioIngreso & _
"            ," & oTabla.IdMedicoIngreso & "," & oTabla.IdEspecialidadMedico & "," & oTabla.IdMedicoEgreso & ",'" & oTabla.FechaEgreso & "'" & _
"            ,'" & oTabla.HoraEgreso & "'," & oTabla.IdOrigenAtencion & ",'" & oTabla.FechaEgresoAdministrativo & "','" & oTabla.HoraEgresoAdministrativo & "'" & _
"            ," & oTabla.IdCondicionAlta & "," & oTabla.IdTipoAlta & "," & oTabla.IdServicioEgreso & _
"            ," & oTabla.IdCamaIngreso & "," & oTabla.IdCamaEgreso & "," & oTabla.IdTipoGravedad & _
"            ," & oTabla.IdTipoEdad & "," & oTabla.IdCuentaAtencion & "," & oTabla.IdTipoServicio & _
"            ," & oTabla.idFormaPago & "," & oTabla.idFuenteFinanciamiento & _
"            ," & oTabla.idEstadoAtencion & "," & IIf(oTabla.EsPacienteExterno = True, "1", "0") & "," & oTabla.idSunasaPacienteHistorico & "," & oTabla.idAtencion & _
"           )"
   oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    oConexionMDB.Close
ErrHistorico:
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing

End Sub

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla Atenciones
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesModificar"
           Set oParameter = .CreateParameter("@HoraIngreso", adChar, adParamInput, 5, IIf(oTabla.HoraIngreso = "", Null, oTabla.HoraIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaIngreso", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaIngreso = 0, Null, oTabla.FechaIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, IIf(oTabla.IdTipoServicio = 0, Null, oTabla.IdTipoServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoCondicionALEstab", adInteger, adParamInput, 0, IIf(oTabla.IdTipoCondicionALEstab = 0, Null, oTabla.IdTipoCondicionALEstab)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaEgresoAdministrativo", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEgresoAdministrativo = 0, Null, oTabla.FechaEgresoAdministrativo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCamaEgreso", adInteger, adParamInput, 0, IIf(oTabla.IdCamaEgreso = 0, Null, oTabla.IdCamaEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCamaIngreso", adInteger, adParamInput, 0, IIf(oTabla.IdCamaIngreso = 0, Null, oTabla.IdCamaIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioEgreso", adInteger, adParamInput, 0, IIf(oTabla.IdServicioEgreso = 0, Null, oTabla.IdServicioEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoAlta", adInteger, adParamInput, 0, IIf(oTabla.IdTipoAlta = 0, Null, oTabla.IdTipoAlta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCondicionAlta", adInteger, adParamInput, 0, IIf(oTabla.IdCondicionAlta = 0, Null, oTabla.IdCondicionAlta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoEdad", adInteger, adParamInput, 0, IIf(oTabla.IdTipoEdad = 0, Null, oTabla.IdTipoEdad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdOrigenAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdOrigenAtencion = 0, Null, oTabla.IdOrigenAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDestinoAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdDestinoAtencion = 0, Null, oTabla.IdDestinoAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraEgresoAdministrativo", adChar, adParamInput, 5, IIf(oTabla.HoraEgresoAdministrativo = "", Null, oTabla.HoraEgresoAdministrativo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoCondicionAlServicio", adInteger, adParamInput, 0, IIf(oTabla.IdTipoCondicionAlServicio = 0, Null, oTabla.IdTipoCondicionAlServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraEgreso", adChar, adParamInput, 5, IIf(oTabla.HoraEgreso = "", Null, oTabla.HoraEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaEgreso", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEgreso = 0, Null, oTabla.FechaEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoEgreso", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoEgreso = 0, Null, oTabla.IdMedicoEgreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Edad", adInteger, adParamInput, 0, IIf(oTabla.Edad = 0, Null, oTabla.Edad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEspecialidadMedico", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidadMedico = 0, Null, oTabla.IdEspecialidadMedico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoIngreso", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoIngreso = 0, Null, oTabla.IdMedicoIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioIngreso", adInteger, adParamInput, 0, IIf(oTabla.IdServicioIngreso = 0, Null, oTabla.IdServicioIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoGravedad", adInteger, adParamInput, 0, IIf(oTabla.IdTipoGravedad = 0, Null, oTabla.IdTipoGravedad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idFormaPago", adInteger, adParamInput, 0, IIf(oTabla.idFormaPago = 0, Null, oTabla.idFormaPago)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idEstadoAtencion", adInteger, adParamInput, 0, oTabla.idEstadoAtencion): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsPacienteExterno", adBoolean, adParamInput, 0, IIf(oTabla.EsPacienteExterno = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idSunasaPacienteHistorico", adInteger, adParamInput, 0, IIf(oTabla.idSunasaPacienteHistorico = 0, Null, oTabla.idSunasaPacienteHistorico)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraInicioAtencion", adChar, adParamInput, 5, IIf(oTabla.HoraInicioAtencion = "", Null, oTabla.HoraInicioAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EpsPorcentaje", adCurrency, adParamInput, 0, oTabla.EpsPorcentaje): .Parameters.Append oParameter
       
       .Execute
   End With
 
   Modificar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   MsgBox Err.Description
   Trace Err.Number & " " + Err.Description, "Modificar"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla Atenciones
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesEliminar"
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   Eliminar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "Eliminar"
Exit Function
End Function

Sub CargaTabla(ByRef oTabla As DOAtencion, oRecordset As Recordset)
       oTabla.HoraIngreso = IIf(IsNull(oRecordset!HoraIngreso), "", oRecordset!HoraIngreso)
       oTabla.FechaIngreso = IIf(IsNull(oRecordset!FechaIngreso), 0, oRecordset!FechaIngreso)
       oTabla.IdTipoServicio = IIf(IsNull(oRecordset!IdTipoServicio), 0, oRecordset!IdTipoServicio)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       oTabla.IdTipoCondicionALEstab = IIf(IsNull(oRecordset!IdTipoCondicionALEstab), 0, oRecordset!IdTipoCondicionALEstab)
       oTabla.FechaEgresoAdministrativo = IIf(IsNull(oRecordset!FechaEgresoAdministrativo), 0, oRecordset!FechaEgresoAdministrativo)
       oTabla.IdCamaEgreso = IIf(IsNull(oRecordset!IdCamaEgreso), 0, oRecordset!IdCamaEgreso)
       oTabla.IdCamaIngreso = IIf(IsNull(oRecordset!IdCamaIngreso), 0, oRecordset!IdCamaIngreso)
       oTabla.IdServicioEgreso = IIf(IsNull(oRecordset!IdServicioEgreso), 0, oRecordset!IdServicioEgreso)
       oTabla.IdTipoAlta = IIf(IsNull(oRecordset!IdTipoAlta), 0, oRecordset!IdTipoAlta)
       oTabla.IdCondicionAlta = IIf(IsNull(oRecordset!IdCondicionAlta), 0, oRecordset!IdCondicionAlta)
       oTabla.IdTipoEdad = IIf(IsNull(oRecordset!IdTipoEdad), 0, oRecordset!IdTipoEdad)
       oTabla.IdOrigenAtencion = IIf(IsNull(oRecordset!IdOrigenAtencion), 0, oRecordset!IdOrigenAtencion)
       oTabla.IdDestinoAtencion = IIf(IsNull(oRecordset!IdDestinoAtencion), 0, oRecordset!IdDestinoAtencion)
       oTabla.HoraEgresoAdministrativo = IIf(IsNull(oRecordset!HoraEgresoAdministrativo), "", oRecordset!HoraEgresoAdministrativo)
       oTabla.IdTipoCondicionAlServicio = IIf(IsNull(oRecordset!IdTipoCondicionAlServicio), 0, oRecordset!IdTipoCondicionAlServicio)
       oTabla.HoraEgreso = IIf(IsNull(oRecordset!HoraEgreso), "", oRecordset!HoraEgreso)
       oTabla.FechaEgreso = IIf(IsNull(oRecordset!FechaEgreso), 0, oRecordset!FechaEgreso)
       oTabla.IdMedicoEgreso = IIf(IsNull(oRecordset!IdMedicoEgreso), 0, oRecordset!IdMedicoEgreso)
       oTabla.Edad = IIf(IsNull(oRecordset!Edad), 0, oRecordset!Edad)
       oTabla.IdEspecialidadMedico = IIf(IsNull(oRecordset!IdEspecialidadMedico), 0, oRecordset!IdEspecialidadMedico)
       oTabla.IdMedicoIngreso = IIf(IsNull(oRecordset!IdMedicoIngreso), 0, oRecordset!IdMedicoIngreso)
       oTabla.IdServicioIngreso = IIf(IsNull(oRecordset!IdServicioIngreso), 0, oRecordset!IdServicioIngreso)
       oTabla.IdTipoGravedad = IIf(IsNull(oRecordset!IdTipoGravedad), 0, oRecordset!IdTipoGravedad)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.idFormaPago = IIf(IsNull(oRecordset!idFormaPago), 0, oRecordset!idFormaPago)
       oTabla.idFuenteFinanciamiento = IIf(IsNull(oRecordset!idFuenteFinanciamiento), 0, oRecordset!idFuenteFinanciamiento)
       oTabla.idEstadoAtencion = IIf(IsNull(oRecordset!idEstadoAtencion), 0, oRecordset!idEstadoAtencion)
       oTabla.EsPacienteExterno = IIf(IsNull(oRecordset!EsPacienteExterno), 0, oRecordset!EsPacienteExterno)
       oTabla.idSunasaPacienteHistorico = IIf(IsNull(oRecordset!idSunasaPacienteHistorico), 0, oRecordset!idSunasaPacienteHistorico)
       oTabla.HoraInicioAtencion = IIf(IsNull(oRecordset!HoraInicioAtencion), "", oRecordset!HoraInicioAtencion)
       oTabla.EpsPorcentaje = IIf(IsNull(oRecordset!EpsPorcentaje), 0, oRecordset!EpsPorcentaje)
End Sub

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla Atenciones
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionesSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oTabla.idAtencion): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       CargaTabla oTabla, oRecordset 'Carga al objeto oTabla los valores del oRecordset
'       oTabla.HoraIngreso = IIf(IsNull(oRecordset!HoraIngreso), "", oRecordset!HoraIngreso)
'       oTabla.FechaIngreso = IIf(IsNull(oRecordset!FechaIngreso), 0, oRecordset!FechaIngreso)
'       oTabla.IdTipoServicio = IIf(IsNull(oRecordset!IdTipoServicio), 0, oRecordset!IdTipoServicio)
'       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
'       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
'       oTabla.IdTipoCondicionALEstab = IIf(IsNull(oRecordset!IdTipoCondicionALEstab), 0, oRecordset!IdTipoCondicionALEstab)
'       oTabla.FechaEgresoAdministrativo = IIf(IsNull(oRecordset!FechaEgresoAdministrativo), 0, oRecordset!FechaEgresoAdministrativo)
'       oTabla.IdCamaEgreso = IIf(IsNull(oRecordset!IdCamaEgreso), 0, oRecordset!IdCamaEgreso)
'       oTabla.IdCamaIngreso = IIf(IsNull(oRecordset!IdCamaIngreso), 0, oRecordset!IdCamaIngreso)
'       oTabla.IdServicioEgreso = IIf(IsNull(oRecordset!IdServicioEgreso), 0, oRecordset!IdServicioEgreso)
'       oTabla.IdTipoAlta = IIf(IsNull(oRecordset!IdTipoAlta), 0, oRecordset!IdTipoAlta)
'       oTabla.IdCondicionAlta = IIf(IsNull(oRecordset!IdCondicionAlta), 0, oRecordset!IdCondicionAlta)
'       oTabla.IdTipoEdad = IIf(IsNull(oRecordset!IdTipoEdad), 0, oRecordset!IdTipoEdad)
'       oTabla.IdOrigenAtencion = IIf(IsNull(oRecordset!IdOrigenAtencion), 0, oRecordset!IdOrigenAtencion)
'       oTabla.IdDestinoAtencion = IIf(IsNull(oRecordset!IdDestinoAtencion), 0, oRecordset!IdDestinoAtencion)
'       oTabla.HoraEgresoAdministrativo = IIf(IsNull(oRecordset!HoraEgresoAdministrativo), "", oRecordset!HoraEgresoAdministrativo)
'       oTabla.IdTipoCondicionAlServicio = IIf(IsNull(oRecordset!IdTipoCondicionAlServicio), 0, oRecordset!IdTipoCondicionAlServicio)
'       oTabla.HoraEgreso = IIf(IsNull(oRecordset!HoraEgreso), "", oRecordset!HoraEgreso)
'       oTabla.FechaEgreso = IIf(IsNull(oRecordset!FechaEgreso), 0, oRecordset!FechaEgreso)
'       oTabla.IdMedicoEgreso = IIf(IsNull(oRecordset!IdMedicoEgreso), 0, oRecordset!IdMedicoEgreso)
'       oTabla.Edad = IIf(IsNull(oRecordset!Edad), 0, oRecordset!Edad)
'       oTabla.IdEspecialidadMedico = IIf(IsNull(oRecordset!IdEspecialidadMedico), 0, oRecordset!IdEspecialidadMedico)
'       oTabla.IdMedicoIngreso = IIf(IsNull(oRecordset!IdMedicoIngreso), 0, oRecordset!IdMedicoIngreso)
'       oTabla.IdServicioIngreso = IIf(IsNull(oRecordset!IdServicioIngreso), 0, oRecordset!IdServicioIngreso)
'       oTabla.IdTipoGravedad = IIf(IsNull(oRecordset!IdTipoGravedad), 0, oRecordset!IdTipoGravedad)
'       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
'       oTabla.idFormaPago = IIf(IsNull(oRecordset!idFormaPago), 0, oRecordset!idFormaPago)
'       oTabla.idFuenteFinanciamiento = IIf(IsNull(oRecordset!idFuenteFinanciamiento), 0, oRecordset!idFuenteFinanciamiento)
'       oTabla.idEstadoAtencion = IIf(IsNull(oRecordset!idEstadoAtencion), 0, oRecordset!idEstadoAtencion)
'       oTabla.EsPacienteExterno = IIf(IsNull(oRecordset!EsPacienteExterno), 0, oRecordset!EsPacienteExterno)
'       oTabla.idSunasaPacienteHistorico = IIf(IsNull(oRecordset!idSunasaPacienteHistorico), 0, oRecordset!idSunasaPacienteHistorico)
'       oTabla.HoraInicioAtencion = IIf(IsNull(oRecordset!HoraInicioAtencion), "", oRecordset!HoraInicioAtencion)
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "SeleccionarPorId"

Exit Function
End Function


Function SeleccionarPorIdCuentaDeAtencion(ByVal oTabla As DOAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorIdCuentaDeAtencion = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "atencionesSeleccionarXIdCuentaAtencion"
     Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oTabla.IdCuentaAtencion): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       CargaTabla oTabla, oRecordset
       SeleccionarPorIdCuentaDeAtencion = True
   Else
       SeleccionarPorIdCuentaDeAtencion = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "SeleccionarPorId"

Exit Function
End Function




Function AtencionesPacientesCitasDatosadicionalesSeleccionarPorId(oTabla As DOAtencion, _
                                          oDOAtencionDatosAdicionales As DOAtencionDatosAdicionales, _
                                          DoPaciente As DoPaciente, _
                                          mo_CuentasAtencion As DOCuentaAtencion, _
                                          lbConsideraTablaCita As Boolean, _
                                          Optional ByRef DoCita As DoCita) As Boolean
On Error GoTo ManejadorDeError

Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     If lbConsideraTablaCita = True Then
        .CommandText = "AtencionesSeleccionarPorIdAtencionEnAtencionesCitasAtencionesdatosadicionalesPacientes "
        Set oParameter = .CreateParameter("@IdCita", adInteger, adParamInput, 0, DoCita.IdCita): .Parameters.Append oParameter
     Else
        .CommandText = "AtencionesSeleccionarPorIdAtencionEnAtencionesAtencionesdatosadicionalesPacientes"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oTabla.idAtencion): .Parameters.Append oParameter
     End If
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       AtencionesPacientesCitasDatosadicionalesSeleccionarPorId = True
       'Atenciones
       oTabla.HoraIngreso = IIf(IsNull(oRecordset!HoraIngreso), "", oRecordset!HoraIngreso)
       oTabla.FechaIngreso = IIf(IsNull(oRecordset!FechaIngreso), 0, oRecordset!FechaIngreso)
       oTabla.IdTipoServicio = IIf(IsNull(oRecordset!IdTipoServicio), 0, oRecordset!IdTipoServicio)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       If lbConsideraTablaCita = True Then
          oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       End If
       oTabla.IdTipoCondicionALEstab = IIf(IsNull(oRecordset!IdTipoCondicionALEstab), 0, oRecordset!IdTipoCondicionALEstab)
       oTabla.FechaEgresoAdministrativo = IIf(IsNull(oRecordset!FechaEgresoAdministrativo), 0, oRecordset!FechaEgresoAdministrativo)
       oTabla.IdCamaEgreso = IIf(IsNull(oRecordset!IdCamaEgreso), 0, oRecordset!IdCamaEgreso)
       oTabla.IdCamaIngreso = IIf(IsNull(oRecordset!IdCamaIngreso), 0, oRecordset!IdCamaIngreso)
       oTabla.IdServicioEgreso = IIf(IsNull(oRecordset!IdServicioEgreso), 0, oRecordset!IdServicioEgreso)
       oTabla.IdTipoAlta = IIf(IsNull(oRecordset!IdTipoAlta), 0, oRecordset!IdTipoAlta)
       oTabla.IdCondicionAlta = IIf(IsNull(oRecordset!IdCondicionAlta), 0, oRecordset!IdCondicionAlta)
       oTabla.IdTipoEdad = IIf(IsNull(oRecordset!IdTipoEdad), 0, oRecordset!IdTipoEdad)
       oTabla.IdOrigenAtencion = IIf(IsNull(oRecordset!IdOrigenAtencion), 0, oRecordset!IdOrigenAtencion)
       oTabla.IdDestinoAtencion = IIf(IsNull(oRecordset!IdDestinoAtencion), 0, oRecordset!IdDestinoAtencion)
       oTabla.HoraEgresoAdministrativo = IIf(IsNull(oRecordset!HoraEgresoAdministrativo), "", oRecordset!HoraEgresoAdministrativo)
       oTabla.IdTipoCondicionAlServicio = IIf(IsNull(oRecordset!IdTipoCondicionAlServicio), 0, oRecordset!IdTipoCondicionAlServicio)
       oTabla.HoraEgreso = IIf(IsNull(oRecordset!HoraEgreso), "", oRecordset!HoraEgreso)
       oTabla.FechaEgreso = IIf(IsNull(oRecordset!FechaEgreso), 0, oRecordset!FechaEgreso)
       oTabla.IdMedicoEgreso = IIf(IsNull(oRecordset!IdMedicoEgreso), 0, oRecordset!IdMedicoEgreso)
       oTabla.Edad = IIf(IsNull(oRecordset!Edad), 0, oRecordset!Edad)
       oTabla.IdEspecialidadMedico = IIf(IsNull(oRecordset!IdEspecialidadMedico), 0, oRecordset!IdEspecialidadMedico)
       oTabla.IdMedicoIngreso = IIf(IsNull(oRecordset!IdMedicoIngreso), 0, oRecordset!IdMedicoIngreso)
       oTabla.IdServicioIngreso = IIf(IsNull(oRecordset!IdServicioIngreso), 0, oRecordset!IdServicioIngreso)
       oTabla.IdTipoGravedad = IIf(IsNull(oRecordset!IdTipoGravedad), 0, oRecordset!IdTipoGravedad)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.idFormaPago = IIf(IsNull(oRecordset!idFormaPago), 0, oRecordset!idFormaPago)
       oTabla.idFuenteFinanciamiento = IIf(IsNull(oRecordset!idFuenteFinanciamiento), 0, oRecordset!idFuenteFinanciamiento)
       oTabla.idEstadoAtencion = IIf(IsNull(oRecordset!idEstadoAtencion), 0, oRecordset!idEstadoAtencion)
       oTabla.EsPacienteExterno = IIf(IsNull(oRecordset!EsPacienteExterno), 0, oRecordset!EsPacienteExterno)
       oTabla.idSunasaPacienteHistorico = IIf(IsNull(oRecordset!idSunasaPacienteHistorico), 0, oRecordset!idSunasaPacienteHistorico)
       'Citas
       If lbConsideraTablaCita = True Then
            DoCita.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
            DoCita.IdProgramacion = IIf(IsNull(oRecordset!IdProgramacion), 0, oRecordset!IdProgramacion)
            DoCita.IdServicio = IIf(IsNull(oRecordset!IdServicio), 0, oRecordset!IdServicio)
            DoCita.HoraFin = IIf(IsNull(oRecordset!HoraFin), "", oRecordset!HoraFin)
            DoCita.HoraInicio = IIf(IsNull(oRecordset!HoraInicio), "", oRecordset!HoraInicio)
            DoCita.IdCita = IIf(IsNull(oRecordset!IdCita), 0, oRecordset!IdCita)
            DoCita.Fecha = IIf(IsNull(oRecordset!Fecha), 0, oRecordset!Fecha)
            DoCita.IdEstadoCita = IIf(IsNull(oRecordset!IdEstadoCita), 0, oRecordset!IdEstadoCita)
            DoCita.IdMedico = IIf(IsNull(oRecordset!IdMedico), 0, oRecordset!IdMedico)
            DoCita.IdEspecialidad = IIf(IsNull(oRecordset!IdEspecialidad), 0, oRecordset!IdEspecialidad)
            DoCita.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
            DoCita.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
            DoCita.FechaSolicitud = IIf(IsNull(oRecordset!FechaSolicitud), 0, oRecordset!FechaSolicitud)
            DoCita.HoraSolicitud = IIf(IsNull(oRecordset!HoraSolicitud), "", oRecordset!HoraSolicitud)
            DoCita.EsCitaAdicional = IIf(IsNull(oRecordset!EsCitaAdicional), 0, oRecordset!EsCitaAdicional)
       End If
       'Datos Adicionales
       oDOAtencionDatosAdicionales.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       oDOAtencionDatosAdicionales.DireccionDomicilio = IIf(IsNull(oRecordset!DireccionDomicilio), "", oRecordset!DireccionDomicilio)
       oDOAtencionDatosAdicionales.NombreAcompaniante = IIf(IsNull(oRecordset!NombreAcompaniante), "", oRecordset!NombreAcompaniante)
       oDOAtencionDatosAdicionales.Observacion = IIf(IsNull(oRecordset!obs_AtencionesDatosAdicionales), "", oRecordset!obs_AtencionesDatosAdicionales)
       oDOAtencionDatosAdicionales.ProximaCita = IIf(IsNull(oRecordset!ProximaCita), 0, oRecordset!ProximaCita)
       oDOAtencionDatosAdicionales.NumeroDeHijos = IIf(IsNull(oRecordset!NumeroDeHijos), 0, oRecordset!NumeroDeHijos)
       oDOAtencionDatosAdicionales.idSiasis = IIf(IsNull(oRecordset!idSiasis), 0, oRecordset!idSiasis)
       oDOAtencionDatosAdicionales.FuaCodigoPrestacion = IIf(IsNull(oRecordset!FuaCodigoPrestacion), "", oRecordset!FuaCodigoPrestacion)
       oDOAtencionDatosAdicionales.SisCodigo = IIf(IsNull(oRecordset!SisCodigo), "", oRecordset!SisCodigo)
       oDOAtencionDatosAdicionales.IdTipoReferenciaDestino = IIf(IsNull(oRecordset!IdTipoReferenciaDestino), 0, oRecordset!IdTipoReferenciaDestino)
       oDOAtencionDatosAdicionales.IdTipoReferenciaOrigen = IIf(IsNull(oRecordset!IdTipoReferenciaOrigen), 0, oRecordset!IdTipoReferenciaOrigen)
       oDOAtencionDatosAdicionales.IdEstablecimientoDestino = IIf(IsNull(oRecordset!IdEstablecimientoDestino), 0, oRecordset!IdEstablecimientoDestino)
       oDOAtencionDatosAdicionales.IdEstablecimientoOrigen = IIf(IsNull(oRecordset!IdEstablecimientoOrigen), 0, oRecordset!IdEstablecimientoOrigen)
       oDOAtencionDatosAdicionales.IdEstablecimientoNoMinsaDestino = IIf(IsNull(oRecordset!IdEstablecimientoNoMinsaDestino), 0, oRecordset!IdEstablecimientoNoMinsaDestino)
       oDOAtencionDatosAdicionales.IdEstablecimientoNoMinsaOrigen = IIf(IsNull(oRecordset!IdEstablecimientoNoMinsaOrigen), 0, oRecordset!IdEstablecimientoNoMinsaOrigen)
       oDOAtencionDatosAdicionales.HuboInfeccionIntraHospitalaria = IIf(IsNull(oRecordset!HuboInfeccionIntraHospitalaria), 0, oRecordset!HuboInfeccionIntraHospitalaria)
       oDOAtencionDatosAdicionales.TieneNecropsia = IIf(IsNull(oRecordset!TieneNecropsia), 0, oRecordset!TieneNecropsia)
       oDOAtencionDatosAdicionales.IdMedicoRespNacimiento = IIf(IsNull(oRecordset!IdMedicoRespNacimiento), 0, oRecordset!IdMedicoRespNacimiento)
       oDOAtencionDatosAdicionales.RecienNacido = IIf(IsNull(oRecordset!RecienNacido), 0, oRecordset!RecienNacido)
       oDOAtencionDatosAdicionales.NroReferenciaOrigen = IIf(IsNull(oRecordset!NroReferenciaOrigen), "", oRecordset!NroReferenciaOrigen)
       oDOAtencionDatosAdicionales.NroReferenciaDestino = IIf(IsNull(oRecordset!NroReferenciaDestino), "", oRecordset!NroReferenciaDestino)
       oDOAtencionDatosAdicionales.referenciaOdniMedico = IIf(IsNull(oRecordset!referenciaOdniMedico), "", oRecordset!referenciaOdniMedico)
       oDOAtencionDatosAdicionales.acompanianteDNI = IIf(IsNull(oRecordset!acompanianteDNI), "", oRecordset!acompanianteDNI)
       oDOAtencionDatosAdicionales.emergenciaCorrelativo = IIf(IsNull(oRecordset!emergenciaCorrelativo), "", oRecordset!emergenciaCorrelativo)
       oDOAtencionDatosAdicionales.SeImprimioFicha = IIf(IsNull(oRecordset!SeImprimioFicha), 0, oRecordset!SeImprimioFicha)
       oDOAtencionDatosAdicionales.sisAfiliacion = IIf(IsNull(oRecordset!sisAfiliacion), "", oRecordset!sisAfiliacion)
       oDOAtencionDatosAdicionales.referenciaOservicio = IIf(IsNull(oRecordset!referenciaOservicio), "", oRecordset!referenciaOservicio)
       oDOAtencionDatosAdicionales.referenciaOidDiagnostico = IIf(IsNull(oRecordset!referenciaOidDiagnostico), 0, oRecordset!referenciaOidDiagnostico)
       oDOAtencionDatosAdicionales.referenciaDservicio = IIf(IsNull(oRecordset!referenciaDservicio), "", oRecordset!referenciaDservicio)
       oDOAtencionDatosAdicionales.referenciaDfextension = IIf(IsNull(oRecordset!referenciaDfextension), 0, oRecordset!referenciaDfextension)
       oDOAtencionDatosAdicionales.referenciaDftramite = IIf(IsNull(oRecordset!referenciaDftramite), 0, oRecordset!referenciaDftramite)
       oDOAtencionDatosAdicionales.idAtencionEmeg_CE = IIf(IsNull(oRecordset!idAtencionEmeg_CE), 0, oRecordset!idAtencionEmeg_CE)
       oDOAtencionDatosAdicionales.ReferenciaMedicoOColeg = IIf(IsNull(oRecordset!ReferenciaMedicoOColeg), "", oRecordset!ReferenciaMedicoOColeg)    'franklin 2017
       oDOAtencionDatosAdicionales.ReferenciaMedicoOIdcolegio = IIf(IsNull(oRecordset!ReferenciaMedicoOIdcolegio), "", oRecordset!ReferenciaMedicoOIdcolegio)    'franklin 2017
       oDOAtencionDatosAdicionales.IdServicioDestino = IIf(IsNull(oRecordset!IdServicioDestino), 0, oRecordset!IdServicioDestino)

       
       'Pacientes
       DoPaciente.IdPaisNacimiento = IIf(IsNull(oRecordset!IdPaisNacimiento), 0, oRecordset!IdPaisNacimiento)
       DoPaciente.ApellidoMaterno = IIf(IsNull(oRecordset!ApellidoMaterno), "", oRecordset!ApellidoMaterno)
       DoPaciente.DireccionDomicilio = IIf(IsNull(oRecordset!DireccionDomicilio), "", oRecordset!DireccionDomicilio)
       DoPaciente.Observacion = IIf(IsNull(oRecordset!obs_pacientes), "", oRecordset!obs_pacientes)
       DoPaciente.IdTipoNumeracion = IIf(IsNull(oRecordset!IdTipoNumeracion), 0, oRecordset!IdTipoNumeracion)
       DoPaciente.IdPaisProcedencia = IIf(IsNull(oRecordset!IdPaisProcedencia), 0, oRecordset!IdPaisProcedencia)
       DoPaciente.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       DoPaciente.ApellidoPaterno = IIf(IsNull(oRecordset!ApellidoPaterno), "", oRecordset!ApellidoPaterno)
       DoPaciente.PrimerNombre = IIf(IsNull(oRecordset!PrimerNombre), "", oRecordset!PrimerNombre)
       DoPaciente.SegundoNombre = IIf(IsNull(oRecordset!SegundoNombre), "", oRecordset!SegundoNombre)
       DoPaciente.TercerNombre = IIf(IsNull(oRecordset!TercerNombre), "", oRecordset!TercerNombre)
       DoPaciente.FechaNacimiento = IIf(IsNull(oRecordset!FechaNacimiento), 0, oRecordset!FechaNacimiento)
       DoPaciente.NroDocumento = IIf(IsNull(oRecordset!NroDocumento), "", oRecordset!NroDocumento)
       DoPaciente.Telefono = IIf(IsNull(oRecordset!Telefono), "", oRecordset!Telefono)
       DoPaciente.Autogenerado = IIf(IsNull(oRecordset!Autogenerado), "", oRecordset!Autogenerado)
       DoPaciente.idTipoSexo = IIf(IsNull(oRecordset!idTipoSexo), 0, oRecordset!idTipoSexo)
       DoPaciente.IdProcedencia = IIf(IsNull(oRecordset!IdProcedencia), 0, oRecordset!IdProcedencia)
       DoPaciente.IdGradoInstruccion = IIf(IsNull(oRecordset!IdGradoInstruccion), 0, oRecordset!IdGradoInstruccion)
       DoPaciente.IdEstadoCivil = IIf(IsNull(oRecordset!IdEstadoCivil), 0, oRecordset!IdEstadoCivil)
       DoPaciente.IdDocIdentidad = IIf(IsNull(oRecordset!IdDocIdentidad), 0, oRecordset!IdDocIdentidad)
       DoPaciente.IdTipoOcupacion = IIf(IsNull(oRecordset!IdTipoOcupacion), 0, oRecordset!IdTipoOcupacion)
       DoPaciente.IdCentroPobladoDomicilio = IIf(IsNull(oRecordset!IdCentroPobladoDomicilio), 0, oRecordset!IdCentroPobladoDomicilio)
       DoPaciente.NombrePadre = IIf(IsNull(oRecordset!NombrePadre), "", oRecordset!NombrePadre)
       DoPaciente.NombreMadre = IIf(IsNull(oRecordset!NombreMadre), "", oRecordset!NombreMadre)
       DoPaciente.IdPaisDomicilio = IIf(IsNull(oRecordset!IdPaisDomicilio), 0, oRecordset!IdPaisDomicilio)
       DoPaciente.NroHistoriaClinica = IIf(IsNull(oRecordset!NroHistoriaClinica), 0, oRecordset!NroHistoriaClinica)
       DoPaciente.IdCentroPobladoNacimiento = IIf(IsNull(oRecordset!IdCentroPobladoNacimiento), 0, oRecordset!IdCentroPobladoNacimiento)
       DoPaciente.IdCentroPobladoProcedencia = IIf(IsNull(oRecordset!IdCentroPobladoProcedencia), 0, oRecordset!IdCentroPobladoProcedencia)
       DoPaciente.IdDistritoProcedencia = IIf(IsNull(oRecordset!IdDistritoProcedencia), 0, oRecordset!IdDistritoProcedencia)
       DoPaciente.IdDistritoDomicilio = IIf(IsNull(oRecordset!IdDistritoDomicilio), 0, oRecordset!IdDistritoDomicilio)
       DoPaciente.IdDistritoNacimiento = IIf(IsNull(oRecordset!IdDistritoNacimiento), 0, oRecordset!IdDistritoNacimiento)
       DoPaciente.FichaFamiliar = IIf(IsNull(oRecordset!FichaFamiliar), "", oRecordset!FichaFamiliar)
       DoPaciente.IdEtnia = IIf(IsNull(oRecordset!IdEtnia), "", oRecordset!IdEtnia)
       DoPaciente.GrupoSanguineo = IIf(IsNull(oRecordset!GrupoSanguineo), "", oRecordset!GrupoSanguineo)
       DoPaciente.FactorRh = IIf(IsNull(oRecordset!FactorRh), "", oRecordset!FactorRh)
       DoPaciente.UsoWebReniec = IIf(IsNull(oRecordset!UsoWebReniec), 0, oRecordset!UsoWebReniec)
       
'<(Inicio) Añadido Por: WABG el: 22/10/2020-10:55:47 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
       DoPaciente.validacionReniec = IIf(IsNull(oRecordset!validacionReniec), 0, oRecordset!validacionReniec)
'</(Fin) Añadido Por: WABG el: 22/10/2020-10:55:47 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
       
       DoPaciente.idIdioma = IIf(IsNull(oRecordset!idIdioma), 0, oRecordset!idIdioma)
       DoPaciente.Email = IIf(IsNull(oRecordset!Email), "", oRecordset!Email)
       DoPaciente.madreDocumento = IIf(IsNull(oRecordset!madreDocumento), "", oRecordset!madreDocumento)
       DoPaciente.madreApellidoPaterno = IIf(IsNull(oRecordset!madreApellidoPaterno), "", oRecordset!madreApellidoPaterno)
       DoPaciente.madreApellidoMaterno = IIf(IsNull(oRecordset!madreApellidoMaterno), "", oRecordset!madreApellidoMaterno)
       DoPaciente.madrePrimerNombre = IIf(IsNull(oRecordset!madrePrimerNombre), "", oRecordset!madrePrimerNombre)
       DoPaciente.madreSegundoNombre = IIf(IsNull(oRecordset!madreSegundoNombre), "", oRecordset!madreSegundoNombre)
       DoPaciente.NroOrdenHijo = IIf(IsNull(oRecordset!NroOrdenHijo), 0, oRecordset!NroOrdenHijo)
       DoPaciente.madreTipoDocumento = IIf(IsNull(oRecordset!madreTipoDocumento), 0, oRecordset!madreTipoDocumento)
       DoPaciente.sector = IIf(IsNull(oRecordset!sector), "", oRecordset!sector)
       DoPaciente.sectorista = IIf(IsNull(oRecordset!sector), 0, oRecordset!sectorista)
       
       
       'CuentasAtencion
       mo_CuentasAtencion.TotalPorPagar = IIf(IsNull(oRecordset!TotalPorPagar), 0, oRecordset!TotalPorPagar)
       mo_CuentasAtencion.IdEstado = IIf(IsNull(oRecordset!IdEstado), 0, oRecordset!IdEstado)
       mo_CuentasAtencion.TotalPagado = IIf(IsNull(oRecordset!TotalPagado), 0, oRecordset!TotalPagado)
       mo_CuentasAtencion.TotalAsegurado = IIf(IsNull(oRecordset!TotalAsegurado), 0, oRecordset!TotalAsegurado)
       mo_CuentasAtencion.TotalExonerado = IIf(IsNull(oRecordset!TotalExonerado), 0, oRecordset!TotalExonerado)
       mo_CuentasAtencion.HoraCierre = IIf(IsNull(oRecordset!HoraCierre), "", oRecordset!HoraCierre)
       mo_CuentasAtencion.FechaCierre = IIf(IsNull(oRecordset!FechaCierre), 0, oRecordset!FechaCierre)
       mo_CuentasAtencion.HoraApertura = IIf(IsNull(oRecordset!HoraApertura), "", oRecordset!HoraApertura)
       mo_CuentasAtencion.FechaApertura = IIf(IsNull(oRecordset!FechaApertura), 0, oRecordset!FechaApertura)
       mo_CuentasAtencion.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       mo_CuentasAtencion.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       mo_CuentasAtencion.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
   Else
       AtencionesPacientesCitasDatosadicionalesSeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "SeleccionarPorId"

Exit Function
End Function










Function ActualizarCitasConIdAtencion(IdCita As Long, idAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CitasActualizarConIdAtencion"
        Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCita", adInteger, adParamInput, 0, IdCita): .Parameters.Append oParameter
        .Execute
   End With
   Set oRecordset = Nothing
   Set oCommand = Nothing
   ActualizarCitasConIdAtencion = True
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ReporteParaHistoriaClinica(ByVal lIdAtencion As Long) As ADODB.Recordset 'scrafet reporte 2
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim oParameter As ADODB.Parameter
     Set ReporteParaHistoriaClinica = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandType = adCmdStoredProc
       .CommandText = "AtencionesReporteParaHistoriaClinica"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
     Set ReporteParaHistoriaClinica = oRecordset
     ms_MensajeError = ""
     Set oCommand = Nothing
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "ReporteParaHistoriaClinica"
Exit Function
End Function

'<(Inicio) Añadido Por: WABG el: 13/05/2021-09:34:11 a.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-74>
'--------------------------------------------------------------------------------
' Project    :       SIGHDatos
' Procedure  :       ObtenerDatosAdicionalesDeAtencionPorIdAtencion
' Description:       [obtiene los datos desde la tabla sigh.dbo.AtencionesDatosAdicionales]
' Created by :       WABG
' Machine    :       SISGALENPLUS-PC
' Date-Time  :       13/05/2021-09:35:32 a.m.
' Parameters :       lIdAtencion (Long)
'--------------------------------------------------------------------------------
Function ObtenerDatosAdicionalesDeAtencionPorIdAtencion(ByVal lIdAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim oParameter As ADODB.Parameter
     Set ObtenerDatosAdicionalesDeAtencionPorIdAtencion = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandType = adCmdStoredProc
       .CommandText = "usp_ObtenerDatosAdicionalesDeAtencionPorIdAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
     Set ObtenerDatosAdicionalesDeAtencionPorIdAtencion = oRecordset
     ms_MensajeError = ""
     Set oCommand = Nothing
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "ObtenerDatosAdicionalesDeAtencionPorIdAtencion"
Exit Function
End Function
'</(Fin) Añadido Por: WABG el: 13/05/2021-09:34:11 a.m. en el Equipo: SISGALENPLUS-PC<CAMBIO-74>

Function SeleccionarUltimaAtencionHospitalizacion(ByVal lIdPaciente As Long, lTipoServicio As Long) As DOAtencion
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oDOAtencion As New DOAtencion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "atencionesMaxAtencionesXPacienteYTipoServicio"
        Set oParameter = .CreateParameter("@lIdPaciente", adInteger, adParamInput, 0, lIdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lTipoServicio", adInteger, adParamInput, 0, lTipoServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   oDOAtencion.idAtencion = IIf(IsNull(oRecordset!IdMaxAtencion), 0, oRecordset!IdMaxAtencion)
   If SeleccionarPorId(oDOAtencion) Then
        Set SeleccionarUltimaAtencionHospitalizacion = oDOAtencion
   Else
        Set SeleccionarUltimaAtencionHospitalizacion = Nothing
   End If
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
Exit Function
End Function
'------------------------------------------

Function SeleccionarUltimaAtencion(ByVal lIdPaciente As Long, IdCuentaAtencion As Long) As DOAtencion
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oDOAtencion As New DOAtencion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesMaxAtencionesXPacienteYCuenta"
        Set oParameter = .CreateParameter("@lIdPaciente", adInteger, adParamInput, 0, lIdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   oDOAtencion.idAtencion = IIf(IsNull(oRecordset!IdMaxAtencion), 0, oRecordset!IdMaxAtencion)
   If SeleccionarPorId(oDOAtencion) Then 'Selecciona de la tabla Atenciones por idAtencion y carga los valores al objeto oDOAtencion
        Set SeleccionarUltimaAtencion = oDOAtencion
   Else
        Set SeleccionarUltimaAtencion = Nothing
   End If
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
Exit Function
End Function

'--------------------------------------------
Function FiltrarConsultaExterna(oDOPaciente As DoPaciente, oDOAtencion As DOAtencion, lcFechaIngreso As String, IdMedicoIngreso As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim sWhere As String
       sWhere = " Atenciones.idEstadoAtencion<>0 and "
       If oDOPaciente.ApellidoPaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoPaterno like '" + oDOPaciente.ApellidoPaterno + "%' and "
       End If
       If oDOPaciente.ApellidoMaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoMaterno like '" + oDOPaciente.ApellidoMaterno + "%' and "
       End If
       If oDOPaciente.PrimerNombre <> "" Then
            sWhere = sWhere + " PAcientes.PrimerNombre like '%" + oDOPaciente.PrimerNombre + "%' and "
       End If
       If oDOPaciente.SegundoNombre <> "" Then
            sWhere = sWhere + " PAcientes.SegundoNombre like '%" + oDOPaciente.SegundoNombre + "%' and "
       End If
       If oDOPaciente.NroHistoriaClinica <> 0 Then
            sWhere = sWhere + " Pacientes.NroHistoriaClinica = " & oDOPaciente.NroHistoriaClinica & " and "
       End If
        If oDOAtencion.idAtencion <> 0 Then
            sWhere = sWhere + " Atenciones.IdAtencion = " & oDOAtencion.idAtencion & " and "
        End If
        If UCase(lcFechaIngreso) <> "TODAS" Then
            sWhere = sWhere + " Atenciones.fechaIngreso Between (CONVERT(DATETIME,'" & lcFechaIngreso & " 00:00:00',103)) and (CONVERT(DATETIME,'" & lcFechaIngreso & " 23:59:59',103)) and "
        End If
        If IdMedicoIngreso > 0 Then
           sWhere = sWhere + " atenciones.idMedicoIngreso=" & IdMedicoIngreso & " and "
        End If

        sSql = " Where Atenciones.IdTipoServicio = 1 "
        If (sWhere <> "") Then
            sSql = sSql + " and " + Left(sWhere, Len(sWhere) - 4)
        End If

       sSql = sSql + " order by Atenciones.FechaIngreso asc, Atenciones.HoraIngreso asc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FiltrarConsultaExterna"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1500, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FiltrarConsultaExterna = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function FiltrarConsultorioEmergencia(oDOPaciente As DoPaciente, oDOAtencion As DOAtencion) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim sWhere As String
        sWhere = ""
        sWhere = " Atenciones.idEstadoAtencion<>0 and "
       If oDOPaciente.ApellidoPaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoPaterno like '" + oDOPaciente.ApellidoPaterno + "%' and "
       End If
       If oDOPaciente.ApellidoMaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoMaterno like '" + oDOPaciente.ApellidoMaterno + "%' and "
       End If
       If oDOPaciente.PrimerNombre <> "" Then
            sWhere = sWhere + " PAcientes.PrimerNombre like '%" + oDOPaciente.PrimerNombre + "%' and "
       End If
       If oDOPaciente.NroHistoriaClinica <> 0 Then
            sWhere = sWhere + " Pacientes.NroHistoriaClinica = " & oDOPaciente.NroHistoriaClinica & " and "
       End If
        If oDOAtencion.idAtencion <> 0 Then
            sWhere = sWhere + " Atenciones.IdAtencion = " & oDOAtencion.idAtencion & " and "
        End If

        sSql = " Where Atenciones.IdTipoServicio = 2 "
        If (sWhere <> "") Then
            sSql = sSql + " and " + Left(sWhere, Len(sWhere) - 4)
        End If

       sSql = sSql + " order by Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre, Atenciones.FechaIngreso"


    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesEmergenciaSegunFiltro"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1500, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FiltrarConsultorioEmergencia = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function FiltrarObservacionEmergencia(oDOPaciente As DoPaciente, oDOAtencion As DOAtencion) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim sWhere As String

       sWhere = " Atenciones.idEstadoAtencion<>0 and "
       If oDOPaciente.ApellidoPaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoPaterno like '" + oDOPaciente.ApellidoPaterno + "%' and "
       End If
       If oDOPaciente.ApellidoMaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoMaterno like '" + oDOPaciente.ApellidoMaterno + "%' and "
       End If
       If oDOPaciente.PrimerNombre <> "" Then
            sWhere = sWhere + " PAcientes.PrimerNombre like '%" + oDOPaciente.PrimerNombre + "%' and "
       End If
       If oDOPaciente.NroHistoriaClinica <> 0 Then
            sWhere = sWhere + " Pacientes.NroHistoriaClinica = " & oDOPaciente.NroHistoriaClinica & " and "
       End If
        If oDOAtencion.idAtencion <> 0 Then
            sWhere = sWhere + " Atenciones.IdAtencion = " & oDOAtencion.idAtencion & " and "
        End If

        sSql = " Where Atenciones.IdTipoServicio = 4 "
        If (sWhere <> "") Then
            sSql = sSql + " and " + Left(sWhere, Len(sWhere) - 4)
        End If

       sSql = sSql + " order by Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre, Atenciones.FechaIngreso "

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesFiltrarObservacionEmergencia"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1500, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FiltrarObservacionEmergencia = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function FiltrarEmergencia(oDOPaciente As DoPaciente, oDOAtencion As DOAtencion, lcFechaIngreso As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim sWhere As String
       sWhere = " Atenciones.idEstadoAtencion<>0 and "
       If oDOPaciente.ApellidoPaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoPaterno like '" + oDOPaciente.ApellidoPaterno + "%' and "
       End If
       If oDOPaciente.ApellidoMaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoMaterno like '" + oDOPaciente.ApellidoMaterno + "%' and "
       End If
       If oDOPaciente.PrimerNombre <> "" Then
            sWhere = sWhere + " PAcientes.PrimerNombre like '%" + oDOPaciente.PrimerNombre + "%' and "
       End If
       If oDOPaciente.NroHistoriaClinica <> 0 Then
            sWhere = sWhere + " Pacientes.NroHistoriaClinica = " & oDOPaciente.NroHistoriaClinica & " and "
       End If
        If oDOAtencion.idAtencion <> 0 Then
            sWhere = sWhere + " Atenciones.IdAtencion = " & oDOAtencion.idAtencion & " and "
        End If
        If UCase(lcFechaIngreso) <> "TODAS" Then
            sWhere = sWhere + " Atenciones.fechaIngreso Between (CONVERT(DATETIME,'" & lcFechaIngreso & " 00:00:00',103)) and (CONVERT(DATETIME,'" & lcFechaIngreso & " 23:59:59',103)) and "
        End If


        sSql = " Where Atenciones.IdTipoServicio in (2,4) "
        If (sWhere <> "") Then
            sSql = sSql + " and " + Left(sWhere, Len(sWhere) - 4)
        End If

       sSql = sSql + " order by Atenciones.FechaIngreso desc,  Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"


    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesFiltrarEmergencia"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1500, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FiltrarEmergencia = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function FiltrarHospitalizacion(oDOPaciente As DoPaciente, oDOAtencion As DOAtencion, lcFechaIngreso As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim sWhere As String
       sWhere = " Atenciones.idEstadoAtencion<>0 and "
       If oDOPaciente.ApellidoPaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoPaterno like '" + oDOPaciente.ApellidoPaterno + "%' and "
       End If
       If oDOPaciente.ApellidoMaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoMaterno like '" + oDOPaciente.ApellidoMaterno + "%' and "
       End If
       If oDOPaciente.PrimerNombre <> "" Then
            sWhere = sWhere + " PAcientes.PrimerNombre like '%" + oDOPaciente.PrimerNombre + "%' and "
       End If
       If oDOPaciente.NroHistoriaClinica <> 0 Then
            sWhere = sWhere + " Pacientes.NroHistoriaClinica = " & oDOPaciente.NroHistoriaClinica & " and "
       End If
        If oDOAtencion.idAtencion <> 0 Then
            sWhere = sWhere + " Atenciones.IdAtencion = " & oDOAtencion.idAtencion & " and "
        End If
        If UCase(lcFechaIngreso) <> "TODAS" Then
            sWhere = sWhere + " Atenciones.fechaIngreso Between (CONVERT(DATETIME,'" & lcFechaIngreso & " 00:00:00',103)) and (CONVERT(DATETIME,'" & lcFechaIngreso & " 23:59:59',103)) and "
        End If

        sSql = " Where Atenciones.IdTipoServicio = 3 "
        If (sWhere <> "") Then
            sSql = sSql + " and " + Left(sWhere, Len(sWhere) - 4)
        End If

       sSql = sSql + " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"


    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesFiltrarHospitalizacion"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1500, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FiltrarHospitalizacion = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function FiltrarPacientesParaIngresarProcedimientos(oDOPaciente As DoPaciente, oDOCuentaAtencion As DOCuentaAtencion) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim sWhere As String
        sWhere = ""
       If oDOPaciente.NroHistoriaClinica <> 0 Then
            sWhere = sWhere + " Pacientes.NroHistoriaClinica = " & oDOPaciente.NroHistoriaClinica & " and "
       End If
        If oDOCuentaAtencion.IdCuentaAtencion <> 0 Then
            sWhere = sWhere + " CuentasAtencion.IdCuentaAtencion = " & oDOCuentaAtencion.IdCuentaAtencion & " and "
        End If

        If (sWhere <> "") Then
            sSql = " where " + Left(sWhere, Len(sWhere) - 4)
            sSql = sSql + " order by Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre, Atenciones.FechaIngreso desc"
        Else
            sSql = " order by Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre, Atenciones.FechaIngreso desc"
        End If

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesFiltrarPacientesParaIngresarProcedimientos"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1500, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FiltrarPacientesParaIngresarProcedimientos = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Sub Trace(sMensaje As String, NombreDeMetodo As String)
    MsgBox "Mensaje: " & sMensaje + Chr(13) + "Metodo:" & NombreDeMetodo + Chr(13) + "Clase de Datos: Atenciones", vbInformation, "Atenciones"
End Sub


Function AtencionesSeleccionarCEPorCuentaPorHistoriaPorApellidosPorServicio(lnHistoriaClinica As Long, lnIdCuentaAtencion As Long, lcApellidoPaterno As String, lcFechaIngreso As String, lnIdServicio As Long, lcDNI As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim sWhere As String
     sSql = ""
     If lnHistoriaClinica > 0 Then
        sSql = sSql & " WHERE dbo.Pacientes.NroHistoriaClinica=" & lnHistoriaClinica & _
                      "      and dbo.Atenciones.IdServicioIngreso=" & lnIdServicio & _
                      "      and dbo.Atenciones.idTipoServicio=1  and dbo.Atenciones.esPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso asc, Atenciones.HoraIngreso asc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     ElseIf lnIdCuentaAtencion > 0 Then
'        sSql = sSql & " WHERE dbo.Atenciones.IdCuentaAtencion=" & lnIdCuentaAtencion & _
'                      "      and dbo.Atenciones.IdServicioIngreso=" & lnIdServicio & _
'                      "      and dbo.Atenciones.idTipoServicio=1  and dbo.Atenciones.esPacienteExterno<>1" & _
'                      " order by Atenciones.FechaIngreso asc, Atenciones.HoraIngreso asc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
        sSql = sSql & " WHERE dbo.Atenciones.IdCuentaAtencion=" & lnIdCuentaAtencion & _
                      "      and dbo.Atenciones.idTipoServicio=1  and dbo.Atenciones.esPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso asc, Atenciones.HoraIngreso asc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     ElseIf lcApellidoPaterno <> "" Then
        sSql = sSql & " WHERE dbo.Pacientes.ApellidoPaterno like '" & lcApellidoPaterno & "%'" & _
                      "      and dbo.Atenciones.IdServicioIngreso=" & lnIdServicio & _
                      "      and dbo.Atenciones.idTipoServicio=1  and dbo.Atenciones.esPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso asc, Atenciones.HoraIngreso asc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     ElseIf lcDNI <> "" Then
        sSql = sSql & " WHERE dbo.Pacientes.idDocIdentidad=1 and dbo.Pacientes.nroDocumento='" & lcDNI & "'" & _
                      "      and dbo.Atenciones.IdServicioIngreso=" & lnIdServicio & _
                      "      and dbo.Atenciones.idTipoServicio=1  and dbo.Atenciones.esPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso asc, Atenciones.HoraIngreso asc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     ElseIf UCase(lcFechaIngreso) = "TODAS" Then
        sSql = sSql & " WHERE dbo.Atenciones.IdServicioIngreso=" & lnIdServicio & _
                      "      and dbo.Atenciones.idTipoServicio=1  and dbo.Atenciones.esPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso asc, Atenciones.HoraIngreso asc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     Else
        sSql = sSql & " WHERE dbo.Atenciones.FechaIngreso=CONVERT(DATETIME,'" & lcFechaIngreso & "',103)" & _
                      "      and dbo.Atenciones.IdServicioIngreso=" & lnIdServicio & _
                      "      and dbo.Atenciones.idTipoServicio=1  and dbo.Atenciones.esPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso asc, Atenciones.HoraIngreso asc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     End If

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarCEPorCuentaPorHistoriaPorApellidosPorServicio"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1500, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarCEPorCuentaPorHistoriaPorApellidosPorServicio = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function AtencionesSeleccionarEmergPorCuentaPorHistoriaPorApellidosPorServicio(lnHistoriaClinica As Long, lnIdCuentaAtencion As Long, lcApellidoPaterno As String, lcFechaIngreso As String, lnIdServicio As Long, lcDNI As String) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim sWhere As String
     sSql = ""
     If lnHistoriaClinica > 0 Then
        sSql = sSql & " WHERE  dbo.Pacientes.NroHistoriaClinica=" & lnHistoriaClinica & _
                      "        and dbo.Atenciones.IdServicioEgreso =" & lnIdServicio & _
                      "        and dbo.Atenciones.idTipoServicio in (2,4) and dbo.atenciones.EsPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     ElseIf lnIdCuentaAtencion > 0 Then
        sSql = sSql & " WHERE dbo.Atenciones.IdCuentaAtencion=" & lnIdCuentaAtencion & _
                      "        and dbo.Atenciones.IdServicioEgreso =" & lnIdServicio & _
                      "        and dbo.Atenciones.idTipoServicio in (2,4) and dbo.atenciones.EsPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     ElseIf lcApellidoPaterno <> "" Then
        sSql = sSql & " WHERE  dbo.Pacientes.ApellidoPaterno like '" & lcApellidoPaterno & "%'" & _
                      "        and dbo.Atenciones.IdServicioEgreso =" & lnIdServicio & _
                      "        and dbo.Atenciones.idTipoServicio in (2,4) and dbo.atenciones.EsPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     ElseIf lcDNI <> "" Then
        sSql = sSql & " WHERE  dbo.Pacientes.idDocIdentidad=1 and dbo.Pacientes.NroDocumento ='" & lcDNI & "'" & _
                      "        and dbo.Atenciones.IdServicioEgreso =" & lnIdServicio & _
                      "        and dbo.Atenciones.idTipoServicio in (2,4) and dbo.atenciones.EsPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     ElseIf UCase(lcFechaIngreso) = "TODAS" Then
        sSql = sSql & " WHERE  dbo.Atenciones.IdServicioEgreso =" & lnIdServicio & _
                      "        and dbo.Atenciones.idTipoServicio in (2,4) and dbo.atenciones.EsPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     Else
        sSql = sSql & " WHERE  dbo.Atenciones.FechaIngreso=CONVERT(DATETIME,'" & lcFechaIngreso & "',103)" & _
                      "        and dbo.Atenciones.IdServicioEgreso =" & lnIdServicio & _
                      "        and dbo.Atenciones.idTipoServicio in (2,4) and dbo.atenciones.EsPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     End If

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarEmergPorCuentaPorHistoriaPorApellidosPorServicio"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1500, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarEmergPorCuentaPorHistoriaPorApellidosPorServicio = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function AtencionesSeleccionarHospPorCuentaPorHistoriaPorApellidosPorServicio(lnHistoriaClinica As Long, lnIdCuentaAtencion As Long, lcApellidoPaterno As String, lcFechaIngreso As String, lnIdServicio As Long, lcDNI As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim sWhere As String
Dim ldFechaIngreso As Date
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lcParametro289 As String
     If SIGHENTIDADES.EsFecha(lcFechaIngreso, "DD/MM/AAAA") Then
        ldFechaIngreso = CDate(lcFechaIngreso)
     Else
        'lcFechaIngreso = Date
     End If
     lcParametro289 = Trim(lcBuscaParametro.SeleccionaFilaParametro(289))
     If lnHistoriaClinica > 0 Then
        sSql = sSql & " WHERE  dbo.Pacientes.NroHistoriaClinica=" & lnHistoriaClinica & _
                      "        and dbo.Atenciones.IdServicioEgreso =" & lnIdServicio & _
                      "        and dbo.Atenciones.idTipoServicio =3   and dbo.Atenciones.esPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     ElseIf lnIdCuentaAtencion > 0 Then
        sSql = sSql & " WHERE  dbo.Atenciones.IdCuentaAtencion=" & lnIdCuentaAtencion & _
                      "        and dbo.Atenciones.IdServicioEgreso =" & lnIdServicio & _
                      "        and dbo.Atenciones.idTipoServicio =3   and dbo.Atenciones.esPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     ElseIf lcApellidoPaterno <> "" Then
        sSql = sSql & "  WHERE dbo.Pacientes.ApellidoPaterno like '" & lcApellidoPaterno & "%'" & _
                      "        and dbo.Atenciones.IdServicioEgreso =" & lnIdServicio & _
                      "        and dbo.Atenciones.idTipoServicio =3   and dbo.Atenciones.esPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     ElseIf lcDNI <> "" Then
        sSql = sSql & " WHERE   dbo.Pacientes.IdDocIdentidad=1 and dbo.Pacientes.NroDocumento ='" & lcDNI & "'" & _
                      "        and dbo.Atenciones.IdServicioEgreso =" & lnIdServicio & _
                      "        and dbo.Atenciones.idTipoServicio =3   and dbo.Atenciones.esPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     ElseIf UCase(lcFechaIngreso) = "TODAS" Then
        sSql = sSql & " WHERE   dbo.Atenciones.IdServicioEgreso =" & lnIdServicio & _
                      "        and dbo.Atenciones.idTipoServicio =3   and dbo.Atenciones.esPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     Else
        sSql = sSql & " WHERE   dbo.Atenciones.FechaIngreso=CONVERT(DATETIME,'" & lcFechaIngreso & "',103) and dbo.Atenciones.IdServicioEgreso =" & lnIdServicio & _
                      "        and dbo.Atenciones.idTipoServicio =3   and dbo.Atenciones.esPacienteExterno<>1" & _
                      " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     End If

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarHospPorCuentaPorHistoriaPorApellidosPorServicio"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1500, sSql): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcParametro289", adVarChar, adParamInput, 1, lcParametro289): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
  Set AtencionesSeleccionarHospPorCuentaPorHistoriaPorApellidosPorServicio = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function





Function AtencionesSeleccionarPacExtPorFechas1(ldFechaIni As Date, ldFechaFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarPacExtPorFechas"
        Set oParameter = .CreateParameter("@ldFechaIni", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ldFechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPacExtPorFechas1 = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function AtencionesSeleccionarPacExtPorCuentaHistoriaApellidosServPARTIC(lnHistoriaClinica As Long, lnIdCuentaAtencion As Long, lcApellidoPaterno As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim sWhere As String
     sSql = ""
     If lnHistoriaClinica > 0 Then
        sSql = sSql & " WHERE dbo.Pacientes.NroHistoriaClinica=" & lnHistoriaClinica & _
                      "       and dbo.atenciones.EsPacienteExterno=1 and dbo.TiposFinanciamiento.GeneraPago=1" & _
                      " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     ElseIf lnIdCuentaAtencion > 0 Then
        sSql = sSql & " WHERE dbo.Atenciones.IdCuentaAtencion=" & lnIdCuentaAtencion & _
                      "       and dbo.atenciones.EsPacienteExterno=1 and dbo.TiposFinanciamiento.GeneraPago=1" & _
                      " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     ElseIf lcApellidoPaterno <> "" Then
        sSql = sSql & " WHERE dbo.Pacientes.ApellidoPaterno like '" & lcApellidoPaterno & "%'" & _
                      "       and dbo.atenciones.EsPacienteExterno=1 and dbo.TiposFinanciamiento.GeneraPago=1" & _
                      " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     Else
        sSql = sSql & " WHERE dbo.atenciones.EsPacienteExterno=1 and dbo.TiposFinanciamiento.GeneraPago=1" & _
                      " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     End If

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarPacExtPorCuentaHistoriaApellidosServPARTIC"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1500, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPacExtPorCuentaHistoriaApellidosServPARTIC = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AtencionesSeleccionarPacExtPorCuentaHistoriaApellidosServSEGUROS(lnHistoriaClinica As Long, _
                                            lnIdCuentaAtencion As Long, lcApellidoPaterno As String, _
                                            lcApellidoVacio As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim sWhere As String
     sSql = ""
     If lnHistoriaClinica > 0 Then
            sSql = sSql & " WHERE dbo.Pacientes.NroHistoriaClinica=" & lnHistoriaClinica & _
                          "       and dbo.atenciones.EsPacienteExterno=1 and dbo.TiposFinanciamiento.GeneraPago>1" & _
                          " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     ElseIf lnIdCuentaAtencion > 0 Then
            sSql = sSql & " WHERE dbo.Atenciones.IdCuentaAtencion=" & lnIdCuentaAtencion & _
                          "       and dbo.atenciones.EsPacienteExterno=1 and dbo.TiposFinanciamiento.GeneraPago>1" & _
                          " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     ElseIf lcApellidoPaterno <> "" Then
            If lcApellidoPaterno = lcApellidoVacio Then
                sSql = sSql & " WHERE dbo.Pacientes.ApellidoPaterno = '" & lcApellidoPaterno & "'" & _
                              "       and dbo.atenciones.EsPacienteExterno=1 and dbo.TiposFinanciamiento.GeneraPago>1" & _
                              " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
            Else
                sSql = sSql & " WHERE dbo.Pacientes.ApellidoPaterno like '" & lcApellidoPaterno & "%'" & _
                              "       and dbo.atenciones.EsPacienteExterno=1 and dbo.TiposFinanciamiento.GeneraPago>1" & _
                              " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
            End If
     Else
            sSql = sSql & " WHERE dbo.atenciones.EsPacienteExterno=1 and dbo.TiposFinanciamiento.GeneraPago>1" & _
                          " order by Atenciones.FechaIngreso desc, Atenciones.HoraIngreso desc, Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
     End If

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarPacExtPorCuentaHistoriaApellidosServSEGUROS"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1500, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPacExtPorCuentaHistoriaApellidosServSEGUROS = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarXidCuentaAtencion(lnIdCuentaAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim oParameter As ADODB.Parameter
     
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandType = adCmdStoredProc
       .CommandText = "atencionesXIdCuentaAtencion"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
     Set SeleccionarXidCuentaAtencion = oRecordset
     ms_MensajeError = ""
     Set oCommand = Nothing
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "ReporteParaHistoriaClinica"
Exit Function
End Function
