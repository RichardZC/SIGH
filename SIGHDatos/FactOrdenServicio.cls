VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FactOrdenServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FactOrdenServicio
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String

Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property


'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla FactOrdenServicio
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoFactOrdenServ) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
'If oTabla.IdCuentaAtencion > 0 Then
'   oTabla.IdServicioPaciente = DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(oTabla.IdCuentaAtencion, oTabla.FechaCreacion, Format(Now, sighentidades.DevuelveHoraSoloFormato_HM))
'End If

Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FactOrdenServicioAgregar"
   Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.IdPuntoCarga = 0, Null, oTabla.IdPuntoCarga))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicioPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdServicioPaciente = 0, Null, oTabla.IdServicioPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaDespacho", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaDespacho = 0, Null, oTabla.FechaDespacho))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioDespacho", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioDespacho = 0, Null, oTabla.IdUsuarioDespacho))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoFacturacion", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoFacturacion = 0, Null, oTabla.IdEstadoFacturacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaHoraRealizaCpt", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaHoraRealizaCpt = 0, Null, oTabla.FechaHoraRealizaCpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   'SCCQ 26-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
   'HRA 10/12/2020 Cambio 46 Inicio
   'Set oParameter = .CreateParameter("@FuaCodigoPrestacion", adVarChar, adParamInput, 3, IIf(oTabla.CodPres = "", Null, oTabla.CodPres))
   '.Parameters.Append oParameter
   'HRA 10/12/2020 Cambio 46 Fin
   'SCCQ 26-04-2021 Cambio 64 Fin (Reversion Cambio 46)
   .Execute
   oTabla.IdOrden = .Parameters("@IdOrden")
End With
 
Insertar = True
ms_MensajeError = ""
GrabaHistorico oTabla
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub GrabaHistorico(oTabla As DoFactOrdenServ)
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
    lcSql = "INSERT INTO FactOrdenServicio" & _
"           (IdPuntoCarga,IdPaciente" & _
"           ,IdCuentaAtencion,IdServicioPaciente" & _
"           ,idTipoFinanciamiento,idFuenteFinanciamiento" & _
"           ,FechaCreacion,IdUsuario" & _
"           ,FechaDespacho,IdUsuarioDespacho" & _
"           ,IdEstadoFacturacion,FechaHoraRealizaCpt,idOrden)" & _
"     Values" & _
"           (" & oTabla.IdPuntoCarga & "," & oTabla.IdPaciente & _
"           ," & oTabla.IdCuentaAtencion & "," & oTabla.IdServicioPaciente & _
"           ," & oTabla.idTipoFinanciamiento & "," & oTabla.idFuenteFinanciamiento & _
"           ,'" & oTabla.FechaCreacion & "'," & oTabla.IdUsuario & _
"           ,'" & oTabla.FechaDespacho & "'," & oTabla.IdUsuarioAuditoria & _
"           ," & oTabla.IdEstadoFacturacion & ",'" & oTabla.FechaHoraRealizaCpt & "'," & oTabla.IdOrden & ")"
   oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    oConexionMDB.Close
ErrHistorico:
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing

End Sub


'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla FactOrdenServicio
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoFactOrdenServ) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter

'If oTabla.IdCuentaAtencion > 0 Then
'   oTabla.IdServicioPaciente = DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(oTabla.IdCuentaAtencion, oTabla.FechaCreacion, Format(Now, sighentidades.DevuelveHoraSoloFormato_HM))
'End If
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FactOrdenServicioModificar"
   Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.IdPuntoCarga = 0, Null, oTabla.IdPuntoCarga))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicioPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdServicioPaciente = 0, Null, oTabla.IdServicioPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaDespacho", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaDespacho = 0, Null, oTabla.FechaDespacho))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioDespacho", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioDespacho = 0, Null, oTabla.IdUsuarioDespacho))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoFacturacion", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoFacturacion = 0, Null, oTabla.IdEstadoFacturacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaHoraRealizaCpt", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaHoraRealizaCpt = 0, Null, oTabla.FechaHoraRealizaCpt))
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
'   Descripción:    elimia o borra un registro a la tabla FactOrdenServicio
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoFactOrdenServ) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FactOrdenServicioEliminar"
   Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, oTabla.IdOrden)
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
'   Descripción:    Selecciona un unico registro de la tabla FactOrdenServicio
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoFactOrdenServ) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     'Set .ActiveConnection = sighentidades.CadenaConexion
     .CommandText = "FactOrdenServicioSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, oTabla.IdOrden)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdOrden = IIf(IsNull(oRecordset!IdOrden), 0, oRecordset!IdOrden)
       oTabla.IdPuntoCarga = IIf(IsNull(oRecordset!IdPuntoCarga), 0, oRecordset!IdPuntoCarga)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.IdServicioPaciente = IIf(IsNull(oRecordset!IdServicioPaciente), 0, oRecordset!IdServicioPaciente)
       oTabla.idTipoFinanciamiento = IIf(IsNull(oRecordset!idTipoFinanciamiento), 0, oRecordset!idTipoFinanciamiento)
       oTabla.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
       oTabla.IdUsuario = IIf(IsNull(oRecordset!IdUsuario), 0, oRecordset!IdUsuario)
       oTabla.IdEstadoFacturacion = IIf(IsNull(oRecordset!IdEstadoFacturacion), 0, oRecordset!IdEstadoFacturacion)
       oTabla.idFuenteFinanciamiento = IIf(IsNull(oRecordset!idFuenteFinanciamiento), 0, oRecordset!idFuenteFinanciamiento)
       oTabla.FechaDespacho = IIf(IsNull(oRecordset!FechaDespacho), 0, oRecordset!FechaDespacho)
       oTabla.IdUsuarioDespacho = IIf(IsNull(oRecordset!IdUsuarioDespacho), 0, oRecordset!IdUsuarioDespacho)
       oTabla.FechaHoraRealizaCpt = IIf(IsNull(oRecordset!FechaHoraRealizaCpt), 0, oRecordset!FechaHoraRealizaCpt)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoFactOrdenServ)
   oTabla.IdOrden = 0
   oTabla.IdPuntoCarga = 0
   oTabla.IdPaciente = 0
   oTabla.IdCuentaAtencion = 0
   oTabla.IdServicioPaciente = 0
   oTabla.idTipoFinanciamiento = 0
   oTabla.idFuenteFinanciamiento = 0
   oTabla.FechaCreacion = Now
   oTabla.IdUsuario = 0
   oTabla.IdEstadoFacturacion = 0
   oTabla.IdUsuarioDespacho = 0
   oTabla.FechaDespacho = Now
End Sub


'Function DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(lidCuenta As Long, ldFecha As Date, lcHora As String) As Long
'On Error GoTo ManejadorDeError
'Dim oRecordset As New ADODB.Recordset
'Dim oCommand As New ADODB.Command
'Dim oParameter As ADODB.Parameter
''Dim oConexion As New ADODB.Connection
'Dim ms_MensajeError As String
'Dim lnIdServicioDelPaciente As Long
'    ms_MensajeError = ""
' '   oConexion.Open sighentidades.CadenaConexion
' '   oConexion.CursorLocation = adUseClient
'    With oCommand
'        .CommandType = adCmdStoredProc
' '       Set .ActiveConnection = oConexion
'        Set .ActiveConnection = mo_Conexion
'        .CommandTimeout = 150
'        .CommandText = "AtencionesEstanciaHospitalariaPorIdCuenta"
'        Set oParameter = .CreateParameter("@idCuenta", adInteger, adParamInput, 0, lidCuenta): .Parameters.Append oParameter
'        Set oRecordset = .Execute
'        Set oRecordset.ActiveConnection = Nothing
'   End With
'   lnIdServicioDelPaciente = 0
'   If oRecordset.RecordCount > 0 Then
'      oRecordset.MoveFirst
'      Do While Not oRecordset.EOF
'         If IsNull(oRecordset.Fields!fechaDesocupacion) Then
'            If ldFecha >= oRecordset.Fields!fechaOcupacion Then
'               lnIdServicioDelPaciente = oRecordset.Fields!IdServicio
'               Exit Do
'            End If
'         Else
'            If ldFecha >= oRecordset.Fields!fechaOcupacion And ldFecha <= oRecordset.Fields!fechaDesocupacion Then
'               If ldFecha = oRecordset.Fields!fechaOcupacion And lcHora >= oRecordset.Fields!HoraOcupacion And lcHora <= oRecordset.Fields!HoraDesocupacion Then
'                  lnIdServicioDelPaciente = oRecordset.Fields!IdServicio
'                  Exit Do
'               End If
'            End If
'         End If
'         oRecordset.MoveNext
'      Loop
'   End If
'   'oConexion.Close
'   'Set oConexion = Nothing
'   Set oCommand = Nothing
'   Set oRecordset = Nothing
'   DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora = lnIdServicioDelPaciente
'Exit Function
'ManejadorDeError:
'   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
'Exit Function
'End Function
'
