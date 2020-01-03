VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TurnoNegocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim ms_MensajeError As String
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim lcBuscaParametro As New SIGHDatos.Parametros
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

Function TurnosAgregar(oDOTurno As turnoComun, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDturno As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTurno As New TurnoDatos
    TurnosAgregar = False
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oTurno.Conexion = oConexion
    If oTurno.Insertar(oDOTurno) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOTurno.IdUsuarioAuditoria, "A", oDOTurno.idTurno, "Turnos", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDturno)          'ListBarItems.idListItem
        TurnosAgregar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oTurno.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oTurno = Nothing
End Function

Function TurnosModificar(oDOTurno As turnoComun, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDturno As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTurno As New TurnoDatos
    TurnosModificar = False
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oTurno.Conexion = oConexion
    If oTurno.Modificar(oDOTurno) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOTurno.IdUsuarioAuditoria, "M", oDOTurno.idTurno, "Turnos", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDturno)          'ListBarItems.idListItem
        TurnosModificar = True
        oConexion.CommitTrans
    Else
        oConexion.RollbackTrans
        ms_MensajeError = oTurno.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oTurno = Nothing
End Function

Function TurnosEliminar(oDOTurno As turnoComun, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDturno As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTurno As New TurnoDatos
    
    TurnosEliminar = False
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oTurno.Conexion = oConexion
    If oTurno.Eliminar(oDOTurno) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOTurno.IdUsuarioAuditoria, "E", oDOTurno.idTurno, "Turnos", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDturno)          'ListBarItems.idListItem
        TurnosEliminar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oTurno.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oTurno = Nothing
End Function

Function TurnosSeleccionarPorId(lnIdTurno As Integer) As turnoComun
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oTabla As New turnoComun
Dim oConexion As New ADODB.Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
   Set TurnosSeleccionarPorId = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = oConexion
     .CommandText = "TurnosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, lnIdTurno): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       ActualizaTabla oTabla, oRecordset
       Set TurnosSeleccionarPorId = oTabla
   End If

   oRecordset.Close
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
   ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Sub ActualizaTabla(oTabla As turnoComun, oRecordset As Recordset)
       oTabla.IdEspecialidad = IIf(IsNull(oRecordset!IdEspecialidad), 0, oRecordset!IdEspecialidad)
       oTabla.codigo = IIf(IsNull(oRecordset!codigo), "", oRecordset!codigo)
       oTabla.IdTipoServicio = IIf(IsNull(oRecordset!IdTipoServicio), 0, oRecordset!IdTipoServicio)
       oTabla.HoraFin = IIf(IsNull(oRecordset!HoraFin), "", oRecordset!HoraFin)
       oTabla.HoraInicio = IIf(IsNull(oRecordset!HoraInicio), "", oRecordset!HoraInicio)
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
       oTabla.idTurno = IIf(IsNull(oRecordset!idTurno), 0, oRecordset!idTurno)

End Sub


Function TiposServicioSeleccionarAsistenciales() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
   oConexion.CommandTimeout = 300
   oConexion.CursorLocation = adUseClient
   oConexion.Open sighentidades.CadenaConexion
   Set TiposServicioSeleccionarAsistenciales = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = oConexion
     .CommandText = "TiposServicioSeleccionarAsistenciales"
     Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set TiposServicioSeleccionarAsistenciales = oRecordset
 
   ms_MensajeError = ""
   Set oCommand = Nothing
   Set oRecordset = Nothing
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function TurnosSeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
   oConexion.CommandTimeout = 300
   oConexion.CursorLocation = adUseClient
   oConexion.Open sighentidades.CadenaConexion
   Set TurnosSeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = oConexion
     .CommandText = "TurnosSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set TurnosSeleccionarTodos = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



