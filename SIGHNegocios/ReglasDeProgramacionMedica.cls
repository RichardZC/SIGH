VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasDeProgMedica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Reglas de Programación Médica
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
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

Function TurnosAgregar(oDOTurno As doTurno, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDturno As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTurno As New Turno
    TurnosAgregar = False
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oTurno.Conexion = oConexion
    If oTurno.Insertar(oDOTurno) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOTurno.IdUsuarioAuditoria, "A", oDOTurno.IdTurno, "Turnos", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDturno)          'ListBarItems.idListItem
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

Function TurnosModificar(oDOTurno As doTurno, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDturno As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTurno As New Turno
    TurnosModificar = False
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oTurno.Conexion = oConexion
    If oTurno.modificar(oDOTurno) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOTurno.IdUsuarioAuditoria, "M", oDOTurno.IdTurno, "Turnos", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDturno)          'ListBarItems.idListItem
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

Function TurnosEliminar(oDOTurno As doTurno, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDturno As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTurno As New Turno
    
    TurnosEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oTurno.Conexion = oConexion
    If oTurno.Eliminar(oDOTurno) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOTurno.IdUsuarioAuditoria, "E", oDOTurno.IdTurno, "Turnos", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDturno)          'ListBarItems.idListItem
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

Function TurnosSeleccionarPorId(IdTurno As Integer) As doTurno
Dim oConexion As New ADODB.Connection
Dim oTurno As New Turno
Dim oDOTurno As New doTurno
    
    Set TurnosSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTurno.Conexion = oConexion
    oDOTurno.IdTurno = IdTurno
    If oTurno.SeleccionarPorId(oDOTurno) Then
        Set TurnosSeleccionarPorId = oDOTurno
    Else
        Set TurnosSeleccionarPorId = Nothing
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oTurno = Nothing
    ms_MensajeError = oTurno.MensajeError

End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function TurnosSeleccionarPorIdTipoServicio(lnIdTipoServicio As Long) As ADODB.Recordset
    Dim oRsTmp As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
  
    Set TurnosSeleccionarPorIdTipoServicio = Nothing
    
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "TurnosSeleccionarPorIdTipoServicio"
      Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, lnIdTipoServicio): .Parameters.Append oParameter
      Set oRsTmp = .Execute
      Set oRsTmp.ActiveConnection = Nothing
    End With
  
    Set TurnosSeleccionarPorIdTipoServicio = oRsTmp
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function TurnosSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTurno As New Turno
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    Set oTurno.Conexion = oConexion
    Set TurnosSeleccionarTodos = oTurno.SeleccionarTodos
    oConexion.Close
    Set oConexion = Nothing
    Set oTurno = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function TurnosSeleccionarPorEspecialidad(lIdEspecialidad As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTurno As New Turno
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTurno.Conexion = oConexion
    Set TurnosSeleccionarPorEspecialidad = oTurno.SeleccionarPorEspecialidad(lIdEspecialidad)
    oConexion.Close
     Set oConexion = Nothing
    Set oTurno = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function TurnosSeleccionarPorCodigo(sCodigo As String) As doTurno
Dim oConexion As New ADODB.Connection
Dim oTabla As New Turno
Dim doTurno As New doTurno
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    doTurno.Codigo = sCodigo
    If oTabla.SeleccionarPorCodigo(doTurno) Then
        Set TurnosSeleccionarPorCodigo = doTurno
    Else
        Set TurnosSeleccionarPorCodigo = Nothing
    End If
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'------------------------------------------------------------------------------
'                               MEDICOS
'------------------------------------------------------------------------------
'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MedicosAgregar(oDoMedico As DOMedico, oDOEmpleado As DOEmpleado, oMedicosEspecialidad As Collection, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcMedico As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oMedico As New Medicos
Dim oEmpleado As New Empleados
Dim oMedicoEspecialidad As New MedicosEspecialidad
Dim oDOMedicoEspecialidad As DOMedicoEspecialidad
Dim bEstadoTransaccion As Boolean
    
    MedicosAgregar = False
    bEstadoTransaccion = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    oConexion.BeginTrans
    
    Set oMedico.Conexion = oConexion
    Set oEmpleado.Conexion = oConexion
    Set oMedicoEspecialidad.Conexion = oConexion
    
    If oEmpleado.Insertar(oDOEmpleado) Then
        oDoMedico.idEmpleado = oDOEmpleado.idEmpleado
        If oMedico.Insertar(oDoMedico) Then
            If oMedicosEspecialidad.Count > 0 Then
                For Each oDOMedicoEspecialidad In oMedicosEspecialidad
                    oDOMedicoEspecialidad.IdMedico = oDoMedico.IdMedico
                    If oMedicoEspecialidad.Insertar(oDOMedicoEspecialidad) Then
                        bEstadoTransaccion = True
                    Else
                        bEstadoTransaccion = False
                        Exit For
                    End If
                Next
            Else
                bEstadoTransaccion = True
            End If
        End If
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMedico.IdUsuarioAuditoria, "A", oDoMedico.IdMedico, "Medicos", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcMedico)          'ListBarItems.idListItem
    
    If bEstadoTransaccion Then
        oConexion.CommitTrans
        MedicosAgregar = True
        
    Else
        oConexion.RollbackTrans
        MedicosAgregar = False
        ms_MensajeError = oMedico.MensajeError
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oMedico = Nothing
    Set oEmpleado = Nothing
    Set oMedicoEspecialidad = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MedicosModificar(oDoMedico As DOMedico, oDOEmpleado As DOEmpleado, oMedicosEspecialidad As Collection, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcMedico As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oMedico As New Medicos
Dim oEmpleado As New Empleados
Dim oMedicoEspecialidad As New MedicosEspecialidad
Dim oDOMedicoEspecialidad As DOMedicoEspecialidad
Dim bEstadoTransaccion As Boolean
    
    MedicosModificar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    oConexion.BeginTrans
    Set oMedico.Conexion = oConexion
    Set oEmpleado.Conexion = oConexion
    Set oMedicoEspecialidad.Conexion = oConexion
    
    If oEmpleado.modificar(oDOEmpleado) Then
        If oMedico.modificar(oDoMedico) Then
            If oMedicoEspecialidad.EliminarPorMedico(oDoMedico) Then
                If oMedicosEspecialidad.Count > 0 Then
                    For Each oDOMedicoEspecialidad In oMedicosEspecialidad
                        If oMedicoEspecialidad.Insertar(oDOMedicoEspecialidad) Then
                            bEstadoTransaccion = True
                        Else
                            bEstadoTransaccion = False
                            Exit For
                        End If
                    Next
                Else
                    bEstadoTransaccion = True
                End If
            End If
        End If
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMedico.IdUsuarioAuditoria, "M", oDoMedico.IdMedico, "Medicos", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcMedico)        'ListBarItems.idListItem
    
    If bEstadoTransaccion Then
        oConexion.CommitTrans
        MedicosModificar = True
    Else
        oConexion.RollbackTrans
        MedicosModificar = False
        ms_MensajeError = oMedico.MensajeError
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oMedico = Nothing
    Set oEmpleado = Nothing
    Set oMedicoEspecialidad = Nothing
    
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MedicosEliminar(oDoMedico As DOMedico, oDOEmpleado As DOEmpleado, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcMedico As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oMedico As New Medicos
Dim oEmpleado As New Empleados
Dim oMedicoEspecialidad As New MedicosEspecialidad
Dim bEstadoTransaccion As Boolean
Dim oProgMedica As New ProgramacionMedica
    
    MedicosEliminar = False
    bEstadoTransaccion = False
        
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 900
    oConexion.Open sighentidades.CadenaConexion
    
    oConexion.BeginTrans
    Set oMedico.Conexion = oConexion
    Set oEmpleado.Conexion = oConexion
    Set oMedicoEspecialidad.Conexion = oConexion
    Set oProgMedica.Conexion = oConexion
    
    Dim rsProgMedica As New Recordset
    Set rsProgMedica = oProgMedica.SeleccionarPorMedico(oDoMedico.IdMedico)
    If rsProgMedica.RecordCount > 0 Then
        MsgBox "No se puede eliminar el médico dado que tiene programaciones médicas, debe eliminar antes las programaciones y luego eliminar el médico", vbInformation, "Programación médica"
    Else
        If oMedicoEspecialidad.EliminarPorMedico(oDoMedico) Then
            If oMedico.Eliminar(oDoMedico) Then
                If oEmpleado.Eliminar(oDOEmpleado) Then
                    bEstadoTransaccion = True
                End If
            End If
        End If
    End If
    rsProgMedica.Close
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMedico.IdUsuarioAuditoria, "E", oDoMedico.IdMedico, "Medicos", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcMedico)         'ListBarItems.idListItem
    
    If bEstadoTransaccion Then
        oConexion.CommitTrans
        MedicosEliminar = True
    Else
        oConexion.RollbackTrans
        MedicosEliminar = False
        ms_MensajeError = oMedicoEspecialidad.MensajeError + oMedico.MensajeError + oEmpleado.MensajeError
    End If
        
    oConexion.Close
    Set oConexion = Nothing
    Set oMedico = Nothing
    Set oEmpleado = Nothing
    Set oMedicoEspecialidad = Nothing
End Function

Function MedicosSeleccionarPorId(IdMedico As Long, oDoMedico As DOMedico, oDOEmpleado As DOEmpleado, oMedicosEspecialidad As Collection, oConexion As Connection) As Boolean
Dim oMedico As New Medicos
Dim oEmpleado As New Empleados
Dim oMedicoEspecialidad As New MedicosEspecialidad
Dim rsMedicoEspecialidad As Recordset
    
    MedicosSeleccionarPorId = False
    Set oMedico.Conexion = oConexion
    Set oEmpleado.Conexion = oConexion
    Set oMedicoEspecialidad.Conexion = oConexion
    
    oDoMedico.IdMedico = IdMedico
    If oMedico.SeleccionarPorId(oDoMedico) Then
        oDOEmpleado.idEmpleado = oDoMedico.idEmpleado
        If oEmpleado.SeleccionarPorId(oDOEmpleado) Then
            MedicosSeleccionarPorId = True
        End If
    End If
    
    Dim oDOMedicoEspecialidad As DOMedicoEspecialidad
    Set rsMedicoEspecialidad = oMedicoEspecialidad.SeleccionarPorMedico(IdMedico)
    Do While Not rsMedicoEspecialidad.EOF
        Set oDOMedicoEspecialidad = New DOMedicoEspecialidad
        oDOMedicoEspecialidad.IdEspecialidad = rsMedicoEspecialidad!IdEspecialidad
        oDOMedicoEspecialidad.IdMedico = rsMedicoEspecialidad!IdMedico
        oDOMedicoEspecialidad.IdMedicoEspecialidad = rsMedicoEspecialidad!IdMedico
        oMedicosEspecialidad.Add oDOMedicoEspecialidad
        rsMedicoEspecialidad.MoveNext
    Loop
    ms_MensajeError = oMedico.MensajeError
    
    Set oMedico = Nothing
    Set oEmpleado = Nothing
    Set oMedicoEspecialidad = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MedicosSeleccionarPorCodigo1(Codigo As String, oDoMedico As DOMedico, oDOEmpleado As DOEmpleado, oMedicosEspecialidad As Collection) As Boolean
Dim oConexion As New ADODB.Connection
Dim oMedico As New Medicos
Dim oEmpleado As New Empleados
Dim oMedicoEspecialidad As New MedicosEspecialidad
Dim rsMedicoEspecialidad As Recordset
    
    
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
      
    Set oMedico.Conexion = oConexion
    Set oEmpleado.Conexion = oConexion
    Set oMedicoEspecialidad.Conexion = oConexion
    
    oDOEmpleado.CodigoPlanilla = Codigo
    If oMedico.SeleccionarPorCodigo(oDoMedico, oDOEmpleado) Then
        oDOEmpleado.idEmpleado = oDoMedico.idEmpleado
        If oEmpleado.SeleccionarPorId(oDOEmpleado) Then
            MedicosSeleccionarPorCodigo1 = True
        End If
    End If
    
    Dim oDOMedicoEspecialidad As DOMedicoEspecialidad
    Set rsMedicoEspecialidad = oMedicoEspecialidad.SeleccionarPorMedico(oDoMedico.IdMedico)
    Do While Not rsMedicoEspecialidad.EOF
        Set oDOMedicoEspecialidad = New DOMedicoEspecialidad
        oDOMedicoEspecialidad.IdEspecialidad = rsMedicoEspecialidad!IdEspecialidad
        oDOMedicoEspecialidad.IdMedico = rsMedicoEspecialidad!IdMedico
        oDOMedicoEspecialidad.IdMedicoEspecialidad = rsMedicoEspecialidad!IdMedico
        oMedicosEspecialidad.Add oDOMedicoEspecialidad
        rsMedicoEspecialidad.MoveNext
    Loop
    oConexion.Close
    
    ms_MensajeError = oMedico.MensajeError

    Set oConexion = Nothing
    Set oMedico = Nothing
    Set oEmpleado = Nothing
    Set oMedicoEspecialidad = Nothing
    Set rsMedicoEspecialidad = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MedicosSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oMedico As New Medicos
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oMedico.Conexion = oConexion
    Set MedicosSeleccionarTodos = oMedico.SeleccionarTodos
    oConexion.Close
    
    ms_MensajeError = oMedico.MensajeError
    Set oConexion = Nothing
    Set oMedico = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MedicosFiltrar(oDoMedico As DOMedico, oDOEmpleado As DOEmpleado, lIdEspecialidad As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oMedico As New Medicos
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    Set oMedico.Conexion = oConexion
    Set MedicosFiltrar = oMedico.Filtrar(oDoMedico, oDOEmpleado, lIdEspecialidad)
    oConexion.Close
    
    ms_MensajeError = oMedico.MensajeError
    Set oConexion = Nothing
    Set oMedico = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MedicosProgramadosFiltrar(oDoMedico As DOMedico, oDOEmpleado As DOEmpleado, lIdEspecialidad As Long, ldFecha As Date, lcHora As String, lnIdTipoServicio As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
'Dim ssql As String
'Dim sWhere As String

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
       
       Set MedicosProgramadosFiltrar = Nothing
       
       With oCommand
'           .CommandType = adCmdText
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
'            .CommandText = ssql
            .CommandTimeout = 150
            .CommandText = "MedicosProgramadosFiltrar"
            Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, lnIdTipoServicio): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, Format(ldFecha, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Hora", adChar, adParamInput, 5, lcHora): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 50, oDOEmpleado.ApellidoPaterno): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 50, oDOEmpleado.ApellidoMaterno): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Nombres", adVarChar, adParamInput, 50, oDOEmpleado.Nombres): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@CodigoPlanilla", adVarChar, adParamInput, 8, oDOEmpleado.CodigoPlanilla): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, lIdEspecialidad): .Parameters.Append oParameter
           Set oRecordset = .Execute
           Set oRecordset.ActiveConnection = Nothing
       End With
       Set MedicosProgramadosFiltrar = oRecordset
       oConexion.Close
       Set oConexion = Nothing
       Set oCommand = Nothing
       ms_MensajeError = ""
       Exit Function
ManejadorDeError:
       ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
       oConexion.Close
       Set oConexion = Nothing
       Set oCommand = Nothing
       Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MedicosFiltrarPorDptoYEspecialidadSql2000(IdDepartamento As Long, IdEspecialidad As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oMedico As New Medicos
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oMedico.Conexion = oConexion
    Set MedicosFiltrarPorDptoYEspecialidadSql2000 = oMedico.FiltrarPorDptosYEspecialidadEsActivo(IdDepartamento, IdEspecialidad)
    oConexion.Close
    
    ms_MensajeError = oMedico.MensajeError
    Set oConexion = Nothing
    Set oMedico = Nothing

End Function
'mgaray20141014
Function MedicosFiltrarPorDptoYEspecialidadConEspecialidad(IdDepartamento As Long, IdEspecialidad As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oMedico As New Medicos
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oMedico.Conexion = oConexion
    Set MedicosFiltrarPorDptoYEspecialidadConEspecialidad = oMedico.FiltrarPorDptosYEspecialidadEsActivoConEspecialidad(IdDepartamento, IdEspecialidad)
    oConexion.Close
    
    ms_MensajeError = oMedico.MensajeError
    Set oConexion = Nothing
    Set oMedico = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MedicosFiltrarPorProgramacionSql2000(IdDepartamento As Long, IdEspecialidad As Long, lnIdServicio As Long, daFecha As Date) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oMedico As New Medicos
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oMedico.Conexion = oConexion
    Set MedicosFiltrarPorProgramacionSql2000 = oMedico.FiltrarPorProgramacion(IdDepartamento, IdEspecialidad, lnIdServicio, daFecha)
    oConexion.Close
    Set oMedico = Nothing
    Set oConexion = Nothing
    ms_MensajeError = oMedico.MensajeError

End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MedicosFiltrarPorProgramacionYidTipoServicio(IdDepartamento As Long, IdEspecialidad As Long, _
                                     lnIdServicio As Long, daFecha As Date, lnIdTipoServicio As Long, _
                                     lnIdMedicoFiltro As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oMedico As New Medicos
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oMedico.Conexion = oConexion
    Set MedicosFiltrarPorProgramacionYidTipoServicio = oMedico.FiltrarPorProgramacion(IdDepartamento, IdEspecialidad, lnIdServicio, daFecha)
    MedicosFiltrarPorProgramacionYidTipoServicio.Filter = "idTipoServicio=" & Trim(Str(lnIdTipoServicio)) & IIf(lnIdMedicoFiltro > 0, " and idMedico=" & Trim(Str(lnIdMedicoFiltro)), "")
    oConexion.Close
    Set oMedico = Nothing
    Set oConexion = Nothing
    ms_MensajeError = oMedico.MensajeError

End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MedicosObtenerDepartamento(lIdMedico As Long) As Long
Dim oConexion As New ADODB.Connection
Dim oMedico As New Medicos
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    Set oMedico.Conexion = oConexion
    MedicosObtenerDepartamento = oMedico.ObtenerDepartamento(lIdMedico)
    oConexion.Close
    
    ms_MensajeError = oMedico.MensajeError
    Set oConexion = Nothing
    Set oMedico = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function ProgramacionMedicaAgregar(oProgramaciones As Collection, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcMedicoServicio As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oProgMedica As New ProgramacionMedica
Dim oDoProgramacion As DOProgramacionMedica
Dim bEstadoTransaccion As Boolean

    
    ProgramacionMedicaAgregar = False
    bEstadoTransaccion = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    oConexion.BeginTrans
    Set oProgMedica.Conexion = oConexion
    
    If oProgramaciones.Count > 0 Then
        For Each oDoProgramacion In oProgramaciones
            If oProgMedica.Insertar(oDoProgramacion) Then
                '
                Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoProgramacion.IdUsuarioAuditoria, "A", oDoProgramacion.idProgramacion, "ProgramacionMedica", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcMedicoServicio)        'ListBarItems.idListItem
                bEstadoTransaccion = True
            Else
                bEstadoTransaccion = False
                Exit For
            End If
        Next
    Else
        bEstadoTransaccion = True
    End If
    
    
    If bEstadoTransaccion Then
        oConexion.CommitTrans
        ProgramacionMedicaAgregar = True
    Else
        oConexion.RollbackTrans
        ProgramacionMedicaAgregar = False
        ms_MensajeError = oProgMedica.MensajeError
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oProgMedica = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function ProgramacionMedicaModificar(oDOProgramacionMedica As DOProgramacionMedica, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcMedicoServicio As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oProgramacionMedica As New ProgramacionMedica
    
    ProgramacionMedicaModificar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
      
    oConexion.BeginTrans
    Set oProgramacionMedica.Conexion = oConexion
    If oProgramacionMedica.modificar(oDOProgramacionMedica) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOProgramacionMedica.IdUsuarioAuditoria, "M", oDOProgramacionMedica.idProgramacion, "ProgramacionMedica", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcMedicoServicio)        'ListBarItems.idListItem
        ProgramacionMedicaModificar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oProgramacionMedica.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oProgramacionMedica = Nothing
    
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function ProgramacionMedicaEliminar(oDoProgramacion As DOProgramacionMedica, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcMedicoServicio As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oProgMedica As New ProgramacionMedica
Dim bEstadoTransaccion As Boolean
    
    ProgramacionMedicaEliminar = False
    bEstadoTransaccion = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    oConexion.BeginTrans
    Set oProgMedica.Conexion = oConexion
    
    If oProgMedica.Eliminar(oDoProgramacion) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoProgramacion.IdUsuarioAuditoria, "E", oDoProgramacion.idProgramacion, "ProgramacionMedica", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcMedicoServicio)        'ListBarItems.idListItem
        bEstadoTransaccion = True
    Else
        bEstadoTransaccion = False
    End If
    
    If bEstadoTransaccion Then
        oConexion.CommitTrans
        ProgramacionMedicaEliminar = True
    Else
        oConexion.RollbackTrans
        ProgramacionMedicaEliminar = False
        ms_MensajeError = oProgMedica.MensajeError
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oProgMedica = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function ProgramacionMedicaGrabar(oProgramaciones As Collection, lIdMedico As Long, iAnio As Integer, iMes As Integer, lIdUsuarioAuditoria As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oProgMedica As New ProgramacionMedica
Dim oDoProgramacion As DOProgramacionMedica
Dim bEstadoTransaccion As Boolean

    
    ProgramacionMedicaGrabar = False
    bEstadoTransaccion = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    oConexion.BeginTrans
    Set oProgMedica.Conexion = oConexion
    
    If oProgMedica.EliminarPorMedicoYMes(lIdMedico, iMes, iAnio, lIdUsuarioAuditoria) Then
        If oProgramaciones.Count > 0 Then
            For Each oDoProgramacion In oProgramaciones
                If oProgMedica.Insertar(oDoProgramacion) Then
                    bEstadoTransaccion = True
                Else
                    bEstadoTransaccion = False
                    Exit For
                End If
            Next
        Else
            bEstadoTransaccion = True
        End If
    End If
    
    
    If bEstadoTransaccion Then
        oConexion.CommitTrans
        ProgramacionMedicaGrabar = True
    Else
        oConexion.RollbackTrans
        ProgramacionMedicaGrabar = False
        ms_MensajeError = oProgMedica.MensajeError
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oProgMedica = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function ProgramacionMedicaLeerPorMedicoYMes(IdMedico As Long, iMes As Integer, iAnio As Integer) As Collection
Dim oConexion As New ADODB.Connection
Dim oProgramacionMedica As New ProgramacionMedica
Dim rsProgramaciones As ADODB.Recordset
Dim oProgramaciones As New Collection
Dim oProgramacion As DOProgramacionMedica
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oProgramacionMedica.Conexion = oConexion
    Set rsProgramaciones = oProgramacionMedica.SeleccionarPorMedicoYMes(IdMedico, iMes, iAnio)
    oConexion.Close
    ms_MensajeError = oProgramacionMedica.MensajeError
   
    Do While Not rsProgramaciones.EOF
       Set oProgramacion = New DOProgramacionMedica
       oProgramacion.idProgramacion = IIf(IsNull(rsProgramaciones!idProgramacion), 0, rsProgramaciones!idProgramacion)
       oProgramacion.IdMedico = IIf(IsNull(rsProgramaciones!IdMedico), 0, rsProgramaciones!IdMedico)
       oProgramacion.IdDepartamento = IIf(IsNull(rsProgramaciones!IdDepartamento), 0, rsProgramaciones!IdDepartamento)
       oProgramacion.idTipoServicio = IIf(IsNull(rsProgramaciones!idTipoServicio), 0, rsProgramaciones!idTipoServicio)
       oProgramacion.fecha = IIf(IsNull(rsProgramaciones!fecha), "", rsProgramaciones!fecha)
       oProgramacion.HoraInicio = IIf(IsNull(rsProgramaciones!HoraInicio), "", rsProgramaciones!HoraInicio)
       oProgramacion.horaFin = IIf(IsNull(rsProgramaciones!horaFin), "", rsProgramaciones!horaFin)
       oProgramacion.Descripcion = IIf(IsNull(rsProgramaciones!Descripcion), "", rsProgramaciones!Descripcion)
       oProgramacion.IdTipoProgramacion = IIf(IsNull(rsProgramaciones!IdTipoProgramacion), 0, rsProgramaciones!IdTipoProgramacion)
       oProgramacion.IdEspecialidad = IIf(IsNull(rsProgramaciones!IdEspecialidad), 0, rsProgramaciones!IdEspecialidad)
       oProgramacion.IdTurno = IIf(IsNull(rsProgramaciones!IdTurno), 0, rsProgramaciones!IdTurno)
       oProgramacion.IdServicio = IIf(IsNull(rsProgramaciones!IdServicio), 0, rsProgramaciones!IdServicio)
       oProgramacion.Color = IIf(IsNull(rsProgramaciones!Color), 0, rsProgramaciones!Color)
       
       oProgramaciones.Add oProgramacion
       rsProgramaciones.MoveNext
    Loop
    
    Set ProgramacionMedicaLeerPorMedicoYMes = oProgramaciones

End Function

Function CitasLeerPorMedicoYFecha(IdMedico As Long, daFecha As Date, oConexion As Connection) As Collection
Dim rsCitas As ADODB.Recordset
Dim oCita As New Citas
Dim doCitas As New Collection
Dim doCita As doCita
    
    Set oCita.Conexion = oConexion
    Set rsCitas = oCita.SeleccionarPorMedicoYFecha(IdMedico, daFecha)
    If oCita.MensajeError <> "" Then
        MsgBox oCita.MensajeError, vbInformation, "Asignación de Citas"
        Exit Function
    End If
       
    Do While Not rsCitas.EOF
       Set doCita = New doCita
       doCita.IdCita = IIf(IsNull(rsCitas!IdCita), 0, rsCitas!IdCita)
       doCita.fecha = IIf(IsNull(rsCitas!fecha), 0, rsCitas!fecha)
       doCita.HoraInicio = IIf(IsNull(rsCitas!HoraInicio), "", rsCitas!HoraInicio)
       doCita.horaFin = IIf(IsNull(rsCitas!horaFin), "", rsCitas!horaFin)
       doCita.IdEstadoCita = IIf(IsNull(rsCitas!IdEstadoCita), 0, rsCitas!IdEstadoCita)
       doCita.IdMedico = IIf(IsNull(rsCitas!IdMedico), 0, rsCitas!IdMedico)
       doCita.IdEspecialidad = IIf(IsNull(rsCitas!IdEspecialidad), 0, rsCitas!IdEspecialidad)
       doCita.idAtencion = IIf(IsNull(rsCitas!idAtencion), 0, rsCitas!idAtencion)
       doCita.idPaciente = IIf(IsNull(rsCitas!idPaciente), 0, rsCitas!idPaciente)
       doCita.idProgramacion = IIf(IsNull(rsCitas!idProgramacion), 0, rsCitas!idProgramacion)
       doCita.fechaSolicitud = IIf(IsNull(rsCitas!fechaSolicitud), 0, rsCitas!fechaSolicitud)
       doCita.horaSolicitud = IIf(IsNull(rsCitas!horaSolicitud), 0, rsCitas!horaSolicitud)
       doCita.EsCitaAdicional = IIf(IsNull(rsCitas!EsCitaAdicional), False, rsCitas!EsCitaAdicional)
       doCitas.Add doCita
       rsCitas.MoveNext
    Loop
    
    Set CitasLeerPorMedicoYFecha = doCitas

End Function
Function CitasLeerBloqueadasPorMedicoYFecha(IdMedico As Long, daFecha As Date, oConexion As Connection) As Collection
Dim oCita As New Citas
Dim rsCitas As ADODB.Recordset
Dim doCitas As New Collection
Dim doCita As DOCitaBloqueada
    
    Set oCita.Conexion = oConexion
    Set rsCitas = oCita.SeleccionarBloqueadasPorMedicoYFecha(IdMedico, daFecha)
    If oCita.MensajeError <> "" Then
        MsgBox oCita.MensajeError, vbInformation, "Asignación de Citas"
        Exit Function
    End If
       
    Do While Not rsCitas.EOF
       Set doCita = New DOCitaBloqueada
       doCita.idCitaBloqueada = IIf(IsNull(rsCitas!idCitaBloqueada), 0, rsCitas!idCitaBloqueada)
       doCita.fecha = IIf(IsNull(rsCitas!fecha), 0, rsCitas!fecha)
       doCita.HoraInicio = IIf(IsNull(rsCitas!HoraInicio), "", rsCitas!HoraInicio)
       doCita.horaFin = IIf(IsNull(rsCitas!horaFin), "", rsCitas!horaFin)
       doCita.IdMedico = IIf(IsNull(rsCitas!IdMedico), 0, rsCitas!IdMedico)
       doCita.idUsuario = IIf(IsNull(rsCitas!idUsuario), 0, rsCitas!idUsuario)
       doCita.FechaBloqueo = IIf(IsNull(rsCitas!FechaBloqueo), 0, rsCitas!FechaBloqueo)
       doCita.HoraBloqueo = IIf(IsNull(rsCitas!HoraBloqueo), 0, rsCitas!HoraBloqueo)
       doCitas.Add doCita
       rsCitas.MoveNext
    Loop
    
    Set CitasLeerBloqueadasPorMedicoYFecha = doCitas

End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function CitasSeleccionarDisponiblesPorMedicoYFecha(IdMedico As Long, daFecha As Date) As Collection
Dim oConexion As New ADODB.Connection
Dim oCita As New Citas
Dim rsCitas As ADODB.Recordset
Dim doCitas As New Collection
Dim doCita As doCita
Dim lHoraInicio As Long
Dim lHoraFin  As Long
Dim lTiempoPromedio As Long
Dim lHoraSiguiente As Long

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    Set oCita.Conexion = oConexion
    Set rsCitas = oCita.SeleccionarPorDisponiblesPorMedicoYFecha(IdMedico, daFecha)
    oConexion.Close
    If oCita.MensajeError <> "" Then
        MsgBox oCita.MensajeError, vbInformation, "Asignación de Citas"
        Exit Function
    End If
    
    Do While Not rsCitas.EOF
        lHoraInicio = ConvertirAMinutos(rsCitas!HoraInicio)
        lHoraFin = ConvertirAMinutos(rsCitas!horaFin)
        lTiempoPromedio = rsCitas!TiempoPromedioConsulta
        lHoraSiguiente = lHoraInicio
        
        Do While lHoraSiguiente < lHoraFin
            lHoraSiguiente = lHoraSiguiente + lTiempoPromedio
            
            Set doCita = New doCita
            doCita.HoraInicio = ConvertirAHora(lHoraInicio)
            doCita.horaFin = ConvertirAHora(lHoraSiguiente)
            doCita.idProgramacion = rsCitas!idProgramacion
            doCitas.Add doCita
                        
            lHoraInicio = lHoraSiguiente
        Loop
        rsCitas.MoveNext
    Loop
    
    Set CitasSeleccionarDisponiblesPorMedicoYFecha = doCitas

End Function

Function CitasSeleccionarDisponiblesPorMedicoEspecialidadYFecha(IdMedico As Long, IdEspecialidad As Long, daFecha As Date, oConexion As Connection) As Collection
Dim oCita As New Citas
Dim rsCitas As ADODB.Recordset
Dim doCitas As New Collection
Dim doCita As doCita
Dim lHoraInicio As Long
Dim lHoraFin  As Long
Dim lTiempoPromedio As Long
Dim lHoraSiguiente As Long


    Set oCita.Conexion = oConexion
    Set rsCitas = oCita.SeleccionarPorDisponiblesPorMedicoEspecialidadYFecha(IdMedico, IdEspecialidad, daFecha)
    If oCita.MensajeError <> "" Then
        MsgBox oCita.MensajeError, vbInformation, "Asignación de Citas"
        Exit Function
    End If
    
    Do While Not rsCitas.EOF
        lHoraInicio = ConvertirAMinutos(rsCitas!HoraInicio)
        lHoraFin = ConvertirAMinutos(rsCitas!horaFin)
        lTiempoPromedio = rsCitas!TiempoPromedioAtencion
        lHoraSiguiente = lHoraInicio
        
        Do While lHoraSiguiente < lHoraFin
            lHoraSiguiente = lHoraSiguiente + lTiempoPromedio
            
            Set doCita = New doCita
            doCita.HoraInicio = ConvertirAHora(lHoraInicio)
            doCita.horaFin = ConvertirAHora(lHoraSiguiente)
            doCita.idProgramacion = rsCitas!idProgramacion
            doCitas.Add doCita
                        
            lHoraInicio = lHoraSiguiente
        Loop
        rsCitas.MoveNext
    Loop
    
    Set CitasSeleccionarDisponiblesPorMedicoEspecialidadYFecha = doCitas

End Function

'debb-23/04/2018
Function CitasSeleccionarDisponiblesPorMedicoEspecialidadYFechaHoras(IdMedico As Long, IdEspecialidad As Long, daFecha As Date, _
                                                                     oConexion As Connection, lcHoraInicio As String, _
                                                                     lcHoraFinal As String, oRsConsultoriosAsignados As Recordset, _
                                                                     wxParametro522 As String) As Collection
Dim oCita As New Citas
Dim rsCitas As ADODB.Recordset
Dim doCitas As New Collection
Dim doCita As doCita
Dim lHoraInicio As Long
Dim lHoraFin  As Long
Dim lTiempoPromedio As Long
Dim lHoraSiguiente As Long

    Set CitasSeleccionarDisponiblesPorMedicoEspecialidadYFechaHoras = doCitas
    Set oCita.Conexion = oConexion
    Set rsCitas = oCita.SeleccionarPorDisponiblesPorMedicoEspecialidadYFecha(IdMedico, IdEspecialidad, daFecha)
    If oCita.MensajeError <> "" Then
        MsgBox oCita.MensajeError, vbInformation, "Asignación de Citas"
        Exit Function
    End If
    If rsCitas.RecordCount = 0 Then
       Exit Function
    End If
    If wxParametro522 = "S" Then
         If oRsConsultoriosAsignados.RecordCount = 0 Then
            Exit Function
         Else
            oRsConsultoriosAsignados.MoveFirst
            oRsConsultoriosAsignados.Find "idServicio=" & rsCitas!IdServicio
            If oRsConsultoriosAsignados.EOF Then
               Exit Function
            End If
         End If
    End If
    Do While Not rsCitas.EOF

       If rsCitas!HoraInicio >= lcHoraInicio And rsCitas!HoraInicio < lcHoraFinal Then
            lHoraInicio = ConvertirAMinutos(rsCitas!HoraInicio)
            lHoraFin = ConvertirAMinutos(rsCitas!horaFin)
'            lTiempoPromedio = rsCitas!TiempoPromedioAtencion
            lTiempoPromedio = IIf(IsNull(rsCitas!TiempoPromedioAtencion), rsCitas!tpaEspecialidadCE, rsCitas!TiempoPromedioAtencion)
            lHoraSiguiente = lHoraInicio
            Do While lHoraSiguiente < lHoraFin
                lHoraSiguiente = lHoraSiguiente + lTiempoPromedio
                
                Set doCita = New doCita
                doCita.HoraInicio = ConvertirAHora(lHoraInicio)
                doCita.horaFin = ConvertirAHora(lHoraSiguiente)
                doCita.idProgramacion = rsCitas!idProgramacion
                doCitas.Add doCita
                            
                lHoraInicio = lHoraSiguiente
            Loop
        End If
        rsCitas.MoveNext
    Loop
    
    Set CitasSeleccionarDisponiblesPorMedicoEspecialidadYFechaHoras = doCitas

End Function

Function ConvertirAMinutos(sHora As String) As Long
Dim sHoras() As String
        
        sHoras = Split(sHora, ":")
        ConvertirAMinutos = Val(sHoras(0)) * 60 + Val(sHoras(1))
        
End Function



Function CalculaHoraFinal(lcHoraInicial As String, lTiempoPromedio As Long) As String
    Dim lHoraInicio As Long
    Dim lHoraSiguiente As Long
    lHoraInicio = ConvertirAMinutos(lcHoraInicial)
    lHoraSiguiente = lHoraInicio + lTiempoPromedio
    CalculaHoraFinal = ConvertirAHora(lHoraSiguiente)
End Function

Function ConvertirAHora(lMinutos As Long) As String
Dim sHoras() As String
Dim lMin As Long
Dim lHora As Long
        
        lMin = lMinutos Mod 60
        lHora = (lMinutos - lMin) / 60
        
        ConvertirAHora = Right("00" & lHora, 2) + ":" + Right("00" & lMin, 2)
        
End Function

Function CitasSeleeccionarPacientePorMedicoFecha(IdMedico As Long, daFecha As Date, oConexion As Connection) As ADODB.Recordset
Dim oCita As New Citas
Dim rsPacientes As ADODB.Recordset
    
    Set oCita.Conexion = oConexion
    Set rsPacientes = oCita.SeleccionarPacientesPorMedicoYFecha(IdMedico, daFecha)
    If oCita.MensajeError <> "" Then
        MsgBox oCita.MensajeError, vbInformation, "Asignación de Citas"
        Exit Function
    End If
    
    Set CitasSeleeccionarPacientePorMedicoFecha = rsPacientes

End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function CitasSeleeccionarPacientePorMedicoFechaHoras(IdMedico As Long, daFecha As Date, oConexion As Connection, lcHoraInicial As String, lcHoraFinal As String) As ADODB.Recordset
Dim rsPacientes As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter


  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "CitasSeleeccionarPacientePorMedicoFechaHoras"
    Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IdMedico): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, daFecha): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@HoraInicial", adChar, adParamInput, 5, lcHoraInicial): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@HoraFinal", adChar, adParamInput, 5, lcHoraFinal): .Parameters.Append oParameter
    Set rsPacientes = .Execute
    Set rsPacientes.ActiveConnection = Nothing
  End With

    Set CitasSeleeccionarPacientePorMedicoFechaHoras = rsPacientes
    Set oCommand = Nothing
End Function

Function ProgramacionMedicaSeleccionarDiasDeCEPorMedicoYMes(IdMedico As Long, iMes As Integer, iAnio As Integer, oConexion As Connection) As ADODB.Recordset
Dim oProgramacion As New ProgramacionMedica
Dim rsPacientes As ADODB.Recordset
    
    Set oProgramacion.Conexion = oConexion
    Set rsPacientes = oProgramacion.SeleccionarDiasDeCEProgPorMedicoYMes(IdMedico, iMes, iAnio)
    If oProgramacion.MensajeError <> "" Then
        MsgBox oProgramacion.MensajeError, vbInformation, "Asignación de Citas"
        Exit Function
    End If
    
    Set ProgramacionMedicaSeleccionarDiasDeCEPorMedicoYMes = rsPacientes
    Set oProgramacion = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function ProgramacionMedicaSeleccionarPorId(IdProgramacionMedica As Long) As DOProgramacionMedica
Dim oConexion As New ADODB.Connection
Dim oProgramacionMedica As New ProgramacionMedica
Dim oDOProgramacionMedica As New DOProgramacionMedica
    
    Set ProgramacionMedicaSeleccionarPorId = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oProgramacionMedica.Conexion = oConexion
    oDOProgramacionMedica.idProgramacion = IdProgramacionMedica
    If oProgramacionMedica.SeleccionarPorId(oDOProgramacionMedica) Then
        Set ProgramacionMedicaSeleccionarPorId = oDOProgramacionMedica
    Else
        Set ProgramacionMedicaSeleccionarPorId = Nothing
    End If
    oConexion.Close

    ms_MensajeError = oProgramacionMedica.MensajeError
    Set oConexion = Nothing
    Set oProgramacionMedica = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function MedicosSeleccionarTodosOrdenadoAlfabeticamente() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
'Dim ssql As String
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
        Set MedicosSeleccionarTodosOrdenadoAlfabeticamente = Nothing
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion

        With oCommand
'           .CommandType = adCmdText
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "MedicosSeleccionarTodosOrdenadoAlfabeticamente"
'           .CommandText = ssql    'esta ok en sql2008r2
           Set oRecordset = .Execute
           Set oRecordset.ActiveConnection = Nothing
        End With
        Set MedicosSeleccionarTodosOrdenadoAlfabeticamente = oRecordset
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
        ms_MensajeError = ""
        Exit Function
ManejadorDeError:
        ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
        Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function DevuelveListaMedicosSegunServicio(lnIdServicioHospital As Long) As Recordset
    Dim lnEspecialidadServicio As Long
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    

    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "DevuelveListaMedicosSegunServicio"
      Set oParameter = .CreateParameter("@IdServicioHospital", adInteger, adParamInput, 0, lnIdServicioHospital): .Parameters.Append oParameter
      Set oRsTmp = .Execute
      Set oRsTmp.ActiveConnection = Nothing
    End With
    Set DevuelveListaMedicosSegunServicio = oRsTmp
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function MedicosSeleccionarPorIdMedicoPlanilla(lnIdMedico As Long) As Recordset
'        Dim lcSql As String
        Dim oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
  

        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "MedicosSeleccionarPorIdMedicoPlanilla"
          Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lnIdMedico): .Parameters.Append oParameter
          Set oRsTmp1 = .Execute
          Set oRsTmp1.ActiveConnection = Nothing
        End With
        
        Set MedicosSeleccionarPorIdMedicoPlanilla = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function ProgramacionMedicaSeleccionarPorMedicoFechaServicio(lnIdMedico As Long, lcFecha As String, lnIdServicio As Long) As Recordset
'        Dim lcSql As String
        Dim oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection


        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
          
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ProgramacionMedicaSeleccionarPorMedicoFechaServicio"
            Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lnIdMedico): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, Format(lcFecha, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
          
        Set ProgramacionMedicaSeleccionarPorMedicoFechaServicio = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Sub CitasActualizaDatosDeReprogramacionXfecha(lcFecha As String, lnIdAtencion As Long, lnIdProgramacionNew As Long, oConexion As Connection)
'        Dim lcSql As String, oRsTmp As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        
        On Error GoTo ErrActCitas

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CitasActualizaDatosDeReprogramacionXfecha"
            Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdProgramacionNew", adInteger, adParamInput, 0, lnIdProgramacionNew): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, Format(lcFecha, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
            .Execute
        End With
        Set oCommand = Nothing
        Exit Sub
ErrActCitas:
'        Set oRsTmp = Nothing
        MsgBox Err.Description
End Sub
'SCCQ 21/02/2020 Cambio7 Inicio
Sub CitasActualizaDatosDeReprogramacionXfechaServicioCE(lcFecha As String, lnIdAtencion As Long, lnIdProgramacionNew As Long, lnIdServicioCENew As Long, oConexion As Connection)
'        Dim lcSql As String, oRsTmp As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        
        On Error GoTo ErrActCitas

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CitasActualizaDatosDeReprogramacionXfecha"
            Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdProgramacionNew", adInteger, adParamInput, 0, lnIdProgramacionNew): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, Format(lcFecha, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdServicioCENew", adInteger, adParamInput, 0, lnIdServicioCENew): .Parameters.Append oParameter
            .Execute
        End With
        Set oCommand = Nothing
        Exit Sub
ErrActCitas:
'        Set oRsTmp = Nothing
        MsgBox Err.Description
End Sub
'SCCQ 21/02/2020 Cambio7 Fin

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Sub ProgramacionMedicaEliminarPorId(lnIdProgramacion As Long, oConexion As Connection)
    Dim lcSql As String
'    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProgramacionMedicaEliminarPorId"
        Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamInput, 0, lnIdProgramacion): .Parameters.Append oParameter
        .Execute
    End With
    Set oCommand = Nothing

End Sub

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Sub CitasActualizaDatosDeReporgramacionXmedico(lnIdResponsableNew As Long, lnIdProgramacionNew As Long, lnIdAtencion As Long, oConexion As Connection)
'        Dim lcSql As String, oRsTmp As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CitasActualizaDatosDeReporgramacionXmedico"
            Set oParameter = .CreateParameter("@IdResponsableNew", adInteger, adParamInput, 0, lnIdResponsableNew): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdProgramacionNew", adInteger, adParamInput, 0, lnIdProgramacionNew): .Parameters.Append oParameter
            .Execute
        End With
        Set oCommand = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function MedicosSeleccionarPorEspecialidad(lnIdEspecialidad As Long) As Recordset
'        Dim lcSql As String
        Dim oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
                

        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
  
        Set MedicosSeleccionarPorEspecialidad = Nothing
        
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "MedicosSeleccionarPorEspecialidad"
          Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, lnIdEspecialidad): .Parameters.Append oParameter
          Set oRsTmp1 = .Execute
          Set oRsTmp1.ActiveConnection = Nothing
        End With

        Set MedicosSeleccionarPorEspecialidad = oRsTmp1
        
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function CitasSeleccionarXfechaMedico(lcFechaCita As String, lnIdMedico As Long) As Recordset
'        Dim lcSql As String
        Dim oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        
        Set CitasSeleccionarXfechaMedico = Nothing
        

        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CitasSeleccionarXfechaMedico"
            Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lnIdMedico): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, Format(lcFechaCita, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With

        Set CitasSeleccionarXfechaMedico = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function ProgramacionMedicaSeleccionarXFechaMedico(lcFechaProgramada As String, lnIdMedico As Long) As Recordset
        Dim oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ProgramacionMedicaSeleccionarXFechaMedico"
            Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lnIdMedico): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaProgramada", adDBTimeStamp, adParamInput, 0, Format(lcFechaProgramada, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
        
        Set ProgramacionMedicaSeleccionarXFechaMedico = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

'debb-20/10/2015
Sub CitasActualizaDatosXpaciente(oConexion As Connection, mo_lcNombrePc As String, lcMedicoActual As String, _
                                 lcServicioActual As String, oDOCita As doCita, lnIdCuentaAtencion As Long, _
                                 lcPaciente As String)
           On Error GoTo ErrorProceso1
           Dim oCitas As New Citas
'           Dim oRsTmp As New Recordset
'           Dim lcSql As String
           Dim oCommand As New ADODB.Command
           Dim oParameter As ADODB.Parameter
           Set oCitas.Conexion = oConexion
           
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "CitasActualizaDatosXpaciente"
                Set oParameter = .CreateParameter("@HoraInicio", adChar, adParamInput, 5, oDOCita.HoraInicio): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@FechaIngreso", adDBTimeStamp, adParamInput, 0, oDOCita.fecha): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdMedicoIngreso", adInteger, adParamInput, 0, oDOCita.IdMedico): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdServicioIngreso", adInteger, adParamInput, 0, oDOCita.IdServicio): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oDOCita.idAtencion): .Parameters.Append oParameter
                .Execute
            End With
            
           'Añade Cita nueva
           If Not oCitas.Insertar(oDOCita) Then
              MsgBox oCitas.MensajeError: GoTo ErrorProceso1
           End If
           Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
           Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCita.IdUsuarioAuditoria, "M", _
                                                     oDOCita.idProgramacion, "ProgramacionMedica", _
                                                     oConexion, 500 + 128, mo_lcNombrePc, _
                                                     Left("Cuenta:" & Trim(Str(lnIdCuentaAtencion)) & _
                                                     " " & lcPaciente, 100))
           Set oCitas = Nothing
'           Set oRsTmp = Nothing
           Set oCommand = Nothing
  
           Exit Sub
ErrorProceso1:
           MsgBox Err.Description
           Set oCitas = Nothing
'           Set oRsTmp = Nothing
           Set oCommand = Nothing
           Set mo_ReglasSeguridad = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function ProgramacionMedicaSeleccionarPorIdServicio(lnIdServicio As Long) As Recordset
'        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        Set ProgramacionMedicaSeleccionarPorIdServicio = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ProgramacionMedicaSeleccionarPorIdServicio"
            Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With

        Set ProgramacionMedicaSeleccionarPorIdServicio = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

Function MedicosSeleccionarXIdMedico(lnIdMedico As Long) As Recordset
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        Set MedicosSeleccionarXIdMedico = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "MedicosSeleccionarXIdMedico"
            Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lnIdMedico): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
        

        Set MedicosSeleccionarXIdMedico = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

Function MedicosDevuelveNombre(lnIdMedico As Long, Optional oConexion1 As Connection) As String
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        Dim mo_reglasComunes As New ReglasComunes
        MedicosDevuelveNombre = ""
        If oConexion1 Is Nothing Then
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
        End If
        With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "MedicosSeleccionarXidentificador"
            Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lnIdMedico): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            If oConexion1 Is Nothing Then
               Set oRsTmp1.ActiveConnection = Nothing
            End If
        End With
        If oRsTmp1.RecordCount > 0 Then
            MedicosDevuelveNombre = Trim(oRsTmp1!ApellidoPaterno) & " " & Trim(oRsTmp1!ApellidoMaterno) & " " & Trim(oRsTmp1!Nombres)
        End If
        If oConexion1 Is Nothing Then
           oConexion.Close
        End If
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function ProgramacionMedicaSeleccionarPorFechas(ldFechaInicial As Date, ldFechaFinal As Date) As Recordset
        '        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        Set ProgramacionMedicaSeleccionarPorFechas = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ProgramacionMedicaSeleccionarPorFechas"
            Set oParameter = .CreateParameter("@FechaInicial", adDBTimeStamp, adParamInput, 0, Format(ldFechaInicial, "dd/mm/yyyy")): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaFinal", adDBTimeStamp, adParamInput, 0, Format(ldFechaFinal, "dd/mm/yyyy")): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set ProgramacionMedicaSeleccionarPorFechas = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function CitasSeleccionarPorServicioYfecha(lnIdServicio As Long, ldFecha As Date) As Recordset
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CitasSeleccionarPorServicioYfecha"
            Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, Format(ldFecha, "dd/mm/yyyy")): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
        
        Set CitasSeleccionarPorServicioYfecha = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

Function CitasWebCuposEliminarPorFechas(ldFechaInicial As Date, ldFechaFinal As Date, oConexion As Connection, lnIdServicio As Long) As Boolean
        On Error GoTo ErrCWE
        CitasWebCuposEliminarPorFechas = False
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "CitasWebCuposEliminarPorFechas"
    Set oParameter = .CreateParameter("@FechaInicial", adDBTimeStamp, adParamInput, 0, ldFechaInicial): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaFinal", adDBTimeStamp, adParamInput, 0, ldFechaFinal): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  Set oCommand = Nothing
    
  CitasWebCuposEliminarPorFechas = True
  Exit Function
ErrCWE:
       ms_MensajeError = Err.Description
End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function CitasBloqueadasEliminarPorFechas(ldFechaInicial As Date, ldFechaFinal As Date, oConexion As Connection) As Boolean
        On Error GoTo ErrCWEB
        CitasBloqueadasEliminarPorFechas = False
        '        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CitasBloqueadasEliminarPorFechas"
            Set oParameter = .CreateParameter("@FechaInicial", adDBTimeStamp, adParamInput, 0, Format(ldFechaInicial, "dd/mm/yyyy")): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaFinal", adDBTimeStamp, adParamInput, 0, Format(ldFechaFinal, "dd/mm/yyyy")): .Parameters.Append oParameter
            .Execute
        End With
        Set oCommand = Nothing
        CitasBloqueadasEliminarPorFechas = True
        Exit Function
ErrCWEB:
       ms_MensajeError = Err.Description
End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function CitasWebCuposSeleccionarPorFechas(ldFechaInicial As Date, ldFechaFinal As Date, _
                                           Optional lnIdMedico As Long, _
                                           Optional lnIdServicio As Long, _
                                           Optional oConexionExterna As Connection, _
                                           Optional lnIdEstadoCitaWeb As Long) As Recordset
        On Error GoTo ErrCWE
'        Dim lcSql As String
        Dim oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New Connection
        Dim lcSql As String
        Set CitasWebCuposSeleccionarPorFechas = Nothing
        If oConexionExterna Is Nothing Then
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        End If
        lcSql = ""
        If lnIdMedico > 0 Then
           lcSql = lcSql & " and idMedico=" & lnIdMedico
        End If
        If lnIdServicio > 0 Then
           lcSql = lcSql & " and idServicio=" & lnIdServicio
        End If
        If lnIdEstadoCitaWeb > 0 Then
           lcSql = lcSql & " and idEstadoCitaWeb=" & lnIdEstadoCitaWeb
        End If
        With oCommand
          .CommandType = adCmdStoredProc
            If oConexionExterna Is Nothing Then
                Set .ActiveConnection = oConexion
            Else
                Set .ActiveConnection = oConexionExterna
            End If
          .CommandTimeout = 150
          .CommandText = "CitasWebCuposSeleccionarPorFechas"
          Set oParameter = .CreateParameter("@FechaInicial", adDBTimeStamp, adParamInput, 0, ldFechaInicial): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@FechaFinal", adDBTimeStamp, adParamInput, 0, ldFechaFinal): .Parameters.Append oParameter
          Set oRsTmp1 = .Execute
          If oConexionExterna Is Nothing Then
             Set oRsTmp1.ActiveConnection = Nothing
          End If
        End With
        If lcSql <> "" Then
           lcSql = Mid(lcSql, 5)
           oRsTmp1.Filter = lcSql
        End If
        Set CitasWebCuposSeleccionarPorFechas = oRsTmp1
        Set oCommand = Nothing
        Set oConexion = Nothing
        Exit Function
ErrCWE:
       ms_MensajeError = Err.Description
End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function MedicosSeleccionarTodosParaWeb() As Recordset
        '        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oConexion As New ADODB.Connection
        
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
  
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "MedicosSeleccionarTodosParaWeb"
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
                
        Set MedicosSeleccionarTodosParaWeb = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function CitasWebCuposSeleccionarPorIdCitaBloqueada(lnIdCitaBloqueada As Long, Optional oConexion1 As Connection) As Recordset
        On Error GoTo ErrCWE
        Dim oRsTmp1 As New Recordset
        Dim oConexion As New Connection
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        If oConexion1 Is Nothing Then
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        End If
        With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "CitasWebCuposSeleccionarPorIdCitaBloqueada"
            Set oParameter = .CreateParameter("@IdCitaBloqueada", adInteger, adParamInput, 0, lnIdCitaBloqueada): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            If oConexion1 Is Nothing Then
               Set oRsTmp1.ActiveConnection = Nothing
            End If
        End With
        Set CitasWebCuposSeleccionarPorIdCitaBloqueada = oRsTmp1
        If oConexion1 Is Nothing Then
           oConexion.Close
        End If
        Set oConexion = Nothing
        Set oCommand = Nothing
        Exit Function
ErrCWE:
       ms_MensajeError = Err.Description
End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function CitasBloqueadasEliminarPorId(lnIdCitaBloqueada As Long, Optional oConexion1 As Connection) As Boolean
        On Error GoTo ErrCWEB
        CitasBloqueadasEliminarPorId = False
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection

        If oConexion1 Is Nothing Then
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
        End If
        With oCommand
          .CommandType = adCmdStoredProc
          If oConexion1 Is Nothing Then
             Set .ActiveConnection = oConexion
          Else
             Set .ActiveConnection = oConexion1
          End If
          .CommandTimeout = 150
          .CommandText = "CitasBloqueadasEliminarPorId"
          Set oParameter = .CreateParameter("@IdCitaBloqueada", adInteger, adParamInput, 0, lnIdCitaBloqueada): .Parameters.Append oParameter
          .Execute
        End With

        CitasBloqueadasEliminarPorId = True
        If oConexion1 Is Nothing Then
            oConexion.Close
        End If
        Set oConexion = Nothing
        Set oCommand = Nothing
        Exit Function
ErrCWEB:
       ms_MensajeError = Err.Description
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function MedicosSeleccionarPorDNI(lcDNI As String) As Recordset

        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        Set MedicosSeleccionarPorDNI = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "MedicosSeleccionarPorDNI"
          Set oParameter = .CreateParameter("@Dni", adChar, adParamInput, 20, lcDNI): .Parameters.Append oParameter
          Set oRsTmp1 = .Execute
          Set oRsTmp1.ActiveConnection = Nothing
        End With
        
        Set MedicosSeleccionarPorDNI = oRsTmp1
        
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 10/10/2013 - se cambio a store procedure
Function ProgramacionMedicaSeleccionarXidentificador(lnIdProgramacion As Long, Optional oConexion1 As Connection) As Recordset

        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New Connection
        Set ProgramacionMedicaSeleccionarXidentificador = Nothing
        
        If oConexion1 Is Nothing Then
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
        End If
        
        With oCommand
          .CommandType = adCmdStoredProc
          If oConexion1 Is Nothing Then
             Set .ActiveConnection = oConexion
          Else
             Set .ActiveConnection = oConexion1
          End If
          .CommandTimeout = 150
          .CommandText = "ProgramacionMedicaSeleccionarXidentificador"
          Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamInput, 0, lnIdProgramacion): .Parameters.Append oParameter
          Set oRsTmp1 = .Execute
        End With
        
    Set ProgramacionMedicaSeleccionarXidentificador = oRsTmp1
    Set oCommand = Nothing
    Set oConexion = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 11/10/2013 - se cambio a store procedure
Function TurnosSeleccionarPorIdentificador(lnIdTurno As Long, Optional oConexion1 As Connection) As Recordset
'    Dim lcSql As String
     Dim oRsTmp1 As New ADODB.Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Dim oConexion As New Connection

    If oConexion1 Is Nothing Then
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
    End If
    With oCommand
          .CommandType = adCmdStoredProc
          If oConexion1 Is Nothing Then
             Set .ActiveConnection = oConexion
          Else
             Set .ActiveConnection = oConexion1
          End If
          .CommandTimeout = 150
          .CommandText = "TurnosSeleccionarPorIdentificador"
          Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, lnIdTurno): .Parameters.Append oParameter
          Set oRsTmp1 = .Execute
    End With
    Set TurnosSeleccionarPorIdentificador = oRsTmp1
    Set oCommand = Nothing
    Set oConexion = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 11/10/2013 - se cambio a store procedure
Function MedicosSeleccionarXidentificador(lnIdMedico As Long, Optional oConexion1 As Connection) As Recordset
'    Dim lcSql As String
     Dim oRsTmp1 As New ADODB.Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Dim oConexion As New Connection

    If oConexion1 Is Nothing Then
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
    End If
    
    With oCommand
          .CommandType = adCmdStoredProc
          If oConexion1 Is Nothing Then
             Set .ActiveConnection = oConexion
          Else
             Set .ActiveConnection = oConexion1
          End If
          .CommandTimeout = 150
          .CommandText = "MedicosSeleccionarXidentificador"
          Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lnIdMedico): .Parameters.Append oParameter
          Set oRsTmp1 = .Execute
    End With
        
    Set MedicosSeleccionarXidentificador = oRsTmp1
    Set oCommand = Nothing
    Set oConexion = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 11/10/2013 - se cambio a store procedure
Function MedicosXidEmpleado(lnIdEmpleado As Long, oConexion As Connection) As Recordset
'    Dim lcSql As String
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    

    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "MedicosXidEmpleado"
            Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, lnIdEmpleado): .Parameters.Append oParameter
            Set oRs = .Execute
    End With
    Set MedicosXidEmpleado = oRs
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 11/10/2013 - se cambio a store procedure
Function ProgramacionMedicaXmedicoFechaHOra(lnIdMedico As Long, lcFecha As String, lcHora As String) As Recordset
'    Dim lcSql As String
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    

    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ProgramacionMedicaXmedicoFechaHOra"
            Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lnIdMedico): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, Format(lcFecha, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Hora", adChar, adParamInput, 5, lcHora): .Parameters.Append oParameter
            Set oRs = .Execute
            Set oRs.ActiveConnection = Nothing
    End With
        
    Set ProgramacionMedicaXmedicoFechaHOra = oRs
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 11/10/2013 - se cambio a store procedure
Function CuposProgramadosPorMedicosPorFechas(ldFechaIni As Date, ldFechaFin As Date) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand  As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
Dim sSql As String
    Set CuposProgramadosPorMedicosPorFechas = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
'        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
'        .CommandText = ssql
        .CommandText = "CuposProgramadosPorMedicosPorFechas"
        Set oParameter = .CreateParameter("@FechaInicial", adDBTimeStamp, adParamInput, 0, Format(ldFechaIni, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFinal", adDBTimeStamp, adParamInput, 0, Format(ldFechaFin, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CuposProgramadosPorMedicosPorFechas = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 11/10/2013 - se cambio a store procedure
Function DevuelveCantidadCitasYaAsignadas(ldFechaIni As Date, ldFechaFin As Date, IdMedico As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand  As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
Dim sSql As String
    Set DevuelveCantidadCitasYaAsignadas = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    

    With oCommand
        .CommandType = adCmdStoredProc
'        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
'        .CommandText = ssql
        .CommandText = "DevuelveCantidadCitasYaAsignadas"
        Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaInicial", adDBTimeStamp, adParamInput, 0, Format(ldFechaIni, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFinal", adDBTimeStamp, adParamInput, 0, Format(ldFechaFin, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set DevuelveCantidadCitasYaAsignadas = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 11/10/2013 - se cambio a store procedure
Function DevuelveCantidadDeCitasAdicionales(ldFechaIni As Date, ldFechaFin As Date, IdMedico As Long) As Integer
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand  As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
'Dim ssql As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  

    With oCommand
        .CommandType = adCmdStoredProc
'        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
'        .CommandText = ssql
        .CommandText = "DevuelveCantidadDeCitasAdicionales"
        Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaInicial", adDBTimeStamp, adParamInput, 0, Format(ldFechaIni, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFinal", adDBTimeStamp, adParamInput, 0, Format(ldFechaFin, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   DevuelveCantidadDeCitasAdicionales = oRecordset.RecordCount
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
Exit Function
End Function


           
Function CitasWebCuposSeleccionarPorIdEstadoCita(lnIdEstadoCitaWeb As Long, oConexion As Connection) As Recordset
        On Error GoTo ErrCWE
        Dim oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CitasWebCuposSeleccionarPorIdEstadoCita"
            Set oParameter = .CreateParameter("@IdEstadoCitaWeb", adInteger, adParamInput, 0, lnIdEstadoCitaWeb): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
        End With
        
        Set CitasWebCuposSeleccionarPorIdEstadoCita = oRsTmp1
        Set oCommand = Nothing
        Exit Function
ErrCWE:
       ms_MensajeError = Err.Description
End Function


'Sub ImportaCitasWebMINSA(ml_idUsuario As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String)
'     '
'     On Error GoTo ErrImpCWEB
'     Dim iFila As Long, lnIdWeb As Long
'     Dim lineas() As String, lnRegistros As Long, Registro As String, Fichero As Integer, Campos() As String, Columna As Single
'     Dim ldFecha As Date, lnIdServicio As Long, lnIdMedico As Long, lcHoraInicio As String
'     Dim lcHoraFinal As String, lnIdEstadoCitaWeb As Long, lnIdCitaBloqueada As Long, lcDNI As String
'     Dim lcApellidoPaterno As String, lcApellidoMaterno As String, lcPrimerNombre As String
'     Dim ldFechaConfirmacion As Date, lcHoraConfirmacion As String, lnNroRegistro As Long
'     Dim lcSegundoNombre As String, lnIdTipoSexo As Long, ldFechaNacimiento As Date, lnUbigeo As Long
'     Dim lnDiasMaximoConfirmacion As Integer, lnDiasQfaltaPcita As Integer, lnIdFuenteFinanciamiento As Long
'     Dim ldFechaHoy As Date, lcHoraHoy As String, lnHistoriaCreada As Long, lnIdAtencion As Long
'     Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
'     Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
'     Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
'     Dim oCitasWebCupos As New CitasWebCupos, oDOCitasWebCupos As New DOCitasWebCupos
'     Dim oConexionExterna As New Connection, oConexion As New Connection
'     Dim oRsTmp1 As New Recordset, lbImportaCitasWeb As Boolean
'     lbImportaCitasWeb = False
'     '
'     '
'     oConexion.CommandTimeout = 300
'     oConexion.CursorLocation = adUseClient
'     oConexion.Open sighentidades.CadenaConexion
'     '
'     oConexionExterna.CommandTimeout = 300
'     oConexionExterna.CursorLocation = adUseClient
'     oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
'     oConexionExterna.BeginTrans
'     Set oCitasWebCupos.Conexion = oConexionExterna
'     '
'     lnDiasMaximoConfirmacion = Val(lcBuscaParametro.SeleccionaFilaParametro(321))
'     ldFechaHoy = CDate(lcBuscaParametro.RetornaFechaServidorSQL): lcHoraHoy = lcBuscaParametro.RetornaHoraServidorSQL
'
'
'
'    '*************************************** Citas Web MINSA ***********************************
'    Set oRsTmp1 = mo_ReglasDeProgMedica.CitasWebCuposSeleccionarPorIdEstadoCita(sghCitaWebEstados.CupoConfirmadoEnCitaWeb, oConexionExterna)
'    If oRsTmp1.RecordCount > 0 Then
'       oRsTmp1.MoveFirst
'       Do While Not oRsTmp1.EOF
'            If oRsTmp1!idCitaBloqueada > 0 And ldFechaHoy <= oRsTmp1!fecha Then
'                If oRsTmp1!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoEnCitaWeb Then
'                     lnDiasQfaltaPcita = DateDiff("d", CDate(Format(ldFechaHoy, sighentidades.DevuelveFechaSoloFormato_DMY) & _
'                                                  " " & lcHoraHoy), _
'                                                 CDate(Format(oRsTmp1!fecha, sighentidades.DevuelveFechaSoloFormato_DMY) & _
'                                                 " " & oRsTmp1!HoraInicio))
'                     Set oDOCitasWebCupos = oCitasWebCupos.SeleccionarPorIdCitaBloqueada(oRsTmp1!idCitaBloqueada)
'                     If oDOCitasWebCupos.idCitaBloqueada > 0 Then
'                         If lnDiasQfaltaPcita < lnDiasMaximoConfirmacion Then
'                            'ANULA cupo WEB
'                            oDOCitasWebCupos.idEstadoCitaWeb = sghCitaWebEstados.CupoANULADO
'                            oDOCitasWebCupos.IdUsuarioAuditoria = 0
'                            If oCitasWebCupos.modificar(oDOCitasWebCupos) = False Then
'                               MsgBox oCitasWebCupos.MensajeError: GoTo ErrImpCWEB
'                            End If
'                            If mo_ReglasDeProgMedica.CitasBloqueadasEliminarPorId(oRsTmp1!idCitaBloqueada, oConexion) = False Then
'                               MsgBox mo_ReglasDeProgMedica.MensajeError: GoTo ErrImpCWEB
'                            End If
'                         Else
'                            'asigna al cupo reservado la cita web, incluyendo el paciente web
'                            '(si es necesario crea HISTORIA)
'                            If Not IsNull(oRsTmp1!IdWeb) Then
'                               oDOCitasWebCupos.IdWeb = oRsTmp1!IdWeb
'                            End If
'                            oDOCitasWebCupos.idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoYconCitaEnGalenhos
'                            oDOCitasWebCupos.DNI = oRsTmp1!DNI
'                            oDOCitasWebCupos.ApellidoPaterno = oRsTmp1!ApellidoPaterno
'                            oDOCitasWebCupos.ApellidoMaterno = oRsTmp1!ApellidoMaterno
'                            oDOCitasWebCupos.PrimerNombre = oRsTmp1!PrimerNombre
'                            oDOCitasWebCupos.SegundoNombre = IIf(IsNull(oRsTmp1!SegundoNombre), "", oRsTmp1!SegundoNombre)
'                            oDOCitasWebCupos.idTipoSexo = IIf(IsNull(oRsTmp1!idTipoSexo), 0, oRsTmp1!idTipoSexo)
'                            oDOCitasWebCupos.FechaNacimiento = IIf(IsNull(oRsTmp1!FechaNacimiento), 0, oRsTmp1!FechaNacimiento)
'                            If Not IsNull(oRsTmp1!Ubigeo) Then
'                               oDOCitasWebCupos.Ubigeo = oRsTmp1!Ubigeo
'                            End If
'                            If Not IsNull(oRsTmp1!FechaConfirmacion) Then
'                               oDOCitasWebCupos.FechaConfirmacion = oRsTmp1!FechaConfirmacion
'                            End If
'                            If Not IsNull(oRsTmp1!HoraConfirmacion) Then
'                               oDOCitasWebCupos.HoraConfirmacion = oRsTmp1!HoraConfirmacion
'                            End If
'                            oDOCitasWebCupos.idFuenteFinanciamiento = IIf(IsNull(oRsTmp1!idFuenteFinanciamiento), 0, oRsTmp1!idFuenteFinanciamiento)
'                            If Not IsNull(oRsTmp1!IdTurno) Then
'                               oDOCitasWebCupos.IdTurno = oRsTmp1!IdTurno
'                            End If
'                            If Not IsNull(oRsTmp1!email) Then
'                               oDOCitasWebCupos.email = oRsTmp1!email
'                            End If
'                            If Not IsNull(oRsTmp1!TELEFONO) Then
'                               oDOCitasWebCupos.TELEFONO = oRsTmp1!TELEFONO
'                            End If
'                            If oCitasWebCupos.modificar(oDOCitasWebCupos) = False Then
'                               MsgBox oCitasWebCupos.MensajeError: GoTo ErrImpCWEB
'                            End If
'                            If CrearCitaAutomatica(oRsTmp1!DNI, oRsTmp1!fecha, oRsTmp1!HoraInicio, oRsTmp1!horaFinal, _
'                                                   oRsTmp1!IdServicio, oRsTmp1!IdMedico, oRsTmp1!ApellidoMaterno, _
'                                                   oRsTmp1!ApellidoPaterno, oRsTmp1!FechaNacimiento, _
'                                                   IIf(IsNull(oRsTmp1!Ubigeo), 0, oRsTmp1!Ubigeo), _
'                                                   oRsTmp1!idTipoSexo, oRsTmp1!PrimerNombre, _
'                                                   IIf(IsNull(oRsTmp1!SegundoNombre), "", oRsTmp1!SegundoNombre), _
'                                                   oRsTmp1!idFuenteFinanciamiento, ldFechaHoy, lcHoraHoy, _
'                                                   lnHistoriaCreada, oConexion, ml_idUsuario, mo_lnIdTablaLISTBARITEMS, _
'                                                   mo_lcNombrePc, False, oDOCitasWebCupos.idPaciente, lnIdAtencion) = False Then
'                               MsgBox ms_MensajeError: GoTo ErrImpCWEB
'                            End If
'                            If mo_ReglasDeProgMedica.CitasBloqueadasEliminarPorId(oRsTmp1!idCitaBloqueada, oConexion) = False Then
'                               MsgBox mo_ReglasDeProgMedica.MensajeError: GoTo ErrImpCWEB
'                            End If
'                         End If
'                     End If
'                End If
'           End If
'           oRsTmp1.MoveNext
'       Loop
'    End If
'    '
'    oConexionExterna.CommitTrans
'    oConexionExterna.Close
'    oConexion.Close
'    '
'    lbImportaCitasWeb = True
'ErrImpCWEB:
'    If lbImportaCitasWeb = False Then
'        oConexionExterna.RollbackTrans
'        oConexionExterna.Close
'        oConexion.Close
'    End If
'    Set mo_ReglasDeProgMedica = Nothing
'    Set mo_ReglasAdmision = Nothing
'    Set mo_ReglasServiciosHosp = Nothing
'    Set oConexionExterna = Nothing
'    Set oConexion = Nothing
'    Set oRsTmp1 = Nothing
'    'Resume
'End Sub


Function CrearCitaAutomatica(lcDNI As String, ldFechaCita As Date, lcHoraInicio As String, lcHoraFinal As String, lnIdServicio As Long, _
                   lnIdMedico As Long, lcApellidoMaterno, lcApellidoPaterno As String, ldFechaNacimiento As Date, _
                   lnUbigeo As Long, lnIdTipoSexo As Long, lcPrimerNombre As String, lcSegundoNombre As String, _
                   lnIdFuenteFinanciamiento As Long, ldFechaHoy As Date, lcHoraHoy As String, _
                   ByRef lnHistoriaCreada As Long, oConexion As Connection, ml_idUsuario As Long, _
                   mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lbEsCitaAdicional As Boolean, _
                   lnIdPaciente As Long, ByRef lnIdAtencion As Long, _
                   lnIdEstablecimientoOrigen As Long, lnNroReferenciaOrigen As String, _
                   lnReferenciaMedicoOColeg As String, lnReferenciaOidDiagnostico As Long, _
                   lcReferenciaOservicio As String) As Boolean
     On Error GoTo ErrCrearCitaWeb
     ms_MensajeError = ""
     CrearCitaAutomatica = False
     Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
     Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
     Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
     Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
     Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
     Dim mo_CuentasAtencion As New DOCuentaAtencion, mo_Atenciones As New DOAtencion, mo_Pacientes As New DOPaciente
     Dim mo_DOFacturacionPaquetes As New DOFacturacionPaquetes, oDoSunasaPacientesHistoricos As New DoSunasaPacientesHistoricos
     Dim mo_Cita As New doCita, mo_Historia As New DOHistoriaClinica, mo_DoAtencionDatosAdicionales As New DoAtencionDatosAdicionales
     Dim mo_FacturacionServicios As New Collection
     Dim oPacientes As New Pacientes, oHistoria As New HistoriasClinicas
     Dim oRsTmp21 As New Recordset
     Dim lnTipoNumeracionAnterior As Long, lnIdHistoriaClinicaAnterior As Long, lbYaSeTransfirioHCdeUnServicioAotro As Boolean
     Dim mo_Diagnosticos As New Collection, mo_Procedimientos As New Collection, mo_Examenes As New Collection
     Dim lnIdtipoFinanciamiento As Long, lnIdProductoCpt As Long, lnIdEspecialidadMedico As Long, lnIdProgramacion As Long
     Dim oEdad As Edad, lbNecesitaCrearHistoria As Boolean
     Dim lnIdCondicionAlServicio As Long, lnIdCondicionAlEstablecimiento As Long
     '
     
     '
     Set oRsTmp21 = mo_ReglasFacturacion.TiposFinanciamientosTarifaSeleccionarPorPlan(lnIdFuenteFinanciamiento)
     lnIdtipoFinanciamiento = 0
     If oRsTmp21.RecordCount > 0 Then
        lnIdtipoFinanciamiento = oRsTmp21.Fields!IdTipoFinanciamiento
     End If
     oRsTmp21.Close
     '
     lnIdProductoCpt = 0
     Set oRsTmp21 = mo_ReglasServiciosHosp.EspecialidadCESeleccionarPorIdServicio(lnIdServicio)
     If oRsTmp21.RecordCount > 0 Then
        lnIdProductoCpt = oRsTmp21.Fields!IdProductoConsulta
     End If
     oRsTmp21.Close
     '
     lnIdEspecialidadMedico = 0: lnIdProgramacion = 0
     Set oRsTmp21 = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarPorMedicoFechaServicio(lnIdMedico, Format(ldFechaCita, "dd/mm/yyyy"), lnIdServicio)
     If oRsTmp21.RecordCount > 0 Then
        oRsTmp21.MoveFirst
        Do While Not oRsTmp21.EOF
            If lcHoraInicio >= oRsTmp21!HoraInicio And lcHoraInicio <= oRsTmp21!horaFin Then
               lnIdEspecialidadMedico = oRsTmp21.Fields!IdEspecialidad: lnIdProgramacion = oRsTmp21.Fields!idProgramacion
               Exit Do
            End If
            oRsTmp21.MoveNext
        Loop
     End If
     oRsTmp21.Close
     If lnIdProgramacion = 0 Then   'hubo reprogramación medica, no se podrá crear cita
        GoTo ErrCrearCitaWeb
     End If
     '
     If lnIdPaciente > 0 Then
        Set mo_Pacientes = mo_AdminAdmision.PacientesSeleccionarPorId(lnIdPaciente, oConexion)
     End If
     If mo_Pacientes.idPaciente = 0 Then
        Set mo_Pacientes = mo_AdminAdmision.PacientesSeleccionarPorDNI(lcDNI)
     End If
     If mo_Pacientes.idPaciente = 0 Then
        lbNecesitaCrearHistoria = True
        mo_Pacientes.ApellidoMaterno = UCase(lcApellidoMaterno)
        mo_Pacientes.ApellidoPaterno = UCase(lcApellidoPaterno)
        mo_Pacientes.FechaNacimiento = ldFechaNacimiento
        mo_Pacientes.IdDistritoDomicilio = lnUbigeo
        mo_Pacientes.idDocIdentidad = 1
        mo_Pacientes.idPaciente = 0
        mo_Pacientes.idTipoSexo = lnIdTipoSexo
        mo_Pacientes.IdUsuarioAuditoria = ml_idUsuario
        mo_Pacientes.NroDocumento = lcDNI
        mo_Pacientes.NroHistoriaClinica = 0
        mo_Pacientes.PrimerNombre = UCase(lcPrimerNombre)
        mo_Pacientes.SegundoNombre = lcSegundoNombre
        mo_Pacientes.IdTipoNumeracion = 1
        mo_Pacientes.autogenerado = mo_AdminAdmision.PacienteCrearNroAutogenerado(mo_Pacientes)
     End If
     oEdad = CalcularEdad(mo_Pacientes.FechaNacimiento, CDate(Format(ldFechaCita, "dd/mm/yyyy") & " " & lcHoraInicio))
     mo_reglasComunes.TiposCondicionPacienteCondicionAlEstablecimientoYservicio lnIdCondicionAlEstablecimiento, _
                              lnIdCondicionAlServicio, mo_Pacientes.idPaciente, _
                              Format(ldFechaCita, sighentidades.DevuelveFechaSoloFormato_DMY), 0, lnIdServicio
     If mo_Pacientes.idPaciente = 0 Then
        mo_Historia.FechaCreacion = ldFechaHoy
        mo_Historia.idEstadoHistoria = 1
        mo_Historia.idPaciente = 0
        mo_Historia.IdTipoNumeracion = 1
        mo_Historia.IdUsuarioAuditoria = ml_idUsuario
        mo_Historia.NroHistoriaClinica = 0
        mo_Historia.IdTipoHistoria = 1
     Else
        mo_Historia.NroHistoriaClinica = mo_Pacientes.NroHistoriaClinica
        Set oHistoria.Conexion = oConexion
        If oHistoria.SeleccionarPorId(mo_Historia) = False Then
           ms_MensajeError = oHistoria.MensajeError: GoTo ErrCrearCitaWeb
        End If
     End If
     '
     mo_CuentasAtencion.FechaApertura = ldFechaHoy
     mo_CuentasAtencion.FechaCreacion = ldFechaHoy
     mo_CuentasAtencion.HoraApertura = lcHoraHoy
     mo_CuentasAtencion.IdEstado = sghEstadoCuenta.sghAbierto
     mo_CuentasAtencion.idPaciente = mo_Pacientes.idPaciente
     mo_CuentasAtencion.IdUsuarioAuditoria = ml_idUsuario
     '
     mo_Atenciones.Edad = oEdad.Edad
     mo_Atenciones.EsPacienteExterno = False
     mo_Atenciones.FechaIngreso = ldFechaCita
     mo_Atenciones.HoraIngreso = lcHoraInicio
     mo_Atenciones.idAtencion = 0
     mo_Atenciones.idCuentaAtencion = 0
     mo_Atenciones.IdEspecialidadMedico = lnIdEspecialidadMedico
     mo_Atenciones.IdEstadoAtencion = sghEstadoTabla.sghRegistrado
     mo_Atenciones.IdFormaPago = lnIdtipoFinanciamiento
     mo_Atenciones.idFuenteFinanciamiento = lnIdFuenteFinanciamiento
     mo_Atenciones.idMedicoIngreso = lnIdMedico
     'mo_Atenciones.IdMedicoEgreso
     mo_Atenciones.IdOrigenAtencion = 10   'domicilio
     mo_Atenciones.idPaciente = mo_Pacientes.idPaciente
     mo_Atenciones.idServicioIngreso = lnIdServicio
     mo_Atenciones.IdTipoEdad = oEdad.TipoEdad
     mo_Atenciones.idTipoServicio = 1
     mo_Atenciones.IdUsuarioAuditoria = ml_idUsuario
     mo_Atenciones.idTipoCondicionAlEstab = lnIdCondicionAlEstablecimiento
     mo_Atenciones.idTipoCondicionAlServicio = lnIdCondicionAlServicio
     '
     mo_Cita.fecha = ldFechaCita
     mo_Cita.fechaSolicitud = ldFechaCita
     mo_Cita.horaFin = lcHoraFinal
     mo_Cita.HoraInicio = lcHoraInicio
     mo_Cita.horaSolicitud = lcHoraInicio
     mo_Cita.idAtencion = 0
     mo_Cita.IdCita = 0
     mo_Cita.IdEspecialidad = lnIdEspecialidadMedico
     mo_Cita.IdEstadoCita = 1
     mo_Cita.IdMedico = lnIdMedico
     mo_Cita.idPaciente = mo_Pacientes.idPaciente
     mo_Cita.idProducto = lnIdProductoCpt
     mo_Cita.idProgramacion = lnIdProgramacion
     mo_Cita.IdServicio = lnIdServicio
     mo_Cita.IdUsuarioAuditoria = ml_idUsuario
     If lbEsCitaAdicional = True Then
        mo_Cita.EsCitaAdicional = True
     End If
     '
     If lnIdEstablecimientoOrigen > 0 Then
        mo_Atenciones.IdOrigenAtencion = 12   'referido
        mo_DoAtencionDatosAdicionales.IdEstablecimientoOrigen = lnIdEstablecimientoOrigen
        mo_DoAtencionDatosAdicionales.IdTipoReferenciaOrigen = 1
        mo_DoAtencionDatosAdicionales.NroReferenciaOrigen = lnNroReferenciaOrigen
        mo_DoAtencionDatosAdicionales.ReferenciaMedicoOColeg = lnReferenciaMedicoOColeg
        mo_DoAtencionDatosAdicionales.referenciaOidDiagnostico = lnReferenciaOidDiagnostico
        mo_DoAtencionDatosAdicionales.referenciaOservicio = lcReferenciaOservicio
     End If
     If Len(Trim(mo_Pacientes.NroDocumento)) = 8 And mo_Atenciones.IdFormaPago = sghTipoFinanciamiento.sghSis Then
        Dim wxParametro302 As String, wxParametro322 As String, wxParametro323 As String
        wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
        wxParametro322 = lcBuscaParametro.SeleccionaFilaParametro(322)
        wxParametro323 = lcBuscaParametro.SeleccionaFilaParametro(323)
        If wxParametro322 = "S" And wxParametro302 = "S" Then
            Dim oRsBuscaPacientesSis As New Recordset
            Dim mo_SisConsumoWeb As New SIGHNegocios.SisConsumoWeb
            Set oRsBuscaPacientesSis = mo_SisConsumoWeb.WebServiceSISBuscarAfiliado(mo_Pacientes.NroDocumento, "", _
                                                "", "", "", "", wxParametro323)
            If oRsBuscaPacientesSis.RecordCount > 0 Then
               If (oRsBuscaPacientesSis!fbajaOK = "" Or IsNull(oRsBuscaPacientesSis!fbajaOK)) And _
                                                                            Val(oRsBuscaPacientesSis!estadoSis) = 0 Then
                    mo_DoAtencionDatosAdicionales.idSiaSis = oRsBuscaPacientesSis!idSiaSis
                    mo_DoAtencionDatosAdicionales.sisCodigo = oRsBuscaPacientesSis!Codigo
                    mo_DoAtencionDatosAdicionales.sisAfiliacion = oRsBuscaPacientesSis!cAfiliacion
               End If
            End If
            oRsBuscaPacientesSis.Close
            Set oRsBuscaPacientesSis = Nothing
            Set mo_SisConsumoWeb = Nothing
        End If
     End If
     
     'mo_Diagnosticos               'no hay datos
     'mo_Procedimientos             'no hay datos
     'mo_Examenes                   'no hay datos
     'mo_DOFacturacionPaquetes      'no hay datos
     'oDoSunasaPacientesHistoricos  'no hay datos
     '
     Set mo_FacturacionServicios = New Collection
     mo_FacturacionServicios.Add CargarTipoDeConsulta(lnIdtipoFinanciamiento, lnIdFuenteFinanciamiento, _
                                                     lnIdProductoCpt, ml_idUsuario)
     '
     lnTipoNumeracionAnterior = 1: lnIdHistoriaClinicaAnterior = 0: lbYaSeTransfirioHCdeUnServicioAotro = False
     
     If mo_AdminAdmision.AdmisionCEAgregar(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, mo_Cita, mo_Historia, _
                       lnTipoNumeracionAnterior, mo_Diagnosticos, mo_Procedimientos, mo_Examenes, _
                       lnIdHistoriaClinicaAnterior, mo_FacturacionServicios, mo_DOFacturacionPaquetes, _
                       lbYaSeTransfirioHCdeUnServicioAotro, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                       "Cita WEB: " & Trim(mo_Pacientes.ApellidoPaterno) & " " & Trim(mo_Pacientes.ApellidoMaterno) & _
                       " " & mo_Pacientes.PrimerNombre & "(" & mo_Atenciones.FechaIngreso & ")", _
                       oDoSunasaPacientesHistoricos, mo_DoAtencionDatosAdicionales) = False Then
        ms_MensajeError = mo_AdminAdmision.MensajeError: GoTo ErrCrearCitaWeb
     End If
     If lbNecesitaCrearHistoria = True Then
        lnHistoriaCreada = mo_Pacientes.NroHistoriaClinica
     Else
        lnHistoriaCreada = 0
     End If
     lnIdAtencion = mo_Atenciones.idAtencion
     
     CrearCitaAutomatica = True
ErrCrearCitaWeb:
     Set mo_AdminAdmision = Nothing
     Set mo_ReglasFacturacion = Nothing
     Set mo_ReglasServiciosHosp = Nothing
     Set mo_ReglasDeProgMedica = Nothing
     Set mo_reglasComunes = Nothing
     Set mo_CuentasAtencion = Nothing
     Set mo_Atenciones = Nothing
     Set mo_Pacientes = Nothing
     Set mo_DOFacturacionPaquetes = Nothing
     Set oDoSunasaPacientesHistoricos = Nothing
     Set mo_Cita = Nothing
     Set mo_Historia = Nothing
     Set mo_DoAtencionDatosAdicionales = Nothing
     Set mo_FacturacionServicios = Nothing
     Set oPacientes = Nothing
     Set oHistoria = Nothing
     Set oRsTmp21 = Nothing
End Function

Function CargarTipoDeConsulta(lnIdtipoFinanciamiento As Long, lnIdFuenteFinanciamiento As Long, _
                              lnIdProductoCpt As Long, ml_idUsuario As Long) As DOFacturacionServicios
Dim oConsulta As New DOFacturacionServicios
Dim oRsBuscaSeguro As New ADODB.Recordset
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim PrSeguro As Double
        With oConsulta
            .idAtencion = 0
            .idFacturacionServicio = 0
            .IdTipoFinanciamiento = lnIdtipoFinanciamiento
            .idFuenteFinanciamiento = lnIdFuenteFinanciamiento
            .cantidad = 1
            .idProducto = lnIdProductoCpt
            .IdUsuarioAuditoria = ml_idUsuario
            .IdEstadoFacturacion = sghEstadoFacturacion.sghRegistrado
            .FechaAutorizaPendiente = 0
            .FechaAutorizaSeguro = 0
            .IdCentroCosto = 0
            .IdUsuarioAutorizaPendiente = 0
            .IdUsuarioAutorizaSeguro = 0
            .PrecioUnitario = 0
            .totalPorPagar = 0
            .IdPuntoCarga = 6
            '********pone Seguros en forma automatica, sin necesidad de ir a SEGUROS-inicio
'            If mi_Opcion = sghModificar Then
'               .IdOrden = lnIdFactServicios
'            End If
            PrSeguro = 0
            If lnIdtipoFinanciamiento = 9 Then
               'Si es EXONERACIONES tomará el PRECIO de un Paciente Normal
               Set oRsBuscaSeguro = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(lnIdProductoCpt, 1)
            Else
               Set oRsBuscaSeguro = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(lnIdProductoCpt, lnIdtipoFinanciamiento)
            End If
            If oRsBuscaSeguro.RecordCount > 0 Then
               PrSeguro = oRsBuscaSeguro.Fields!PrecioUnitario
            End If
            '.idTipoFinanciamiento = lnIdTipoFinanciamiento
            .CantidadSIS = 0
            .PrecioSIS = 0
            .ImporteSIS = 0
            .CantidadSOAT = 0
            .PrecioSOAT = 0
            .ImporteSOAT = 0
            .importeEXO = 0
            .cantidadConv = 0
            .precConv = 0
            .ImporteConv = 0
            Select Case lnIdtipoFinanciamiento
            Case 1  'Contado
            Case 2  'SIS
                 If PrSeguro > 0 Then
                    .CantidadSIS = 1
                    .PrecioSIS = PrSeguro
                    .ImporteSIS = PrSeguro
                    .IdEstadoFacturacion = 10
                    .FechaAutorizaSeguro = Now
                 Else
                 End If
            Case 3  'Soat
                 If PrSeguro > 0 Then
                    .CantidadSOAT = 1
                    .PrecioSOAT = PrSeguro
                    .ImporteSOAT = PrSeguro
                    .IdEstadoFacturacion = 10
                    .FechaAutorizaSeguro = Now
                 Else
                    .IdTipoFinanciamiento = 1
                 End If
            Case 4  'Convenios
                 If PrSeguro > 0 Then
                    .cantidadConv = 1
                    .precConv = PrSeguro
                    .ImporteConv = PrSeguro
                    .IdEstadoFacturacion = 10
                    .PrecioUnitario = PrSeguro
                    .FechaAutorizaSeguro = Now
                 Else
                    .IdTipoFinanciamiento = 1
                 End If
            Case 9  'Exonerados
                .importeEXO = PrSeguro
                .IdEstadoFacturacion = 10
                .IdTipoFinanciamiento = 1
                .FechaAutorizaEXO = Now
            End Select
            '********pone Seguros en forma automatica, sin necesidad de ir a SEGUROS-fin
        End With
        Set CargarTipoDeConsulta = oConsulta
        oRsBuscaSeguro.Close
        Set oRsBuscaSeguro = Nothing
        Set mo_AdminFacturacion = Nothing
End Function


Function ProgramacionMedicaServiciosSeleccionarPorFechas(ldFechaInicial As Date, ldFechaFinal As Date) As Recordset
        '        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        Set ProgramacionMedicaServiciosSeleccionarPorFechas = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ProgramacionMedicaServiciosSeleccionarPorFechas"
            Set oParameter = .CreateParameter("@FechaInicial", adDBTimeStamp, adParamInput, 0, Format(ldFechaInicial, "dd/mm/yyyy")): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaFinal", adDBTimeStamp, adParamInput, 0, Format(ldFechaFinal, "dd/mm/yyyy")): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set ProgramacionMedicaServiciosSeleccionarPorFechas = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

Function CitasSeleccionarPorServicioTurnoFecha(ldFecha As Date, lnIdServicio As Long, lnIdTurno As Long, lcHora As String) As Recordset
        '        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        Set CitasSeleccionarPorServicioTurnoFecha = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CitasSeleccionarPorServicioTurnoFecha"
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, Format(ldFecha, "dd/mm/yyyy")): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, lnIdTurno): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Hora", adVarChar, adParamInput, 5, lcHora): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set CitasSeleccionarPorServicioTurnoFecha = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

Function CitasSeleccionarPorIdProgramacion(lnIdProgramacion As Long, oConexion As Connection) As Recordset
        '        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CitasSeleccionarPorIdProgramacion"
            Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdProgramacion): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set CitasSeleccionarPorIdProgramacion = oRsTmp1
        Set oCommand = Nothing
End Function


Function ProgramacionMedicaPorIdMedicoMesAnio(IdMedico As Long, iMes As Integer, iAnio As Integer, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProgramacionMedicaPorIdMedicoMesAnio"
        Set oParameter = .CreateParameter("@lIdMedico", adInteger, adParamInput, 0, IdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@iMes", adInteger, adParamInput, 0, iMes): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@iAnio", adInteger, adParamInput, 0, iAnio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ProgramacionMedicaPorIdMedicoMesAnio = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description

End Function


Function ProgramacionMedicaSeleccionarXfechaConsultorio(lcFechaProgramada As String, lnIdServicio As Long) As Recordset
        Dim oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ProgramacionMedicaSeleccionarXfechaConsultorio"
            Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaProgramada", adDBTimeStamp, adParamInput, 0, Format(lcFechaProgramada, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
        
        Set ProgramacionMedicaSeleccionarXfechaConsultorio = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

Function ProgramacionMedicaSeleccionarPorEspecialidad(lnIdEspecialidad As Long) As Recordset
        Dim oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ProgramacionMedicaSeleccionarPorEspecialidad"
            Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, lnIdEspecialidad): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
        
        Set ProgramacionMedicaSeleccionarPorEspecialidad = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

'MARIO 20/08/2014
Function CitasBloqueadasTodas(oConexion As Connection) As Recordset
        Dim oRsTmp As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CitasBloqueadasSeleccionarTodas"
            Set oRsTmp = .Execute
            Set oRsTmp.ActiveConnection = Nothing
        End With
        Set CitasBloqueadasTodas = oRsTmp
        Set oCommand = Nothing
End Function




Sub CitasWebActualizaMantCita(oDOCita As doCita, mi_opcion As sghOpciones)
   On Error GoTo ErrACWEB
   Dim oRsTmp1 As New Recordset
   Dim oRsTmp2 As New Recordset
   Dim oConexionExterna As New Connection
   Dim mo_ReglasAdmision As New ReglasAdmision
   Dim oDOCitasWebCupos As New DOCitasWebCupos, oCitasWebCupos As New CitasWebCupos
   Dim ldFechaCita As Date, lcHoraInicio As String, lcHoraFin As String, lnIdPaciente As Long
   oConexionExterna.CursorLocation = adUseClient
   oConexionExterna.CommandTimeout = 300
   oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
   Set oCitasWebCupos.Conexion = oConexionExterna
   oDOCitasWebCupos.fecha = oDOCita.fecha
   oDOCitasWebCupos.HoraInicio = oDOCita.HoraInicio
   oDOCitasWebCupos.idEstadoCitaWeb = 1
   oDOCitasWebCupos.IdMedico = oDOCita.IdMedico
   oDOCitasWebCupos.IdServicio = oDOCita.IdServicio
   If mi_opcion = sghEliminar Then
          oDOCitasWebCupos.ApellidoMaterno = ""
          oDOCitasWebCupos.ApellidoPaterno = ""
          oDOCitasWebCupos.DNI = ""
          oDOCitasWebCupos.email = ""
          oDOCitasWebCupos.FechaNacimiento = 0
          oDOCitasWebCupos.idFuenteFinanciamiento = 0
          oDOCitasWebCupos.idTipoSexo = 0
          oDOCitasWebCupos.IdUsuarioAuditoria = 0
          oDOCitasWebCupos.PrimerNombre = ""
          oDOCitasWebCupos.SegundoNombre = ""
          oDOCitasWebCupos.TELEFONO = ""
          oDOCitasWebCupos.Ubigeo = 0
          oDOCitasWebCupos.idPaciente = 0
   Else
          Set oRsTmp2 = mo_ReglasAdmision.AtencionesSeleccionarPorIdAtencion(oDOCita.idAtencion)
          If oRsTmp2.RecordCount > 0 Then
               oDOCitasWebCupos.ApellidoMaterno = oRsTmp2!ApellidoMaterno
               oDOCitasWebCupos.ApellidoPaterno = oRsTmp2!ApellidoPaterno
               If oRsTmp2!idDocIdentidad = 1 Then
                 oDOCitasWebCupos.DNI = oRsTmp2!NroDocumento
               End If
               oDOCitasWebCupos.email = IIf(IsNull(oRsTmp2!email), "", oRsTmp2!email)
               oDOCitasWebCupos.FechaNacimiento = oRsTmp2!FechaNacimiento
               oDOCitasWebCupos.idFuenteFinanciamiento = oRsTmp2!idFuenteFinanciamiento
               oDOCitasWebCupos.idTipoSexo = oRsTmp2!idTipoSexo
               oDOCitasWebCupos.IdUsuarioAuditoria = 0
               oDOCitasWebCupos.PrimerNombre = oRsTmp2!PrimerNombre
               oDOCitasWebCupos.SegundoNombre = IIf(IsNull(oRsTmp2!SegundoNombre), "", oRsTmp2!SegundoNombre)
               oDOCitasWebCupos.TELEFONO = IIf(IsNull(oRsTmp2!TELEFONO), "", oRsTmp2!TELEFONO)
               oDOCitasWebCupos.Ubigeo = IIf(IsNull(oRsTmp2!IdDistritoDomicilio), 0, oRsTmp2!IdDistritoDomicilio)
               oDOCitasWebCupos.idPaciente = oDOCita.idPaciente
          End If
   End If
   If oCitasWebCupos.ModificarDcitas(oDOCitasWebCupos) = True Then
   End If
   oConexionExterna.Close
ErrACWEB:
   Set oRsTmp1 = Nothing
   Set oConexionExterna = Nothing
   Set oRsTmp2 = Nothing
   Set mo_ReglasAdmision = Nothing
   Set oDOCitasWebCupos = Nothing
   Set oCitasWebCupos = Nothing
End Sub

'MARIO 12/09/2014
Function CitasSeleccionarPorAtenciones(lnIdServicio As Long, ldFecha As Date) As Recordset
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CitasPorAtencion"
            Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, Format(ldFecha, "dd/mm/yyyy")): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
        
        Set CitasSeleccionarPorAtenciones = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

'Actualizado 30102014 yamill palomino
Function ProgramacionMedicaSeleccionarPorMedicoEnOtrosServicio(oDOProgramacionMedica As DOProgramacionMedica) As Recordset
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        Set ProgramacionMedicaSeleccionarPorMedicoEnOtrosServicio = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ProgramacionMedicaSeleccionarPorMedicoEnOtrosServicio"
            Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, oDOProgramacionMedica.IdServicio): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, oDOProgramacionMedica.IdMedico): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oDOProgramacionMedica.fecha = 0, Null, oDOProgramacionMedica.fecha)): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With

        Set ProgramacionMedicaSeleccionarPorMedicoEnOtrosServicio = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function


'mgaray201411c
Function EspecialidadMedicoValidaEliminar(oDOMedicoEspecialidad As DOMedicoEspecialidad) As Boolean
On Error GoTo miError
    Dim returnValue As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMedicoEspecialidad As New MedicosEspecialidad
    Dim oRsResult As ADODB.Recordset
    Dim sErrorMessage As String

    
    returnValue = True
    sErrorMessage = ""
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oMedicoEspecialidad.Conexion = oConexion
    
    Set oRsResult = oMedicoEspecialidad.EspecialidadMedicoValidaEliminar(oDOMedicoEspecialidad)
    
    If Not (oRsResult Is Nothing) Then
        If oRsResult.RecordCount > 0 Then
            Select Case oRsResult.Fields!ErrorNumber
                Case 0:
                    returnValue = True
                Case 1
                    returnValue = False
                    sErrorMessage = "Tiene Programación registradas en la especialidad"
                Case 2
                    returnValue = False
                    sErrorMessage = "Tiene Admisiones(CE, Emergencia u Hospitalización) registradas en la especialidad"
                Case Else
                    returnValue = False
                    sErrorMessage = "Se ha registrado datos para lña especialidad"
            End Select
        End If
    End If
    
    If returnValue = False Then
        MsgBox "No se puede eliminar especialidad, " & sErrorMessage, vbInformation, "Registro de Médicos"
    End If
    
    EspecialidadMedicoValidaEliminar = returnValue
    
    oConexion.Close
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical, "Registro de Médicos"
    End If
    Set oConexion = Nothing
    Set oMedicoEspecialidad = Nothing
    
End Function

'mgaray201503
Function ProgramacionMedicaCitaMasProxima(IdMedico As Long, _
            lIdEspecialidad As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    'iMes As Integer, iAnio As Integer
    Set ProgramacionMedicaCitaMasProxima = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProgramacionMedicaCitaMasProxima"
        Set oParameter = .CreateParameter("@lIdMedico", adInteger, adParamInput, 0, IIf(IdMedico = 0, Null, IdMedico)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idEspecialidad", adInteger, adParamInput, 0, IIf(lIdEspecialidad = 0, Null, lIdEspecialidad)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, sghTipoServicio.sghConsultaExterna): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ProgramacionMedicaCitaMasProxima = oRecordset
   Set oCommand = Nothing
   oConexion.Close
    Set oConexion = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'debb-30/10/2015
Function MedicosSeleccionarXId(lnIdMedico As Long, Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
    If oConexion1 Is Nothing Then
       oConexion.CommandTimeout = 900
       oConexion.CursorLocation = adUseClient
       oConexion.Open sighentidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "MedicosSeleccionarPorId"
        Set oParameter = .CreateParameter("@lIdMedico", adInteger, adParamInput, 0, lnIdMedico): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set MedicosSeleccionarXId = oRecordset
   Set oCommand = Nothing
   If oConexion1 Is Nothing Then
      oConexion.Close
      Set oConexion = Nothing
      Set oRecordset = Nothing
   End If
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
'debb-15/07/2016
Function ProgramacionMedicaSeleccionarXfechas(ldFechaInicial As Date, ldFechaFinal As Date, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProgramacionMedicaSeleccionarXfechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 8, ldFechaInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 8, ldFechaFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set ProgramacionMedicaSeleccionarXfechas = oRecordset
   Set oCommand = Nothing

   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'debb-15/07/2016
Function MedicosFiltrarDatos(lcFiltro As String, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "MedicosPorFiltro"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 2000, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set MedicosFiltrarDatos = oRecordset
   Set oCommand = Nothing

   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



Function MedicoDevuelveColegiatura(lnIdEmpleado As Long) As String
On Error GoTo miError
    Dim oRsTmp As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
    Dim bUsuarioEsMedico As Boolean
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    MedicoDevuelveColegiatura = ""
    Set oRsTmp = oReglasDeProgMedica.MedicosXidEmpleado(lnIdEmpleado, oConexion)
    If oRsTmp.RecordCount > 0 Then
        If Not IsNull(oRsTmp!colegiatura) Then
           MedicoDevuelveColegiatura = oRsTmp!colegiatura
        End If
    End If
    oRsTmp.Close
    oConexion.Close
    Set oConexion = Nothing
    Set oRsTmp = Nothing
    Set oReglasDeProgMedica = Nothing
miError:
'    If Err Then
'        MsgBox Err.Number & " : " & Err.Description, vbExclamation, "Mensaje"
'    End If
End Function


Sub CitasBloqueadasWEBEliminarXFechas(ldFechaInicial As Date, ldFechaFinal As Date)
        On Error GoTo ErrCWEB
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CitasBloqueadasWEBEliminarXFechas"
            Set oParameter = .CreateParameter("@FechaInicial", adDBTimeStamp, adParamInput, 0, Format(ldFechaInicial, "dd/mm/yyyy")): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaFinal", adDBTimeStamp, adParamInput, 0, Format(ldFechaFinal, "dd/mm/yyyy")): .Parameters.Append oParameter
            .Execute
        End With
        Set oCommand = Nothing
        oConexion.Close
       
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        CitasWebCuposEliminarPorFechas ldFechaInicial, ldFechaFinal, oConexion, 0
        oConexion.Close
        
         Set oConexion = Nothing
        Exit Sub
ErrCWEB:
       ms_MensajeError = Err.Description
End Sub

