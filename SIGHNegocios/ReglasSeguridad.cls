VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasDeSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Reglas de Seguridad
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property

Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

Function RolesItemsSeleccionarGruposPorUsuarioSql2000(lIdUsuario As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New RolesItems
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set RolesItemsSeleccionarGruposPorUsuarioSql2000 = oTabla.SeleccionarGruposPorUsuario(lIdUsuario)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
    
End Function

Function RolesItemsSeleccionarItemsPorUsuarioYGrupoSql2000(lIdUsuario As Long, lIdGrupo As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New RolesItems
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set RolesItemsSeleccionarItemsPorUsuarioYGrupoSql2000 = oTabla.SeleccionarItemsPorUsuarioYGrupo(lIdUsuario, lIdGrupo)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function RolesItemsSeleccionarPorRol(lIdRol As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New RolesItems
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set RolesItemsSeleccionarPorRol = oTabla.SeleccionarPorRol(lIdRol)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function RolesItemsSeleccionarPermisosPorEmpleadoYListItem(lIdEmpleado As Long, lIdListItem As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New RolesItems
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set RolesItemsSeleccionarPermisosPorEmpleadoYListItem = oTabla.SeleccionarPermisosPorIdEmpleadoYIdListItem(lIdEmpleado, lIdListItem)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function ListItemsSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New ListbarItems
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set ListItemsSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function RolesPermisosSeleccionarPorRol(lIdRol As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New RolesPermisos
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set RolesPermisosSeleccionarPorRol = oTabla.SeleccionarPorRol(lIdRol)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function PermisosSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Permisos
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set PermisosSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


'-------------------------------------------------------------------------------------
'                                       Roles
'-------------------------------------------------------------------------------------
Function RolesAgregar(oDoRol As DORol, oRolesItem As Collection, oRolesPermiso As Collection, oRsReportes As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreRol As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oRoles As New Roles
Dim oRolesItems As New RolesItems
Dim oRolesPermisos As New RolesPermisos
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    RolesAgregar = False
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    oConexion.BeginTrans
    Set oRoles.Conexion = oConexion
    Set oRolesItems.Conexion = oConexion
    Set oRolesPermisos.Conexion = oConexion
        
    If oRoles.Insertar(oDoRol) Then
        '
        Call AuditoriaAgregarV(oDoRol.IdUsuarioAuditoria, "A", oDoRol.IdRol, "Roles", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreRol)           'ListBarItems.idListItem
        If oRolesItems.ActualizarRolesItems(oRolesItem, oDoRol.IdRol) Then
            If oRolesPermisos.ActualizarRolesPermisos(oRolesPermiso, oDoRol.IdRol) Then
                Dim oRsTmp As New Recordset

                With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = oConexion
                  .CommandTimeout = 150
                  .CommandText = "RolesReportesConsultar"
                  Set oRsTmp = .Execute
                  Set oRsTmp.ActiveConnection = Nothing
                End With
                
                If oRsReportes.RecordCount > 0 Then
                   oRsReportes.MoveFirst
                   Do While Not oRsReportes.EOF
                        With oCommand
                            .CommandType = adCmdStoredProc
                            Set .ActiveConnection = oConexion
                            .CommandTimeout = 150
                            .CommandText = "RolesReportesAgregar"
                            Set oParameter = .CreateParameter("@IdRol", adInteger, adParamInput, 0, oDoRol.IdRol): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@idReporte", adInteger, adParamInput, 0, oRsReportes.Fields!idReporte): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@tieneAcceso", adBoolean, adParamInput, 0, oRsReportes.Fields!tieneAcceso): .Parameters.Append oParameter
                           .Execute
                        End With
                        Set oCommand = Nothing
                        Set oParameter = Nothing
'                      oRsTmp.AddNew
'                      oRsTmp.Fields!IdRol = oDoRol.IdRol
'                      oRsTmp.Fields!idReporte = oRsReportes.Fields!idReporte
'                      oRsTmp.Fields!tieneAcceso = oRsReportes.Fields!tieneAcceso
'                      oRsTmp.Update
                      oRsReportes.MoveNext
                   Loop
                End If
                
                RolesAgregar = True
                oConexion.CommitTrans
            End If
        End If
    Else
        ms_MensajeError = oRoles.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oRoles = Nothing
    Set oRolesItems = Nothing
    Set oRolesPermisos = Nothing
    Set oCommand = Nothing
    
End Function

Function RolesModificar(oDoRol As DORol, oRolesItem As Collection, oRolesPermiso As Collection, oRsReportes As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreRol As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oRoles As New Roles
Dim oRolesItems As New RolesItems
Dim oRolesPermisos As New RolesPermisos
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
  
    RolesModificar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    oConexion.BeginTrans
    Set oRoles.Conexion = oConexion
    Set oRolesItems.Conexion = oConexion
    Set oRolesPermisos.Conexion = oConexion
    
    If oRoles.modificar(oDoRol) Then
        '
        Call AuditoriaAgregarV(oDoRol.IdUsuarioAuditoria, "M", oDoRol.IdRol, "Roles", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreRol)            'ListBarItems.idListItem
        If oRolesItems.ActualizarRolesItems(oRolesItem, oDoRol.IdRol) Then
            If oRolesPermisos.ActualizarRolesPermisos(oRolesPermiso, oDoRol.IdRol) Then
                '
                Dim oRsTmp As New Recordset

                With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = oConexion
                  .CommandTimeout = 150
                  .CommandText = "RolesReportesEliminarConsultar"
                  Set oParameter = .CreateParameter("@IdRol", adInteger, adParamInput, 0, oDoRol.IdRol): .Parameters.Append oParameter
                   .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
                With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = oConexion
                  .CommandTimeout = 150
                  .CommandText = "RolesReportesConsultar"
                  Set oRsTmp = .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
                
                If oRsReportes.RecordCount > 0 Then
                   oRsReportes.MoveFirst
                   Do While Not oRsReportes.EOF
                        With oCommand
                          .CommandType = adCmdStoredProc
                          Set .ActiveConnection = oConexion
                          .CommandTimeout = 150
                          .CommandText = "RolesReportesAgregar"
                          Set oParameter = .CreateParameter("@IdRol", adInteger, adParamInput, 0, oDoRol.IdRol): .Parameters.Append oParameter
                          Set oParameter = .CreateParameter("@idReporte", adInteger, adParamInput, 0, oRsReportes.Fields!idReporte): .Parameters.Append oParameter
                          Set oParameter = .CreateParameter("@tieneAcceso", adBoolean, adParamInput, 0, oRsReportes.Fields!tieneAcceso): .Parameters.Append oParameter
                           .Execute
                        End With
                        Set oCommand = Nothing
                        Set oParameter = Nothing
'                      oRsTmp.AddNew
'                      oRsTmp.Fields!IdRol = oDoRol.IdRol
'                      oRsTmp.Fields!idReporte = oRsReportes.Fields!idReporte
'                      oRsTmp.Fields!tieneAcceso = oRsReportes.Fields!tieneAcceso
'                      oRsTmp.Update
                      oRsReportes.MoveNext
                   Loop
                End If
                
                RolesModificar = True
                oConexion.CommitTrans
            End If
        End If
    Else
        ms_MensajeError = oRoles.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oRoles = Nothing
    Set oRolesItems = Nothing
    Set oRolesPermisos = Nothing
    
End Function

Function RolesEliminar(oDoRol As DORol, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreRol As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oRoles As New Roles
Dim oRolesItems As New RolesItems
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
  
    RolesEliminar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    oConexion.BeginTrans
    Set oRoles.Conexion = oConexion
    Set oRolesItems.Conexion = oConexion
     
    If oRolesItems.EliminarRolesItems(oDoRol.IdRol) Then
        Dim oRsTmp As New Recordset

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "RolesReportesRolesPermisosEliminar"
            Set oParameter = .CreateParameter("@IdRol", adInteger, adParamInput, 0, oDoRol.IdRol): .Parameters.Append oParameter
            .Execute
        End With
        Set oParameter = Nothing
        
        Set oRsTmp = Nothing
        If oRoles.Eliminar(oDoRol) Then
           RolesEliminar = True
           oConexion.CommitTrans
        Else
            ms_MensajeError = oRoles.MensajeError
            oConexion.RollbackTrans
        End If
    Else
        ms_MensajeError = oRoles.MensajeError
    End If
    '
    Call AuditoriaAgregarV(oDoRol.IdUsuarioAuditoria, "E", oDoRol.IdRol, "Roles", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreRol)           'ListBarItems.idListItem
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRoles = Nothing
    Set oRolesItems = Nothing
End Function

Function RolesSeleccionarPorId(IdRol As Long) As DORol
Dim oConexion As New ADODB.Connection
Dim oRoles As New Roles
Dim oDoRol As New DORol
    
    Set RolesSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oRoles.Conexion = oConexion
    oDoRol.IdRol = IdRol
    If oRoles.SeleccionarPorId(oDoRol) Then
        Set RolesSeleccionarPorId = oDoRol
    End If
    oConexion.Close

    ms_MensajeError = oRoles.MensajeError
    Set oConexion = Nothing
    Set oRoles = Nothing

End Function
Function RolesSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Roles
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set RolesSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function UsuariosRolesSeleccionarPorEmpleado(lIdEmpleado As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New UsuariosRoles
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set UsuariosRolesSeleccionarPorEmpleado = oTabla.SeleccionarPorEmpleado(lIdEmpleado)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function UsuariosRolesSeleccionarPermisosFacturacion(lIdEmpleado As Long) As PermisosFacturacion
Dim oConexion As New ADODB.Connection
Dim oTabla As New RolesPermisos
Dim rsPermisos As New Recordset
Dim oPermisosFacturacion As New PermisosFacturacion
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set rsPermisos = oTabla.SeleccionarPermisosFacturacionPorUsuario(lIdEmpleado)
    
    Do While Not rsPermisos.EOF
        Select Case rsPermisos!IdPermiso
        Case 100
            oPermisosFacturacion.AgregarServicios = True
        Case 101
            oPermisosFacturacion.AgregarBienesInsumos = True
        Case 102
            oPermisosFacturacion.AgregarExoneraciones = True
        Case 104
            oPermisosFacturacion.AutorizarPendientesDePago = True
        Case 105
            oPermisosFacturacion.AgregarPolizasSeguro = True
        Case 106
            oPermisosFacturacion.AutorizarSIS = True
        Case 107
            oPermisosFacturacion.AutorizarSOAT = True
        Case 108
            oPermisosFacturacion.AutorizarConvenios = True
        Case 109
            oPermisosFacturacion.AutorizarDevoluciones = True
        Case 110
            oPermisosFacturacion.AbrirCuentaAtencion = True
        Case 111
            oPermisosFacturacion.CerrarCuentaAtencion = True
        Case 301
            oPermisosFacturacion.ActualizaFechaDocumentoES = True
        End Select
        rsPermisos.MoveNext
    Loop
    
    Set UsuariosRolesSeleccionarPermisosFacturacion = oPermisosFacturacion
    
    rsPermisos.Close
    oConexion.Close
    Set oConexion = Nothing
    Set rsPermisos = Nothing
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oPermisosFacturacion = Nothing
End Function


'debb-19/03/2018
Function EmpleadosAutenticar(sUsuario As String) As ADODB.Recordset
'SCCQ 10/02/2020 Cambio6 Inicio
Dim contador As Integer
contador = 0
Inicio:
contador = contador + 1
'SCCQ 10/02/2020 Cambio6 Fin
Dim oConexion As New ADODB.Connection
Dim oTabla As New Empleados
Dim lnErrorEnOpen As Boolean

On Error GoTo ErrMens
    lnErrorEnOpen = False
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    lnErrorEnOpen = True
    oConexion.Open sighentidades.CadenaConexion
    lnErrorEnOpen = False
    'sighentidades.ConexionCopia = sighentidades.CadenaConexion
   
    Set oTabla.Conexion = oConexion
    Set EmpleadosAutenticar = oTabla.Autenticar(sUsuario)
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Exit Function
ErrMens:
   ' If lnErrorEnOpen = True Then
   '    sighentidades.CadenaConexion = sighentidades.ConexionCopia
   '    Resume
   ' Else
       '
   ' End If
     'SCCQ 10/02/2020 Cambio6 Inicio
        If contador <= 1 Then
            'Si se genera un error
            Dim oReg As New sighentidades.Configuracion
            oReg.Acceso = "Acceso1"
            '*********************
        Else
    'SCCQ 10/02/2020 Cambio6 Fin
            MsgBox "Error en cadena 'CadenaConexion' del regedit" & Chr(13) & Err.Number & Chr(13) & Err.Description
    'SCCQ 10/02/2020 Cambio6 Inicio
        End If
    GoTo Inicio
    'SCCQ 10/02/2020 Cambio6 Fin
End Function
Function EmpleadosAutenticarMaquina(sUsuario As String, nMaquina As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Empleados
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
   ' MsgBox "Abrio Conexion BD"
    Set oTabla.Conexion = oConexion
    Set EmpleadosAutenticarMaquina = oTabla.Autenticarmaquina(sUsuario, nMaquina)
    oConexion.Close
    Set oConexion = Nothing
   ' MsgBox "Autentico BD"
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function


Function ParametrosObtenerCadenaConexion() As String
Dim oConexion As New ADODB.Connection
Dim oTabla As New Parametros
   
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    ParametrosObtenerCadenaConexion = oTabla.ObtenerCadenaConexion()
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function
Function ParametrosObtenerCadenaConexionShape() As String
Dim oConexion As New ADODB.Connection
Dim oTabla As New Parametros
   
    oConexion.Open sighentidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    ParametrosObtenerCadenaConexionShape = oTabla.ObtenerCadenaConexionShape()
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function

Function UsuariosRolesSeleccionarPermisosTodos(lnIdEmpleado As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New RolesPermisos
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set UsuariosRolesSeleccionarPermisosTodos = oTabla.SeleccionarPermisosFacturacionPorUsuario(lnIdEmpleado)

    Set oTabla = Nothing
End Function


Sub AuditoriaAgregar(lnIdUsuario As Long, lcAccionAME As String, lnIdRegistroDeTabla As Long, lcTablaSql As String, wxCadenaConexion As String)
    Dim oConexion As New ADODB.Connection
    Dim oRsTmp As New ADODB.Recordset
    Dim sSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lcHora As String
    On Error GoTo ErrCierre
    oConexion.Open wxCadenaConexion
    oConexion.BeginTrans
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = oConexion
       .CommandText = "AuditoriaAgregar"
        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, lnIdUsuario): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Accion", adChar, adParamInput, 1, lcAccionAME): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdRegistro", adInteger, adParamInput, 0, lnIdRegistroDeTabla): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Tabla", adChar, adParamInput, 50, lcTablaSql): .Parameters.Append oParameter
       .Execute
    End With
    '
    oConexion.CommitTrans
    oConexion.Close
    Set oCommand = Nothing
    Set oConexion = Nothing
    Exit Sub
ErrCierre:
    oConexion.RollbackTrans
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Sub


Sub AuditoriaAgregarV(lnIdUsuario As Long, lcAccionAME As String, lnIdRegistroDeTabla As Long, lcTablaSql As String, wxConexion As Connection, lnIdListItem As Long, lcNombrePc As String, lcObservaciones As String)
    Dim oRsTmp As New ADODB.Recordset
    Dim sSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lcHora As String
    On Error GoTo ErrCierre
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = wxConexion
       .CommandTimeout = 150
       .CommandText = "AuditoriaAgregarV"
        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, lnIdUsuario): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Accion", adChar, adParamInput, 1, lcAccionAME): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdRegistro", adInteger, adParamInput, 0, lnIdRegistroDeTabla): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Tabla", adChar, adParamInput, 50, lcTablaSql): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idListItem", adInteger, adParamInput, 0, lnIdListItem): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@nombrePC", adChar, adParamInput, 30, lcNombrePc): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@observaciones", adChar, adParamInput, 100, Left(lcObservaciones & "<>" & sighentidades.ParaAuditoria, 100)): .Parameters.Append oParameter
       .Execute
    End With
    Set oCommand = Nothing
    sighentidades.ParaAuditoriaPorCadaDato sghAudLimpiar, ""
    Exit Sub
ErrCierre:
    Set oCommand = Nothing
End Sub

'Procedimiento para actualizar el estado de logeo de un usuario
Sub LogueaUsuario(Estado As Long, Usuario As Long, PC As String)
  Dim oConexSQL As New ADODB.Connection
'  Dim oTabla As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  
  oConexSQL.CursorLocation = adUseClient
  oConexSQL.CommandTimeout = 300
  oConexSQL.Open sighentidades.CadenaConexion
  
    'MODIFICADO POR FRANKLIN CACHAY 22/10/2013 - se cambio a store procedure
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexSQL
        .CommandTimeout = 150
        .CommandText = "LogueaUsuario"
        Set oParameter = .CreateParameter("@loginEstado", adInteger, adParamInput, 0, Estado): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@loginPC", adVarChar, adParamInput, 50, PC): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, Usuario): .Parameters.Append oParameter
        .Execute
    End With
    Set oParameter = Nothing
    
    oConexSQL.Close
    Set oConexSQL = Nothing
    Set oCommand = Nothing
'    Set oTabla = Nothing
End Sub

Function RetornaOpcionesReporteQueNoTieneAcceso(lnIdUsuario As Long) As Recordset
        Dim oRsTmp As New Recordset
        Dim oRsRolesRep As New Recordset
        Dim mrs_Tmp As New Recordset
        Dim lcSql As String
        Dim lnCorr As Long
        Dim lnIdReporte As Long
        '
        With mrs_Tmp
             .Fields.Append "id_menuReporte", adVarChar, 100, adFldIsNullable
             .LockType = adLockOptimistic
             .Open
        End With
        
        Set oRsRolesRep = ListBarReportesSeleccionarTodos
        '
         Dim oCommand As New ADODB.Command
         Dim oParameter As ADODB.Parameter
         Dim oConexion As New ADODB.Connection
         Dim ms_MensajeError  As String
         ms_MensajeError = ""
         Set oRsTmp = RetornaOpcionesReporteQueTieneAcceso(lnIdUsuario)
'         oConexion.CursorLocation = adUseClient
'         oConexion.CommandTimeout = 300
'         oConexion.Open sighentidades.CadenaConexion
'         With oCommand
'             .CommandType = adCmdStoredProc
'             Set .ActiveConnection = oConexion
'             .CommandTimeout = 150
'             .CommandText = "RetornaOpcionesReporteQueTieneAcceso"
'             Set oParameter = .CreateParameter("@lnIdUsuario", adInteger, adParamInput, 0, lnIdUsuario): .Parameters.Append oParameter
'             Set oRsTmp = .Execute
'             Set oRsTmp.ActiveConnection = Nothing
'        End With
'        oConexion.Close
'        Set oConexion = Nothing
'        Set oCommand = Nothing
        '
        If oRsTmp.RecordCount = 0 Then
            mrs_Tmp.AddNew
            mrs_Tmp.Fields!id_menuReporte = "ID_Reportes"
            mrs_Tmp.Update
            mrs_Tmp.AddNew
            mrs_Tmp.Fields!id_menuReporte = "ID_Herramientas"
            mrs_Tmp.Update
        Else
           oRsRolesRep.MoveFirst
           Do While Not oRsRolesRep.EOF
                lcSql = "id_menuReporte='" & oRsRolesRep.Fields!id_menuReporte & "'"
                oRsTmp.MoveFirst
                oRsTmp.Find lcSql
                If oRsTmp.EOF Then
                   mrs_Tmp.AddNew
                   mrs_Tmp.Fields!id_menuReporte = oRsRolesRep.Fields!id_menuReporte
                   mrs_Tmp.Update
                   
                End If
                oRsRolesRep.MoveNext
           Loop
        End If
        oRsTmp.Close
        Set oRsTmp = Nothing
        Set RetornaOpcionesReporteQueNoTieneAcceso = mrs_Tmp
End Function

Function ListBarReportesSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ListBarReporteSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ListBarReportesSeleccionarTodos = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description

End Function

Function RolesReportesSeleccionarXrol(lnIdRol As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RolesReportesSeleccionarXrol"
        Set oParameter = .CreateParameter("@lnIdRol", adInteger, adParamInput, 0, lnIdRol): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set RolesReportesSeleccionarXrol = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function EmpleadoVisualizaListaMedicamentosVencidos(lnIdUsuario As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRsTmp As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosLugarDeTrabajoSeleccionarPorId"
        Set oParameter = .CreateParameter("@lnIdUsuario", adInteger, adParamInput, 0, lnIdUsuario): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   


   EmpleadoVisualizaListaMedicamentosVencidos = False
   If oRsTmp.RecordCount > 0 Then
      oRsTmp.MoveFirst
      Do While Not oRsTmp.EOF
         If oRsTmp.Fields!idLaboraArea = 1 Then
            EmpleadoVisualizaListaMedicamentosVencidos = True
            Exit Do
         End If
         oRsTmp.MoveNext
      Loop
   End If
   oRsTmp.Close
   oConexion.Close
   Set oRsTmp = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function RolesSeleccionarXidentificador(lIdRol As Long, Optional oConexion1 As Connection) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
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
        .CommandText = "RolesSeleccionarXidentificador"
        Set oParameter = .CreateParameter("@lIdRol", adInteger, adParamInput, 0, lIdRol): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set RolesSeleccionarXidentificador = oRecordset
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function UsuariosRolesXidEmpleadoEsDeFarmacia(ml_idUsuario As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "UsuariosRolesXidEmpleadoEsDeFarmacia"
        Set oParameter = .CreateParameter("@ml_idUsuario", adInteger, adParamInput, 0, ml_idUsuario): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set UsuariosRolesXidEmpleadoEsDeFarmacia = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function PermisosSoloFarmacia(oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PermisosSoloFarmacia"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PermisosSoloFarmacia = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function TieneRolAdministrador(lnIdUsuario As Long) As Boolean

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TieneRolAdministrador"
        Set oParameter = .CreateParameter("@lnIdUsuario", adInteger, adParamInput, 0, lnIdUsuario): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   TieneRolAdministrador = False
   If oRecordset.RecordCount > 0 Then
      TieneRolAdministrador = True
   End If
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



Function AuditoriaFiltrarSoloRecalculos() As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AuditoriaFiltrarSoloRecalculos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AuditoriaFiltrarSoloRecalculos = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'debb-10/05/2016
Function AuditoriaFiltrarCitasPorIdAtencion(lnIdAtencion As Long, lnIdListBarItem As sghOpcionGalenHos, _
                                            Optional oConexion1 As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
    End If

    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "AuditoriaFiltrarCitasPorIdAtencion"
        Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdListBarItem", adInteger, adParamInput, 0, lnIdListBarItem): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set AuditoriaFiltrarCitasPorIdAtencion = oRecordset
   If oConexion1 Is Nothing Then
'        oRecordset.Close
        oConexion.Close
        Set oRecordset = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'mgaray20141008
Public Function UsuarioEsMedico(lnIdUsuario As Long) As Boolean
On Error GoTo miError
    Dim oRsTmp As ADODB.Recordset
    Dim oReglasDeProgMedica As New ReglasDeProgMedica
    Dim oConexion As New ADODB.Connection
    Dim bUsuarioEsMedico As Boolean
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    bUsuarioEsMedico = False
    Set oRsTmp = oReglasDeProgMedica.MedicosXidEmpleado(lnIdUsuario, oConexion)
    If oRsTmp.RecordCount > 0 Then
        bUsuarioEsMedico = True
    End If
    oConexion.Close
    Set oConexion = Nothing
    UsuarioEsMedico = bUsuarioEsMedico
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation, "Mensaje"
    End If
End Function

'A.Yañez 30-10-2014 ****************************************************
Public Function UsuarioEsMedicomejorado(lnIdUsuario As Long) As String
On Error GoTo miError
    Dim oRsTmp As ADODB.Recordset
    Dim oReglasDeProgMedica As New ReglasDeProgMedica
    Dim oConexion As New ADODB.Connection
    Dim bUsuarioEsMedico As String
    Dim ms_idColegioHis As String
            
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    bUsuarioEsMedico = "99"
    Set oRsTmp = oReglasDeProgMedica.MedicosXidEmpleado(lnIdUsuario, oConexion)
    If oRsTmp.RecordCount > 0 Then
       ms_idColegioHis = oRsTmp!idColegioHIS
       If ms_idColegioHis = "01" Then
          bUsuarioEsMedico = "1"
       Else
          bUsuarioEsMedico = "2"
       End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    UsuarioEsMedicomejorado = bUsuarioEsMedico
    If Err Then
miError:
        MsgBox Err.Number & " : " & Err.Description, vbExclamation, "Mensaje"
    End If
End Function
'*******************************************************************************

'debb2014d
Function SoloTieneOpcionCONSULTA(ml_idUsuario As Long, mo_lnIdTablaLISTBARITEMS As Long) As Boolean
  Dim rsTemp As New ADODB.Recordset
  Set rsTemp = RolesItemsSeleccionarPermisosPorEmpleadoYListItem(ml_idUsuario, mo_lnIdTablaLISTBARITEMS)
  SoloTieneOpcionCONSULTA = False
  If rsTemp.RecordCount > 0 Then
     If rsTemp!Agregar = 0 And rsTemp!modificar = 0 And rsTemp!Eliminar = 0 And rsTemp!consultar = 1 Then
        SoloTieneOpcionCONSULTA = True
     End If
  End If
End Function

Function RetornaOpcionesReporteQueTieneAcceso(lnIdUsuario As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RetornaOpcionesReporteQueTieneAcceso"
        Set oParameter = .CreateParameter("@lnIdUsuario", adInteger, adParamInput, 0, lnIdUsuario): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set RetornaOpcionesReporteQueTieneAcceso = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



Function ListbarItemsSeleccionarPorClave(lcClave As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ListbarItemsSeleccionarPorClave"
        Set oParameter = .CreateParameter("@clave", adVarChar, adParamInput, 50, lcClave): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ListbarItemsSeleccionarPorClave = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
'SCCQ 28/08/2020 Cambio27 Inicio
Function EmpleadoVisualizaListaHCSinDevolver(lnIdUsuario As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRsTmp As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "usp_EmpleadosRolSeleccionarPorId"
        Set oParameter = .CreateParameter("@lnIdUsuario", adInteger, adParamInput, 0, lnIdUsuario): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   EmpleadoVisualizaListaHCSinDevolver = False
   If oRsTmp.RecordCount > 0 Then
      oRsTmp.MoveFirst
      Do While Not oRsTmp.EOF
         If oRsTmp.Fields!IdRol = 62 Then 'IdRol=62 es para "ArchivoClinico"
            EmpleadoVisualizaListaHCSinDevolver = True
            Exit Do
         End If
         oRsTmp.MoveNext
      Loop
   End If
   oRsTmp.Close
   oConexion.Close
   Set oRsTmp = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
'SCCQ 28/08/2020 Cambio27 Fin
