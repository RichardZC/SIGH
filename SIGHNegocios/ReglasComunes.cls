VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasComunes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Reglas Comunes
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim lcSql As String
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Function TiposProgramacionSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposProgramacion
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposProgramacionSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function TiposOcupacionSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposOcupacion
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposOcupacionSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oConexion = Nothing

End Function
Function TiposOcupacionTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposOcupacion
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposOcupacionTodos = oTabla.Todos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oConexion = Nothing

End Function

Function TiposDocIdentidadSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposDocIdentidad
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposDocIdentidadSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oConexion = Nothing

End Function

Function TiposEstadoCivilSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposEstadoCivil
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposEstadoCivilSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oConexion = Nothing

End Function

Function TiposEstadoCivilTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposEstadoCivil
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposEstadoCivilTodos = oTabla.Todos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oConexion = Nothing

End Function


Function TiposGradosInstruccionSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposGradoInstruccion
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposGradosInstruccionSeleccionarTodos = oTabla.Todos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oConexion = Nothing

End Function

Function TiposGradosInstruccionTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposGradoInstruccion
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposGradosInstruccionTodos = oTabla.Todos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oConexion = Nothing

End Function
Function TiposProcedenciaSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposProcedencia
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposProcedenciaSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oConexion = Nothing

End Function
Function TiposProcedenciaTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposProcedencia
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposProcedenciaTodos = oTabla.Todos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oConexion = Nothing

End Function

Function TiposSexoSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposSexo
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposSexoSeleccionarTodos = oTabla.SeleccionarTodos
    
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oConexion = Nothing

End Function


Function PaisCodigoPostalPorIdentificador(lnIdPais As Long, oConexion As Connection) As String
On Error GoTo miError
    Dim oPais As New Paises
    Dim oDOPais As New DOPais
    PaisCodigoPostalPorIdentificador = ""
    Set oPais.Conexion = oConexion
    oDOPais.idPais = lnIdPais
    If oPais.SeleccionarPorId(oDOPais) Then
        PaisCodigoPostalPorIdentificador = oDOPais.codigoPostal
    End If
    ms_MensajeError = oPais.MensajeError
    Set oPais = Nothing
    Set oDOPais = Nothing
miError:
    If Err Then
        ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    End If
End Function

Function TiposEmpleadoSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposEmpleado
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposEmpleadoSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function TiposCondicionTrabajoSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposCondicionTrabajo
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposCondicionTrabajoSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'-------------------------------------------------------------------------------------
'                                       Empleados
'-------------------------------------------------------------------------------------

Function EmpleadosAgregar(oDOEmpleado As DOEmpleado, oUsuariosRoles As Collection, oRsCargos As Recordset, oRsLaboraEn As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreEmpleado As String, lcDniEncriptado As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oEmpleado As New Empleados
Dim oUsuarioRol As New UsuariosRoles
Dim oRsTmp As New ADODB.Recordset
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim lnIdEspecialidad  As Long
Dim lcSql As String
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    On Error GoTo ErrEmpl
    EmpleadosAgregar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    
    Set oEmpleado.Conexion = oConexion
    Set oUsuarioRol.Conexion = oConexion
    
    If oEmpleado.Insertar(oDOEmpleado) Then
        If oUsuarioRol.ActualizarPorEmpleado(oUsuariosRoles, oDOEmpleado.idEmpleado) Then
            'cargos
            If oRsCargos.RecordCount > 0 Then
                oRsCargos.MoveFirst
                Do While Not oRsCargos.EOF
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "EmpleadosCargosAgregar"
                        Set oParameter = .CreateParameter("@IdTipoCargo", adInteger, adParamInput, 0, oRsCargos.Fields!idTipoCargo): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, oDOEmpleado.idEmpleado): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oParameter = Nothing
                    oRsCargos.MoveNext
                Loop
            End If
            'Lugar donde Labora
            Dim lnIdEspecialidadAC As Long
            Dim lbTrabajaEnArchivoClinico As Boolean
            lbTrabajaEnArchivoClinico = False

            Set oCommand = Nothing
            Set oParameter = Nothing
            With oCommand
                 .CommandType = adCmdStoredProc
                 Set .ActiveConnection = oConexion
                 .CommandTimeout = 150
                 .CommandText = "ConsultaServiciosXidServicio"
                 Set oParameter = .CreateParameter("@idServicio", adInteger, adParamInput, 0, Trim(Str(mo_AdminComun.ParametrosIdServicioArchivoClinico()))): .Parameters.Append oParameter
                 Set oRsTmp = .Execute
                 Set oRsTmp.ActiveConnection = Nothing
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
            If oRsTmp.RecordCount = 0 Then
               MsgBox "     Se ha eliminado ARCHIVOS CLINICOS de la tabla SERVICIOS     " & Chr(13) & _
                      "Solución: agregarlo manualmente en la tabla SERVICIOS (DESDE SQL)" & _
                      "                     el IdServicio ANOTARLO                      " & _
                      "En la tabla PARAMETROS ubicarse en el Id=9 poner en VALORINT el idServicio ANOTADO", vbInformation, "Empleados"
               GoTo ErrEmpl
            Else
               lnIdEspecialidadAC = oRsTmp.Fields!IdEspecialidad
            End If
            oRsTmp.Close
            If oRsLaboraEn.RecordCount > 0 Then
                oRsLaboraEn.MoveFirst
                Do While Not oRsLaboraEn.EOF
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "EmpleadosLugarDeTrabajoAgregar"
                        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, oDOEmpleado.idEmpleado): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdLaboraArea", adInteger, adParamInput, 0, oRsLaboraEn.Fields!idLaboraArea): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdLaboraSubArea", adInteger, adParamInput, 0, oRsLaboraEn.Fields!idLaboraSubArea): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oParameter = Nothing
                    
                   If lnIdEspecialidadAC = oRsLaboraEn.Fields!idLaboraSubArea Then
                      lbTrabajaEnArchivoClinico = True
                   End If
                   oRsLaboraEn.MoveNext
                Loop
            End If
            'archivero (Archivo Clinico)
            If lbTrabajaEnArchivoClinico = True Then
                Dim oBuscaServicios As New SIGHNegocios.ReglasAdmision
                Dim oRsSubArea As New Recordset
                Set oRsSubArea = oBuscaServicios.DevuelveServiciosDelHospital("(1,2,3,4)", "", sghFiltraAnuladosYactivos, sghPorDescTipoServicio)
                If oRsSubArea.RecordCount > 0 Then
                   oRsSubArea.MoveFirst
                   Do While Not oRsSubArea.EOF
                        Set oCommand = Nothing
                        Set oParameter = Nothing
                        With oCommand
                            .CommandType = adCmdStoredProc
                            Set .ActiveConnection = oConexion
                            .CommandTimeout = 150
                            .CommandText = "ArchiveroServicioAgregar"
                            Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, oRsSubArea.Fields!IdServicio): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, oDOEmpleado.idEmpleado): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@IdArchivero", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@EsConsultorioAsignado", adBoolean, adParamInput, 0, False): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                            .Execute
                        End With
                    
                      oRsSubArea.MoveNext
                   Loop
                End If
                oRsSubArea.Close
                Set oRsSubArea = Nothing
                Set oBuscaServicios = Nothing
            End If
            '
            If lcDniEncriptado <> "" Then
                If oRsTmp.State = 1 Then oRsTmp.Close
                Set oCommand = Nothing
                Set oParameter = Nothing
                With oCommand
                     .CommandType = adCmdStoredProc
                     Set .ActiveConnection = oConexion
                     .CommandTimeout = 150
                     .CommandText = "parametrosModificarDniEncriptado"
                     Set oParameter = .CreateParameter("@DniEncriptado", adVarChar, adParamInput, 355, lcDniEncriptado): .Parameters.Append oParameter
                     .Execute
                End With
                
            End If
            '
            Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOEmpleado.IdUsuarioAuditoria, "A", oDOEmpleado.idEmpleado, "Empleados", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreEmpleado)          'ListBarItems.idListItem
            EmpleadosAgregar = True
        End If
    Else
        ms_MensajeError = oEmpleado.MensajeError
    End If
ErrEmpl:
   If Err.Number = 0 Then
      oConexion.CommitTrans
   Else
      ms_MensajeError = Err.Description
   End If
   Set oConexion = Nothing
   Set oEmpleado = Nothing
   Set oUsuarioRol = Nothing
   Set oRsTmp = Nothing
   Set mo_AdminComun = Nothing
End Function

Function EmpleadosModificar(oDOEmpleado As DOEmpleado, oUsuariosRoles As Collection, oRsCargos As Recordset, oRsLaboraEn As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreEmpleado As String, lcDniEncriptado As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oEmpleado As New Empleados
Dim oUsuarioRol As New UsuariosRoles
Dim oRsTmp As New ADODB.Recordset
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim lcSql As String
Dim lnIdEspecialidad As Long
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    
    On Error GoTo ErrEmplM
    
    EmpleadosModificar = False
    oConexion.BeginTrans
    
    Set oEmpleado.Conexion = oConexion
    Set oUsuarioRol.Conexion = oConexion
    
    If oEmpleado.modificar(oDOEmpleado) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOEmpleado.IdUsuarioAuditoria, "M", oDOEmpleado.idEmpleado, "Empleados", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")
        If oUsuarioRol.ActualizarPorEmpleado(oUsuariosRoles, oDOEmpleado.idEmpleado) Then
            'Cargos
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "EmpleadosCargosEliminar"
                Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, oDOEmpleado.idEmpleado): .Parameters.Append oParameter
                .Execute
            End With
            Set oParameter = Nothing
            
            If oRsCargos.RecordCount > 0 Then
                oRsCargos.MoveFirst
                Do While Not oRsCargos.EOF
                   Set oCommand = Nothing
                   Set oParameter = Nothing
                   With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "EmpleadosCargosAgregar"
                        Set oParameter = .CreateParameter("@IdTipoCargo", adInteger, adParamInput, 0, oRsCargos.Fields!idTipoCargo): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, oDOEmpleado.idEmpleado): .Parameters.Append oParameter
                        .Execute
                   End With
                   
                   oRsCargos.MoveNext
                Loop
            End If
            'Lugar donde Labora
            Set oCommand = Nothing
            Set oParameter = Nothing
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "EmpleadosLugarDeTrabajoEliminar"
                Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, oDOEmpleado.idEmpleado): .Parameters.Append oParameter
                .Execute
            End With

            Dim lnIdEspecialidadAC As Long
            Dim lbTrabajaEnArchivoClinico As Boolean
            lbTrabajaEnArchivoClinico = False

                Set oCommand = Nothing
                Set oParameter = Nothing
                With oCommand
                     .CommandType = adCmdStoredProc
                     Set .ActiveConnection = oConexion
                     .CommandTimeout = 150
                     .CommandText = "ConsultaServiciosXidServicio"
                     Set oParameter = .CreateParameter("@idServicio", adInteger, adParamInput, 0, Trim(Str(mo_AdminComun.ParametrosIdServicioArchivoClinico()))): .Parameters.Append oParameter
                     Set oRsTmp = .Execute
                     Set oRsTmp.ActiveConnection = Nothing
                End With
                
            lnIdEspecialidadAC = 0
            If oRsTmp.RecordCount > 0 Then
               lnIdEspecialidadAC = oRsTmp.Fields!IdEspecialidad
            End If
            oRsTmp.Close
            If oRsLaboraEn.RecordCount > 0 Then
                oRsLaboraEn.MoveFirst
                Do While Not oRsLaboraEn.EOF
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "EmpleadosLugarDeTrabajoAgregar"
                        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, oDOEmpleado.idEmpleado): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdLaboraArea", adInteger, adParamInput, 0, oRsLaboraEn.Fields!idLaboraArea): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdLaboraSubArea", adInteger, adParamInput, 0, oRsLaboraEn.Fields!idLaboraSubArea): .Parameters.Append oParameter
                        .Execute
                    End With
                   
                   If lnIdEspecialidadAC = oRsLaboraEn.Fields!idLaboraSubArea Then
                      lbTrabajaEnArchivoClinico = True
                   End If
                   oRsLaboraEn.MoveNext
                Loop
            End If
            'archivero (Archivo Clinico)
            Set oCommand = Nothing
            Set oParameter = Nothing
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "ArchiveroServicioEliminarXIdEmpleado"
                Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, oDOEmpleado.idEmpleado): .Parameters.Append oParameter
                .Execute
            End With

            If lbTrabajaEnArchivoClinico = True Then
                Dim oBuscaServicios As New SIGHNegocios.ReglasAdmision
                Dim oRsSubArea As New Recordset
                Set oRsSubArea = oBuscaServicios.DevuelveServiciosDelHospital("(1,2,3,4)", "", sghFiltraAnuladosYactivos, sghPorDescTipoServicio)
                If oRsSubArea.RecordCount > 0 Then
                   oRsSubArea.MoveFirst
                   Do While Not oRsSubArea.EOF
                        Set oCommand = Nothing
                        Set oParameter = Nothing
                        With oCommand
                            .CommandType = adCmdStoredProc
                            Set .ActiveConnection = oConexion
                            .CommandTimeout = 150
                            .CommandText = "ArchiveroServicioAgregar"
                            Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, oRsSubArea.Fields!IdServicio): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, oDOEmpleado.idEmpleado): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@IdArchivero", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@EsConsultorioAsignado", adBoolean, adParamInput, 0, False): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                            .Execute
                        End With
                        
                      oRsSubArea.MoveNext
                   Loop
                End If
                oRsSubArea.Close
                Set oRsSubArea = Nothing
                Set oBuscaServicios = Nothing
            End If
            '
            If lcDniEncriptado <> "" Then
                If oRsTmp.State = 1 Then oRsTmp.Close

                Set oCommand = Nothing
                Set oParameter = Nothing
                With oCommand
                     .CommandType = adCmdStoredProc
                     Set .ActiveConnection = oConexion
                     .CommandTimeout = 150
                     .CommandText = "parametrosModificarDniEncriptado"
                     Set oParameter = .CreateParameter("@DniEncriptado", adVarChar, adParamInput, 355, lcDniEncriptado): .Parameters.Append oParameter
                     .Execute
                End With
                
            End If
            EmpleadosModificar = True
        End If
    Else
        ms_MensajeError = oEmpleado.MensajeError
    End If
ErrEmplM:
 If Err.Number = 0 Then
      oConexion.CommitTrans
   Else
      oConexion.RollbackTrans
      ms_MensajeError = Err.Description
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oEmpleado = Nothing
   Set oUsuarioRol = Nothing
   Set oRsTmp = Nothing
   Set mo_AdminComun = Nothing
End Function

'Modificado por Yamill Palomino 07-10-13 Se cambio a Store Procedure
Function EmpleadosEliminar(oDOEmpleado As DOEmpleado, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, _
                                             lcNombreEmpleado As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oEmpleado As New Empleados
Dim oUsuarioRol As New UsuariosRoles
Dim oRsTmp As New ADODB.Recordset
Dim oRsTmp1 As New ADODB.Recordset
Dim lcSql As String
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion

    EmpleadosEliminar = False
    oConexion.BeginTrans
    Set oEmpleado.Conexion = oConexion
    Set oUsuarioRol.Conexion = oConexion
    'Medicos
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosMedicosConsultar"
        Set oParameter = .CreateParameter("@idEmpleado", adInteger, adParamInput, 0, oDOEmpleado.idEmpleado): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With

    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
            Set oCommand = Nothing
            Set oParameter = Nothing
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "MedicosEspecialidadEliminarPorIdMedico"
                Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, oRsTmp.Fields!IdMedico): .Parameters.Append oParameter
                .Execute
            End With
            '
            Set oCommand = Nothing
            Set oParameter = Nothing
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "MedicosEliminar"
                Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, oRsTmp.Fields!IdMedico): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oDOEmpleado.IdUsuarioAuditoria): .Parameters.Append oParameter
                .Execute
            End With
            'oRsTmp.Delete
            'oRsTmp.Update
          '
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
'    'cargos
    Set oCommand = Nothing
    Set oParameter = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosCargosLugarDeTrabajoArchiveroServicioEliminar"
        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, oDOEmpleado.idEmpleado): .Parameters.Append oParameter
        .Execute
    End With
    
    If oUsuarioRol.EliminarPorEmpleado(oDOEmpleado.idEmpleado) Then
        If oEmpleado.Eliminar(oDOEmpleado) Then
            '
            Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOEmpleado.IdUsuarioAuditoria, "E", oDOEmpleado.idEmpleado, "Empleados", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreEmpleado)          'ListBarItems.idListItem
            EmpleadosEliminar = True
            oConexion.CommitTrans
        Else
            ms_MensajeError = oEmpleado.MensajeError
            oConexion.RollbackTrans
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oEmpleado = Nothing
    Set oUsuarioRol = Nothing
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
End Function

Function EmpleadosSeleccionarPorId(idEmpleado As Long) As DOEmpleado
Dim oConexion As New ADODB.Connection
Dim oEmpleado As New Empleados
Dim oDOEmpleado As New DOEmpleado
    
    Set EmpleadosSeleccionarPorId = Nothing
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oEmpleado.Conexion = oConexion
    oDOEmpleado.idEmpleado = idEmpleado
    
    If oEmpleado.SeleccionarPorId(oDOEmpleado) Then
        Set EmpleadosSeleccionarPorId = oDOEmpleado
    End If
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oEmpleado.MensajeError
    Set oEmpleado = Nothing
End Function
Function EmpleadosSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Empleados
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EmpleadosSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function
Function EmpleadosFiltrar(oDOEmpleado As DOEmpleado) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Empleados
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EmpleadosFiltrar = oTabla.Filtrar(oDOEmpleado)
    oConexion.Close

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function
Function EmpleadosSeleccionarPorCodigo(Codigo As String, oDOEmpleado As DOEmpleado) As Boolean
Dim oConexion As New ADODB.Connection
Dim oEmpleado As New Empleados
Dim rsMedicoEspecialidad As Recordset

    
    EmpleadosSeleccionarPorCodigo = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oEmpleado.Conexion = oConexion
    oDOEmpleado.CodigoPlanilla = Codigo
    
    If oEmpleado.SeleccionarPorCodigo(oDOEmpleado) Then
            EmpleadosSeleccionarPorCodigo = True
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oEmpleado = Nothing

End Function
Function EmpleadosObtenerConElMismoUsuario(oDOEmpleado As DOEmpleado) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oEmpleado As New Empleados
Dim rsMedicoEspecialidad As Recordset
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oEmpleado.Conexion = oConexion
    Set EmpleadosObtenerConElMismoUsuario = oEmpleado.ObtenerConElMismoUsuario(oDOEmpleado)
    oConexion.Close
    Set oConexion = Nothing
    Set oEmpleado = Nothing
    
End Function

Function EmpleadosObtenerConElMismoCodigoPlanilla(oDOEmpleado As DOEmpleado) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oEmpleado As New Empleados
Dim rsMedicoEspecialidad As Recordset

    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oEmpleado.Conexion = oConexion
    Set EmpleadosObtenerConElMismoCodigoPlanilla = oEmpleado.ObtenerConElMismoCodigoPlanilla(oDOEmpleado)
    oConexion.Close
    
    Set oConexion = Nothing
    Set oEmpleado = Nothing
End Function

Function EmpleadosObtenerConelMismoDNI(lcDNI As String, lnIdTipoDocumento As Long) As Recordset
    Dim oEmpleado As New Empleados
    Set EmpleadosObtenerConelMismoDNI = oEmpleado.ObtenerConElMismoDNI(lcDNI, lnIdTipoDocumento)
    Set oEmpleado = Nothing
End Function

Function EmpleadosObtenerConLaMismaCOLEGIATURA(lcCOLEGIATURA As String) As Recordset
   Dim oEmpleado As New Empleados
   Set EmpleadosObtenerConLaMismaCOLEGIATURA = oEmpleado.ObtenerConLaMismaCOLEGIATURA(lcCOLEGIATURA)
    Set oEmpleado = Nothing
End Function


Function TiposCitaSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposCita
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposCitaSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function TiposConsultaSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposConsulta
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposConsultaSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


Function ParametrosComparadorFechas() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Parametros
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set ParametrosComparadorFechas = oTabla.SeleccionarComparadorFechas()
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function ParametrosIdServicioArchivoClinico() As Long
Dim oConexion As New ADODB.Connection
Dim oTabla As New Parametros
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    ParametrosIdServicioArchivoClinico = oTabla.SeleccionarIdServicioArchivoClinico()
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'-------------------------------------------------------------------------------------
'                                       Procedimientos
'-------------------------------------------------------------------------------------

Function ProcedimientosAgregar(oDOProcedimiento As doProcedimiento) As Boolean
Dim oConexion As New ADODB.Connection
Dim oProcedimiento As New Procedimientos
    
    ProcedimientosAgregar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    Set oProcedimiento.Conexion = oConexion
    If oProcedimiento.Insertar(oDOProcedimiento) Then
        ProcedimientosAgregar = True
    Else
        ms_MensajeError = oProcedimiento.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oProcedimiento = Nothing
    
End Function

Function ProcedimientosModificar(oDOProcedimiento As doProcedimiento) As Boolean
Dim oConexion As New ADODB.Connection
Dim oProcedimiento As New Procedimientos
    
    ProcedimientosModificar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    Set oProcedimiento.Conexion = oConexion
    If oProcedimiento.modificar(oDOProcedimiento) Then
        ProcedimientosModificar = True
    Else
        ms_MensajeError = oProcedimiento.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oProcedimiento = Nothing
    
End Function

Function ProcedimientosEliminar(oDOProcedimiento As doProcedimiento) As Boolean
Dim oConexion As New ADODB.Connection
Dim oProcedimiento As New Procedimientos
    
    ProcedimientosEliminar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    Set oProcedimiento.Conexion = oConexion
    If oProcedimiento.Eliminar(oDOProcedimiento) Then
        ProcedimientosEliminar = True
    Else
        ms_MensajeError = oProcedimiento.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oProcedimiento = Nothing
    
End Function

Function ProcedimientosSeleccionarPorId(IdProcedimiento As Long) As doProcedimiento
Dim oConexion As New ADODB.Connection
Dim oProcedimiento As New Procedimientos
Dim oDOProcedimiento As New doProcedimiento
    
    Set ProcedimientosSeleccionarPorId = Nothing
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oProcedimiento.Conexion = oConexion
    oDOProcedimiento.IdProcedimiento = IdProcedimiento
    If oProcedimiento.SeleccionarPorId(oDOProcedimiento) Then
        Set ProcedimientosSeleccionarPorId = oDOProcedimiento
    End If
    oConexion.Close

    ms_MensajeError = oProcedimiento.MensajeError
    Set oConexion = Nothing
    Set oProcedimiento = Nothing

End Function
Function ProcedimientosSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Procedimientos
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set ProcedimientosSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function ProcedimientosFiltrar(oDOProcedimiento As doProcedimiento) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Procedimientos
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set ProcedimientosFiltrar = oTabla.Filtrar(oDOProcedimiento)
    oConexion.Close

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'-------------------------------------------------------------------------------------
'                                       Diagnosticos
'-------------------------------------------------------------------------------------

Function DiagnosticosAgregar(oDODiagnostico As doDiagnostico) As Boolean
Dim oConexion As New ADODB.Connection
Dim oDiagnostico As New Diagnosticos
    
    DiagnosticosAgregar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    Set oDiagnostico.Conexion = oConexion
    If oDiagnostico.Insertar(oDODiagnostico) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDODiagnostico.IdUsuarioAuditoria, "A", oDODiagnostico.IdDiagnostico, _
                                                    "Diagnostico", oConexion, 1202, SIGHEntidades.RetornaNombrePC, _
                                                    Trim(oDODiagnostico.CodigoCIE2004) & "-" & oDODiagnostico.Descripcion)
        DiagnosticosAgregar = True
    Else
        ms_MensajeError = oDiagnostico.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oDiagnostico = Nothing
    
End Function

Function DiagnosticosModificar(oDODiagnostico As doDiagnostico) As Boolean
Dim oConexion As New ADODB.Connection
Dim oDiagnostico As New Diagnosticos
    
    DiagnosticosModificar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    Set oDiagnostico.Conexion = oConexion
    If oDiagnostico.modificar(oDODiagnostico) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDODiagnostico.IdUsuarioAuditoria, "M", oDODiagnostico.IdDiagnostico, _
                                                    "Diagnostico", oConexion, 1202, SIGHEntidades.RetornaNombrePC, _
                                                    Trim(oDODiagnostico.CodigoCIE2004) & "-" & oDODiagnostico.Descripcion)
        DiagnosticosModificar = True
    Else
        ms_MensajeError = oDiagnostico.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oDiagnostico = Nothing
    
End Function

Function DiagnosticosEliminar(oDODiagnostico As doDiagnostico) As Boolean
Dim oConexion As New ADODB.Connection
Dim oDiagnostico As New Diagnosticos
    
    DiagnosticosEliminar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    Set oDiagnostico.Conexion = oConexion
    If oDiagnostico.Eliminar(oDODiagnostico) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDODiagnostico.IdUsuarioAuditoria, "E", oDODiagnostico.IdDiagnostico, _
                                                    "Diagnostico", oConexion, 1202, SIGHEntidades.RetornaNombrePC, _
                                                    Trim(oDODiagnostico.CodigoCIE2004) & "-" & oDODiagnostico.Descripcion)
        DiagnosticosEliminar = True
    Else
        ms_MensajeError = oDiagnostico.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oDiagnostico = Nothing
    
End Function

Function DiagnosticosSeleccionarPorId(IdDiagnostico As Long) As doDiagnostico
Dim oConexion As New ADODB.Connection
Dim oDiagnostico As New Diagnosticos
Dim oDODiagnostico As New doDiagnostico
    
    Set DiagnosticosSeleccionarPorId = Nothing
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oDiagnostico.Conexion = oConexion
    oDODiagnostico.IdDiagnostico = IdDiagnostico
    If oDiagnostico.SeleccionarPorId(oDODiagnostico) Then
        Set DiagnosticosSeleccionarPorId = oDODiagnostico
    End If
    oConexion.Close
    ms_MensajeError = oDiagnostico.MensajeError
    Set oConexion = Nothing
    Set oDiagnostico = Nothing
    

End Function
Function DiagnosticosSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Diagnosticos
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set DiagnosticosSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


Function TiposReferenciaSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposReferencia
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposReferenciaSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function EstablecimientosSeleccionarPorId(IdEstablecimiento As Long) As DOEstablecimiento
Dim oConexion As New ADODB.Connection
Dim oEstablecimiento As New Establecimientos
Dim oDOEstablecimiento As New DOEstablecimiento
    
    Set EstablecimientosSeleccionarPorId = Nothing
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oEstablecimiento.Conexion = oConexion
    oDOEstablecimiento.IdEstablecimiento = IdEstablecimiento
    If oEstablecimiento.SeleccionarPorId(oDOEstablecimiento) Then
        Set EstablecimientosSeleccionarPorId = oDOEstablecimiento
    End If
    oConexion.Close

    ms_MensajeError = oEstablecimiento.MensajeError
    Set oConexion = Nothing
    Set oEstablecimiento = Nothing

End Function

Function EstablecimientosFiltrar(ByVal oEstablecimiento As DOEstablecimiento, _
                lDepartamento As Long, lProvincia As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Establecimientos
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EstablecimientosFiltrar = oTabla.Filtrar(oEstablecimiento, lDepartamento, lProvincia)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EstablecimientosSeleccionarPorCodigo(Codigo As String, oDOEstablecimiento As DOEstablecimiento) As Boolean
Dim oConexion As New ADODB.Connection
Dim oEstablecimiento As New Establecimientos
Dim rsMedicoEspecialidad As Recordset

    
    EstablecimientosSeleccionarPorCodigo = False
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    
    
    Set oEstablecimiento.Conexion = oConexion
    oDOEstablecimiento.Codigo = Codigo
    
    If oEstablecimiento.SeleccionarPorCodigo(oDOEstablecimiento) Then
            EstablecimientosSeleccionarPorCodigo = True
    End If
    Set oEstablecimiento = Nothing
End Function

'-------------------------------------------------------------------------------------
'                                       Establecimientos No Minsa
'-------------------------------------------------------------------------------------

Function EstablecimientosNoMinsaAgregar(oDoEstablecimientoNoMinsa As DOEstablecimientoNoMinsa, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcEstablecimiento As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim EstablecimientosNoMinsa As New EstablecimientosNoMinsa
    
    EstablecimientosNoMinsaAgregar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    Set EstablecimientosNoMinsa.Conexion = oConexion
    If EstablecimientosNoMinsa.Insertar(oDoEstablecimientoNoMinsa) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoEstablecimientoNoMinsa.IdUsuarioAuditoria, "A", oDoEstablecimientoNoMinsa.IdEstablecimientoNoMinsa, "EstablecimientosNoMinsa", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcEstablecimiento)           'ListBarItems.idListItem
        EstablecimientosNoMinsaAgregar = True
    Else
        ms_MensajeError = EstablecimientosNoMinsa.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set EstablecimientosNoMinsa = Nothing
    
End Function

Function EstablecimientosNoMinsaModificar(oDoEstablecimientoNoMinsa As DOEstablecimientoNoMinsa, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcEstablecimiento As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim EstablecimientosNoMinsa As New EstablecimientosNoMinsa
    
    EstablecimientosNoMinsaModificar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    Set EstablecimientosNoMinsa.Conexion = oConexion
    If EstablecimientosNoMinsa.modificar(oDoEstablecimientoNoMinsa) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoEstablecimientoNoMinsa.IdUsuarioAuditoria, "M", oDoEstablecimientoNoMinsa.IdEstablecimientoNoMinsa, "EstablecimientosNoMinsa", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcEstablecimiento)           'ListBarItems.idListItem
        EstablecimientosNoMinsaModificar = True
    Else
        ms_MensajeError = EstablecimientosNoMinsa.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set EstablecimientosNoMinsa = Nothing
    
End Function

Function EstablecimientosNoMinsaEliminar(oDoEstablecimientoNoMinsa As DOEstablecimientoNoMinsa, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcEstablecimiento As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim EstablecimientosNoMinsa As New EstablecimientosNoMinsa
    
    EstablecimientosNoMinsaEliminar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    Set EstablecimientosNoMinsa.Conexion = oConexion
    If EstablecimientosNoMinsa.Eliminar(oDoEstablecimientoNoMinsa) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoEstablecimientoNoMinsa.IdUsuarioAuditoria, "E", oDoEstablecimientoNoMinsa.IdEstablecimientoNoMinsa, "EstablecimientosNoMinsa", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcEstablecimiento)           'ListBarItems.idListItem
        EstablecimientosNoMinsaEliminar = True
    Else
        ms_MensajeError = EstablecimientosNoMinsa.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set EstablecimientosNoMinsa = Nothing
    
End Function

Function EstablecimientosNoMinsaSeleccionarPorId(IdEstablecimientoNoMinsa As Long) As DOEstablecimientoNoMinsa
Dim oConexion As New ADODB.Connection
Dim oEstablecimientoNoMinsa As New EstablecimientosNoMinsa
Dim oDoEstablecimientoNoMinsa As New DOEstablecimientoNoMinsa
    
    Set EstablecimientosNoMinsaSeleccionarPorId = Nothing
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oEstablecimientoNoMinsa.Conexion = oConexion
    oDoEstablecimientoNoMinsa.IdEstablecimientoNoMinsa = IdEstablecimientoNoMinsa
    If oEstablecimientoNoMinsa.SeleccionarPorId(oDoEstablecimientoNoMinsa) Then
        Set EstablecimientosNoMinsaSeleccionarPorId = oDoEstablecimientoNoMinsa
    End If
    oConexion.Close

    ms_MensajeError = oEstablecimientoNoMinsa.MensajeError
    Set oConexion = Nothing

End Function

Function EstablecimientosNoMinsaSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New EstablecimientosNoMinsa
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EstablecimientosNoMinsaSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function EstablecimientosNoMinsaFiltrar(ByVal oEstablecimiento As DOEstablecimientoNoMinsa, lDepartamento As Long, lProvincia As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New EstablecimientosNoMinsa
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EstablecimientosNoMinsaFiltrar = oTabla.Filtrar(oEstablecimiento, lDepartamento, lProvincia)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function TiposSubsectorSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposSubsector
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposSubsectorSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function TiposCondicionAltaSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposCondicionAlta
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposCondicionAltaSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function TiposAltaSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposAlta
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposAltaSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function DiagnosticosCapitulosSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New DiagnosticosCapitulos
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set DiagnosticosCapitulosSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function DiagnosticosGruposSeleccionarPorCapitulo(IdCapitulo As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New DiagnosticosGrupos
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set DiagnosticosGruposSeleccionarPorCapitulo = oTabla.SeleccionarPorCapitulo(IdCapitulo)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


Function DiagnosticosCategoriaSeleccionarPorGrupo(IdGrupo As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New DiagnosticosCategorias
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set DiagnosticosCategoriaSeleccionarPorGrupo = oTabla.SeleccionarPorGrupo(IdGrupo)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function DiagnosticosFiltrar(oDODiagnostico As doDiagnostico, lbSoloMuestraDxGalenHos As Boolean, lbUSAcodigoCIEsinPto As Boolean) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Diagnosticos
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    
    
    Set oTabla.Conexion = oConexion
    Set DiagnosticosFiltrar = oTabla.Filtrar(oDODiagnostico, lbSoloMuestraDxGalenHos, lbUSAcodigoCIEsinPto)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'mgaray09
Function DiagnosticosFiltrarSoloActivos(oDODiagnostico As doDiagnostico, lbSoloMuestraDxGalenHos As Boolean, lbUSAcodigoCIEsinPto As Boolean) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Diagnosticos
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    
    
    Set oTabla.Conexion = oConexion
    Set DiagnosticosFiltrarSoloActivos = oTabla.FiltrarSoloActivos(oDODiagnostico, lbSoloMuestraDxGalenHos, lbUSAcodigoCIEsinPto)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function DiagnosticosSeleccionarPorCodigoCIE2004(sCodigo As String, lbSoloMuestraDxGalenHos As Boolean) As doDiagnostico
Dim oConexion As New ADODB.Connection
Dim oTabla As New Diagnosticos
Dim doDiagnostico As New doDiagnostico
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    doDiagnostico.CodigoCIE2004 = sCodigo
    If oTabla.SeleccionarPorCodigoCIE2004(doDiagnostico, lbSoloMuestraDxGalenHos) Then
        Set DiagnosticosSeleccionarPorCodigoCIE2004 = doDiagnostico
    Else
        Set DiagnosticosSeleccionarPorCodigoCIE2004 = Nothing
    End If
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
    Set doDiagnostico = Nothing
End Function

Function SubclasificacionDiagnosticosSeleccionarDxConsultaExterna() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New SubclasificacionDiagnosticos
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set SubclasificacionDiagnosticosSeleccionarDxConsultaExterna = oTabla.SeleccionarDxConsultaExterna
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function SubclasificacionDiagnosticosSeleccionarDxHospIngreso() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New SubclasificacionDiagnosticos
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set SubclasificacionDiagnosticosSeleccionarDxHospIngreso = oTabla.SeleccionarDxHospitalizacionIngreso
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function SubclasificacionDiagnosticosSeleccionarDxHospEgreso() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New SubclasificacionDiagnosticos
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set SubclasificacionDiagnosticosSeleccionarDxHospEgreso = oTabla.SeleccionarDxHospitalizacionEgreso
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function SubclasificacionDiagnosticosSeleccionarDxHospMortalidad() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New SubclasificacionDiagnosticos
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set SubclasificacionDiagnosticosSeleccionarDxHospMortalidad = oTabla.SeleccionarDxHospitalizacionMortalidad
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function SubclasificacionDiagnosticosSeleccionarDxHospMuerteFetal() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New SubclasificacionDiagnosticos
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set SubclasificacionDiagnosticosSeleccionarDxHospMuerteFetal = oTabla.SeleccionarDxHospitalizacionMuerteFetal
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function SubclasificacionDiagnosticosSeleccionarDxHospComplicaciones() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New SubclasificacionDiagnosticos
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set SubclasificacionDiagnosticosSeleccionarDxHospComplicaciones = oTabla.SeleccionarDxHospitalizacionComplicaciones
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function SubclasificacionDiagnosticosSeleccionarDxInterconsultas() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New SubclasificacionDiagnosticos
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set SubclasificacionDiagnosticosSeleccionarDxInterconsultas = oTabla.SeleccionarDxInterconsultas
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function TiposCondicionPacienteSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposCondicionPaciente
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposCondicionPacienteSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function TiposCondicionPacienteCondicionAlEstablecimiento(idPaciente As Long, FechaIngreso As Date, idAtencion As Long) As Integer
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposCondicionPaciente
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    TiposCondicionPacienteCondicionAlEstablecimiento = oTabla.CondicionAlEstablecimiento(idPaciente, FechaIngreso, idAtencion)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function TiposCondicionPacienteCondicionAlServicio(idPaciente As Long, FechaIngreso As Date, IdServicio As Long, idAtencion As Long) As Integer
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposCondicionPaciente
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    TiposCondicionPacienteCondicionAlServicio = oTabla.CondicionAlServicio(idPaciente, FechaIngreso, IdServicio, idAtencion)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


Function ProcedimientosSeleccionarPorCodigoCPT(sCodigo As String) As doProcedimiento
Dim oConexion As New ADODB.Connection
Dim oTabla As New Procedimientos
Dim doProcedimiento As New doProcedimiento
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    doProcedimiento.CodigoCPT2004 = sCodigo
    If oTabla.SeleccionarPorCodigoCPT(doProcedimiento) Then
        Set ProcedimientosSeleccionarPorCodigoCPT = doProcedimiento
    Else
        Set ProcedimientosSeleccionarPorCodigoCPT = Nothing
    End If
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function TiposCondicionRNSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposCondicionRecienNacido
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposCondicionRNSeleccionarTodos = oTabla.SeleccionarTodos
    
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function SubclasificacionDiagnosticosSeleccionarIdPorCodigoYClasificacion(sCodigo As String, lIdClasificacion As Long) As Long
Dim oConexion As New ADODB.Connection
Dim oTabla As New SubclasificacionDiagnosticos
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    SubclasificacionDiagnosticosSeleccionarIdPorCodigoYClasificacion = oTabla.SeleccionarIdPorCodigoYClasificacion(sCodigo, lIdClasificacion)
    
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function TipoGravedadAtencionSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposGravedadAtencion
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TipoGravedadAtencionSeleccionarTodos = oTabla.SeleccionarTodos
    
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function TiposEdadSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposEdad
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposEdadSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function TiposEdadSeleccionarPorCodigo(sCodigo As String) As DOTipoEdad
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposEdad
Dim oDOTipoEdad As New DOTipoEdad
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    oDOTipoEdad.Codigo = sCodigo
    If oTabla.SeleccionarPorCodigo(oDOTipoEdad) Then
        Set TiposEdadSeleccionarPorCodigo = oDOTipoEdad
    Else
        Set TiposEdadSeleccionarPorCodigo = Nothing
    End If
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'-------------------------------------------------------------------------------------
'                                       CatalogoBienesInsumos
'-------------------------------------------------------------------------------------
Function CatalogoBienesInsumosAgregar(oDOCatalogoBienesInsumos As DOCatalogoBienesInsumos, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreProducto As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCatalogoBienesInsumos As New CatalogoBienesInsumos
    CatalogoBienesInsumosAgregar = False
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    Set oCatalogoBienesInsumos.Conexion = oConexion
    If oCatalogoBienesInsumos.Insertar(oDOCatalogoBienesInsumos) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCatalogoBienesInsumos.IdUsuarioAuditoria, "A", oDOCatalogoBienesInsumos.idProducto, "FactCatalogoBienesInsumos", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreProducto)             'ListBarItems.idListItem
        CatalogoBienesInsumosAgregar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oCatalogoBienesInsumos.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCatalogoBienesInsumos = Nothing
End Function

Function CatalogoBienesInsumosModificar(oDOCatalogoBienesInsumos As DOCatalogoBienesInsumos, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreProducto As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCatalogoBienesInsumos As New CatalogoBienesInsumos
    CatalogoBienesInsumosModificar = False
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    Set oCatalogoBienesInsumos.Conexion = oConexion
    If oCatalogoBienesInsumos.modificar(oDOCatalogoBienesInsumos) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCatalogoBienesInsumos.IdUsuarioAuditoria, "M", oDOCatalogoBienesInsumos.idProducto, "FactCatalogoBienesInsumos", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreProducto)             'ListBarItems.idListItem
        CatalogoBienesInsumosModificar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oCatalogoBienesInsumos.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCatalogoBienesInsumos = Nothing
    
End Function

Function CatalogoBienesInsumosEliminar(oDOCatalogoBienesInsumos As DOCatalogoBienesInsumos, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreProducto As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCatalogoBienesInsumos As New CatalogoBienesInsumos

    CatalogoBienesInsumosEliminar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    Set oCatalogoBienesInsumos.Conexion = oConexion
    If oCatalogoBienesInsumos.Eliminar(oDOCatalogoBienesInsumos) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCatalogoBienesInsumos.IdUsuarioAuditoria, "E", oDOCatalogoBienesInsumos.idProducto, "FactCatalogoBienesInsumos", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreProducto)             'ListBarItems.idListItem
        CatalogoBienesInsumosEliminar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oCatalogoBienesInsumos.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCatalogoBienesInsumos = Nothing
    
End Function

Function CatalogoBienesInsumosSeleccionarPorId(idProducto As Long) As DOCatalogoBienesInsumos
Dim oConexion As New ADODB.Connection
Dim oProducto As New CatalogoBienesInsumos
Dim oDOBienesInsumos As New DOCatalogoBienesInsumos
    
    Set CatalogoBienesInsumosSeleccionarPorId = Nothing
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oProducto.Conexion = oConexion
    oDOBienesInsumos.idProducto = idProducto
    
    If oProducto.SeleccionarPorId(oDOBienesInsumos) Then
        Set CatalogoBienesInsumosSeleccionarPorId = oDOBienesInsumos
    End If
    oConexion.Close
    ms_MensajeError = oProducto.MensajeError
    Set oConexion = Nothing
    Set oProducto = Nothing
End Function

Function CatalogoBienesInsumosFiltrar(doBienesInsumos As DOCatalogoBienesInsumos, ml_TipoCatalogo As Long) As ADODB.Recordset

Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoBienesInsumos
    
    oConexion.Open SIGHEntidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    
    Select Case ml_TipoCatalogo
    Case 0
        Set CatalogoBienesInsumosFiltrar = oTabla.FiltrarCatalogoBase(doBienesInsumos)
    Case Else
        Set CatalogoBienesInsumosFiltrar = oTabla.FiltrarCatalogo(doBienesInsumos, ml_TipoCatalogo)
    End Select
    
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function CatalogoBienesInsumosFiltrarDEBB(doBienesInsumos As DOCatalogoBienesInsumos, ml_TipoCatalogo As Long) As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New CatalogoBienesInsumos
    
    oConexion.Open SIGHEntidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    
    Select Case ml_TipoCatalogo
    Case 0
        Set CatalogoBienesInsumosFiltrarDEBB = oTabla.FiltrarCatalogoBase(doBienesInsumos)
    Case Else
        Set CatalogoBienesInsumosFiltrarDEBB = oTabla.FiltrarCatalogodebb(doBienesInsumos, ml_TipoCatalogo)
    End Select
    
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing

End Function

'-------------------------------------------------------------------------------------
'                                       InsumosGrupoFarmacologico
'-------------------------------------------------------------------------------------
Function InsumosGrupoFarmacologicoSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New InsumosGrupoFarmacologico
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set InsumosGrupoFarmacologicoSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


'-------------------------------------------------------------------------------------
'                                       InsumosSubGrupoFarmacologico
'-------------------------------------------------------------------------------------
Function InsumosSubGrupoFarmacologicoSeleccionarPorGrupo(lIdGrupoFarmacologico As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New InsumosSubGrupoFarmacologico
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set InsumosSubGrupoFarmacologicoSeleccionarPorGrupo = oTabla.SeleccionarPorGrupo(lIdGrupoFarmacologico)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
     Set oConexion = Nothing
    Set oTabla = Nothing
   
End Function

'-------------------------------------------------------------------------------------
'                                       CentrosCosto
'-------------------------------------------------------------------------------------
Function CentrosCostoAgregar(oDOCentrosCosto As DOCentrosCosto, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreCC As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCentrosCosto As New CentrosCosto
    
    CentrosCostoAgregar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    Set oCentrosCosto.Conexion = oConexion
    If oCentrosCosto.Insertar(oDOCentrosCosto) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCentrosCosto.IdUsuarioAuditoria, "A", oDOCentrosCosto.IdCentroCosto, "CentrosCosto", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreCC)         'ListBarItems.idListItem
        CentrosCostoAgregar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oCentrosCosto.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCentrosCosto = Nothing
End Function

Function CentrosCostoModificar(oDOCentrosCosto As DOCentrosCosto, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreCC As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCentrosCosto As New CentrosCosto
    
    CentrosCostoModificar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    Set oCentrosCosto.Conexion = oConexion
    If oCentrosCosto.modificar(oDOCentrosCosto) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCentrosCosto.IdUsuarioAuditoria, "M", oDOCentrosCosto.IdCentroCosto, "CentrosCosto", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreCC)         'ListBarItems.idListItem
        CentrosCostoModificar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oCentrosCosto.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCentrosCosto = Nothing
    
End Function

Function CentrosCostoEliminar(oDOCentrosCosto As DOCentrosCosto, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreCC As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCentrosCosto As New CentrosCosto

    CentrosCostoEliminar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    Set oCentrosCosto.Conexion = oConexion
    If oCentrosCosto.Eliminar(oDOCentrosCosto) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCentrosCosto.IdUsuarioAuditoria, "E", oDOCentrosCosto.IdCentroCosto, "CentrosCosto", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreCC)         'ListBarItems.idListItem
        CentrosCostoEliminar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oCentrosCosto.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCentrosCosto = Nothing
    
End Function

Function CentrosCostoSeleccionarPorId(IdCentroCosto As Long) As DOCentrosCosto
Dim oConexion As New ADODB.Connection
Dim oProducto As New CentrosCosto
Dim oDOCentrosCosto As New DOCentrosCosto
    
    Set CentrosCostoSeleccionarPorId = Nothing
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oProducto.Conexion = oConexion
    oDOCentrosCosto.IdCentroCosto = IdCentroCosto
    
    If oProducto.SeleccionarPorId(oDOCentrosCosto) Then
        Set CentrosCostoSeleccionarPorId = oDOCentrosCosto
    End If
    oConexion.Close

    ms_MensajeError = oProducto.MensajeError
    Set oConexion = Nothing
    Set oProducto = Nothing

End Function
Function CentrosCostoFiltrar(oDOCentrosCosto As DOCentrosCosto) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CentrosCosto
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set CentrosCostoFiltrar = oTabla.Filtrar(oDOCentrosCosto)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing

End Function




'-------------------------------------------------------------------------------------
'                                       Partidas Presupuestales
'-------------------------------------------------------------------------------------
Function PartidasPresupuestalesAgregar(oDOPartidasPresupuestales As DOPartidasPresupuestales, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDpartida As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPartidasPresupuestales As New PartidasPresupuestales
    
    PartidasPresupuestalesAgregar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    Set oPartidasPresupuestales.Conexion = oConexion
    If oPartidasPresupuestales.Insertar(oDOPartidasPresupuestales) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOPartidasPresupuestales.IdUsuarioAuditoria, "A", oDOPartidasPresupuestales.IdPartidaPresupuestal, "FactPartidasPresupuestales", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDpartida)            'ListBarItems.idListItem
        PartidasPresupuestalesAgregar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oPartidasPresupuestales.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPartidasPresupuestales = Nothing
End Function

Function PartidasPresupuestalesModificar(oDOPartidasPresupuestales As DOPartidasPresupuestales, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDpartida As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPartidasPresupuestales As New PartidasPresupuestales
    
    PartidasPresupuestalesModificar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    Set oPartidasPresupuestales.Conexion = oConexion
    If oPartidasPresupuestales.modificar(oDOPartidasPresupuestales) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOPartidasPresupuestales.IdUsuarioAuditoria, "M", oDOPartidasPresupuestales.IdPartidaPresupuestal, "FactPartidasPresupuestales", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDpartida)            'ListBarItems.idListItem
        PartidasPresupuestalesModificar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oPartidasPresupuestales.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPartidasPresupuestales = Nothing
    
End Function

Function PartidasPresupuestalesEliminar(oDOPartidasPresupuestales As DOPartidasPresupuestales, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDpartida As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPartidasPresupuestales As New PartidasPresupuestales

    PartidasPresupuestalesEliminar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    Set oPartidasPresupuestales.Conexion = oConexion
    If oPartidasPresupuestales.Eliminar(oDOPartidasPresupuestales) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOPartidasPresupuestales.IdUsuarioAuditoria, "E", oDOPartidasPresupuestales.IdPartidaPresupuestal, "FactPartidasPresupuestales", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDpartida)            'ListBarItems.idListItem
        PartidasPresupuestalesEliminar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oPartidasPresupuestales.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPartidasPresupuestales = Nothing
    
End Function

Function PartidasPresupuestalesSeleccionarPorId(IdPartidasPresupuestales As Long) As DOPartidasPresupuestales
Dim oConexion As New ADODB.Connection
Dim oProducto As New PartidasPresupuestales
Dim oDOPartidasPresupuestales As New DOPartidasPresupuestales
    
    Set PartidasPresupuestalesSeleccionarPorId = Nothing
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oProducto.Conexion = oConexion
    oDOPartidasPresupuestales.IdPartidaPresupuestal = IdPartidasPresupuestales
    
    If oProducto.SeleccionarPorId(oDOPartidasPresupuestales) Then
        Set PartidasPresupuestalesSeleccionarPorId = oDOPartidasPresupuestales
    End If
    oConexion.Close

    ms_MensajeError = oProducto.MensajeError
    Set oConexion = Nothing
    Set oProducto = Nothing

End Function


Function PartidasPresupuestalesFiltrar(oDoPartidaPresupuestales As DOPartidasPresupuestales) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New PartidasPresupuestales
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set PartidasPresupuestalesFiltrar = oTabla.Filtrar(oDoPartidaPresupuestales)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing


End Function

Function CentrosCostoSeleccionarTodos(Optional oConexion1 As Connection) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CentrosCosto
    
    If oConexion1 Is Nothing Then
       oConexion.CommandTimeout = 300
       oConexion.CursorLocation = adUseClient
       oConexion.Open SIGHEntidades.CadenaConexion
       Set oTabla.Conexion = oConexion
    Else
       Set oTabla.Conexion = oConexion1
    End If
    Set CentrosCostoSeleccionarTodos = oTabla.SeleccionarTodos
    
    ms_MensajeError = oTabla.MensajeError
       Set oConexion = Nothing
    Set oTabla = Nothing

End Function


'-------------------------------------------------------------------------------------
'                                       PartidasPresupuestales
'-------------------------------------------------------------------------------------
Function PartidasPresupuestalesSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New PartidasPresupuestales
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set PartidasPresupuestalesSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function



'-------------------------------------------------------------------------------------
'                                       CatalogoServicios
'-------------------------------------------------------------------------------------

'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function CatalogoServiciosAgregar(oDOCatalogoServicio As DOCatalogoServicio, oRsPuntoCarga As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreServicio As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCatalogoServicios As New CatalogoServicios
Dim oFactCatalogoServiciosPtos As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    CatalogoServiciosAgregar = False
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    Set oCatalogoServicios.Conexion = oConexion
    If oCatalogoServicios.Insertar(oDOCatalogoServicio) Then
       If oRsPuntoCarga.RecordCount > 0 Then
          oRsPuntoCarga.MoveFirst
          Do While Not oRsPuntoCarga.EOF
             Set oCommand = Nothing
             Set oParameter = Nothing
             With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "FactCatalogoServiciosPtosAgregar"
                Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, oRsPuntoCarga.Fields!IdPuntoCarga): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, oDOCatalogoServicio.idProducto): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@EsPreVenta", adBoolean, adParamInput, 0, IIf(oRsPuntoCarga.Fields!EsPreVenta = True, 1, 0)): .Parameters.Append oParameter
                .Execute
             End With
             '
             oRsPuntoCarga.MoveNext
          Loop
       End If
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCatalogoServicio.IdUsuarioAuditoria, "A", oDOCatalogoServicio.idProducto, "FactCatalogoServicios", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreServicio)          'ListBarItems.idListItem
        CatalogoServiciosAgregar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oCatalogoServicios.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCatalogoServicios = Nothing
    Set oFactCatalogoServiciosPtos = Nothing
    Set oCommand = Nothing
End Function

'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function CatalogoServiciosModificar(oDOCatalogoServicio As DOCatalogoServicio, oRsPuntoCarga As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreServicio As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCatalogoServicios As New CatalogoServicios
Dim oFactCatalogoServiciosPtos As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    CatalogoServiciosModificar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    
    oConexion.BeginTrans
    Set oCatalogoServicios.Conexion = oConexion
    If oCatalogoServicios.modificar(oDOCatalogoServicio) Then
        With oCommand
           .CommandType = adCmdStoredProc
           Set .ActiveConnection = oConexion
           .CommandTimeout = 150
           .CommandText = "FactCatalogoServiciosPtosEliminar"
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, oDOCatalogoServicio.idProducto): .Parameters.Append oParameter
           .Execute
        End With
       
       If oRsPuntoCarga.RecordCount > 0 Then
          oRsPuntoCarga.MoveFirst
          Do While Not oRsPuntoCarga.EOF
             Set oCommand = Nothing
             Set oParameter = Nothing
             With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "FactCatalogoServiciosPtosAgregar"
                Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, oRsPuntoCarga.Fields!IdPuntoCarga): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, oDOCatalogoServicio.idProducto): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@EsPreVenta", adBoolean, adParamInput, 0, IIf(oRsPuntoCarga.Fields!EsPreVenta = True, 1, 0)): .Parameters.Append oParameter
                .Execute
             End With
             '
             oRsPuntoCarga.MoveNext
          Loop
       End If
       Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCatalogoServicio.IdUsuarioAuditoria, "M", oDOCatalogoServicio.idProducto, "FactCatalogoServicios", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreServicio)          'ListBarItems.idListItem
       CatalogoServiciosModificar = True
       oConexion.CommitTrans
    Else
        ms_MensajeError = oCatalogoServicios.MensajeError
        oConexion.RollbackTrans
    End If
    '
    
    oConexion.Close
    Set oConexion = Nothing
    Set oCatalogoServicios = Nothing
    Set oFactCatalogoServiciosPtos = Nothing
    Set oCommand = Nothing
    
End Function

Function CatalogoServiciosEliminar(oDOCatalogoServicio As DOCatalogoServicio, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreServicio As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oCatalogoServicios As New CatalogoServicios
Dim oFactCatalogoServiciosPtos As New Recordset
    CatalogoServiciosEliminar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    Set oCatalogoServicios.Conexion = oConexion
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = oConexion
       .CommandTimeout = 150
       .CommandText = "FactCatalogoServiciosPtosEliminar"
       Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, oDOCatalogoServicio.idProducto): .Parameters.Append oParameter
       .Execute
    End With
    Set oParameter = Nothing
    
    If oCatalogoServicios.Eliminar(oDOCatalogoServicio) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCatalogoServicio.IdUsuarioAuditoria, "E", oDOCatalogoServicio.idProducto, "FactCatalogoServicios", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreServicio)          'ListBarItems.idListItem
        
        CatalogoServiciosEliminar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oCatalogoServicios.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCatalogoServicios = Nothing
    Set oFactCatalogoServiciosPtos = Nothing
End Function

Function CatalogoServiciosSeleccionarPorId(idProducto As Long) As DOCatalogoServicio
Dim oConexion As New ADODB.Connection
Dim oProducto As New CatalogoServicios
Dim oDOBienesInsumos As New DOCatalogoServicio
    
    Set CatalogoServiciosSeleccionarPorId = Nothing
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oProducto.Conexion = oConexion
    oDOBienesInsumos.idProducto = idProducto
    
    If oProducto.SeleccionarPorId(oDOBienesInsumos) Then
        Set CatalogoServiciosSeleccionarPorId = oDOBienesInsumos
    End If
    oConexion.Close

    ms_MensajeError = oProducto.MensajeError
    Set oConexion = Nothing
    Set oProducto = Nothing

End Function

Function CatalogoServiciosFiltrar(oDOCatalogoServicio As DOCatalogoServicio, ml_TipoCatalogo As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoServicios
    
    oConexion.Open SIGHEntidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    
    Select Case ml_TipoCatalogo
    Case 0
        Set CatalogoServiciosFiltrar = oTabla.FiltrarCatalogoBase(oDOCatalogoServicio)
    Case Else
        Set CatalogoServiciosFiltrar = oTabla.FiltrarCatalogo(oDOCatalogoServicio, ml_TipoCatalogo)
    End Select
    
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function CatalogoServiciosFiltrarDEBB(oDOCatalogoServicio As DOCatalogoServicio, ml_TipoCatalogo As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoServicios
    
    oConexion.Open SIGHEntidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    
    Select Case ml_TipoCatalogo
    Case 0
        Set CatalogoServiciosFiltrarDEBB = oTabla.FiltrarCatalogoBase(oDOCatalogoServicio)
    Case Else
        Set CatalogoServiciosFiltrarDEBB = oTabla.FiltrarCatalogodebb(oDOCatalogoServicio, ml_TipoCatalogo)
    End Select
    
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


Function CatalogoServiciosGrupoSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoServiciosGrupo
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set CatalogoServiciosGrupoSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function CatalogoServiciosSubGrupoSeleccionarPorGrupo(lIdGrupo As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoServiciosSubGrupo
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set CatalogoServiciosSubGrupoSeleccionarPorGrupo = oTabla.SeleccionarPorGrupo(lIdGrupo)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function CatalogoServiciosSeccionSeleccionarPorSubGrupo(lIdSubGrupo As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoServiciosSeccion
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set CatalogoServiciosSeccionSeleccionarPorSubGrupo = oTabla.SeleccionarPorSubGrupo(lIdSubGrupo)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function CatalogoServiciosSubSeccionSeleccionarPorSeccion(lIdServicioSeccion As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoServiciosSubSeccion
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set CatalogoServiciosSubSeccionSeleccionarPorSeccion = oTabla.SeleccionarPorSeccion(lIdServicioSeccion)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


Function SeleccionarParametrosPorTipo(tipo As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Parametros
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set SeleccionarParametrosPorTipo = oTabla.SeleccionarTipo(tipo)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function SeleccionarPuntosDeCarga() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FactPuntosCarga
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    
    
    Set oTabla.Conexion = oConexion
    Set SeleccionarPuntosDeCarga = oTabla.SeleccionarTodos()
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function SeleccionarPuntosDeCargaSegunFiltro(lcFiltro As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FactPuntosCarga
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set SeleccionarPuntosDeCargaSegunFiltro = oTabla.SeleccionarSegunFiltro(lcFiltro)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function



Function ParametrosSeleccionarValorIntPorTipoYCodigo(sTipo As String, sCodigo As String) As Long
Dim oConexion As New ADODB.Connection
Dim oTabla As New Parametros
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    
    ParametrosSeleccionarValorIntPorTipoYCodigo = oTabla.ObtenerValorIntPorTipoYCodigo(sTipo, sCodigo)
    
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


'***************Retorna Permiso del Usuario Galenhos para poder usar solo SIS o SOAT
'***************o EXONERAR o CONVENIO_FOSPOLIS
'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function DevuelveSubAreaDondeLaboraElUsuarioDelSistema(lnIdAreaLaboraUsuarioSistema As sghAreasLaboraEmpleado, lnIdUsuarioSistema As Long) As Recordset
    Dim oRsTmp As New ADODB.Recordset
     
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set DevuelveSubAreaDondeLaboraElUsuarioDelSistema = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "DevuelveSubAreaDondeLaboraElUsuarioDelSistema"
            Set oParameter = .CreateParameter("@idLaboraArea", adInteger, adParamInput, 0, lnIdAreaLaboraUsuarioSistema): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@idEmpleado", adInteger, adParamInput, 0, lnIdUsuarioSistema): .Parameters.Append oParameter
            Set oRsTmp = .Execute
            Set oRsTmp.ActiveConnection = Nothing
       End With
       Set DevuelveSubAreaDondeLaboraElUsuarioDelSistema = oRsTmp
       oConexion.Close
       Set oRsTmp = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing

End Function


Function DevuelveNombreMedicamentoOServicioSegunId(lnIdBienOservicio As Long, lnBienOServicio As sghTipoProducto) As String
    Dim oDOCatalogoBienesInsumos As New DOCatalogoBienesInsumos
    Dim oCatalogoBienesInsumos As New CatalogoBienesInsumos
    Dim oDOCatalogoServicio As New DOCatalogoServicio
    Dim oCatalogoServicios As New CatalogoServicios
    Dim oConexion As New ADODB.Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    DevuelveNombreMedicamentoOServicioSegunId = ""
    If lnBienOServicio = sghbien Then
        Set oCatalogoBienesInsumos.Conexion = oConexion
        oDOCatalogoBienesInsumos.idProducto = lnIdBienOservicio
        If oCatalogoBienesInsumos.SeleccionarPorId(oDOCatalogoBienesInsumos) Then
           DevuelveNombreMedicamentoOServicioSegunId = oDOCatalogoBienesInsumos.Codigo & " " & oDOCatalogoBienesInsumos.nombre
        End If
    Else
        Set oCatalogoServicios.Conexion = oConexion
        oDOCatalogoServicio.idProducto = lnIdBienOservicio
        If oCatalogoServicios.SeleccionarPorId(oDOCatalogoServicio) Then
           DevuelveNombreMedicamentoOServicioSegunId = oDOCatalogoServicio.Codigo & " " & oDOCatalogoServicio.nombre
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oDOCatalogoBienesInsumos = Nothing
    Set oCatalogoBienesInsumos = Nothing
    Set oDOCatalogoServicio = Nothing
    Set oCatalogoServicios = Nothing
End Function


Function FactPuntosCargaSeleccionarPorId(lnIdPuntoCarga As Long, Optional oConexion1 As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactPuntosCargaSeleccionarPorId = Nothing
    ms_MensajeError = ""
    If oConexion1 Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open SIGHEntidades.CadenaConexion
        
    End If

    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "FactPuntosCargaSeleccionarPorId"
        Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set FactPuntosCargaSeleccionarPorId = oRecordset
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function OPCsDevuelveCodigoOPCporCodigoCPT(lcCodigoCpt As String) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "OpcsSeleccionarPorCPT"
        Set oParameter = .CreateParameter("@codigoCpt", adVarChar, adParamInput, 7, Left(lcCodigoCpt, 7)): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   OPCsDevuelveCodigoOPCporCodigoCPT = ""
   If oRecordset.RecordCount > 0 Then
      If Not IsNull(oRecordset.Fields!opcsCodigo) Then
         OPCsDevuelveCodigoOPCporCodigoCPT = Trim(oRecordset.Fields!opcsCodigo)
      End If
   End If
   oRecordset.Close
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
Resume
End Function

Function CatalogoBienesInsumosSeleccionarPorCodigo(lcCodigo As String, Optional oConexion As Connection) As ADODB.Recordset
    Dim oTabla As New CatalogoBienesInsumos
    Set CatalogoBienesInsumosSeleccionarPorCodigo = oTabla.SeleccionarPorCodigo(lcCodigo, oConexion)
End Function

Function CatalogoServiciosSeleccionarPorCodigo(lcCodigo As String, Optional oConexion As Connection) As ADODB.Recordset
    Dim oTabla As New CatalogoServicios
    Set CatalogoServiciosSeleccionarPorCodigo = oTabla.SeleccionarPorCodigo(lcCodigo, oConexion)
End Function


Public Sub DevuelveApellidosYnombresDesdeRazonSocial(lcRazonSocial As String, ByRef lcApellidoPaterno As String, ByRef lcApellidoMaterno As String, ByRef lcPrimerNombre As String)
    Dim lnLen As Integer
    Dim lnPos As Integer
    Dim mo_Teclado As New SIGHEntidades.Teclado
    lnPos = 1
    lcApellidoPaterno = ""
    lcApellidoMaterno = ""
    lcPrimerNombre = ""
    If lcRazonSocial <> "" Then
        For lnLen = 1 To Len(lcRazonSocial)
            If Mid(lcRazonSocial, lnLen, 1) = " " Then
               lnPos = lnPos + 1
            Else
               Select Case lnPos
               Case 1
                    lcApellidoPaterno = lcApellidoPaterno & Mid(lcRazonSocial, lnLen, 1)
               Case 2
                    lcApellidoMaterno = lcApellidoMaterno & Mid(lcRazonSocial, lnLen, 1)
               Case 3
                    lcPrimerNombre = lcPrimerNombre & Mid(lcRazonSocial, lnLen, 1)
               End Select
            End If
        Next
        lcApellidoPaterno = Left(lcApellidoPaterno, 20)
        lcApellidoMaterno = Left(lcApellidoMaterno, 20)
        lcPrimerNombre = Left(lcPrimerNombre, 20)
        lcApellidoPaterno = mo_Teclado.CapitalizarNombres(lcApellidoPaterno)
        lcApellidoMaterno = mo_Teclado.CapitalizarNombres(lcApellidoMaterno)
        lcPrimerNombre = mo_Teclado.CapitalizarNombres(lcPrimerNombre)
    End If
End Sub

Function DevuelveEspecialidadDelServicio(lnIdServicio As Long) As Long
      Dim oServicios As New Servicios
      Dim oConexion As New Connection
      oConexion.Open SIGHEntidades.CadenaConexion
      oConexion.CursorLocation = adUseClient
      Set oServicios.Conexion = oConexion
      DevuelveEspecialidadDelServicio = oServicios.DevuelveEspecialidadDelServicio(lnIdServicio)
      oConexion.Close
      Set oConexion = Nothing
      Set oServicios = Nothing
End Function

Function FuentesFinanciamientoSegunFiltro(lcFiltro As String) As Recordset
      Dim oFuentesFinanciamiento As New FuentesFinanciamiento
      Dim oConexion As New Connection
      oConexion.Open SIGHEntidades.CadenaConexion
      oConexion.CursorLocation = adUseClient
      Set oFuentesFinanciamiento.Conexion = oConexion
      Set FuentesFinanciamientoSegunFiltro = oFuentesFinanciamiento.FuentesFinanciamientoSegunFiltro(lcFiltro)
      oConexion.Close
      Set oConexion = Nothing
End Function

Function FuentesFinanciamientoDevuelveDescripcion(lnIdFuenteFinanciamiento As Long) As String
        FuentesFinanciamientoDevuelveDescripcion = ""
        Dim oRsTmp As New Recordset
        Set oRsTmp = FuentesFinanciamientoSegunFiltro("idFuenteFinanciamiento=" & lnIdFuenteFinanciamiento)
        If oRsTmp.RecordCount > 0 Then
           FuentesFinanciamientoDevuelveDescripcion = oRsTmp!Descripcion
        End If
        oRsTmp.Close
        Set oRsTmp = Nothing
End Function

Function TiposFinanciamientoSegunFiltro(lcFiltro As String) As Recordset
      Dim oTiposFinanciamiento As New TiposFinanciamiento
      Dim oConexion As New Connection
      oConexion.Open SIGHEntidades.CadenaConexion
      oConexion.CursorLocation = adUseClient
      Set oTiposFinanciamiento.Conexion = oConexion
      Set TiposFinanciamientoSegunFiltro = oTiposFinanciamiento.TiposFinanciamientoSegunFiltro(lcFiltro)
      oConexion.Close
      Set oConexion = Nothing
End Function

'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Sub LlenaListBoxConTablaMensajesEnVentana(ctrListBox As Object, lcVentana As String)
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "LlenaListBoxConTablaMensajesEnVentana"
            Set oParameter = .CreateParameter("@Ventana", adVarChar, adParamInput, 50, lcVentana): .Parameters.Append oParameter
            Set oRsTmp = .Execute
            Set oRsTmp.ActiveConnection = Nothing
       End With
      
    ctrListBox.Clear
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          ctrListBox.AddItem oRsTmp.Fields!mensajeAusuario
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    oConexion.Close
    Set oRsTmp = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    
End Sub

'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Sub LlenaDataComboTipoSalidaBienSegunAlmacen(ctrDataCombo As Object, lnIdAlmacen As Long, _
                                              ByRef lcidTipoSuministro As String)
    Dim oRsTmp As New Recordset
    Dim mo_Formulario As New SIGHEntidades.Formulario
    
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
 
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "LlenaDataComboTipoSalidaBienSegunAlmacen"
            Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
            Set oRsTmp = .Execute
            Set oRsTmp.ActiveConnection = Nothing
       End With
    
    
    ctrDataCombo.Clear
    mo_Formulario.HabilitarDeshabilitar ctrDataCombo, False
    lcidTipoSuministro = ""
    If oRsTmp.RecordCount > 0 Then
       lcidTipoSuministro = oRsTmp.Fields!idTipoSuministro
       If oRsTmp.Fields!idTipoSuministro = "02" Then
          ctrDataCombo.AddItem ""
          ctrDataCombo.AddItem ""
          ctrDataCombo.AddItem ""
          ctrDataCombo.AddItem ""
          ctrDataCombo.AddItem "Donaciones"
          ctrDataCombo.ListIndex = 4
       Else
          ctrDataCombo.AddItem ""
          ctrDataCombo.AddItem "Venta"
          ctrDataCombo.AddItem "InterSanit"
          mo_Formulario.HabilitarDeshabilitar ctrDataCombo, True
       End If
       
    End If
    oRsTmp.Close
    oConexion.Close
    Set oRsTmp = Nothing
    Set mo_Formulario = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing


End Sub

'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function SunasaTiposAfiliacionSeleccionarTodos() As Recordset
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
        Set SunasaTiposAfiliacionSeleccionarTodos = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "SunasaTiposAfiliacionSeleccionarTodos"
            Set oRsTmp = .Execute
            Set oRsTmp.ActiveConnection = Nothing
       End With
       Set SunasaTiposAfiliacionSeleccionarTodos = oRsTmp
       oConexion.Close
       Set oRsTmp = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
End Function

'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function SunasaTiposRegimenSeleccionarTodos() As Recordset
    Dim oRsTmp As New Recordset
    
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
        Set SunasaTiposRegimenSeleccionarTodos = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "SunasaTiposRegimenSeleccionarTodos"
            Set oRsTmp = .Execute
            Set oRsTmp.ActiveConnection = Nothing
       End With
       Set SunasaTiposRegimenSeleccionarTodos = oRsTmp
       oConexion.Close
       Set oRsTmp = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
    
End Function

'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function SunasaTiposParentescoSeleccionarTodos() As Recordset
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
        Set SunasaTiposParentescoSeleccionarTodos = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "SunasaTiposParentescoSeleccionarTodos"
            Set oRsTmp = .Execute
            Set oRsTmp.ActiveConnection = Nothing
       End With
       Set SunasaTiposParentescoSeleccionarTodos = oRsTmp
       oConexion.Close
       Set oRsTmp = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
     
End Function
'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function SunasaTiposOperacionSeleccionarTodos() As Recordset
    Dim oRsTmp As New Recordset
    
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
        Set SunasaTiposOperacionSeleccionarTodos = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "SunasaTiposOperacionSeleccionarTodos"
            Set oRsTmp = .Execute
            Set oRsTmp.ActiveConnection = Nothing
       End With
       Set SunasaTiposOperacionSeleccionarTodos = oRsTmp
       oConexion.Close
       Set oRsTmp = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
    
End Function

'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Sub ActualizaPreciosParaFarmacia(mrs_Precios As Recordset, lnIdProducto As Long)
        On Error GoTo ErrActPr
        'Graba Precios
        Dim oRsTmp As New Recordset
        Dim lcSql As String
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        Dim oDoFinanciamientoCatalogoBien As New DoFinanciamientoCatalogoBien, oFinanciamientoCatalogoBien As New FinanciamientoCatalogoBien
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        If mrs_Precios.RecordCount > 0 Then
           Set oFinanciamientoCatalogoBien.Conexion = oConexion
           mrs_Precios.MoveFirst
           Do While Not mrs_Precios.EOF
              oDoFinanciamientoCatalogoBien.idProducto = lnIdProducto
              oDoFinanciamientoCatalogoBien.IdTipoFinanciamiento = mrs_Precios.Fields!IdTipoFinanciamiento
              If oFinanciamientoCatalogoBien.SeleccionarPorProductoYTipoFinanciamiento(oDoFinanciamientoCatalogoBien) = True Then
                 If mrs_Precios.Fields!PrecioVenta > 0 Then
                    oDoFinanciamientoCatalogoBien.PrecioUnitario = mrs_Precios.Fields!PrecioVenta
                    If oFinanciamientoCatalogoBien.modificar(oDoFinanciamientoCatalogoBien) = True Then
                    End If
                 Else
                    If oFinanciamientoCatalogoBien.Eliminar(oDoFinanciamientoCatalogoBien) = True Then
                    End If
                 End If
              Else
                 oDoFinanciamientoCatalogoBien.idProducto = lnIdProducto
                 oDoFinanciamientoCatalogoBien.IdTipoFinanciamiento = mrs_Precios.Fields!IdTipoFinanciamiento
                 oDoFinanciamientoCatalogoBien.Activo = 1
                 oDoFinanciamientoCatalogoBien.PrecioUnitario = mrs_Precios.Fields!PrecioVenta
                 If oFinanciamientoCatalogoBien.Insertar(oDoFinanciamientoCatalogoBien) = True Then
                 End If
              End If
              '
              mrs_Precios.MoveNext
           Loop
        End If
        oConexion.Close
        Set oRsTmp = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
        Set oDoFinanciamientoCatalogoBien = Nothing
        Set oFinanciamientoCatalogoBien = Nothing
        Exit Sub
ErrActPr:
        MsgBox Err.Description
        Resume
End Sub
'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function PerinatalCatalogoCie10SeleccionarPorIdModulo(lnIdModulo As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set PerinatalCatalogoCie10SeleccionarPorIdModulo = Nothing
     
'
      With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PerinatalCatalogoCie10SeleccionarPorIdModulo"
        Set oParameter = .CreateParameter("@IdModulo", adInteger, adParamInput, 0, lnIdModulo): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set PerinatalCatalogoCie10SeleccionarPorIdModulo = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function PerinatalCatalogoCptSeleccionarPorIdModulo(lnIdModulo As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set PerinatalCatalogoCptSeleccionarPorIdModulo = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PerinatalCatalogoCptSeleccionarPorIdModulo"
        Set oParameter = .CreateParameter("@IdModulo", adInteger, adParamInput, 0, lnIdModulo): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set PerinatalCatalogoCptSeleccionarPorIdModulo = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function PerinatalCatalogoMedicamentosSeleccionarPorIdModulo(lnIdModulo As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set PerinatalCatalogoMedicamentosSeleccionarPorIdModulo = Nothing

   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PerinatalCatalogoMedicamentosSeleccionarPorIdModulo"
        Set oParameter = .CreateParameter("@IdModulo", adInteger, adParamInput, 0, lnIdModulo): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set PerinatalCatalogoMedicamentosSeleccionarPorIdModulo = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function PerinatalAtencionCptSeleccionarPorIdPerinatalAtencion(mo_idPerinatalAtencion As Long, oConexion As Connection) As Recordset
      Dim oPerinatalAtencioncpt As New PerinatalAtencionCpt
      Set oPerinatalAtencioncpt.Conexion = oConexion
      Set PerinatalAtencionCptSeleccionarPorIdPerinatalAtencion = oPerinatalAtencioncpt.PerinatalAtencionCptSeleccionarPorIdPerinatalAtencion(mo_idPerinatalAtencion)
      Set oPerinatalAtencioncpt = Nothing
End Function
Function PerinatalAtencionMedicamentoSeleccionarPorIdPerinatalAtencion(mo_idPerinatalAtencion As Long, oConexion As Connection) As Recordset
      Dim oPerinatalAtencionMed As New PerinatalAtencionMed
      Set oPerinatalAtencionMed.Conexion = oConexion
      Set PerinatalAtencionMedicamentoSeleccionarPorIdPerinatalAtencion = oPerinatalAtencionMed.PerinatalAtencionMedicamentoSeleccionarPorIdPerinatalAtencion(mo_idPerinatalAtencion)
      Set oPerinatalAtencionMed = Nothing
End Function
Function PerinatalAtencionDxSeleccionarPorIdPerinatalAtencion(mo_idPerinatalAtencion As Long, oConexion As Connection) As Recordset
      Dim oPerinatalAtencionDx As New PerinatalAtencionDx
      Set oPerinatalAtencionDx.Conexion = oConexion
      Set PerinatalAtencionDxSeleccionarPorIdPerinatalAtencion = oPerinatalAtencionDx.PerinatalAtencionDxSeleccionarPorIdPerinatalAtencion(mo_idPerinatalAtencion)
      Set oPerinatalAtencionDx = Nothing
End Function
Function PerinatalAtencionCredSeleccionarPorIdPaciente(ml_idPaciente As Long, oConexion As Connection) As Recordset
      Dim oPerinatalAtencionCred As New PerinatalAtencionCred
      Set oPerinatalAtencionCred.Conexion = oConexion
      Set PerinatalAtencionCredSeleccionarPorIdPaciente = oPerinatalAtencionCred.PerinatalAtencionCredSeleccionarPorIdPaciente(ml_idPaciente)
      Set oPerinatalAtencionCred = Nothing
End Function

'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function PerinatalAtencionSeleccionarPorIdPaciente(lnIdPaciente As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim lcSql As String
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set PerinatalAtencionSeleccionarPorIdPaciente = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PerinatalAtencionSeleccionarPorIdPaciente"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PerinatalAtencionSeleccionarPorIdPaciente = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing

Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function TiposDestacadosSeleccionarTodos() As Recordset
     Dim oRecordset As New ADODB.Recordset
     Dim oCommand As New ADODB.Command
     Dim oConexion As New ADODB.Connection
         Set TiposDestacadosSeleccionarTodos = Nothing
         oConexion.CursorLocation = adUseClient
         oConexion.CommandTimeout = 300
         oConexion.Open SIGHEntidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "TiposDestacadosSeleccionarTodos"
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
        End With
        Set TiposDestacadosSeleccionarTodos = oRecordset
        oConexion.Close
        Set oRecordset = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
     
End Function

'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function TiposEmpleadosSeleccionarSiSeProgramaPorId(lnIdTipoEmpleado As Long) As Boolean
     Dim oRecordset As New ADODB.Recordset
     Dim oCommand As New ADODB.Command
     Dim oConexion As New ADODB.Connection
     Dim oParameter As New ADODB.Parameter
         oConexion.CursorLocation = adUseClient
         oConexion.CommandTimeout = 300
         oConexion.Open SIGHEntidades.CadenaConexion
     TiposEmpleadosSeleccionarSiSeProgramaPorId = False
     

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "TiposEmpleadosSeleccionarSiSeProgramaPorId"
            Set oParameter = .CreateParameter("@IdTipoEmpleado", adInteger, adParamInput, 0, lnIdTipoEmpleado): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
        End With
     
     If oRecordset.RecordCount > 0 Then
        If oRecordset.Fields!esProgramado = True Then
           TiposEmpleadosSeleccionarSiSeProgramaPorId = True
        End If
     End If
        Set oRecordset = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function ColegiosHISseleccionarTodos() As Recordset
     Dim oRecordset As New ADODB.Recordset
     Dim oCommand As New ADODB.Command
     Dim oConexion As New ADODB.Connection
        Set ColegiosHISseleccionarTodos = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ColegiosHISseleccionarTodos"
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set ColegiosHISseleccionarTodos = oRecordset
       oConexion.Close
       Set oRecordset = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
End Function
'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function EtniaHISseleccionarTodos() As Recordset
     Dim oRecordset As New ADODB.Recordset
     Dim oCommand As New ADODB.Command
     Dim oConexion As New ADODB.Connection
        Set EtniaHISseleccionarTodos = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

       With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "EtniaHISseleccionarTodos"
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
        Set EtniaHISseleccionarTodos = oRecordset
        oConexion.Close
        Set oRecordset = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function TiposEmpleadosSeleccionarSegunFiltro(lcWhere As String) As Recordset
     Dim oRecordset As New ADODB.Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As New ADODB.Parameter
     Dim oConexion As New ADODB.Connection
         Set TiposEmpleadosSeleccionarSegunFiltro = Nothing
         oConexion.CursorLocation = adUseClient
         oConexion.CommandTimeout = 300
         oConexion.Open SIGHEntidades.CadenaConexion
       With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "TiposEmpleadosSeleccionarSegunFiltro"
             'SCCQ 16/04/2020 Cambio2 Inicio -->Se incrementó el tamaño del parámetro de 50 a 80
            Set oParameter = .CreateParameter("@Where", adVarChar, adParamInput, 80, lcWhere): .Parameters.Append oParameter
            'SCCQ 16/04/2020 Cambio2 Fin
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
        Set TiposEmpleadosSeleccionarSegunFiltro = oRecordset
        oConexion.Close
        Set oRecordset = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function ConsumoCptEnLaboratorioPorIdAntencionParaParticulares(lnIdAtencion As Long) As Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set ConsumoCptEnLaboratorioPorIdAntencionParaParticulares = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ConsumoCptEnLaboratorioPorIdAntencionParaParticulares"
            Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set ConsumoCptEnLaboratorioPorIdAntencionParaParticulares = oRecordset
       oConexion.Close
       Set oRecordset = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function ConsumoCptEnLaboratorioPorIdAntencionConSeguro(lnIdAtencion As Long) As Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set ConsumoCptEnLaboratorioPorIdAntencionConSeguro = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

       With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ConsumoCptEnLaboratorioPorIdAntencionConSeguro"
            Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set ConsumoCptEnLaboratorioPorIdAntencionConSeguro = oRecordset
       oConexion.Close
       Set oRecordset = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function ConsumoMedicamentosPorIdAntencionConSeguroYparticular(lnIdAtencion As Long) As Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set ConsumoMedicamentosPorIdAntencionConSeguroYparticular = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

       With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ConsumoMedicamentosPorIdAntencionConSeguroYparticular"
            Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set ConsumoMedicamentosPorIdAntencionConSeguroYparticular = oRecordset
       oConexion.Close
       Set oRecordset = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function ConsumoCptEnImagenesPorIdAntencionParaParticulares(lnIdAtencion As Long) As Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set ConsumoCptEnImagenesPorIdAntencionParaParticulares = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

       With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ConsumoCptEnImagenesPorIdAntencionParaParticulares"
            Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set ConsumoCptEnImagenesPorIdAntencionParaParticulares = oRecordset
       oConexion.Close
       Set oRecordset = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 09-10-13 Se cambio a Store Procedure
Function ConsumoCptEnImagenesPorIdAntencionConSeguro(lnIdAtencion As Long) As Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set ConsumoCptEnImagenesPorIdAntencionConSeguro = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ConsumoCptEnImagenesPorIdAntencionConSeguro"
            Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set ConsumoCptEnImagenesPorIdAntencionConSeguro = oRecordset
       oConexion.Close
       Set oRecordset = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
     
End Function
'debb-18/05/2016
Function RecetasConCabeceraYdetalleSoloCpt(lnIdReceta As Long, lnIdEstado As Long) As Recordset
       Dim oRsTmp1 As New Recordset
       Dim oCommand As New ADODB.Command
       Dim oParameter As ADODB.Parameter
       Dim oConexion As New ADODB.Connection
           oConexion.CursorLocation = adUseClient
           oConexion.CommandTimeout = 300
           oConexion.Open SIGHEntidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "RecetasConCabeceraYdetalleSoloCpt"
            Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
            If lnIdEstado > 0 Then
               oRsTmp1.Filter = "idEstado=" & Trim(Str(sghRecetaEstados.sighRecetaRegistrada))
            End If
       End With
       Set RecetasConCabeceraYdetalleSoloCpt = oRsTmp1
       oConexion.Close
       Set oRsTmp1 = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function FactCatalogoServiciosPtosSeleccionar(lcFiltro As String, Optional oConexion1 As Connection) As Recordset
       Dim oRsTmp1 As New Recordset
       Dim oCommand As New ADODB.Command
       Dim oParameter As ADODB.Parameter
       Dim oConexion As New ADODB.Connection
       If oConexion1 Is Nothing Then
           oConexion.CursorLocation = adUseClient
           oConexion.CommandTimeout = 300
           oConexion.Open SIGHEntidades.CadenaConexion
       End If

       With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
                Set .ActiveConnection = oConexion
            Else
                Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "FactCatalogoServiciosPtosSeleccionar"
            Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 100, lcFiltro): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
        End With
        Set FactCatalogoServiciosPtosSeleccionar = oRsTmp1
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
        Set oRsTmp1 = Nothing
       
'       If oConexion Is Nothing Then
'          oRsTmp1.Open lcSql, sighentidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
'       Else
'          oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'       End If
'       Set FactCatalogoServiciosPtosSeleccionar = oRsTmp1
End Function

'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function RecetasDevuelveDatosDelDetalle(lnIdReceta As Long, lnIdPuntoCarga As Long) As Recordset
      Dim oRsTmp As New Recordset
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      Dim oConexion As New ADODB.Connection
          Set RecetasDevuelveDatosDelDetalle = Nothing
          oConexion.CursorLocation = adUseClient
          oConexion.CommandTimeout = 300
          oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetasDevuelveDatosDelDetalle"
        Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@PtoCargaFarmacia", adInteger, adParamInput, 0, sghPtoCargaFarmacia): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set RecetasDevuelveDatosDelDetalle = oRsTmp
   oConexion.Close
   Set oRsTmp = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing

End Function
'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function RecetaCabeceraDetalleSeleccionaPorNroReceta(lnIdReceta As Long) As Recordset
       Dim oRsTmp1 As New Recordset
       Dim oCommand As New ADODB.Command
       Dim oParameter As ADODB.Parameter
       Dim oConexion As New ADODB.Connection
           Set RecetaCabeceraDetalleSeleccionaPorNroReceta = Nothing
           oConexion.CursorLocation = adUseClient
           oConexion.CommandTimeout = 300
           oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaCabeceraDetalleSeleccionaPorNroReceta"
        Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set RecetaCabeceraDetalleSeleccionaPorNroReceta = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function DevuelveNombreEstablecimientoRenaes(lcFuaReferidoOrigenCodigoRENAES As String) As String
      Dim oRsTmp As New Recordset
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      Dim oConexion As New ADODB.Connection
          oConexion.CursorLocation = adUseClient
          oConexion.CommandTimeout = 300
          oConexion.Open SIGHEntidades.CadenaConexion
          DevuelveNombreEstablecimientoRenaes = ""

        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "DevuelveNombreEstablecimientoRenaes"
             Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 100, lcFuaReferidoOrigenCodigoRENAES): .Parameters.Append oParameter
             Set oRsTmp = .Execute
             Set oRsTmp.ActiveConnection = Nothing
        End With
      
      If oRsTmp.RecordCount > 0 Then
         DevuelveNombreEstablecimientoRenaes = oRsTmp.Fields!nombre
      End If
      oRsTmp.Close
      oConexion.Close
      Set oRsTmp = Nothing
      Set oConexion = Nothing
      Set oCommand = Nothing
      Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function DiagnosticosSeleccionarPorCie10devuelveDescripcion(lcCodigoCie10 As String) As String
      Dim oRsTmp As New Recordset
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      Dim oConexion As New ADODB.Connection
          oConexion.CursorLocation = adUseClient
          oConexion.CommandTimeout = 300
          oConexion.Open SIGHEntidades.CadenaConexion
      DiagnosticosSeleccionarPorCie10devuelveDescripcion = ""

        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "DiagnosticosSeleccionarPorCie10devuelveDescripcion"
             Set oParameter = .CreateParameter("@CodigoCie2004", adVarChar, adParamInput, 100, lcCodigoCie10): .Parameters.Append oParameter
             Set oRsTmp = .Execute
             Set oRsTmp.ActiveConnection = Nothing
        End With
      
      If oRsTmp.RecordCount > 0 Then
         DiagnosticosSeleccionarPorCie10devuelveDescripcion = oRsTmp.Fields!Descripcion
      End If
      oRsTmp.Close
      oConexion.Close
      Set oRsTmp = Nothing
      Set oConexion = Nothing
      Set oCommand = Nothing
      Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function MedicamentosInsumosSeleccionarPorCodigoDevuelveNombre(lcCodigoDigemid As String) As String
      Dim oRsTmp As New Recordset
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      Dim oConexion As New ADODB.Connection
          oConexion.CursorLocation = adUseClient
          oConexion.CommandTimeout = 300
          oConexion.Open SIGHEntidades.CadenaConexion
      MedicamentosInsumosSeleccionarPorCodigoDevuelveNombre = ""

        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "MedicamentosInsumosSeleccionarPorCodigoDevuelveNombre"
             Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 100, lcCodigoDigemid): .Parameters.Append oParameter
             Set oRsTmp = .Execute
             Set oRsTmp.ActiveConnection = Nothing
        End With
      If oRsTmp.RecordCount > 0 Then
         MedicamentosInsumosSeleccionarPorCodigoDevuelveNombre = oRsTmp.Fields!nombre
      End If
      oRsTmp.Close
      oConexion.Close
      Set oRsTmp = Nothing
      Set oConexion = Nothing
      Set oCommand = Nothing
      Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function ProcedimientosSeleccionarPorCodigo(lcCodigoCpt As String, Optional oConexion1 As Connection) As Recordset
      Dim oRsTmp As New Recordset
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      Dim oConexion As New ADODB.Connection
      Set ProcedimientosSeleccionarPorCodigo = Nothing
      If oConexion1 Is Nothing Then
          SIGHEntidades.AbreConexionSIGH oConexion
      End If
   With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "ProcedimientosSeleccionarPorCodigo"
        Set oParameter = .CreateParameter("@CodigoCpt", adVarChar, adParamInput, 100, lcCodigoCpt): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        If oConexion1 Is Nothing Then
           Set oRsTmp.ActiveConnection = Nothing
        End If
   End With
   Set ProcedimientosSeleccionarPorCodigo = oRsTmp
   If oConexion1 Is Nothing Then
        oConexion.Close
        Set oRsTmp = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Sub TiposCondicionPacienteCondicionAlEstablecimientoYservicio(ByRef lnIdCondicionEstablecimiento As Long, _
                                                              ByRef lnIdCondicionServicio As Long, _
                                                              lnIdPaciente As Long, ldFechaIngreso As Date, _
                                                              lnIdAtencion As Long, lnIdServicio As Long)
      Dim oRsTmp As New Recordset
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      Dim oConexion As New ADODB.Connection
          oConexion.CursorLocation = adUseClient
          oConexion.CommandTimeout = 300
          oConexion.Open SIGHEntidades.CadenaConexion

      With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposCondicionPacienteCondicionAlEstablecimientoYservicio"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ldFechaIngreso", adDBTimeStamp, adParamInput, 0, Format(ldFechaIngreso, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
      End With
       
      If oRsTmp.RecordCount = 0 Then
         lnIdCondicionEstablecimiento = sghTipoCondicion.sghTipoCondicionNuevo
         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionNuevo
      Else
         oRsTmp.MoveFirst
         Select Case oRsTmp.Fields!idTipoCondicionAlEstab
         Case sghTipoCondicion.sghTipoCondicionNuevo
              Select Case oRsTmp.Fields!idTipoCondicionAlServicio
              Case sghTipoCondicion.sghTipoCondicionNuevo    'N,N
                   If Year(oRsTmp.Fields!FechaIngreso) = Year(ldFechaIngreso) Then
                      lnIdCondicionEstablecimiento = sghTipoCondicion.sghTipoCondicionContinuador
                      oRsTmp.Filter = "idServicioIngreso=" & lnIdServicio
                      If oRsTmp.RecordCount = 0 Then
                         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionNuevo
                      Else
                         oRsTmp.MoveFirst
                         If Year(oRsTmp.Fields!FechaIngreso) = Year(ldFechaIngreso) Then
                            lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionContinuador
                         Else
                            lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionReingresante
                         End If
                      End If
                   Else
                      lnIdCondicionEstablecimiento = sghTipoCondicion.sghTipoCondicionReingresante
                      oRsTmp.Filter = "idServicioIngreso=" & lnIdServicio
                      If oRsTmp.RecordCount = 0 Then
                         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionNuevo
                      Else
                         If Year(oRsTmp.Fields!FechaIngreso) = Year(ldFechaIngreso) Then
                             lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionContinuador
                         Else
                             lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionReingresante
                         End If
                      End If
                   End If
              Case sghTipoCondicion.sghTipoCondicionReingresante   'N,R
                   If Year(oRsTmp.Fields!FechaIngreso) = Year(ldFechaIngreso) Then
                      lnIdCondicionEstablecimiento = sghTipoCondicion.sghTipoCondicionContinuador
                      oRsTmp.Filter = "idServicioIngreso=" & lnIdServicio
                      If oRsTmp.RecordCount = 0 Then
                         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionNuevo
                      Else
                         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionContinuador
                      End If
                   Else
                      lnIdCondicionEstablecimiento = sghTipoCondicion.sghTipoCondicionReingresante
                      oRsTmp.Filter = "idServicioIngreso=" & lnIdServicio
                      If oRsTmp.RecordCount = 0 Then
                         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionNuevo
                      Else
                         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionContinuador
                      End If
                   End If
              Case sghTipoCondicion.sghTipoCondicionContinuador    'N,C
                   If Year(oRsTmp.Fields!FechaIngreso) = Year(ldFechaIngreso) Then
                      lnIdCondicionEstablecimiento = sghTipoCondicion.sghTipoCondicionContinuador
                      oRsTmp.Filter = "idServicioIngreso=" & lnIdServicio
                      If oRsTmp.RecordCount = 0 Then
                         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionNuevo
                      Else
                         oRsTmp.MoveFirst
                         If Year(oRsTmp.Fields!FechaIngreso) = Year(ldFechaIngreso) Then
                            lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionContinuador
                         Else
                            lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionReingresante
                         End If
                      End If
                   Else
                      lnIdCondicionEstablecimiento = sghTipoCondicion.sghTipoCondicionReingresante
                      oRsTmp.Filter = "idServicioIngreso=" & lnIdServicio
                      If oRsTmp.RecordCount = 0 Then
                         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionNuevo
                      Else
                         If Year(oRsTmp.Fields!FechaIngreso) = Year(ldFechaIngreso) Then
                             lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionContinuador
                         Else
                             lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionReingresante
                         End If
                      End If
                   End If
              End Select
         Case sghTipoCondicion.sghTipoCondicionReingresante
              Select Case oRsTmp.Fields!idTipoCondicionAlServicio
              Case sghTipoCondicion.sghTipoCondicionNuevo          'R,N
                    lnIdCondicionEstablecimiento = sghTipoCondicion.sghTipoCondicionContinuador
                    oRsTmp.Filter = "idServicioIngreso=" & lnIdServicio
                    If oRsTmp.RecordCount = 0 Then
                       lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionNuevo
                    Else
                       oRsTmp.MoveFirst
                       If Year(oRsTmp.Fields!FechaIngreso) = Year(ldFechaIngreso) Then
                          lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionContinuador
                       Else
                          lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionReingresante
                       End If
                    End If
              Case sghTipoCondicion.sghTipoCondicionReingresante   'R,R
                   lnIdCondicionEstablecimiento = sghTipoCondicion.sghTipoCondicionContinuador
                   oRsTmp.Filter = "idServicioIngreso=" & lnIdServicio
                   If oRsTmp.RecordCount = 0 Then
                       lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionNuevo
                   Else
                       lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionContinuador
                   End If
              Case sghTipoCondicion.sghTipoCondicionContinuador    'R,C
                   lnIdCondicionEstablecimiento = sghTipoCondicion.sghTipoCondicionContinuador
                   oRsTmp.Filter = "idServicioIngreso=" & lnIdServicio
                   If oRsTmp.RecordCount = 0 Then
                       lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionNuevo
                   Else
                       oRsTmp.MoveFirst
                       If Year(oRsTmp.Fields!FechaIngreso) = Year(ldFechaIngreso) Then
                          lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionContinuador
                       Else
                          lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionReingresante
                       End If
                   End If
              End Select
         Case sghTipoCondicion.sghTipoCondicionContinuador
              Select Case oRsTmp.Fields!idTipoCondicionAlServicio
              Case sghTipoCondicion.sghTipoCondicionNuevo          'C,N
                   If Year(oRsTmp.Fields!FechaIngreso) = Year(ldFechaIngreso) Then
                      lnIdCondicionEstablecimiento = sghTipoCondicion.sghTipoCondicionContinuador
                      oRsTmp.Filter = "idServicioIngreso=" & lnIdServicio
                      If oRsTmp.RecordCount = 0 Then
                         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionNuevo
                      Else
                         oRsTmp.MoveFirst
                         If Year(oRsTmp.Fields!FechaIngreso) = Year(ldFechaIngreso) Then
                            lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionContinuador
                         Else
                            lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionReingresante
                         End If
                      End If
                   Else
                      lnIdCondicionEstablecimiento = sghTipoCondicion.sghTipoCondicionReingresante
                      oRsTmp.Filter = "idServicioIngreso=" & lnIdServicio
                      If oRsTmp.RecordCount = 0 Then
                         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionNuevo
                      Else
                         oRsTmp.MoveFirst
                         If Year(oRsTmp.Fields!FechaIngreso) = Year(ldFechaIngreso) Then
                            lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionContinuador
                         Else
                            lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionReingresante
                         End If
                      End If
                   End If
              Case sghTipoCondicion.sghTipoCondicionReingresante   'C,R
                   If Year(oRsTmp.Fields!FechaIngreso) = Year(ldFechaIngreso) Then
                      lnIdCondicionEstablecimiento = sghTipoCondicion.sghTipoCondicionContinuador
                      oRsTmp.Filter = "idServicioIngreso=" & lnIdServicio
                      If oRsTmp.RecordCount = 0 Then
                         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionNuevo
                      Else
                         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionContinuador
                      End If
                   Else
                      lnIdCondicionEstablecimiento = sghTipoCondicion.sghTipoCondicionReingresante
                      oRsTmp.Filter = "idServicioIngreso=" & lnIdServicio
                      If oRsTmp.RecordCount = 0 Then
                         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionNuevo
                      Else
                         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionContinuador
                      End If
                   End If
              Case sghTipoCondicion.sghTipoCondicionContinuador    'C,C
                   If Year(oRsTmp.Fields!FechaIngreso) = Year(ldFechaIngreso) Then
                      lnIdCondicionEstablecimiento = sghTipoCondicion.sghTipoCondicionContinuador
                      oRsTmp.Filter = "idServicioIngreso=" & lnIdServicio
                      If oRsTmp.RecordCount = 0 Then
                         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionNuevo
                      Else
                         oRsTmp.MoveFirst
                         If Year(oRsTmp.Fields!FechaIngreso) = Year(ldFechaIngreso) Then
                            lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionContinuador
                         Else
                            lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionReingresante
                         End If
                      End If
                   Else
                      lnIdCondicionEstablecimiento = sghTipoCondicion.sghTipoCondicionReingresante
                      oRsTmp.Filter = "idServicioIngreso=" & lnIdServicio
                      If oRsTmp.RecordCount = 0 Then
                         lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionNuevo
                      Else
                         oRsTmp.MoveFirst
                         If Year(oRsTmp.Fields!FechaIngreso) = Year(ldFechaIngreso) Then
                            lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionContinuador
                         Else
                            lnIdCondicionServicio = sghTipoCondicion.sghTipoCondicionReingresante
                         End If
                      End If
                   End If
              End Select
         End Select
      End If
      oRsTmp.Close
      Set oRsTmp = Nothing
      Set oConexion = Nothing
      Set oCommand = Nothing
      Set oParameter = Nothing
End Sub

Function FactCatalogoServiciosPtosSeleccionarXidProducto(lnIdProducto As Long, Optional oConexion1 As Connection) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
            If oConexion1 Is Nothing Then
                oConexion.CursorLocation = adUseClient
                oConexion.CommandTimeout = 300
                oConexion.Open SIGHEntidades.CadenaConexion
            End If
            With oCommand
                 .CommandType = adCmdStoredProc
                 If oConexion1 Is Nothing Then
                    Set .ActiveConnection = oConexion
                 Else
                    Set .ActiveConnection = oConexion1
                 End If
                 .CommandTimeout = 150
                 .CommandText = "FactCatalogoServiciosPtosSeleccionarXidProducto"
                 Set oParameter = .CreateParameter("    ", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                 Set oRsTmp1 = .Execute
                 If oConexion1 Is Nothing Then
                    Set oRsTmp1.ActiveConnection = Nothing
                 End If
            End With
            Set FactCatalogoServiciosPtosSeleccionarXidProducto = oRsTmp1
            If oConexion1 Is Nothing Then
                oConexion.Close
                Set oRsTmp1 = Nothing
            End If
            Set oConexion = Nothing
            Set oCommand = Nothing
            Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function FactPuntosCargaSeleccionarPorFiltro(lcFiltro As String, Optional oConexion1 As Connection) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        If oConexion1 Is Nothing Then
            oConexion.Open SIGHEntidades.CadenaConexion
            oConexion.CursorLocation = adUseClient
        End If

        With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "FactPuntosCargaSeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 500, lcFiltro): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
        End With
        Set FactPuntosCargaSeleccionarPorFiltro = oRsTmp1
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
        Set oRsTmp1 = Nothing
'        If oConexion Is Nothing Then
'           oRsTmp1.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'        Else
'          oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'        End If
'        Set FactPuntosCargaSeleccionarPorFiltro = oRsTmp1
End Function
'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function EmpleadosCargosSeleccionarPorFiltro(lcFiltro As String) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
            Set EmpleadosCargosSeleccionarPorFiltro = Nothing
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "EmpleadosCargosSeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 100, lcFiltro): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set EmpleadosCargosSeleccionarPorFiltro = oRsTmp1
        Set oCommand = Nothing
        Set oParameter = Nothing
        Set oRsTmp1 = Nothing
End Function
'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function EmpleadosLugarDeTrabajoSeleccionarPorFiltro(lcFiltro As String) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
            Set EmpleadosLugarDeTrabajoSeleccionarPorFiltro = Nothing
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open SIGHEntidades.CadenaConexion

       With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "EmpleadosLugarDeTrabajoSeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 100, lcFiltro): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set EmpleadosLugarDeTrabajoSeleccionarPorFiltro = oRsTmp1
        Set oCommand = Nothing
        Set oParameter = Nothing
        Set oRsTmp1 = Nothing
End Function
'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function EstablecimientosSeleccionarPorFiltro(lcFiltro As String, Optional oConexion1 As Connection) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New Connection
        If oConexion1 Is Nothing Then
            oConexion.CommandTimeout = 300
            oConexion.CursorLocation = adUseClient
            oConexion.Open SIGHEntidades.CadenaConexion
        End If
        With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "EstablecimientosSeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 100, lcFiltro): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
        End With
        Set EstablecimientosSeleccionarPorFiltro = oRsTmp1
        Set oCommand = Nothing
        Set oParameter = Nothing
        Set oConexion = Nothing
End Function



Function EspecialidadesSeleccionarPorFiltro(lcFiltro As String, Optional oConexion1 As Connection) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Set EspecialidadesSeleccionarPorFiltro = Nothing
        Dim oConexion As New ADODB.Connection
        If oConexion1 Is Nothing Then
            oConexion.CommandTimeout = 300
            oConexion.CursorLocation = adUseClient
            oConexion.Open SIGHEntidades.CadenaConexion
        End If
        

        With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "EspecialidadesSeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 100, lcFiltro): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
        End With
        Set EspecialidadesSeleccionarPorFiltro = oRsTmp1
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
'        If oConexion Is Nothing Then
'           oRsTmp1.Open lcSql, sighentidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
'        Else
'           oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'        End If
'        Set EspecialidadesSeleccionarPorFiltro = oRsTmp1
End Function
'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function FactPuntosCargaSeleccionarPorIdPuntoCarga(lnIdPuntoCarga As Long) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
            Set FactPuntosCargaSeleccionarPorIdPuntoCarga = Nothing
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "FactPuntosCargaSeleccionarPorIdPuntoCarga"
             Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
             Set oRsTmp1 = .Execute
             Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set FactPuntosCargaSeleccionarPorIdPuntoCarga = oRsTmp1
        oConexion.Close
        Set oRsTmp1 = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
End Function

Function CatalogoBienesInsumosSeleccionarPorFiltro(lcFiltro As String, lnOrdenadoPor As sghTipoOrden, Optional oConexion1 As Connection) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset
                Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Set CatalogoBienesInsumosSeleccionarPorFiltro = Nothing
        Dim oConexion As New ADODB.Connection
        If oConexion1 Is Nothing Then
            oConexion.CommandTimeout = 300
            oConexion.CursorLocation = adUseClient
            oConexion.Open SIGHEntidades.CadenaConexion
        End If

        With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "CatalogoBienesInsumosSeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Order", adVarChar, adParamInput, 50, IIf(lnOrdenadoPor = sghPorCodigo, " Order by Codigo", " ORDER BY Nombre")): .Parameters.Append oParameter

            Set oRsTmp1 = .Execute
        End With
        Set CatalogoBienesInsumosSeleccionarPorFiltro = oRsTmp1
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing

'        If oConexion Is Nothing Then
'           oRsTmp1.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'        Else
'           oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'        End If
'        Set CatalogoBienesInsumosSeleccionarPorFiltro = oRsTmp1
End Function
'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function CatalogoBienesInsumosResumenSeleccionarPorFiltro(lcFiltro As String, lnOrdenadoPor As sghTipoOrden) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
            Set CatalogoBienesInsumosResumenSeleccionarPorFiltro = Nothing
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CatalogoBienesInsumosResumenSeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Order", adVarChar, adParamInput, 50, IIf(lnOrdenadoPor = sghPorCodigo, " Order by Codigo", " ORDER BY Nombre")): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set CatalogoBienesInsumosResumenSeleccionarPorFiltro = oRsTmp1
        Set oCommand = Nothing
        Set oParameter = Nothing
        Set oRsTmp1 = Nothing
        
End Function

'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function CatalogoServiciosSeleccionarSoloAdministrativos() As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oConexion As New ADODB.Connection
            Set CatalogoServiciosSeleccionarSoloAdministrativos = Nothing
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "CatalogoServiciosSeleccionarSoloAdministrativos"
             Set oRsTmp1 = .Execute
             Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set CatalogoServiciosSeleccionarSoloAdministrativos = oRsTmp1
        oConexion.Close
        Set oRsTmp1 = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
        
End Function

Function CatalogoServiciosSeleccionarSoloConPreciosEnSIS() As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oConexion As New ADODB.Connection
            Set CatalogoServiciosSeleccionarSoloConPreciosEnSIS = Nothing
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "CatalogoServiciosSeleccionarSoloConPreciosEnSIS"
             Set oRsTmp1 = .Execute
             Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set CatalogoServiciosSeleccionarSoloConPreciosEnSIS = oRsTmp1
        oConexion.Close
        Set oRsTmp1 = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function CatalogoBienesInsumosSeleccionarSoloConPreciosEnSIS() As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oConexion As New ADODB.Connection
            Set CatalogoBienesInsumosSeleccionarSoloConPreciosEnSIS = Nothing
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "CatalogoBienesInsumosSeleccionarSoloConPreciosEnSIS"
             Set oRsTmp1 = .Execute
             Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set CatalogoBienesInsumosSeleccionarSoloConPreciosEnSIS = oRsTmp1
        oConexion.Close
        Set oRsTmp1 = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function
'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function CatalogoServiciosSeleccionarSoloConPreciosEnParticular(lnIdPuntoCarga As Long) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CatalogoServiciosSeleccionarSoloConPreciosEnParticular"
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set CatalogoServiciosSeleccionarSoloConPreciosEnParticular = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function EmpleadosSeleccionarPorIdEmpleado(lnIdEmpleado As Long, Optional oConexion1 As Connection) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        If oConexion1 Is Nothing Then
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open SIGHEntidades.CadenaConexion
         End If

         With oCommand
             .CommandType = adCmdStoredProc
             If oConexion1 Is Nothing Then
                Set .ActiveConnection = oConexion
             Else
                Set .ActiveConnection = oConexion1
             End If
             .CommandTimeout = 150
             .CommandText = "EmpleadosSeleccionarPorIdEmpleado"
             Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, lnIdEmpleado): .Parameters.Append oParameter
             Set oRsTmp1 = .Execute
        End With
        Set EmpleadosSeleccionarPorIdEmpleado = oRsTmp1
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
'        If oConexion Is Nothing Then
'           oRsTmp1.Open lcSql, sighentidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
'        Else
'           oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'        End If
'        Set EmpleadosSeleccionarPorIdEmpleado = oRsTmp1
End Function
'Modificado por Yamill Palomino 10-10-13 Se cambio a Store Procedure
Function AtencionesSeleccionarMedicoPorCuenta(lnIdCuentaAtencion As Long) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set AtencionesSeleccionarMedicoPorCuenta = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "AtencionesSeleccionarMedicoPorCuenta"
             Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
             Set oRsTmp1 = .Execute
             Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set AtencionesSeleccionarMedicoPorCuenta = oRsTmp1
        oConexion.Close
        Set oRsTmp1 = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
End Function

Sub CatalogoServiciosActualizaPartidaXidProducto(lnIdPartida As Long, lnIdProducto As Long, lbValorIdPartidaNULL As Boolean)
'   Dim lcSql As String, oRsTmp1 As New Recordset
   Dim oCommand As New ADODB.Command
   Dim oParameter As ADODB.Parameter
   Dim oConexion As New ADODB.Connection
   Dim lnValorIdPartidaNULL As Integer
   
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        
   If lbValorIdPartidaNULL = True Then
        lnValorIdPartidaNULL = 1
   Else
        lnValorIdPartidaNULL = 0
   End If

        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "CatalogoServiciosActualizaPartidaXidProducto"
             Set oParameter = .CreateParameter("@ValorIdPartidaNULL", adInteger, adParamInput, 0, lnValorIdPartidaNULL): .Parameters.Append oParameter
             Set oParameter = .CreateParameter("@IdPartida", adInteger, adParamInput, 0, Trim(Str(lnIdPartida))): .Parameters.Append oParameter
             Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
             .Execute
        End With
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
End Sub

Sub CatalogoServicioActualizaNULLenIdPartida(lnIdPartida As Long)
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
  
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CatalogoServicioActualizaNULLenIdPartida"
        Set oParameter = .CreateParameter("@IdPartida", adInteger, adParamInput, 0, lnIdPartida): .Parameters.Append oParameter
        .Execute
    End With
    Set oParameter = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Sub

Sub CatalogoBienesInsumosActualizaNULLenIdPartida(lnIdPartida As Long)
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CatalogoBienesInsumosActualizaNULLenIdPartida"
        Set oParameter = .CreateParameter("@IdPartida", adInteger, adParamInput, 0, lnIdPartida): .Parameters.Append oParameter
        .Execute
    End With
    Set oParameter = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Sub
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function ServiciosSeleccionarEmergenciaPorEspecialidad(lnEspecialidad As Long) As ADODB.Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
        Set ServiciosSeleccionarEmergenciaPorEspecialidad = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
    On Error GoTo ManejadorDeError
    ms_MensajeError = ""

        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "ServiciosSeleccionarEmergenciaPorEspecialidad"
             Set oParameter = .CreateParameter("@Especialidad", adInteger, adParamInput, 0, lnEspecialidad): .Parameters.Append oParameter
             Set oRecordset = .Execute
             Set oRecordset.ActiveConnection = Nothing
        End With
   Set ServiciosSeleccionarEmergenciaPorEspecialidad = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    Exit Function
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function DiagnosticosSeleccionarXCodigo(lcCodigoDx As String) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set DiagnosticosSeleccionarXCodigo = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "DiagnosticosSeleccionarXCodigo"
             Set oParameter = .CreateParameter("@CodigoDx", adVarChar, adParamInput, 7, Trim(lcCodigoDx)): .Parameters.Append oParameter
             Set oRsTmp1 = .Execute
             Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set DiagnosticosSeleccionarXCodigo = oRsTmp1
        oConexion.Close
        Set oRsTmp1 = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function RecetaDetalleSoloFarmaciaSeleccionarXidCuentaAtencion(ml_idCuentaAtencion As Long, lbSoloRegistradasYnoDespachadas As Boolean) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set RecetaDetalleSoloFarmaciaSeleccionarXidCuentaAtencion = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaDetalleSoloFarmaciaSeleccionarXidCuentaAtencion"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_idCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@SoloRegistradasYnoDespachadas", adVarChar, adParamInput, 50, IIf(lbSoloRegistradasYnoDespachadas = True, " and  dbo.RecetaCabecera.idEstado=1", "")): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set RecetaDetalleSoloFarmaciaSeleccionarXidCuentaAtencion = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function RecetaDetalleSoloServiciosSeleccionarXidCuentaAtencion(ml_idCuentaAtencion As Long, lbSoloRegistradasYnoDespachadas As Boolean) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set RecetaDetalleSoloServiciosSeleccionarXidCuentaAtencion = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "RecetaDetalleSoloServiciosSeleccionarXidCuentaAtencion"
             Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_idCuentaAtencion): .Parameters.Append oParameter
             Set oParameter = .CreateParameter("@SoloRegistradasYnoDespachadas", adVarChar, adParamInput, 50, IIf(lbSoloRegistradasYnoDespachadas = True, " and  dbo.RecetaCabecera.idEstado=1", "")): .Parameters.Append oParameter
             Set oRsTmp1 = .Execute
             Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set RecetaDetalleSoloServiciosSeleccionarXidCuentaAtencion = oRsTmp1
        oConexion.Close
        Set oRsTmp1 = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
         
End Function

'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function RecetaDetalleSeleccionarXidCuentaIdItemDocumento(ml_idCuentaAtencion As Long, lnIdProducto As Long, lcDocumentoNumero As String) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set RecetaDetalleSeleccionarXidCuentaIdItemDocumento = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaDetalleSeleccionarXidCuentaIdItemDocumento"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_idCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@DocumentoNumero", adVarChar, adParamInput, 50, lcDocumentoNumero): .Parameters.Append oParameter

        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set RecetaDetalleSeleccionarXidCuentaIdItemDocumento = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function TiposDocIdentidadDevuelveTipoDocumentoSis(ml_idTipoDocumentoGalenhos As Long) As String
    TiposDocIdentidadDevuelveTipoDocumentoSis = "2"
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposDocIdentidadDevuelveTipoDocumentoSis"
        Set oParameter = .CreateParameter("@idTipoDocumentoGalenhos", adInteger, adParamInput, 0, ml_idTipoDocumentoGalenhos): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   
    If oRsTmp1.RecordCount > 0 Then
       TiposDocIdentidadDevuelveTipoDocumentoSis = oRsTmp1.Fields!CodigoSIS
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Sub CatalogoBienesInsumosActualizaPRECIOS(lnPrCompraNew As Double, lnPrDistribucionNew As Double, lnIdProducto As Long, oConexion1 As Connection)
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        If oConexion1 Is Nothing Then
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open SIGHEntidades.CadenaConexion
        End If
        With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "CatalogoBienesInsumosActualizaPRECIOS"
            Set oParameter = .CreateParameter("@PrCompraNew", adCurrency, adParamInput, 0, FormatCurrency(lnPrCompraNew, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@PrDistribucionNew", adCurrency, adParamInput, 0, FormatCurrency(lnPrDistribucionNew, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
            .Execute
        End With
        Set oParameter = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
End Sub

'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function EmpleadosDevuelveNombreUsuario(lnIdEmpleado As Long) As String
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosDevuelveNombreUsuario"
        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, lnIdEmpleado): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
 EmpleadosDevuelveNombreUsuario = oRsTmp1.Fields!Usuario
   oConexion.Close
   Set oCommand = Nothing
   Set oParameter = Nothing
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
End Function


'HRA- para el Reporte RESUMEN POR CENTRO DE COSTOS
Sub ActualizaCentroCostosParaItems(lbEnVersionDespuesOct2012 As Boolean, oConexion As Connection)
    ActualizaPuntoCargaImagenesLaboratorio lbEnVersionDespuesOct2012, oConexion, sghPtoCargaRayosX, "AD029"
    ActualizaPuntoCargaImagenesLaboratorio lbEnVersionDespuesOct2012, oConexion, sghPtoCargaEcogObstetrica, "008.04"
    ActualizaPuntoCargaImagenesLaboratorio lbEnVersionDespuesOct2012, oConexion, sghPtoCargaRayosX, "73562"
    ActualizaPuntoCargaImagenesLaboratorio lbEnVersionDespuesOct2012, oConexion, sghPtoCargaRayosX, "73560"
    ActualizaPuntoCargaImagenesLaboratorio lbEnVersionDespuesOct2012, oConexion, sghPtoCargaPatologiaClinica, "10001"
    ActualizaPuntoCargaImagenesLaboratorio lbEnVersionDespuesOct2012, oConexion, sghPtoCargaTomografia, "AD041"
    ActualizaPuntoCargaImagenesLaboratorio lbEnVersionDespuesOct2012, oConexion, sghPtoCargaRayosX, "90782"
    ActualizaPuntoCargaImagenesLaboratorio lbEnVersionDespuesOct2012, oConexion, sghPtoCargaEcogObstetrica, "76818"
    ActualizaPuntoCargaImagenesLaboratorio lbEnVersionDespuesOct2012, oConexion, sghPtoCargaTomografia, "70450"
End Sub

Sub ActualizaPuntoCargaImagenesLaboratorio(lbEsVersionDespuesOct2012 As Boolean, oConexion As Connection, lnIdPuntoCarga As sghPuntosCargaBasicos, lcCodigoCpt As String)
End Sub

Function TiposTarifaFiltrar(oDOTiposTarifa As DOTiposTarifa) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposTarifa
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set TiposTarifaFiltrar = oTabla.Filtrar(oDOTiposTarifa)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function
'-------------------------------------------------------------------------------------
'                                       CentrosCosto
'-------------------------------------------------------------------------------------
Function TiposTarifaAgregar(oDOTiposTarifa As DOTiposTarifa, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreCC As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTiposTarifa As New TiposTarifa
    
    TiposTarifaAgregar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    Set oTiposTarifa.Conexion = oConexion
    If oTiposTarifa.Insertar(oDOTiposTarifa) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOTiposTarifa.IdUsuarioAuditoria, "A", oDOTiposTarifa.IdTipoTarifa, "TiposTarifa", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreCC)          'ListBarItems.idListItem
        TiposTarifaAgregar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oTiposTarifa.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oTiposTarifa = Nothing
End Function

Function TiposTarifaModificar(oDOTiposTarifa As DOTiposTarifa, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreCC As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTiposTarifa As New TiposTarifa
    
    TiposTarifaModificar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    Set oTiposTarifa.Conexion = oConexion
    If oTiposTarifa.modificar(oDOTiposTarifa) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOTiposTarifa.IdUsuarioAuditoria, "M", oDOTiposTarifa.IdTipoTarifa, "TiposTarifa", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreCC)           'ListBarItems.idListItem
        TiposTarifaModificar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oTiposTarifa.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oTiposTarifa = Nothing
    
End Function

Function TiposTarifaEliminar(oDOTiposTarifa As DOTiposTarifa, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreCC As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTiposTarifa As New TiposTarifa

    TiposTarifaEliminar = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    Set oTiposTarifa.Conexion = oConexion
    If oTiposTarifa.Eliminar(oDOTiposTarifa) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOTiposTarifa.IdUsuarioAuditoria, "E", oDOTiposTarifa.IdTipoTarifa, "TiposTarifa", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreCC)          'ListBarItems.idListItem
        TiposTarifaEliminar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oTiposTarifa.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oTiposTarifa = Nothing
End Function
Function TiposTarifaSeleccionarPorId(IdTipoTarifa As Long) As DOTiposTarifa
    Dim oConexion As New ADODB.Connection
    Dim oProducto As New TiposTarifa
    Dim oDOTiposTarifa As New DOTiposTarifa
    Set TiposTarifaSeleccionarPorId = Nothing
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oProducto.Conexion = oConexion
    oDOTiposTarifa.IdTipoTarifa = IdTipoTarifa
    If oProducto.SeleccionarPorId(oDOTiposTarifa) Then
        Set TiposTarifaSeleccionarPorId = oDOTiposTarifa
    End If
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oProducto.MensajeError
    Set oProducto = Nothing
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Sub TiposTarifaCptActualizar(lnIdTipoTarifa As Long, lnIdProducto As Long)
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open SIGHEntidades.CadenaConexion

         With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "TiposTarifaCptActualizar"
             Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
             Set oRsTmp1 = .Execute
             Set oRsTmp1.ActiveConnection = Nothing
        End With
        If oRsTmp1.RecordCount = 0 Then
             Set oCommand = Nothing
             Set oParameter = Nothing
             With oCommand
                 .CommandType = adCmdStoredProc
                 Set .ActiveConnection = oConexion
                 .CommandTimeout = 150
                 .CommandText = "TiposTarifaCptAgregar"
                 Set oParameter = .CreateParameter("@idTipoTarifa", adInteger, adParamInput, 0, lnIdTipoTarifa): .Parameters.Append oParameter
                 Set oParameter = .CreateParameter("@IdProductoCpt", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                 Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 0): .Parameters.Append oParameter
                  .Execute
            End With
        Else
             Set oCommand = Nothing
             Set oParameter = Nothing
             With oCommand
                 .CommandType = adCmdStoredProc
                 Set .ActiveConnection = oConexion
                 .CommandTimeout = 150
                 .CommandText = "TiposTarifaCptModificarXidproducto"
                 Set oParameter = .CreateParameter("@idTipoTarifa", adInteger, adParamInput, 0, lnIdTipoTarifa): .Parameters.Append oParameter
                 Set oParameter = .CreateParameter("@IdProductoCpt", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                  .Execute
            End With
        
        End If
     
        Set oRsTmp1 = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
End Sub

Sub TiposTarifaCptEliminar(lnIdProducto As Long, lnIdTipoTarifa As Long)
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
  

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "TiposTarifaCptEliminarPorIdProducto"
            Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdTipoTarifa", adInteger, adParamInput, 0, lnIdTipoTarifa): .Parameters.Append oParameter
            .Execute
        End With
        Set oParameter = Nothing
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Sub

Sub TiposTarifaCptEliminarPorTipoTarifa(lnIdTipoTarifa As Long)
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "TiposTarifaCptEliminar"
            Set oParameter = .CreateParameter("@idTipoTarifa", adInteger, adParamInput, 0, lnIdTipoTarifa): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
            .Execute
        End With
        Set oParameter = Nothing
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Sub

'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function TiposTarifaCptSeleccionarTodos() As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
        Set TiposTarifaCptSeleccionarTodos = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposTarifaCptSeleccionarTodos"
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set TiposTarifaCptSeleccionarTodos = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function TiposTarifaSeleccionarTodos(Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
       
        If oConexion1 Is Nothing Then
            oConexion = New Connection
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open SIGHEntidades.CadenaConexion
        End If
        

        With oCommand
'           .CommandType = adCmdText
            .CommandType = adCmdStoredProc
           If oConexion1 Is Nothing Then
              Set .ActiveConnection = oConexion
           Else
              Set .ActiveConnection = oConexion1
              
           End If
            .CommandTimeout = 150
'           .CommandText = ssql
            .CommandText = "TiposTarifaSeleccionarTodos"
           Set oRecordset = .Execute
           If oConexion1 Is Nothing Then
              Set oRecordset.ActiveConnection = Nothing
           End If
        End With
        Set TiposTarifaSeleccionarTodos = oRecordset
        Set oConexion = Nothing
        Set oCommand = Nothing
   
        ms_MensajeError = ""
        Exit Function
ManejadorDeError:
        ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
        Exit Function
End Function

'---------------------------------------------------------------------------------------
' Procedure : ConsultarDepartamentos
' Author    : JVALLE
' Date      : 04/05/2012 17:43
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function ConsultarDepartamentos() As Recordset
On Error GoTo ManejadorDeError
Dim oConexion As New ADODB.Connection
Dim mo_DatosDepartamentos As New SIGHDatos.Departamentos

oConexion.Open SIGHEntidades.CadenaConexion
oConexion.CursorLocation = adUseClient
    
Set mo_DatosDepartamentos.Conexion = oConexion
Set ConsultarDepartamentos = mo_DatosDepartamentos.SeleccionarTodos
Exit Function
ManejadorDeError:
ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function

'---------------------------------------------------------------------------------------
' Procedure : ConsultarProvinciasPorDepartamento
' Author    : JVALLE
' Date      : 04/05/2012 17:43
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function ConsultarProvinciasPorDepartamento(IdDepartamento As Integer) As Recordset
On Error GoTo ManejadorDeError
Dim oConexion As New ADODB.Connection
Dim mo_DatosProvincias As New SIGHDatos.Provincias

oConexion.Open SIGHEntidades.CadenaConexion
oConexion.CursorLocation = adUseClient
    
Set mo_DatosProvincias.Conexion = oConexion
Set ConsultarProvinciasPorDepartamento = mo_DatosProvincias.SeleccionarPorDepartamento(IdDepartamento)
Exit Function
ManejadorDeError:
ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function

'---------------------------------------------------------------------------------------
' Procedure : SeleccionarPorProvinciaYDescripcion
' Author    : JVALLE
' Date      : 04/05/2012 17:43
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function SeleccionarPorProvinciaYDescripcion(IdProvicia As Long, Descripcion As String) As Recordset
On Error GoTo ManejadorDeError
Dim oConexion As New ADODB.Connection
Dim mo_DatosDistritos As New SIGHDatos.Distritos
oConexion.CommandTimeout = 300
oConexion.CursorLocation = adUseClient
oConexion.Open SIGHEntidades.CadenaConexion

    
Set mo_DatosDistritos.Conexion = oConexion
Set SeleccionarPorProvinciaYDescripcion = mo_DatosDistritos.SeleccionarPorProvinciaYDescripcion(IdProvicia, Descripcion)
Exit Function
ManejadorDeError:
ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function


Function DistritosSeleccionarPorDepartamentoYDescripcion(lnIdDepartamento As Long, Descripcion As String) As Recordset
On Error GoTo ManejadorDeError
Dim oConexion As New ADODB.Connection
Dim mo_DatosDistritos As New SIGHDatos.Distritos
oConexion.CommandTimeout = 300
oConexion.CursorLocation = adUseClient
oConexion.Open SIGHEntidades.CadenaConexion

    
Set mo_DatosDistritos.Conexion = oConexion
Set DistritosSeleccionarPorDepartamentoYDescripcion = mo_DatosDistritos.SeleccionarPorDepartamentoYDescripcion(lnIdDepartamento, Descripcion)
Exit Function
ManejadorDeError:
ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function




'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function FuentesFinanciamientoSeleccionarTodos() As Recordset
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
        Set FuentesFinanciamientoSeleccionarTodos = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ConsultarFuentesFinanciamientoXidFuenteFinanciamiento"
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set FuentesFinanciamientoSeleccionarTodos = oRsTmp
   oConexion.Close
   Set oRsTmp = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing

End Function

'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function MedicamentosInsumosSeleccionarPorCodigo(lcCodigoDigemid As String) As Recordset
      Dim oRsTmp As New Recordset
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      Dim oConexion As New ADODB.Connection
          Set MedicamentosInsumosSeleccionarPorCodigo = Nothing
          oConexion.CursorLocation = adUseClient
          oConexion.CommandTimeout = 300
          oConexion.Open SIGHEntidades.CadenaConexion
          
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "MedicamentosInsumosSeleccionarPorCodigo"
        Set oParameter = .CreateParameter("@CodigoDigemid", adVarChar, adParamInput, 50, lcCodigoDigemid): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set MedicamentosInsumosSeleccionarPorCodigo = oRsTmp
   oConexion.Close
   Set oRsTmp = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function EstablecimientosDevuelveCodigo(IdEstablecimiento As Long) As String
      Dim oRsTmp As New Recordset

      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      Dim oConexion As New ADODB.Connection
          oConexion.CursorLocation = adUseClient
          oConexion.CommandTimeout = 300
          oConexion.Open SIGHEntidades.CadenaConexion

         With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "EstablecimientosDevuelveCodigo"
             Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IdEstablecimiento): .Parameters.Append oParameter
             Set oRsTmp = .Execute
             Set oRsTmp.ActiveConnection = Nothing
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
           
      If oRsTmp.RecordCount > 0 Then
         EstablecimientosDevuelveCodigo = oRsTmp.Fields!Codigo
      Else
         EstablecimientosDevuelveCodigo = ""
      End If
      oRsTmp.Close
      Set oRsTmp = Nothing
End Function

'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function TiposIdiomasSeleccionarTodos() As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
        Set TiposIdiomasSeleccionarTodos = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposIdiomasSeleccionarTodos"
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set TiposIdiomasSeleccionarTodos = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function EstablecimientosNoMinsaSeleccionarPorCodigo(lcCodigo As String, Optional oConexion1 As Connection) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
            Set .ActiveConnection = oConexion
        Else
            Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "EstablecimientosNoMinsaSeleccionarPorCodigo"
        Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 50, lcCodigo): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        If oConexion1 Is Nothing Then
           Set oRsTmp1.ActiveConnection = Nothing
        End If
   End With
   Set EstablecimientosNoMinsaSeleccionarPorCodigo = oRsTmp1
   If oConexion1 Is Nothing Then
        oConexion.Close
        Set oRsTmp1 = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
 
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function DiagnosticosSeleccionarXcodigoCIEsinPto(lcCodigoCIEsinPto As String, Optional oConexion1 As Connection) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set DiagnosticosSeleccionarXcodigoCIEsinPto = Nothing
        If oConexion1 Is Nothing Then
           SIGHEntidades.AbreConexionSIGH oConexion
        End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "DiagnosticosSeleccionarXcodigoCIEsinPto"
        Set oParameter = .CreateParameter("@CodigoCIEsinPto", adVarChar, adParamInput, 50, lcCodigoCIEsinPto): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        If oConexion1 Is Nothing Then
           Set oRsTmp1.ActiveConnection = Nothing
        End If
   End With
   Set DiagnosticosSeleccionarXcodigoCIEsinPto = oRsTmp1
   If oConexion1 Is Nothing Then
        oConexion.Close
        Set oRsTmp1 = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
    
End Function

'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function DiagnosticosSeleccionarXcodigoCIEsinPtoDevuelveDescripcion(lcCodigoCIEsinPto As String) As String
      Dim oRsTmp As New Recordset
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      Dim oConexion As New ADODB.Connection
          oConexion.CursorLocation = adUseClient
          oConexion.CommandTimeout = 300
          oConexion.Open SIGHEntidades.CadenaConexion
      DiagnosticosSeleccionarXcodigoCIEsinPtoDevuelveDescripcion = ""
          With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "DiagnosticosSeleccionarXcodigoCIEsinPtoDevuelveDescripcion"
             Set oParameter = .CreateParameter("@CodigoCIEsinPto", adVarChar, adParamInput, 50, lcCodigoCIEsinPto): .Parameters.Append oParameter
             Set oRsTmp = .Execute
             Set oRsTmp.ActiveConnection = Nothing
        End With
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
      If oRsTmp.RecordCount > 0 Then
         DiagnosticosSeleccionarXcodigoCIEsinPtoDevuelveDescripcion = oRsTmp.Fields!Descripcion
      End If
      oRsTmp.Close
      Set oRsTmp = Nothing
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function EmpleadosDevuelveNombre(lnIdEmpleado As Long, Optional oConexion1 As Connection) As String
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "EmpleadosDevuelveNombre"
        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, lnIdEmpleado): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        If oConexion1 Is Nothing Then
           Set oRsTmp1.ActiveConnection = Nothing
        End If
   End With
   EmpleadosDevuelveNombre = oRsTmp1.Fields!dEmpleado
   If oConexion1 Is Nothing Then
      oConexion.Close
   End If
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing

End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function TiposEmpleadosSeleccionarIdMedico(lnIdMedico As Long) As String
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposEmpleadosSeleccionarIdMedico"
        Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, lnIdMedico): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
    
    If oRsTmp1.RecordCount > 0 Then
       TiposEmpleadosSeleccionarIdMedico = oRsTmp1.Fields!Descripcion
    Else
       TiposEmpleadosSeleccionarIdMedico = ""
    End If
    
       oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function FactPuntosCargaSeleccionarPorIdServicio(lnIdServicio As Long) As Recordset
   On Error GoTo ErrCPSt
   Dim oFactPuntosCarga As New FactPuntosCarga
   Dim oConexion As New Connection
   oConexion.CommandTimeout = 300
   oConexion.CursorLocation = adUseClient
   oConexion.Open SIGHEntidades.CadenaConexion
   Set oFactPuntosCarga.Conexion = oConexion
   Set FactPuntosCargaSeleccionarPorIdServicio = oFactPuntosCarga.SeleccionarPorIdServicio(lnIdServicio)
   Exit Function
ErrCPSt:
   MsgBox Err.Description
End Function


Sub CatalogoBienesInsumosHospActualizaPrecioSegunIdProducto(lnIdProducto As Long, lnPrecioNuevo As Double)
    Dim oCatalogoBienesInsumos As New CatalogoBienesInsumos
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    Set oCatalogoBienesInsumos.Conexion = oConexion
    oCatalogoBienesInsumos.ActualizaPrecioSegunIdProducto lnIdProducto, lnPrecioNuevo
End Sub


Sub FactPuntosCargaActualiza1(lnIdServicio As Long, lnTipoServicio As Long, chkPuntoCarga As Integer, _
                             nombrePto As String, mi_opcion As sghOpciones)
    Dim oRsTmp As New Recordset
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim oDOFactPuntoCarga As New DOFactPuntoCarga, oFactPuntosCarga As New FactPuntosCarga
    On Error GoTo errPtoCarga
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    Set oFactPuntosCarga.Conexion = oConexion

    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "FactPuntosCargaSeleccionarIdServicio"
      Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
      Set oRsTmp = .Execute
      Set oRsTmp.ActiveConnection = Nothing
    End With
    Select Case mi_opcion
    Case sghAgregar
         If chkPuntoCarga = 1 And lnTipoServicio = 1 Then
            oDOFactPuntoCarga.IdPuntoCarga = 500 + lnIdServicio
            oDOFactPuntoCarga.Descripcion = nombrePto
            oDOFactPuntoCarga.IdUPS = 5
            oDOFactPuntoCarga.IdServicio = lnIdServicio
            If oFactPuntosCarga.Insertar(oDOFactPuntoCarga) = True Then
            End If
         End If
    Case sghModificar
         If chkPuntoCarga = 1 Then
            If oRsTmp.RecordCount > 0 Then
                oDOFactPuntoCarga.IdPuntoCarga = oRsTmp!IdPuntoCarga
                If oFactPuntosCarga.SeleccionarPorId(oDOFactPuntoCarga) = True Then
                   oDOFactPuntoCarga.Descripcion = nombrePto
                   If oFactPuntosCarga.modificar(oDOFactPuntoCarga) = True Then
                   End If
                End If
            Else
                oDOFactPuntoCarga.IdPuntoCarga = 500 + lnIdServicio
                oDOFactPuntoCarga.Descripcion = nombrePto
                oDOFactPuntoCarga.IdUPS = 5
                oDOFactPuntoCarga.IdServicio = lnIdServicio
                If oFactPuntosCarga.Insertar(oDOFactPuntoCarga) = True Then
                End If
            End If
         Else
            If oRsTmp.RecordCount > 0 Then
                oDOFactPuntoCarga.IdPuntoCarga = oRsTmp!IdPuntoCarga
                If oFactPuntosCarga.Eliminar(oDOFactPuntoCarga) = True Then
                End If
            End If
         End If
    Case sghEliminar
         If oRsTmp.RecordCount > 0 Then
                oDOFactPuntoCarga.IdPuntoCarga = oRsTmp!IdPuntoCarga
                If oFactPuntosCarga.Eliminar(oDOFactPuntoCarga) = True Then
                End If
         End If
    End Select
    oRsTmp.Close
    Set oRsTmp = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set oDOFactPuntoCarga = Nothing
    Set oFactPuntosCarga = Nothing
    Exit Sub
errPtoCarga:
    MsgBox Err.Description
    oConexion.Close
    Set oConexion = Nothing
End Sub

Sub AtencionesActualizaCreditoHospitalario(lnIdCuentaAtencion As Long)
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
  
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "AtencionesActualizaCreditoHospitalario"
            Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
            .Execute
        End With
        
        Set oParameter = Nothing
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing

End Sub
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function CatalogoBienesInsumosFiltrarDonacionesXcodigo(lcCodigo As String) As Recordset
        Dim lcSql As String, oRs As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
            
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CatalogoBienesInsumosFiltrarDonacionesXcodigo"
        Set oParameter = .CreateParameter("@lcCodigo", adVarChar, adParamInput, 50, lcCodigo): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set CatalogoBienesInsumosFiltrarDonacionesXcodigo = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing

        
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function RecetaCabeceraFiltraXcuentaYDocumentodespacho(lcDocumentoDespacho As String, lnIdCuentaAtencion As Long) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaCabeceraFiltraXcuentaYDocumentodespacho"
        Set oParameter = .CreateParameter("@DocumentoDespacho", adVarChar, adParamInput, 50, lcDocumentoDespacho): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set RecetaCabeceraFiltraXcuentaYDocumentodespacho = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing

End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function CentrosCostosSeleccionarXidentificador(lnIdCentroCosto As Long, Optional oConexion1 As Connection)
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
    End If


     With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "CentrosCostosSeleccionarXidentificador"
            Set oParameter = .CreateParameter("@IdCentroCosto", adInteger, adParamInput, 0, lnIdCentroCosto): .Parameters.Append oParameter
            Set oRs = .Execute
        End With
        Set CentrosCostosSeleccionarXidentificador = oRs
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing

'    If oConexion Is Nothing Then
'       oRs.Open lcSql, sighentidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
'    Else
'       oRs.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'    End If
'    Set CentrosCostosSeleccionarXidentificador = oRs
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function CatalogoServiciosSeleccionarXidentificador(lnIdProducto As Long, Optional oConexion1 As Connection) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        If oConexion1 Is Nothing Then
                oConexion.CursorLocation = adUseClient
                oConexion.CommandTimeout = 300
                oConexion.Open SIGHEntidades.CadenaConexion
        End If
                
         With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "CatalogoServiciosSeleccionarXidentificador"
            Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
            Set oRs = .Execute
            If oConexion1 Is Nothing Then
               Set oRs.ActiveConnection = Nothing
            End If
        End With
        Set CatalogoServiciosSeleccionarXidentificador = oRs
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing


End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function FactPartidasPresupuestalesXidentificador(lnIdPartidaPresupuestal As Long, Optional oConexion1 As Connection) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        If oConexion1 Is Nothing Then
                oConexion.CursorLocation = adUseClient
                oConexion.CommandTimeout = 300
                oConexion.Open SIGHEntidades.CadenaConexion
        End If
    

        With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "FactPartidasPresupuestalesXidentificador"
            Set oParameter = .CreateParameter("@IdPartidaPresupuestal", adInteger, adParamInput, 0, lnIdPartidaPresupuestal): .Parameters.Append oParameter
            Set oRs = .Execute
        End With
        Set FactPartidasPresupuestalesXidentificador = oRs
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
    
'    If oConexion Is Nothing Then
'       oRs.Open lcSql, sighentidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
'    Else
'       oRs.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'    End If
'    Set FactPartidasPresupuestalesXidentificador = oRs
End Function
Function lbElMedicoNOregistraDatosCE(lnIdServicio As Long) As String
    Dim oRsTmp1 As New Recordset
    lbElMedicoNOregistraDatosCE = "N"
    Set oRsTmp1 = ServiciosSeleccionarXidentificador(lnIdServicio)
    If oRsTmp1.RecordCount > 0 Then
       If oRsTmp1.Fields!UsaGalenHos = True Then
          lbElMedicoNOregistraDatosCE = "S"
       End If
    End If
    Set oRsTmp1 = Nothing
End Function
Function ServiciosSeleccionarXidentificador(lnIdServicio As Long, Optional oConexion1 As Connection) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        If oConexion1 Is Nothing Then
                oConexion.CursorLocation = adUseClient
                oConexion.CommandTimeout = 300
                oConexion.Open SIGHEntidades.CadenaConexion
        End If
    

        With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "ServiciosSeleccionarXidentificador"
            Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
            Set oRs = .Execute
        End With
        Set ServiciosSeleccionarXidentificador = oRs
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
    
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function EstablecimientosNoMinsaXidentificador(lnidEstablecimientoNoMinsa As Long, Optional oConexion1 As Connection) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        If oConexion1 Is Nothing Then
                oConexion.CursorLocation = adUseClient
                oConexion.CommandTimeout = 300
                oConexion.Open SIGHEntidades.CadenaConexion
        End If
    

        With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "EstablecimientosNoMinsaXidentificador"
            Set oParameter = .CreateParameter("@idEstablecimientoNoMinsa", adInteger, adParamInput, 0, lnidEstablecimientoNoMinsa): .Parameters.Append oParameter
            Set oRs = .Execute
        End With
        Set EstablecimientosNoMinsaXidentificador = oRs
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
    
'    If oConexion Is Nothing Then
'       oRs.Open lcSql, sighentidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
'    Else
'       oRs.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'    End If
'    Set EstablecimientosNoMinsaXidentificador = oRs
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function AtencionesSeleccionarXidentificador(lnIdAtencion As Long, Optional oConexion1 As Connection) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Set AtencionesSeleccionarXidentificador = Nothing
    
        If oConexion1 Is Nothing Then
                oConexion.CursorLocation = adUseClient
                oConexion.CommandTimeout = 300
                oConexion.Open SIGHEntidades.CadenaConexion
        End If
    
        With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "AtencionesSeleccionarXidentificador"
            Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
            Set oRs = .Execute
            If oConexion1 Is Nothing Then
               Set oRs.ActiveConnection = Nothing
            End If
        End With
        Set AtencionesSeleccionarXidentificador = oRs
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
    
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function BuscaDx(lcFechaAtencionInicial As String, lcFechaAtencionFinal As String, lnIdEspecialidad As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL1 As String: Dim sSQL2 As String
Dim oConexion As New ADODB.Connection
Dim oParameter As ADODB.Parameter
    Set BuscaDx = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion

     With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "BuscaDx"
         Set oParameter = .CreateParameter("@FechaAtencionInicial", adDBTimeStamp, adParamInput, 0, lcFechaAtencionInicial): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@FechaAtencionFinal", adDBTimeStamp, adParamInput, 0, lcFechaAtencionFinal): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@lnIdEspecialidad", adInteger, adParamInput, 0, lnIdEspecialidad): .Parameters.Append oParameter
         Set oRecordset = .Execute
         Set oRecordset.ActiveConnection = Nothing
    End With
    Set BuscaDx = oRecordset
    oConexion.Close
    Set oRecordset = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing

    Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function EstadosCuentaDevuelveTodos() As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
        Set EstadosCuentaDevuelveTodos = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EstadosCuentaDevuelveTodos"
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set EstadosCuentaDevuelveTodos = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function FactPuntosCargaDevuelveTodos() As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
        Set FactPuntosCargaDevuelveTodos = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
    
        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "FactPuntosCargaDevuelveTodos"
             Set oRs = .Execute
             Set oRs.ActiveConnection = Nothing
        End With
        Set FactPuntosCargaDevuelveTodos = oRs
        oConexion.Close
        Set oRs = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
         
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function DevuelveHIS_SITUACIOporDescripcion() As ADODB.Recordset
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
        Set DevuelveHIS_SITUACIOporDescripcion = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "DevuelveHIS_SITUACIOporDescripcion"
             Set oRsTmp = .Execute
             Set oRsTmp.ActiveConnection = Nothing
        End With
        Set DevuelveHIS_SITUACIOporDescripcion = oRsTmp
        oConexion.Close
        Set oRsTmp = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
    
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function EstablecimientosXidentificador(lnIdEstablecimiento As Long, oConexion As Connection) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
        
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EstablecimientosXidentificador"
        Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, lnIdEstablecimiento): .Parameters.Append oParameter
        Set oRs = .Execute
   End With
   Set EstablecimientosXidentificador = oRs
   Set oRs = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function RecetaCabeceraXidCuentaAtencion(ml_idCuentaAtencion As Long, Optional oConexion1 As Connection) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
       If oConexion1 Is Nothing Then
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open SIGHEntidades.CadenaConexion
      End If
        

        With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "RecetaCabeceraXidCuentaAtencion"
            Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_idCuentaAtencion): .Parameters.Append oParameter
            Set oRs = .Execute
        End With
        Set RecetaCabeceraXidCuentaAtencion = oRs
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
    

End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function DescripcionOcupacion(idTipoOcupacion As Long) As String
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim sSQL1 As String
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  oConexion.CommandTimeout = 300
  oConexion.CursorLocation = adUseClient
  oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DescripcionOcupacion"
        Set oParameter = .CreateParameter("@idTipoOcupacion", adInteger, adParamInput, 0, idTipoOcupacion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With

  If oRecordset.EOF = True And oRecordset.BOF = True Then
    
  Else
    DescripcionOcupacion = oRecordset!Descripcion
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
  End If
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Set oParameter = Nothing
  Exit Function
    
ManejadorDeError:
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function FactCatalogoServiciosPtosXidProducto(lnIdProducto As Long) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosPtosXidProducto"
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set FactCatalogoServiciosPtosXidProducto = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function CatalogoBienesSeleccionarPorIdYtipoFinanciamiento(mo_conexion As Connection, idProducto As Long, IdTipoFinanciamiento As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oRecordset As New Recordset
    

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_conexion
        .CommandTimeout = 150
        .CommandText = "CatalogoBienesSeleccionarPorIdYtipoFinanciamiento"
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, idProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TipoFinanciamiento", adInteger, adParamInput, 0, IdTipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CatalogoBienesSeleccionarPorIdYtipoFinanciamiento = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
   ms_MensajeError = ""
    Set oCommand = Nothing
Exit Function
ManejadorDeError:
   MsgBox Err.Description
Exit Function
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function CatalogoServicioSeleccionarPorIdYtipoFinanciamiento(mo_conexion As Connection, idProducto As Long, IdTipoFinanciamiento As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oRecordset As New Recordset
    

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_conexion
        .CommandTimeout = 150
        .CommandText = "CatalogoServicioSeleccionarPorIdYtipoFinanciamiento"
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, idProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IdTipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CatalogoServicioSeleccionarPorIdYtipoFinanciamiento = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing

    ms_MensajeError = ""
Exit Function
ManejadorDeError:
   MsgBox Err.Description
Exit Function
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function CatalogoBienesInsumosHospSeleccionarTodos(oConexion As Connection) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
        Set CatalogoBienesInsumosHospSeleccionarTodos = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CatalogoBienesInsumosHospSeleccionarTodos"
        Set oRs = .Execute
   End With
   Set CatalogoBienesInsumosHospSeleccionarTodos = oRs
   Set oRs = Nothing
   Set oCommand = Nothing
   
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function CatalogoServiciosSeleccionarTodos(Optional oConexion1 As Connection) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        If oConexion1 Is Nothing Then
            oConexion.CommandTimeout = 300
            oConexion.CursorLocation = adUseClient
            oConexion.Open SIGHEntidades.CadenaConexion
        End If


        With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "ConsultaFactCatalogoServiciosXSeleccionarTodos"
            Set oRs = .Execute
        End With
        Set CatalogoServiciosSeleccionarTodos = oRs
        Set oConexion = Nothing
        Set oCommand = Nothing
   
End Function

Function FactPuntosCargaSeleccionarPorIdDevDescripcion(lnIdPuntoCarga As Long) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    FactPuntosCargaSeleccionarPorIdDevDescripcion = ""
    ms_MensajeError = ""
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactPuntosCargaSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   FactPuntosCargaSeleccionarPorIdDevDescripcion = oRecordset.Fields("descripcion").value
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function DevuelveEstadoReceta(lnIdEstadoReceta As Long) As String
      Dim oRsTmp As New Recordset
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      Dim oConexion As New ADODB.Connection
          oConexion.CursorLocation = adUseClient
          oConexion.CommandTimeout = 300
          oConexion.Open SIGHEntidades.CadenaConexion
      

    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "DevuelveEstadoReceta"
         Set oParameter = .CreateParameter("@IdEstadoReceta", adInteger, adParamInput, 0, lnIdEstadoReceta): .Parameters.Append oParameter
         Set oRsTmp = .Execute
         Set oRsTmp.ActiveConnection = Nothing
    End With

      If oRsTmp.RecordCount > 0 Then
         DevuelveEstadoReceta = Trim(oRsTmp.Fields!Estado)
      Else
         DevuelveEstadoReceta = ""
      End If
      Set oRsTmp = Nothing
      Set oCommand = Nothing
      Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function FactCatalogoBienesInsumosHospXfiltro(lcFiltro As String, Optional oConexion1 As Connection) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "FactCatalogoBienesInsumosHospXfiltro"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRs = .Execute
        If oConexion1 Is Nothing Then
           Set oRs.ActiveConnection = Nothing
        End If
   End With
   Set FactCatalogoBienesInsumosHospXfiltro = oRs
   If oConexion1 Is Nothing Then
        oConexion.Close
        Set oRs = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing

End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function FactCatalogoServiciosHospXfiltro(lcFiltro As String, Optional oConexion1 As Connection) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosHospXfiltro"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRs = .Execute
        If oConexion1 Is Nothing Then
           Set oRs.ActiveConnection = Nothing
        End If
   End With
   Set FactCatalogoServiciosHospXfiltro = oRs
   If oConexion1 Is Nothing Then
        oConexion.Close
        Set oRs = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function DistritosSeleccionarXfiltro(lcFiltro As String) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
       
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        

        With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DistritosSeleccionarXfiltro"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set DistritosSeleccionarXfiltro = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Sub PreparaTablasDBFhis(ByRef oRsFox1 As Recordset, ByRef oRsFox2 As Recordset, ByRef oRsFox3 As Recordset, oConexionFox As Connection)
        Dim lcSql As String
        
        lcSql = "delete from hisLote"
        oRsFox3.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        oRsFox3.Open "select * from hisLote", oConexionFox, adOpenKeyset, adLockOptimistic
        
        
        lcSql = "delete from his1"
        oRsFox2.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        oRsFox2.Open "select * from his1", oConexionFox, adOpenKeyset, adLockOptimistic

        
        lcSql = "delete from hisA"
        oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        oRsFox1.Open "select * from hisA", oConexionFox, adOpenKeyset, adLockOptimistic
End Sub
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Sub PreparaTablasDBFsip2000(ByRef oRsFox1 As Recordset, oConexionFox As Connection)
    
        lcSql = "delete from gale_sip"
        oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        lcSql = "select * from gale_sip"
        oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic

End Sub
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Sub HisLoteXfiltro(ByRef oRsFox As Recordset, oConexionFox As Connection, lcFiltro As String)
   On Error Resume Next
   Dim lcSql As String
   lcSql = "select * from hisLote where " & lcFiltro
   If oRsFox.State = 1 Then oRsFox.Close
   oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
End Sub

'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function DevuelveCodigoEnTablaHIS_TabDiag(lcCodigoDx As String, oConexionFox As Connection) As String
    Dim lcSql As String, oRsTmp As New Recordset
    lcSql = "select * from tabDiag where Diag='" & Trim(lcCodigoDx) & "'"
    oRsTmp.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic

    DevuelveCodigoEnTablaHIS_TabDiag = "0"
    If oRsTmp.RecordCount > 0 Then
       DevuelveCodigoEnTablaHIS_TabDiag = Trim(Str(oRsTmp.Fields!Codigo))
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function
'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Sub AbreTablaFox(ByRef oRsFox As Recordset, oConexionFox As Connection, lcTablaDBF As String, lcFiltro As String, lcOrder As String)
        Dim lcSql As String
    
        lcSql = "SELECT *  FROM " & lcTablaDBF & _
                IIf(lcFiltro = "", "", " where " & lcFiltro) & _
                IIf(lcOrder = "", "", " order by " & lcOrder)
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
End Sub

Function HIS_ESTABseleccionarXcodigo(lcCodigoEstablecimiento As String, oConexion As Connection) As Recordset
    '    Dim lcSql As String
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set HIS_ESTABseleccionarXcodigo = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HIS_ESTABseleccionarXcodigo"
        Set oParameter = .CreateParameter("@CodigoEstablecimiento", adVarChar, adParamInput, 200, lcCodigoEstablecimiento): .Parameters.Append oParameter
        Set oRs = .Execute
    End With
    Set HIS_ESTABseleccionarXcodigo = oRs
    Set oCommand = Nothing
End Function

'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function TiposEstadoCivilXidEstadoCivil(lnIdEstadoCivil As Long, oConexion As Connection) As Recordset
    '    Dim lcSql As String
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set TiposEstadoCivilXidEstadoCivil = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposEstadoCivilXidEstadoCivil"
        Set oParameter = .CreateParameter("@IdEstadoCivil", adInteger, adParamInput, 0, lnIdEstadoCivil): .Parameters.Append oParameter
        Set oRs = .Execute
    End With
    Set TiposEstadoCivilXidEstadoCivil = oRs
    Set oCommand = Nothing
End Function

'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function TiposGradoInstruccionXidGradoInstruccion(lnIdGradoInstruccion As Long, oConexion As Connection) As Recordset
    '    Dim lcSql As String
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    

   
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposGradoInstruccionXidGradoInstruccion"
        Set oParameter = .CreateParameter("@IdGradoInstruccion", adInteger, adParamInput, 0, lnIdGradoInstruccion): .Parameters.Append oParameter
        Set oRs = .Execute
    End With
            
    Set TiposGradoInstruccionXidGradoInstruccion = oRs
    Set oCommand = Nothing
End Function

'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function ServiciosXEspecialidadYtipoServicio(lnIdEspecialidad As Long, lnIdTipoServicio As Long, oConexion As Connection) As Recordset
    '    Dim lcSql As String
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set ServiciosXEspecialidadYtipoServicio = Nothing
   
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosXEspecialidadYtipoServicio"
        Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, lnIdEspecialidad): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, lnIdTipoServicio): .Parameters.Append oParameter
        Set oRs = .Execute
    End With
            
    Set ServiciosXEspecialidadYtipoServicio = oRs
    Set oCommand = Nothing
End Function

Sub ActualizaIdPacienteEnTodasLasTablasSegunNroCuenta(ml_IdPacienteNew As Long, lnIdCuentaAtencion As Long, _
                                                      lnIdAtencion As Long, lnIdCama As Long, ml_idUsuario As Long, _
                                                      mo_lcNombrePc As String, lcObservacionesAuditoria As String)
        On Error GoTo ErrorProceso
        Dim oRsTmp As New Recordset
        Dim oConexion As New Connection
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        
        oConexion.BeginTrans
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ActualizaIdPacienteEnTodasLasTablasSegunNroCuenta"
            Set oParameter = .CreateParameter("@IdPacienteNew", adInteger, adParamInput, 0, ml_IdPacienteNew): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdCama", adInteger, adParamInput, 0, lnIdCama): .Parameters.Append oParameter
            .Execute
        End With
        Set oParameter = Nothing
        
        Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "M", lnIdAtencion, "Atenciones", oConexion, 500 + 129, mo_lcNombrePc, lcObservacionesAuditoria)    '500+ ListBarReporte.idReporte
        
        oConexion.CommitTrans
        oConexion.Close
        Exit Sub
ErrorProceso:
'        oConexion.RollbackTrans
        MsgBox Err.Description
End Sub

Sub AuditoriaInsertaDesdeReniec(lcDniAbuscar As String, lcObservaciones As String, lcFechaServidor As String)
        On Error GoTo ErrorProceso
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "AuditoriaInsertaDesdeReniec"
            Set oParameter = .CreateParameter("@FechaServidor", adDBTimeStamp, adParamInput, 0, Format(lcFechaServidor, "dd/mm/yyyy")): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@DniAbuscar", adInteger, adParamInput, 0, CLng(lcDniAbuscar)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, SIGHEntidades.Usuario): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@NombrePC", adVarChar, adParamInput, 30, SIGHEntidades.RetornaNombrePC): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Observaciones", adVarChar, adParamInput, 100, lcObservaciones): .Parameters.Append oParameter
            .Execute
        End With
        Set oParameter = Nothing
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing

        Exit Sub
ErrorProceso:
        MsgBox Err.Description
End Sub

'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function FactPartidasPresupuestalesSeleccionarTodas(oConexion As Connection) As Recordset
'    Dim lcSql As String
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Set FactPartidasPresupuestalesSeleccionarTodas = Nothing
    

    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactPartidasPresupuestalesSeleccionarTodas"
        Set oRs = .Execute
    End With
            
    Set FactPartidasPresupuestalesSeleccionarTodas = oRs
    Set oCommand = Nothing

End Function

'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function FactCatalogoServiciosHospXidProductoOrdenadoXtipofinanciamiento(lnIdProducto As Long) As Recordset
'    Dim lcSql As String
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    
    

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosHospXidProductoOrdenadoXtipofinanciamiento"
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
            
    Set FactCatalogoServiciosHospXidProductoOrdenadoXtipofinanciamiento = oRs
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing

End Function

'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function FactCatalogoServiciosPtosOrdenasXcpt() As Recordset
'    Dim lcSql As String
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Set FactCatalogoServiciosPtosOrdenasXcpt = Nothing
    

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosPtosOrdenasXcpt"
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
            
    Set FactCatalogoServiciosPtosOrdenasXcpt = oRs
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function FactCatalogoServiciosHospXidProducto(lnIdProducto As Long) As Recordset
'    Dim lcSql As String
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Set FactCatalogoServiciosHospXidProducto = Nothing
    

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosHospXidProducto"
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
            
    Set FactCatalogoServiciosHospXidProducto = oRs
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function TiposTarifaSoloFarmacia(oConexion As Connection) As Recordset
'    Dim lcSql As String
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Set TiposTarifaSoloFarmacia = Nothing
    

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposTarifaSoloFarmacia"
        Set oRs = .Execute
    End With
    Set TiposTarifaSoloFarmacia = oRs
    Set oCommand = Nothing
End Function

Function CatalogoServiciosSeleccionarTodosOrdenadoXnombre() As Recordset
'    Dim lcSql As String
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
  

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CatalogoServiciosSeleccionarTodosOrdenadoXnombre"
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    
    Set CatalogoServiciosSeleccionarTodosOrdenadoXnombre = oRs
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'Modificado por Yamill Palomino 11-10-13 Se cambio a Store Procedure
Function RecetaDetalleSeleccioarPorIdReceta(lnIdReceta As Long, oConexion As Connection) As Recordset
'    Dim lcSql As String
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set RecetaDetalleSeleccioarPorIdReceta = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaDetalleSeleccioarPorIdReceta"
        Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
        Set oRs = .Execute
    End With

    Set RecetaDetalleSeleccioarPorIdReceta = oRs
    Set oCommand = Nothing
End Function


Function UPServiciosSeleccionarTodos() As Recordset
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
  
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "UPServiciosSeleccionarTodos"
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    
    Set UPServiciosSeleccionarTodos = oRs
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function


Sub UPServiciosActualizar(lcUPScodigo As String, lcUPSdescripcion As String, oConexion As ADODB.Connection)
   
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "UPServiciosActualizar"
        Set oParameter = .CreateParameter("@UPScodigo", adVarChar, adParamInput, 6, lcUPScodigo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@UPSdescripcion", adVarChar, adParamInput, 50, lcUPSdescripcion): .Parameters.Append oParameter
        .Execute
    End With
    Set oCommand = Nothing
End Sub

Sub ParametrosActualizaValorTextoXid(lnIdParametro As Long, lcValorTexto As String)
    
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ParametrosActualizaValorTextoXid"
        Set oParameter = .CreateParameter("@idParametro", adInteger, adParamInput, 0, lnIdParametro): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ValorTexto", adVarChar, adParamInput, 150, lcValorTexto): .Parameters.Append oParameter
        .Execute
    End With
    Set oCommand = Nothing
    Set oConexion = Nothing
End Sub



Function FactCatalogoServiciosHospPorIdProducto(lnIdProducto As Long, oConexion As Connection) As Recordset
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosHospPorIdProducto"
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRs = .Execute
    End With
    Set FactCatalogoServiciosHospPorIdProducto = oRs
    Set oCommand = Nothing
End Function


Function RecetaCabeceraPorRangoFechas(ldFechaInicio As Date, ldFechaFin As Date, lcApellidoPaterno As String, _
                                     lnIdEstado As Long) As Recordset
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaCabeceraPorRangoFechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@apellidoP", adVarChar, adParamInput, 20, lcApellidoPaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, lnIdEstado): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    Set RecetaCabeceraPorRangoFechas = oRs
    Set oCommand = Nothing
    Set oConexion = Nothing
End Function


Function ConsumoMedicamentosPorIdCuentaConSeguroYparticular(lnIdCuentaAtencion As Long) As Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set ConsumoMedicamentosPorIdCuentaConSeguroYparticular = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

       With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ConsumoMedicamentosPorIdCuentaConSeguroYparticular"
            Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set ConsumoMedicamentosPorIdCuentaConSeguroYparticular = oRecordset
       oConexion.Close
       Set oRecordset = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function


'<<<<<<<<  FUNCION  Y METODOS - MODULO PROGRAMAS DE SALUD >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Function ProControlConfigDatoSeleccionarPorIdPrograma(lnIdPrograma As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set ProControlConfigDatoSeleccionarPorIdPrograma = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProControlConfigDatoSeleccionarPorIdPrograma"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set ProControlConfigDatoSeleccionarPorIdPrograma = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function ProCabeceraConfigDatosSeleccionarPorIdPrograma(lnIdPrograma As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set ProCabeceraConfigDatosSeleccionarPorIdPrograma = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProCabeceraConfigDatosSeleccionarPorIdPrograma"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set ProCabeceraConfigDatosSeleccionarPorIdPrograma = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function ProCatalogoDiagnosticosSeleccionarPorIdPrograma(lnIdPrograma As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set ProCatalogoDiagnosticosSeleccionarPorIdPrograma = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProCatalogoDiagnosticosSeleccionarPorIdPrograma"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set ProCatalogoDiagnosticosSeleccionarPorIdPrograma = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function TiposResultadoServInterm(oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Set TiposResultadoServInterm = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposResultadoServInterm"
        Set oRsTmp = .Execute
   End With
   Set TiposResultadoServInterm = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
End Function

Function ProCatalogoControlesPorIdPrograma(lnIdPrograma As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set ProCatalogoControlesPorIdPrograma = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProCatalogoControlesPorIdPrograma"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set ProCatalogoControlesPorIdPrograma = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function ProCatalogoTratamientosSeleccionarPorIdPrograma(lnIdPrograma As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set ProCatalogoTratamientosSeleccionarPorIdPrograma = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProCatalogoTratamientosSeleccionarPorIdPrograma"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set ProCatalogoTratamientosSeleccionarPorIdPrograma = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function ProConsultarCabeceraControlPorIdPacienteIdAtencion(lnIdPrograma As Long, lnIdPaciente As Long, lnIdAtencion As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set ProConsultarCabeceraControlPorIdPacienteIdAtencion = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProConsultarCabeceraControlPorIdPacienteIdAtencion"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set ProConsultarCabeceraControlPorIdPacienteIdAtencion = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function ProDiagnosticosSeleccionarPorIdControl(lnIdPrograma As Long, lnIdProCabecera As Long, lnIdControl As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set ProDiagnosticosSeleccionarPorIdControl = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProDiagnosticosSeleccionarPorIdControl"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, lnIdProCabecera): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, lnIdControl): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set ProDiagnosticosSeleccionarPorIdControl = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function ProProcedimientosSeleccionarPorIdControl(lnIdPrograma As Long, lnIdProCabecera As Long, lnIdControl As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set ProProcedimientosSeleccionarPorIdControl = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProProcedimientosSeleccionarPorIdControl"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, lnIdProCabecera): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, lnIdControl): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set ProProcedimientosSeleccionarPorIdControl = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function ProTratamientosSeleccionarPorIdControl(lnIdPrograma As Long, lnIdProCabecera As Long, lnIdControl As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set ProTratamientosSeleccionarPorIdControl = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProTratamientosSeleccionarPorIdControl"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, lnIdProCabecera): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, lnIdControl): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set ProTratamientosSeleccionarPorIdControl = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function ProConsultarActualProgramaPaciente(lnIdPrograma As Long, lnIdPaciente As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set ProConsultarActualProgramaPaciente = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProConsultarActualProgramaPaciente"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set ProConsultarActualProgramaPaciente = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function ProConsultarDatosCabecera(lnIdPrograma As Long, lnIdProCabecera As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set ProConsultarDatosCabecera = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProConsultarDatosCabecera"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, lnIdProCabecera): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set ProConsultarDatosCabecera = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function ProConsultarControles(lnIdPrograma As Long, lnIdProCabecera As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set ProConsultarControles = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProConsultarControles"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, lnIdProCabecera): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set ProConsultarControles = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function ProControlesSeleccionarPorId(lnIdPrograma As Long, lnIdProCabecera As Long, lnIdControl As Long, oConexion As Connection) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set ProControlesSeleccionarPorId = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProControlesSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, lnIdProCabecera): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, lnIdControl): .Parameters.Append oParameter
        Set oRsTmp = .Execute
    End With
    
    Set ProControlesSeleccionarPorId = oRsTmp
    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Function ProControlDatoSeleccionarPorId(lnIdPrograma As Long, lnIdProCabecera As Long, lnIdControl As Long, oConexion As Connection) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set ProControlDatoSeleccionarPorId = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProControlDatoSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, lnIdProCabecera): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, lnIdControl): .Parameters.Append oParameter
        Set oRsTmp = .Execute
    End With
    
    Set ProControlDatoSeleccionarPorId = oRsTmp
    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Function ProCatalogoProcedimientosSelecionarPorIdProgramaIdDiagnostico(lnIdPrograma As Long, lnIdDiagnostico As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set ProCatalogoProcedimientosSelecionarPorIdProgramaIdDiagnostico = Nothing
     With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "ProCatalogoProcedimientosSelecionarPorIdProgramaIdDiagnostico"
         Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, lnIdDiagnostico): .Parameters.Append oParameter
         Set oRsTmp = .Execute
    End With
    Set ProCatalogoProcedimientosSelecionarPorIdProgramaIdDiagnostico = oRsTmp
    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

'Actualizado 28102014
Function ProgramaControlesAgregar(oDoProCabecera As DoProCabecera, oRsDatosProCabecera As Recordset, oDOProControles As DOProControles, oRsDatosProControles As Recordset, oRsProDiagnosticos As Recordset, oRsProProcedimientos As Recordset, oRsProTratamientos As Recordset, lnIdAtencion As Long, oRsCatalogoControles As Recordset) As Boolean
    On Error GoTo ErrCeAgr
    Dim oConexion As New ADODB.Connection
    
    Dim oProCabecera As New ProCabecera
    Dim oProCabeceraDato As New ProCabeceraDato, oDOProCabeceraDato As New DOProCabeceraDato
    Dim oProControles As New ProControles
    Dim oProControlDato As New ProControlDato, oDoProControlDato As New DoProControlDato
    Dim oProDiagnosticos As New ProDiagnosticos, oDOProDiagnosticos As New DOProDiagnosticos
    Dim oProProcedimientos As New ProProcedimientos, oDOProProcedimientos As New DOProProcedimientos
    Dim oProTratamientos As New ProTratamientos, oDOProTratamientos As New DOProTratamientos
    Dim oRsTmpAutomatico As New Recordset, oRsTmpCodigoHis As New Recordset
    Dim oDOAtencionDiagnostico As New DOAtencionDiagnostico, oAtencionesDiagnosticos As New AtencionesDiagnosticos
    Dim lbContinuar As Boolean
    Dim oDOProControlesTemp As New DOProControles
    
    ProgramaControlesAgregar = False
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    '
    oConexion.BeginTrans
    '
    Set oProCabecera.Conexion = oConexion
    Set oProCabeceraDato.Conexion = oConexion
    Set oProControles.Conexion = oConexion
    Set oProControlDato.Conexion = oConexion
    Set oProDiagnosticos.Conexion = oConexion
    Set oAtencionesDiagnosticos.Conexion = oConexion
    Set oProProcedimientos.Conexion = oConexion
    Set oProTratamientos.Conexion = oConexion

    'Inserta en la tabla ProCabecera y Obtiene el Id
    If Not oProCabecera.Insertar(oDoProCabecera) Then
        ms_MensajeError = oProCabecera.MensajeError: GoTo ErrCeAgr
    End If
    'Inserta los datos configurados de la cabecera
    If oRsDatosProCabecera.RecordCount > 0 Then
        oRsDatosProCabecera.MoveFirst
        Do While Not oRsDatosProCabecera.EOF
           oDOProCabeceraDato.IdPrograma = oDoProCabecera.IdPrograma
           oDOProCabeceraDato.IdProCabecera = oDoProCabecera.IdProCabecera
           oDOProCabeceraDato.IdCabDato = oRsDatosProCabecera.Fields!IdCabDato
           oDOProCabeceraDato.CabDato = oRsDatosProCabecera.Fields!CabDato
           oDOProCabeceraDato.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
           If Not oProCabeceraDato.Insertar(oDOProCabeceraDato) Then
             ms_MensajeError = oProCabeceraDato.MensajeError: GoTo ErrCeAgr
           End If
           oRsDatosProCabecera.MoveNext
        Loop
    End If
    'Completa IdProCabecera en DoProControles
    oDOProControles.IdProCabecera = oDoProCabecera.IdProCabecera
    oDOProControles.idAtencion = oDOProControles.idAtencion
    oDOProControles.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
    'Inserta en la tabla ProControles
    If Not oProControles.Insertar(oDOProControles) Then
        ms_MensajeError = oProControles.MensajeError: GoTo ErrCeAgr
    End If
    
    'Insertar los controles que se realizaron en otro establecimiento
    'Actualizado 28102014
    If oRsCatalogoControles.RecordCount > 0 Then
        oRsCatalogoControles.MoveFirst
        Do While Not oRsCatalogoControles.EOF
            If oRsCatalogoControles.Fields!FechaControl = "" Then
                If oDOProControles.IdControl <> oRsCatalogoControles.Fields!IdControl Then
                    oDOProControlesTemp.IdPrograma = oDoProCabecera.IdPrograma
                    oDOProControlesTemp.IdProCabecera = oDoProCabecera.IdProCabecera
                    oDOProControlesTemp.IdControl = oRsCatalogoControles.Fields!IdControl
                    oDOProControlesTemp.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
                    If oProControles.SeleccionarPorId(oDOProControlesTemp) Then
                        If Not oProControles.Eliminar(oDOProControlesTemp) Then
                            ms_MensajeError = oProControles.MensajeError: GoTo ErrCeAgr
                        End If
                    End If
                End If
            Else
                If oRsCatalogoControles.Fields!ControlOtroEESS = True Then
                    oDOProControlesTemp.IdPrograma = oDoProCabecera.IdPrograma
                    oDOProControlesTemp.IdProCabecera = oDoProCabecera.IdProCabecera
                    oDOProControlesTemp.IdControl = oRsCatalogoControles.Fields!IdControl
                    
                    If oProControles.SeleccionarPorId(oDOProControlesTemp) Then
                        oDOProControlesTemp.IdPrograma = oDoProCabecera.IdPrograma
                        oDOProControlesTemp.IdProCabecera = oDoProCabecera.IdProCabecera
                        oDOProControlesTemp.IdControl = oRsCatalogoControles.Fields!IdControl
                        oDOProControlesTemp.idAtencion = 0
                        oDOProControlesTemp.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
                        oDOProControlesTemp.ControlOtroEESS = oRsCatalogoControles.Fields!ControlOtroEESS
                        oDOProControlesTemp.IdEstablecimiento = oRsCatalogoControles.Fields!IdEstablecimiento
                        oDOProControlesTemp.FechaControl = oRsCatalogoControles.Fields!FechaControl
                        If Not oProControles.modificar(oDOProControlesTemp) Then
                            ms_MensajeError = oProControles.MensajeError: GoTo ErrCeAgr
                        End If
                    Else
                        oDOProControlesTemp.IdPrograma = oDoProCabecera.IdPrograma
                        oDOProControlesTemp.IdProCabecera = oDoProCabecera.IdProCabecera
                        oDOProControlesTemp.IdControl = oRsCatalogoControles.Fields!IdControl
                        oDOProControlesTemp.idAtencion = 0
                        oDOProControlesTemp.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
                        oDOProControlesTemp.ControlOtroEESS = oRsCatalogoControles.Fields!ControlOtroEESS
                        oDOProControlesTemp.IdEstablecimiento = oRsCatalogoControles.Fields!IdEstablecimiento
                        oDOProControlesTemp.FechaControl = oRsCatalogoControles.Fields!FechaControl
                        If Not oProControles.Insertar(oDOProControlesTemp) Then
                            ms_MensajeError = oProControles.MensajeError: GoTo ErrCeAgr
                        End If
                    End If
                End If
           End If
            oRsCatalogoControles.MoveNext
        Loop
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Inserta los datos configurados del nuevo control
    If oRsDatosProControles.RecordCount > 0 Then
        oRsDatosProControles.MoveFirst
        Do While Not oRsDatosProControles.EOF
           oDoProControlDato.IdPrograma = oDoProCabecera.IdPrograma
           oDoProControlDato.IdProCabecera = oDoProCabecera.IdProCabecera
           oDoProControlDato.IdControl = oDOProControles.IdControl
           oDoProControlDato.IdControlDato = oRsDatosProControles.Fields!IdControlDato
           oDoProControlDato.ControlDato = oRsDatosProControles.Fields!ControlDato
           oDoProControlDato.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
           If Not oProControlDato.Insertar(oDoProControlDato) Then
             ms_MensajeError = oProControlDato.MensajeError: GoTo ErrCeAgr
           End If
           oRsDatosProControles.MoveNext
        Loop
    End If
    
    'Inserta Diagnosticos
    If oRsProDiagnosticos.RecordCount > 0 Then
        oRsProDiagnosticos.MoveFirst
        Do While Not oRsProDiagnosticos.EOF
           oDOProDiagnosticos.IdPrograma = oDoProCabecera.IdPrograma
           oDOProDiagnosticos.IdProCabecera = oDoProCabecera.IdProCabecera
           oDOProDiagnosticos.IdControl = oDOProControles.IdControl
           oDOProDiagnosticos.IdDiagnostico = oRsProDiagnosticos.Fields!IdDiagnostico
           oDOProDiagnosticos.Principal = oRsProDiagnosticos.Fields!Principal
           oDOProDiagnosticos.labConfHIS = oRsProDiagnosticos.Fields!labConfHIS
           oDOProDiagnosticos.IdSubclasificacionDx = oRsProDiagnosticos.Fields!IdSubclasificacionDx
           oDOProDiagnosticos.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
           If Not oProDiagnosticos.Insertar(oDOProDiagnosticos) Then
             ms_MensajeError = oProDiagnosticos.MensajeError: GoTo ErrCeAgr
           End If
           
           'Actualizado 15102014
           oDOAtencionDiagnostico.idAtencion = lnIdAtencion
           oDOAtencionDiagnostico.IdDiagnostico = oRsProDiagnosticos.Fields!IdDiagnostico
           If Not oAtencionesDiagnosticos.SeleccionarPorIdAtencionIdDx(oDOAtencionDiagnostico) Then
                oDOAtencionDiagnostico.IdClasificacionDx = 1
                oDOAtencionDiagnostico.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
                oDOAtencionDiagnostico.labConfHIS = Space(3)
                
                oDOAtencionDiagnostico.IdSubclasificacionDx = oRsProDiagnosticos.Fields!IdSubclasificacionDx
'                If oRsProDiagnosticos.Fields!Principal Then
'                     oDOAtencionDiagnostico.IdSubClasificacionDX = 102 'Definitivo
'                Else
'                     oDOAtencionDiagnostico.IdSubClasificacionDX = 101 'Presuntivo
'                End If
                If Not oAtencionesDiagnosticos.Insertar(oDOAtencionDiagnostico) Then
                     ms_MensajeError = oAtencionesDiagnosticos.MensajeError: GoTo ErrCeAgr
                End If
           End If
           
           oRsProDiagnosticos.MoveNext
        Loop
    End If
    
    'Inserta Procedimientos
    If oRsProProcedimientos.RecordCount > 0 Then
        oRsProProcedimientos.MoveFirst
        Do While Not oRsProProcedimientos.EOF
           oDOProProcedimientos.IdPrograma = oDoProCabecera.IdPrograma
           oDOProProcedimientos.IdProCabecera = oDoProCabecera.IdProCabecera
           oDOProProcedimientos.IdControl = oDOProControles.IdControl
           oDOProProcedimientos.IdDiagnostico = oRsProProcedimientos.Fields!IdDiagnostico
           oDOProProcedimientos.idProducto = oRsProProcedimientos.Fields!idProducto
           oDOProProcedimientos.IdResultado = oRsProProcedimientos.Fields!IdResultado
           oDOProProcedimientos.labConfHIS = oRsProProcedimientos.Fields!labConfHIS
           oDOProProcedimientos.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
           If Not oProProcedimientos.Insertar(oDOProProcedimientos) Then
             ms_MensajeError = oProProcedimientos.MensajeError: GoTo ErrCeAgr
           End If
           oRsProProcedimientos.MoveNext
        Loop
    End If
    
        'Inserta Tratamientos
    If oRsProTratamientos.RecordCount > 0 Then
        oRsProTratamientos.MoveFirst
        Do While Not oRsProTratamientos.EOF
           oDOProTratamientos.IdPrograma = oDoProCabecera.IdPrograma
           oDOProTratamientos.IdProCabecera = oDoProCabecera.IdProCabecera
           oDOProTratamientos.IdControl = oDOProControles.IdControl
           oDOProTratamientos.idProducto = oRsProTratamientos.Fields!idProducto
           oDOProTratamientos.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
           If Not oProTratamientos.Insertar(oDOProTratamientos) Then
                ms_MensajeError = oProTratamientos.MensajeError: GoTo ErrCeAgr
           End If
           oRsProTratamientos.MoveNext
        Loop
    End If

    ProgramaControlesAgregar = True
    oConexion.Close
ErrCeAgr:
'Resume
    If ProgramaControlesAgregar = True Then
       oConexion.CommitTrans
    Else
       oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oProCabecera = Nothing
    Set oProCabeceraDato = Nothing
    Set oDOProCabeceraDato = Nothing
    Set oProControles = Nothing
    Set oProControlDato = Nothing
    Set oDoProControlDato = Nothing
    Set oProDiagnosticos = Nothing
    Set oDOProDiagnosticos = Nothing
    Set oAtencionesDiagnosticos = Nothing
    Set oDOAtencionDiagnostico = Nothing
    Set oProProcedimientos = Nothing
    Set oDOProProcedimientos = Nothing
    Set oProTratamientos = Nothing
    Set oDOProTratamientos = Nothing
End Function

Function ProgramaControlesModificar(oDoProCabecera As DoProCabecera, oRsDatosProCabecera As Recordset, oDOProControles As DOProControles, oRsDatosProControles As Recordset, oRsProDiagnosticos As Recordset, oRsProProcedimientos As Recordset, oRsProTratamientos As Recordset, lnIdAtencion As Long, lbActualizarControl As Boolean, oRsCatalogoControles As Recordset) As Boolean
    On Error GoTo ErrCeAgr
    Dim oConexion As New ADODB.Connection
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
  
    Dim oProCabeceraDato As New ProCabeceraDato, oDOProCabeceraDato As New DOProCabeceraDato
    Dim oProControles As New ProControles
    Dim oProControlDato As New ProControlDato, oDoProControlDato As New DoProControlDato
    Dim oProDiagnosticos As New ProDiagnosticos, oDOProDiagnosticos As New DOProDiagnosticos
    Dim oProProcedimientos As New ProProcedimientos, oDOProProcedimientos As New DOProProcedimientos
    Dim oProTratamientos As New ProTratamientos, oDOProTratamientos As New DOProTratamientos
    Dim oRsTmpAutomatico As New Recordset, oRsTmpCodigoHis As New Recordset
    Dim oDOAtencionDiagnostico As New DOAtencionDiagnostico, oAtencionesDiagnosticos As New AtencionesDiagnosticos
    Dim lbContinuar As Boolean
    Dim lbExisteDatoCabecera As Boolean
    Dim oDOProControlesTemp As New DOProControles
    
    ProgramaControlesModificar = False
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    '
    oConexion.BeginTrans
    '
    Set oProCabeceraDato.Conexion = oConexion
    Set oProControles.Conexion = oConexion
    Set oProControlDato.Conexion = oConexion
    Set oProDiagnosticos.Conexion = oConexion
    Set oAtencionesDiagnosticos.Conexion = oConexion
    Set oProProcedimientos.Conexion = oConexion
    Set oProTratamientos.Conexion = oConexion

    'Actualiza los datos configurados de la cabecera
    If oRsDatosProCabecera.RecordCount > 0 Then
        oRsDatosProCabecera.MoveFirst
        Do While Not oRsDatosProCabecera.EOF
           oDOProCabeceraDato.IdPrograma = oDoProCabecera.IdPrograma
           oDOProCabeceraDato.IdProCabecera = oDoProCabecera.IdProCabecera
           oDOProCabeceraDato.IdCabDato = oRsDatosProCabecera.Fields!IdCabDato
           
           lbExisteDatoCabecera = False
           If oProCabeceraDato.SeleccionarPorId(oDOProCabeceraDato) Then lbExisteDatoCabecera = True
           
           oDOProCabeceraDato.CabDato = oRsDatosProCabecera.Fields!CabDato
           oDOProCabeceraDato.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
           
           If lbExisteDatoCabecera = True Then
                If Not oProCabeceraDato.modificar(oDOProCabeceraDato) Then
                  ms_MensajeError = oProCabeceraDato.MensajeError: GoTo ErrCeAgr
                End If
           Else
                If Not oProCabeceraDato.Insertar(oDOProCabeceraDato) Then
                  ms_MensajeError = oProCabeceraDato.MensajeError: GoTo ErrCeAgr
                End If
           End If
           oRsDatosProCabecera.MoveNext
        Loop
    End If
    'Completa IdProCabecera en DoProControles
    oDOProControles.IdProCabecera = oDoProCabecera.IdProCabecera
    oDOProControles.idAtencion = oDOProControles.idAtencion
    oDOProControles.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
    'Inserta en la tabla ProControles
    If lbActualizarControl = True Then
        If Not oProControles.EliminarDatosControl(oDOProControles) Then
            ms_MensajeError = oProControles.MensajeError: GoTo ErrCeAgr
        End If
        If Not oProControles.modificar(oDOProControles) Then
            ms_MensajeError = oProControles.MensajeError: GoTo ErrCeAgr
        End If
    Else
        If Not oProControles.Insertar(oDOProControles) Then
            ms_MensajeError = oProControles.MensajeError: GoTo ErrCeAgr
        End If
    End If
    
    'Insertar los controles que se realizaron en otro establecimiento
    'Actualizado 28102014
    If oRsCatalogoControles.RecordCount > 0 Then
        oRsCatalogoControles.MoveFirst
        Do While Not oRsCatalogoControles.EOF
            If oRsCatalogoControles.Fields!FechaControl = "" Then
                If oDOProControles.IdControl <> oRsCatalogoControles.Fields!IdControl Then
                    oDOProControlesTemp.IdPrograma = oDoProCabecera.IdPrograma
                    oDOProControlesTemp.IdProCabecera = oDoProCabecera.IdProCabecera
                    oDOProControlesTemp.IdControl = oRsCatalogoControles.Fields!IdControl
                    oDOProControlesTemp.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
                    If oProControles.SeleccionarPorId(oDOProControlesTemp) Then
                        If Not oProControles.Eliminar(oDOProControlesTemp) Then
                            ms_MensajeError = oProControles.MensajeError: GoTo ErrCeAgr
                        End If
                    End If
                End If
            Else
                If oRsCatalogoControles.Fields!ControlOtroEESS = True Then
                    oDOProControlesTemp.IdPrograma = oDoProCabecera.IdPrograma
                    oDOProControlesTemp.IdProCabecera = oDoProCabecera.IdProCabecera
                    oDOProControlesTemp.IdControl = oRsCatalogoControles.Fields!IdControl
                    
                    If oProControles.SeleccionarPorId(oDOProControlesTemp) Then
                        oDOProControlesTemp.IdPrograma = oDoProCabecera.IdPrograma
                        oDOProControlesTemp.IdProCabecera = oDoProCabecera.IdProCabecera
                        oDOProControlesTemp.IdControl = oRsCatalogoControles.Fields!IdControl
                        oDOProControlesTemp.idAtencion = 0
                        oDOProControlesTemp.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
                        oDOProControlesTemp.ControlOtroEESS = oRsCatalogoControles.Fields!ControlOtroEESS
                        oDOProControlesTemp.IdEstablecimiento = oRsCatalogoControles.Fields!IdEstablecimiento
                        oDOProControlesTemp.FechaControl = oRsCatalogoControles.Fields!FechaControl
                        If Not oProControles.modificar(oDOProControlesTemp) Then
                            ms_MensajeError = oProControles.MensajeError: GoTo ErrCeAgr
                        End If
                    Else
                        oDOProControlesTemp.IdPrograma = oDoProCabecera.IdPrograma
                        oDOProControlesTemp.IdProCabecera = oDoProCabecera.IdProCabecera
                        oDOProControlesTemp.IdControl = oRsCatalogoControles.Fields!IdControl
                        oDOProControlesTemp.idAtencion = 0
                        oDOProControlesTemp.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
                        oDOProControlesTemp.ControlOtroEESS = oRsCatalogoControles.Fields!ControlOtroEESS
                        oDOProControlesTemp.IdEstablecimiento = oRsCatalogoControles.Fields!IdEstablecimiento
                        oDOProControlesTemp.FechaControl = oRsCatalogoControles.Fields!FechaControl
                        If Not oProControles.Insertar(oDOProControlesTemp) Then
                            ms_MensajeError = oProControles.MensajeError: GoTo ErrCeAgr
                        End If
                    End If
                End If
           End If
            oRsCatalogoControles.MoveNext
        Loop
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Inserta los datos configurados del nuevo control
    If oRsDatosProControles.RecordCount > 0 Then
        oRsDatosProControles.MoveFirst
        Do While Not oRsDatosProControles.EOF
           oDoProControlDato.IdPrograma = oDoProCabecera.IdPrograma
           oDoProControlDato.IdProCabecera = oDoProCabecera.IdProCabecera
           oDoProControlDato.IdControl = oDOProControles.IdControl
           oDoProControlDato.IdControlDato = oRsDatosProControles.Fields!IdControlDato
           oDoProControlDato.ControlDato = oRsDatosProControles.Fields!ControlDato
           oDoProControlDato.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
           If Not oProControlDato.Insertar(oDoProControlDato) Then
                ms_MensajeError = oProControlDato.MensajeError: GoTo ErrCeAgr
           End If
           oRsDatosProControles.MoveNext
        Loop
    End If
    
    If lbActualizarControl = True Then
        'Eliminar Diagnosticos
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ProDiagnosticosEliminarTodos"
            Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, oDoProCabecera.IdPrograma): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, oDOProControles.IdProCabecera): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, oDOProControles.IdControl): .Parameters.Append oParameter
            .Execute
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
    End If
''Eliminar Dx de la atención
'    If Not oAtencionesDiagnosticos.EliminarDiagnosticosPorAtencion(lnIdAtencion) Then
'       ms_MensajeError = oAtencionesDiagnosticos.MensajeError: GoTo ErrCeAgr
'    End If

'Inserta Diagnosticos
    If oRsProDiagnosticos.RecordCount > 0 Then
        oRsProDiagnosticos.MoveFirst
        Do While Not oRsProDiagnosticos.EOF
           oDOProDiagnosticos.IdPrograma = oDoProCabecera.IdPrograma
           oDOProDiagnosticos.IdProCabecera = oDoProCabecera.IdProCabecera
           oDOProDiagnosticos.IdControl = oDOProControles.IdControl
           oDOProDiagnosticos.IdDiagnostico = oRsProDiagnosticos.Fields!IdDiagnostico
           oDOProDiagnosticos.Principal = oRsProDiagnosticos.Fields!Principal
           oDOProDiagnosticos.IdSubclasificacionDx = oRsProDiagnosticos.Fields!IdSubclasificacionDx
           oDOProDiagnosticos.labConfHIS = oRsProDiagnosticos.Fields!labConfHIS
           oDOProDiagnosticos.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
           If Not oProDiagnosticos.Insertar(oDOProDiagnosticos) Then
             ms_MensajeError = oProDiagnosticos.MensajeError: GoTo ErrCeAgr
           End If
           'Insertar Dx en atenciones
           oDOAtencionDiagnostico.idAtencion = lnIdAtencion
           oDOAtencionDiagnostico.IdDiagnostico = oRsProDiagnosticos.Fields!IdDiagnostico
           
           'Actualizado 15102014
           If Not oAtencionesDiagnosticos.SeleccionarPorIdAtencionIdDx(oDOAtencionDiagnostico) Then
                oDOAtencionDiagnostico.IdClasificacionDx = 1
                oDOAtencionDiagnostico.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
                oDOAtencionDiagnostico.labConfHIS = Space(3)
                oDOAtencionDiagnostico.IdSubclasificacionDx = oRsProDiagnosticos.Fields!IdSubclasificacionDx
'                If oRsProDiagnosticos.Fields!Principal Then
'                     oDOAtencionDiagnostico.IdSubClasificacionDX = 102 'Definitivo
'                Else
'                     oDOAtencionDiagnostico.IdSubClasificacionDX = 101 'Presuntivo
'                End If
                
                If Not oAtencionesDiagnosticos.Insertar(oDOAtencionDiagnostico) Then
                     ms_MensajeError = oAtencionesDiagnosticos.MensajeError: GoTo ErrCeAgr
                End If
           End If
           
           oRsProDiagnosticos.MoveNext
        Loop
    End If
    
    If lbActualizarControl = True Then
        'Eliminar Procedimientos
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ProProcedimientosEliminarTodos"
            Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, oDoProCabecera.IdPrograma): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, oDOProControles.IdProCabecera): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, oDOProControles.IdControl): .Parameters.Append oParameter
            .Execute
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
    End If
    'Inserta Procedimientos
    If oRsProProcedimientos.RecordCount > 0 Then
        oRsProProcedimientos.MoveFirst
        Do While Not oRsProProcedimientos.EOF
           oDOProProcedimientos.IdPrograma = oDoProCabecera.IdPrograma
           oDOProProcedimientos.IdProCabecera = oDoProCabecera.IdProCabecera
           oDOProProcedimientos.IdControl = oRsProProcedimientos.Fields!IdControl
           oDOProProcedimientos.IdDiagnostico = 1
           oDOProProcedimientos.idProducto = oRsProProcedimientos.Fields!idProducto
           oDOProProcedimientos.IdResultado = oRsProProcedimientos.Fields!IdResultado
           oDOProProcedimientos.labConfHIS = oRsProProcedimientos.Fields!labConfHIS
           oDOProProcedimientos.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
           If Not oProProcedimientos.Insertar(oDOProProcedimientos) Then
             ms_MensajeError = oProProcedimientos.MensajeError: GoTo ErrCeAgr
           End If
           oRsProProcedimientos.MoveNext
        Loop
    End If
    
    If lbActualizarControl = True Then
        'Eliminar Tratamientos
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ProTratamientosEliminarTodos"
            Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, oDoProCabecera.IdPrograma): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, oDOProControles.IdProCabecera): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, oDOProControles.IdControl): .Parameters.Append oParameter
            .Execute
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
    End If
    'Inserta Tratamientos
    If oRsProTratamientos.RecordCount > 0 Then
        oRsProTratamientos.MoveFirst
        Do While Not oRsProTratamientos.EOF
           oDOProTratamientos.IdPrograma = oDoProCabecera.IdPrograma
           oDOProTratamientos.IdProCabecera = oDoProCabecera.IdProCabecera
           oDOProTratamientos.IdControl = oDOProControles.IdControl
           oDOProTratamientos.idProducto = oRsProTratamientos.Fields!idProducto
           oDOProTratamientos.IdUsuarioAuditoria = oDoProCabecera.IdUsuarioAuditoria
           If Not oProTratamientos.Insertar(oDOProTratamientos) Then
                ms_MensajeError = oProTratamientos.MensajeError: GoTo ErrCeAgr
           End If
           oRsProTratamientos.MoveNext
        Loop
    End If

    ProgramaControlesModificar = True
    oConexion.Close
ErrCeAgr:
'Resume
    If ProgramaControlesModificar = True Then
       oConexion.CommitTrans
    Else
       oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oProCabeceraDato = Nothing
    Set oDOProCabeceraDato = Nothing
    Set oProControles = Nothing
    Set oProControlDato = Nothing
    Set oDoProControlDato = Nothing
    Set oProDiagnosticos = Nothing
    Set oDOProDiagnosticos = Nothing
    Set oAtencionesDiagnosticos = Nothing
    Set oDOAtencionDiagnostico = Nothing
    Set oProProcedimientos = Nothing
    Set oDOProProcedimientos = Nothing
    Set oProTratamientos = Nothing
    Set oDOProTratamientos = Nothing
End Function

Function ProgramaControlesEliminar(oDOProControles As DOProControles, lnIdAtencion As Long, lbActualizarControl As Boolean) As Boolean
    On Error GoTo ErrCeAgr
    Dim oConexion As New ADODB.Connection
    Dim oRsTmp1 As New ADODB.Recordset

    Dim oProControles As New ProControles
    ProgramaControlesEliminar = False
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    '
    oConexion.BeginTrans
    '
    Set oProControles.Conexion = oConexion

    If lbActualizarControl = True Then
        If Not oProControles.Eliminar(oDOProControles) Then
            ms_MensajeError = oProControles.MensajeError: GoTo ErrCeAgr
        End If
    End If
    
    ProgramaControlesEliminar = True
    oConexion.Close
ErrCeAgr:
'Resume
    If ProgramaControlesEliminar = True Then
       oConexion.CommitTrans
    Else
       oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oProControles = Nothing
End Function


Function ProSeleccionarControlesAnteriores(lnIdPrograma As Long, lnIdProCabecera As Long, lnIdControlSeleccionado As Long, oConexion As Connection) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set ProSeleccionarControlesAnteriores = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProSeleccionarControlesAnteriores"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, lnIdProCabecera): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdControlSeleccionado", adInteger, adParamInput, 0, lnIdControlSeleccionado): .Parameters.Append oParameter
        Set oRsTmp = .Execute
    End With
    
    Set ProSeleccionarControlesAnteriores = oRsTmp
    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Function ProControlConfigGrafico(lnIdPrograma As Long, lbEsGraficoX As Boolean, oConexion As Connection) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set ProControlConfigGrafico = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProControlConfigGrafico"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EsGraficoX", adBoolean, adParamInput, 0, lbEsGraficoX): .Parameters.Append oParameter
        Set oRsTmp = .Execute
    End With
    
    Set ProControlConfigGrafico = oRsTmp
    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Function ProHistoricoDatoControlGrafico(lnIdPrograma As Long, lnIdProCabecera As Long, lnIdControl As Long, lnIdControlDatoX As Long, lnIdControlDatoY As Long, oConexion As Connection) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set ProHistoricoDatoControlGrafico = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProHistoricoDatoControlGrafico"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, lnIdProCabecera): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Idcontrol", adInteger, adParamInput, 0, lnIdControl): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdControlDatoX", adInteger, adParamInput, 0, lnIdControlDatoX): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdControlDatoY", adInteger, adParamInput, 0, lnIdControlDatoY): .Parameters.Append oParameter
        Set oRsTmp = .Execute
    End With
    Set ProHistoricoDatoControlGrafico = oRsTmp
    Set oRsTmp = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>



Function SeleccionarGrupoParametros(lcParametro As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Parametros
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
   
    
    Set oTabla.Conexion = oConexion
    Set SeleccionarGrupoParametros = oTabla.SeleccionarPorGrupo(lcParametro)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function LlenadoParametros() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection

Set LlenadoParametros = Nothing
        
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open SIGHEntidades.CadenaConexion
        

        With oCommand
           .CommandType = adCmdStoredProc
           Set .ActiveConnection = oConexion
           .CommandTimeout = 150
           .CommandText = "ParametrosGrupos"
           Set oRecordset = .Execute
           Set oRecordset.ActiveConnection = Nothing
        End With

        Set LlenadoParametros = oRecordset
        ms_MensajeError = ""
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
        Exit Function
ManejadorDeError:
        ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
        Exit Function
End Function


Function PerinatalAtencionSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        With oCommand
           .CommandType = adCmdStoredProc
           Set .ActiveConnection = oConexion
           .CommandTimeout = 150
           .CommandText = "PerinatalAtencionSeleccionarTodos"
           Set oRecordset = .Execute
           Set oRecordset.ActiveConnection = Nothing
        End With
        Set PerinatalAtencionSeleccionarTodos = oRecordset
        ms_MensajeError = ""
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
        Exit Function
ManejadorDeError:
        ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
        Exit Function
End Function

Function ProgramaControlMaternoConsultarTotalControles(TotalControl As Integer, oConexion As Connection) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set ProgramaControlMaternoConsultarTotalControles = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProgramaControlMaternoConsultarTotalControles"
        Set oParameter = .CreateParameter("@Controles", adInteger, adParamInput, 0, TotalControl): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set ProgramaControlMaternoConsultarTotalControles = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function ProgramaControlReporteFPP(lcFechaInicio As String, lcFechaFin As String, oConexion As Connection) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set ProgramaControlReporteFPP = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProgramaControlReporteFPP"
        Set oParameter = .CreateParameter("@FechaInicio", adVarChar, adParamInput, 10, lcFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adVarChar, adParamInput, 10, lcFechaFin): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set ProgramaControlReporteFPP = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function his_historicoAtencionesSeleccionarPorId(lnIdPaciente As Long, oConexion As Connection) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "his_historicoAtencionesSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set his_historicoAtencionesSeleccionarPorId = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function FactPuntosCargaSeleccionarIdServicio(lnIdServicio As Long, oConexion As Connection) As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactPuntosCargaSeleccionarIdServicio"
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set FactPuntosCargaSeleccionarIdServicio = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function


Function RecetaDosisSelecionarTodos() As Recordset
    Dim lcSql As String
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaDosisSelecionarTodos"
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set RecetaDosisSelecionarTodos = oRsTmp
   oConexion.Close
   Set oConexion = Nothing
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function EstablecimientosDevuelveUltimoID(oConexion As Connection) As Long
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EstablecimientosDevuelveUltimoID"
        Set oRs = .Execute
   End With
   EstablecimientosDevuelveUltimoID = oRs!IdEstablecimiento
   Set oRs = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
End Function
Function FarmAlmacenesTipoSuministroSeleccionarTodos() As Recordset
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Set FarmAlmacenesTipoSuministroSeleccionarTodos = Nothing
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FarmAlmacenesTipoSuministroSeleccionarTodos"
            Set oRsTmp = .Execute
            Set oRsTmp.ActiveConnection = Nothing
       End With
       
       Set FarmAlmacenesTipoSuministroSeleccionarTodos = oRsTmp
       oConexion.Close
       Set oRsTmp = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
End Function

Sub CabeceraReportes(oWorkSheet As Object, lbEsOpenOffice As Boolean, Optional lbNoMuestraFechaHoraImpresion As Boolean, _
                     Optional lcRutaImagen As String)
    On Error GoTo miError

    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim NOMBREEESS As String
    Dim DIRECCIONEESS As String
    Dim TELEFONO As String
        
    NOMBREEESS = lcBuscaParametro.SeleccionaFilaParametro(205)
    DIRECCIONEESS = lcBuscaParametro.SeleccionaFilaParametro(206)
    TELEFONO = lcBuscaParametro.SeleccionaFilaParametro(207)
    
    Dim oEstilosPagina As Object
    Dim oEstiloPredeterminado As Object
    Dim sEstilo As String
    Dim DefPage As Object
    Dim PageStyles As Object
    Dim Sheet As Object
    Dim StyleFamilies As Object
    Dim Htext As Object
    Dim Hcontent As Object
    Dim ret As Long
    Dim Document As Object

    If lbEsOpenOffice = True Then
        'Encabezado de Pagina
        Set StyleFamilies = oWorkSheet.StyleFamilies
        Set PageStyles = StyleFamilies.getByName("PageStyles")
        Set DefPage = PageStyles.getByName("Default")
        With DefPage
            'Propiedades para el encabezado
            .HeaderIsOn = True                'Activar encabezado
            .HeaderIsDynamicHeight = True     'Ajuste dinamico de altura
            .HeaderBodyDistance = 20          'Distancia al cuerpo del documento
        End With
        Set Hcontent = DefPage.RightPageHeaderContent
        Set Htext = Hcontent.centerText
        lcSql = NOMBREEESS & Chr(13) & DIRECCIONEESS & Chr(13) & "TELEFONO:  " & TELEFONO
        Htext.String = lcSql
        DefPage.RightPageHeaderContent = Hcontent
        grabaTablaAuditoria ("(excel) " & oWorkSheet.Name)   'debb-27/05/2015
    Else
        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
        oWorkSheet.PageSetup.CenterHeader = "&""Arial Narrow,Normal""&8" & "NOMBREEESS" & Chr(13) & "DIRECCIONEESS" & Chr(13) & "TELEFONO:  "
        
        With oWorkSheet.PageSetup
           If lcRutaImagen = "" Then
              .LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
           Else
              .LeftHeaderPicture.FileName = lcRutaImagen
              
           End If
           .CenterHeader = "&""Arial Narrow,Normal""&8" & NOMBREEESS & Chr(13) & DIRECCIONEESS & Chr(13) & "TELEFONO:  " & TELEFONO
           If lbNoMuestraFechaHoraImpresion = False Then
              .RightHeader = "&""Arial Narrow,Normal""&8" & "&D" & Chr(13) & "&T" & vbCrLf & "Pagina " & "&P" & " de " & "&N"
           End If
        End With
        grabaTablaAuditoria ("(excel) " & oWorkSheet.Name)   'debb-27/05/2015
    End If
    'Actualizado 16092014
miError:
    Select Case Err.Number
    Case 1004, 0
        'MsgBox "No se pudo asignar la cabecera al reporte", vbExclamation, "Reporte"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
End Sub

'yamill palomino
Function EmpleadosSeleccionarPorDNI(lcDNI As String) As Recordset
    Dim lcSql As String
    Dim oRsTmp5 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Set EmpleadosSeleccionarPorDNI = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosSeleccionarPorDNI"
        Set oParameter = .CreateParameter("@DNI", adVarChar, adParamInput, 8, Trim(lcDNI)): .Parameters.Append oParameter
        Set oRsTmp5 = .Execute
   End With
   Set EmpleadosSeleccionarPorDNI = oRsTmp5
   Set oRsTmp5 = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Public Function cargarDatosTriajeAObjetoDatos(ByRef oDOAtencionesCE As DOAtencionesCE, _
            ByVal oDOAtencionesCETriaje As DOAtencionesCE) As DOAtencionesCE
    If oDOAtencionesCETriaje Is Nothing Then
        Set oDOAtencionesCETriaje = New DOAtencionesCE
    End If
    With oDOAtencionesCE
        .TriajeFrecCardiaca = oDOAtencionesCETriaje.TriajeFrecCardiaca
        .TriajeFrecRespiratoria = Val(oDOAtencionesCETriaje.TriajeFrecRespiratoria)
        .TriajeOrigen = oDOAtencionesCETriaje.TriajeOrigen
        .TriajePerimCefalico = oDOAtencionesCETriaje.TriajePerimCefalico
        .TriajePeso = oDOAtencionesCETriaje.TriajePeso
        .TriajePresion = oDOAtencionesCETriaje.TriajePresion
        .TriajePulso = Val(oDOAtencionesCETriaje.TriajePulso)
        .TriajeTalla = oDOAtencionesCETriaje.TriajeTalla
        .TriajeTemperatura = oDOAtencionesCETriaje.TriajeTemperatura
    End With
    Set cargarDatosTriajeAObjetoDatos = oDOAtencionesCE
End Function

'mgaray09
Public Function diagnosticoValidarEliminar(oDODiagnostico As doDiagnostico) As Boolean
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oDiagnostico As New Diagnosticos
    
    diagnosticoValidarEliminar = False
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    Set oDiagnostico.Conexion = oConexion
    
    ms_MensajeError = DiagnosticoParseMessageError(oDiagnostico.validarEliminar(oDODiagnostico))
    If ms_MensajeError = "" Then
        diagnosticoValidarEliminar = True
    End If
    oConexion.Close
miError:
   If Err Then
        MsgBox Err.Description, vbCritical, "Acceso a Datos"
   End If
    Set oConexion = Nothing
    Set oDiagnostico = Nothing
End Function

Public Function DiagnosticoParseMessageError(oRsError As ADODB.Recordset) As String
    Dim messageError As String
    messageError = ""
    If oRsError.RecordCount > 0 Then
        Select Case oRsError.Fields!ErrorNumber
            Case 1:
                messageError = "No se Puede Eliminar, Se han registrado este diagnostico en atenciones"
            Case 2:
                messageError = "No se Puede Eliminar, Se han configurado este diagnostico para módulo perinatal"
            Case 3:
                messageError = "No se Puede Eliminar, registro esta en uso"
            Case 4:
                messageError = "No se Puede Eliminar, registro esta en uso"
            Case 5:
                messageError = "No se Puede Eliminar, Se han registrado este diagnostico en His"
        End Select
    End If
    DiagnosticoParseMessageError = messageError
End Function



Sub RolesEliminaHospEmergDelADMINISTRADOR(oConexion As Connection)
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RolesEliminaHospEmergDelADMINISTRADOR"
         .Execute
   End With
   Set oCommand = Nothing
End Sub


'Actualizado 22092014
Function RecetaCabeceraPorIdReceta(lnIdReceta As Long) As Recordset
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaCabeceraPorIdReceta"
        Set oParameter = .CreateParameter("@lnidReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    Set RecetaCabeceraPorIdReceta = oRs
    Set oCommand = Nothing
    Set oConexion = Nothing
End Function

'Actualizado 22092014
Function RecetaDetallePorIdReceta(lnIdReceta As Long) As Recordset
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaDetallePorIdReceta"
        Set oParameter = .CreateParameter("@lnidReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    Set RecetaDetallePorIdReceta = oRs
    Set oCommand = Nothing
    Set oConexion = Nothing
End Function

'Actualizado 22092014
Function RecetaDetalleServicioPorIdReceta(lnIdReceta As Long) As Recordset
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaDetalleServicioPorIdReceta"
        Set oParameter = .CreateParameter("@lnidReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    Set RecetaDetalleServicioPorIdReceta = oRs
    Set oCommand = Nothing
    Set oConexion = Nothing
End Function

'Actualizado 26092014
Function RecetasListadoViasAdministracion() As Recordset
'    Dim lcSql As String
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    
    Set RecetasListadoViasAdministracion = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetasListadoViasAdministracion"
        Set oRs = .Execute
    End With
    Set RecetasListadoViasAdministracion = oRs
    Set oCommand = Nothing
    Set oConexion = Nothing
End Function

'mgaray20140926
Function RecetaDiagnosticosPorIdAtencion(idAtencion As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    
    ms_MensajeError = ""
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaDiagnosticosPorIdAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set RecetaDiagnosticosPorIdAtencion = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
    
ManejadorDeError:
    ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    Exit Function
End Function

'Actualizado 16102014
Function TiposEmpleadosSeleccionarId(lnIdTipoEmpleado As Long) As Integer
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposEmpleadosSeleccionarId"
        Set oParameter = .CreateParameter("@IdTipoEmpleado", adInteger, adParamInput, 0, lnIdTipoEmpleado): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
    
    If oRsTmp1.RecordCount > 0 Then
       TiposEmpleadosSeleccionarId = IIf(oRsTmp1.Fields!EsColegiatura = True, 1, 0)
    Else
       TiposEmpleadosSeleccionarId = 0
    End If
       oConexion.Close
       Set oRsTmp1 = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function

'mgaray20141012
Function EstablecimientosSeleccionarTodos() As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    
    Dim oEstablecimientos As New Establecimientos
    
    Set EstablecimientosSeleccionarTodos = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    
    Set oEstablecimientos.Conexion = oConexion
    Set oRsTmp1 = oEstablecimientos.SeleccionarTodos
    
   Set EstablecimientosSeleccionarTodos = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
End Function

'mgaray20141020
Public Function PaisSeleccionarPorId(idPais As Long) As DOPais
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oPais As New Paises
    Dim oDOPais As New DOPais
    
    Set PaisSeleccionarPorId = Nothing
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oPais.Conexion = oConexion
    oDOPais.idPais = idPais
    
    If oPais.SeleccionarPorId(oDOPais) Then
        Set PaisSeleccionarPorId = oDOPais
    End If
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oPais.MensajeError
    Set oPais = Nothing
miError:
    If Err Then
        ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    End If
End Function


Public Function DistritoSeleccionarPorId(IdDistrito As Long) As DODistrito
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oDistrito As New Distritos
    Dim oDODistrito As New DODistrito
    
    Set DistritoSeleccionarPorId = Nothing
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oDistrito.Conexion = oConexion
    oDODistrito.IdDistrito = IdDistrito
    
    If oDistrito.SeleccionarPorId(oDODistrito) Then
        Set DistritoSeleccionarPorId = oDODistrito
    End If
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oDistrito.MensajeError
    Set oDistrito = Nothing
miError:
    If Err Then
        ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    End If
End Function

'mgaray201410e
Public Function ProcedimientoEsParaReceta(ml_idProducto As Long) As Boolean
On Error GoTo miError
    Dim oDoCatalogoServicios As DOCatalogoServicio
    Dim returnValue As Boolean
    
    returnValue = False
    If CatalogoServiciosSeleccionarEsParaReceta(ml_idProducto) = True Then
        returnValue = True
    End If
'    Set oDoCatalogoServicios = CatalogoServiciosSeleccionarPorId(ml_idProducto)
'    If Not (oDoCatalogoServicios Is Nothing) Then
'        If oDoCatalogoServicios.esCpt = 0 Then
'            returnValue = True
'        End If
'    End If
    
    ProcedimientoEsParaReceta = returnValue
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical, "Verificar Procedmiento para receta"
    End If
End Function

Public Function ProcedimientoEsParaCpt(ml_idProducto As Long) As Boolean
On Error GoTo miError
    Dim oDoCatalogoServicios As DOCatalogoServicio
    Dim returnValue As Boolean

    returnValue = True
    If CatalogoServiciosSeleccionarEsParaReceta(ml_idProducto) = True Then
        returnValue = False
    End If
    If returnValue = False Then
        'mgaray201411e
        Dim oRsExcepcionesRecetas As ADODB.Recordset
        Set oRsExcepcionesRecetas = CptExcepcionesRecetasHISSeleccionPorId(ml_idProducto)
        If Not (oRsExcepcionesRecetas Is Nothing) Then
            If oRsExcepcionesRecetas.RecordCount > 0 Then
                returnValue = True
            End If
        End If
    End If

    ProcedimientoEsParaCpt = returnValue
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical, "Verificar Procedmiento para receta"
    End If
End Function

'Actualizado 30102014
Function ProCatalogoCptSeleccionarPorIdPrograma(lnIdPrograma As Long, oConexion As Connection) As Recordset
     Dim lcSql As String
     Dim oRsTmp As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set ProCatalogoCptSeleccionarPorIdPrograma = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProCatalogoCptSeleccionarPorIdPrograma"
        Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, lnIdPrograma): .Parameters.Append oParameter
        Set oRsTmp = .Execute
   End With
   Set ProCatalogoCptSeleccionarPorIdPrograma = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function CatalogoServiciosSeleccionarEsParaReceta(idProducto As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oProducto As New CatalogoServicios
Dim oDOBienesInsumos As New DOCatalogoServicio
    
    CatalogoServiciosSeleccionarEsParaReceta = False
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oProducto.Conexion = oConexion
    oDOBienesInsumos.idProducto = idProducto
    
    CatalogoServiciosSeleccionarEsParaReceta = oProducto.SeleccionarPorIdEsParaReceta(oDOBienesInsumos)
        
    oConexion.Close
    ms_MensajeError = oProducto.MensajeError
    Set oConexion = Nothing
    Set oProducto = Nothing

End Function

'mgaray201411a
Public Function existeCodigoLabHis(oRsLabHis As ADODB.Recordset, cLabHIS As String) As Boolean
    On Error Resume Next
    Dim returnValue As Boolean
    
    returnValue = False
    
'    If Trim(cLabHIS) <> "" Then
        If oRsLabHis.RecordCount > 0 Then
            oRsLabHis.MoveFirst
            oRsLabHis.Find "valores='" & cLabHIS & "'"
            If oRsLabHis.EOF = False Then
                returnValue = True
            End If
        End If
'    End If
    existeCodigoLabHis = returnValue
    Err = 0
End Function

Function CptExcepcionesRecetasHISSeleccionPorId(lnIdProducto As Long) As Recordset
On Error GoTo miError
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CptExcepcionesRecetasHISSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    Set CptExcepcionesRecetasHISSeleccionPorId = oRs
    Set oCommand = Nothing
    Set oConexion = Nothing
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical, "Verificar Procedmiento para receta"
    End If
End Function

'mgaray201411f
Function TiposDocIdentidadSeleccionarTodosIncSinTipoDoc() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposDocIdentidad
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposDocIdentidadSeleccionarTodosIncSinTipoDoc = oTabla.SeleccionarTodosIncSinTipoDoc
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oConexion = Nothing

End Function
'''Mostrar todos los paises - GLCC
''Function NacionalidadSeleccionarTodos() As ADODB.Recordset
''Dim oConexion As New ADODB.Connection
''Dim oTabla As New Nacionalidad
''
''    oConexion.Open sighentidades.CadenaConexion
''    oConexion.CursorLocation = adUseClient
''
''    Set oTabla.Conexion = oConexion
''    Set NacionalidadSeleccionarTodos = oTabla.NacionalidadSeleccionarTodos
''    oConexion.Close
''    ms_MensajeError = oTabla.MensajeError
''    Set oTabla = Nothing
''    Set oConexion = Nothing
''End Function
'debb2014b
Function CatalogoBienesInsumosSeleccionarXid(lnIdProducto As Long, Optional oConexion1 As Connection) As Recordset
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
   If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
   End If
   With oCommand
     .CommandType = adCmdStoredProc
     If oConexion1 Is Nothing Then
         Set .ActiveConnection = oConexion
     Else
         Set .ActiveConnection = oConexion1
     End If
     .CommandText = "FactCatalogoBienesInsumosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
   Set CatalogoBienesInsumosSeleccionarXid = oRecordset
   Set oCommand = Nothing
   Set oConexion = Nothing
End Function

Function AtenHospCensoSeleccionarPorId(lnIdRangoCensoHosp As Long) As DoAtencionHospCenso
Dim oConexion As New ADODB.Connection
Dim oAtenHospCenso As New AtenHospCenso
Dim oDoAtencionHospCenso As New DoAtencionHospCenso
    
    Set AtenHospCensoSeleccionarPorId = Nothing
    
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oAtenHospCenso.Conexion = oConexion
    oDoAtencionHospCenso.IdRangoCensoHosp = lnIdRangoCensoHosp
    
    If oAtenHospCenso.SeleccionarPorId(oDoAtencionHospCenso) Then
        Set AtenHospCensoSeleccionarPorId = oDoAtencionHospCenso
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oAtenHospCenso.MensajeError
    Set oAtenHospCenso = Nothing
End Function

Function AtenHospCensoSeleccionarModificar(oDoAtencionHospCenso1 As DoAtencionHospCenso, _
                                            oDoAtencionHospCenso2 As DoAtencionHospCenso, _
                                                oDoAtencionHospCenso3 As DoAtencionHospCenso) As Boolean
    On Error GoTo ErrCeAgr
    Dim oConexion As New ADODB.Connection
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
  
    Dim oAtenHospCenso As New AtenHospCenso
    
    AtenHospCensoSeleccionarModificar = False
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    '
    oConexion.BeginTrans
    '
    Set oAtenHospCenso.Conexion = oConexion
        
    If Not oAtenHospCenso.modificar(oDoAtencionHospCenso1) Then
        ms_MensajeError = oAtenHospCenso.MensajeError: GoTo ErrCeAgr
    End If
    
    If Not oAtenHospCenso.modificar(oDoAtencionHospCenso2) Then
        ms_MensajeError = oAtenHospCenso.MensajeError: GoTo ErrCeAgr
    End If
    
    If Not oAtenHospCenso.modificar(oDoAtencionHospCenso3) Then
        ms_MensajeError = oAtenHospCenso.MensajeError: GoTo ErrCeAgr
    End If
    
    AtenHospCensoSeleccionarModificar = True
    oConexion.Close
    
ErrCeAgr:
'Resume
    If AtenHospCensoSeleccionarModificar = True Then
       oConexion.CommitTrans
    Else
       oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oAtenHospCenso = Nothing
    Set oDoAtencionHospCenso1 = Nothing
    Set oDoAtencionHospCenso2 = Nothing
    Set oDoAtencionHospCenso3 = Nothing
End Function


'debb-2/3/2015
Sub ActualizaTablaPacientesMovimientos(oConexion As Connection, ldFecha As Date, lnIdCuentaAtencion As Long, _
                                       lnIdListBarItem As sghOpcionGalenHos, lbEsEliminarRegistro As Boolean, _
                                       Optional lnPeso As Double, _
                                       Optional lnTalla As Double, Optional lnIdDxNutricional As Long, _
                                       Optional lnGrafXedadEnMeses As Long, Optional lnGrafYpercentilTE As Long, _
                                       Optional lnGrafYpercentilPT As Long, Optional lnGrafYpercentilPE As Long, _
                                       Optional lnZetaPT As Double, Optional lnZetaTE As Double, _
                                       Optional lnZetaPE As Double, Optional lnHemoglobina As Double, _
                                       Optional lnParasitosis As Integer)
                                       
            Dim oDOPacienteMovimientos As New DOPacienteMovimientos
            Dim oPacientesMovimientos As New PacientesMovimientos
            Dim lbEncontro As Boolean
            Set oPacientesMovimientos.Conexion = oConexion
            
            oDOPacienteMovimientos.idCuentaAtencion = lnIdCuentaAtencion
            lbEncontro = oPacientesMovimientos.SeleccionarPorId(oDOPacienteMovimientos)
            If lnPeso > 0 Then
               oDOPacienteMovimientos.Peso = lnPeso
            End If
            If lnTalla > 0 Then
               oDOPacienteMovimientos.Talla = lnTalla
            End If
            If lnPeso > 0 Then
               oDOPacienteMovimientos.idDxNutricional = lnIdDxNutricional
            End If
            If lnGrafXedadEnMeses > 0 Then
                oDOPacienteMovimientos.GrafXedadEnMeses = lnGrafXedadEnMeses
            End If
            If lnPeso > 0 Then
               oDOPacienteMovimientos.GrafYpercentilTE = lnGrafYpercentilTE
            End If
            If lnPeso > 0 Then
                oDOPacienteMovimientos.GrafYpercentilPT = lnGrafYpercentilPT
            End If
            If lnPeso > 0 Then
                oDOPacienteMovimientos.GrafYpercentilPE = lnGrafYpercentilPE
            End If
            If lnPeso > 0 Then
                oDOPacienteMovimientos.ZetaPT = lnZetaPT
            End If
            If lnPeso > 0 Then
                oDOPacienteMovimientos.ZetaTE = lnZetaTE
            End If
            If lnPeso > 0 Then
                oDOPacienteMovimientos.zetaPE = lnZetaPE
            End If
            If lnHemoglobina > 0 And oDOPacienteMovimientos.Peso > 0 Then
                oDOPacienteMovimientos.Hemoglobina = lnHemoglobina
            End If
            If lnParasitosis > 0 And oDOPacienteMovimientos.Peso > 0 Then
               oDOPacienteMovimientos.Parasitosis = IIf(lnParasitosis = 1, "Si", "No") '1=verdadero, 2=falso
            ElseIf lnParasitosis = 3 And oDOPacienteMovimientos.Peso > 0 Then                                            '3=se quitó elección V/F
               oDOPacienteMovimientos.Parasitosis = ""
            End If
            If lbEncontro = False Then
                If oPacientesMovimientos.Insertar(oDOPacienteMovimientos) = True Then
                End If
            Else
                If oPacientesMovimientos.modificar(oDOPacienteMovimientos) = True Then
                End If
            End If
'            If oPacientesMovimientos.PacientesMovimientosActualizar(oDOPacienteMovimientos, lbEsEliminarRegistro) = True Then
'            End If
            Set oDOPacienteMovimientos = Nothing
            Set oPacientesMovimientos = Nothing
End Sub

'debb-27/05/2015
Sub grabaTablaAuditoria(lcNombreReporte As String)
    On Error GoTo errGTA
    Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    Call mo_ReglasSeguridad.AuditoriaAgregarV(SIGHEntidades.Usuario, "C", 999, "Reportes", oConexion, _
                             999999, SIGHEntidades.RetornaNombrePC, Left(lcNombreReporte, 100))    'ListBarItems.idListItem
    oConexion.Close
    Set oConexion = Nothing
errGTA:
End Sub

'debb-09/07/2015
Function DevuelvePrecioItem(lnIdProducto As Long, lnIdPuntoCarga As Long, ml_IdTipoFinanciamiento As Long, Optional oConexion1 As Connection) As Double
      Dim oRsTmp As New Recordset
      Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
      If lnIdPuntoCarga = sghPtoCargaFarmacia Then
         Set oRsTmp = mo_reglasComunes.FactCatalogoBienesInsumosHospXfiltro("idProducto=" & lnIdProducto & " and idTipoFinanciamiento=" & ml_IdTipoFinanciamiento)
      Else
         Set oRsTmp = mo_reglasComunes.FactCatalogoServiciosHospXfiltro("idProducto=" & lnIdProducto & " and idTipoFinanciamiento=" & ml_IdTipoFinanciamiento)
      End If
      DevuelvePrecioItem = 0
      If oRsTmp.RecordCount > 0 Then
         DevuelvePrecioItem = oRsTmp.Fields!PrecioUnitario
      End If
      Set mo_reglasComunes = Nothing
End Function


'debb-19/08/2015
Function RecetaDetalleUltimaViaAdministracion(lnIdProductoFarmacia As Long) As Long
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaDetalleXultimaViaAdministracion"
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProductoFarmacia): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    RecetaDetalleUltimaViaAdministracion = 0
    If oRs.RecordCount > 0 Then
       If Not IsNull(oRs!IdViaAdministracion) Then
          RecetaDetalleUltimaViaAdministracion = oRs!IdViaAdministracion
       End If
    End If
    oConexion.Close
    Set oCommand = Nothing
    Set oConexion = Nothing
End Function

'***debb-05/08/2015
Function SuSalud_upsSeleccionarTodos() As Recordset
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
  
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SuSalud_upsSeleccionarTodos"
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    
    Set SuSalud_upsSeleccionarTodos = oRs
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'***debb-05/08/2015
Function SisFuaUPServiciosSeleccionarTodos() As Recordset
 On Error GoTo errsfup
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim lnError As Integer
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    lnError = 0
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SisFuaUPServiciosSeleccionarTodos"
        Set oRs = oCommand.Execute
        If lnError > 0 Then
           If oRs.State = 1 Then oRs.Close
           oRs.Open "select descripcion,ups from SisFuaUPServicios order by descripcion", oConexion, adOpenKeyset, adLockOptimistic
        End If
        Set oRs.ActiveConnection = Nothing
    End With
    Set SisFuaUPServiciosSeleccionarTodos = oRs
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set lcBuscaParametro = Nothing
    Exit Function
errsfup:
    If Err.Number = -2147217900 Then
       lnError = 1
       Resume Next
    End If
End Function
'***debb-29/04/2016
Function novafis_edadesSeleccionarTodos() As Recordset
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim lcBuscaParametro As New SIGHDatos.Parametros

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
  
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "novafis_edadesSeleccionarTodos"
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    
    Set novafis_edadesSeleccionarTodos = oRs
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set lcBuscaParametro = Nothing
End Function
'debb-29/04/2016
Function novafis_ElDxEsDeEmbarazada(lcDx As String, lnTipoSexo As Long) As Boolean
    If lnTipoSexo = 2 And (Left(UCase(lcDx), 3) = "Z35" Or Left(UCase(lcDx), 1) = "O") Then
       novafis_ElDxEsDeEmbarazada = True
    Else
       novafis_ElDxEsDeEmbarazada = False
    End If
End Function

'debb-02/05/2016
Function FUAvalidaCodigoPrestacionSegunAdmision(mo_lnIdTablaLISTBARITEMS As sghOpcionGalenHos, lcCodigoPrestacion As String) As Boolean
     FUAvalidaCodigoPrestacionSegunAdmision = True
     Select Case mo_lnIdTablaLISTBARITEMS
     Case sghOpcionGalenHos.sghAdmisionHospitalizacion
          If Val(lcCodigoPrestacion) = 62 Or Val(lcCodigoPrestacion) = 63 Or _
             Val(lcCodigoPrestacion) = 457 Or Val(lcCodigoPrestacion) = 458 Or _
             Val(lcCodigoPrestacion) = 56 Or Val(lcCodigoPrestacion) = 451 Or _
             Val(lcCodigoPrestacion) = 906 Then
             MsgBox "El CODIGO DE PRESTACION no corresponde a HOSPITALIZACION", vbInformation, ""
             FUAvalidaCodigoPrestacionSegunAdmision = False
             Exit Function
          End If
     Case sghOpcionGalenHos.sghAdmisionEmergencia
          If Val(lcCodigoPrestacion) = 56 Or Val(lcCodigoPrestacion) = 451 Or _
             Val(lcCodigoPrestacion) = 906 Then
             MsgBox "El CODIGO DE PRESTACION no corresponde a EMERGENCIA", vbInformation, ""
             FUAvalidaCodigoPrestacionSegunAdmision = False
             Exit Function
          End If
     Case sghOpcionGalenHos.sghRegistroAtencionCE Or sghOpcionGalenHos.sghRegistroCitaCE
          If Val(lcCodigoPrestacion) = 62 Or Val(lcCodigoPrestacion) = 63 Or _
             Val(lcCodigoPrestacion) = 457 Or Val(lcCodigoPrestacion) = 458 Then
             MsgBox "El CODIGO DE PRESTACION no corresponde a CONSULTORIO EXTERNO", vbInformation, ""
             FUAvalidaCodigoPrestacionSegunAdmision = False
             Exit Function
          End If
     End Select

End Function

'debb-18/05/2016
Sub RecetaChequeaEstadoActual(lnCuenta As Long, lnIdEstado As sghRecetaEstados, _
                              lnIdComprobantePago As Long, lcDocumentoDespacho As String)
    Dim oRsTmp1 As New Recordset
    Dim mo_AdminCaja As New ReglasCaja
    Select Case lnIdEstado
    Case sghRecetaEstados.sighRecetaAnulada
         MsgBox "Esa Receta está ANULADA" & Chr(13) & "N° Cuenta: " & Trim(Str(lnCuenta)), vbInformation, ""
    Case sghRecetaEstados.sighRecetaDespachada
         If IsNull(lnIdComprobantePago) Then
            lcSql = "   N° Movimiento: " & Trim(lcDocumentoDespacho)
            MsgBox "Esa Receta  fué ATENDIDA" & Chr(13) & "N° Cuenta: " & Trim(Str(lnCuenta)) & lcSql, vbInformation, ""
         ElseIf lnIdComprobantePago = 0 Then
            lcSql = "   N° Movimiento: " & Trim(lcDocumentoDespacho)
            MsgBox "Esa Receta  fué ATENDIDA" & Chr(13) & "N° Cuenta: " & Trim(Str(lnCuenta)) & lcSql, vbInformation, ""
         Else
            Set oRsTmp1 = mo_AdminCaja.CajaComprobantesSeleccionarPorId(lnIdComprobantePago)
            lcSql = ""
            If oRsTmp1.RecordCount > 0 Then
               lcSql = "  Boleta: " & oRsTmp1.Fields!NroSerie & "-" & oRsTmp1.Fields!NroDocumento
            End If
            MsgBox "Esa Receta tiene BOLETA y fué ATENDIDA" & Chr(13) & "N° Cuenta: " & Trim(Str(lnCuenta)) & lcSql, vbInformation, ""
         End If
    Case sghRecetaEstados.sighRecetaConBoleta
         lcSql = ""
         Set oRsTmp1 = mo_AdminCaja.CajaComprobantesSeleccionarPorId(lnIdComprobantePago)
         If oRsTmp1.RecordCount > 0 Then
            lcSql = "  Boleta: " & oRsTmp1.Fields!NroSerie & "-" & oRsTmp1.Fields!NroDocumento
         End If
         MsgBox "Esa Receta solo tiene BOLETA, aún no es ATENDIDA" & Chr(13) & "N° Cuenta: " & Trim(Str(lnCuenta)) & lcSql, vbInformation, ""
    End Select
    Set oRsTmp1 = Nothing
    Set mo_AdminCaja = Nothing
End Sub



'debb-18/05/2016
Function RecetaDetalleItemPorIdReceta(lnIdReceta As Long) As Recordset
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaDetalleItemSeleccionarPorId"
        Set oParameter = .CreateParameter("@lnidReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    Set RecetaDetalleItemPorIdReceta = oRs
    Set oCommand = Nothing
    Set oConexion = Nothing
End Function

'debb-18/05/2016
Sub WaitSeconds(nSecs As Integer)
        ' Esperar los segundos indicados
        ' Añadirle la diferencia a la hora actual
        Dim t1 As Date
        t1 = Format(DateAdd("s", nSecs, Now), SIGHEntidades.DevuelveHoraSoloFormato_HMS)
        
        ' Esta asignación solo es necesaria
        ' si la comprobación se hace al principio del bucle
        Dim t2 As Date
        t2 = Format(Now, SIGHEntidades.DevuelveHoraSoloFormato_HMS)

        ' Mientras no haya pasado el tiempo indicado
        Do While t2 < t1
            ' Un respiro para el sitema
            'System.Windows.Forms.Application.DoEvents()
            DoEvents
            ' Asignar la hora actual
            t2 = Format(Now, SIGHEntidades.DevuelveHoraSoloFormato_HMS)
        Loop
End Sub
'debb-21/06/2016
Sub DiagnosticosSeleccionarPorIdDevuelveDescripcion(IdDiagnostico As Long, oConexion As Connection, _
                                                         ByRef lcCodigo As String, ByRef lcDescripcion As String)
Dim oDiagnostico As New Diagnosticos
Dim oDODiagnostico As New doDiagnostico
    lcCodigo = "": lcDescripcion = ""
    Set oDiagnostico.Conexion = oConexion
    oDODiagnostico.IdDiagnostico = IdDiagnostico
    If oDiagnostico.SeleccionarPorId(oDODiagnostico) Then
        lcCodigo = oDODiagnostico.CodigoCIE2004
        lcDescripcion = oDODiagnostico.Descripcion
    End If
    Set oDiagnostico = Nothing
End Sub

'debb-14/07/2016
Function AtencionesSeleccionarMedicoEgresoPorCuenta(lnIdCuentaAtencion As Long) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set AtencionesSeleccionarMedicoEgresoPorCuenta = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "AtencionesSeleccionarMedicoEgresoPorCuenta"
             Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
             Set oRsTmp1 = .Execute
             Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set AtencionesSeleccionarMedicoEgresoPorCuenta = oRsTmp1
        oConexion.Close
        Set oRsTmp1 = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
End Function

'debb-09/08/2016
Function NOpuedeModificarResponsable(mi_opcion As sghOpciones, ml_idUsuario As Long, _
                                     oRsResponsables As Recordset) As Boolean
    If mi_opcion = sghAgregar Or mi_opcion = sghModificar Then
        If oRsResponsables.RecordCount > 0 Then
           oRsResponsables.MoveFirst
           Do While Not oRsResponsables.EOF
              If oRsResponsables!idEmpleado = ml_idUsuario Then
                 NOpuedeModificarResponsable = True
                 If mi_opcion = sghAgregar Then
                 Else
                    MsgBox "No podrá modificar el RESPONSABLE porque Ud. es uno de ellos", vbInformation, "Mensaje"
                 End If
                 Exit Do
              End If
              oRsResponsables.MoveNext
           Loop
        End If
    End If
End Function
'debb-09/09/2016
Function DevuelveFUMenUltimaAtencion(lnIdPaciente As Long, ldFechaAtencion As Date, oConexion As Connection) As String
    Dim mo_ReglasLaboratorio As New ReglasLaboratorio
    Dim mo_ReglasImagenes As New ReglasImagenes
    Dim lbEncontro As Boolean
    Dim oRsTmp1432 As New Recordset
    Dim lnEg As Integer
    DevuelveFUMenUltimaAtencion = SIGHEntidades.FECHA_VACIA_DMY
    lbEncontro = False
    Set oRsTmp1432 = mo_ReglasLaboratorio.LabMovimientoLaboratorioSeleccionarPorIdPaciente(lnIdPaciente, oConexion)
    oRsTmp1432.Filter = "eo_fum<>null"
    If oRsTmp1432.RecordCount > 0 Then
       DevuelveFUMenUltimaAtencion = Format(oRsTmp1432!eo_fum, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
       lbEncontro = True
    End If
    oRsTmp1432.Close
    Set oRsTmp1432 = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorIdPaciente(lnIdPaciente, oConexion)
    oRsTmp1432.Filter = "eo_fum<>null"
    If oRsTmp1432.RecordCount > 0 Then
       If lbEncontro = True Then
          If DevuelveFUMenUltimaAtencion < Format(oRsTmp1432!eo_fum, SIGHEntidades.DevuelveFechaSoloFormato_DMY) Then
             DevuelveFUMenUltimaAtencion = Format(oRsTmp1432!eo_fum, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
          End If
       Else
          DevuelveFUMenUltimaAtencion = Format(oRsTmp1432!eo_fum, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
          lbEncontro = True
       End If
    End If
    If IsDate(DevuelveFUMenUltimaAtencion) Then
       lnEg = SIGHEntidades.DevuelveEdadGestacional(CDate(DevuelveFUMenUltimaAtencion), ldFechaAtencion)
       If lnEg = 0 Then
          DevuelveFUMenUltimaAtencion = SIGHEntidades.FECHA_VACIA_DMY
       End If
    End If
    oRsTmp1432.Close
    Set oRsTmp1432 = Nothing
    Set mo_ReglasLaboratorio = Nothing
    Set mo_ReglasImagenes = Nothing
End Function
'debb-14/09/2016
Function PerinatalAtencionCredSeleccionarControles(lnIdPaciente As Long) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "PerinatalAtencionCredSeleccionarControles"
             Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
             Set oRsTmp1 = .Execute
             Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set PerinatalAtencionCredSeleccionarControles = oRsTmp1
        oConexion.Close
        Set oRsTmp1 = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
End Function

'debb-20/09/2016
Function HospitalizadoConCtaEmergNOabierta(lnIdPaciente As Long, ldFechaAltaEmergencia As Date, _
                                           lnidDestinoAtencionCtaEmerg As Long) As Boolean
    Dim mo_ReglasAdmision As New ReglasAdmision
    Dim oRsTmp86 As New Recordset
    HospitalizadoConCtaEmergNOabierta = False
    If lnidDestinoAtencionCtaEmerg <> 23 And lnidDestinoAtencionCtaEmerg <> 24 Then  'referencia/contraReferencia
        If lnidDestinoAtencionCtaEmerg = 2 Then    'hospitalizacion
            Set oRsTmp86 = mo_ReglasAdmision.AtencionesSeleccionarPorIdPaciente(lnIdPaciente, sghTipoServicio.sghHospitalizacion)
            If oRsTmp86.RecordCount > 0 Then
               oRsTmp86.MoveLast
               If oRsTmp86!IdEstadoAtencion = 1 And _
                  CDate(Format(oRsTmp86!FechaIngreso & " " & oRsTmp86!HoraIngreso, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)) > ldFechaAltaEmergencia Then
                    HospitalizadoConCtaEmergNOabierta = True
               End If
            End If
            oRsTmp86.Close
        Else
            HospitalizadoConCtaEmergNOabierta = True
        End If
    End If
    
    Set oRsTmp86 = Nothing
    Set mo_ReglasAdmision = Nothing
End Function


'debb-30/09/2016
Function CalculaNUMEROREFERENCIA(lbEsContraReferencia As Boolean) As String
       Dim oRsTmp As New Recordset
       Dim oConexion As New Connection
       Dim mo_AdminAdmision As New ReglasAdmision
       Dim lcBuscaParametro As New SIGHDatos.Parametros
       Dim lcDestinos As String
       Dim lcHoyAnio As String
       lcHoyAnio = Right(lcBuscaParametro.RetornaFechaServidorSQL, 4)
       If lbEsContraReferencia = True Then
           lcDestinos = "13,24,32"                           'Contrareferencias en CE, Emer,Hosp
       Else
           lcDestinos = "12,23,31"                           'Referencias en CE, Emer,Hosp
       End If
       oConexion.CommandTimeout = 900
       oConexion.CursorLocation = adUseClient
       oConexion.Open SIGHEntidades.CadenaConexion
'       Set oRsTmp = mo_AdminAdmision.atencionesDatosAdicionalesXfiltro(" NOT (dbo.AtencionesDatosAdicionales.NroReferenciaDestino IS NULL) " & _
'                      " and dbo.atenciones.idEstadoAtencion<>0 and dbo.atenciones.idDestinoAtencion in (" & lcDestinos & ")" & _
'                      " order by cast(dbo.AtencionesDatosAdicionales.NroReferenciaDestino as int) desc", _
'                      oConexion)
'       If oRsTmp.RecordCount = 0 Then
'          CalculaNUMEROREFERENCIA = "1"
'       Else
'          CalculaNUMEROREFERENCIA = Trim(Str(Val(oRsTmp!NroReferenciaDestino) + 1))
'       End If
       
       Set oRsTmp = mo_AdminAdmision.atencionesDatosAdicionalesXfiltro(" NOT (dbo.AtencionesDatosAdicionales.NroReferenciaDestino IS NULL) " & _
                      " and dbo.atenciones.idEstadoAtencion<>0 and dbo.atenciones.idDestinoAtencion in (" & lcDestinos & ")" & _
                      " and year(referenciaDfextension)=" & lcHoyAnio, _
                      oConexion)
                      
       If oRsTmp.RecordCount = 0 Then
          CalculaNUMEROREFERENCIA = lcHoyAnio & "1"
       Else
          CalculaNUMEROREFERENCIA = lcHoyAnio & Trim(Str(oRsTmp.RecordCount + 1))
       End If
       oRsTmp.Close
       oConexion.Close
       Set oRsTmp = Nothing
       Set oConexion = Nothing
       Set mo_AdminAdmision = Nothing
       Set lcBuscaParametro = Nothing
End Function
'debb-30/09/2016
Function BuscarSiExisteNUMEROREFERENCIA(lcReferenciaDestino As String, lnIdAtencion As Long, lbEsContraReferencia As Boolean) As Boolean
        Dim oRsTmp1 As New Recordset
        Dim oRsTmp2 As New Recordset
        Dim oConexion As New Connection
        Dim mo_AdminAdmision As New ReglasAdmision
        Dim lcDestinos As String
        If lbEsContraReferencia = True Then
           lcDestinos = "13,24,32"                           'Contrareferencias en CE, Emer,Hosp
        Else
           lcDestinos = "12,23,31"                           'Referencias en CE, Emer,Hosp
        End If
        BuscarSiExisteNUMEROREFERENCIA = False
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open SIGHEntidades.CadenaConexion
        Set oRsTmp1 = mo_AdminAdmision.atencionesDatosAdicionalesXfiltro(" dbo.AtencionesDatosAdicionales.NroReferenciaDestino='" & _
                     lcReferenciaDestino & "' and dbo.atenciones.idEstadoAtencion<>0 and dbo.atenciones.idDestinoAtencion in (" & _
                     lcDestinos & ")", oConexion)
        oRsTmp1.Filter = "idAtencion<>" & lnIdAtencion
        If oRsTmp1.RecordCount > 0 Then
             ' Set oRsTmp2 = mo_AdminAdmision.AtencionesSeleccionarPorIdAtencion(oRsTmp1!idAtencion)
             ' If oRsTmp2!IdEstadoAtencion <> 0 Then
                    MsgBox "Ese N° " & IIf(lbEsContraReferencia = True, "ContraReferencia", "Referencia") & " YA EXISTE  (Destino)(ficha 3.1)", vbInformation, "REFERENCIA"
                    BuscarSiExisteNUMEROREFERENCIA = True
                    'CalculaNUMEROREFERENCIA
             ' End If
        End If
        oRsTmp1.Close
        oConexion.Close
        Set oRsTmp1 = Nothing
        Set oConexion = Nothing
        Set oRsTmp2 = Nothing
        Set mo_AdminAdmision = Nothing
End Function

Function DevuelveCuentaEmergenciaOceDelPacienteHospitalizado(lnIdPaciente As Long, ldFechaIngresoHosp As Date, _
                                                          ByRef lnIdAtencionEmergencia As Long, lnIdOrigenAtencion As Long) As Long
    Dim mo_ReglasAdmision As New ReglasAdmision
    Dim oRsTmp86 As New Recordset
    DevuelveCuentaEmergenciaOceDelPacienteHospitalizado = 0
    If lnIdOrigenAtencion = 30 Then
       Set oRsTmp86 = mo_ReglasAdmision.AtencionesSeleccionarPorIdPaciente(lnIdPaciente, sghTipoServicio.sghConsultaExterna)
    Else
       Set oRsTmp86 = mo_ReglasAdmision.AtencionesSeleccionarPorIdPaciente(lnIdPaciente, sghTipoServicio.sghEmergenciaConsultorios)
    End If
    If lnIdAtencionEmergencia > 0 Then
        oRsTmp86.Filter = "IdAtencion=" & lnIdAtencionEmergencia
        If oRsTmp86.RecordCount > 0 Then
           oRsTmp86.MoveFirst
           DevuelveCuentaEmergenciaOceDelPacienteHospitalizado = oRsTmp86!idCuentaAtencion
        Else
           lnIdAtencionEmergencia = 0
        End If
    Else
        oRsTmp86.Filter = "IdEstadoAtencion<>0 and fechaEgreso<>null"
        If oRsTmp86.RecordCount > 0 Then
           oRsTmp86.MoveLast
           If ldFechaIngresoHosp >= CDate(Format(oRsTmp86!fechaEgreso, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)) Then
                      DevuelveCuentaEmergenciaOceDelPacienteHospitalizado = oRsTmp86!idCuentaAtencion
                      lnIdAtencionEmergencia = oRsTmp86!idAtencion
           End If
        End If
    End If
    oRsTmp86.Close
    Set oRsTmp86 = Nothing
    Set mo_ReglasAdmision = Nothing
End Function
Function HospitalizacionConOrigenEmergenciaOconsultorios(lnIdTipoServicioActual As Long, lnIdOrigenAtencion As Long) As Boolean
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        HospitalizacionConOrigenEmergenciaOconsultorios = False
        If lnIdTipoServicioActual = sghOpcionGalenHos.sghAdmisionHospitalizacion And _
                                    (lnIdOrigenAtencion = 31 Or lnIdOrigenAtencion = 32 Or lnIdOrigenAtencion = 43) And _
                                    lcBuscaParametro.SeleccionaFilaParametro(512) = "S" Then   'Emergencia/ObsEmergencia/EmergObstetrica
           HospitalizacionConOrigenEmergenciaOconsultorios = True
        ElseIf lnIdTipoServicioActual = sghOpcionGalenHos.sghAdmisionHospitalizacion And _
                                    (lnIdOrigenAtencion = 30) And _
                                    lcBuscaParametro.SeleccionaFilaParametro(553) = "S" Then   'Consulta Externa
           HospitalizacionConOrigenEmergenciaOconsultorios = True
        End If
        Set lcBuscaParametro = Nothing
End Function

Function EsUnaClinicaNOminsa() As Boolean
        Dim lcSql As String, lcEESS As String
        Dim oRsTmp1 As New Recordset
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        lcEESS = Right("00000" & lcBuscaParametro.SeleccionaFilaParametro(208), 5)
        EsUnaClinicaNOminsa = False
        lcSql = "select * from establecimientos  where codigo='" & lcEESS & "'"
        If oRsTmp1.State = 1 Then oRsTmp1.Close
        oRsTmp1.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
        If oRsTmp1.RecordCount = 0 Then        'Es una Clinica
           EsUnaClinicaNOminsa = True
        End If
        oRsTmp1.Close
        Set oRsTmp1 = Nothing
        Set lcBuscaParametro = Nothing
End Function




'**************************** pasar a sighnegocios     'debb-29/12/2016
Function AuditoriaPorTablaFechas(lcTabla As String, ldFinicio As Date, ldFfin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AuditoriaPorTablaFechas = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AuditoriaPorTablaFechas"
        Set oParameter = .CreateParameter("@tabla", adVarChar, adParamInput, 50, lcTabla): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaInicial", adDBTimeStamp, adParamInput, 8, ldFinicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFinal", adDBTimeStamp, adParamInput, 8, ldFfin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AuditoriaPorTablaFechas = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"

End Function


'debb-01/02/2017
Function CatalogoServiciosSeleccionarSoloPtoCargaYtipoFinanciamiento(lnIdPuntoCarga As Long, lnIdtipoFinanciamiento As Long) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CatalogoServiciosSeleccionarSoloPtoCargaYtipoFinanciamiento"
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set CatalogoServiciosSeleccionarSoloPtoCargaYtipoFinanciamiento = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

'debb-15/03/2017
Function CatalogoBienesInsumosHospPorCuenta(oConexion As Connection, lnIdCuentaAtencion As Long) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
        Set CatalogoBienesInsumosHospPorCuenta = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CatalogoBienesInsumosHospPorCuenta"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRs = .Execute
   End With
   Set CatalogoBienesInsumosHospPorCuenta = oRs
   Set oRs = Nothing
   Set oCommand = Nothing
End Function

'debb-15/03/2017
Function CatalogoServiciosHospPorCuenta(oConexion As Connection, lnIdCuentaAtencion As Long) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
        Set CatalogoServiciosHospPorCuenta = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CatalogoServiciosHospPorCuenta"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRs = .Execute
   End With
   Set CatalogoServiciosHospPorCuenta = oRs
   Set oRs = Nothing
   Set oCommand = Nothing
End Function
Sub FactCatalogoBienesInsumosHospActualizaPrecio(oConexion As Connection, lnIdProducto As Long, _
                                                 lnIdtipoFinanciamiento As Long, lnPrecio As Double)
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoBienesInsumosHospActualizaPrecio"
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@precio", adCurrency, adParamInput, 0, lnPrecio): .Parameters.Append oParameter
        .Execute
   End With
   Set oCommand = Nothing
End Sub
Function TiposActividadesSeleccionarTodos() As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposActividadesSeleccionarTodos"
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set TiposActividadesSeleccionarTodos = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

Function EmpleadosChequeaNoTengaMovimientos(lnIdEmpleado As Integer) As String
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim lcMensaje As String
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 900
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosChequeaNoTengaMovimientos"
        Set oParameter = .CreateParameter("@mensaje", adVarChar, adParamOutput, 50, lcMensaje): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, lnIdEmpleado): .Parameters.Append oParameter
        .Execute
        lcMensaje = .Parameters("@mensaje")
   End With
   EmpleadosChequeaNoTengaMovimientos = lcMensaje
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
End Function





Function DevuelveFechaAfiliacionSIS(ByRef lcSIScodigo As String, ByRef lnIdSiaSis As Long, _
                                    ByRef lcAfiliacionSIS As String, lnIdFormaPago As Long, _
                                    wxParametro301 As String, wxParametro302 As String, _
                                    wxParametro322 As String, ml_DNI As String, _
                                    wxParametro323 As String, wxParametroJAMO As String, _
                                    ByRef lblMensajeNuevoPlan As String) As Date
    Dim LbSisOK As Boolean, lcSql As String
    Dim mo_SisConsumoWeb As New SIGHNegocios.SisConsumoWeb
    Dim oRsBuscaPacientesSis As New Recordset
    DevuelveFechaAfiliacionSIS = 0
    lblMensajeNuevoPlan = ""
    If lnIdFormaPago = sghTipoFinanciamiento.sghSis Then
       If wxParametro301 = "S" And wxParametro302 = "S" Then
            LbSisOK = False
            If wxParametro322 = "S" Then
               '******* busca en la WEB
               Set oRsBuscaPacientesSis = mo_SisConsumoWeb.WebServiceSISBuscarAfiliado(ml_DNI, _
                                                           "", "", _
                                                           "", "", _
                                                           "", wxParametro323)
               If oRsBuscaPacientesSis.RecordCount > 0 Then
                     LbSisOK = True
                     DevuelveFechaAfiliacionSIS = oRsBuscaPacientesSis!AfiliacionFecha
                     lnIdSiaSis = oRsBuscaPacientesSis!idSiaSis
                     lcSIScodigo = oRsBuscaPacientesSis!Codigo
                     lcAfiliacionSIS = oRsBuscaPacientesSis!cDisa & "-" & _
                                       oRsBuscaPacientesSis!cFormato & "-" & _
                                       oRsBuscaPacientesSis!cNumero
               End If
               oRsBuscaPacientesSis.Close
            End If
            If LbSisOK = False Then
               '******** busca en la bd sigh_Externa
               lcSql = "  where documentoNumero='" & ml_DNI & "'"
               Set oRsBuscaPacientesSis = SisFiltraPacientesAfiliados(lcSql, wxParametroJAMO)
               If oRsBuscaPacientesSis.RecordCount > 0 Then
                        LbSisOK = True
                        DevuelveFechaAfiliacionSIS = oRsBuscaPacientesSis!AfiliacionFecha
                        lnIdSiaSis = oRsBuscaPacientesSis!idSiaSis
                        lcSIScodigo = oRsBuscaPacientesSis!Codigo
                        lcAfiliacionSIS = oRsBuscaPacientesSis!cDisa & "-" & _
                                          oRsBuscaPacientesSis!cFormato & "-" & _
                                          oRsBuscaPacientesSis!cNumero
               End If
               oRsBuscaPacientesSis.Close
            End If
            If LbSisOK = True Then
               lblMensajeNuevoPlan = "Los CONSUMOS los cambia desde su FECHA DE AFILIACION: " & DevuelveFechaAfiliacionSIS
            End If
       End If
    End If
    Set mo_SisConsumoWeb = Nothing
    Set oRsBuscaPacientesSis = Nothing
End Function
Function SisFiltraPacientesAfiliados(lcWhereOrder As String, wxParametroJAMO As String) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ms_MensajeError  As String
    Set SisFiltraPacientesAfiliados = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SisFiltraPacientesAfiliados"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 2000, lcWhereOrder):   .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SisFiltraPacientesAfiliados = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set lcBuscaParametro = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function



Sub DevuelveCodigoYdescripcionSegunId(lnIdProducto As Long, ByRef lcCodigo As String, ByRef lcDescripcion As String)
    Dim oRsTmp As New Recordset
    lcCodigo = ""
    lcDescripcion = ""
    Set oRsTmp = CatalogoServiciosSeleccionarXidentificador(lnIdProducto)
    If oRsTmp.RecordCount > 0 Then
       lcDescripcion = oRsTmp.Fields!nombre
       lcCodigo = oRsTmp.Fields!Codigo
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Sub


Sub ServiciosGrabaYdevueveCorrelativoEmergencia(lbSoloDevuelveCorrelativo As Boolean, ByRef lcNuevoCorrelativo As String, _
                                                    lnIdServicio As Long, oConexion As Connection)
    Dim oDOServicio As New DOServicio
    Dim oServicios As New Servicios
    Set oServicios.Conexion = oConexion
    lcNuevoCorrelativo = Trim(Str(Year(Date))) & "00001"
    oDOServicio.IdServicio = lnIdServicio
    oDOServicio.IdUsuarioAuditoria = SIGHEntidades.Usuario
    If oServicios.SeleccionarPorId(oDOServicio) = True Then
       If oDOServicio.emergenciaCorrelativo <> "" Then
          lcNuevoCorrelativo = Trim(Str(Val(oDOServicio.emergenciaCorrelativo) + 1))
       End If
       If lbSoloDevuelveCorrelativo = False Then
          oDOServicio.emergenciaCorrelativo = lcNuevoCorrelativo
          If oServicios.modificar(oDOServicio) = True Then
             
          End If
       End If
    End If
    Set oDOServicio = Nothing
    Set oServicios = Nothing
End Sub
Function FuentesFinanciamientoTarifasSoloTarifa() As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ms_MensajeError  As String
    Set FuentesFinanciamientoTarifasSoloTarifa = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FuentesFinanciamientoTarifasSoloTarifa"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FuentesFinanciamientoTarifasSoloTarifa = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set lcBuscaParametro = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description

End Function

Function FuentesFinanciamientoTarifasSeleccionarPorIdFuenteFinanciamiento(lnIdtipoFinanciamiento As Long, _
                                                                          lnIdFuenteFinanciamiento As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection

Dim ms_MensajeError  As String
    Set FuentesFinanciamientoTarifasSeleccionarPorIdFuenteFinanciamiento = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FuentesFinanciamientoTarifasSeleccionarPorIdFuenteFinanciamiento"
        Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FuentesFinanciamientoTarifasSeleccionarPorIdFuenteFinanciamiento = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:

End Function
Function DevuelvePrioridadEmergencia(lcDx As String) As String
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    DevuelvePrioridadEmergencia = ""
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DiagnosticosPrioridadesXdx"
        Set oParameter = .CreateParameter("@dx", adVarChar, adParamInput, 7, Left(lcDx, 5)):  .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount > 0 Then
      DevuelvePrioridadEmergencia = oRecordset!prioridad
   End If
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function



Function RecetaDetallePorCuenta(lnIdCuentaAtencion As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
    Set RecetaDetallePorCuenta = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaDetallePorCuenta"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set RecetaDetallePorCuenta = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
End Function


Function InteroperaEquiposSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
Dim oConexion As New Connection
Dim lcBuscaParametro As New SIGHDatos.Parametros
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    Set InteroperaEquiposSeleccionarTodos = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "InteroperaEquiposSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set InteroperaEquiposSeleccionarTodos = oRecordset
   oConexion.Close
   Set oCommand = Nothing
   Set oConexion = Nothing
   Set lcBuscaParametro = Nothing
   Exit Function
ManejadorDeError:
End Function

Function EmpleadosDevuelveDNI(lnIdEmpleado As Long) As String
   Dim oRsTmp1 As New Recordset
   Set oRsTmp1 = EmpleadosSeleccionarPorIdEmpleado(lnIdEmpleado)
   EmpleadosDevuelveDNI = ""
   If oRsTmp1.RecordCount > 0 Then
      EmpleadosDevuelveDNI = Trim(oRsTmp1!DNI)
   End If
   oRsTmp1.Close
   Set oRsTmp1 = Nothing
End Function

Sub DevuelveCamaYdniMedico(ByRef lcMedico As String, ByRef lcMedicoDNI As String, ByRef lcCAma As String, _
                           lnMedicoIdEmpleado As Long, lnIdMedico As Long, ml_idPaciente As Long)
    lcMedicoDNI = "": lcCAma = "": lcMedico = ""
    Dim oRsTmp1 As New Recordset
    Dim oReglasHoteleria As New SIGHNegocios.ReglasHoteleria
    Dim oReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
    If lnIdMedico > 0 Or lnMedicoIdEmpleado > 0 Then
        If lnMedicoIdEmpleado > 0 Then
           Set oRsTmp1 = EmpleadosSeleccionarPorIdEmpleado(lnMedicoIdEmpleado)
        Else
           Set oRsTmp1 = oReglasDeProgMedica.MedicosSeleccionarXIdMedico(lnIdMedico)
        End If
        If oRsTmp1.RecordCount > 0 Then
           If Not IsNull(oRsTmp1!DNI) Then
              lcMedicoDNI = Trim(oRsTmp1!DNI)
              lcMedico = oRsTmp1!ApellidoPaterno & " " & oRsTmp1!ApellidoMaterno & " " & oRsTmp1!Nombres
           End If
        End If
        oRsTmp1.Close
    End If
    If ml_idPaciente > 0 Then
        Set oRsTmp1 = oReglasHoteleria.CamasSeleccionarPorIdPaciente(ml_idPaciente)
        oRsTmp1.Filter = "idEstadoCama=3"
        If oRsTmp1.RecordCount > 0 Then
           lcCAma = Trim(oRsTmp1!Codigo)
        End If
        oRsTmp1.Close
    End If
    Set oRsTmp1 = Nothing
    Set oReglasHoteleria = Nothing
    Set oReglasDeProgMedica = Nothing
End Sub

Function ImagMovimientoSeleccionarIdOrden(lnIdOrden As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
Dim oConexion As New Connection
Dim lcBuscaParametro As New SIGHDatos.Parametros
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    Set ImagMovimientoSeleccionarIdOrden = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagMovimientoSeleccionarIdOrden"
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ImagMovimientoSeleccionarIdOrden = oRecordset
   oConexion.Close
   Set oCommand = Nothing
   Set oConexion = Nothing
   Set lcBuscaParametro = Nothing
   Exit Function
ManejadorDeError:
End Function


Function InteroperaEquiposSeleccionarPorId(lcCodigo As String) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
Dim oConexion As New Connection
Dim lcBuscaParametro As New SIGHDatos.Parametros
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "InteroperaEquiposSeleccionarPorId"
        Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 2, lcCodigo):  .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set InteroperaEquiposSeleccionarPorId = oRecordset
   oConexion.Close
   Set oCommand = Nothing
   Set oConexion = Nothing
   Set lcBuscaParametro = Nothing
   Exit Function
ManejadorDeError:
End Function


Sub FactPuntosCargaActualizaCupos(lnIdPuntoCarga As Long, lnCupos As Long, lnCuposMinutos As Long, _
                                  lcHoraInicioCita As String)
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactPuntosCargaActualizaCupos"
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@nroCupos", adInteger, adParamInput, 0, lnCupos): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@nroCuposMinutos", adInteger, adParamInput, 0, lnCuposMinutos): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@HoraInicioDiaCita", adVarChar, adParamInput, 5, lcHoraInicioCita):  .Parameters.Append oParameter
        .Execute
   End With
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Sub
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Sub
End Sub
Function FuenteFinanciamientoEsUnidosis(lnIdFuenteFinanciamiento As Long, ByRef lnCuentaUnidosis As Long) As Boolean
    Dim oRsFF As New Recordset
    Set oRsFF = FuentesFinanciamientoSegunFiltro("idFuenteFinanciamiento=" & lnIdFuenteFinanciamiento)
    lnCuentaUnidosis = 0
    If oRsFF.RecordCount > 0 Then
       If oRsFF!usadoEnFUnidosis = 1 Then
          FuenteFinanciamientoEsUnidosis = True
          lnCuentaUnidosis = oRsFF!CuentaParaUnidosis
       End If
    End If
    oRsFF.Close
    Set oRsFF = Nothing
End Function
Function CuentaDeEmergenciaCerrada(lnIdCuenta As Long, lnIdPuntoCarga As sghPuntosCargaBasicos) As Boolean
    CuentaDeEmergenciaCerrada = False
    Dim oRsTmp2 As New Recordset
    Dim oConexion As New ADODB.Connection
    Dim oReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    Set oRsTmp2 = oReglasFarmacia.AtencionesSelecionarPorCuenta(lnIdCuenta, oConexion)
    If oRsTmp2.RecordCount > 0 Then
       If (oRsTmp2!IdEstado <> 1) Then
 '      If (oRsTmp2!IdEstado <> 1) And oRsTmp2!idTipoServicio = 2 Then
'          If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oRsTmp2!IdFormaPago, oConexion) = True Then
'             If lnIdPuntoCarga = sghPtoCargaFarmacia Then
'                CuentaDeEmergenciaCerrada = True
'                MsgBox "El Paciente es PAGANTE, la Cuenta es de EMERGENCIA y no está ABIERTA" & Chr(13) & Chr(13) & _
'                       "Deberá crear PRE-VENTA, que pague en CAJA y recién se le entrega Medicamentos/Insumos", vbInformation, ""
'             Else
'                MsgBox "El Paciente es PAGANTE, la Cuenta es de EMERGENCIA y no está ABIERTA" & Chr(13) & Chr(13) & _
'                       "Deberá pagar en CAJA la RECETA", vbInformation, ""
'             End If
'          Else
            CuentaDeEmergenciaCerrada = True
'          End If
       End If
    End If
    oRsTmp2.Close
    oConexion.Close
    Set oConexion = Nothing
    Set oRsTmp2 = Nothing
    Set oReglasFarmacia = Nothing
    Set mo_ReglasFacturacion = Nothing
End Function
Function AuditoriaPorTabla(lcTabla As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AuditoriaPorTabla = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AuditoriaPorTabla"
        Set oParameter = .CreateParameter("@tabla", adVarChar, adParamInput, 50, lcTabla): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AuditoriaPorTabla = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"

End Function
Function DevuelveNumeroColaEnCita(lnIdCuenta As Long, lcHoraCita As String) As String
    Dim oRsTmp3 As New Recordset
    Dim mo_ReglasAdmision As New ReglasAdmision
    Dim mo_ReglasDeProgMedica As New ReglasDeProgMedica
    Dim lnIdProgramacion As Long, lnNumeroCola As Integer
    Dim lHoraCita As Long
    Dim lHoraInicio As Long
    Dim lHoraFin  As Long
    Dim lTiempoPromedio As Long
    Dim lHoraSiguiente As Long
    DevuelveNumeroColaEnCita = ""
    Set oRsTmp3 = mo_ReglasAdmision.CitasSeleccionarPorNroCuenta(lnIdCuenta)
    If oRsTmp3.RecordCount > 0 Then
       lnNumeroCola = 0
       lnIdProgramacion = oRsTmp3!idProgramacion
       oRsTmp3.Close
       Set oRsTmp3 = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarXidentificador(lnIdProgramacion)
       If oRsTmp3.RecordCount > 0 Then
            lHoraCita = mo_ReglasDeProgMedica.ConvertirAMinutos(lcHoraCita)
            lHoraInicio = mo_ReglasDeProgMedica.ConvertirAMinutos(oRsTmp3!HoraInicio)
            lHoraFin = mo_ReglasDeProgMedica.ConvertirAMinutos(oRsTmp3!horaFin)
            lTiempoPromedio = oRsTmp3!TiempoPromedioAtencion
            lHoraSiguiente = lHoraInicio
            Do While lHoraSiguiente < lHoraFin
                lnNumeroCola = lnNumeroCola + 1
                If lHoraSiguiente = lHoraCita Then
                   Exit Do
                End If
                lHoraSiguiente = lHoraSiguiente + lTiempoPromedio
            Loop
            DevuelveNumeroColaEnCita = Trim(Str(lnNumeroCola))
       End If
    End If
    oRsTmp3.Close
    Set oRsTmp3 = Nothing
    Set mo_ReglasAdmision = Nothing
    Set mo_ReglasDeProgMedica = Nothing
End Function
Function RecetaCabeceraFiltrar(lcFiltro As String) As Recordset
    Dim lcSql As String
    Dim oRsTmp5 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Set RecetaCabeceraFiltrar = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaCabeceraFiltrar"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 500, lcFiltro): .Parameters.Append oParameter
        Set oRsTmp5 = .Execute
        Set oRsTmp5.ActiveConnection = Nothing
   End With
   Set RecetaCabeceraFiltrar = oRsTmp5
   oConexion.Close
   Set oRsTmp5 = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
   Set oConexion = Nothing
End Function
Function RetornaNombrePuntoCargaRecetaConsumoEnServicio(lnIdReceta As Long) As String
    Dim oRsTmp5 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    RetornaNombrePuntoCargaRecetaConsumoEnServicio = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaDetalleDevuelvePuntoCarga"
        Set oParameter = .CreateParameter("@lnIdReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
        Set oRsTmp5 = .Execute
   End With
   If oRsTmp5.RecordCount > 0 Then
      If Not IsNull(oRsTmp5!Descripcion) Then
      RetornaNombrePuntoCargaRecetaConsumoEnServicio = UCase(Trim(oRsTmp5!Descripcion))
      End If
   End If
   oRsTmp5.Close
   Set oRsTmp5 = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
   oConexion.Close
   Set oConexion = Nothing
End Function
Function RecetaCabeceraFiltraXpuntoCargaYDocumentodespacho(lcDocumentoDespacho As String, lnIdPuntoCarga As sghPuntosCargaBasicos) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaCabeceraFiltraXpuntoCargaYDocumentodespacho"
        Set oParameter = .CreateParameter("@DocumentoDespacho", adVarChar, adParamInput, 50, lcDocumentoDespacho): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set RecetaCabeceraFiltraXpuntoCargaYDocumentodespacho = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing

End Function
Function RecetaDetalleItemFiltraXidComprobantePagoYpuntoCarga(lnIdComprobantePago As Long, lnIdPuntoCarga As sghPuntosCargaBasicos) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaDetalleItemFiltraXidComprobantePagoYpuntoCarga"
        Set oParameter = .CreateParameter("@idComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set RecetaDetalleItemFiltraXidComprobantePagoYpuntoCarga = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing

End Function
Sub AnularCptGeneradasEnCitas(lnIdCuentaAtencion As Long)
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AnularCptGeneradasEnCitas"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        .Execute
   End With
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
   Exit Sub
ManejadorDeError:
  '  MsgBox "Terapias: " & Err.Description
End Sub
