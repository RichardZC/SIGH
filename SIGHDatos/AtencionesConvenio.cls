VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionesConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase:
'        Autor: DAniel Barrantes
'        Fecha: 05/09/2007
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:Mantenimiento de la tabla "Atenciones Convenio" MINSA-ESSALUD
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------
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
'   Descripción:    Inserta un registro a la tabla AtencionesConvenio
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOAtencionesConvenio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesConvenioAgregar"
           Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, oTabla.idProducto): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@nroOficio", adVarChar, adParamInput, 20, oTabla.nroOficio): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaSesion", adDBTimeStamp, adParamInput, 8, oTabla.fechaSesion): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@importeSesion", adCurrency, adParamInput, 0, oTabla.importeSesion): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionesConvenio", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
       oTabla.IdAtencionesConvenio = .Parameters("@IdAtencionesConvenio")
   End With
   Insertar = True
   ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla AtencionesConvenio
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencionesConvenio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesConvenioModificar"
           Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, oTabla.idProducto): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@nroOficio", adVarChar, adParamInput, 20, oTabla.nroOficio): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaSesion", adDBTimeStamp, adParamInput, 8, oTabla.fechaSesion): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@importeSesion", adCurrency, adParamInput, 0, oTabla.importeSesion): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionesConvenio", adInteger, adParamInput, 0, oTabla.IdAtencionesConvenio): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CajaCaja
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencionesConvenio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesConvenioEliminar"
           Set oParameter = .CreateParameter("@IdAtencionesConvenio", adInteger, adParamInput, 0, oTabla.IdAtencionesConvenio): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla AtencionesConvenio
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function SeleccionarPorId(ByVal oTabla As DOAtencionesConvenio) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
   sSql = "SELECT  dbo.AtencionesConvenio.*, LTRIM(dbo.Pacientes.ApellidoPaterno) + ' ' + LTRIM(dbo.Pacientes.ApellidoMaterno) " & _
                   " + ' ' + LTRIM(dbo.Pacientes.PrimerNombre) AS nombrePaciente, dbo.Pacientes.NroHistoriaClinica AS nroHistoriaClinica," & _
                   " dbo.FactCatalogoServicios.Codigo AS CodServicio, dbo.FactCatalogoServicios.Nombre AS dServicio" & _
          " FROM   dbo.AtencionesConvenio LEFT OUTER JOIN" & _
                   " dbo.FactCatalogoServicios ON dbo.AtencionesConvenio.idProducto = dbo.FactCatalogoServicios.IdProducto LEFT OUTER JOIN" & _
                   " dbo.Pacientes ON dbo.AtencionesConvenio.idPaciente = dbo.Pacientes.IdPaciente" & _
          " where dbo.AtencionesConvenio.IdAtencionesConvenio = " & oTabla.IdAtencionesConvenio
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdText
     Set .ActiveConnection = mo_Conexion
     .CommandText = sSql
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.fechaSesion = oRecordset!fechaSesion
       oTabla.IdPaciente = oRecordset!IdPaciente
       oTabla.importeSesion = oRecordset!importeSesion
       oTabla.nroOficio = oRecordset!nroOficio
       oTabla.nombrePaciente = oRecordset!nombrePaciente
       oTabla.idProducto = oRecordset!idProducto
       oTabla.NombreProducto = oRecordset!dServicio
       oTabla.nroHistoria = oRecordset!NroHistoriaClinica
       oTabla.CodProducto = oRecordset!CodServicio
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


Function FiltrarAtencionesConvenio(oDOPaciente As DOPaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String

     Set FiltrarAtencionesConvenio = Nothing
     
        sWhere = ""
        sSql = "SELECT   dbo.AtencionesConvenio.idAtencionesConvenio,dbo.AtencionesConvenio.fechaSesion, dbo.Pacientes.ApellidoPaterno as ApellidoPaterno,dbo.Pacientes.ApellidoMaterno as apellidoMaterno, " & _
                      " dbo.Pacientes.PrimerNombre AS PrimerNombre, dbo.Pacientes.NroHistoriaClinica AS nroHistoriaClinica," & _
                      " dbo.FactCatalogoServicios.Codigo AS CodServicio, dbo.FactCatalogoServicios.Nombre AS Servicio" & _
               " FROM  dbo.AtencionesConvenio LEFT OUTER JOIN" & _
                      " dbo.FactCatalogoServicios ON dbo.AtencionesConvenio.idProducto = dbo.FactCatalogoServicios.IdProducto LEFT OUTER JOIN" & _
                      " dbo.Pacientes ON dbo.AtencionesConvenio.idPaciente = dbo.Pacientes.IdPaciente"
       If oDOPaciente.ApellidoPaterno <> "" Then
            sWhere = sWhere + " dbo.Pacientes.ApellidoPaterno like '" + oDOPaciente.ApellidoPaterno + "%' and "
       End If
       If oDOPaciente.ApellidoMaterno <> "" Then
            sWhere = sWhere + " dbo.Pacientes.ApellidoMaterno like '" + oDOPaciente.ApellidoMaterno + "%' and "
       End If
       If oDOPaciente.PrimerNombre <> "" Then
            sWhere = sWhere + " dbo.PAcientes.PrimerNombre like '%" + oDOPaciente.PrimerNombre + "%' and "
       End If
       If oDOPaciente.NroHistoriaClinica <> 0 Then
            sWhere = sWhere + " dbo.Pacientes.NroHistoriaClinica = " & oDOPaciente.NroHistoriaClinica & " and "
       End If
        If (sWhere <> "") Then
            sSql = sSql + " Where  "
            sSql = sSql + Left(sWhere, Len(sWhere) - 4)
        End If
       sSql = sSql + " order by dbo.AtencionesConvenio.fechaSesion desc,  dbo.Pacientes.ApellidoPaterno, dbo.Pacientes.ApellidoMaterno, dbo.Pacientes.PrimerNombre"
        oRecordset.Open sSql, mo_Conexion, adOpenKeyset, adLockOptimistic
        Set FiltrarAtencionesConvenio = oRecordset
        ms_MensajeError = ""
Exit Function
ManejadorDeError:
   MsgBox Err.Number & " " + Err.Description, "FiltrarConsultaExterna"
Exit Function
End Function


'
'CREATE  procedure AtencionesConvenioAgregar
'(@idPaciente int ,
'@idProducto int ,
'@nroOficio varchar(20) ,
'@fechaSesion datetime,
'@importeSesion money,
'@IdAtencionesConvenio int output,
'@IdUsuarioAuditoria int)
' as insert into AtencionesConvenio(idPaciente,nroOficio,FechaSesion,ImporteSesion,idUsuario,idProducto)
' values (@idPaciente,@nroOficio,@fechaSesion,@importeSesion,@IdUsuarioAuditoria,@idProducto)
'set @IdAtencionesConvenio = @@Identity
'exec AuditoriaAgregar @IdUsuarioAuditoria ,'A',@IdAtencionesConvenio,'AtencionesConvenio'
'GO
'
'
'create procedure AtencionesConvenioEliminar
'(@IdAtencionesConvenio int ,
'@IdUsuarioAuditoria int)
'as delete from AtencionesConvenio
'Where
'IdAtencionesConvenio = @IdAtencionesConvenio
'exec AuditoriaAgregar @IdUsuarioAuditoria ,'E',@IdAtencionesConvenio,'AtencionesConvenio'
'GO
'
'
'CREATE  procedure AtencionesConvenioModificar
'(@idPaciente int ,
'@idProducto int,
'@nroOficio varchar(20) ,
'@fechaSesion datetime,
'@importeSesion money,
'@IdAtencionesConvenio int ,
'@IdUsuarioAuditoria int)
'as update AtencionesConvenio
'set
'idPaciente=@idPaciente,
'nroOficio=@nroOficio,
'fechaSesion=@fechaSesion,
'importeSesion=@importeSesion,
'idProducto=@idProducto
'where IdAtencionesConvenio = @IdAtencionesConvenio
'exec AuditoriaAgregar @IdUsuarioAuditoria ,'M',@IdAtencionesConvenio,'AtencionesConvenio'
'GO
