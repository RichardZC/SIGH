VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionProcedimientoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: AtencionProcedimientoDetalle
'        Autor: William Castro Grijalva
'        Fecha: 14/06/2005 10:49:27 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
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
'   Descripción:    Inserta un registro a la tabla AtencionProcedimientoDetalle
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOAtencionProcDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionProcedimientoDetalleAgregar"
           Set oParameter = .CreateParameter("@IdAtencionProcedimiento", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionProcedimiento = 0, Null, oTabla.IdAtencionProcedimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFacturacionServicio", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionServicio = 0, Null, oTabla.IdFacturacionServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoRealiza", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoRealiza = 0, Null, oTabla.IdMedicoRealiza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioRealiza", adInteger, adParamInput, 0, IIf(oTabla.IdServicioRealiza = 0, Null, oTabla.IdServicioRealiza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProcedimiento", adInteger, adParamInput, 0, IIf(oTabla.IdProcedimiento = 0, Null, oTabla.IdProcedimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraRealizacion", adChar, adParamInput, 5, IIf(oTabla.HoraRealizacion = "", Null, oTabla.HoraRealizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaRealizacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaRealizacion = 0, Null, oTabla.FechaRealizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionProcDetalle", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdAtencionProcDetalle = .Parameters("@IdAtencionProcDetalle")
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
'   Descripción:    modifica un registro a la tabla AtencionProcedimientoDetalle
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencionProcDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionProcedimientoDetalleModificar"
           Set oParameter = .CreateParameter("@IdAtencionProcedimiento", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionProcedimiento = 0, Null, oTabla.IdAtencionProcedimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFacturacionServicio", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionServicio = 0, Null, oTabla.IdFacturacionServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoRealiza", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoRealiza = 0, Null, oTabla.IdMedicoRealiza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioRealiza", adInteger, adParamInput, 0, IIf(oTabla.IdServicioRealiza = 0, Null, oTabla.IdServicioRealiza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProcedimiento", adInteger, adParamInput, 0, IIf(oTabla.IdProcedimiento = 0, Null, oTabla.IdProcedimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraRealizacion", adChar, adParamInput, 5, IIf(oTabla.HoraRealizacion = "", Null, oTabla.HoraRealizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaRealizacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaRealizacion = 0, Null, oTabla.FechaRealizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionProcDetalle", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionProcDetalle = 0, Null, oTabla.IdAtencionProcDetalle)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla AtencionProcedimientoDetalle
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencionProcDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionProcedimientoDetalleEliminar"
           Set oParameter = .CreateParameter("@IdAtencionProcDetalle", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionProcDetalle = 0, Null, oTabla.IdAtencionProcDetalle)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla AtencionProcedimientoDetalle
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtencionProcDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionProcedimientoDetalleSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdAtencionProcDetalle", adInteger, adParamInput, 0, oTabla.IdAtencionProcDetalle): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdAtencionProcedimiento = IIf(IsNull(oRecordset!IdAtencionProcedimiento), 0, oRecordset!IdAtencionProcedimiento)
       oTabla.IdFacturacionServicio = IIf(IsNull(oRecordset!IdFacturacionServicio), 0, oRecordset!IdFacturacionServicio)
       oTabla.IdMedicoRealiza = IIf(IsNull(oRecordset!IdMedicoRealiza), 0, oRecordset!IdMedicoRealiza)
       oTabla.IdServicioRealiza = IIf(IsNull(oRecordset!IdServicioRealiza), 0, oRecordset!IdServicioRealiza)
       oTabla.IdProcedimiento = IIf(IsNull(oRecordset!IdProcedimiento), 0, oRecordset!IdProcedimiento)
       oTabla.HoraRealizacion = IIf(IsNull(oRecordset!HoraRealizacion), "", oRecordset!HoraRealizacion)
       oTabla.FechaRealizacion = IIf(IsNull(oRecordset!FechaRealizacion), 0, oRecordset!FechaRealizacion)
       oTabla.IdAtencionProcDetalle = IIf(IsNull(oRecordset!IdAtencionProcDetalle), 0, oRecordset!IdAtencionProcDetalle)
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


Function SeleccionarPorIdAtencionProcedimiento(lIdAtencionProcedimiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String
Dim sWhere As String

        Set SeleccionarPorIdAtencionProcedimiento = Nothing
     
        sSQL = "select AtencionProcedimientoDetalle.IdAtencionProcDetalle, " + _
        "AtencionProcedimientoDetalle.IdProcedimiento, " + _
        " Procedimientos.CodigoCPT2004 as CodigoCPT , " + _
        " Procedimientos.Descripcion, " + _
        " AtencionProcedimientoDetalle.IdMedicoRealiza, " + _
        " Empleados.ApellidoPaterno + ' ' + Empleados.ApellidoMaterno + ' ' + Empleados.Nombres as NombreMedico, " + _
        " AtencionProcedimientoDetalle.IdServicioRealiza, " + _
        " Servicios.Nombre as NombreServicio, " + _
        " AtencionProcedimientoDetalle.FechaRealizacion, " + _
        " AtencionProcedimientoDetalle.HoraRealizacion, " + _
        " AtencionProcedimientoDetalle.IdFacturacionServicio " + _
        " From " + _
        " (((AtencionProcedimientoDetalle " + _
        " left join Servicios on AtencionProcedimientoDetalle.IdServicioRealiza = Servicios.IdServicio) " + _
        " left join Medicos on AtencionProcedimientoDetalle.IdMedicoRealiza = Medicos.IdMedico) " + _
        " left join Empleados on Medicos.IdEmpleado = Empleados.IdEmpleado) " + _
        " left join Procedimientos on AtencionProcedimientoDetalle.IdProcedimiento =  Procedimientos.IdProcedimiento " + _
        " where IdAtencionProcedimiento = " & lIdAtencionProcedimiento & _
        " order by FechaRealizacion desc"
                
        With oCommand
           .CommandType = adCmdText
           Set .ActiveConnection = mo_Conexion
           .CommandText = sSQL
           Set oRecordset = .Execute
           Set oRecordset.ActiveConnection = Nothing
        End With
    
        Set SeleccionarPorIdAtencionProcedimiento = oRecordset
    
        ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function
Function Filtrar(ByVal oTabla As DOAtencionProcDetalle) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String
Dim sWhere As String

     Set Filtrar = Nothing
     
'        sWhere = ""
'        sSQL = "select Pacientes.IdPaciente, Pacientes.ApellidoPaterno," + _
'        " Pacientes.ApellidoMaterno, Pacientes.PrimerNombre, Pacientes.SegundoNombre," + _
'        " Pacientes.NroHistoriaClinica, TiposNumeracionHistoria.Descripcion as TipoNumeracion, Pacientes.IdTipoNumeracion" + _
'        " from Pacientes left join TiposNumeracionHistoria " + _
'        " on Pacientes.IdTipoNumeracion  = TiposNumeracionHistoria.IdTipoNumeracion "
'
'       If oTabla.ApellidoPaterno <> "" Then
'            sWhere = sWhere + " Pacientes.ApellidoPaterno like '" + oTabla.ApellidoPaterno + "%' and "
'       End If
'       If oTabla.ApellidoMaterno <> "" Then
'            sWhere = sWhere + " Pacientes.ApellidoMaterno like '" + oTabla.ApellidoMaterno + "%' and "
'       End If
'       If oTabla.PrimerNombre <> "" Then
'            sWhere = sWhere + " Pacientes.PrimerNombre like '" + oTabla.PrimerNombre + "%' and "
'       End If
'       If oTabla.SegundoNombre <> "" Then
'            sWhere = sWhere + " Pacientes.SegundoNombre like '" + oTabla.SegundoNombre + "%' and "
'       End If
'       If oTabla.NroHistoriaClinica <> 0 Then
'            sWhere = sWhere + " Pacientes.NroHistoriaClinica = " & oTabla.NroHistoriaClinica & " and "
'       End If
'       If oTabla.NroDocumento <> "" Then
'            sWhere = sWhere + " Pacientes.NroDocumento ='" & oTabla.NroDocumento & "' and "
'       End If
'
'       If sWhere <> "" Then
'            sSQL = sSQL + " where " + Left(sWhere, Len(sWhere) - 4)
'       End If
'
'        sSQL = sSQL + " order by Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre, Pacientes.SegundoNombre "
        
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandText = sSQL
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    
    Set Filtrar = oRecordset
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function EliminarPorIdAtencionProcedimiento(lIdAtencionProcedimiento As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSQL As String
 
   EliminarPorIdAtencionProcedimiento = False
   
   sSQL = "delete from AtencionProcedimientoDetalle where IdAtencionProcedimiento = " & lIdAtencionProcedimiento
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSQL
       .Execute
   End With
   
   EliminarPorIdAtencionProcedimiento = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



