VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionesInterconsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: AtencionesInterconsultas
'        Autor: William Castro Grijalva
'        Fecha: 14/04/2006 07:06:50 p.m.
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
'   Descripción:    Inserta un registro a la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOAtencionInterconsulta) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesInterconsultasAgregar"
           Set oParameter = .CreateParameter("@IdDetalleProducto", adInteger, adParamInput, 0, IIf(oTabla.IdDetalleProducto = 0, Null, oTabla.IdDetalleProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdAtencion = 0, Null, oTabla.IdAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraSolicitud", adChar, adParamInput, 5, IIf(oTabla.HoraSolicitud = "", Null, oTabla.HoraSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraRealizacion", adChar, adParamInput, 5, IIf(oTabla.HoraRealizacion = "", Null, oTabla.HoraRealizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaSolicitud", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaSolicitud = 0, Null, oTabla.FechaSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaRealizacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaRealizacion = 0, Null, oTabla.FechaRealizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoRealiza", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoRealiza = 0, Null, oTabla.IdMedicoRealiza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoSolicita", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoSolicita = 0, Null, oTabla.IdMedicoSolicita)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdInterconsulta", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdInterconsulta = .Parameters("@IdInterconsulta")
   End With
 
   Insertar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla AtencionesInterconsultas
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencionInterconsulta) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesInterconsultasModificar"
           Set oParameter = .CreateParameter("@IdDetalleProducto", adInteger, adParamInput, 0, IIf(oTabla.IdDetalleProducto = 0, Null, oTabla.IdDetalleProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdAtencion = 0, Null, oTabla.IdAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraSolicitud", adChar, adParamInput, 5, IIf(oTabla.HoraSolicitud = "", Null, oTabla.HoraSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraRealizacion", adChar, adParamInput, 5, IIf(oTabla.HoraRealizacion = "", Null, oTabla.HoraRealizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaSolicitud", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaSolicitud = 0, Null, oTabla.FechaSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaRealizacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaRealizacion = 0, Null, oTabla.FechaRealizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoRealiza", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoRealiza = 0, Null, oTabla.IdMedicoRealiza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoSolicita", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoSolicita = 0, Null, oTabla.IdMedicoSolicita)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdInterconsulta", adInteger, adParamInput, 0, IIf(oTabla.IdInterconsulta = 0, Null, oTabla.IdInterconsulta)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencionInterconsulta) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesInterconsultasEliminar"
           Set oParameter = .CreateParameter("@IdInterconsulta", adInteger, adParamInput, 0, IIf(oTabla.IdInterconsulta = 0, Null, oTabla.IdInterconsulta)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtencionInterconsulta) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionesInterconsultasSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdInterconsulta", adInteger, adParamInput, 0, oTabla.IdInterconsulta): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdDetalleProducto = IIf(IsNull(oRecordset!IdDetalleProducto), 0, oRecordset!IdDetalleProducto)
       oTabla.IdAtencion = IIf(IsNull(oRecordset!IdAtencion), 0, oRecordset!IdAtencion)
       oTabla.HoraSolicitud = IIf(IsNull(oRecordset!HoraSolicitud), "", oRecordset!HoraSolicitud)
       oTabla.HoraRealizacion = IIf(IsNull(oRecordset!HoraRealizacion), "", oRecordset!HoraRealizacion)
       oTabla.FechaSolicitud = IIf(IsNull(oRecordset!FechaSolicitud), 0, oRecordset!FechaSolicitud)
       oTabla.FechaRealizacion = IIf(IsNull(oRecordset!FechaRealizacion), 0, oRecordset!FechaRealizacion)
       oTabla.IdMedicoRealiza = IIf(IsNull(oRecordset!IdMedicoRealiza), 0, oRecordset!IdMedicoRealiza)
       oTabla.IdMedicoSolicita = IIf(IsNull(oRecordset!IdMedicoSolicita), 0, oRecordset!IdMedicoSolicita)
       oTabla.IdInterconsulta = IIf(IsNull(oRecordset!IdInterconsulta), 0, oRecordset!IdInterconsulta)
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

Function FiltrarConsultaExterna(oDOPaciente As doPaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String
Dim sWhere As String

     Set FiltrarConsultaExterna = Nothing
     
        sWhere = ""
        
'        sSQL = "select AtencionesInterconsultas.IdInterconsulta, CuentasAtencion.IdCuentaAtencion," + _
'        "Atenciones.IdAtencion, Pacientes.ApellidoPaterno + ' ' + Pacientes.ApellidoMaterno + ' ' +  IsNull(Pacientes.PrimerNombre,'') + ' ' + IsNull(Pacientes.SegundoNombre,'') as NombrePaciente, " + _
'        "convert(char(10),AtencionesInterconsultas.FechaRealizacion,103) as FechaRealizacion," + _
'        "AtencionesInterconsultas.HoraRealizacion, " + _
'        "convert(char(10),AtencionesInterconsultas.FechaSolicitud,103) as FechaSolicitud, " + _
'        "AtencionesInterconsultas.HoraSolicitud, " + _
'        " " + _
'        "Empleados.ApellidoPaterno+ ' ' + Empleados.ApellidoMaterno + ' ' + Empleados.Nombres as NombreMedico " + _
'        "from ((((AtencionesInterconsultas " + _
'        "left join CuentasAtencion on AtencionesInterconsultas.IdCuentaAtencion = CuentasAtencion.IdCuentaAtencion) " + _
'        "left join Atenciones on CuentasAtencion.IdAtencion = Atenciones.IdAtencion) " + _
'        "left join Pacientes on Atenciones.IdPaciente = Pacientes.IdPaciente) " + _
'        "left join Medicos on AtencionesInterconsultas.IdMedicoRealiza = Medicos.IdMedico) " + _
'        "left join Empleados on Medicos.IdEmpleado = Empleados.IdEmpleado "

        sSQL = "select AtencionesInterconsultas.IdInterconsulta, FacturacionCuentasAtencion.IdCuentaAtencion," + _
                "Atenciones.IdAtencion, Pacientes.ApellidoPaterno + ' ' + Pacientes.ApellidoMaterno + ' ' +  IsNull(Pacientes.PrimerNombre,'') + ' ' + IsNull(Pacientes.SegundoNombre,'') as NombrePaciente, " + _
                "convert(char(10),AtencionesInterconsultas.FechaRealizacion,103) as FechaRealizacion," + _
                "AtencionesInterconsultas.HoraRealizacion, " + _
                "convert(char(10),AtencionesInterconsultas.FechaSolicitud,103) as FechaSolicitud, " + _
                "AtencionesInterconsultas.HoraSolicitud, " + _
                " " + _
                "Empleados.ApellidoPaterno+ ' ' + Empleados.ApellidoMaterno + ' ' + Empleados.Nombres as NombreMedico " + _
                "from ((((AtencionesInterconsultas " + _
                "left join Atenciones on AtencionesInterconsultas.IdAtencion = Atenciones.IdAtencion) " + _
                "left join FacturacionCuentasAtencion on Atenciones.IdCuentaAtencion = FacturacionCuentasAtencion.IdCuentaAtencion) " + _
                "left join Pacientes on Atenciones.IdPaciente = Pacientes.IdPaciente) " + _
                "left join Medicos on AtencionesInterconsultas.IdMedicoRealiza = Medicos.IdMedico) " + _
                "left join Empleados on Medicos.IdEmpleado = Empleados.IdEmpleado "

       If oDOPaciente.ApellidoPaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoPaterno like '" + oDOPaciente.ApellidoPaterno + "%' and "
       End If
       If oDOPaciente.ApellidoMaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoMaterno like '" + oDOPaciente.ApellidoMaterno + "%' and "
       End If
       If oDOPaciente.PrimerNombre <> "" Then
            sWhere = sWhere + " Pacientes.Nombres like '%" + oDOPaciente.PrimerNombre + "%' and "
       End If
       If oDOPaciente.NroHistoriaClinica <> 0 Then
            sWhere = sWhere + " Pacientes.NroHistoriaClinica = " & oDOPaciente.NroHistoriaClinica & " and "
       End If
       
        sSQL = sSQL + " Where Atenciones.IdTipoServicio = 1 "
        If (sWhere <> "") Then
            sSQL = sSQL + " and " + Left(sWhere, Len(sWhere) - 4)
        End If
        
       sSQL = sSQL + " order by Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
       
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandText = sSQL
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    
    Set FiltrarConsultaExterna = oRecordset
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FiltrarConsultorioEmergencia(oDOPaciente As doPaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String
Dim sWhere As String

     Set FiltrarConsultorioEmergencia = Nothing
     
        sWhere = ""
        
        sSQL = "select AtencionesInterconsultas.IdInterconsulta, CuentasAtencion.IdCuentaAtencion," + _
        "Atenciones.IdAtencion, Pacientes.ApellidoPaterno + ' ' + Pacientes.ApellidoMaterno + ' ' +  IsNull(Pacientes.PrimerNombre,'') + ' ' + IsNull(Pacientes.SegundoNombre,'') as NombrePaciente, " + _
        "convert(char(10),AtencionesInterconsultas.FechaRealizacion,103) as FechaRealizacion," + _
        "AtencionesInterconsultas.HoraRealizacion, " + _
        "convert(char(10),AtencionesInterconsultas.FechaSolicitud,103) as FechaSolicitud, " + _
        "AtencionesInterconsultas.HoraSolicitud, " + _
        " " + _
        "Empleados.ApellidoPaterno+ ' ' + Empleados.ApellidoMaterno + ' ' + Empleados.Nombres as NombreMedico " + _
        "from ((((AtencionesInterconsultas " + _
        "left join CuentasAtencion on AtencionesInterconsultas.IdCuentaAtencion = CuentasAtencion.IdCuentaAtencion) " + _
        "left join Atenciones on CuentasAtencion.IdAtencion = Atenciones.IdAtencion) " + _
        "left join Pacientes on Atenciones.IdPaciente = Pacientes.IdPaciente) " + _
        "left join Medicos on AtencionesInterconsultas.IdMedicoRealiza = Medicos.IdMedico) " + _
        "left join Empleados on Medicos.IdEmpleado = Empleados.IdEmpleado "

       If oDOPaciente.ApellidoPaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoPaterno like '" + oDOPaciente.ApellidoPaterno + "%' and "
       End If
       If oDOPaciente.ApellidoMaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoMaterno like '" + oDOPaciente.ApellidoMaterno + "%' and "
       End If
       If oDOPaciente.PrimerNombre <> "" Then
            sWhere = sWhere + " PAcientes.Nombres like '%" + oDOPaciente.PrimerNombre + "%' and "
       End If
       If oDOPaciente.NroHistoriaClinica <> 0 Then
            sWhere = sWhere + " Pacientes.NroHistoriaClinica = " & oDOPaciente.NroHistoriaClinica & " and "
       End If
       
        sSQL = sSQL + " Where Atenciones.IdTipoServicio = 2 "
        If (sWhere <> "") Then
            sSQL = sSQL + " and " + Left(sWhere, Len(sWhere) - 4)
        End If
       
       sSQL = sSQL + " order by Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
       
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandText = sSQL
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    
    Set FiltrarConsultorioEmergencia = oRecordset
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FiltrarObservacionEmergencia(oDOPaciente As doPaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String
Dim sWhere As String

     Set FiltrarObservacionEmergencia = Nothing
     
        sWhere = ""
        
        sSQL = "select AtencionesInterconsultas.IdInterconsulta, CuentasAtencion.IdCuentaAtencion," + _
        "Atenciones.IdAtencion, Pacientes.ApellidoPaterno + ' ' + Pacientes.ApellidoMaterno + ' ' +  IsNull(Pacientes.PrimerNombre,'') + ' ' + IsNull(Pacientes.SegundoNombre,'') as NombrePaciente, " + _
        "convert(char(10),AtencionesInterconsultas.FechaRealizacion,103) as FechaRealizacion," + _
        "AtencionesInterconsultas.HoraRealizacion, " + _
        "convert(char(10),AtencionesInterconsultas.FechaSolicitud,103) as FechaSolicitud, " + _
        "AtencionesInterconsultas.HoraSolicitud, " + _
        "" + _
        "Empleados.ApellidoPaterno+ ' ' + Empleados.ApellidoMaterno + ' ' + Empleados.Nombres as NombreMedico " + _
        "from ((((AtencionesInterconsultas " + _
        "left join CuentasAtencion on AtencionesInterconsultas.IdCuentaAtencion = CuentasAtencion.IdCuentaAtencion) " + _
        "left join Atenciones on CuentasAtencion.IdAtencion = Atenciones.IdAtencion) " + _
        "left join Pacientes on Atenciones.IdPaciente = Pacientes.IdPaciente) " + _
        "left join Medicos on AtencionesInterconsultas.IdMedicoRealiza = Medicos.IdMedico) " + _
        "left join Empleados on Medicos.IdEmpleado = Empleados.IdEmpleado "

       If oDOPaciente.ApellidoPaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoPaterno like '" + oDOPaciente.ApellidoPaterno + "%' and "
       End If
       If oDOPaciente.ApellidoMaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoMaterno like '" + oDOPaciente.ApellidoMaterno + "%' and "
       End If
       If oDOPaciente.PrimerNombre <> "" Then
            sWhere = sWhere + " PAcientes.Nombres like '%" + oDOPaciente.PrimerNombre + "%' and "
       End If
       If oDOPaciente.NroHistoriaClinica <> 0 Then
            sWhere = sWhere + " Pacientes.NroHistoriaClinica = " & oDOPaciente.NroHistoriaClinica & " and "
       End If
       
        sSQL = sSQL + " Where Atenciones.IdTipoServicio = 4 "
        If (sWhere <> "") Then
            sSQL = sSQL + " and " + Left(sWhere, Len(sWhere) - 4)
        End If
        
       sSQL = sSQL + " order by Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
       
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandText = sSQL
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    
    Set FiltrarObservacionEmergencia = oRecordset
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FiltrarHospitalizacion(oDOPaciente As doPaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String
Dim sWhere As String

        Set FiltrarHospitalizacion = Nothing
     
        sWhere = ""
        
        sSQL = "select AtencionesInterconsultas.IdInterconsulta, CuentasAtencion.IdCuentaAtencion," + _
        "Atenciones.IdAtencion, Pacientes.ApellidoPaterno + ' ' + Pacientes.ApellidoMaterno + ' ' +  IsNull(Pacientes.PrimerNombre,'') + ' ' + IsNull(Pacientes.SegundoNombre,'') as NombrePaciente, " + _
        "convert(char(10),AtencionesInterconsultas.FechaRealizacion,103) as FechaRealizacion," + _
        "AtencionesInterconsultas.HoraRealizacion, " + _
        "convert(char(10),AtencionesInterconsultas.FechaSolicitud,103) as FechaSolicitud, " + _
        "AtencionesInterconsultas.HoraSolicitud, " + _
        "" + _
        "Empleados.ApellidoPaterno+ ' ' + Empleados.ApellidoMaterno + ' ' + Empleados.Nombres as NombreMedico " + _
        "from ((((AtencionesInterconsultas " + _
        "left join CuentasAtencion on AtencionesInterconsultas.IdCuentaAtencion = CuentasAtencion.IdCuentaAtencion) " + _
        "left join Atenciones on CuentasAtencion.IdAtencion = Atenciones.IdAtencion) " + _
        "left join Pacientes on Atenciones.IdPaciente = Pacientes.IdPaciente) " + _
        "left join Medicos on AtencionesInterconsultas.IdMedicoRealiza = Medicos.IdMedico) " + _
        "left join Empleados on Medicos.IdEmpleado = Empleados.IdEmpleado "

       If oDOPaciente.ApellidoPaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoPaterno like '" + oDOPaciente.ApellidoPaterno + "%' and "
       End If
       If oDOPaciente.ApellidoMaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoMaterno like '" + oDOPaciente.ApellidoMaterno + "%' and "
       End If
       If oDOPaciente.PrimerNombre <> "" Then
            sWhere = sWhere + " PAcientes.Nombres like '%" + oDOPaciente.PrimerNombre + "%' and "
       End If
       If oDOPaciente.NroHistoriaClinica <> 0 Then
            sWhere = sWhere + " Pacientes.NroHistoriaClinica = " & oDOPaciente.NroHistoriaClinica & " and "
       End If
       
        sSQL = sSQL + " Where Atenciones.IdTipoServicio = 3 "
        If (sWhere <> "") Then
            sSQL = sSQL + " and " + Left(sWhere, Len(sWhere) - 4)
        End If
        
       sSQL = sSQL + " order by Pacientes.ApellidoPaterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
       
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandText = sSQL
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    
    Set FiltrarHospitalizacion = oRecordset
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function ActualizarDiagnosticosInterconsultas(oDiagnosticos As Collection, lIdInterconsulta As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSQL As String
Dim oDOInterconsultaDiagnostico As New DOInterconsultaDiagnostico
Dim oDOAtencionDiagnostico As DOAtencionDiagnostico
Dim oInterconsultaDiagnostico As New InterconsultasDiagnosticos
 
   ActualizarDiagnosticosInterconsultas = False
   
'   ''elimina solo los dx de interconsultas
   sSQL = "delete from InterconsultasDiagnosticos where IdInterconsulta = " & lIdInterconsulta
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSQL
       .Execute
   End With

    Set oInterconsultaDiagnostico.Conexion = mo_Conexion
   If Not oDiagnosticos Is Nothing Then
        For Each oDOAtencionDiagnostico In oDiagnosticos
            oDOInterconsultaDiagnostico.IdInterconsulta = lIdInterconsulta
            oDOInterconsultaDiagnostico.IdDiagnostico = oDOAtencionDiagnostico.IdDiagnostico
            oDOInterconsultaDiagnostico.IdClasificacionDx = oDOAtencionDiagnostico.IdClasificacionDx
            oDOInterconsultaDiagnostico.IdSubclasificacionDx = oDOAtencionDiagnostico.IdSubclasificacionDx
            oDOInterconsultaDiagnostico.IdUsuarioAuditoria = oDOAtencionDiagnostico.IdUsuarioAuditoria
            If Not oInterconsultaDiagnostico.Insertar(oDOInterconsultaDiagnostico) Then
                Exit Function
            End If
         Next
    End If
    
   ActualizarDiagnosticosInterconsultas = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function SeleccionarDiagnosticosPorInterconsulta(lIdInterconsulta As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarDiagnosticosPorInterconsulta = Nothing
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtencionesInterconsultasSeleccionarDiagnosticos"
        Set oParameter = .CreateParameter("@lIdInterconsulta", adInteger, adParamInput, 4, lIdInterconsulta): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarDiagnosticosPorInterconsulta = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

