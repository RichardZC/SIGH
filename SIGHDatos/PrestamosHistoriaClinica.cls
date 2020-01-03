VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrestamosHistoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: PrestamosHistoriaClinica
'        Autor: William Castro Grijalva
'        Fecha: 15/01/2005 03:29:02 p.m.
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
'   Descripción:    Inserta un registro a la tabla PrestamosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOPrestamoHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "PrestamosHistoriaClinicaAgregar"
           Set oParameter = .CreateParameter("@IdMotivo", adInteger, adParamInput, 4, IIf(oTabla.IdMotivo = 0, Null, oTabla.IdMotivo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaPrestamoRequerida", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaPrestamoRequerida = 0, Null, oTabla.FechaPrestamoRequerida)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraPrestamoRequerida", adChar, adParamInput, 5, IIf(oTabla.HoraPrestamoRequerida = "", Null, oTabla.HoraPrestamoRequerida)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaPrestamoReal", adDBTimeStamp, adParamInput, 10, IIf(oTabla.FechaPrestamoReal = 0, Null, oTabla.FechaPrestamoReal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPrestamo", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaSolicitud", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaSolicitud = 0, Null, oTabla.FechaSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraSolicitud", adChar, adParamInput, 5, IIf(oTabla.HoraSolicitud = "", Null, oTabla.HoraSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoPrestamo", adInteger, adParamInput, 4, IIf(oTabla.IdEstadoPrestamo = 0, Null, oTabla.IdEstadoPrestamo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 4, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEnvio", adInteger, adParamInput, 4, IIf(oTabla.IdEnvio = 0, Null, oTabla.IdEnvio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 100, IIf(oTabla.Observacion = "", Null, oTabla.Observacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 4, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraPrestamoReal", adChar, adParamInput, 5, IIf(oTabla.HoraPrestamoReal = "", Null, oTabla.HoraPrestamoReal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraDevolucion", adChar, adParamInput, 5, IIf(oTabla.HoraDevolucion = "", Null, oTabla.HoraDevolucion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroFolios", adInteger, adParamInput, 4, IIf(oTabla.NroFolios = 0, Null, oTabla.NroFolios)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaDevolucion", adDBTimeStamp, adParamInput, 4, IIf(oTabla.FechaDevolucion = 0, Null, oTabla.FechaDevolucion)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdPrestamo = .Parameters("@IdPrestamo")
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
'   Descripción:    modifica un registro a la tabla PrestamosHistoriaClinica
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOPrestamoHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "PrestamosHistoriaClinicaModificar"
           Set oParameter = .CreateParameter("@IdMotivo", adInteger, adParamInput, 4, IIf(oTabla.IdMotivo = 0, Null, oTabla.IdMotivo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaPrestamoRequerida", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaPrestamoRequerida = 0, Null, oTabla.FechaPrestamoRequerida)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraPrestamoRequerida", adChar, adParamInput, 5, IIf(oTabla.HoraPrestamoRequerida = "", Null, oTabla.HoraPrestamoRequerida)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaPrestamoReal", adDBTimeStamp, adParamInput, 10, IIf(oTabla.FechaPrestamoReal = 0, Null, oTabla.FechaPrestamoReal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPrestamo", adInteger, adParamInput, 4, IIf(oTabla.IdPrestamo = 0, Null, oTabla.IdPrestamo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaSolicitud", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaSolicitud = 0, Null, oTabla.FechaSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraSolicitud", adChar, adParamInput, 5, IIf(oTabla.HoraSolicitud = "", Null, oTabla.HoraSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoPrestamo", adInteger, adParamInput, 4, IIf(oTabla.IdEstadoPrestamo = 0, Null, oTabla.IdEstadoPrestamo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 4, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEnvio", adInteger, adParamInput, 4, IIf(oTabla.IdEnvio = 0, Null, oTabla.IdEnvio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 100, IIf(oTabla.Observacion = "", Null, oTabla.Observacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 4, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraPrestamoReal", adChar, adParamInput, 5, IIf(oTabla.HoraPrestamoReal = "", Null, oTabla.HoraPrestamoReal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraDevolucion", adChar, adParamInput, 5, IIf(oTabla.HoraDevolucion = "", Null, oTabla.HoraDevolucion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroFolios", adInteger, adParamInput, 4, IIf(oTabla.NroFolios = 0, Null, oTabla.NroFolios)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaDevolucion", adDBTimeStamp, adParamInput, 4, IIf(oTabla.FechaDevolucion = 0, Null, oTabla.FechaDevolucion)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla PrestamosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOPrestamoHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "PrestamosHistoriaClinicaEliminar"
           Set oParameter = .CreateParameter("@IdPrestamo", adInteger, adParamInput, 4, IIf(oTabla.IdPrestamo = 0, Null, oTabla.IdPrestamo)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla PrestamosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOPrestamoHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "PrestamosHistoriaClinicaSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdPrestamo", adInteger, adParamInput, 4, oTabla.IdPrestamo): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdMotivo = IIf(IsNull(oRecordset!IdMotivo), 0, oRecordset!IdMotivo)
       oTabla.FechaPrestamoRequerida = IIf(IsNull(oRecordset!FechaPrestamoRequerida), 0, oRecordset!FechaPrestamoRequerida)
       oTabla.HoraPrestamoRequerida = IIf(IsNull(oRecordset!HoraPrestamoRequerida), "", oRecordset!HoraPrestamoRequerida)
       oTabla.FechaPrestamoReal = IIf(IsNull(oRecordset!FechaPrestamoReal), 0, oRecordset!FechaPrestamoReal)
       oTabla.IdPrestamo = IIf(IsNull(oRecordset!IdPrestamo), 0, oRecordset!IdPrestamo)
       oTabla.FechaSolicitud = IIf(IsNull(oRecordset!FechaSolicitud), 0, oRecordset!FechaSolicitud)
       oTabla.HoraSolicitud = IIf(IsNull(oRecordset!HoraSolicitud), "", oRecordset!HoraSolicitud)
       oTabla.IdEstadoPrestamo = IIf(IsNull(oRecordset!IdEstadoPrestamo), 0, oRecordset!IdEstadoPrestamo)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.IdEnvio = IIf(IsNull(oRecordset!IdEnvio), 0, oRecordset!IdEnvio)
       oTabla.Observacion = IIf(IsNull(oRecordset!Observacion), "", oRecordset!Observacion)
       oTabla.IdServicio = IIf(IsNull(oRecordset!IdServicio), "", oRecordset!IdServicio)
       oTabla.HoraPrestamoReal = IIf(IsNull(oRecordset!HoraPrestamoReal), "", oRecordset!HoraPrestamoReal)
       oTabla.HoraDevolucion = IIf(IsNull(oRecordset!HoraDevolucion), "", oRecordset!HoraDevolucion)
       oTabla.NroFolios = IIf(IsNull(oRecordset!NroFolios), 0, oRecordset!NroFolios)
       oTabla.FechaDevolucion = IIf(IsNull(oRecordset!FechaDevolucion), 0, oRecordset!FechaDevolucion)
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



Function Filtrar(ByVal oDOPaciente As doPaciente, ByVal oDOPrestamo As DOPrestamoHistoriaClinica) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String
Dim sWhere As String

     Set Filtrar = Nothing
     
        sWhere = ""
        sSQL = "select PrestamosHistoriaClinica.IdPrestamo, HistoriasClinicas.IdHistoriaClinica as HistoriaClinica,"
        sSQL = sSQL + " Pacientes.ApellidoPaterno + ' ' + Pacientes.ApellidoMaterno + ' ' + Pacientes.PrimerNombre as Nombres,"
        sSQL = sSQL + " convert(char(10), PrestamosHistoriaClinica.FechaSolicitud,103) as FechaSolicitud,"
        sSQL = sSQL + " convert(char(10), PrestamosHistoriaClinica.FechaPrestamoRequerida,103) as FechaPrestamoRequerida,"
        sSQL = sSQL + " convert(Char(10), PrestamosHistoriaClinica.FechaDevolucion, 103) As FechaDevolucion,"
        sSQL = sSQL + " Servicios.Nombre, PrestamosHistoriaClinica.Observacion, PrestamosHistoriaClinica.NroFolios"
        sSQL = sSQL + " from (((PrestamosHistoriaClinica left join  Pacientes"
        sSQL = sSQL + " on PrestamosHistoriaClinica.IdPaciente = Pacientes.IdPaciente)"
        sSQL = sSQL + " left join Servicios on PrestamosHistoriaClinica.IdServicio = Servicios.IdServicio)"
        sSQL = sSQL + " left join HistoriasClinicas on HistoriasClinicas.IdPaciente = Pacientes.IdPaciente)"

        If oDOPaciente.ApellidoPaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoPaterno like '" + oDOPaciente.ApellidoPaterno + "%' and "
        End If
        If oDOPaciente.ApellidoMaterno <> "" Then
             sWhere = sWhere + " Pacientes.ApellidoMaterno like '" + oDOPaciente.ApellidoMaterno + "%' and "
        End If
        If oDOPaciente.PrimerNombre <> "" Then
             sWhere = sWhere + " Pacientes.PrimerNombre like '" + oDOPaciente.PrimerNombre + "%' and "
        End If
        If oDOPaciente.NroHistoriaClinica <> 0 Then
             sWhere = sWhere + " Pacientes.NroHistoriaClinica ='" & oDOPaciente.NroHistoriaClinica & "' and "
        End If
        If oDOPrestamo.FechaDevolucion <> 0 Then
             sWhere = sWhere + " PrestamosHistoriaClinica.FechaDevolucion =" & oDOPrestamo.FechaDevolucion & " and "
        End If
        If oDOPrestamo.FechaPrestamoRequerida <> 0 Then
             sWhere = sWhere + " PrestamosHistoriaClinica.FechaPrestamoRequerida =" & oDOPrestamo.FechaPrestamoRequerida & " and "
        End If
        If oDOPrestamo.FechaSolicitud <> 0 Then
             sWhere = sWhere + " PrestamosHistoriaClinica.FechaSolicitud =" & oDOPrestamo.FechaSolicitud & " and "
        End If
        If oDOPrestamo.IdEnvio <> 0 Then
             sWhere = sWhere + " PrestamosHistoriaClinica.IdEnvio =" & oDOPrestamo.IdEnvio & " and "
        End If
        
        If sWhere <> "" Then
             sSQL = sSQL + " where " + Left(sWhere, Len(sWhere) - 4)
        End If
       
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
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Function FiltrarEnviados(ByVal oDOPaciente As doPaciente, ByVal oDOPrestamo As DOPrestamoHistoriaClinica) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String
Dim sWhere As String

     Set FiltrarEnviados = Nothing
     
        sWhere = ""
        sSQL = "select PrestamosHistoriaClinica.IdPrestamo, HistoriasClinicas.IdHistoriaClinica as HistoriaClinica, "
        sSQL = sSQL + " Pacientes.ApellidoPaterno + ' ' + Pacientes.ApellidoMaterno + ' ' + Pacientes.PrimerNombre as Nombres, "
        sSQL = sSQL + " convert(char(10), PrestamosHistoriaClinica.FechaSolicitud,103) as FechaSolicitud, "
        sSQL = sSQL + " convert(char(10), PrestamosHistoriaClinica.FechaPrestamoRequerida,103) as FechaPrestamo, "
        sSQL = sSQL + " convert(Char(10), PrestamosHistoriaClinica.FechaDevolucion, 103) As FechaDevolucion, PrestamosHistoriaClinica.IdEnvio "
        sSQL = sSQL + " from (PrestamosHistoriaClinica left join  Pacientes "
        sSQL = sSQL + " on PrestamosHistoriaClinica.IdPaciente = Pacientes.IdPaciente)"
        sSQL = sSQL + " left join HistoriasClinicas "
        sSQL = sSQL + " on HistoriasClinicas.IdPaciente = Pacientes.IdPaciente where IdEnvio is not null "

        If oDOPaciente.ApellidoPaterno <> "" Then
            sWhere = sWhere + " Pacientes.ApellidoPaterno like '" + oDOPaciente.ApellidoPaterno + "%' and "
        End If
        If oDOPaciente.ApellidoMaterno <> "" Then
             sWhere = sWhere + " Pacientes.ApellidoMaterno like '" + oDOPaciente.ApellidoMaterno + "%' and "
        End If
        If oDOPaciente.PrimerNombre <> "" Then
             sWhere = sWhere + " Pacientes.PrimerNombre like '" + oDOPaciente.PrimerNombre + "%' and "
        End If
        If oDOPaciente.NroHistoriaClinica <> 0 Then
             sWhere = sWhere + " Pacientes.NroHistoriaClinica ='" & oDOPaciente.NroHistoriaClinica & "' and "
        End If
        If oDOPrestamo.FechaDevolucion <> 0 Then
             sWhere = sWhere + " PrestamosHistoriaClinica.FechaDevolucion =" & oDOPrestamo.FechaDevolucion & " and "
        End If
        If oDOPrestamo.FechaPrestamoReal <> 0 Then
             sWhere = sWhere + " PrestamosHistoriaClinica.FechaPrestamoReal =" & oDOPrestamo.FechaPrestamoReal & " and "
        End If
        If oDOPrestamo.FechaSolicitud <> 0 Then
             sWhere = sWhere + " PrestamosHistoriaClinica.FechaSolicitud =" & oDOPrestamo.FechaSolicitud & " and "
        End If
        
        If sWhere <> "" Then
             sSQL = sSQL + " and " + Left(sWhere, Len(sWhere) - 4)
        End If
       
        With oCommand
           .CommandType = adCmdText
           Set .ActiveConnection = mo_Conexion
           .CommandText = sSQL
           Set oRecordset = .Execute
           Set oRecordset.ActiveConnection = Nothing
        End With
    
        Set FiltrarEnviados = oRecordset
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


Function FiltrarParaEnvio(oBusqueda As SIGHComun.sghBusquedaPrestamoHistorias) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String
Dim sWhere As String

     Set FiltrarParaEnvio = Nothing
     
        sWhere = ""
        sSQL = "select PrestamosHistoriaClinica.IdPrestamo, HistoriasClinicas.IdHistoriaClinica as HistoriaClinica, "
        sSQL = sSQL + " Pacientes.ApellidoPaterno + ' ' + Pacientes.ApellidoMaterno + ' ' + Pacientes.PrimerNombre as Nombres,"
        sSQL = sSQL + " convert(Char(10), PrestamosHistoriaClinica.FechaPrestamoRequerida, 103) As FechaPrestamoRequerida, "
        sSQL = sSQL + " Servicios.Nombre as Servicio, PrestamosHistoriaClinica.NroFolios "
        sSQL = sSQL + " From"
        sSQL = sSQL + " (((PrestamosHistoriaClinica left join  Pacientes"
        sSQL = sSQL + " on PrestamosHistoriaClinica.IdPaciente = Pacientes.IdPaciente) "
        sSQL = sSQL + " left join HistoriasClinicas on HistoriasClinicas.IdPaciente = Pacientes.IdPaciente)"
        sSQL = sSQL + " left join Servicios on PrestamosHistoriaClinica.IdServicio = Servicios.IdServicio)"

        If oBusqueda.IdEstadoPrestamo = 1 Then
            sWhere = sWhere + " PrestamosHistoriaClinica.IdEstadoPrestamo =" & oBusqueda.IdEstadoPrestamo & " and "
        End If

        If oBusqueda.IdEnvio <> 0 Then
            sWhere = sWhere + " PrestamosHistoriaClinica.IdEnvio =" & oBusqueda.IdEnvio & " and "
        End If

        If oBusqueda.IdTipoServicio <> 0 Then
             sWhere = sWhere + " Servicios.IdTipoServicio = " & oBusqueda.IdTipoServicio & " and "
        End If
        If oBusqueda.IdServicio <> 0 Then
             sWhere = sWhere + " PrestamosHistoriaClinica.IdServicio =" & oBusqueda.IdServicio & " and "
        End If
        
        Select Case oBusqueda.ComparadorFecha
        Case "=", ">", "<="
            sWhere = sWhere + " PrestamosHistoriaClinica.FechaPrestamoRequerida " & oBusqueda.ComparadorFecha & " " & oBusqueda.FechaPrestamoRequeridaDesde & " and "
        Case "Entre"
            sWhere = sWhere + " (PrestamosHistoriaClinica.FechaPrestamoRequerida between " & oBusqueda.FechaPrestamoRequeridaDesde & " and " & oBusqueda.FechaPrestamoRequeridaHasta & ") and "
        End Select
        
        If sWhere <> "" Then
             sSQL = sSQL + " where " + Left(sWhere, Len(sWhere) - 4)
        End If
       
        sSQL = sSQL + " order by PrestamosHistoriaClinica.FechaPrestamoRequerida desc"
       
        With oCommand
           .CommandType = adCmdText
           Set .ActiveConnection = mo_Conexion
           .CommandText = sSQL
           oRecordset.CursorType = adOpenKeyset
           oRecordset.LockType = adLockOptimistic
           Set oRecordset = .Execute
           Set oRecordset.ActiveConnection = Nothing
        End With

        Set FiltrarParaEnvio = oRecordset
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


Function ActualizarIdEnvio(oDOEnvio As DOEnvioHistoriaClinica, oPrestamos As Collection) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSQL As String
 
   ActualizarIdEnvio = False
   
   'Resetea los registros que tengan el Id Envio
   sSQL = "update PrestamosHistoriaClinica set IdEnvio = null , IdEstadoPrestamo = 1, NroFolios = 0, FechaPrestamoReal=null, HoraPrestamoReal=null where IdEnvio = " & oDOEnvio.IdEnvio
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSQL
       .Execute
   End With
   
   
   Dim oPrestamo As DOPrestamoHistoriaClinica
   
   If Not oPrestamos Is Nothing Then
   For Each oPrestamo In oPrestamos
        'Atualiza la nueva lista con el Id de envio y el nro de folio
        sSQL = "update PrestamosHistoriaClinica " + _
        " set IdEnvio = " & oDOEnvio.IdEnvio & _
        ", FechaPrestamoReal  = '" & Format(oDOEnvio.FechaPrestamoReal, "dd/mm/yyyy") & "'" + _
        ", HoraPrestamoReal = '" & Format(oDOEnvio.HoraPrestamoReal, "hh:mm") & "'" + _
        ", NroFolios = " & oPrestamo.NroFolios & _
        ", IdEstadoPrestamo = 2" + _
        " where IdPrestamo = " & oPrestamo.IdPrestamo
        With oCommand
            .CommandType = adCmdText
            Set .ActiveConnection = mo_Conexion
            .CommandText = sSQL
            .Execute
        End With
    Next
    End If
   ActualizarIdEnvio = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

