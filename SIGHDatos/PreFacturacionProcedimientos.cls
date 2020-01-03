VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionProcedimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: AtencionProcedimientos
'        Autor: William Castro Grijalva
'        Fecha: 17/07/2005 05:19:38 p.m.
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
'   Descripción:    Inserta un registro a la tabla AtencionProcedimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOAtencionProcedimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionProcedimientosAgregar"
           Set oParameter = .CreateParameter("@IdServicioOrdena", adInteger, adParamInput, 0, IIf(oTabla.IdServicioOrdena = 0, Null, oTabla.IdServicioOrdena)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoOrdena", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoOrdena = 0, Null, oTabla.IdMedicoOrdena)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraOrden", adChar, adParamInput, 5, IIf(oTabla.HoraOrden = "", Null, oTabla.HoraOrden)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaOrden", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaOrden = 0, Null, oTabla.FechaOrden)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroOrden", adChar, adParamInput, 10, IIf(oTabla.NroOrden = "", Null, oTabla.NroOrden)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionProcedimiento", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdAtencionProcedimiento = .Parameters("@IdAtencionProcedimiento")
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
'   Descripción:    modifica un registro a la tabla AtencionProcedimientos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencionProcedimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionProcedimientosModificar"
           Set oParameter = .CreateParameter("@IdServicioOrdena", adInteger, adParamInput, 0, IIf(oTabla.IdServicioOrdena = 0, Null, oTabla.IdServicioOrdena)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoOrdena", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoOrdena = 0, Null, oTabla.IdMedicoOrdena)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraOrden", adChar, adParamInput, 5, IIf(oTabla.HoraOrden = "", Null, oTabla.HoraOrden)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaOrden", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaOrden = 0, Null, oTabla.FechaOrden)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroOrden", adChar, adParamInput, 10, IIf(oTabla.NroOrden = "", Null, oTabla.NroOrden)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionProcedimiento", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionProcedimiento = 0, Null, oTabla.IdAtencionProcedimiento)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla AtencionProcedimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencionProcedimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionProcedimientosEliminar"
           Set oParameter = .CreateParameter("@IdAtencionProcedimiento", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionProcedimiento = 0, Null, oTabla.IdAtencionProcedimiento)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla AtencionProcedimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtencionProcedimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionProcedimientosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdAtencionProcedimiento", adInteger, adParamInput, 0, oTabla.IdAtencionProcedimiento): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdServicioOrdena = IIf(IsNull(oRecordset!IdServicioOrdena), 0, oRecordset!IdServicioOrdena)
       oTabla.IdMedicoOrdena = IIf(IsNull(oRecordset!IdMedicoOrdena), 0, oRecordset!IdMedicoOrdena)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.HoraOrden = IIf(IsNull(oRecordset!HoraOrden), "", oRecordset!HoraOrden)
       oTabla.FechaOrden = IIf(IsNull(oRecordset!FechaOrden), 0, oRecordset!FechaOrden)
       oTabla.NroOrden = IIf(IsNull(oRecordset!NroOrden), "", oRecordset!NroOrden)
       oTabla.IdAtencionProcedimiento = IIf(IsNull(oRecordset!IdAtencionProcedimiento), 0, oRecordset!IdAtencionProcedimiento)
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




Function Filtrar(ByVal oTabla As DOAtencionProcedimiento, oDOPaciente As doPaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String
Dim sWhere As String

     Set Filtrar = Nothing
     
        sWhere = ""
        
        sSQL = "select AtencionProcedimientos.IdCuentaAtencion, AtencionProcedimientos.NroOrden," + _
        "AtencionProcedimientos.FechaOrden, AtencionProcedimientos.HoraOrden, " + _
        "Pacientes.NroHistoriaClinica , TiposNumeracionHistoria.Descripcion as TipoHistoria, AtencionProcedimientos.IdAtencionProcedimiento " + _
        "from ((AtencionProcedimientos " + _
        "left join CuentasAtencion on AtencionProcedimientos.IdCuentaAtencion = CuentasAtencion.IdCuentaAtencion) " + _
        "left join Pacientes on CuentasAtencion.IdPaciente = Pacientes.IdPaciente)" + _
        "left join TiposNumeracionHistoria on Pacientes.IdTipoNumeracion = TiposNumeracionHistoria.IdTipoNumeracion"

       If oDOPaciente.NroHistoriaClinica <> 0 Then
            sWhere = sWhere + " Pacientes.NroHistoriaClinica = " & oDOPaciente.NroHistoriaClinica & " and "
       End If
       If oTabla.IdCuentaAtencion <> 0 Then
            sWhere = sWhere + " AtencionProcedimientos.IdCuentaAtencion = " & oTabla.IdCuentaAtencion & " and "
       End If
       If oTabla.NroOrden <> "" Then
            sWhere = sWhere + " AtencionProcedimientos.NroOrden = '" + oTabla.NroOrden + "' and "
       End If

       If sWhere <> "" Then
            sSQL = sSQL + " where " + Left(sWhere, Len(sWhere) - 4)
       End If

        sSQL = sSQL + " order by Pacientes.NroHistoriaClinica, AtencionProcedimientos.IdCuentaAtencion "
        
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

Function SeleccionarPorCuentaAtencionSinIterconsultas(lIdCuentaAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorCuentaAtencionSinIterconsultas = Nothing
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtencionesProcedimientosSeleccionarPorCuentaAtencionSinInterconsulta"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 4, lIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorCuentaAtencionSinIterconsultas = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Function SeleccionarPorCuentaAtencionEIterconsulta(lIdCuentaAtencion As Long, lIdInterconsulta As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorCuentaAtencionEIterconsulta = Nothing
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtencionesProcedimientosSeleccionarPorCuentaAtencionEInterconsulta"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 4, lIdCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdInterconsulta", adInteger, adParamInput, 4, lIdInterconsulta): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorCuentaAtencionEIterconsulta = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function ActualizarProcedimientosAtencion(oProcedimientos As Collection, lIdCuentaAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSQL As String
Dim oDOAtencionProcedimiento As New DOAtencionProcedimiento
 
   ActualizarProcedimientosAtencion = False
   
   'elimina los complicaciones
   sSQL = "delete from AtencionesProcedimientos where IdCuentaAtencion = " & lIdCuentaAtencion
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSQL
       .Execute
   End With
   
   If Not oProcedimientos Is Nothing Then
        For Each oDOAtencionProcedimiento In oProcedimientos
            oDOAtencionProcedimiento.IdCuentaAtencion = lIdCuentaAtencion
            If Not Insertar(oDOAtencionProcedimiento) Then
                Exit Function
            End If
         Next
    End If
    
   ActualizarProcedimientosAtencion = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Function ActualizarProcedimientosInterconsultas(oProcedimientos As Collection, lIdCuentaAtencion As Long, lIdInterconsulta As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSQL As String
Dim oDOAtencionProcedimiento As New DOAtencionProcedimiento
 
   ActualizarProcedimientosInterconsultas = False
   
'   ''elimina solo los proc de interconsultas
'   sSQL = "delete from AtencionesProcedimientos where IdCuentaAtencion = " & lIdCuentaAtencion & " and IdInterconsulta = " & lIdInterconsulta
'   With oCommand
'       .CommandType = adCmdText
'       Set .ActiveConnection = mo_Conexion
'       .CommandText = sSQL
'       .Execute
'   End With
'
'   If Not oProcedimientos Is Nothing Then
'        For Each oDOAtencionProcedimiento In oProcedimientos
'            oDOAtencionProcedimiento.IdCuentaAtencion = lIdCuentaAtencion
'            oDOAtencionProcedimiento.IdInterconsulta = lIdInterconsulta
'            If Not Insertar(oDOAtencionProcedimiento) Then
'                Exit Function
'            End If
'         Next
'    End If
    
   ActualizarProcedimientosInterconsultas = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function EliminarProcedimientosPorCuentaAtencion(lIdCuentaAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSQL As String
 
   EliminarProcedimientosPorCuentaAtencion = False
   
   'elimina los diagnosticos
   sSQL = "delete from AtencionesProcedimientos where IdCuentaAtencion = " & lIdCuentaAtencion
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSQL
       .Execute
   End With
   
   EliminarProcedimientosPorCuentaAtencion = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function SeleccionarProcedimientos(lIdCuentaAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarProcedimientos = Nothing
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtencionesProcedimientosSeleccionarPorCuentaAtencion"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 4, lIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarProcedimientos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
