VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CitasBloqueadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla CitasBloqueadas
'        Programado por: Castro W
'        Fecha: Agosto 2006
'
'------------------------------------------------------------------------------------
Option Explicit
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
'   Descripción:    Inserta un registro a la tabla CitasBloqueadas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCitaBloqueada) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CitasBloqueadasAgregar"
           Set oParameter = .CreateParameter("@IdCitaBloqueada", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraBloqueo", adChar, adParamInput, 5, IIf(oTabla.HoraBloqueo = "", Null, oTabla.HoraBloqueo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaBloqueo", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaBloqueo = 0, Null, oTabla.FechaBloqueo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IIf(oTabla.IdMedico = 0, Null, oTabla.IdMedico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraFin", adChar, adParamInput, 5, IIf(oTabla.HoraFin = "", Null, oTabla.HoraFin)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraInicio", adChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fecha = 0, Null, oTabla.fecha)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.idCitaBloqueada = .Parameters("@IdCitaBloqueada")
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
'   Descripción:    modifica un registro a la tabla CitasBloqueadas
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCitaBloqueada) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CitasBloqueadasModificar"
           Set oParameter = .CreateParameter("@IdCitaBloqueada", adInteger, adParamInput, 0, IIf(oTabla.idCitaBloqueada = 0, Null, oTabla.idCitaBloqueada)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraBloqueo", adChar, adParamInput, 5, IIf(oTabla.HoraBloqueo = "", Null, oTabla.HoraBloqueo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaBloqueo", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaBloqueo = 0, Null, oTabla.FechaBloqueo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IIf(oTabla.IdMedico = 0, Null, oTabla.IdMedico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraFin", adChar, adParamInput, 5, IIf(oTabla.HoraFin = "", Null, oTabla.HoraFin)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraInicio", adChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fecha = 0, Null, oTabla.fecha)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CitasBloqueadas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCitaBloqueada) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CitasBloqueadasEliminar"
           Set oParameter = .CreateParameter("@IdCitaBloqueada", adInteger, adParamInput, 0, IIf(oTabla.idCitaBloqueada = 0, Null, oTabla.idCitaBloqueada)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla CitasBloqueadas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCitaBloqueada) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CitasBloqueadasSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdCitaBloqueada", adInteger, adParamInput, 0, oTabla.idCitaBloqueada): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.idCitaBloqueada = IIf(IsNull(oRecordset!idCitaBloqueada), 0, oRecordset!idCitaBloqueada)
       oTabla.HoraBloqueo = IIf(IsNull(oRecordset!HoraBloqueo), "", oRecordset!HoraBloqueo)
       oTabla.FechaBloqueo = IIf(IsNull(oRecordset!FechaBloqueo), 0, oRecordset!FechaBloqueo)
       oTabla.IdMedico = IIf(IsNull(oRecordset!IdMedico), 0, oRecordset!IdMedico)
       oTabla.HoraFin = IIf(IsNull(oRecordset!HoraFin), "", oRecordset!HoraFin)
       oTabla.HoraInicio = IIf(IsNull(oRecordset!HoraInicio), "", oRecordset!HoraInicio)
       oTabla.fecha = IIf(IsNull(oRecordset!fecha), 0, oRecordset!fecha)
       oTabla.IdUsuario = IIf(IsNull(oRecordset!IdUsuario), 0, oRecordset!IdUsuario)
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


Function EliminarPorUsuario(lIdUsuario As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CitasBloqueadasEliminarXusuario"
        Set oParameter = .CreateParameter("@lIdUsuario", adInteger, adParamInput, 0, lIdUsuario): .Parameters.Append oParameter
        .Execute
   End With
   Set oRecordset = Nothing
   Set oCommand = Nothing
   EliminarPorUsuario = True
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'MARIO 20/08/2014
Function CitasBloqueadasEliminarPorFechaHora(ldFecha As Date, lcHoraInicio As String, ldIdMedico As Long, oConexion As Connection) As Boolean
        On Error GoTo ErrCWEB
        CitasBloqueadasEliminarPorFechaHora = False
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CitasBloqueadasEliminarPorFechayHora"
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, Format(ldFecha, "dd/mm/yyyy")): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@HoraInicio", adVarChar, adParamInput, 5, lcHoraInicio): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 4, ldIdMedico): .Parameters.Append oParameter
            .Execute
        End With
        Set oCommand = Nothing
        CitasBloqueadasEliminarPorFechaHora = True
        Exit Function
ErrCWEB:
       ms_MensajeError = Err.Description
End Function
