VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TurnoDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla Turnos
'        Programado por: Castro W
'        Fecha: Marzo 2006
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
'   Descripción:    Inserta un registro a la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As turnoComun) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "TurnosAgregar"
           Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidad = 0, Null, oTabla.IdEspecialidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adChar, adParamInput, 3, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, IIf(oTabla.IdTipoServicio = 0, Null, oTabla.IdTipoServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraFin", adChar, adParamInput, 5, IIf(oTabla.HoraFin = "", Null, oTabla.HoraFin)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraInicio", adChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 50, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.idTurno = .Parameters("@IdTurno")
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
'   Descripción:    modifica un registro a la tabla Turnos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As turnoComun) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "TurnosModificar"
           Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidad = 0, Null, oTabla.IdEspecialidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adChar, adParamInput, 3, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, IIf(oTabla.IdTipoServicio = 0, Null, oTabla.IdTipoServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraFin", adChar, adParamInput, 5, IIf(oTabla.HoraFin = "", Null, oTabla.HoraFin)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraInicio", adChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 50, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.idTurno = 0, Null, oTabla.idTurno)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As turnoComun) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "TurnosEliminar"
           Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.idTurno = 0, Null, oTabla.idTurno)): .Parameters.Append oParameter
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

Sub ActualizaTabla(oTabla As turnoComun, oRecordset As Recordset)
       oTabla.IdEspecialidad = IIf(IsNull(oRecordset!IdEspecialidad), 0, oRecordset!IdEspecialidad)
       oTabla.codigo = IIf(IsNull(oRecordset!codigo), "", oRecordset!codigo)
       oTabla.IdTipoServicio = IIf(IsNull(oRecordset!IdTipoServicio), 0, oRecordset!IdTipoServicio)
       oTabla.HoraFin = IIf(IsNull(oRecordset!HoraFin), "", oRecordset!HoraFin)
       oTabla.HoraInicio = IIf(IsNull(oRecordset!HoraInicio), "", oRecordset!HoraInicio)
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
       oTabla.idTurno = IIf(IsNull(oRecordset!idTurno), 0, oRecordset!idTurno)

End Sub
'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As turnoComun) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TurnosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, oTabla.idTurno): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       ActualizaTabla oTabla, oRecordset
       SeleccionarPorId = True
'       oTabla.IdEspecialidad = IIf(IsNull(oRecordset!IdEspecialidad), 0, oRecordset!IdEspecialidad)
'       oTabla.codigo = IIf(IsNull(oRecordset!codigo), "", oRecordset!codigo)
'       oTabla.IdTipoServicio = IIf(IsNull(oRecordset!IdTipoServicio), 0, oRecordset!IdTipoServicio)
'       oTabla.HoraFin = IIf(IsNull(oRecordset!HoraFin), "", oRecordset!HoraFin)
'       oTabla.HoraInicio = IIf(IsNull(oRecordset!HoraInicio), "", oRecordset!HoraInicio)
'       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
'       oTabla.IdTurno = IIf(IsNull(oRecordset!IdTurno), 0, oRecordset!IdTurno)
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

Function SeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TurnosSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarPorEspecialidad(lIdEspecialidad As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorEspecialidad = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TurnosSeleccionarPorEspecialidad"
     Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, lIdEspecialidad): .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarPorEspecialidad = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarPorCodigo(ByVal oTabla As DoTurno) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "TurnosXcodigo"
        Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 3, oTabla.codigo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If Not (oRecordset.EOF And oRecordset.BOF) Then
        ActualizaTabla oTabla, oRecordset
        SeleccionarPorCodigo = True
   Else
        SeleccionarPorCodigo = False
   End If
   oRecordset.Close
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description

End Function

