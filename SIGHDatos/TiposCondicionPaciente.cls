VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TiposCondicionPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla TiposCondicionPaciente
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
'   Descripción:    Inserta un registro a la tabla TiposCondicionPaciente
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOTipoCondicionPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "TiposCondicionPacienteAgregar"
        Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 50, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoCondicionPaciente", adInteger, adParamOutput, 0): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdTipoCondicionPaciente = .Parameters("@IdTipoCondicionPaciente")
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
'   Descripción:    modifica un registro a la tabla TiposCondicionPaciente
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOTipoCondicionPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "TiposCondicionPacienteModificar"
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 50, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoCondicionPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdTipoCondicionPaciente = 0, Null, oTabla.IdTipoCondicionPaciente)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla TiposCondicionPaciente
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOTipoCondicionPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "TiposCondicionPacienteEliminar"
           Set oParameter = .CreateParameter("@IdTipoCondicionPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdTipoCondicionPaciente = 0, Null, oTabla.IdTipoCondicionPaciente)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla TiposCondicionPaciente
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOTipoCondicionPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TiposCondicionPacienteSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdTipoCondicionPaciente", adInteger, adParamInput, 0, oTabla.IdTipoCondicionPaciente): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
       oTabla.IdTipoCondicionPaciente = IIf(IsNull(oRecordset!IdTipoCondicionPaciente), 0, oRecordset!IdTipoCondicionPaciente)
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
     .CommandText = "TiposCondicionPacienteSeleccionarTodos"
     Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
    Set oCommand = Nothing
   Set oRecordset = Nothing

Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function CondicionAlEstablecimiento(IdPaciente As Long, FechaIngreso As Date, idAtencion As Long) As Integer
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   CondicionAlEstablecimiento = False
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "CondicionPacienteAlEstablecimiento"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaIngreso", adDate, adParamInput, 4, FechaIngreso): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoCondicionPaciente", adInteger, adParamOutput, 0): .Parameters.Append oParameter
        .Execute
        CondicionAlEstablecimiento = .Parameters("@IdTipoCondicionPaciente")
   End With
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Function CondicionAlServicio(IdPaciente As Long, FechaIngreso As Date, IdServicio As Long, idAtencion As Long) As Integer
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   CondicionAlServicio = False
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "CondicionPacienteAlServicio"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaIngreso", adDate, adParamInput, 4, FechaIngreso): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IdServicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoCondicionPaciente", adInteger, adParamOutput, 0): .Parameters.Append oParameter
        .Execute
        CondicionAlServicio = .Parameters("@IdTipoCondicionPaciente")
   End With
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

