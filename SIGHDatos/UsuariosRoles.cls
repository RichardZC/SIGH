VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UsuariosRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla UsuariosRoles
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
'   Descripción:    Inserta un registro a la tabla UsuariosRoles
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOUsuarioRol) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "UsuariosRolesAgregar"
           Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleado = 0, Null, oTabla.IdEmpleado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdRol", adInteger, adParamInput, 0, IIf(oTabla.IdRol = 0, Null, oTabla.IdRol)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioRol", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdUsuarioRol = .Parameters("@IdUsuarioRol")
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
'   Descripción:    modifica un registro a la tabla UsuariosRoles
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOUsuarioRol) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "UsuariosRolesModificar"
           Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleado = 0, Null, oTabla.IdEmpleado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdRol", adInteger, adParamInput, 0, IIf(oTabla.IdRol = 0, Null, oTabla.IdRol)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioRol", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioRol = 0, Null, oTabla.IdUsuarioRol)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla UsuariosRoles
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOUsuarioRol) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "UsuariosRolesEliminar"
           Set oParameter = .CreateParameter("@IdUsuarioRol", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioRol = 0, Null, oTabla.IdUsuarioRol)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla UsuariosRoles
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOUsuarioRol) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "UsuariosRolesSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdUsuarioRol", adInteger, adParamInput, 0, oTabla.IdUsuarioRol): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdEmpleado = IIf(IsNull(oRecordset!IdEmpleado), 0, oRecordset!IdEmpleado)
       oTabla.IdRol = IIf(IsNull(oRecordset!IdRol), 0, oRecordset!IdRol)
       oTabla.IdUsuarioRol = IIf(IsNull(oRecordset!IdUsuarioRol), 0, oRecordset!IdUsuarioRol)
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

Function ActualizarPorEmpleado(oUsuariosRoles As Collection, lIdEmpleado As Long) As Boolean
Dim oDOUsuarioRol As New DOUsuarioRol
On Error GoTo ManejadorDeError
   EliminarPorEmpleado lIdEmpleado
   If Not oUsuariosRoles Is Nothing Then
        For Each oDOUsuarioRol In oUsuariosRoles
            oDOUsuarioRol.IdEmpleado = lIdEmpleado
            If Not Insertar(oDOUsuarioRol) Then
                Exit Function
            End If
         Next
    End If

   ActualizarPorEmpleado = True
   ms_MensajeError = ""
   
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function EliminarPorEmpleado(lIdEmpleado As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "UsuariosRolesElimnaXidEmpleado"
        Set oParameter = .CreateParameter("@lIdEmpleado", adInteger, adParamInput, 0, lIdEmpleado): .Parameters.Append oParameter
         .Execute
   End With
   ms_MensajeError = ""
   EliminarPorEmpleado = True
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarPorEmpleado(IdEmpleado As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorEmpleado = Nothing
   With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_Conexion
    .CommandText = "UsuariosRolesSeleccionarPorEmpleados"
    Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, IdEmpleado)
    .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarPorEmpleado = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
