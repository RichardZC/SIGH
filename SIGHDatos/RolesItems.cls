VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RolesItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla RolesItems
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
'   Descripción:    Inserta un registro a la tabla RolesItems
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DORolItem) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "RolesItemsAgregar"
           Set oParameter = .CreateParameter("@Consultar", adBoolean, adParamInput, 0, IIf(oTabla.Consultar = 0, 0, oTabla.Consultar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Eliminar", adBoolean, adParamInput, 0, IIf(oTabla.Eliminar = 0, 0, oTabla.Eliminar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Modificar", adBoolean, adParamInput, 0, IIf(oTabla.Modificar = 0, 0, oTabla.Modificar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Agregar", adBoolean, adParamInput, 0, IIf(oTabla.Agregar = 0, 0, oTabla.Agregar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdRol", adInteger, adParamInput, 0, IIf(oTabla.IdRol = 0, 0, oTabla.IdRol)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdListItem", adInteger, adParamInput, 0, IIf(oTabla.IdListItem = 0, 0, oTabla.IdListItem)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdRolItem", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdRolItem = .Parameters("@IdRolItem")
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
'   Descripción:    modifica un registro a la tabla RolesItems
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DORolItem) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "RolesItemsModificar"
           Set oParameter = .CreateParameter("@Consultar", adBoolean, adParamInput, 0, IIf(oTabla.Consultar = 0, 0, oTabla.Consultar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Eliminar", adBoolean, adParamInput, 0, IIf(oTabla.Eliminar = 0, 0, oTabla.Eliminar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Modificar", adBoolean, adParamInput, 0, IIf(oTabla.Modificar = 0, 0, oTabla.Modificar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Agregar", adBoolean, adParamInput, 0, IIf(oTabla.Agregar = 0, 0, oTabla.Agregar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdRol", adInteger, adParamInput, 0, IIf(oTabla.IdRol = 0, 0, oTabla.IdRol)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdListItem", adInteger, adParamInput, 0, IIf(oTabla.IdListItem = 0, 0, oTabla.IdListItem)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdRolItem", adInteger, adParamInput, 0, IIf(oTabla.IdRolItem = 0, 0, oTabla.IdRolItem)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla RolesItems
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DORolItem) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "RolesItemsEliminar"
           Set oParameter = .CreateParameter("@IdRolItem", adInteger, adParamInput, 0, IIf(oTabla.IdRolItem = 0, Null, oTabla.IdRolItem)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla RolesItems
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DORolItem) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "RolesItemsSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdRolItem", adInteger, adParamInput, 0, oTabla.IdRolItem): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.Consultar = IIf(IsNull(oRecordset!Consultar), 0, oRecordset!Consultar)
       oTabla.Eliminar = IIf(IsNull(oRecordset!Eliminar), 0, oRecordset!Eliminar)
       oTabla.Modificar = IIf(IsNull(oRecordset!Modificar), 0, oRecordset!Modificar)
       oTabla.Agregar = IIf(IsNull(oRecordset!Agregar), 0, oRecordset!Agregar)
       oTabla.IdRol = IIf(IsNull(oRecordset!IdRol), 0, oRecordset!IdRol)
       oTabla.IdListItem = IIf(IsNull(oRecordset!IdListItem), 0, oRecordset!IdListItem)
       oTabla.IdRolItem = IIf(IsNull(oRecordset!IdRolItem), 0, oRecordset!IdRolItem)
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

Function SeleccionarGruposPorUsuario(lIdUsuario As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "RolesItemsSeleccionarGruposPorUsuario"
        Set oParameter = .CreateParameter("@lIdUsuario", adInteger, adParamInput, 0, lIdUsuario): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarGruposPorUsuario = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarItemsPorUsuarioYGrupo(lIdUsuario As Long, lIdGrupo As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "RolesItemsSeleccionarItemsPorUsuarioYGrupo"
        Set oParameter = .CreateParameter("@lIdGrupo", adInteger, adParamInput, 0, lIdGrupo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdUsuario", adInteger, adParamInput, 0, lIdUsuario): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarItemsPorUsuarioYGrupo = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ActualizarRolesItems(oRolesItems As Collection, lIdRol As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oDORolItem As New DORolItem
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "RolesItemsEliminarXidRol"
        Set oParameter = .CreateParameter("@lIdRol", adInteger, adParamInput, 0, lIdRol): .Parameters.Append oParameter
         .Execute
   End With
   If Not oRolesItems Is Nothing Then
        For Each oDORolItem In oRolesItems
            oDORolItem.IdRol = lIdRol
            If Not Insertar(oDORolItem) Then
                Exit Function
            End If
         Next
    End If

   ActualizarRolesItems = True
   ms_MensajeError = ""
   
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function EliminarRolesItems(lIdRol As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oDORolItem As New DORolItem
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "RolesItemsEliminarXidRol"
        Set oParameter = .CreateParameter("@lIdRol", adInteger, adParamInput, 0, lIdRol): .Parameters.Append oParameter
         .Execute
   End With
   EliminarRolesItems = True
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarPorRol(lIdRol As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorRol = Nothing
   With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_Conexion
    .CommandText = "RolesItemsSeleccionarPorRol"
    Set oParameter = .CreateParameter("@IdRol", adInteger, adParamInput, 0, lIdRol)
    .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarPorRol = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarPermisosPorIdEmpleadoYIdListItem(lIdEmpleado As Long, lIdListItem As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPermisosPorIdEmpleadoYIdListItem = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "RolesItemSeleccionarPermisosPorIdEmpleadoYIdListItem"
        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, lIdEmpleado): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdListItem", adInteger, adParamInput, 0, lIdListItem): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
   
   Set SeleccionarPermisosPorIdEmpleadoYIdListItem = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

