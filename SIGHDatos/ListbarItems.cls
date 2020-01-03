VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListbarItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla ListBarItems
'        Programado por: Castro W
'        Fecha: Mayo 2006
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
'   Descripción:    Inserta un registro a la tabla ListbarItems
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOListbarItem) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ListbarItemsAgregar"
           Set oParameter = .CreateParameter("@KeyIcon", adChar, adParamInput, 50, IIf(oTabla.KeyIcon = "", Null, oTabla.KeyIcon)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Indice", adInteger, adParamInput, 0, IIf(oTabla.Indice = 0, Null, oTabla.Indice)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Clave", adVarChar, adParamInput, 20, IIf(oTabla.Clave = "", Null, oTabla.Clave)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Texto", adVarChar, adParamInput, 20, IIf(oTabla.Texto = "", Null, oTabla.Texto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdListGrupo", adInteger, adParamInput, 0, IIf(oTabla.IdListGrupo = 0, Null, oTabla.IdListGrupo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdListItem", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdListItem = .Parameters("@IdListItem")
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
'   Descripción:    modifica un registro a la tabla ListbarItems
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOListbarItem) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ListbarItemsModificar"
           Set oParameter = .CreateParameter("@KeyIcon", adChar, adParamInput, 50, IIf(oTabla.KeyIcon = "", Null, oTabla.KeyIcon)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Indice", adInteger, adParamInput, 0, IIf(oTabla.Indice = 0, Null, oTabla.Indice)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Clave", adVarChar, adParamInput, 20, IIf(oTabla.Clave = "", Null, oTabla.Clave)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Texto", adVarChar, adParamInput, 20, IIf(oTabla.Texto = "", Null, oTabla.Texto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdListGrupo", adInteger, adParamInput, 0, IIf(oTabla.IdListGrupo = 0, Null, oTabla.IdListGrupo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdListItem", adInteger, adParamInput, 0, IIf(oTabla.IdListItem = 0, Null, oTabla.IdListItem)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla ListbarItems
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOListbarItem) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ListbarItemsEliminar"
           Set oParameter = .CreateParameter("@IdListItem", adInteger, adParamInput, 0, IIf(oTabla.IdListItem = 0, Null, oTabla.IdListItem)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla ListbarItems
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOListbarItem) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ListbarItemsSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdListItem", adInteger, adParamInput, 0, oTabla.IdListItem): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.KeyIcon = IIf(IsNull(oRecordset!KeyIcon), "", oRecordset!KeyIcon)
       oTabla.Indice = IIf(IsNull(oRecordset!Indice), 0, oRecordset!Indice)
       oTabla.Clave = IIf(IsNull(oRecordset!Clave), "", oRecordset!Clave)
       oTabla.Texto = IIf(IsNull(oRecordset!Texto), "", oRecordset!Texto)
       oTabla.IdListGrupo = IIf(IsNull(oRecordset!IdListGrupo), 0, oRecordset!IdListGrupo)
       oTabla.IdListItem = IIf(IsNull(oRecordset!IdListItem), 0, oRecordset!IdListItem)
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
     .CommandText = "ListbarItemsSeleccionarTodos"
     Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

