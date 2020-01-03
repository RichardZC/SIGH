VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CentroPoblados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla CentrosPoblados
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
'   Descripción:    Inserta un registro a la tabla CentrosPoblados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCentroPoblado) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CentrosPobladosAgregar"
           Set oParameter = .CreateParameter("@IdDistrito", adInteger, adParamInput, 0, oTabla.IdDistrito): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 50, oTabla.nombre): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCentroPoblado", adInteger, adParamInput, 0, oTabla.IdCentroPoblado): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
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
'   Descripción:    modifica un registro a la tabla CentrosPoblados
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCentroPoblado) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CentrosPobladosModificar"
       Set oParameter = .CreateParameter("@IdDistrito", adInteger, adParamInput, 0, oTabla.IdDistrito): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 50, oTabla.nombre): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdCentroPoblado", adInteger, adParamInput, 0, oTabla.IdCentroPoblado): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CentrosPoblados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCentroPoblado) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CentrosPobladosEliminar"
       Set oParameter = .CreateParameter("@IdCentroPoblado", adInteger, adParamInput, 0, oTabla.IdCentroPoblado): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla CentrosPoblados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCentroPoblado) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CentrosPobladosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdCentroPoblado", adInteger, adParamInput, 0, oTabla.IdCentroPoblado): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdDistrito = oRecordset("IdDistrito")
       oTabla.nombre = oRecordset("Nombre")
       oTabla.IdCentroPoblado = oRecordset("IdCentroPoblado")
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


Function SeleccionarPorDistrito(IdDistrito As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorDistrito = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CentrosPobladosSeleccionarPorDistrito"
     Set oParameter = .CreateParameter("@IdDistrito", adInteger, adParamInput, 0, IdDistrito): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorDistrito = oRecordset
 
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
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CentrosPobladosSeleccionarTodos"
     Set oRecordset = .Execute
   End With
   Set SeleccionarTodos = oRecordset
   Set oCommand = Nothing
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


