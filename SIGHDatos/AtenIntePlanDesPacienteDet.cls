VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtenIntePlanDesPacienteDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtenIntePlanDesPacienteDet
'        Programado por: Garay M
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla AtenIntePlanDesPacienteDet
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtenIntePlanDesPacienteDet) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanDesPacienteDetAgregar"
   Set oParameter = .CreateParameter("@IdPlanDesarrolloPaciente", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanIntegralPaciente", adBigInt, adParamInput, 0, oTabla.IdPlanIntegralPaciente)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdItemDesarrollo", adInteger, adParamInput, 0, IIf(oTabla.IdItemDesarrollo = 0, Null, oTabla.IdItemDesarrollo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OrdenItem", adInteger, adParamInput, 0, IIf(oTabla.OrdenItem = 0, Null, oTabla.OrdenItem))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EjecutaAccion", adBoolean, adParamInput, 0, IIf(oTabla.EjecutaAccion = 0, Null, oTabla.EjecutaAccion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdPlanDesarrolloPaciente = .Parameters("@IdPlanDesarrolloPaciente")
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
'   Descripción:    modifica un registro a la tabla AtenIntePlanDesPacienteDet
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtenIntePlanDesPacienteDet) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanDesPacienteDetModificar"
   Set oParameter = .CreateParameter("@IdPlanDesarrolloPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPlanDesarrolloPaciente = 0, Null, oTabla.IdPlanDesarrolloPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanIntegralPaciente", adBigInt, adParamInput, 0, oTabla.IdPlanIntegralPaciente)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdItemDesarrollo", adInteger, adParamInput, 0, IIf(oTabla.IdItemDesarrollo = 0, Null, oTabla.IdItemDesarrollo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OrdenItem", adInteger, adParamInput, 0, IIf(oTabla.OrdenItem = 0, Null, oTabla.OrdenItem))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EjecutaAccion", adBoolean, adParamInput, 0, IIf(oTabla.RespondioEjecutaAccion = False, Null, oTabla.EjecutaAccion))
   .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla AtenIntePlanDesPacienteDet
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtenIntePlanDesPacienteDet) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanDesPacienteDetEliminar"
   Set oParameter = .CreateParameter("@IdPlanDesarrolloPaciente", adInteger, adParamInput, 0, oTabla.IdPlanDesarrolloPaciente)
   .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla AtenIntePlanDesPacienteDet
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtenIntePlanDesPacienteDet) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtenIntePlanDesPacienteDetSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdPlanDesarrolloPaciente", adInteger, adParamInput, 0, oTabla.IdPlanDesarrolloPaciente)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdPlanDesarrolloPaciente = IIf(IsNull(oRecordset!IdPlanDesarrolloPaciente), 0, oRecordset!IdPlanDesarrolloPaciente)
       oTabla.IdPlanIntegralPaciente = oRecordset!IdPlanIntegralPaciente
       oTabla.IdItemDesarrollo = IIf(IsNull(oRecordset!IdItemDesarrollo), 0, oRecordset!IdItemDesarrollo)
       oTabla.OrdenItem = IIf(IsNull(oRecordset!OrdenItem), 0, oRecordset!OrdenItem)
       oTabla.EjecutaAccion = IIf(IsNull(oRecordset!EjecutaAccion), 0, oRecordset!EjecutaAccion)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub SetDefaults(ByVal oTabla As DOAtenIntePlanDesPacienteDet)
   oTabla.IdPlanDesarrolloPaciente = 0
   oTabla.IdPlanIntegralPaciente = 0
   oTabla.IdItemDesarrollo = 0
   oTabla.OrdenItem = 0
   oTabla.EjecutaAccion = False
End Sub

'mgaray20141012
Public Function ListarPlanDesarrolloPacienteDetallePorId(ByVal oTabla As DOAtenIntePlanDesPacienteDet) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim sSql As String
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtenInteListarDesarrolloPacienteDetPorId"
        Set oParameter = .CreateParameter("@IdPlanIntegralPaciente", adInteger, adParamInput, 0, oTabla.IdPlanIntegralPaciente)
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPlanDesarrolloPaciente", adInteger, adParamInput, 0, oTabla.IdPlanDesarrolloPaciente)
        .Parameters.Append oParameter
        
'        Set oRecordset = .Execute
        oRecordset.Open oCommand, , adOpenStatic, adLockOptimistic
        Set oRecordset.ActiveConnection = Nothing
    End With
    
    Set ListarPlanDesarrolloPacienteDetallePorId = oRecordset
    ms_MensajeError = ""
    Exit Function
ManejadorDeError:
    If Err Then
        ms_MensajeError = Err.Number & " " + Err.Description
    End If
End Function

