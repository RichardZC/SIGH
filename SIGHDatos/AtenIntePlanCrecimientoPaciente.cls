VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtenIntePlanCrecPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtenIntePlanCrecimientoPaciente
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
'   Descripción:    Inserta un registro a la tabla AtenIntePlanCrecimientoPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtenIntePlanCrecPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanCrecimientoPacienteAgregar"
   Set oParameter = .CreateParameter("@IdPlanCrecimientoPaciente", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanIntegralPaciente", adBigInt, adParamInput, 0, oTabla.IdPlanIntegralPaciente)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdPlanAtencion = 0, Null, oTabla.IdPlanAtencion))
   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@IdTriajeVariable", adUnsignedTinyInt, adParamInput, 0, oTabla.IdTriajeVariable)
'   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAtenInteItemPlan", adInteger, adParamInput, 0, IIf(oTabla.IdAtenInteItemPlan = 0, Null, oTabla.IdAtenInteItemPlan))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaProgramada", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaProgramada = 0, Null, oTabla.FechaProgramada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaEjecucion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEjecucion = 0, Null, oTabla.FechaEjecucion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NumeroSesion", adInteger, adParamInput, 0, IIf(oTabla.NumeroSesion = 0, Null, oTabla.NumeroSesion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdPlanCrecimientoPaciente = .Parameters("@IdPlanCrecimientoPaciente")
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
'   Descripción:    modifica un registro a la tabla AtenIntePlanCrecimientoPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtenIntePlanCrecPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanCrecimientoPacienteModificar"
   Set oParameter = .CreateParameter("@IdPlanCrecimientoPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPlanCrecimientoPaciente = 0, Null, oTabla.IdPlanCrecimientoPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanIntegralPaciente", adBigInt, adParamInput, 0, oTabla.IdPlanIntegralPaciente)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdPlanAtencion = 0, Null, oTabla.IdPlanAtencion))
   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@IdTriajeVariable", adUnsignedTinyInt, adParamInput, 0, oTabla.IdTriajeVariable)
'   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAtenInteItemPlan", adInteger, adParamInput, 0, IIf(oTabla.IdAtenInteItemPlan = 0, Null, oTabla.IdAtenInteItemPlan))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaProgramada", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaProgramada = 0, Null, oTabla.FechaProgramada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaEjecucion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEjecucion = 0, Null, oTabla.FechaEjecucion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NumeroSesion", adInteger, adParamInput, 0, IIf(oTabla.NumeroSesion = 0, Null, oTabla.NumeroSesion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
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
'   Descripción:    elimia o borra un registro a la tabla AtenIntePlanCrecimientoPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtenIntePlanCrecPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanCrecimientoPacienteEliminar"
   Set oParameter = .CreateParameter("@IdPlanCrecimientoPaciente", adInteger, adParamInput, 0, oTabla.IdPlanCrecimientoPaciente)
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
'   Descripción:    Selecciona un unico registro de la tabla AtenIntePlanCrecimientoPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtenIntePlanCrecPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtenIntePlanCrecimientoPacienteSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdPlanCrecimientoPaciente", adInteger, adParamInput, 0, oTabla.IdPlanCrecimientoPaciente)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdPlanCrecimientoPaciente = IIf(IsNull(oRecordset!IdPlanCrecimientoPaciente), 0, oRecordset!IdPlanCrecimientoPaciente)
       oTabla.IdPlanIntegralPaciente = oRecordset!IdPlanIntegralPaciente
       oTabla.IdPlanAtencion = IIf(IsNull(oRecordset!IdPlanAtencion), 0, oRecordset!IdPlanAtencion)
'       oTabla.IdTriajeVariable = oRecordset!IdTriajeVariable
       oTabla.IdAtenInteItemPlan = IIf(IsNull(oRecordset!IdAtenInteItemPlan), 0, oRecordset!IdAtenInteItemPlan)
       oTabla.FechaProgramada = IIf(IsNull(oRecordset!FechaProgramada), 0, oRecordset!FechaProgramada)
       oTabla.FechaEjecucion = IIf(IsNull(oRecordset!FechaEjecucion), 0, oRecordset!FechaEjecucion)
       oTabla.NumeroSesion = IIf(IsNull(oRecordset!NumeroSesion), 0, oRecordset!NumeroSesion)
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       oTabla.IdEstablecimiento = IIf(IsNull(oRecordset!IdEstablecimiento), 0, oRecordset!IdEstablecimiento)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOAtenIntePlanCrecPaciente)
   oTabla.IdPlanCrecimientoPaciente = 0
   oTabla.IdPlanIntegralPaciente = 0
   oTabla.IdPlanAtencion = 0
'   oTabla.IdTriajeVariable = 0
   oTabla.IdAtenInteItemPlan = 0
   oTabla.FechaProgramada = Now
   oTabla.FechaEjecucion = Now
   oTabla.NumeroSesion = 0
   oTabla.idAtencion = 0
   oTabla.IdEstablecimiento = 0
End Sub
