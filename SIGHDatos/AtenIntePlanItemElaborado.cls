VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtenIntePlanItemElaborado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtenIntePlanItemElaborado
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
'   Descripción:    Inserta un registro a la tabla AtenIntePlanItemElaborado
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtenIntePlanItemElaborado) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanItemElaboradoAgregar"
   Set oParameter = .CreateParameter("@IdAtenInteItemPlan", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanIntegralPaciente", adBigInt, adParamInput, 0, oTabla.IdPlanIntegralPaciente)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsElaborado", adBoolean, adParamInput, 0, IIf(oTabla.EsElaborado = 0, Null, oTabla.EsElaborado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdAtenInteItemPlan = .Parameters("@IdAtenInteItemPlan")
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
'   Descripción:    modifica un registro a la tabla AtenIntePlanItemElaborado
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtenIntePlanItemElaborado) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanItemElaboradoModificar"
   Set oParameter = .CreateParameter("@IdAtenInteItemPlan", adInteger, adParamInput, 0, IIf(oTabla.IdAtenInteItemPlan = 0, Null, oTabla.IdAtenInteItemPlan))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanIntegralPaciente", adBigInt, adParamInput, 0, oTabla.IdPlanIntegralPaciente)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsElaborado", adBoolean, adParamInput, 0, IIf(oTabla.EsElaborado = 0, Null, oTabla.EsElaborado))
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
'   Descripción:    elimia o borra un registro a la tabla AtenIntePlanItemElaborado
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtenIntePlanItemElaborado) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanItemElaboradoEliminar"
   Set oParameter = .CreateParameter("@IdAtenInteItemPlan", adInteger, adParamInput, 0, oTabla.IdAtenInteItemPlan)
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
'   Descripción:    Selecciona un unico registro de la tabla AtenIntePlanItemElaborado
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtenIntePlanItemElaborado) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtenIntePlanItemElaboradoSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdAtenInteItemPlan", adInteger, adParamInput, 0, oTabla.IdAtenInteItemPlan)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdAtenInteItemPlan = IIf(IsNull(oRecordset!IdAtenInteItemPlan), 0, oRecordset!IdAtenInteItemPlan)
       oTabla.IdPlanIntegralPaciente = oRecordset!IdPlanIntegralPaciente
       oTabla.EsElaborado = IIf(IsNull(oRecordset!EsElaborado), 0, oRecordset!EsElaborado)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOAtenIntePlanItemElaborado)
   oTabla.IdAtenInteItemPlan = 0
   oTabla.IdPlanIntegralPaciente = 0
   oTabla.EsElaborado = False
End Sub
