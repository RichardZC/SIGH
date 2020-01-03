VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtenIntePlanProcedPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtenIntePlanProcedimientoPaciente
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
'   Descripción:    Inserta un registro a la tabla AtenIntePlanProcedimientoPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtenIntePlanProcedPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanProcedimientoPacienteAgregar"
   Set oParameter = .CreateParameter("@IdPlanProcedimientoPaciente", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanIntegralPaciente", adBigInt, adParamInput, 0, oTabla.IdPlanIntegralPaciente)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdPlanAtencion = 0, Null, oTabla.IdPlanAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAtenInteItemPlan", adInteger, adParamInput, 0, IIf(oTabla.IdAtenInteItemPlan = 0, Null, oTabla.IdAtenInteItemPlan))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaProgramada", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaProgramada = 0, Null, oTabla.FechaProgramada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaEjecucion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEjecucion = 0, Null, oTabla.FechaEjecucion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NumeroDosis", adUnsignedTinyInt, adParamInput, 0, oTabla.NumeroDosis)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoHIS", adVarChar, adParamInput, 7, IIf(oTabla.CodigoHIS = "", Null, oTabla.CodigoHIS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdPlanProcedimientoPaciente = .Parameters("@IdPlanProcedimientoPaciente")
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
'   Descripción:    modifica un registro a la tabla AtenIntePlanProcedimientoPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtenIntePlanProcedPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanProcedimientoPacienteModificar"
   Set oParameter = .CreateParameter("@IdPlanProcedimientoPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPlanProcedimientoPaciente = 0, Null, oTabla.IdPlanProcedimientoPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanIntegralPaciente", adBigInt, adParamInput, 0, oTabla.IdPlanIntegralPaciente)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdPlanAtencion = 0, Null, oTabla.IdPlanAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAtenInteItemPlan", adInteger, adParamInput, 0, IIf(oTabla.IdAtenInteItemPlan = 0, Null, oTabla.IdAtenInteItemPlan))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaProgramada", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaProgramada = 0, Null, oTabla.FechaProgramada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaEjecucion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEjecucion = 0, Null, oTabla.FechaEjecucion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NumeroDosis", adUnsignedTinyInt, adParamInput, 0, oTabla.NumeroDosis)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoHIS", adVarChar, adParamInput, 7, IIf(oTabla.CodigoHIS = "", Null, oTabla.CodigoHIS))
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
'   Descripción:    elimia o borra un registro a la tabla AtenIntePlanProcedimientoPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtenIntePlanProcedPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanProcedimientoPacienteEliminar"
   Set oParameter = .CreateParameter("@IdPlanProcedimientoPaciente", adInteger, adParamInput, 0, oTabla.IdPlanProcedimientoPaciente)
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
'   Descripción:    Selecciona un unico registro de la tabla AtenIntePlanProcedimientoPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtenIntePlanProcedPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtenIntePlanProcedimientoPacienteSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdPlanProcedimientoPaciente", adInteger, adParamInput, 0, oTabla.IdPlanProcedimientoPaciente)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdPlanProcedimientoPaciente = IIf(IsNull(oRecordset!IdPlanProcedimientoPaciente), 0, oRecordset!IdPlanProcedimientoPaciente)
       oTabla.IdPlanIntegralPaciente = oRecordset!IdPlanIntegralPaciente
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.IdPlanAtencion = IIf(IsNull(oRecordset!IdPlanAtencion), 0, oRecordset!IdPlanAtencion)
       oTabla.IdAtenInteItemPlan = IIf(IsNull(oRecordset!IdAtenInteItemPlan), 0, oRecordset!IdAtenInteItemPlan)
       oTabla.FechaProgramada = IIf(IsNull(oRecordset!FechaProgramada), 0, oRecordset!FechaProgramada)
       oTabla.FechaEjecucion = IIf(IsNull(oRecordset!FechaEjecucion), 0, oRecordset!FechaEjecucion)
       oTabla.NumeroDosis = oRecordset!NumeroDosis
       oTabla.CodigoHIS = IIf(IsNull(oRecordset!CodigoHIS), "", oRecordset!CodigoHIS)
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
Sub SetDefaults(ByVal oTabla As DOAtenIntePlanProcedPaciente)
   oTabla.IdPlanProcedimientoPaciente = 0
   oTabla.IdPlanIntegralPaciente = 0
   oTabla.idProducto = 0
   oTabla.IdPlanAtencion = 0
   oTabla.IdAtenInteItemPlan = 0
   oTabla.FechaProgramada = Now
   oTabla.FechaEjecucion = Now
   oTabla.NumeroDosis = 0
   oTabla.CodigoHIS = ""
   oTabla.idAtencion = 0
   oTabla.IdEstablecimiento = 0
End Sub


Public Function ListarPlanProcedimientosPaciente(ByVal oTabla As DOAtenIntePlanIntePaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim sSql As String
  
'    With oCommand
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = mo_Conexion
'        .CommandText = "AtenInteListarPlanProcedimientosPaciente"
'        Set oParameter = .CreateParameter("@IdAtenInteGrupo", adInteger, adParamInput, 0, oTabla.IdAtenInteGrupo)
'        .Parameters.Append oParameter
'        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
'        .Parameters.Append oParameter
'        Set oParameter = .CreateParameter("@IdAtenInteItemPlan", adInteger, adParamInput, 0, oTabla.IdAtenInteItemPlan)
'        .Parameters.Append oParameter
'        Set oRecordset = .Execute
'        Set oRecordset.ActiveConnection = Nothing
'    End With
    
    sSql = "SHAPE {EXEC AtenInteListarPlanIntegralProcedimientosPaciente " & oTabla.IdAtenInteGrupo _
                    & " , " & IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente) _
                    & "," & oTabla.IdAtenInteItemPlan & "} " + _
            "APPEND ({EXEC AtenInteListarPlanProcedimientosPaciente " & oTabla.IdAtenInteGrupo _
                    & " , " & IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente) _
                    & "," & oTabla.IdAtenInteItemPlan & "} " + _
            " AS detalleProcedimiento RELATE IdPlanAtencion TO IdPlanAtencion) "
            
    oRecordset.Open sSql, mo_Conexion, adOpenStatic, adLockOptimistic
    Set oRecordset.ActiveConnection = Nothing
    
    Set ListarPlanProcedimientosPaciente = oRecordset
    ms_MensajeError = ""
    Exit Function
ManejadorDeError:
    If Err Then
        ms_MensajeError = Err.Number & " " + Err.Description
    End If
End Function

Public Function ListarPlanProcedimientosPacientePendientes(ByVal oTabla As DOAtenIntePlanIntePaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim sSql As String
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtenInteListarProcedimientosPacientePendientes"
        Set oParameter = .CreateParameter("@IdAtenInteGrupo", adInteger, adParamInput, 0, oTabla.IdAtenInteGrupo)
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtenInteItemPlan", adInteger, adParamInput, 0, oTabla.IdAtenInteItemPlan)
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oTabla.idAtencion)
        .Parameters.Append oParameter


'        Set oRecordset = .Execute
        oRecordset.Open oCommand, , adOpenStatic, adLockOptimistic
        Set oRecordset.ActiveConnection = Nothing
    End With
    
'    sSql = "SHAPE {EXEC AtenInteListarPlanProcedimientosPacientePendientes " & oTabla.IdAtenInteGrupo _
'                    & " , " & IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente) _
'                    & "," & oTabla.IdAtenInteItemPlan & "," & oTabla.idAtencion & "} " + _
'            "APPEND ({EXEC AtenInteListarProcedimientosPacientePendientes " & oTabla.IdAtenInteGrupo _
'                    & " , " & IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente) _
'                    & "," & oTabla.IdAtenInteItemPlan & "," & oTabla.idAtencion & "} " + _
'            " AS title_chap RELATE IdPlanAtencion TO IdPlanAtencion) "
'
'    oRecordset.Open sSql, mo_Conexion, adOpenStatic, adLockOptimistic
'    Set oRecordset.ActiveConnection = Nothing
    
    Set ListarPlanProcedimientosPacientePendientes = oRecordset
    ms_MensajeError = ""
    Exit Function
ManejadorDeError:
    If Err Then
        ms_MensajeError = Err.Number & " " + Err.Description
    End If
End Function


Function ModificarFechaProgramada(ByVal oTabla As DOAtenIntePlanProcedPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
ModificarFechaProgramada = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanProcedimientoPacienteModFechaProgramada"
   Set oParameter = .CreateParameter("@IdPlanProcedimientoPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPlanProcedimientoPaciente = 0, Null, oTabla.IdPlanProcedimientoPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanIntegralPaciente", adBigInt, adParamInput, 0, oTabla.IdPlanIntegralPaciente)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaProgramada", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaProgramada = 0, Null, oTabla.FechaProgramada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
ModificarFechaProgramada = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
