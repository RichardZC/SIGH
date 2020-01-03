VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtenIntePlanDesPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtenIntePlanDesarrolloPaciente
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
'   Descripción:    Inserta un registro a la tabla AtenIntePlanDesarrolloPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtenIntePlanDesPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanDesarrolloPacienteAgregar"
   Set oParameter = .CreateParameter("@IdPlanDesarrolloPaciente", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanIntegralPaciente", adBigInt, adParamInput, 0, oTabla.IdPlanIntegralPaciente)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Evaluacion", adInteger, adParamInput, 0, IIf(oTabla.Evaluacion = 0, Null, oTabla.Evaluacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdPlanAtencion = 0, Null, oTabla.IdPlanAtencion))
   .Parameters.Append oParameter
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
'   Descripción:    modifica un registro a la tabla AtenIntePlanDesarrolloPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtenIntePlanDesPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanDesarrolloPacienteModificar"
   Set oParameter = .CreateParameter("@IdPlanDesarrolloPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPlanDesarrolloPaciente = 0, Null, oTabla.IdPlanDesarrolloPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanIntegralPaciente", adBigInt, adParamInput, 0, oTabla.IdPlanIntegralPaciente)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Evaluacion", adInteger, adParamInput, 0, IIf(oTabla.Evaluacion = 0, Null, oTabla.Evaluacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdPlanAtencion = 0, Null, oTabla.IdPlanAtencion))
   .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla AtenIntePlanDesarrolloPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtenIntePlanDesPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanDesarrolloPacienteEliminar"
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
'   Descripción:    Selecciona un unico registro de la tabla AtenIntePlanDesarrolloPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
'mgaray20141012
Function SeleccionarPorId(ByVal oTabla As DOAtenIntePlanDesPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
'   With oCommand
'     .CommandType = adCmdStoredProc
'     Set .ActiveConnection = mo_Conexion
'     .CommandText = "AtenIntePlanDesarrolloPacienteSeleccionarPorId"
'     Set oParameter = .CreateParameter("@IdPlanDesarrolloPaciente", adInteger, adParamInput, 0, oTabla.IdPlanDesarrolloPaciente)
'     .Parameters.Append oParameter
'     Set oParameter = .CreateParameter("@IdPlanIntegralPaciente", adBigInt, adParamInput, 0, oTabla.IdPlanIntegralPaciente)
'     .Parameters.Append oParameter
'     Set oRecordset = .Execute
'   End With
   
   Set oRecordset = setRsSeleccionarPorId(oTabla)
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdPlanDesarrolloPaciente = IIf(IsNull(oRecordset!IdPlanDesarrolloPaciente), 0, oRecordset!IdPlanDesarrolloPaciente)
       oTabla.IdPlanIntegralPaciente = oRecordset!IdPlanIntegralPaciente
       oTabla.Evaluacion = IIf(IsNull(oRecordset!Evaluacion), 0, oRecordset!Evaluacion)
       oTabla.IdPlanAtencion = IIf(IsNull(oRecordset!IdPlanAtencion), 0, oRecordset!IdPlanAtencion)
       oTabla.IdAtenInteItemPlan = IIf(IsNull(oRecordset!IdAtenInteItemPlan), 0, oRecordset!IdAtenInteItemPlan)
       oTabla.FechaProgramada = IIf(IsNull(oRecordset!FechaProgramada), 0, oRecordset!FechaProgramada)
       oTabla.FechaEjecucion = IIf(IsNull(oRecordset!FechaEjecucion), 0, oRecordset!FechaEjecucion)
       oTabla.NumeroSesion = IIf(IsNull(oRecordset!NumeroSesion), 0, oRecordset!NumeroSesion)
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       oTabla.IdEstablecimiento = IIf(IsNull(oRecordset!IdEstablecimiento), 0, oRecordset!IdEstablecimiento)
       
       oTabla.EdadAnio = IIf(IsNull(oRecordset!EdadAnio), 0, oRecordset!EdadAnio)
       oTabla.EdadMes = IIf(IsNull(oRecordset!EdadMes), 0, oRecordset!EdadMes)
       oTabla.EdadDia = IIf(IsNull(oRecordset!EdadDia), 0, oRecordset!EdadDia)
       
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
       oTabla.EvaluacionDesc = IIf(IsNull(oRecordset!EvaluacionDesc), "", oRecordset!EvaluacionDesc)
       oTabla.Establecimiento = IIf(IsNull(oRecordset!Establecimiento), "", oRecordset!Establecimiento)
      SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'mgaray20141012
Function setRsSeleccionarPorId(ByVal oTabla As DOAtenIntePlanDesPaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
 
'    Set setRsSeleccionarPorId = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtenIntePlanDesarrolloPacienteSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdPlanDesarrolloPaciente", adInteger, adParamInput, 0, oTabla.IdPlanDesarrolloPaciente)
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPlanIntegralPaciente", adBigInt, adParamInput, 0, oTabla.IdPlanIntegralPaciente)
        .Parameters.Append oParameter
'        Set oRecordset = .Execute
'        oRecordset.Open
        oRecordset.Open oCommand, , adOpenStatic, adLockOptimistic
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set setRsSeleccionarPorId = oRecordset
   
   ms_MensajeError = ""
'   oRecordset.Close
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub SetDefaults(ByVal oTabla As DOAtenIntePlanDesPaciente)
   oTabla.IdPlanDesarrolloPaciente = 0
   oTabla.IdPlanIntegralPaciente = 0
   oTabla.Evaluacion = 0
   oTabla.IdPlanAtencion = 0
   oTabla.IdAtenInteItemPlan = 0
   oTabla.FechaProgramada = Now
   oTabla.FechaEjecucion = Now
   oTabla.NumeroSesion = 0
   oTabla.idAtencion = 0
   oTabla.IdEstablecimiento = 0
End Sub

Public Function ListarPlanDesarrolloPaciente(ByVal oTabla As DOAtenIntePlanIntePaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim sSql As String
    
    sSql = "SHAPE {EXEC AtenInteListarPlanIntegralDesarrolloPaciente " & oTabla.IdAtenInteGrupo _
                    & " , " & IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente) _
                    & "," & oTabla.IdAtenInteItemPlan & "} " + _
            "APPEND ({EXEC AtenInteListarPlanDesarrolloPacienteDet " & oTabla.IdAtenInteGrupo _
                    & " , " & IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente) _
                    & "," & oTabla.IdAtenInteItemPlan & "} " + _
            " AS detalleDesarrollo RELATE IdPlanAtencion TO IdPlanAtencion) "
            
    oRecordset.Open sSql, mo_Conexion, adOpenStatic, adLockOptimistic
    Set oRecordset.ActiveConnection = Nothing
    
    Set ListarPlanDesarrolloPaciente = oRecordset
    ms_MensajeError = ""
    Exit Function
ManejadorDeError:
    If Err Then
        ms_MensajeError = Err.Number & " " + Err.Description
    End If
End Function

Public Function ListarPlanDesarrolloPacientePendientes(ByVal oTabla As DOAtenIntePlanIntePaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim sSql As String
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtenInteListarDesarrolloPacientePendientesDet"
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
    
    Set ListarPlanDesarrolloPacientePendientes = oRecordset
    ms_MensajeError = ""
    Exit Function
ManejadorDeError:
    If Err Then
        ms_MensajeError = Err.Number & " " + Err.Description
    End If
End Function

Public Function ItemDesarrolloPacientePendiente(ByVal oTabla As DOAtenIntePlanIntePaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim sSql As String
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtenInteItemDesarrolloPacientePendiente"
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
    
    Set ItemDesarrolloPacientePendiente = oRecordset
    ms_MensajeError = ""
    Exit Function
ManejadorDeError:
    If Err Then
        ms_MensajeError = Err.Number & " " + Err.Description
    End If
End Function

Function ModificarFechaProgramada(ByVal oTabla As DOAtenIntePlanDesPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
ModificarFechaProgramada = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanDesarrolloPacienteModFechaProgramada"
   Set oParameter = .CreateParameter("@IdPlanDesarrolloPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPlanDesarrolloPaciente = 0, Null, oTabla.IdPlanDesarrolloPaciente))
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

Public Function ListarPlanDesarrolloPacientePendientesParaImpresion(ByVal oTabla As DOAtenIntePlanIntePaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim sSql As String
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtenInteListarPlanDesarrolloPacienteDet"
        Set oParameter = .CreateParameter("@IdAtenInteGrupo", adInteger, adParamInput, 0, oTabla.IdAtenInteGrupo)
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtenInteItemPlan", adInteger, adParamInput, 0, oTabla.IdAtenInteItemPlan)
        .Parameters.Append oParameter
'        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oTabla.idAtencion)
'        .Parameters.Append oParameter


        Set oRecordset = .Execute
'        oRecordset.Open oCommand, , adOpenStatic, adLockOptimistic
        Set oRecordset.ActiveConnection = Nothing
    End With
    
    Set ListarPlanDesarrolloPacientePendientesParaImpresion = oRecordset
    ms_MensajeError = ""
    Exit Function
ManejadorDeError:
    If Err Then
        ms_MensajeError = Err.Number & " " + Err.Description
    End If
End Function


Function EliminarAtencionIntegral(ByVal oTabla As DOAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
EliminarAtencionIntegral = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanAtecionPacienteEliminaEjecucion"
   Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oTabla.idAtencion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente)
   .Parameters.Append oParameter
   .Execute
End With
 
EliminarAtencionIntegral = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'mgaray20141012
Public Function ListarPlanDesarrolloPacienteVencidos(ByVal oTabla As DOAtenIntePlanIntePaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim sSql As String
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtenInteListarPlanDesarrolloPacienteVencidos"
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
    
    Set ListarPlanDesarrolloPacienteVencidos = oRecordset
    ms_MensajeError = ""
    Exit Function
ManejadorDeError:
    If Err Then
        ms_MensajeError = Err.Number & " " + Err.Description
    End If
End Function
