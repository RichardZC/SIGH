VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TriajeVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla TriajeVariable
'        Programado por: Garay M
'        Fecha: Agosto 2014
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
'   Descripción:    Inserta un registro a la tabla TriajeVariable
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOTriajeVariable) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "TriajeVariableAgregar"
   Set oParameter = .CreateParameter("@IdTriajeVariable", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeVariable", adVarChar, adParamInput, 50, IIf(oTabla.TriajeVariable = "", Null, oTabla.TriajeVariable))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsAntropometrica", adBoolean, adParamInput, 0, IIf(oTabla.EsAntropometrica = 0, Null, oTabla.EsAntropometrica))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TieneLimiteMedicion", adBoolean, adParamInput, 0, IIf(oTabla.TieneLimiteMedicion = 0, Null, oTabla.TieneLimiteMedicion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadDiaLimiteMinima", adInteger, adParamInput, 0, IIf(oTabla.EdadDiaLimiteMinima = 0, Null, oTabla.EdadDiaLimiteMinima))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadDiaLimiteMaxima", adInteger, adParamInput, 0, IIf(oTabla.EdadDiaLimiteMaxima = 0, Null, oTabla.EdadDiaLimiteMaxima))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsDatoObligatorio", adBoolean, adParamInput, 0, IIf(oTabla.EsDatoObligatorio = 0, Null, oTabla.EsDatoObligatorio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsActivo", adBoolean, adParamInput, 0, IIf(oTabla.EsActivo = 0, Null, oTabla.EsActivo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdTriajeVariable = .Parameters("@IdTriajeVariable")
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
'   Descripción:    modifica un registro a la tabla TriajeVariable
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOTriajeVariable) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "TriajeVariableModificar"
   Set oParameter = .CreateParameter("@IdTriajeVariable", adInteger, adParamInput, 0, IIf(oTabla.IdTriajeVariable = 0, Null, oTabla.IdTriajeVariable))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeVariable", adVarChar, adParamInput, 50, IIf(oTabla.TriajeVariable = "", Null, oTabla.TriajeVariable))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsAntropometrica", adBoolean, adParamInput, 0, IIf(oTabla.EsAntropometrica = 0, Null, oTabla.EsAntropometrica))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TieneLimiteMedicion", adBoolean, adParamInput, 0, IIf(oTabla.TieneLimiteMedicion = 0, Null, oTabla.TieneLimiteMedicion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadDiaLimiteMinima", adInteger, adParamInput, 0, IIf(oTabla.EdadDiaLimiteMinima = 0, Null, oTabla.EdadDiaLimiteMinima))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadDiaLimiteMaxima", adInteger, adParamInput, 0, IIf(oTabla.EdadDiaLimiteMaxima = 0, Null, oTabla.EdadDiaLimiteMaxima))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsDatoObligatorio", adBoolean, adParamInput, 0, IIf(oTabla.EsDatoObligatorio = 0, Null, oTabla.EsDatoObligatorio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsActivo", adBoolean, adParamInput, 0, IIf(oTabla.EsActivo = 0, Null, oTabla.EsActivo))
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
'   Descripción:    elimia o borra un registro a la tabla TriajeVariable
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOTriajeVariable) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "TriajeVariableEliminar"
   Set oParameter = .CreateParameter("@IdTriajeVariable", adInteger, adParamInput, 0, oTabla.IdTriajeVariable)
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
'   Descripción:    Selecciona un unico registro de la tabla TriajeVariable
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOTriajeVariable) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TriajeVariableSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdTriajeVariable", adInteger, adParamInput, 0, oTabla.IdTriajeVariable)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdTriajeVariable = IIf(IsNull(oRecordset!IdTriajeVariable), 0, oRecordset!IdTriajeVariable)
       oTabla.TriajeVariable = IIf(IsNull(oRecordset!TriajeVariable), "", oRecordset!TriajeVariable)
       oTabla.EsAntropometrica = IIf(IsNull(oRecordset!EsAntropometrica), 0, oRecordset!EsAntropometrica)
       oTabla.TieneLimiteMedicion = IIf(IsNull(oRecordset!TieneLimiteMedicion), 0, oRecordset!TieneLimiteMedicion)
       oTabla.EdadDiaLimiteMinima = IIf(IsNull(oRecordset!EdadDiaLimiteMinima), 0, oRecordset!EdadDiaLimiteMinima)
       oTabla.EdadDiaLimiteMaxima = IIf(IsNull(oRecordset!EdadDiaLimiteMaxima), 0, oRecordset!EdadDiaLimiteMaxima)
       oTabla.EsDatoObligatorio = IIf(IsNull(oRecordset!EsDatoObligatorio), 0, oRecordset!EsDatoObligatorio)
       oTabla.EsActivo = IIf(IsNull(oRecordset!EsActivo), 0, oRecordset!EsActivo)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOTriajeVariable)
   oTabla.IdTriajeVariable = 0
   oTabla.TriajeVariable = ""
   oTabla.EsAntropometrica = False
   oTabla.TieneLimiteMedicion = False
   oTabla.EdadDiaLimiteMinima = 0
   oTabla.EdadDiaLimiteMaxima = 0
   oTabla.EsDatoObligatorio = False
   oTabla.EsActivo = False
End Sub


Function ListarVariableTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set ListarVariableTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TriajeListarVariableTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set ListarVariableTodos = oRecordset
 
   ms_MensajeError = ""
   Set oCommand = Nothing
   Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function ListarVariablesAntropometricas() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set ListarVariablesAntropometricas = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TriajeListarVariableAntropometricas"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set ListarVariablesAntropometricas = oRecordset
 
   ms_MensajeError = ""
   Set oCommand = Nothing
   Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

