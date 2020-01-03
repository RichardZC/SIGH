VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtenIntePregunta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtenIntePregunta
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
'   Descripción:    Inserta un registro a la tabla AtenIntePregunta
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtenIntePregunta) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePreguntaAgregar"
   Set oParameter = .CreateParameter("@IdPregunta", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Pregunta", adVarChar, adParamInput, 70, IIf(oTabla.Pregunta = "", Null, oTabla.Pregunta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoRespuesta", adInteger, adParamInput, 0, IIf(oTabla.TipoRespuesta = 0, Null, oTabla.TipoRespuesta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoValorRespuesta", adInteger, adParamInput, 0, IIf(oTabla.TipoValorRespuesta = 0, Null, oTabla.TipoValorRespuesta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdPregunta = .Parameters("@IdPregunta")
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
'   Descripción:    modifica un registro a la tabla AtenIntePregunta
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtenIntePregunta) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePreguntaModificar"
   Set oParameter = .CreateParameter("@IdPregunta", adInteger, adParamInput, 0, IIf(oTabla.IdPregunta = 0, Null, oTabla.IdPregunta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Pregunta", adVarChar, adParamInput, 70, IIf(oTabla.Pregunta = "", Null, oTabla.Pregunta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoRespuesta", adInteger, adParamInput, 0, IIf(oTabla.TipoRespuesta = 0, Null, oTabla.TipoRespuesta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoValorRespuesta", adInteger, adParamInput, 0, IIf(oTabla.TipoValorRespuesta = 0, Null, oTabla.TipoValorRespuesta))
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
'   Descripción:    elimia o borra un registro a la tabla AtenIntePregunta
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtenIntePregunta) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePreguntaEliminar"
   Set oParameter = .CreateParameter("@IdPregunta", adInteger, adParamInput, 0, oTabla.IdPregunta)
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
'   Descripción:    Selecciona un unico registro de la tabla AtenIntePregunta
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtenIntePregunta) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtenIntePreguntaSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdPregunta", adInteger, adParamInput, 0, oTabla.IdPregunta)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdPregunta = IIf(IsNull(oRecordset!IdPregunta), 0, oRecordset!IdPregunta)
       oTabla.Pregunta = IIf(IsNull(oRecordset!Pregunta), "", oRecordset!Pregunta)
       oTabla.TipoRespuesta = IIf(IsNull(oRecordset!TipoRespuesta), 0, oRecordset!TipoRespuesta)
       oTabla.TipoValorRespuesta = IIf(IsNull(oRecordset!TipoValorRespuesta), 0, oRecordset!TipoValorRespuesta)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub SetDefaults(ByVal oTabla As DOAtenIntePregunta)
   oTabla.IdPregunta = 0
   oTabla.Pregunta = ""
   oTabla.TipoRespuesta = 0
   oTabla.TipoValorRespuesta = 0
End Sub
