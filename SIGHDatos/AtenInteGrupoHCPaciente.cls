VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtenInteGrupoHCPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtenInteGrupoHCPaciente
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
'   Descripción:    Inserta un registro a la tabla AtenInteGrupoHCPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtenInteGrupoHCPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteGrupoHCPacienteAgregar"
   Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdGrupoHCPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdGrupoHCPaciente = 0, Null, oTabla.IdGrupoHCPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAtenInteGrupo", adInteger, adParamInput, 0, IIf(oTabla.IdAtenInteGrupo = 0, Null, oTabla.IdAtenInteGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPregunta", adInteger, adParamInput, 0, IIf(oTabla.IdPregunta = 0, Null, oTabla.IdPregunta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdPaciente = .Parameters("@IdPaciente")
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
'   Descripción:    modifica un registro a la tabla AtenInteGrupoHCPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtenInteGrupoHCPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteGrupoHCPacienteModificar"
   Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdGrupoHCPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdGrupoHCPaciente = 0, Null, oTabla.IdGrupoHCPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAtenInteGrupo", adInteger, adParamInput, 0, IIf(oTabla.IdAtenInteGrupo = 0, Null, oTabla.IdAtenInteGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPregunta", adInteger, adParamInput, 0, IIf(oTabla.IdPregunta = 0, Null, oTabla.IdPregunta))
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
'   Descripción:    elimia o borra un registro a la tabla AtenInteGrupoHCPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtenInteGrupoHCPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteGrupoHCPacienteEliminar"
   Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente)
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
'   Descripción:    Selecciona un unico registro de la tabla AtenInteGrupoHCPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtenInteGrupoHCPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtenInteGrupoHCPacienteSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.IdGrupoHCPaciente = IIf(IsNull(oRecordset!IdGrupoHCPaciente), 0, oRecordset!IdGrupoHCPaciente)
       oTabla.IdAtenInteGrupo = IIf(IsNull(oRecordset!IdAtenInteGrupo), 0, oRecordset!IdAtenInteGrupo)
       oTabla.IdPregunta = IIf(IsNull(oRecordset!IdPregunta), 0, oRecordset!IdPregunta)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOAtenInteGrupoHCPaciente)
   oTabla.IdPaciente = 0
   oTabla.IdGrupoHCPaciente = 0
   oTabla.IdAtenInteGrupo = 0
   oTabla.IdPregunta = 0
End Sub

Public Function ListarPreguntasPorPacienteYGrupo(ByVal oGrupoHcPaciente As DOAtenInteGrupoHCPaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
  
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtenInteListarPreguntasPorPacienteYGrupo"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oGrupoHcPaciente.IdPaciente)
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtenInteGrupo", adInteger, adParamInput, 0, IIf(oGrupoHcPaciente.IdAtenInteGrupo = 0, Null, oGrupoHcPaciente.IdAtenInteGrupo))
        .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set ListarPreguntasPorPacienteYGrupo = oRecordset
    ms_MensajeError = ""
    Exit Function
ManejadorDeError:
    If Err Then
        ms_MensajeError = Err.Number & " " + Err.Description
    End If
End Function


Function verificarInsertar(ByRef oTabla As DOAtenInteGrupoHCPaciente) As Boolean
On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
     
    verificarInsertar = False
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtenInteVerificaIngresaPreguntaPaciente"
       Set oParameter = .CreateParameter("@IdGrupoHCPaciente", adInteger, adParamOutput, 0)
       .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
       .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdAtenInteGrupo", adInteger, adParamInput, 0, IIf(oTabla.IdAtenInteGrupo = 0, Null, oTabla.IdAtenInteGrupo))
       .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdPregunta", adInteger, adParamInput, 0, IIf(oTabla.IdPregunta = 0, Null, oTabla.IdPregunta))
       .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
       oTabla.IdGrupoHCPaciente = .Parameters("@IdGrupoHCPaciente")
    End With
     
    verificarInsertar = True
    ms_MensajeError = ""
     
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
End Function

 


