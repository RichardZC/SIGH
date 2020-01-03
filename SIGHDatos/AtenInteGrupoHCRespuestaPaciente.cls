VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtenInteHCRespuestaPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtenInteGrupoHCRespuestaPaciente
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
'   Descripción:    Inserta un registro a la tabla AtenInteGrupoHCRespuestaPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtenInteHCRespuestaPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteGrupoHCRespuestaPacienteAgregar"
   Set oParameter = .CreateParameter("@IdGrupoHCPaciente", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ItemRespuesta", adInteger, adParamInput, 0, IIf(oTabla.ItemRespuesta = 0, Null, oTabla.ItemRespuesta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorTexto", adVarChar, adParamInput, 255, IIf(oTabla.ValorTexto = "", Null, oTabla.ValorTexto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorNumero", adCurrency, adParamInput, 0, oTabla.ValorNumero)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorFecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.ValorFecha = 0, Null, oTabla.ValorFecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorNumeroFin", adCurrency, adParamInput, 0, oTabla.ValorNumeroFin)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorFechaFin", adDBTimeStamp, adParamInput, 0, IIf(oTabla.ValorFechaFin = 0, Null, oTabla.ValorFechaFin))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorEspecificacion", adVarChar, adParamInput, 1500, IIf(oTabla.ValorEspecificacion = "", Null, oTabla.ValorEspecificacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsActivo", adBoolean, adParamInput, 0, IIf(oTabla.EsActivo = 0, Null, oTabla.EsActivo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdGrupoHCPaciente = .Parameters("@IdGrupoHCPaciente")
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
'   Descripción:    modifica un registro a la tabla AtenInteGrupoHCRespuestaPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtenInteHCRespuestaPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteGrupoHCRespuestaPacienteModificar"
   Set oParameter = .CreateParameter("@IdGrupoHCPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdGrupoHCPaciente = 0, Null, oTabla.IdGrupoHCPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ItemRespuesta", adInteger, adParamInput, 0, IIf(oTabla.ItemRespuesta = 0, Null, oTabla.ItemRespuesta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorTexto", adVarChar, adParamInput, 255, IIf(oTabla.ValorTexto = "", Null, oTabla.ValorTexto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorNumero", adCurrency, adParamInput, 0, oTabla.ValorNumero)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorFecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.ValorFecha = 0, Null, oTabla.ValorFecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorNumeroFin", adCurrency, adParamInput, 0, oTabla.ValorNumeroFin)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorFechaFin", adDBTimeStamp, adParamInput, 0, IIf(oTabla.ValorFechaFin = 0, Null, oTabla.ValorFechaFin))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorEspecificacion", adVarChar, adParamInput, 1500, IIf(oTabla.ValorEspecificacion = "", Null, oTabla.ValorEspecificacion))
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
'   Descripción:    elimia o borra un registro a la tabla AtenInteGrupoHCRespuestaPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtenInteHCRespuestaPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteGrupoHCRespuestaPacienteEliminar"
   Set oParameter = .CreateParameter("@IdGrupoHCPaciente", adInteger, adParamInput, 0, oTabla.IdGrupoHCPaciente)
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
'   Descripción:    Selecciona un unico registro de la tabla AtenInteGrupoHCRespuestaPaciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtenInteHCRespuestaPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtenInteGrupoHCRespuestaPacienteSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdGrupoHCPaciente", adInteger, adParamInput, 0, oTabla.IdGrupoHCPaciente)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdGrupoHCPaciente = IIf(IsNull(oRecordset!IdGrupoHCPaciente), 0, oRecordset!IdGrupoHCPaciente)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.ItemRespuesta = IIf(IsNull(oRecordset!ItemRespuesta), 0, oRecordset!ItemRespuesta)
       oTabla.ValorTexto = IIf(IsNull(oRecordset!ValorTexto), "", oRecordset!ValorTexto)
       oTabla.ValorNumero = oRecordset!ValorNumero
       oTabla.ValorFecha = IIf(IsNull(oRecordset!ValorFecha), 0, oRecordset!ValorFecha)
       oTabla.ValorNumeroFin = oRecordset!ValorNumeroFin
       oTabla.ValorFechaFin = IIf(IsNull(oRecordset!ValorFechaFin), 0, oRecordset!ValorFechaFin)
       oTabla.ValorEspecificacion = IIf(IsNull(oRecordset!ValorEspecificacion), "", oRecordset!ValorEspecificacion)
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
Sub SetDefaults(ByVal oTabla As DOAtenInteHCRespuestaPaciente)
   oTabla.IdGrupoHCPaciente = 0
   oTabla.IdPaciente = 0
   oTabla.ItemRespuesta = 0
   oTabla.ValorTexto = ""
   oTabla.ValorNumero = 0
   oTabla.ValorFecha = Now
   oTabla.ValorNumeroFin = 0
   oTabla.ValorFechaFin = Now
   oTabla.ValorEspecificacion = ""
   oTabla.EsActivo = False
End Sub

Public Function ListarRespuestasPorPacienteYGrupo(ByVal oGrupoHcPaciente As DOAtenInteGrupoHCPaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
  
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtenInteListarRespuestasPorPacienteYGrupo"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oGrupoHcPaciente.IdPaciente)
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtenInteGrupo", adInteger, adParamInput, 0, IIf(oGrupoHcPaciente.IdAtenInteGrupo = 0, Null, oGrupoHcPaciente.IdAtenInteGrupo))
        .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set ListarRespuestasPorPacienteYGrupo = oRecordset
    ms_MensajeError = ""
    Exit Function
ManejadorDeError:
    If Err Then
        ms_MensajeError = Err.Number & " " + Err.Description
    End If
End Function

Function verificaInsertar(ByVal oTabla As DOAtenInteHCRespuestaPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
verificaInsertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteGrupoHCRespuestaPacienteVerificarAgregar"
   Set oParameter = .CreateParameter("@IdGrupoHCPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdGrupoHCPaciente = 0, Null, oTabla.IdGrupoHCPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ItemRespuesta", adInteger, adParamInput, 0, IIf(oTabla.ItemRespuesta = 0, Null, oTabla.ItemRespuesta))
'   Set oParameter = .CreateParameter("@ItemRespuesta", adInteger, adParamInput, 1)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorTexto", adVarChar, adParamInput, 255, IIf(oTabla.ValorTexto = "", Null, oTabla.ValorTexto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorNumero", adCurrency, adParamInput, 0, oTabla.ValorNumero)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorFecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.ValorFecha = 0, Null, oTabla.ValorFecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorNumeroFin", adCurrency, adParamInput, 0, oTabla.ValorNumeroFin)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorFechaFin", adDBTimeStamp, adParamInput, 0, IIf(oTabla.ValorFechaFin = 0, Null, oTabla.ValorFechaFin))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorEspecificacion", adVarChar, adParamInput, 1500, IIf(oTabla.ValorEspecificacion = "", Null, oTabla.ValorEspecificacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsActivo", adBoolean, adParamInput, 0, IIf(oTabla.EsActivo = 0, Null, oTabla.EsActivo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   'oTabla.ItemRespuesta = .Parameters("@ItemRespuesta")
End With
 
verificaInsertar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


Function EliminarRespuestaPorPregunta(ByVal oTabla As DOAtenInteGrupoHCPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
EliminarRespuestaPorPregunta = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteGrupoHCRptaPacienteEliminarPorPregunta"
   Set oParameter = .CreateParameter("@IdGrupoHCPaciente", adInteger, adParamInput, 0, oTabla.IdGrupoHCPaciente)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente)
    .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
EliminarRespuestaPorPregunta = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

