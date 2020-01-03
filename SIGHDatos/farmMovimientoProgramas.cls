VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "farmMovimientoProgramas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FarmMovimientoProgramas
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String

Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property


'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla farmMovimientoProgramas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOfarmMovimientoProgramas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmMovimientoProgramasAgregar"
   Set oParameter = .CreateParameter("@movNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@movTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCoordinador", adInteger, adParamInput, 0, IIf(oTabla.idCoordinador = 0, Null, oTabla.idCoordinador))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPrescriptor", adChar, adParamInput, 10, IIf(oTabla.idPrescriptor = "", Null, oTabla.idPrescriptor))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idComponente", adInteger, adParamInput, 0, IIf(oTabla.idComponente = 0, Null, oTabla.idComponente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idSubComponente", adInteger, adParamInput, 0, IIf(oTabla.idSubComponente = 0, Null, oTabla.idSubComponente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaHoraPrescribe", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaHoraPrescribe = 0, Null, oTabla.FechaHoraPrescribe))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
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
'   Descripción:    modifica un registro a la tabla farmMovimientoProgramas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOfarmMovimientoProgramas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmMovimientoProgramasModificar"
   Set oParameter = .CreateParameter("@movNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@movTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCoordinador", adInteger, adParamInput, 0, IIf(oTabla.idCoordinador = 0, Null, oTabla.idCoordinador))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPrescriptor", adChar, adParamInput, 10, IIf(oTabla.idPrescriptor = "", Null, oTabla.idPrescriptor))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idComponente", adInteger, adParamInput, 0, IIf(oTabla.idComponente = 0, Null, oTabla.idComponente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idSubComponente", adInteger, adParamInput, 0, IIf(oTabla.idSubComponente = 0, Null, oTabla.idSubComponente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaHoraPrescribe", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaHoraPrescribe = 0, Null, oTabla.FechaHoraPrescribe))
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
'   Descripción:    elimia o borra un registro a la tabla farmMovimientoProgramas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOfarmMovimientoProgramas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmMovimientoProgramasEliminar"
   Set oParameter = .CreateParameter("@movNumero", adVarChar, adParamInput, 9, oTabla.movNumero)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@movTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
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
'   Descripción:    Selecciona un unico registro de la tabla farmMovimientoProgramas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOfarmMovimientoProgramas) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "farmMovimientoProgramasSeleccionarPorId"
     Set oParameter = .CreateParameter("@movNumero", adVarChar, adParamInput, 9, oTabla.movNumero)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@movTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.movNumero = IIf(IsNull(oRecordset!movNumero), "", oRecordset!movNumero)
       oTabla.movTipo = IIf(IsNull(oRecordset!movTipo), "", oRecordset!movTipo)
       oTabla.idCoordinador = IIf(IsNull(oRecordset!idCoordinador), 0, oRecordset!idCoordinador)
       oTabla.idPrescriptor = IIf(IsNull(oRecordset!idPrescriptor), "", oRecordset!idPrescriptor)
       oTabla.IdDiagnostico = IIf(IsNull(oRecordset!IdDiagnostico), 0, oRecordset!IdDiagnostico)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.idComponente = IIf(IsNull(oRecordset!idComponente), 0, oRecordset!idComponente)
       oTabla.idSubComponente = IIf(IsNull(oRecordset!idSubComponente), 0, oRecordset!idSubComponente)
       oTabla.FechaHoraPrescribe = IIf(IsNull(oRecordset!FechaHoraPrescribe), 0, oRecordset!FechaHoraPrescribe)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOfarmMovimientoProgramas)
   oTabla.movNumero = ""
   oTabla.movTipo = ""
   oTabla.idCoordinador = 0
   oTabla.idPrescriptor = ""
   oTabla.IdDiagnostico = 0
   oTabla.IdPaciente = 0
   oTabla.idComponente = 0
   oTabla.idSubComponente = 0
End Sub
