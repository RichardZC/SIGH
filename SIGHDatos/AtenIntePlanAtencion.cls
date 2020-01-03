VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtenIntePlanAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtenIntePlanAtencion
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
'   Descripción:    Inserta un registro a la tabla AtenIntePlanAtencion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtenIntePlanAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanAtencionAgregar"
   Set oParameter = .CreateParameter("@IdPlanAtencion", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAtenInteGrupo", adUnsignedTinyInt, adParamInput, 0, oTabla.IdAtenInteGrupo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPeriodoTiempo", adUnsignedTinyInt, adParamInput, 0, oTabla.IdPeriodoTiempo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadAnio", adInteger, adParamInput, 0, IIf(oTabla.EdadAnio = 0, Null, oTabla.EdadAnio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadMes", adUnsignedTinyInt, adParamInput, 0, oTabla.EdadMes)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadDia", adUnsignedTinyInt, adParamInput, 0, oTabla.EdadDia)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 30, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdPlanAtencion = .Parameters("@IdPlanAtencion")
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
'   Descripción:    modifica un registro a la tabla AtenIntePlanAtencion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtenIntePlanAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanAtencionModificar"
   Set oParameter = .CreateParameter("@IdPlanAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdPlanAtencion = 0, Null, oTabla.IdPlanAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAtenInteGrupo", adUnsignedTinyInt, adParamInput, 0, oTabla.IdAtenInteGrupo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPeriodoTiempo", adUnsignedTinyInt, adParamInput, 0, oTabla.IdPeriodoTiempo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadAnio", adInteger, adParamInput, 0, IIf(oTabla.EdadAnio = 0, Null, oTabla.EdadAnio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadMes", adUnsignedTinyInt, adParamInput, 0, oTabla.EdadMes)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadDia", adUnsignedTinyInt, adParamInput, 0, oTabla.EdadDia)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 30, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion))
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
'   Descripción:    elimia o borra un registro a la tabla AtenIntePlanAtencion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtenIntePlanAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenIntePlanAtencionEliminar"
   Set oParameter = .CreateParameter("@IdPlanAtencion", adInteger, adParamInput, 0, oTabla.IdPlanAtencion)
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
'   Descripción:    Selecciona un unico registro de la tabla AtenIntePlanAtencion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtenIntePlanAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtenIntePlanAtencionSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdPlanAtencion", adInteger, adParamInput, 0, oTabla.IdPlanAtencion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdPlanAtencion = IIf(IsNull(oRecordset!IdPlanAtencion), 0, oRecordset!IdPlanAtencion)
       oTabla.IdAtenInteGrupo = oRecordset!IdAtenInteGrupo
       oTabla.IdPeriodoTiempo = oRecordset!IdPeriodoTiempo
       oTabla.EdadAnio = IIf(IsNull(oRecordset!EdadAnio), 0, oRecordset!EdadAnio)
       oTabla.EdadMes = oRecordset!EdadMes
       oTabla.EdadDia = oRecordset!EdadDia
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOAtenIntePlanAtencion)
   oTabla.IdPlanAtencion = 0
   oTabla.IdAtenInteGrupo = 0
   oTabla.IdPeriodoTiempo = 0
   oTabla.EdadAnio = 0
   oTabla.EdadMes = 0
   oTabla.EdadDia = 0
   oTabla.Descripcion = ""
End Sub


