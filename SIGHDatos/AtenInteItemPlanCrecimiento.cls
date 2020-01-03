VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtenInteItemPlanCrecimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtenInteItemPlanCrecimiento
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
'   Descripción:    Inserta un registro a la tabla AtenInteItemPlanCrecimiento
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtenInteItemPlanCrecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteItemPlanCrecimientoAgregar"
   Set oParameter = .CreateParameter("@IdItemPlanCrecimiento", adBigInt, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdPlanAtencion = 0, Null, oTabla.IdPlanAtencion))
   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@IdTriajeVariable", adUnsignedTinyInt, adParamInput, 0, oTabla.IdTriajeVariable)
'   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NumeroSesion", adInteger, adParamInput, 0, IIf(oTabla.NumeroSesion = 0, Null, oTabla.NumeroSesion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdItemPlanCrecimiento = .Parameters("@IdItemPlanCrecimiento")
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
'   Descripción:    modifica un registro a la tabla AtenInteItemPlanCrecimiento
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtenInteItemPlanCrecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteItemPlanCrecimientoModificar"
   Set oParameter = .CreateParameter("@IdItemPlanCrecimiento", adBigInt, adParamInput, 0, oTabla.IdItemPlanCrecimiento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPlanAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdPlanAtencion = 0, Null, oTabla.IdPlanAtencion))
   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@IdTriajeVariable", adUnsignedTinyInt, adParamInput, 0, oTabla.IdTriajeVariable)
'   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NumeroSesion", adInteger, adParamInput, 0, IIf(oTabla.NumeroSesion = 0, Null, oTabla.NumeroSesion))
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
'   Descripción:    elimia o borra un registro a la tabla AtenInteItemPlanCrecimiento
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtenInteItemPlanCrecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteItemPlanCrecimientoEliminar"
   Set oParameter = .CreateParameter("@IdItemPlanCrecimiento", adBigInt, adParamInput, 0, oTabla.IdItemPlanCrecimiento)
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
'   Descripción:    Selecciona un unico registro de la tabla AtenInteItemPlanCrecimiento
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtenInteItemPlanCrecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtenInteItemPlanCrecimientoSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdItemPlanCrecimiento", adBigInt, adParamInput, 0, oTabla.IdItemPlanCrecimiento)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdItemPlanCrecimiento = oRecordset!IdItemPlanCrecimiento
       oTabla.IdPlanAtencion = IIf(IsNull(oRecordset!IdPlanAtencion), 0, oRecordset!IdPlanAtencion)
'       oTabla.IdTriajeVariable = oRecordset!IdTriajeVariable
       oTabla.NumeroSesion = IIf(IsNull(oRecordset!NumeroSesion), 0, oRecordset!NumeroSesion)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOAtenInteItemPlanCrecimiento)
   oTabla.IdItemPlanCrecimiento = 0
   oTabla.IdPlanAtencion = 0
'   oTabla.IdTriajeVariable = 0
   oTabla.NumeroSesion = 0
End Sub

