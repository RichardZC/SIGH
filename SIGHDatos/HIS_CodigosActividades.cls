VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HIS_CodigosActividades"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla HIS_CodigosActividades
'        Programado por: Cachay F
'        Fecha: Noviembre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ml_IdHisCodActvidad As Long
Dim ml_IdTipoAtencion As Long
Dim ms_CodigoActividad As String
Dim ms_Descripcion As String

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla HIS_CodigosActividades
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOHIS_CodigosActividades) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "HIS_CodigosActividadesSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdHisCodActvidad", adInteger, adParamInput, 0, oTabla.IdHisCodActvidad)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdHisCodActvidad = IIf(IsNull(oRecordset!IdHisCodActvidad), 0, oRecordset!IdHisCodActvidad)
       oTabla.IdTipoAtencion = IIf(IsNull(oRecordset!IdTipoAtencion), 0, oRecordset!IdTipoAtencion)
       oTabla.CodigoActividad = IIf(IsNull(oRecordset!CodigoActividad), "", oRecordset!CodigoActividad)
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub SetDefaults(ByVal oTabla As DOHIS_CodigosActividades)
   oTabla.IdHisCodActvidad = 0
   oTabla.IdTipoAtencion = 0
   oTabla.CodigoActividad = ""
   oTabla.Descripcion = ""
End Sub

Function ObtenerListaCodigosActividades(ml_CodigoTipoActividad As Integer) As ADODB.Recordset
'On Error GoTo ManejadorDeError
'Dim oRecordSet As New ADODB.Recordset
'Dim oCommand As New ADODB.Command
'Dim sSql As String
'
'sSql = "SELECT [IdHisCodActvidad], [IdTipoAtencion], [CodigoActividad], [Descripcion] "
'sSql = sSql & " From [sigh].[dbo].[HIS_CodigosActividades]"
'If ml_CodigoTipoActividad <> 0 Then
'    sSql = sSql & " WHERE IdTipoAtencion = " & ml_CodigoTipoActividad
'End If
'With oCommand
'    .CommandType = adCmdText
'    Set .ActiveConnection = mo_Conexion
'    .CommandText = sSql
'    Set oRecordSet = .Execute
'    Set oRecordSet.ActiveConnection = Nothing
'End With
'
'Set ObtenerListaCodigosActividades = oRecordSet
'ms_MensajeError = ""
'Exit Function
'ManejadorDeError:
'       ms_MensajeError = Err.Number & " " + Err.Description

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_CodigosActividadesObtenerListaCodigosActividades"
        Set oParameter = .CreateParameter("@ml_CodigoTipoActividad", adInteger, adParamInput, 0, ml_CodigoTipoActividad): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerListaCodigosActividades = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'Modificado por YEPE _ NOVIEMBRE
Function ObtenerListaCodigosActividadesporCodigoyNombre(ByVal CodigoActividad As String, ByVal Descripcion As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_CodigosActividadesSeleccionarPorCodigoDescripcion"
        Set oParameter = .CreateParameter("@CodigoActividad", adVarChar, adParamInput, 6, CodigoActividad): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 100, Descripcion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerListaCodigosActividadesporCodigoyNombre = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

