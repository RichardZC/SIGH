VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HIS_situacio"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla HIS_situacio
'        Programado por: Cachay F
'        Fecha: Noviembre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String


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
'   Descripción:    Selecciona un unico registro de la tabla HIS_situacio
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOHIS_situacio) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "HIS_situacioSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdHisSituacio", adInteger, adParamInput, 0, oTabla.IdHisSituacio)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdHisSituacio = IIf(IsNull(oRecordset!IdHisSituacio), 0, oRecordset!IdHisSituacio)
       oTabla.valores = IIf(IsNull(oRecordset!valores), "", oRecordset!valores)
       oTabla.descripcio = IIf(IsNull(oRecordset!descripcio), "", oRecordset!descripcio)
       oTabla.codigo = oRecordset!codigo
       oTabla.est = IIf(IsNull(oRecordset!est), "", oRecordset!est)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOHIS_situacio)
   oTabla.IdHisSituacio = 0
   oTabla.valores = ""
   oTabla.descripcio = ""
   oTabla.codigo = 0
   oTabla.est = ""
End Sub

Function ObtenerListaCodigosLAB() As Recordset
'On Error GoTo ManejadorDeError
'Dim oRecordset As New ADODB.Recordset
'Dim oCommand As New ADODB.Command
'Dim sSql As String
'sSql = "SELECT [IdHisSituacio], [valores], [descripcio], [codigo], [est] FROM [HIS_situacio] ORDER BY 3"
'
'With oCommand
'    .CommandType = adCmdText
'    Set .ActiveConnection = mo_Conexion
'    .CommandText = sSql
'    Set oRecordset = .Execute
'    Set oRecordset.ActiveConnection = Nothing
'End With
'
'Set ObtenerListaCodigosLAB = oRecordset
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
        .CommandText = "HIS_situacioSeleccionarPorCodigoDescripcion"
        
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerListaCodigosLAB = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
'Modificado por YEPE _ NOVIEMBRE
Function ObtenerListaCodigosLABporCodigoyNombre(ByVal valores As String, ByVal descripcio As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_situacioSeleccionarPorCodigoDescripcion"
        Set oParameter = .CreateParameter("@Valores", adChar, adParamInput, 5, valores): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Descripcio", adVarChar, adParamInput, 40, descripcio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerListaCodigosLABporCodigoyNombre = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

