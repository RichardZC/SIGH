VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HIS_DetalleDiagnostico"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla HIS_DetalleDiagnostico
'        Programado por: Cachay F
'        Fecha: Noviembre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ml_IdHisDetalleDiagnostico As Long
Dim ml_IdHisDetalle As Long
Dim ml_IdCIE As Long
Dim ml_IdSubClasificacionDX As Long
Dim ms_CodLAB As String

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
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla HIS_DetalleDiagnostico
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOHIS_DetalleDiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_DetalleDiagnosticoAgregar"
   Set oParameter = .CreateParameter("@IdHisDetalleDiagnostico", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdHisDetalle", adInteger, adParamInput, 0, IIf(oTabla.IdHisDetalle = 0, Null, oTabla.IdHisDetalle))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdCIE", adInteger, adParamInput, 0, IIf(oTabla.IdCIE = 0, Null, oTabla.IdCIE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdSubClasificacionDX", adInteger, adParamInput, 0, IIf(oTabla.IdSubClasificacionDX = 0, Null, oTabla.IdSubClasificacionDX))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodLAB", adVarChar, adParamInput, 30, IIf(oTabla.CodLAB = "", Null, oTabla.CodLAB))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdHisDetalleDiagnostico = .Parameters("@IdHisDetalleDiagnostico")
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
'   Descripción:    modifica un registro a la tabla HIS_DetalleDiagnostico
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOHIS_DetalleDiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_DetalleDiagnosticoModificar"
   Set oParameter = .CreateParameter("@IdHisDetalleDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdHisDetalleDiagnostico = 0, Null, oTabla.IdHisDetalleDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdHisDetalle", adInteger, adParamInput, 0, IIf(oTabla.IdHisDetalle = 0, Null, oTabla.IdHisDetalle))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdCIE", adInteger, adParamInput, 0, IIf(oTabla.IdCIE = 0, Null, oTabla.IdCIE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdSubClasificacionDX", adInteger, adParamInput, 0, IIf(oTabla.IdSubClasificacionDX = 0, Null, oTabla.IdSubClasificacionDX))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodLAB", adVarChar, adParamInput, 30, IIf(oTabla.CodLAB = "", Null, oTabla.CodLAB))
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
'   Descripción:    elimia o borra un registro a la tabla HIS_DetalleDiagnostico
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal lnIdHisDetalle As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_DetalleDiagnosticoEliminar"
   Set oParameter = .CreateParameter("@IdHisDetalleDiagnostico", adInteger, adParamInput, 0, lnIdHisDetalle)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 0)
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
'   Descripción:    Selecciona un unico registro de la tabla HIS_DetalleDiagnostico
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOHIS_DetalleDiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "HIS_DetalleDiagnosticoSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdHisDetalleDiagnostico", adInteger, adParamInput, 0, oTabla.IdHisDetalleDiagnostico)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdHisDetalleDiagnostico = IIf(IsNull(oRecordset!IdHisDetalleDiagnostico), 0, oRecordset!IdHisDetalleDiagnostico)
       oTabla.IdHisDetalle = IIf(IsNull(oRecordset!IdHisDetalle), 0, oRecordset!IdHisDetalle)
       oTabla.IdCIE = IIf(IsNull(oRecordset!IdCIE), 0, oRecordset!IdCIE)
       oTabla.IdSubClasificacionDX = IIf(IsNull(oRecordset!IdSubClasificacionDX), 0, oRecordset!IdSubClasificacionDX)
       oTabla.CodLAB = IIf(IsNull(oRecordset!CodLAB), "", oRecordset!CodLAB)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOHIS_DetalleDiagnostico)
   oTabla.IdHisDetalleDiagnostico = 0
   oTabla.IdHisDetalle = 0
   oTabla.IdCIE = 0
   oTabla.IdSubClasificacionDX = 0
   oTabla.CodLAB = ""
End Sub

Function ListaIdsDiagnosticosPorIdCabecera(RegCabeceraHIS) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HisDetalleDiagnosticoListaIdsDiagnosticosPorIdCabecera"
        Set oParameter = .CreateParameter("@IdHisCabecera", adInteger, adParamInput, 0, RegCabeceraHIS.IdHisCabecera): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ListaIdsDiagnosticosPorIdCabecera = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Public Function ObtenerDatosDetalleDiagnostico(ml_IdCabeceraHIS As Long) As Recordset
'On Error GoTo ManejadorDeError
'Dim oRecordSet As New ADODB.Recordset
'Dim oCommand As New ADODB.Command
'Dim sSql As String
'
'sSql = "SELECT"
'sSql = sSql & " HD.IdHisDetalleDiagnostico,"
'sSql = sSql & " HD.IdHisDetalle,"
'sSql = sSql & " HD.IdCIE,"
'sSql = sSql & " ('(' + D.CodigoCIE10 + ') ' + D.Descripcion) AS DESCRIPCION_CIE,"
'sSql = sSql & " HD.IdSubClasificacionDX,"
'sSql = sSql & " HD.CodLAB,"
'sSql = sSql & " '' AS MSG_ALERTA,"
'sSql = sSql & " 0 [IdEstado]"
'sSql = sSql & " From HIS_DetalleDiagnostico HD"
'sSql = sSql & " inner join Diagnosticos D on HD.IdCIE = D.IdDiagnostico"
'sSql = sSql & " WHERE  HD.IdHisDetalle IN (SELECT [IdHISDetalle]"
'                sSql = sSql & " From [HIS_Detalle]"
'                sSql = sSql & " WHERE  [IdHisCabecera] = " & ml_IdCabeceraHIS & ")"
'
'With oCommand
'    .CommandType = adCmdText
'    Set .ActiveConnection = mo_Conexion
'    .CommandText = sSql
'    Set oRecordSet = .Execute
'    Set oRecordSet.ActiveConnection = Nothing
'End With
'
'Set ObtenerDatosDetalleDiagnostico = oRecordSet
'ms_MensajeError = ""
'Exit Function
'
'ManejadorDeError:
'ms_MensajeError = Err.Number & " " + Err.Description
Stop
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_DetalleDiagnosticoObtenerDatosDetalleDiagnostico"
        Set oParameter = .CreateParameter("@ml_IdCabeceraHIS", adInteger, adParamInput, 0, ml_IdCabeceraHIS): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerDatosDetalleDiagnostico = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Public Function ObtenerDatosDetalleDiagnosticoPorIdDetalle(ml_IdDetalleHIS As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_DetalleDiagnosticoObtenerDatosDetalleDiagnosticoPorIdDetalle"
        Set oParameter = .CreateParameter("@ml_IdDetalleHIS", adInteger, adParamInput, 0, ml_IdDetalleHIS): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerDatosDetalleDiagnosticoPorIdDetalle = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Public Function ExportacionHIS_Diagnosticos(IdUsuario As Long, ml_Mes As Long, mi_anio As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String

sSql = ""

With oCommand
    .CommandType = adCmdText
    Set .ActiveConnection = mo_Conexion
    .CommandText = sSql
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
End With

Set ExportacionHIS_Diagnosticos = oRecordset
ms_MensajeError = ""
Exit Function

ManejadorDeError:
ms_MensajeError = Err.Number & " " + Err.Description
End Function


