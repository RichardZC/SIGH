VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HIS_DetalleDx_Verifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla HIS_DetalleDx_Verifica
'        Programado por: Cachay F
'        Fecha: Noviembre 2013
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
'   Descripción:    Inserta un registro a la tabla HIS_DetalleDx_Verifica
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOHIS_DetalleDx_Verifica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_DetalleDx_VerificaAgregar"
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
 
Insertar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla HIS_DetalleDx_Verifica
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOHIS_DetalleDx_Verifica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_DetalleDx_VerificaModificar"
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
'   Descripción:    elimia o borra un registro a la tabla HIS_DetalleDx_Verifica
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(lnIdHisDetalle As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_DetalleDx_VerificaEliminar"
   Set oParameter = .CreateParameter("@IdHisDetalle", adInteger, adParamInput, 0, lnIdHisDetalle)
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
'   Descripción:    Selecciona un unico registro de la tabla HIS_DetalleDx_Verifica
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOHIS_DetalleDx_Verifica) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "HIS_DetalleDx_VerificaSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdHisDetalle", adInteger, adParamInput, 0, oTabla.IdHisDetalle)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdHisDetalle = IIf(IsNull(oRecordset!IdHisDetalle), 0, oRecordset!IdHisDetalle)
       oTabla.IdCIE = IIf(IsNull(oRecordset!IdCIE), 0, oRecordset!IdCIE)
       oTabla.IdSubClasificacionDX = IIf(IsNull(oRecordset!IdSubClasificacionDX), 0, oRecordset!IdSubClasificacionDX)
       oTabla.CodLAB = IIf(IsNull(oRecordset!CodLAB), "", oRecordset!CodLAB)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Public Function His_ConsultaDxHisDetalleVerif(ml_IdDetalleHIS As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "His_ConsultaDxHisDetalleVerif"
        Set oParameter = .CreateParameter("@ml_IdDetalleHIS", adInteger, adParamInput, 0, ml_IdDetalleHIS): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set His_ConsultaDxHisDetalleVerif = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Sub SetDefaults(ByVal oTabla As DOHIS_DetalleDx_Verifica)
   oTabla.IdHisDetalle = 0
   oTabla.IdCIE = 0
   oTabla.IdSubClasificacionDX = 0
   oTabla.CodLAB = ""
End Sub
