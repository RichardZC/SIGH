VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CajaGestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla CajaGestion
'        Programado por: Castro W
'        Fecha: Agosto 2006
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
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla CajaGestion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCajaGestion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaGestionAgregar"
           Set oParameter = .CreateParameter("@TotalCobrado", adChar, adParamInput, 10, IIf(oTabla.TotalCobrado = "", Null, oTabla.TotalCobrado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCierre", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaCierre = 0, Null, oTabla.FechaCierre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IIf(oTabla.IdCaja = 0, Null, oTabla.IdCaja)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EstadoLote", adChar, adParamInput, 1, IIf(oTabla.EstadoLote = "", Null, oTabla.EstadoLote)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaApertura", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaApertura = 0, Null, oTabla.FechaApertura)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGestionCaja", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdGestionCaja = .Parameters("@IdGestionCaja")
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
'   Descripción:    modifica un registro a la tabla CajaGestion
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCajaGestion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaGestionModificar"
           Set oParameter = .CreateParameter("@TotalCobrado", adChar, adParamInput, 10, IIf(oTabla.TotalCobrado = "", Null, oTabla.TotalCobrado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCierre", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaCierre = 0, Null, oTabla.FechaCierre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IIf(oTabla.IdCaja = 0, Null, oTabla.IdCaja)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EstadoLote", adChar, adParamInput, 1, IIf(oTabla.EstadoLote = "", Null, oTabla.EstadoLote)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaApertura", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaApertura = 0, Null, oTabla.FechaApertura)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGestionCaja", adInteger, adParamInput, 0, IIf(oTabla.IdGestionCaja = 0, Null, oTabla.IdGestionCaja)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CajaGestion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCajaGestion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaGestionEliminar"
           Set oParameter = .CreateParameter("@IdGestionCaja", adInteger, adParamInput, 0, IIf(oTabla.IdGestionCaja = 0, Null, oTabla.IdGestionCaja)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla CajaGestion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCajaGestion) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CajaGestionSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdGestionCaja", adInteger, adParamInput, 0, oTabla.IdGestionCaja): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.TotalCobrado = IIf(IsNull(oRecordset!TotalCobrado), "", oRecordset!TotalCobrado)
       oTabla.FechaCierre = IIf(IsNull(oRecordset!FechaCierre), "", oRecordset!FechaCierre)
       oTabla.IdTurno = IIf(IsNull(oRecordset!IdTurno), 0, oRecordset!IdTurno)
       oTabla.IdCaja = IIf(IsNull(oRecordset!IdCaja), 0, oRecordset!IdCaja)
       oTabla.IdCajero = IIf(IsNull(oRecordset!IdCajero), 0, oRecordset!IdCajero)
       oTabla.EstadoLote = IIf(IsNull(oRecordset!EstadoLote), "", oRecordset!EstadoLote)
       oTabla.FechaApertura = IIf(IsNull(oRecordset!FechaApertura), "", oRecordset!FechaApertura)
       oTabla.IdGestionCaja = IIf(IsNull(oRecordset!IdGestionCaja), 0, oRecordset!IdGestionCaja)
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function




Function FiltrarCajaGestion(ByVal oDOCajaGestion As DOCajaGestion, oDOCajaComprobantesPago As DOCajaComprobantesPago, oDOPaciente As DoPaciente, lcFechaComprobante As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
Dim sSQL1 As String, sSQL2 As String
    sSql = ""
    sWhere = ""
    sSQL1 = ""
    sSQL2 = ""
    If oDOCajaGestion.IdCaja <> 0 Then
        sWhere = sWhere + " and CajaGestion.IdCaja = " & oDOCajaGestion.IdCaja
    End If
    If oDOCajaGestion.IdTurno <> 0 Then
        sWhere = sWhere + " and CajaGestion.IdTurno = " & oDOCajaGestion.IdTurno
    End If
    If oDOCajaComprobantesPago.NroDocumento <> "" Then
        sWhere = sWhere + " and CajaComprobantesPago.NroDocumento = " & oDOCajaComprobantesPago.NroDocumento
    End If
    If oDOCajaComprobantesPago.NroSerie <> "" Then
        sWhere = sWhere + " and CajaComprobantesPago.NroSerie = " & oDOCajaComprobantesPago.NroSerie
    End If
    If oDOPaciente.NroHistoriaClinica <> 0 Then
        sWhere = sWhere + " and Pacientes.NroHistoriaClinica = " & oDOPaciente.NroHistoriaClinica
    End If
    If UCase(lcFechaComprobante) <> "TODAS" Then
        sWhere = sWhere + " and CajaComprobantesPago.FechaCobranza Between (CONVERT(DATETIME,'" & lcFechaComprobante & " 00:00:00',103)) and (CONVERT(DATETIME,'" & lcFechaComprobante & " 23:59:59',103))"
    End If
    If oDOCajaGestion.IdCajero <> 0 Then
        sWhere = sWhere + " and CajaGestion.idCajero = " & oDOCajaGestion.IdCajero
    End If
    If sWhere <> "" Then sSQL1 = sSQL1 & " where " & Mid(sWhere, 5)
    If sWhere <> "" Then sSQL2 = sSQL2 & " where " & Mid(sWhere, 5)
    sSQL1 = sSQL1 & " order by CajaGestion.FechaApertura desc"
    sSQL2 = sSQL2 & " order by CajaComprobantesPago.IdComprobantePago desc"
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CajaGestionFiltrar"
        Set oParameter = .CreateParameter("@lcFiltro1", adVarChar, adParamInput, 2000, sSQL1): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcFiltro2", adVarChar, adParamInput, 2000, sSQL2): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FiltrarCajaGestion = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



Function CajaGestionPorCajeroCajaYTurno(IdCaja As Long, IdCajero As Long, IdTurno As Long) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CajaGestionPorCajeroCajaYTurno"
        Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, IdCajero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IdCaja): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IdTurno): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaGestionPorCajeroCajaYTurno = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
Exit Function
End Function

Function TotalPorGestion(IdGestionCaja As Long) As Double

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantesPagoTotalPorGestion"
        Set oParameter = .CreateParameter("@IdGestionCaja", adInteger, adParamInput, 0, IdGestionCaja): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   TotalPorGestion = 0
   If oRecordset.Fields("total").Value <> Null Then
        TotalPorGestion = Val(oRecordset.Fields("total").Value)
   End If
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



