VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FactOrdenesBienesInsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FactOrdenesBienesInsumo
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
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
'   Descripción:    Inserta un registro a la tabla FactOrdenesBienesInsumo
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOFactOrdenBienInsumo) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter

   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FactOrdenesBienesInsumoAgregar"
           Set oParameter = .CreateParameter("@FechaModificacion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaModificacion = 0, Null, oTabla.FechaModificacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioModifica", adInteger, adParamInput, 0, IIf(oTabla.idUsuarioModifica = 0, Null, oTabla.idUsuarioModifica)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioCrea", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioCrea = 0, Null, oTabla.IdUsuarioCrea)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaOrden", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaOrden = 0, Null, oTabla.FechaOrden)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoOrden", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoOrden = 0, Null, oTabla.IdEstadoOrden)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.idPuntoCarga = 0, Null, oTabla.idPuntoCarga)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
           
           Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idFormaPago", adInteger, adParamInput, 0, IIf(oTabla.idFormaPago = 0, Null, oTabla.idFormaPago)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idFarmacia", adInteger, adParamInput, 0, IIf(oTabla.idFarmacia = 0, Null, oTabla.idFarmacia)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter

       .Execute
           oTabla.IdOrden = .Parameters("@IdOrden")
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
'   Descripción:    modifica un registro a la tabla FactOrdenesBienesInsumo
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFactOrdenBienInsumo) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FactOrdenesBienesInsumoModificar"
           Set oParameter = .CreateParameter("@FechaModificacion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaModificacion = 0, Null, oTabla.FechaModificacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioModifica", adInteger, adParamInput, 0, IIf(oTabla.idUsuarioModifica = 0, Null, oTabla.idUsuarioModifica)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioCrea", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioCrea = 0, Null, oTabla.IdUsuarioCrea)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaOrden", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaOrden = 0, Null, oTabla.FechaOrden)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoOrden", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoOrden = 0, Null, oTabla.IdEstadoOrden)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.idPuntoCarga = 0, Null, oTabla.idPuntoCarga)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       
           Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFormaPago", adInteger, adParamInput, 0, IIf(oTabla.idFormaPago = 0, Null, oTabla.idFormaPago)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idFarmacia", adInteger, adParamInput, 0, IIf(oTabla.idFarmacia = 0, Null, oTabla.idFarmacia)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter

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
'   Descripción:    elimia o borra un registro a la tabla FactOrdenesBienesInsumo
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFactOrdenBienInsumo) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FactOrdenesBienesInsumoEliminar"
           Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla FactOrdenesBienesInsumo
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFactOrdenBienInsumo) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandTimeout = 150
     .CommandText = "FactOrdenesBienesInsumoSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, oTabla.IdOrden): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.FechaModificacion = IIf(IsNull(oRecordset!FechaModificacion), 0, oRecordset!FechaModificacion)
       oTabla.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
       oTabla.idUsuarioModifica = IIf(IsNull(oRecordset!idUsuarioModifica), 0, oRecordset!idUsuarioModifica)
       oTabla.IdUsuarioCrea = IIf(IsNull(oRecordset!IdUsuarioCrea), 0, oRecordset!IdUsuarioCrea)
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       If Not IsNull(oRecordset!FechaOrden) Then
          oTabla.FechaOrden = IIf(IsNull(oRecordset!FechaOrden), 0, oRecordset!FechaOrden)
       End If
       oTabla.IdOrden = IIf(IsNull(oRecordset!IdOrden), 0, oRecordset!IdOrden)
       oTabla.IdEstadoOrden = IIf(IsNull(oRecordset!IdEstadoOrden), 0, oRecordset!IdEstadoOrden)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.idPuntoCarga = IIf(IsNull(oRecordset!idPuntoCarga), 0, oRecordset!idPuntoCarga)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.idFormaPago = IIf(IsNull(oRecordset!idFormaPago), 0, oRecordset!idFormaPago)
       oTabla.idFarmacia = IIf(IsNull(oRecordset!idFarmacia), 0, oRecordset!idFarmacia)
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

Function Filtrar(oDOFactOrdenBienInsumo As DOFactOrdenBienInsumo, oDOPaciente As DoPaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    If oDOFactOrdenBienInsumo.IdOrden <> 0 Then
         sWhere = sWhere + " FactOrdenesBienesInsumo.IdOrden = " & oDOFactOrdenBienInsumo.IdOrden & " and "
    End If
    If oDOFactOrdenBienInsumo.idPuntoCarga <> 0 Then
         sWhere = sWhere + " FactOrdenesBienesInsumo.IdPuntoCarga = " & oDOFactOrdenBienInsumo.idPuntoCarga & " and "
    End If
    If oDOPaciente.NroHistoriaClinica <> 0 Then
         sWhere = sWhere + " Pacientes.NroHistoriaClinica = " & oDOPaciente.NroHistoriaClinica & " and "
    End If
    If sWhere <> "" Then
         sSql = sSql + " where " + Left(sWhere, Len(sWhere) - 4)
    End If
    sSql = sSql + " order by  FactOrdenesBienesInsumo.IdOrden desc, FactOrdenesBienesInsumo.FechaOrden desc "
   
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "EstablecimientosNoMinsaFiltrar"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set Filtrar = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function FiltrarDEBB(oDOFactOrdenBienInsumo As DOFactOrdenBienInsumo, oDOPaciente As DoPaciente, ldFechaInicial As Date, ldFechaFinal As Date, lnFarmacia As Integer) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    If oDOFactOrdenBienInsumo.IdOrden <> 0 Then
         sWhere = sWhere + " FactOrdenesBienesInsumo.IdOrden = " & oDOFactOrdenBienInsumo.IdOrden & " and "
    End If
    If oDOFactOrdenBienInsumo.idPuntoCarga <> 0 Then
         sWhere = sWhere + " FactOrdenesBienesInsumo.IdPuntoCarga = " & oDOFactOrdenBienInsumo.idPuntoCarga & " and "
    End If
    If oDOPaciente.NroHistoriaClinica <> 0 Then
         sWhere = sWhere + " Pacientes.NroHistoriaClinica = " & oDOPaciente.NroHistoriaClinica & " and "
    End If
    If lnFarmacia > 0 Then
        sWhere = sWhere + " FactOrdenesBienesInsumo.idFarmacia=" & lnFarmacia & " And "
    Else
        sWhere = sWhere + " FactOrdenesBienesInsumo.idFarmacia is null And "
    End If
    sWhere = sWhere + " FactOrdenesBienesInsumo.FechaOrden Between (CONVERT(DATETIME,'" & ldFechaInicial & "',103)) and (CONVERT(DATETIME,'" & ldFechaFinal & "',103))" & " and "
    If sWhere <> "" Then
         sSql = sSql + " where " + Left(sWhere, Len(sWhere) - 4)
    End If
    sSql = sSql + " order by  FactOrdenesBienesInsumo.IdOrden desc, FactOrdenesBienesInsumo.FechaOrden desc "
   
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenesBienesInsumoFiltrado"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FiltrarDEBB = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function EliminarPorIdAtencion(lIdAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   EliminarPorIdAtencion = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FactOrdenesBienesInsumoEliminarV2"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(lIdAtencion = 0, Null, lIdAtencion)): .Parameters.Append oParameter
       .Execute
   End With
 
   EliminarPorIdAtencion = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
    Trace Err.Number & " " + Err.Description, "EliminarPorIdAtencion"
Exit Function
End Function

Function EliminarPorIdOrden(lIdOrden As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   EliminarPorIdOrden = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FactOrdenesBienesInsumoEliminarV3"
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, IIf(lIdOrden = 0, Null, lIdOrden)): .Parameters.Append oParameter
       .Execute
   End With
 
   EliminarPorIdOrden = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
    Trace Err.Number & " " + Err.Description, "EliminarPorIdOrden"
Exit Function
End Function

Sub Trace(sMensaje As String, NombreDeMetodo As String)
    MsgBox "Mensaje: " & sMensaje + Chr(13) + "Metodo:" & NombreDeMetodo + Chr(13) + "Clase de Datos: FactOrdenesServicio", vbInformation, "FactOrdenesServicio"
End Sub
Function SeleccionarPorCuentaAtencion(IdCuentaAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenesBienesInsumoXidCuentaAtencion"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorCuentaAtencion = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description

End Function


Function ActualizarPagoDeOrdenesProcesadas(sOrdenesProcesadas As String, IdComprobantePago As Long, lIdUsuario As Long) As Boolean


End Function

Function SeleccionarPorIdComprobante(lIdComprobantePago As Long) As DOFactOrdenBienInsumo
On Error GoTo ManejadorDeError
Dim rs As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oDOFactOrdenBienInsumo As New DOFactOrdenBienInsumo
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenesBienesInsumoXidComprobantePago"
        Set oParameter = .CreateParameter("@lIdComprobantePago", adInteger, adParamInput, 0, lIdComprobantePago): .Parameters.Append oParameter
        Set rs = .Execute
        Set rs.ActiveConnection = Nothing
   End With
   If Not (rs.EOF And rs.BOF) Then
        oDOFactOrdenBienInsumo.IdOrden = rs!IdOrden
        If SeleccionarPorId(oDOFactOrdenBienInsumo) Then
            Set SeleccionarPorIdComprobante = oDOFactOrdenBienInsumo
        End If
   End If
   Set oCommand = Nothing
   Set rs = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
    
    

End Function

Function SeleccionarPorIdComprobanteDevolucion(lIdComprobantePago As Long) As DOFactOrdenBienInsumo

    
End Function
