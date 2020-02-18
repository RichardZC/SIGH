VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CajaComprobantesPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla CajaComprobantesPago
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
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
'   Descripción:    Inserta un registro a la tabla CajaComprobantesPago
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCajaComprobantesPago) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesPagoAgregar"
           Set oParameter = .CreateParameter("@TipoCambio", adCurrency, adParamInput, 0, IIf(oTabla.TipoCambio = 0, Null, oTabla.TipoCambio)): .Parameters.Append oParameter
           'Sunat Facturador
           Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, IIf(oTabla.NroSerie = "", Null, oTabla.NroSerie)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroDocumento", adChar, adParamInput, 8, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento)): .Parameters.Append oParameter
           'Sunat Facturador
           Set oParameter = .CreateParameter("@RazonSocial", adVarChar, adParamInput, 100, IIf(oTabla.RazonSocial = "", Null, oTabla.RazonSocial)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@RUC", adChar, adParamInput, 11, IIf(oTabla.RUC = "", Null, oTabla.RUC)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SubTotal", adCurrency, adParamInput, 0, oTabla.SubTotal): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IGV", adCurrency, adParamInput, 0, oTabla.IGV): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCobranza", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCobranza = 0, Null, oTabla.FechaCobranza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoOrden", adInteger, adParamInput, 0, IIf(oTabla.IdTipoOrden = 0, Null, oTabla.IdTipoOrden)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observaciones", adVarChar, adParamInput, 250, IIf(oTabla.Observaciones = "", Null, oTabla.Observaciones)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoComprobante", adInteger, adParamInput, 0, IIf(oTabla.IdTipoComprobante = 0, Null, oTabla.IdTipoComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoComprobante", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoComprobante = 0, Null, oTabla.IdEstadoComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGestionCaja", adInteger, adParamInput, 0, IIf(oTabla.IdGestionCaja = 0, Null, oTabla.IdGestionCaja)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoPago", adInteger, adParamInput, 0, IIf(oTabla.IdTipoPago = 0, Null, oTabla.IdTipoPago)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, oTabla.Total): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
           
           Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idFormaPago", adInteger, adParamInput, 0, IIf(oTabla.idFormaPago = 0, Null, oTabla.idFormaPago)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idFarmacia", adInteger, adParamInput, 0, IIf(oTabla.idFarmacia = 0, Null, oTabla.idFarmacia)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idCaja", adInteger, adParamInput, 0, IIf(oTabla.IdCaja = 0, Null, oTabla.IdCaja)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idCajero", adInteger, adParamInput, 0, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Exoneraciones", adCurrency, adParamInput, 0, oTabla.Exoneraciones): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Dctos", adCurrency, adParamInput, 0, oTabla.Dctos): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Adelantos", adCurrency, adParamInput, 0, oTabla.Adelantos): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@DNI", adVarChar, adParamInput, 11, oTabla.DNI): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TieneCredito", adVarChar, adParamInput, 11, IIf(oTabla.TieneCredito = "", Null, oTabla.TieneCredito)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@fechaEmision", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fechaEmision = 0, Null, oTabla.fechaEmision)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SunatTotalIgv", adCurrency, adParamInput, 0, oTabla.SunatTotalIgv): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SunatOpeInafectas", adCurrency, adParamInput, 0, oTabla.SunatOpeInafectas): .Parameters.Append oParameter
       .Execute
           oTabla.IdComprobantePago = .Parameters("@IdComprobantePago")
   End With
 
   Insertar = True
   ms_MensajeError = ""
   GrabaHistorico oTabla
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub GrabaHistorico(oTabla As DOCajaComprobantesPago)
    On Error GoTo ErrHistorico
    Dim lcSql As String
    Dim oRsParametros As New Recordset
    Dim oRsMDB As New Recordset
    Dim oConexionMDB As New Connection
    lcSql = "select ValorTexto from Parametros where idparametro=581"
    oRsParametros.Open lcSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    oConexionMDB.CommandTimeout = 900
    oConexionMDB.CursorLocation = adUseClient
    oConexionMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & Trim(oRsParametros!ValorTexto) & "\parametros.mdb"
    lcSql = "INSERT INTO CajaComprobantesPago" & _
"           (idComprobantePago,NroSerie,NroDocumento,RazonSocial,RUC" & _
"           ,SubTotal,IGV,Total,FechaCobranza" & _
"           ,TipoCambio,Observaciones,IdTipoComprobante" & _
"           ,IdCuentaAtencion,IdEstadoComprobante,IdGestionCaja,IdTipoPago" & _
"           ,IdTipoOrden,Dctos,IdPaciente,IdCajero" & _
"           ,idTurno,idCaja,idFormaPago,idFarmacia" & _
"           ,Exoneraciones,Adelantos,idTipoFinanciamiento)" & _
"     Values" & _
"           (" & _
"           " & oTabla.IdComprobantePago & ",'" & oTabla.NroSerie & "','" & oTabla.NroDocumento & "','" & oTabla.RazonSocial & "','" & Trim(oTabla.RUC) & _
"           '," & oTabla.SubTotal & "," & oTabla.IGV & "," & oTabla.Total & ",'" & oTabla.FechaCobranza & _
"           '," & oTabla.TipoCambio & ",'" & oTabla.Observaciones & "'," & oTabla.IdTipoComprobante & _
"           ," & oTabla.IdCuentaAtencion & "," & oTabla.IdEstadoComprobante & "," & oTabla.IdGestionCaja & "," & oTabla.IdTipoPago & _
"           ," & oTabla.IdTipoOrden & "," & oTabla.Dctos & "," & oTabla.IdPaciente & "," & oTabla.IdCajero & _
"           ," & oTabla.IdTurno & "," & oTabla.IdCaja & "," & oTabla.idFormaPago & "," & oTabla.idFarmacia & _
"           ," & oTabla.Exoneraciones & "," & oTabla.Adelantos & "," & oTabla.idTipoFinanciamiento & _
"           )"
   oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    oConexionMDB.Close
ErrHistorico:
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing
End Sub


'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla CajaComprobantesPago
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCajaComprobantesPago) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesPagoModificar"
           Set oParameter = .CreateParameter("@TipoCambio", adCurrency, adParamInput, 0, IIf(oTabla.TipoCambio = 0, Null, oTabla.TipoCambio)): .Parameters.Append oParameter
           'Sunat Facturador
           Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, IIf(oTabla.NroSerie = "", Null, oTabla.NroSerie)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento)): .Parameters.Append oParameter
           'Sunat Facturador
           Set oParameter = .CreateParameter("@RazonSocial", adVarChar, adParamInput, 100, IIf(oTabla.RazonSocial = "", Null, oTabla.RazonSocial)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@RUC", adChar, adParamInput, 11, IIf(oTabla.RUC = "", Null, oTabla.RUC)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SubTotal", adCurrency, adParamInput, 0, IIf(oTabla.SubTotal = 0, Null, oTabla.SubTotal)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IGV", adCurrency, adParamInput, 0, oTabla.IGV): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCobranza", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCobranza = 0, Null, oTabla.FechaCobranza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoOrden", adInteger, adParamInput, 0, IIf(oTabla.IdTipoOrden = 0, Null, oTabla.IdTipoOrden)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observaciones", adVarChar, adParamInput, 250, IIf(oTabla.Observaciones = "", Null, oTabla.Observaciones)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoComprobante", adInteger, adParamInput, 0, IIf(oTabla.IdTipoComprobante = 0, Null, oTabla.IdTipoComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoComprobante", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoComprobante = 0, Null, oTabla.IdEstadoComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGestionCaja", adInteger, adParamInput, 0, IIf(oTabla.IdGestionCaja = 0, Null, oTabla.IdGestionCaja)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoPago", adInteger, adParamInput, 0, IIf(oTabla.IdTipoPago = 0, Null, oTabla.IdTipoPago)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, oTabla.Total): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
           
           Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idFormaPago", adInteger, adParamInput, 0, IIf(oTabla.idFormaPago = 0, Null, oTabla.idFormaPago)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idFarmacia", adInteger, adParamInput, 0, IIf(oTabla.idFarmacia = 0, Null, oTabla.idFarmacia)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idCaja", adInteger, adParamInput, 0, IIf(oTabla.IdCaja = 0, Null, oTabla.IdCaja)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idCajero", adInteger, adParamInput, 0, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@DNI", adVarChar, adParamInput, 11, oTabla.DNI): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TieneCredito", adVarChar, adParamInput, 11, IIf(oTabla.TieneCredito = "", Null, oTabla.TieneCredito)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@fechaEmision", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fechaEmision = 0, Null, oTabla.fechaEmision)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SunatTotalIgv", adCurrency, adParamInput, 0, oTabla.SunatTotalIgv): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SunatOpeInafectas", adCurrency, adParamInput, 0, oTabla.SunatOpeInafectas): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CajaComprobantesPago
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCajaComprobantesPago) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesPagoEliminar"
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
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

Sub ActualizaTabla(ByRef oTabla As DOCajaComprobantesPago, oRecordset As Recordset)
       oTabla.TipoCambio = IIf(IsNull(oRecordset!TipoCambio), 0, oRecordset!TipoCambio)
       oTabla.NroSerie = IIf(IsNull(oRecordset!NroSerie), "", oRecordset!NroSerie)
       oTabla.NroDocumento = IIf(IsNull(oRecordset!NroDocumento), "", oRecordset!NroDocumento)
       oTabla.RazonSocial = IIf(IsNull(oRecordset!RazonSocial), "", oRecordset!RazonSocial)
       oTabla.RUC = IIf(IsNull(oRecordset!RUC), "", oRecordset!RUC)
       oTabla.SubTotal = IIf(IsNull(oRecordset!SubTotal), 0, oRecordset!SubTotal)
       oTabla.IGV = IIf(IsNull(oRecordset!IGV), 0, oRecordset!IGV)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.FechaCobranza = IIf(IsNull(oRecordset!FechaCobranza), 0, oRecordset!FechaCobranza)
       oTabla.IdTipoOrden = IIf(IsNull(oRecordset!IdTipoOrden), 0, oRecordset!IdTipoOrden)
       oTabla.Observaciones = IIf(IsNull(oRecordset!Observaciones), "", oRecordset!Observaciones)
       oTabla.IdTipoComprobante = IIf(IsNull(oRecordset!IdTipoComprobante), 0, oRecordset!IdTipoComprobante)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.IdEstadoComprobante = IIf(IsNull(oRecordset!IdEstadoComprobante), 0, oRecordset!IdEstadoComprobante)
       oTabla.IdGestionCaja = IIf(IsNull(oRecordset!IdGestionCaja), 0, oRecordset!IdGestionCaja)
       oTabla.IdTipoPago = IIf(IsNull(oRecordset!IdTipoPago), 0, oRecordset!IdTipoPago)
       oTabla.Total = IIf(IsNull(oRecordset!Total), 0, oRecordset!Total)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.idFormaPago = IIf(IsNull(oRecordset!idFormaPago), 0, oRecordset!idFormaPago)
       oTabla.idFarmacia = IIf(IsNull(oRecordset!idFarmacia), 0, oRecordset!idFarmacia)
       oTabla.IdCaja = IIf(IsNull(oRecordset!IdCaja), 0, oRecordset!IdCaja)
       oTabla.IdTurno = IIf(IsNull(oRecordset!IdTurno), 0, oRecordset!IdTurno)
       oTabla.IdCajero = IIf(IsNull(oRecordset!IdCajero), 0, oRecordset!IdCajero)
       oTabla.Dctos = IIf(IsNull(oRecordset!Dctos), 0, oRecordset!Dctos)
       oTabla.Exoneraciones = IIf(IsNull(oRecordset!Exoneraciones), 0, oRecordset!Exoneraciones)
       oTabla.Adelantos = IIf(IsNull(oRecordset!Adelantos), 0, oRecordset!Adelantos)
       oTabla.idTipoFinanciamiento = IIf(IsNull(oRecordset!idTipoFinanciamiento), 0, oRecordset!idTipoFinanciamiento)
       oTabla.DNI = IIf(IsNull(oRecordset!DNI), "", oRecordset!DNI)
       oTabla.TieneCredito = IIf(IsNull(oRecordset!TieneCredito), "", oRecordset!TieneCredito)
       oTabla.fechaEmision = IIf(IsNull(oRecordset!fechaEmision), 0, oRecordset!fechaEmision)
       oTabla.SunatTotalIgv = IIf(IsNull(oRecordset!SunatTotalIgv), 0, oRecordset!SunatTotalIgv)
       oTabla.SunatOpeInafectas = IIf(IsNull(oRecordset!SunatOpeInafectas), 0, oRecordset!SunatOpeInafectas)
End Sub

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla CajaComprobantesPago
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCajaComprobantesPago) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CajaComprobantesPagoSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, oTabla.IdComprobantePago): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       ActualizaTabla oTabla, oRecordset
       SeleccionarPorId = True
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

Function SeleccionarPorCuentaAtencion(ByVal oTabla As DOCajaComprobantesPago) As Boolean

On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CajaComprobantesPagoXidCuentaAtencion"
     Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oTabla.IdCuentaAtencion): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       ActualizaTabla oTabla, oRecordset
       SeleccionarPorCuentaAtencion = True
   Else
       SeleccionarPorCuentaAtencion = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function AnularComprobanteYOrdenServicio(lIdComprobantePago As Long, lIdUsuario As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   AnularComprobanteYOrdenServicio = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesAnularComprobanteYOrdenServicio"
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lIdComprobantePago): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, lIdUsuario)
       .Parameters.Append oParameter
       .Execute
   End With
 
   AnularComprobanteYOrdenServicio = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "AnularComprobanteYOrdenServicio"
Exit Function
End Function

Function AnularComprobanteYOrdenBienInsumo(lIdComprobantePago As Long, lIdUsuario As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   AnularComprobanteYOrdenBienInsumo = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesAnularComprobanteYOrdenBienInsumo"
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lIdComprobantePago): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, lIdUsuario)
       .Parameters.Append oParameter
       .Execute
   End With
 
   AnularComprobanteYOrdenBienInsumo = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "AnularComprobanteYOrdenBienInsumo"
Exit Function
End Function
Function DevolverComprobanteYOrdenServicio(lIdComprobantePago As Long, lIdComprobanteDeDevolucion As Long, lIdUsuario As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   DevolverComprobanteYOrdenServicio = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesDevolverComprobanteYOrdenServicio"
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lIdComprobantePago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdComprobantePagoDevolucion", adInteger, adParamInput, 0, lIdComprobanteDeDevolucion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, lIdUsuario): .Parameters.Append oParameter
       .Execute
   End With
 
   DevolverComprobanteYOrdenServicio = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "DevolverComprobanteYOrdenServicio"
Exit Function
End Function

Function DevolverComprobanteYOrdenBienInsumo(lIdComprobantePago As Long, lIdComprobanteDeDevolucion As Long, lIdUsuario As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   DevolverComprobanteYOrdenBienInsumo = False
   With oCommand
       .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaComprobantesDevolverComprobanteYOrdenBienInsumo"
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lIdComprobantePago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdComprobantePagoDevolucion", adInteger, adParamInput, 0, lIdComprobanteDeDevolucion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, lIdUsuario): .Parameters.Append oParameter
       .Execute
   End With
 
   DevolverComprobanteYOrdenBienInsumo = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "DevolverComprobanteYOrdenBienInsumo"
Exit Function
End Function


Sub Trace(sMensaje As String, NombreDeMetodo As String)
    MsgBox "Mensaje: " & sMensaje + Chr(13) + "Metodo:" & NombreDeMetodo + Chr(13) + "Clase de Datos: CajaComprobantesPago", vbInformation, "Atenciones"
End Sub

