VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotaCreditoDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: CajaNotaAfectaComprobante
'        Autor: Daniel Enrique Barrantes
'        Fecha: 31/05/2015 06:34:17
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

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
'   Descripción:    Inserta un registro a la tabla NotaCreditoDebito
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DONotaCreditoDebito) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "NotaCreditoDebitoAgregar"
   Set oParameter = .CreateParameter("@IdNota", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoNota", adInteger, adParamInput, 0, IIf(oTabla.IdTipoNota = 0, Null, oTabla.IdTipoNota))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroSerie", adChar, adParamInput, 4, IIf(oTabla.NroSerie = "", Null, oTabla.NroSerie))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 12, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RazonSocial", adVarChar, adParamInput, 50, IIf(oTabla.RazonSocial = "", Null, oTabla.RazonSocial))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RUC", adChar, adParamInput, 11, IIf(oTabla.RUC = "", Null, oTabla.RUC))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SubTotal", adCurrency, adParamInput, 0, oTabla.SubTotal)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IGV", adCurrency, adParamInput, 0, oTabla.IGV)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, oTabla.Total)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAutoriza", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutoriza = 0, Null, oTabla.IdUsuarioAutoriza))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaAprueba", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaAprueba = 0, Null, oTabla.FechaAprueba))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoCambio", adCurrency, adParamInput, 0, oTabla.TipoCambio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Observaciones", adVarChar, adParamInput, 500, IIf(oTabla.Observaciones = "", Null, oTabla.Observaciones))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoNota", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoNota = 0, Null, oTabla.IdEstadoNota))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaPagado", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaPagado = 0, Null, oTabla.FechaPagado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdGestionCaja", adInteger, adParamInput, 0, IIf(oTabla.IdGestionCaja = 0, Null, oTabla.IdGestionCaja))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCaja", adInteger, adParamInput, 0, IIf(oTabla.IdCaja = 0, Null, oTabla.IdCaja))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idFarmacia", adInteger, adParamInput, 0, IIf(oTabla.idFarmacia = 0, Null, oTabla.idFarmacia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idMotivo", adInteger, adParamInput, 0, IIf(oTabla.idMotivo = 0, Null, oTabla.idMotivo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Direccion", adVarChar, adParamInput, 50, IIf(oTabla.Direccion = "", Null, oTabla.Direccion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoAnulacion", adBoolean, adParamInput, 0, IIf(oTabla.TipoAnulacion = 0, Null, oTabla.TipoAnulacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdNota = .Parameters("@IdNota")
End With
 
Insertar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub GrabaHistorico(oTabla As DONotaCreditoDebito)
    If oTabla.FechaPagado = 0 Then
       Exit Sub
    End If

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
    lcSql = "INSERT INTO NotaCreditoDebito" & _
"           (idNota,IdComprobantePago,IdTipoNota,NroSerie" & _
"           ,NroDocumento,RazonSocial,RUC" & _
"           ,SubTotal,IGV,Total" & _
"           ,IdUsuarioAutoriza,FechaAprueba" & _
"           ,Observaciones,IdEstadoNota,FechaPagado" & _
"           ,IdGestionCaja,IdPaciente,IdCajero" & _
"           ,idTurno,idCaja,idFarmacia" & _
"           ,idMotivo,Direccion,TipoAnulacion)" & _
"     Values" & _
"           (" & _
"           " & oTabla.IdNota & "," & oTabla.IdComprobantePago & "," & oTabla.IdTipoNota & ",'" & oTabla.NroSerie & _
"           ','" & oTabla.NroDocumento & "','" & oTabla.RazonSocial & "','" & oTabla.RUC & _
"           '," & oTabla.SubTotal & "," & oTabla.IGV & "," & oTabla.Total & _
"           ," & oTabla.IdUsuarioAutoriza & ",'" & oTabla.FechaAprueba & _
"           ','" & Left(oTabla.Observaciones, 100) & "'," & oTabla.IdEstadoNota & ",'" & oTabla.FechaPagado & _
"           '," & oTabla.IdGestionCaja & "," & oTabla.IdPaciente & "," & oTabla.IdCajero & _
"           ," & oTabla.IdTurno & "," & oTabla.IdCaja & "," & oTabla.idFarmacia & _
"           ," & oTabla.idMotivo & ",'" & oTabla.Direccion & "'," & IIf(oTabla.TipoAnulacion = True, "1", "0") & _
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
'   Descripción:    modifica un registro a la tabla NotaCreditoDebito
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DONotaCreditoDebito) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "NotaCreditoDebitoModificar"
   Set oParameter = .CreateParameter("@IdNota", adInteger, adParamInput, 0, IIf(oTabla.IdNota = 0, Null, oTabla.IdNota))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoNota", adInteger, adParamInput, 0, IIf(oTabla.IdTipoNota = 0, Null, oTabla.IdTipoNota))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, IIf(oTabla.NroSerie = "", Null, oTabla.NroSerie))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 12, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RazonSocial", adVarChar, adParamInput, 50, IIf(oTabla.RazonSocial = "", Null, oTabla.RazonSocial))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RUC", adChar, adParamInput, 11, IIf(oTabla.RUC = "", Null, oTabla.RUC))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SubTotal", adCurrency, adParamInput, 0, oTabla.SubTotal)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IGV", adCurrency, adParamInput, 0, oTabla.IGV)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, oTabla.Total)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAutoriza", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutoriza = 0, Null, oTabla.IdUsuarioAutoriza))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaAprueba", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaAprueba = 0, Null, oTabla.FechaAprueba))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoCambio", adCurrency, adParamInput, 0, oTabla.TipoCambio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Observaciones", adVarChar, adParamInput, 500, IIf(oTabla.Observaciones = "", Null, oTabla.Observaciones))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoNota", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoNota = 0, Null, oTabla.IdEstadoNota))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaPagado", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaPagado = 0, Null, oTabla.FechaPagado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdGestionCaja", adInteger, adParamInput, 0, IIf(oTabla.IdGestionCaja = 0, Null, oTabla.IdGestionCaja))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCaja", adInteger, adParamInput, 0, IIf(oTabla.IdCaja = 0, Null, oTabla.IdCaja))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idFarmacia", adInteger, adParamInput, 0, IIf(oTabla.idFarmacia = 0, Null, oTabla.idFarmacia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idMotivo", adInteger, adParamInput, 0, IIf(oTabla.idMotivo = 0, Null, oTabla.idMotivo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Direccion", adVarChar, adParamInput, 50, IIf(oTabla.Direccion = "", Null, oTabla.Direccion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoAnulacion", adBoolean, adParamInput, 0, IIf(oTabla.TipoAnulacion = 0, Null, oTabla.TipoAnulacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Modificar = True
ms_MensajeError = ""
 GrabaHistorico oTabla
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla NotaCreditoDebito
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DONotaCreditoDebito) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "NotaCreditoDebitoEliminar"
   Set oParameter = .CreateParameter("@IdNota", adInteger, adParamInput, 0, oTabla.IdNota)
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
'   Descripción:    Selecciona un unico registro de la tabla NotaCreditoDebito
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DONotaCreditoDebito) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "NotaCreditoDebitoSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdNota", adInteger, adParamInput, 0, oTabla.IdNota)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdNota = IIf(IsNull(oRecordset!IdNota), 0, oRecordset!IdNota)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.IdTipoNota = IIf(IsNull(oRecordset!IdTipoNota), 0, oRecordset!IdTipoNota)
       oTabla.NroSerie = IIf(IsNull(oRecordset!NroSerie), "", oRecordset!NroSerie)
       oTabla.NroDocumento = IIf(IsNull(oRecordset!NroDocumento), "", oRecordset!NroDocumento)
       oTabla.RazonSocial = IIf(IsNull(oRecordset!RazonSocial), "", oRecordset!RazonSocial)
       oTabla.RUC = IIf(IsNull(oRecordset!RUC), "", oRecordset!RUC)
       oTabla.SubTotal = oRecordset!SubTotal
       oTabla.IGV = oRecordset!IGV
       oTabla.Total = oRecordset!Total
       oTabla.IdUsuarioAutoriza = IIf(IsNull(oRecordset!IdUsuarioAutoriza), 0, oRecordset!IdUsuarioAutoriza)
       oTabla.FechaAprueba = IIf(IsNull(oRecordset!FechaAprueba), 0, oRecordset!FechaAprueba)
       oTabla.TipoCambio = oRecordset!TipoCambio
       oTabla.Observaciones = IIf(IsNull(oRecordset!Observaciones), "", oRecordset!Observaciones)
       oTabla.IdEstadoNota = IIf(IsNull(oRecordset!IdEstadoNota), 0, oRecordset!IdEstadoNota)
       oTabla.FechaPagado = IIf(IsNull(oRecordset!FechaPagado), 0, oRecordset!FechaPagado)
       oTabla.IdGestionCaja = IIf(IsNull(oRecordset!IdGestionCaja), 0, oRecordset!IdGestionCaja)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.IdCajero = IIf(IsNull(oRecordset!IdCajero), 0, oRecordset!IdCajero)
       oTabla.IdTurno = IIf(IsNull(oRecordset!IdTurno), 0, oRecordset!IdTurno)
       oTabla.IdCaja = IIf(IsNull(oRecordset!IdCaja), 0, oRecordset!IdCaja)
       oTabla.idFarmacia = IIf(IsNull(oRecordset!idFarmacia), 0, oRecordset!idFarmacia)
       oTabla.idMotivo = IIf(IsNull(oRecordset!idMotivo), 0, oRecordset!idMotivo)
       oTabla.Direccion = IIf(IsNull(oRecordset!Direccion), "", oRecordset!Direccion)
       oTabla.TipoAnulacion = IIf(IsNull(oRecordset!TipoAnulacion), 0, oRecordset!TipoAnulacion)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DONotaCreditoDebito)
   oTabla.IdNota = 0
   oTabla.IdComprobantePago = 0
   oTabla.IdTipoNota = 0
   oTabla.NroSerie = ""
   oTabla.NroDocumento = ""
   oTabla.RazonSocial = ""
   oTabla.RUC = ""
   oTabla.SubTotal = 0
   oTabla.IGV = 0
   oTabla.Total = 0
   oTabla.IdUsuarioAutoriza = 0
   oTabla.FechaAprueba = Now
   oTabla.TipoCambio = 0
   oTabla.Observaciones = ""
   oTabla.IdEstadoNota = 0
   oTabla.FechaPagado = Now
   oTabla.IdGestionCaja = 0
   oTabla.IdPaciente = 0
   oTabla.IdCajero = 0
   oTabla.IdTurno = 0
   oTabla.IdCaja = 0
   oTabla.idFarmacia = 0
   oTabla.idMotivo = 0
   oTabla.Direccion = ""
   oTabla.TipoAnulacion = False
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: NotaCreditoDebito
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: NotaCreditoDebito
'        Autor:   DAniel Barrantes B.
'        Fecha: 13/08/2015 10:54:51
'------------------------------------------------------------------------------------
 

