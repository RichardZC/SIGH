VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "farmMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla farmMovimiento
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String

Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property




'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla farmMovimiento
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOfarmMovimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
If Format(oTabla.FechaCreacion, "hh:mm:ss") = "00:00:00" Then
   oTabla.FechaCreacion = CDate(Format(oTabla.FechaCreacion, "dd/mm/yyyy ") & "00:00:01")
End If
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmMovimientoAgregar"
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idAlmacenOrigen", adInteger, adParamInput, 0, oTabla.idAlmacenOrigen)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idAlmacenDestino", adInteger, adParamInput, 0, oTabla.idAlmacenDestino)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoConcepto", adInteger, adParamInput, 0, IIf(oTabla.idTipoConcepto = 0, Null, oTabla.idTipoConcepto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoIdtipo", adInteger, adParamInput, 0, IIf(oTabla.DocumentoIdtipo = 0, Null, oTabla.DocumentoIdtipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoNumero", adVarChar, adParamInput, 20, IIf(oTabla.DocumentoNumero = "", Null, oTabla.DocumentoNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Observaciones", adVarChar, adParamInput, 100, IIf(oTabla.Observaciones = "", Null, oTabla.Observaciones))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Total", adDouble, adParamInput, 0, oTabla.Total)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idMotivoAnulacion", adInteger, adParamInput, 0, IIf(oTabla.idMotivoAnulacion = 0, Null, oTabla.idMotivoAnulacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fechaAnulacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fechaAnulacion = 0, Null, oTabla.fechaAnulacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuarioAnulacion", adInteger, adParamInput, 0, IIf(oTabla.idUsuarioAnulacion = 0, Null, oTabla.idUsuarioAnulacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoMovimiento", adInteger, adParamInput, 0, oTabla.IdEstadoMovimiento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   
    'RHA 12/01/2021 CAMBIO 50 INCIO
    Set oParameter = .CreateParameter("@DocExterno", adVarChar, adParamInput, 20, IIf(oTabla.docExterno = "", Null, oTabla.docExterno))
   .Parameters.Append oParameter
   'RHA 12/01/2021 CAMBIO 50 FIN
   
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
'   Descripción:    modifica un registro a la tabla farmMovimiento
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOfarmMovimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmMovimientoModificar"
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idAlmacenOrigen", adInteger, adParamInput, 0, oTabla.idAlmacenOrigen)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idAlmacenDestino", adInteger, adParamInput, 0, oTabla.idAlmacenDestino)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoConcepto", adInteger, adParamInput, 0, IIf(oTabla.idTipoConcepto = 0, Null, oTabla.idTipoConcepto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoIdtipo", adInteger, adParamInput, 0, IIf(oTabla.DocumentoIdtipo = 0, Null, oTabla.DocumentoIdtipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoNumero", adVarChar, adParamInput, 20, IIf(oTabla.DocumentoNumero = "", Null, oTabla.DocumentoNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Observaciones", adVarChar, adParamInput, 100, IIf(oTabla.Observaciones = "", Null, oTabla.Observaciones))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Total", adDouble, adParamInput, 0, oTabla.Total)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idMotivoAnulacion", adInteger, adParamInput, 0, IIf(oTabla.idMotivoAnulacion = 0, Null, oTabla.idMotivoAnulacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fechaAnulacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fechaAnulacion = 0, Null, oTabla.fechaAnulacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuarioAnulacion", adInteger, adParamInput, 0, IIf(oTabla.idUsuarioAnulacion = 0, Null, oTabla.idUsuarioAnulacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoMovimiento", adInteger, adParamInput, 0, oTabla.IdEstadoMovimiento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   
   'RHA 12/01/2021 CAMBIO 50 INCIO
    Set oParameter = .CreateParameter("@DocExterno", adVarChar, adParamInput, 20, IIf(oTabla.docExterno = "", Null, oTabla.docExterno))
   .Parameters.Append oParameter
   'RHA 12/01/2021 CAMBIO 50 FIN
   
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
'   Descripción:    elimia o borra un registro a la tabla farmMovimiento
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOfarmMovimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmMovimientoEliminar"
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, oTabla.movNumero)
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
'   Descripción:    Selecciona un unico registro de la tabla farmMovimiento
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOfarmMovimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "farmMovimientoSeleccionarPorId"
     Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, oTabla.movNumero)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, oTabla.movTipo)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
'       oTabla.MovNumero = IIf(IsNull(oRecordset!MovNumero), "", oRecordset!MovNumero)
'       oTabla.MovTipo = IIf(IsNull(oRecordset!MovTipo), "", oRecordset!MovTipo)
       oTabla.idAlmacenOrigen = IIf(IsNull(oRecordset!idAlmacenOrigen), 0, oRecordset!idAlmacenOrigen)
       oTabla.idAlmacenDestino = IIf(IsNull(oRecordset!idAlmacenDestino), 0, oRecordset!idAlmacenDestino)
       oTabla.idTipoConcepto = IIf(IsNull(oRecordset!idTipoConcepto), 0, oRecordset!idTipoConcepto)
       oTabla.DocumentoIdtipo = IIf(IsNull(oRecordset!DocumentoIdtipo), 0, oRecordset!DocumentoIdtipo)
       oTabla.DocumentoNumero = IIf(IsNull(oRecordset!DocumentoNumero), "", oRecordset!DocumentoNumero)
       oTabla.Observaciones = IIf(IsNull(oRecordset!Observaciones), "", oRecordset!Observaciones)
       oTabla.Total = oRecordset!Total
       oTabla.idMotivoAnulacion = IIf(IsNull(oRecordset!idMotivoAnulacion), 0, oRecordset!idMotivoAnulacion)
       oTabla.fechaAnulacion = IIf(IsNull(oRecordset!fechaAnulacion), 0, oRecordset!fechaAnulacion)
       oTabla.idUsuarioAnulacion = IIf(IsNull(oRecordset!idUsuarioAnulacion), 0, oRecordset!idUsuarioAnulacion)
       oTabla.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
       oTabla.IdUsuario = IIf(IsNull(oRecordset!IdUsuario), 0, oRecordset!IdUsuario)
       oTabla.IdEstadoMovimiento = IIf(IsNull(oRecordset!IdEstadoMovimiento), 0, oRecordset!IdEstadoMovimiento)
       
       oTabla.docExterno = IIf(IsNull(oRecordset!docExterno), "", oRecordset!docExterno) 'RHA 12/01/2021 CAMBIO 50
       
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOfarmMovimiento)
   oTabla.movNumero = ""
   oTabla.movTipo = ""
   oTabla.idAlmacenOrigen = 0
   oTabla.idAlmacenDestino = 0
   oTabla.idTipoConcepto = 0
   oTabla.DocumentoIdtipo = 0
   oTabla.DocumentoNumero = ""
   oTabla.Observaciones = ""
   oTabla.Total = 0
   oTabla.idMotivoAnulacion = 0
   oTabla.fechaAnulacion = Now
   oTabla.idUsuarioAnulacion = 0
   oTabla.FechaCreacion = Now
   oTabla.IdUsuario = 0
   oTabla.IdEstadoMovimiento = 0
    oTabla.docExterno = "" 'RHA 12/01/2021 CAMBIO 50
End Sub

'********************************   Correlativo  ***************************************
Function DevuelveYactualizaCorrelativosDeDocumentosES(lnIdTipoDocumento As Long) As String
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lcNewCorrelativo As String
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    DevuelveYactualizaCorrelativosDeDocumentosES = ""
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FarmDevuelveYactualizaCorrelativosDeDocumentosES"
       Set oParameter = .CreateParameter("@IdTipoDocumento", adInteger, adParamInput, 0, lnIdTipoDocumento): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@NewCorrelativo", adInteger, adParamOutput, 0): .Parameters.Append oParameter
       .Execute
       lcNewCorrelativo = Trim(Str(.Parameters("@NewCorrelativo")))
    End With
    If lnIdTipoDocumento = 1 Or lnIdTipoDocumento = 2 Then
        DevuelveYactualizaCorrelativosDeDocumentosES = Right(lcBuscaParametro.RetornaFechaServidorSQL, 2) + Right("0000000" + lcNewCorrelativo, 7)
    Else
       DevuelveYactualizaCorrelativosDeDocumentosES = lcNewCorrelativo
    End If
    ms_MensajeError = ""
    Exit Function
ManejadorDeError:
       ms_MensajeError = Err.Number & " " + Err.Description

End Function
'SCCQ 14/10/2020 Cambio28 Inicio
Function DevuelveYactualizaCorrelativoDisponibleFarmRelMod(lnMovTipo As String, lnIdTipoLocales As String, lnIdTipoSuministro As String, lnDocumentoIdtipo As Long) As String
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lcNewCorrelativo As String
    DevuelveYactualizaCorrelativoDisponibleFarmRelMod = ""
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "usp_FarmDevuelveYactualizaCorrelativoDisponibleFarmRelMod"
       Set oParameter = .CreateParameter("@MovTipo", adChar, adParamInput, 1, lnMovTipo): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@idTipoLocales", adChar, adParamInput, 1, lnIdTipoLocales): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@idTipoSuministro", adChar, adParamInput, 2, lnIdTipoSuministro): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@DocumentoIdtipo", adInteger, adParamInput, 0, lnDocumentoIdtipo): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@NewCorrelativo", adInteger, adParamOutput, 0): .Parameters.Append oParameter
       .Execute
       lcNewCorrelativo = Trim(Str(.Parameters("@NewCorrelativo")))
    End With
       DevuelveYactualizaCorrelativoDisponibleFarmRelMod = lcNewCorrelativo
    ms_MensajeError = ""
    Exit Function
ManejadorDeError:
       ms_MensajeError = Err.Number & " " + Err.Description
       End Function
'SCCQ 14/10/2020 Cambio28 Fin

'SCCQ 22/10/2020 Cambio28 Inicio
Function DevuelveYactualizaCorrelativosDisponibleFarmTipoDocumentos(lnIdTipoDocumento As Long) As String
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lcNewCorrelativo As String
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    DevuelveYactualizaCorrelativosDisponibleFarmTipoDocumentos = ""
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "usp_FarmDevuelveYactualizaCorrelativosDisponibleFarmTipoDocumentos"
       Set oParameter = .CreateParameter("@IdTipoDocumento", adInteger, adParamInput, 0, lnIdTipoDocumento): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@NewCorrelativo", adInteger, adParamOutput, 0): .Parameters.Append oParameter
       .Execute
       lcNewCorrelativo = Trim(Str(.Parameters("@NewCorrelativo")))
    End With
    If lnIdTipoDocumento = 1 Or lnIdTipoDocumento = 2 Then
        DevuelveYactualizaCorrelativosDisponibleFarmTipoDocumentos = Right(lcBuscaParametro.RetornaFechaServidorSQL, 2) + Right("0000000" + lcNewCorrelativo, 7)
    Else
       DevuelveYactualizaCorrelativosDisponibleFarmTipoDocumentos = lcNewCorrelativo
    End If
    ms_MensajeError = ""
    Exit Function
ManejadorDeError:
       ms_MensajeError = Err.Number & " " + Err.Description

End Function
'SCCQ 22/10/2020 Cambio28 Fin
