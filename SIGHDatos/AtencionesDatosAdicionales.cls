VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionesDatosAdicionales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtencionesDatosAdicionales
'        Programado por: Barrantes D
'        Fecha: Agosto 2011
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
'   Descripción:    Inserta un registro a la tabla AtencionesDatosAdicionales
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtencionDatosAdicionales) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtencionesDatosAdicionalesAgregar"
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, oTabla.idAtencion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DireccionDomicilio", adVarChar, adParamInput, 100, IIf(oTabla.DireccionDomicilio = "", Null, oTabla.DireccionDomicilio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NombreAcompaniante", adVarChar, adParamInput, 100, IIf(oTabla.NombreAcompaniante = "", Null, oTabla.NombreAcompaniante))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 200, IIf(oTabla.Observacion = "", Null, oTabla.Observacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@proximaCita", adDBTimeStamp, adParamInput, 0, IIf(oTabla.ProximaCita = 0, Null, oTabla.ProximaCita))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NumeroDeHijos", adInteger, adParamInput, 0, oTabla.NumeroDeHijos)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdSiaSis", adInteger, adParamInput, 0, oTabla.idSiasis)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaCodigoPrestacion", adVarChar, adParamInput, 200, IIf(oTabla.FuaCodigoPrestacion = "", Null, oTabla.FuaCodigoPrestacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SisCodigo", adVarChar, adParamInput, 2, IIf(oTabla.SisCodigo = "", Null, oTabla.SisCodigo))
   .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdTipoReferenciaDestino", adInteger, adParamInput, 0, IIf(oTabla.IdTipoReferenciaDestino = 0, Null, oTabla.IdTipoReferenciaDestino)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdTipoReferenciaOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdTipoReferenciaOrigen = 0, Null, oTabla.IdTipoReferenciaOrigen)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdEstablecimientoDestino", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoDestino = 0, Null, oTabla.IdEstablecimientoDestino)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdEstablecimientoOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoOrigen = 0, Null, oTabla.IdEstablecimientoOrigen)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdEstablecimientoNoMinsaDestino", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoNoMinsaDestino = 0, Null, oTabla.IdEstablecimientoNoMinsaDestino)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdEstablecimientoNoMinsaOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoNoMinsaOrigen = 0, Null, oTabla.IdEstablecimientoNoMinsaOrigen)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@HuboInfeccionIntraHospitalaria", adBoolean, adParamInput, 0, IIf(oTabla.HuboInfeccionIntraHospitalaria = 0, Null, oTabla.HuboInfeccionIntraHospitalaria)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@TieneNecropsia", adBoolean, adParamInput, 0, IIf(oTabla.TieneNecropsia = 0, Null, oTabla.TieneNecropsia)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdMedicoRespNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoRespNacimiento = 0, Null, oTabla.IdMedicoRespNacimiento)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@RecienNacido", adBoolean, adParamInput, 0, IIf(oTabla.RecienNacido = 0, Null, oTabla.RecienNacido)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@NroReferenciaOrigen", adVarChar, adParamInput, 20, IIf(oTabla.NroReferenciaOrigen = "", Null, oTabla.NroReferenciaOrigen)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@NroReferenciaDestino", adVarChar, adParamInput, 20, IIf(oTabla.NroReferenciaDestino = "", Null, oTabla.NroReferenciaDestino)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@referenciaOservicio", adVarChar, adParamInput, 7, IIf(oTabla.referenciaOservicio = "", Null, oTabla.referenciaOservicio)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@referenciaOidDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.referenciaOidDiagnostico = 0, Null, oTabla.referenciaOidDiagnostico)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@referenciaDservicio", adVarChar, adParamInput, 7, IIf(oTabla.referenciaDservicio = "", Null, oTabla.referenciaDservicio)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@referenciaDfextension", adDBTimeStamp, adParamInput, 0, IIf(oTabla.referenciaDfextension = 0, Null, oTabla.referenciaDfextension)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@referenciaDftramite", adDBTimeStamp, adParamInput, 0, IIf(oTabla.referenciaDftramite = 0, Null, oTabla.referenciaDftramite)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@referenciaOdniMedico", adVarChar, adParamInput, 8, IIf(oTabla.referenciaOdniMedico = "", Null, oTabla.referenciaOdniMedico)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@acompanianteDNI", adVarChar, adParamInput, 8, IIf(oTabla.acompanianteDNI = "", Null, oTabla.acompanianteDNI)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@emergenciaCorrelativo", adVarChar, adParamInput, 10, IIf(oTabla.emergenciaCorrelativo = "", Null, oTabla.emergenciaCorrelativo)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@sisAfiliacion", adVarChar, adParamInput, 30, IIf(oTabla.sisAfiliacion = "", Null, oTabla.sisAfiliacion)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idAtencionEmeg_CE", adInteger, adParamInput, 0, IIf(oTabla.idAtencionEmeg_CE = 0, Null, oTabla.idAtencionEmeg_CE)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ReferenciaMedicoOColeg", adVarChar, adParamInput, 6, oTabla.ReferenciaMedicoOColeg): .Parameters.Append oParameter     'franklin 2017
    Set oParameter = .CreateParameter("@ReferenciaMedicoOIdcolegio", adVarChar, adParamInput, 2, oTabla.ReferenciaMedicoOIdcolegio): .Parameters.Append oParameter    'franklin 2017
    Set oParameter = .CreateParameter("@idServicioDestino", adInteger, adParamInput, 0, IIf(oTabla.IdServicioDestino = 0, Null, oTabla.IdServicioDestino)): .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Insertar = True
ms_MensajeError = ""
 GrabaHistorico oTabla
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
Resume
End Function


Sub GrabaHistorico(oTabla As DOAtencionDatosAdicionales)
    On Error GoTo ErrHistorico
    Dim lcSql As String
    Dim oRsParametros As New Recordset
    Dim oRsMDB As New Recordset
    Dim oConexionMDB As New Connection
    Dim lnErrorN As Long
    lnErrorN = 1
    lcSql = "select ValorTexto from Parametros where idparametro=581"
    oRsParametros.Open lcSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    lnErrorN = 2
    oConexionMDB.CommandTimeout = 900
    oConexionMDB.CursorLocation = adUseClient
    oConexionMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & Trim(oRsParametros!ValorTexto) & "\parametros.mdb"
    '
    lnErrorN = 21
    lcSql = "select fechaIngreso from atenciones where idatencion=" & oTabla.idAtencion
    If oRsParametros.State = 1 Then oRsParametros.Close
    oRsParametros.Open lcSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    '
    lnErrorN = 3
    lcSql = "INSERT INTO AtencionesDatosAdicionales" & _
"           (idAtencion,DireccionDomicilio,NombreAcompaniante" & _
"           ,Observacion,ProximaCita,NumeroDeHijos" & _
"           ,IdSiaSis,FuaCodigoPrestacion,SisCodigo" & _
"           ,IdTipoReferenciaOrigen,IdTipoReferenciaDestino,IdEstablecimientoOrigen" & _
"           ,IdEstablecimientoDestino,IdEstablecimientoNoMinsaOrigen,IdEstablecimientoNoMinsaDestino" & _
"           ,HuboInfeccionIntraHospitalaria,TieneNecropsia,IdMedicoRespNacimiento" & _
"           ,RecienNacido,NroReferenciaOrigen,NroReferenciaDestino,fecha0)" & _
"     Values(" & _
"           " & oTabla.idAtencion & ",'" & oTabla.DireccionDomicilio & "','" & oTabla.NombreAcompaniante & _
"           ','" & oTabla.Observacion & "','" & oTabla.ProximaCita & "'," & oTabla.NumeroDeHijos & _
"           ," & oTabla.idSiasis & ",'" & oTabla.FuaCodigoPrestacion & "','" & oTabla.SisCodigo & _
"           '," & oTabla.IdTipoReferenciaOrigen & "," & oTabla.IdTipoReferenciaDestino & "," & oTabla.IdEstablecimientoOrigen & _
"           ," & oTabla.IdEstablecimientoDestino & "," & oTabla.IdEstablecimientoNoMinsaOrigen & "," & oTabla.IdEstablecimientoNoMinsaDestino & _
"           ," & IIf(oTabla.HuboInfeccionIntraHospitalaria = True, "1", "0") & "," & IIf(oTabla.TieneNecropsia = True, "1", "0") & "," & oTabla.IdMedicoRespNacimiento & _
"           ," & IIf(oTabla.RecienNacido = True, "1", "0") & ",'" & oTabla.NroReferenciaOrigen & "','" & oTabla.NroReferenciaDestino & "','" & Format(oRsParametros!FechaIngreso, "dd/mm/yyyy") & "'" & _
"           )"
   oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
   lnErrorN = 4
    oConexionMDB.Close
    lnErrorN = 0
ErrHistorico:
    If lnErrorN > 0 Then
        lcSql = "update parametros set ValorInt=" & lnErrorN & " where idparametro=581"
        oRsParametros.Open lcSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    End If
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing
End Sub

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla AtencionesDatosAdicionales
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencionDatosAdicionales) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtencionesDatosAdicionalesModificar"
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, oTabla.idAtencion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DireccionDomicilio", adVarChar, adParamInput, 100, IIf(oTabla.DireccionDomicilio = "", Null, oTabla.DireccionDomicilio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NombreAcompaniante", adVarChar, adParamInput, 100, IIf(oTabla.NombreAcompaniante = "", Null, oTabla.NombreAcompaniante))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 200, IIf(oTabla.Observacion = "", Null, oTabla.Observacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@proximaCita", adDBTimeStamp, adParamInput, 0, IIf(oTabla.ProximaCita = 0, Null, oTabla.ProximaCita))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NumeroDeHijos", adInteger, adParamInput, 0, oTabla.NumeroDeHijos)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdSiaSis", adInteger, adParamInput, 0, oTabla.idSiasis)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaCodigoPrestacion", adVarChar, adParamInput, 200, IIf(oTabla.FuaCodigoPrestacion = "", Null, oTabla.FuaCodigoPrestacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SisCodigo", adVarChar, adParamInput, 2, IIf(oTabla.SisCodigo = "", Null, oTabla.SisCodigo))
   .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdTipoReferenciaDestino", adInteger, adParamInput, 0, IIf(oTabla.IdTipoReferenciaDestino = 0, Null, oTabla.IdTipoReferenciaDestino)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdTipoReferenciaOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdTipoReferenciaOrigen = 0, Null, oTabla.IdTipoReferenciaOrigen)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdEstablecimientoDestino", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoDestino = 0, Null, oTabla.IdEstablecimientoDestino)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdEstablecimientoOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoOrigen = 0, Null, oTabla.IdEstablecimientoOrigen)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdEstablecimientoNoMinsaDestino", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoNoMinsaDestino = 0, Null, oTabla.IdEstablecimientoNoMinsaDestino)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdEstablecimientoNoMinsaOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoNoMinsaOrigen = 0, Null, oTabla.IdEstablecimientoNoMinsaOrigen)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@HuboInfeccionIntraHospitalaria", adBoolean, adParamInput, 0, IIf(oTabla.HuboInfeccionIntraHospitalaria = 0, Null, oTabla.HuboInfeccionIntraHospitalaria)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@TieneNecropsia", adBoolean, adParamInput, 0, IIf(oTabla.TieneNecropsia = 0, Null, oTabla.TieneNecropsia)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdMedicoRespNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoRespNacimiento = 0, Null, oTabla.IdMedicoRespNacimiento)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@RecienNacido", adBoolean, adParamInput, 0, IIf(oTabla.RecienNacido = 0, Null, oTabla.RecienNacido)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@NroReferenciaOrigen", adVarChar, adParamInput, 20, IIf(oTabla.NroReferenciaOrigen = "", Null, oTabla.NroReferenciaOrigen)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@NroReferenciaDestino", adVarChar, adParamInput, 20, IIf(oTabla.NroReferenciaDestino = "", Null, oTabla.NroReferenciaDestino)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@referenciaOservicio", adVarChar, adParamInput, 7, IIf(oTabla.referenciaOservicio = "", Null, oTabla.referenciaOservicio)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@referenciaOidDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.referenciaOidDiagnostico = 0, Null, oTabla.referenciaOidDiagnostico)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@referenciaDservicio", adVarChar, adParamInput, 7, IIf(oTabla.referenciaDservicio = "", Null, oTabla.referenciaDservicio)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@referenciaDfextension", adDBTimeStamp, adParamInput, 0, IIf(oTabla.referenciaDfextension = 0, Null, oTabla.referenciaDfextension)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@referenciaDftramite", adDBTimeStamp, adParamInput, 0, IIf(oTabla.referenciaDftramite = 0, Null, oTabla.referenciaDftramite)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@referenciaOdniMedico", adVarChar, adParamInput, 8, IIf(oTabla.referenciaOdniMedico = "", Null, oTabla.referenciaOdniMedico)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@acompanianteDNI", adVarChar, adParamInput, 8, IIf(oTabla.acompanianteDNI = "", Null, oTabla.acompanianteDNI)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@emergenciaCorrelativo", adVarChar, adParamInput, 10, IIf(oTabla.emergenciaCorrelativo = "", Null, oTabla.emergenciaCorrelativo)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@sisAfiliacion", adVarChar, adParamInput, 30, IIf(oTabla.sisAfiliacion = "", Null, oTabla.sisAfiliacion)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idAtencionEmeg_CE", adInteger, adParamInput, 0, IIf(oTabla.idAtencionEmeg_CE = 0, Null, oTabla.idAtencionEmeg_CE)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ReferenciaMedicoOColeg", adVarChar, adParamInput, 6, oTabla.ReferenciaMedicoOColeg): .Parameters.Append oParameter     'franklin 2017
    Set oParameter = .CreateParameter("@ReferenciaMedicoOIdcolegio", adVarChar, adParamInput, 2, oTabla.ReferenciaMedicoOIdcolegio): .Parameters.Append oParameter    'franklin 2017
    Set oParameter = .CreateParameter("@idServicioDestino", adInteger, adParamInput, 0, IIf(oTabla.IdServicioDestino = 0, Null, oTabla.IdServicioDestino)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla AtencionesDatosAdicionales
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencionDatosAdicionales) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtencionesDatosAdicionalesEliminar"
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, oTabla.idAtencion)
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
'   Descripción:    Selecciona un unico registro de la tabla AtencionesDatosAdicionales
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtencionDatosAdicionales) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionesDatosAdicionalesSeleccionarPorId"
     Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, oTabla.idAtencion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       oTabla.DireccionDomicilio = IIf(IsNull(oRecordset!DireccionDomicilio), "", oRecordset!DireccionDomicilio)
       oTabla.NombreAcompaniante = IIf(IsNull(oRecordset!NombreAcompaniante), "", oRecordset!NombreAcompaniante)
       oTabla.Observacion = IIf(IsNull(oRecordset!Observacion), "", oRecordset!Observacion)
       oTabla.ProximaCita = IIf(IsNull(oRecordset!ProximaCita), 0, oRecordset!ProximaCita)
       oTabla.NumeroDeHijos = IIf(IsNull(oRecordset!NumeroDeHijos), 0, oRecordset!NumeroDeHijos)
       oTabla.idSiasis = IIf(IsNull(oRecordset!idSiasis), 0, oRecordset!idSiasis)
       oTabla.FuaCodigoPrestacion = IIf(IsNull(oRecordset!FuaCodigoPrestacion), "", oRecordset!FuaCodigoPrestacion)
       oTabla.SisCodigo = IIf(IsNull(oRecordset!SisCodigo), "", oRecordset!SisCodigo)
       oTabla.IdTipoReferenciaDestino = IIf(IsNull(oRecordset!IdTipoReferenciaDestino), 0, oRecordset!IdTipoReferenciaDestino)
       oTabla.IdTipoReferenciaOrigen = IIf(IsNull(oRecordset!IdTipoReferenciaOrigen), 0, oRecordset!IdTipoReferenciaOrigen)
       oTabla.IdEstablecimientoDestino = IIf(IsNull(oRecordset!IdEstablecimientoDestino), 0, oRecordset!IdEstablecimientoDestino)
       oTabla.IdEstablecimientoOrigen = IIf(IsNull(oRecordset!IdEstablecimientoOrigen), 0, oRecordset!IdEstablecimientoOrigen)
       oTabla.IdEstablecimientoNoMinsaDestino = IIf(IsNull(oRecordset!IdEstablecimientoNoMinsaDestino), 0, oRecordset!IdEstablecimientoNoMinsaDestino)
       oTabla.IdEstablecimientoNoMinsaOrigen = IIf(IsNull(oRecordset!IdEstablecimientoNoMinsaOrigen), 0, oRecordset!IdEstablecimientoNoMinsaOrigen)
       oTabla.HuboInfeccionIntraHospitalaria = IIf(IsNull(oRecordset!HuboInfeccionIntraHospitalaria), 0, oRecordset!HuboInfeccionIntraHospitalaria)
       oTabla.TieneNecropsia = IIf(IsNull(oRecordset!TieneNecropsia), 0, oRecordset!TieneNecropsia)
       oTabla.IdMedicoRespNacimiento = IIf(IsNull(oRecordset!IdMedicoRespNacimiento), 0, oRecordset!IdMedicoRespNacimiento)
       oTabla.RecienNacido = IIf(IsNull(oRecordset!RecienNacido), 0, oRecordset!RecienNacido)
       oTabla.NroReferenciaOrigen = IIf(IsNull(oRecordset!NroReferenciaOrigen), "", oRecordset!NroReferenciaOrigen)
       oTabla.NroReferenciaDestino = IIf(IsNull(oRecordset!NroReferenciaDestino), "", oRecordset!NroReferenciaDestino)
       oTabla.SeImprimioFicha = IIf(IsNull(oRecordset!SeImprimioFicha), 0, oRecordset!SeImprimioFicha)
       oTabla.referenciaOservicio = IIf(IsNull(oRecordset!referenciaOservicio), "", oRecordset!referenciaOservicio)
       oTabla.referenciaOidDiagnostico = IIf(IsNull(oRecordset!referenciaOidDiagnostico), 0, oRecordset!referenciaOidDiagnostico)
       oTabla.referenciaDservicio = IIf(IsNull(oRecordset!referenciaDservicio), "", oRecordset!referenciaDservicio)
       oTabla.referenciaDfextension = IIf(IsNull(oRecordset!referenciaDfextension), 0, oRecordset!referenciaDfextension)
       oTabla.referenciaDftramite = IIf(IsNull(oRecordset!referenciaDftramite), 0, oRecordset!referenciaDftramite)
       oTabla.referenciaOdniMedico = IIf(IsNull(oRecordset!referenciaOdniMedico), "", oRecordset!referenciaOdniMedico)
       oTabla.acompanianteDNI = IIf(IsNull(oRecordset!acompanianteDNI), "", oRecordset!acompanianteDNI)
       oTabla.emergenciaCorrelativo = IIf(IsNull(oRecordset!emergenciaCorrelativo), "", oRecordset!emergenciaCorrelativo)
       oTabla.sisAfiliacion = IIf(IsNull(oRecordset!sisAfiliacion), "", oRecordset!sisAfiliacion)
       oTabla.idAtencionEmeg_CE = IIf(IsNull(oRecordset!idAtencionEmeg_CE), 0, oRecordset!idAtencionEmeg_CE)
       oTabla.ReferenciaMedicoOColeg = IIf(IsNull(oRecordset!ReferenciaMedicoOColeg), "", oRecordset!ReferenciaMedicoOColeg)    'franklin 2017
       oTabla.ReferenciaMedicoOIdcolegio = IIf(IsNull(oRecordset!ReferenciaMedicoOIdcolegio), "", oRecordset!ReferenciaMedicoOIdcolegio)    'franklin 2017
       oTabla.IdServicioDestino = IIf(IsNull(oRecordset!IdServicioDestino), 0, oRecordset!IdServicioDestino)
       
       'En sighDatos.Atenciones, tambien se carga los mismos datos, si hay algun cambio aquí también se
       'hará en sighDatos.Atenciones
       
       SeleccionarPorId = True
       ms_MensajeError = ""
   Else
       oTabla.idAtencion = 0
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOAtencionDatosAdicionales)
   oTabla.idAtencion = 0
   oTabla.DireccionDomicilio = ""
   oTabla.NombreAcompaniante = ""
   oTabla.Observacion = ""
End Sub

Function ModificarImprimioFicha(ByVal oTabla As DOAtencionDatosAdicionales) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
ModificarImprimioFicha = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtencionesDatosAdicionalesModificarImprimioFicha"
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, oTabla.idAtencion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SeImprimioFicha", adBoolean, adParamInput, 0, oTabla.SeImprimioFicha)
   .Parameters.Append oParameter
   .Execute
End With
 
ModificarImprimioFicha = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
