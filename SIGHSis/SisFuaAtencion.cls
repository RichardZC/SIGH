VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SisFuaAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para el Mantenimiento de la tabla SisAtencion
'        Programado por: Barrantes D
'        Fecha: Enero 2013
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
'   Descripción:    Inserta un registro a la tabla SisFuaAtencion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoSisFuaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaAtencionAgregar"
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaDisa", adVarChar, adParamInput, 3, IIf(oTabla.FuaDisa = "", Null, oTabla.FuaDisa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaLote", adVarChar, adParamInput, 2, IIf(oTabla.FuaLote = "", Null, oTabla.FuaLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaNumero", adVarChar, adParamInput, 16, IIf(oTabla.FuaNumero = "", Null, oTabla.FuaNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EstablecimientoCodigoRenaes", adVarChar, adParamInput, 10, IIf(oTabla.EstablecimientoCodigoRenaes = "", Null, oTabla.EstablecimientoCodigoRenaes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Reconsideracion", adVarChar, adParamInput, 1, IIf(oTabla.Reconsideracion = "", Null, oTabla.Reconsideracion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ReconsideracionCodigoDisa", adVarChar, adParamInput, 3, IIf(oTabla.ReconsideracionCodigoDisa = "", Null, oTabla.ReconsideracionCodigoDisa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ReconsideracionLote", adVarChar, adParamInput, 2, IIf(oTabla.ReconsideracionLote = "", Null, oTabla.ReconsideracionLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ReconsideracionNroFormato", adVarChar, adParamInput, 8, IIf(oTabla.ReconsideracionNroFormato = "", Null, oTabla.ReconsideracionNroFormato))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaComponente", adVarChar, adParamInput, 1, IIf(oTabla.FuaComponente = "", Null, oTabla.FuaComponente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Situacion", adVarChar, adParamInput, 1, IIf(oTabla.Situacion = "", Null, oTabla.Situacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionDisa", adVarChar, adParamInput, 3, IIf(oTabla.AfiliacionDisa = "", Null, oTabla.AfiliacionDisa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionTipoFormato", adVarChar, adParamInput, 2, IIf(oTabla.AfiliacionTipoFormato = "", Null, oTabla.AfiliacionTipoFormato))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionNroFormato", adVarChar, adParamInput, 10, IIf(oTabla.AfiliacionNroFormato = "", Null, oTabla.AfiliacionNroFormato))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoTipoFormato", adVarChar, adParamInput, 2, IIf(oTabla.CodigoTipoFormato = "", Null, oTabla.CodigoTipoFormato))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OrigenAseguradoInstitucion", adVarChar, adParamInput, 1, IIf(oTabla.OrigenAseguradoInstitucion = "", "0", oTabla.OrigenAseguradoInstitucion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OrigenAseguradoCodigo", adVarChar, adParamInput, 16, IIf(oTabla.OrigenAseguradoCodigo = "", Null, oTabla.OrigenAseguradoCodigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Edad", adInteger, adParamInput, 0, IIf(oTabla.Edad = 0, Null, oTabla.Edad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrupoEtareo", adVarChar, adParamInput, 2, IIf(oTabla.GrupoEtareo = "", "0", oTabla.GrupoEtareo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Genero", adVarChar, adParamInput, 1, IIf(oTabla.Genero = "", Null, oTabla.Genero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaAtencion", adInteger, adParamInput, 0, IIf(oTabla.FuaAtencion = 0, Null, oTabla.FuaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaCondicionMaterna", adVarChar, adParamInput, 1, IIf(oTabla.FuaCondicionMaterna = "", Null, oTabla.FuaCondicionMaterna))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaNrohistoria", adVarChar, adParamInput, 20, IIf(oTabla.FuaNrohistoria = "", Null, oTabla.FuaNrohistoria))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaConceptoPr", adInteger, adParamInput, 0, IIf(oTabla.FuaConceptoPr = 0, 1, oTabla.FuaConceptoPr))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaConceptoPrAutoriz", adVarChar, adParamInput, 15, IIf(oTabla.FuaConceptoPrAutoriz = "", Null, oTabla.FuaConceptoPrAutoriz))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaConceptoPrMonto", adCurrency, adParamInput, 0, oTabla.FuaConceptoPrMonto)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaAtencionFecha", adVarChar, adParamInput, 10, IIf(oTabla.FuaAtencionFecha = "", Null, oTabla.FuaAtencionFecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaAtencionHora", adVarChar, adParamInput, 5, IIf(oTabla.FuaAtencionHora = "", Null, oTabla.FuaAtencionHora))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaReferidoOrigenCodigoRENAES", adVarChar, adParamInput, 10, IIf(oTabla.FuaReferidoOrigenCodigoRENAES = "", Null, oTabla.FuaReferidoOrigenCodigoRENAES))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaReferidoOrigenNreferencia", adVarChar, adParamInput, 20, IIf(oTabla.FuaReferidoOrigenNreferencia = "", Null, oTabla.FuaReferidoOrigenNreferencia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaCodigoPrestacion", adVarChar, adParamInput, 3, IIf(oTabla.FuaCodigoPrestacion = "", Null, oTabla.FuaCodigoPrestacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaPersonalQatiende", adInteger, adParamInput, 0, IIf(oTabla.FuaPersonalQatiende = 0, Null, oTabla.FuaPersonalQatiende))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaAtencionLugar", adVarChar, adParamInput, 1, IIf(oTabla.FuaAtencionLugar = "", Null, oTabla.FuaAtencionLugar))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaDestino", adVarChar, adParamInput, 1, IIf(oTabla.FuaDestino = "", Null, oTabla.FuaDestino))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaHospitalizadoFingreso", adVarChar, adParamInput, 10, IIf(oTabla.FuaHospitalizadoFingreso = "", Null, oTabla.FuaHospitalizadoFingreso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaHospitalizadoFalta", adVarChar, adParamInput, 10, IIf(oTabla.FuaHospitalizadoFalta = "", Null, oTabla.FuaHospitalizadoFalta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaReferidoDestinoCodigoRENAES", adVarChar, adParamInput, 10, IIf(oTabla.FuaReferidoDestinoCodigoRENAES = "", Null, oTabla.FuaReferidoDestinoCodigoRENAES))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaReferidoDestinoNreferencia", adVarChar, adParamInput, 20, IIf(oTabla.FuaReferidoDestinoNreferencia = "", Null, oTabla.FuaReferidoDestinoNreferencia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaMedicoDNI", adVarChar, adParamInput, 9, IIf(oTabla.FuaMedicoDNI = "", Null, Left(oTabla.FuaMedicoDNI, 9)))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaMedico", adVarChar, adParamInput, 120, IIf(oTabla.FuaMedico = "", Null, oTabla.FuaMedico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaMedicoTipo", adVarChar, adParamInput, 2, IIf(oTabla.FuaMedicoTipo = "", Null, oTabla.FuaMedicoTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionNroIntegrante", adVarChar, adParamInput, 2, IIf(oTabla.AfiliacionNroIntegrante = "", Null, oTabla.AfiliacionNroIntegrante))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 2, IIf(oTabla.codigo = "", Null, oTabla.codigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idSiasis", adVarChar, adParamInput, 10, IIf(oTabla.idSiasis = "", Null, oTabla.idSiasis))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaObservaciones", adVarChar, adParamInput, 200, IIf(oTabla.FuaObservaciones = "", Null, oTabla.FuaObservaciones))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabDniUsuarioRegistra", adVarChar, adParamInput, 10, IIf(oTabla.CabDniUsuarioRegistra = "", Null, Left(oTabla.CabDniUsuarioRegistra, 10)))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@UltimaFechaAddMod", adVarChar, adParamInput, 10, IIf(oTabla.UltimaFechaAddMod = "", Null, oTabla.UltimaFechaAddMod))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabEstado", adVarChar, adParamInput, 1, IIf(oTabla.CabEstado = "", Null, oTabla.CabEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaFechaParto", adVarChar, adParamInput, 10, IIf(oTabla.FuaFechaParto = "", Null, oTabla.FuaFechaParto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EstablecimientoDistrito", adVarChar, adParamInput, 6, IIf(oTabla.EstablecimientoDistrito = "", Null, oTabla.EstablecimientoDistrito))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Anio", adVarChar, adParamInput, 10, IIf(oTabla.Anio = "", Null, oTabla.Anio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Mes", adVarChar, adParamInput, 10, IIf(oTabla.Mes = "", Null, oTabla.Mes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CostoTotal", adCurrency, adParamInput, 0, oTabla.CostoTotal)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Apaterno", adVarChar, adParamInput, 40, IIf(oTabla.Apaterno = "", Null, oTabla.Apaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Amaterno", adVarChar, adParamInput, 40, IIf(oTabla.Amaterno = "", Null, oTabla.Amaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Pnombre", adVarChar, adParamInput, 35, IIf(oTabla.Pnombre = "", Null, oTabla.Pnombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Onombre", adVarChar, adParamInput, 35, IIf(oTabla.Onombre = "", Null, oTabla.Onombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fnacimiento", adVarChar, adParamInput, 10, IIf(oTabla.Fnacimiento = "", Null, oTabla.Fnacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Autogenerado", adVarChar, adParamInput, 20, IIf(oTabla.Autogenerado = "", Null, oTabla.Autogenerado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoTipo", adVarChar, adParamInput, 1, IIf(oTabla.DocumentoTipo = "", Null, oTabla.DocumentoTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoNumero", adVarChar, adParamInput, 10, IIf(oTabla.DocumentoNumero = "", Null, Left(oTabla.DocumentoNumero, 10)))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EstablecimientoCategoria", adVarChar, adParamInput, 2, IIf(oTabla.EstablecimientoCategoria = "", Null, oTabla.EstablecimientoCategoria))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CostoServicio", adCurrency, adParamInput, 0, oTabla.CostoServicio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CostoMedicamento", adCurrency, adParamInput, 0, oTabla.CostoMedicamento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CostoProcedimiento", adCurrency, adParamInput, 0, oTabla.CostoProcedimiento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CostoInsumo", adCurrency, adParamInput, 0, oTabla.CostoInsumo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MedicoDocumentoTipo", adVarChar, adParamInput, 2, IIf(oTabla.MedicoDocumentoTipo = "", Null, oTabla.MedicoDocumentoTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ate_grupoRiesgo", adVarChar, adParamInput, 2, IIf(oTabla.ate_grupoRiesgo = "", Null, oTabla.ate_grupoRiesgo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabCodigoPuntoDigitacion", adInteger, adParamInput, 0, IIf(oTabla.CabCodigoPuntoDigitacion = 0, Null, oTabla.CabCodigoPuntoDigitacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabCodigoUDR", adVarChar, adParamInput, 3, IIf(oTabla.CabCodigoUDR = "", Null, oTabla.CabCodigoUDR))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabNroEnvioAlSIS", adVarChar, adParamInput, 2, IIf(oTabla.CabNroEnvioAlSIS = "", Null, oTabla.CabNroEnvioAlSIS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabOrigenDelRegistro", adVarChar, adParamInput, 9, IIf(oTabla.CabOrigenDelRegistro = "", Null, oTabla.CabOrigenDelRegistro))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabVersionAplicativo", adVarChar, adParamInput, 9, IIf(oTabla.CabVersionAplicativo = "", Null, oTabla.CabVersionAplicativo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabIdentificacionPaquete", adInteger, adParamInput, 0, oTabla.CabIdentificacionPaquete)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdentificacionArfsis", adInteger, adParamInput, 0, IIf(oTabla.IdentificacionArfsis = 0, Null, oTabla.IdentificacionArfsis))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabFechaFuaPrimeraVez", adVarChar, adParamInput, 10, IIf(oTabla.CabFechaFuaPrimeraVez = "", Null, oTabla.CabFechaFuaPrimeraVez))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PeriodoOrigen", adVarChar, adParamInput, 6, IIf(oTabla.PeriodoOrigen = "", Null, oTabla.PeriodoOrigen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuacolegioCodigo", adVarChar, adParamInput, 20, IIf(oTabla.FuacolegioCodigo = "", Null, oTabla.FuacolegioCodigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuacolegioNivel", adVarChar, adParamInput, 1, IIf(oTabla.FuacolegioNivel = "", Null, oTabla.FuacolegioNivel))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuacolegioGrado", adVarChar, adParamInput, 1, IIf(oTabla.FuacolegioGrado = "", Null, oTabla.FuacolegioGrado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuacolegioSeccion", adVarChar, adParamInput, 5, IIf(oTabla.FuacolegioSeccion = "", Null, oTabla.FuacolegioSeccion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuacolegioTurno", adVarChar, adParamInput, 1, IIf(oTabla.FuacolegioTurno = "", Null, oTabla.FuacolegioTurno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Fuaetnia", adVarChar, adParamInput, 2, IIf(oTabla.Fuaetnia = "", Null, oTabla.Fuaetnia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuafechaFallecimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FuafechaFallecimiento = 0, Null, oTabla.FuafechaFallecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaUPS", adVarChar, adParamInput, 6, IIf(oTabla.FuaUPS = "", Null, oTabla.FuaUPS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaCodAutorizacion", adVarChar, adParamInput, 50, IIf(oTabla.FuaCodAutorizacion = "", Null, oTabla.FuaCodAutorizacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaFechaCorteAdm", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FuaFechaCorteAdm = 0, Null, oTabla.FuaFechaCorteAdm))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaVersionFormato", adVarChar, adParamInput, 2, IIf(oTabla.FuaVersionFormato = "", Null, oTabla.FuaVersionFormato))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaTipoAnexo2015", adInteger, adParamInput, 0, IIf(oTabla.FuaTipoAnexo2015 = 0, Null, oTabla.FuaTipoAnexo2015))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaCodOferFlexible", adVarChar, adParamInput, 20, IIf(oTabla.FuaCodOferFlexible = "", Null, oTabla.FuaCodOferFlexible))
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
'   Descripción:    modifica un registro a la tabla SisFuaAtencion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoSisFuaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaAtencionModificar"
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaDisa", adVarChar, adParamInput, 3, IIf(oTabla.FuaDisa = "", Null, oTabla.FuaDisa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaLote", adVarChar, adParamInput, 2, IIf(oTabla.FuaLote = "", Null, oTabla.FuaLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaNumero", adVarChar, adParamInput, 16, IIf(oTabla.FuaNumero = "", Null, oTabla.FuaNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EstablecimientoCodigoRenaes", adVarChar, adParamInput, 10, IIf(oTabla.EstablecimientoCodigoRenaes = "", Null, oTabla.EstablecimientoCodigoRenaes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Reconsideracion", adVarChar, adParamInput, 1, IIf(oTabla.Reconsideracion = "", Null, oTabla.Reconsideracion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ReconsideracionCodigoDisa", adVarChar, adParamInput, 3, IIf(oTabla.ReconsideracionCodigoDisa = "", Null, oTabla.ReconsideracionCodigoDisa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ReconsideracionLote", adVarChar, adParamInput, 2, IIf(oTabla.ReconsideracionLote = "", Null, oTabla.ReconsideracionLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ReconsideracionNroFormato", adVarChar, adParamInput, 8, IIf(oTabla.ReconsideracionNroFormato = "", Null, oTabla.ReconsideracionNroFormato))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaComponente", adVarChar, adParamInput, 1, IIf(oTabla.FuaComponente = "", Null, oTabla.FuaComponente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Situacion", adVarChar, adParamInput, 1, IIf(oTabla.Situacion = "", Null, oTabla.Situacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionDisa", adVarChar, adParamInput, 3, IIf(oTabla.AfiliacionDisa = "", Null, oTabla.AfiliacionDisa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionTipoFormato", adVarChar, adParamInput, 2, IIf(oTabla.AfiliacionTipoFormato = "", Null, oTabla.AfiliacionTipoFormato))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionNroFormato", adVarChar, adParamInput, 10, IIf(oTabla.AfiliacionNroFormato = "", Null, oTabla.AfiliacionNroFormato))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoTipoFormato", adVarChar, adParamInput, 2, IIf(oTabla.CodigoTipoFormato = "", Null, oTabla.CodigoTipoFormato))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OrigenAseguradoInstitucion", adVarChar, adParamInput, 1, IIf(oTabla.OrigenAseguradoInstitucion = "", "0", oTabla.OrigenAseguradoInstitucion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OrigenAseguradoCodigo", adVarChar, adParamInput, 16, IIf(oTabla.OrigenAseguradoCodigo = "", Null, oTabla.OrigenAseguradoCodigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Edad", adInteger, adParamInput, 0, IIf(oTabla.Edad = 0, Null, oTabla.Edad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrupoEtareo", adVarChar, adParamInput, 2, IIf(oTabla.GrupoEtareo = "", "0", oTabla.GrupoEtareo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Genero", adVarChar, adParamInput, 1, IIf(oTabla.Genero = "", Null, oTabla.Genero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaAtencion", adInteger, adParamInput, 0, IIf(oTabla.FuaAtencion = 0, Null, oTabla.FuaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaCondicionMaterna", adVarChar, adParamInput, 1, IIf(oTabla.FuaCondicionMaterna = "", Null, oTabla.FuaCondicionMaterna))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaNrohistoria", adVarChar, adParamInput, 20, IIf(oTabla.FuaNrohistoria = "", Null, oTabla.FuaNrohistoria))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaConceptoPr", adInteger, adParamInput, 0, IIf(oTabla.FuaConceptoPr = 0, 1, oTabla.FuaConceptoPr))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaConceptoPrAutoriz", adVarChar, adParamInput, 15, IIf(oTabla.FuaConceptoPrAutoriz = "", Null, oTabla.FuaConceptoPrAutoriz))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaConceptoPrMonto", adCurrency, adParamInput, 0, oTabla.FuaConceptoPrMonto)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaAtencionFecha", adVarChar, adParamInput, 10, IIf(oTabla.FuaAtencionFecha = "", Null, oTabla.FuaAtencionFecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaAtencionHora", adVarChar, adParamInput, 5, IIf(oTabla.FuaAtencionHora = "", Null, oTabla.FuaAtencionHora))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaReferidoOrigenCodigoRENAES", adVarChar, adParamInput, 10, IIf(oTabla.FuaReferidoOrigenCodigoRENAES = "", Null, oTabla.FuaReferidoOrigenCodigoRENAES))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaReferidoOrigenNreferencia", adVarChar, adParamInput, 20, IIf(oTabla.FuaReferidoOrigenNreferencia = "", Null, oTabla.FuaReferidoOrigenNreferencia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaCodigoPrestacion", adVarChar, adParamInput, 3, IIf(oTabla.FuaCodigoPrestacion = "", Null, oTabla.FuaCodigoPrestacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaPersonalQatiende", adInteger, adParamInput, 0, IIf(oTabla.FuaPersonalQatiende = 0, Null, oTabla.FuaPersonalQatiende))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaAtencionLugar", adVarChar, adParamInput, 1, IIf(oTabla.FuaAtencionLugar = "", Null, oTabla.FuaAtencionLugar))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaDestino", adVarChar, adParamInput, 1, IIf(oTabla.FuaDestino = "", Null, oTabla.FuaDestino))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaHospitalizadoFingreso", adVarChar, adParamInput, 10, IIf(oTabla.FuaHospitalizadoFingreso = "", Null, oTabla.FuaHospitalizadoFingreso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaHospitalizadoFalta", adVarChar, adParamInput, 10, IIf(oTabla.FuaHospitalizadoFalta = "", Null, oTabla.FuaHospitalizadoFalta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaReferidoDestinoCodigoRENAES", adVarChar, adParamInput, 10, IIf(oTabla.FuaReferidoDestinoCodigoRENAES = "", Null, oTabla.FuaReferidoDestinoCodigoRENAES))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaReferidoDestinoNreferencia", adVarChar, adParamInput, 20, IIf(oTabla.FuaReferidoDestinoNreferencia = "", Null, oTabla.FuaReferidoDestinoNreferencia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaMedicoDNI", adVarChar, adParamInput, 9, IIf(oTabla.FuaMedicoDNI = "", Null, Left(oTabla.FuaMedicoDNI, 9)))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaMedico", adVarChar, adParamInput, 120, IIf(oTabla.FuaMedico = "", Null, oTabla.FuaMedico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaMedicoTipo", adVarChar, adParamInput, 2, IIf(oTabla.FuaMedicoTipo = "", Null, oTabla.FuaMedicoTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionNroIntegrante", adVarChar, adParamInput, 2, IIf(oTabla.AfiliacionNroIntegrante = "", Null, oTabla.AfiliacionNroIntegrante))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 2, IIf(oTabla.codigo = "", Null, oTabla.codigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idSiasis", adVarChar, adParamInput, 10, IIf(oTabla.idSiasis = "", Null, oTabla.idSiasis))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaObservaciones", adVarChar, adParamInput, 200, IIf(oTabla.FuaObservaciones = "", Null, oTabla.FuaObservaciones))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabDniUsuarioRegistra", adVarChar, adParamInput, 10, IIf(oTabla.CabDniUsuarioRegistra = "", Null, Left(oTabla.CabDniUsuarioRegistra, 10)))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@UltimaFechaAddMod", adVarChar, adParamInput, 10, IIf(oTabla.UltimaFechaAddMod = "", Null, oTabla.UltimaFechaAddMod))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabEstado", adVarChar, adParamInput, 1, IIf(oTabla.CabEstado = "", Null, oTabla.CabEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaFechaParto", adVarChar, adParamInput, 10, IIf(oTabla.FuaFechaParto = "", Null, oTabla.FuaFechaParto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EstablecimientoDistrito", adVarChar, adParamInput, 6, IIf(oTabla.EstablecimientoDistrito = "", Null, oTabla.EstablecimientoDistrito))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Anio", adVarChar, adParamInput, 10, IIf(oTabla.Anio = "", Null, oTabla.Anio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Mes", adVarChar, adParamInput, 10, IIf(oTabla.Mes = "", Null, oTabla.Mes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CostoTotal", adCurrency, adParamInput, 0, oTabla.CostoTotal)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Apaterno", adVarChar, adParamInput, 40, IIf(oTabla.Apaterno = "", Null, oTabla.Apaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Amaterno", adVarChar, adParamInput, 40, IIf(oTabla.Amaterno = "", Null, oTabla.Amaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Pnombre", adVarChar, adParamInput, 35, IIf(oTabla.Pnombre = "", Null, oTabla.Pnombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Onombre", adVarChar, adParamInput, 35, IIf(oTabla.Onombre = "", Null, oTabla.Onombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fnacimiento", adVarChar, adParamInput, 10, IIf(oTabla.Fnacimiento = "", Null, oTabla.Fnacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Autogenerado", adVarChar, adParamInput, 20, IIf(oTabla.Autogenerado = "", Null, oTabla.Autogenerado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoTipo", adVarChar, adParamInput, 1, IIf(oTabla.DocumentoTipo = "", Null, oTabla.DocumentoTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoNumero", adVarChar, adParamInput, 10, IIf(oTabla.DocumentoNumero = "", Null, oTabla.DocumentoNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EstablecimientoCategoria", adVarChar, adParamInput, 2, IIf(oTabla.EstablecimientoCategoria = "", Null, oTabla.EstablecimientoCategoria))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CostoServicio", adCurrency, adParamInput, 0, oTabla.CostoServicio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CostoMedicamento", adCurrency, adParamInput, 0, oTabla.CostoMedicamento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CostoProcedimiento", adCurrency, adParamInput, 0, oTabla.CostoProcedimiento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CostoInsumo", adCurrency, adParamInput, 0, oTabla.CostoInsumo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MedicoDocumentoTipo", adVarChar, adParamInput, 2, IIf(oTabla.MedicoDocumentoTipo = "", Null, oTabla.MedicoDocumentoTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ate_grupoRiesgo", adVarChar, adParamInput, 2, IIf(oTabla.ate_grupoRiesgo = "", Null, oTabla.ate_grupoRiesgo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabCodigoPuntoDigitacion", adInteger, adParamInput, 0, IIf(oTabla.CabCodigoPuntoDigitacion = 0, Null, oTabla.CabCodigoPuntoDigitacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabCodigoUDR", adVarChar, adParamInput, 3, IIf(oTabla.CabCodigoUDR = "", Null, oTabla.CabCodigoUDR))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabNroEnvioAlSIS", adVarChar, adParamInput, 2, IIf(oTabla.CabNroEnvioAlSIS = "", Null, oTabla.CabNroEnvioAlSIS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabOrigenDelRegistro", adVarChar, adParamInput, 9, IIf(oTabla.CabOrigenDelRegistro = "", Null, oTabla.CabOrigenDelRegistro))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabVersionAplicativo", adVarChar, adParamInput, 9, IIf(oTabla.CabVersionAplicativo = "", Null, oTabla.CabVersionAplicativo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabIdentificacionPaquete", adInteger, adParamInput, 0, oTabla.CabIdentificacionPaquete)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdentificacionArfsis", adInteger, adParamInput, 0, IIf(oTabla.IdentificacionArfsis = 0, Null, oTabla.IdentificacionArfsis))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabFechaFuaPrimeraVez", adVarChar, adParamInput, 10, IIf(oTabla.CabFechaFuaPrimeraVez = "", Null, oTabla.CabFechaFuaPrimeraVez))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PeriodoOrigen", adVarChar, adParamInput, 6, IIf(oTabla.PeriodoOrigen = "", Null, oTabla.PeriodoOrigen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuacolegioCodigo", adVarChar, adParamInput, 20, IIf(oTabla.FuacolegioCodigo = "", Null, oTabla.FuacolegioCodigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuacolegioNivel", adVarChar, adParamInput, 1, IIf(oTabla.FuacolegioNivel = "", Null, oTabla.FuacolegioNivel))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuacolegioGrado", adVarChar, adParamInput, 1, IIf(oTabla.FuacolegioGrado = "", Null, oTabla.FuacolegioGrado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuacolegioSeccion", adVarChar, adParamInput, 5, IIf(oTabla.FuacolegioSeccion = "", Null, oTabla.FuacolegioSeccion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuacolegioTurno", adVarChar, adParamInput, 1, IIf(oTabla.FuacolegioTurno = "", Null, oTabla.FuacolegioTurno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Fuaetnia", adVarChar, adParamInput, 2, IIf(oTabla.Fuaetnia = "", Null, oTabla.Fuaetnia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuafechaFallecimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FuafechaFallecimiento = 0, Null, oTabla.FuafechaFallecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaUPS", adVarChar, adParamInput, 6, IIf(oTabla.FuaUPS = "", Null, oTabla.FuaUPS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaCodAutorizacion", adVarChar, adParamInput, 50, IIf(oTabla.FuaCodAutorizacion = "", Null, oTabla.FuaCodAutorizacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaFechaCorteAdm", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FuaFechaCorteAdm = 0, Null, oTabla.FuaFechaCorteAdm))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaVersionFormato", adVarChar, adParamInput, 2, IIf(oTabla.FuaVersionFormato = "", Null, oTabla.FuaVersionFormato))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaTipoAnexo2015", adInteger, adParamInput, 0, IIf(oTabla.FuaTipoAnexo2015 = 0, Null, oTabla.FuaTipoAnexo2015))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaCodOferFlexible", adVarChar, adParamInput, 20, IIf(oTabla.FuaCodOferFlexible = "", Null, oTabla.FuaCodOferFlexible))
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
'   Descripción:    elimia o borra un registro a la tabla SisFuaAtencion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoSisFuaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaAtencionEliminar"
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, oTabla.IdCuentaAtencion)
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
'   Descripción:    Selecciona un unico registro de la tabla SisFuaAtencion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoSisFuaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "SisFuaAtencionSeleccionarPorId"
     Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, oTabla.IdCuentaAtencion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.FuaDisa = IIf(IsNull(oRecordset!FuaDisa), "", oRecordset!FuaDisa)
       oTabla.FuaLote = IIf(IsNull(oRecordset!FuaLote), "", oRecordset!FuaLote)
       oTabla.FuaNumero = IIf(IsNull(oRecordset!FuaNumero), "", oRecordset!FuaNumero)
       oTabla.EstablecimientoCodigoRenaes = IIf(IsNull(oRecordset!EstablecimientoCodigoRenaes), "", oRecordset!EstablecimientoCodigoRenaes)
       oTabla.Reconsideracion = IIf(IsNull(oRecordset!Reconsideracion), "", oRecordset!Reconsideracion)
       oTabla.ReconsideracionCodigoDisa = IIf(IsNull(oRecordset!ReconsideracionCodigoDisa), "", oRecordset!ReconsideracionCodigoDisa)
       oTabla.ReconsideracionLote = IIf(IsNull(oRecordset!ReconsideracionLote), "", oRecordset!ReconsideracionLote)
       oTabla.ReconsideracionNroFormato = IIf(IsNull(oRecordset!ReconsideracionNroFormato), "", oRecordset!ReconsideracionNroFormato)
       oTabla.FuaComponente = IIf(IsNull(oRecordset!FuaComponente), "", oRecordset!FuaComponente)
       oTabla.Situacion = IIf(IsNull(oRecordset!Situacion), "", oRecordset!Situacion)
       oTabla.AfiliacionDisa = IIf(IsNull(oRecordset!AfiliacionDisa), "", oRecordset!AfiliacionDisa)
       oTabla.AfiliacionTipoFormato = IIf(IsNull(oRecordset!AfiliacionTipoFormato), "", oRecordset!AfiliacionTipoFormato)
       oTabla.AfiliacionNroFormato = IIf(IsNull(oRecordset!AfiliacionNroFormato), "", oRecordset!AfiliacionNroFormato)
       oTabla.CodigoTipoFormato = IIf(IsNull(oRecordset!CodigoTipoFormato), "", oRecordset!CodigoTipoFormato)
       oTabla.OrigenAseguradoInstitucion = IIf(IsNull(oRecordset!OrigenAseguradoInstitucion), "", oRecordset!OrigenAseguradoInstitucion)
       oTabla.OrigenAseguradoCodigo = IIf(IsNull(oRecordset!OrigenAseguradoCodigo), "", oRecordset!OrigenAseguradoCodigo)
       oTabla.Edad = IIf(IsNull(oRecordset!Edad), 0, oRecordset!Edad)
       oTabla.GrupoEtareo = IIf(IsNull(oRecordset!GrupoEtareo), "", oRecordset!GrupoEtareo)
       oTabla.Genero = IIf(IsNull(oRecordset!Genero), "", oRecordset!Genero)
       oTabla.FuaAtencion = IIf(IsNull(oRecordset!FuaAtencion), 0, oRecordset!FuaAtencion)
       oTabla.FuaCondicionMaterna = IIf(IsNull(oRecordset!FuaCondicionMaterna), "", oRecordset!FuaCondicionMaterna)
       oTabla.FuaNrohistoria = IIf(IsNull(oRecordset!FuaNrohistoria), "", oRecordset!FuaNrohistoria)
       oTabla.FuaConceptoPr = IIf(IsNull(oRecordset!FuaConceptoPr), 0, oRecordset!FuaConceptoPr)
       oTabla.FuaConceptoPrAutoriz = IIf(IsNull(oRecordset!FuaConceptoPrAutoriz), "", oRecordset!FuaConceptoPrAutoriz)
       oTabla.FuaConceptoPrMonto = oRecordset!FuaConceptoPrMonto
       oTabla.FuaAtencionFecha = IIf(IsNull(oRecordset!FuaAtencionFecha), "", oRecordset!FuaAtencionFecha)
       oTabla.FuaAtencionHora = IIf(IsNull(oRecordset!FuaAtencionHora), "", oRecordset!FuaAtencionHora)
       oTabla.FuaReferidoOrigenCodigoRENAES = IIf(IsNull(oRecordset!FuaReferidoOrigenCodigoRENAES), "", oRecordset!FuaReferidoOrigenCodigoRENAES)
       oTabla.FuaReferidoOrigenNreferencia = IIf(IsNull(oRecordset!FuaReferidoOrigenNreferencia), "", oRecordset!FuaReferidoOrigenNreferencia)
       oTabla.FuaCodigoPrestacion = IIf(IsNull(oRecordset!FuaCodigoPrestacion), "", oRecordset!FuaCodigoPrestacion)
       oTabla.FuaPersonalQatiende = IIf(IsNull(oRecordset!FuaPersonalQatiende), 0, oRecordset!FuaPersonalQatiende)
       oTabla.FuaAtencionLugar = IIf(IsNull(oRecordset!FuaAtencionLugar), "", oRecordset!FuaAtencionLugar)
       oTabla.FuaDestino = IIf(IsNull(oRecordset!FuaDestino), "", oRecordset!FuaDestino)
       oTabla.FuaHospitalizadoFingreso = IIf(IsNull(oRecordset!FuaHospitalizadoFingreso), "", oRecordset!FuaHospitalizadoFingreso)
       oTabla.FuaHospitalizadoFalta = IIf(IsNull(oRecordset!FuaHospitalizadoFalta), "", oRecordset!FuaHospitalizadoFalta)
       oTabla.FuaReferidoDestinoCodigoRENAES = IIf(IsNull(oRecordset!FuaReferidoDestinoCodigoRENAES), "", oRecordset!FuaReferidoDestinoCodigoRENAES)
       oTabla.FuaReferidoDestinoNreferencia = IIf(IsNull(oRecordset!FuaReferidoDestinoNreferencia), "", oRecordset!FuaReferidoDestinoNreferencia)
       oTabla.FuaMedicoDNI = IIf(IsNull(oRecordset!FuaMedicoDNI), "", oRecordset!FuaMedicoDNI)
       oTabla.FuaMedico = IIf(IsNull(oRecordset!FuaMedico), "", oRecordset!FuaMedico)
       oTabla.FuaMedicoTipo = IIf(IsNull(oRecordset!FuaMedicoTipo), "", oRecordset!FuaMedicoTipo)
       oTabla.AfiliacionNroIntegrante = IIf(IsNull(oRecordset!AfiliacionNroIntegrante), "", oRecordset!AfiliacionNroIntegrante)
       oTabla.codigo = IIf(IsNull(oRecordset!codigo), "", oRecordset!codigo)
       oTabla.idSiasis = IIf(IsNull(oRecordset!idSiasis), "", oRecordset!idSiasis)
       oTabla.FuaObservaciones = IIf(IsNull(oRecordset!FuaObservaciones), "", oRecordset!FuaObservaciones)
       oTabla.CabDniUsuarioRegistra = IIf(IsNull(oRecordset!CabDniUsuarioRegistra), "", oRecordset!CabDniUsuarioRegistra)
       oTabla.UltimaFechaAddMod = IIf(IsNull(oRecordset!UltimaFechaAddMod), "", oRecordset!UltimaFechaAddMod)
       oTabla.CabEstado = IIf(IsNull(oRecordset!CabEstado), "", oRecordset!CabEstado)
       oTabla.FuaFechaParto = IIf(IsNull(oRecordset!FuaFechaParto), "", oRecordset!FuaFechaParto)
       oTabla.EstablecimientoDistrito = IIf(IsNull(oRecordset!EstablecimientoDistrito), "", oRecordset!EstablecimientoDistrito)
       oTabla.Anio = IIf(IsNull(oRecordset!Anio), "", oRecordset!Anio)
       oTabla.Mes = IIf(IsNull(oRecordset!Mes), "", oRecordset!Mes)
       oTabla.CostoTotal = oRecordset!CostoTotal
       oTabla.Apaterno = IIf(IsNull(oRecordset!Apaterno), "", oRecordset!Apaterno)
       oTabla.Amaterno = IIf(IsNull(oRecordset!Amaterno), "", oRecordset!Amaterno)
       oTabla.Pnombre = IIf(IsNull(oRecordset!Pnombre), "", oRecordset!Pnombre)
       oTabla.Onombre = IIf(IsNull(oRecordset!Onombre), "", oRecordset!Onombre)
       oTabla.Fnacimiento = IIf(IsNull(oRecordset!Fnacimiento), "", oRecordset!Fnacimiento)
       oTabla.Autogenerado = IIf(IsNull(oRecordset!Autogenerado), "", oRecordset!Autogenerado)
       oTabla.DocumentoTipo = IIf(IsNull(oRecordset!DocumentoTipo), "", oRecordset!DocumentoTipo)
       oTabla.DocumentoNumero = IIf(IsNull(oRecordset!DocumentoNumero), "", oRecordset!DocumentoNumero)
       oTabla.EstablecimientoCategoria = IIf(IsNull(oRecordset!EstablecimientoCategoria), "", oRecordset!EstablecimientoCategoria)
       oTabla.CostoServicio = oRecordset!CostoServicio
       oTabla.CostoMedicamento = oRecordset!CostoMedicamento
       oTabla.CostoProcedimiento = oRecordset!CostoProcedimiento
       oTabla.CostoInsumo = oRecordset!CostoInsumo
       oTabla.MedicoDocumentoTipo = IIf(IsNull(oRecordset!MedicoDocumentoTipo), "", oRecordset!MedicoDocumentoTipo)
       oTabla.ate_grupoRiesgo = IIf(IsNull(oRecordset!ate_grupoRiesgo), "", oRecordset!ate_grupoRiesgo)
       oTabla.CabCodigoPuntoDigitacion = IIf(IsNull(oRecordset!CabCodigoPuntoDigitacion), 0, oRecordset!CabCodigoPuntoDigitacion)
       oTabla.CabCodigoUDR = IIf(IsNull(oRecordset!CabCodigoUDR), "", oRecordset!CabCodigoUDR)
       oTabla.CabNroEnvioAlSIS = IIf(IsNull(oRecordset!CabNroEnvioAlSIS), "", oRecordset!CabNroEnvioAlSIS)
       oTabla.CabOrigenDelRegistro = IIf(IsNull(oRecordset!CabOrigenDelRegistro), "", oRecordset!CabOrigenDelRegistro)
       oTabla.CabVersionAplicativo = IIf(IsNull(oRecordset!CabVersionAplicativo), "", oRecordset!CabVersionAplicativo)
       oTabla.CabIdentificacionPaquete = IIf(IsNull(oRecordset!CabIdentificacionPaquete), 0, oRecordset!CabIdentificacionPaquete)
       oTabla.IdentificacionArfsis = IIf(IsNull(oRecordset!IdentificacionArfsis), 0, oRecordset!IdentificacionArfsis)
       oTabla.CabFechaFuaPrimeraVez = IIf(IsNull(oRecordset!CabFechaFuaPrimeraVez), "", oRecordset!CabFechaFuaPrimeraVez)
       oTabla.PeriodoOrigen = IIf(IsNull(oRecordset!PeriodoOrigen), "", oRecordset!PeriodoOrigen)
       oTabla.FuacolegioCodigo = IIf(IsNull(oRecordset!FuacolegioCodigo), "", oRecordset!FuacolegioCodigo)
       oTabla.FuacolegioNivel = IIf(IsNull(oRecordset!FuacolegioNivel), "", oRecordset!FuacolegioNivel)
       oTabla.FuacolegioGrado = IIf(IsNull(oRecordset!FuacolegioGrado), "", oRecordset!FuacolegioGrado)
       oTabla.FuacolegioSeccion = IIf(IsNull(oRecordset!FuacolegioSeccion), "", oRecordset!FuacolegioSeccion)
       oTabla.FuacolegioTurno = IIf(IsNull(oRecordset!FuacolegioTurno), "", oRecordset!FuacolegioTurno)
       oTabla.Fuaetnia = IIf(IsNull(oRecordset!Fuaetnia), "", oRecordset!Fuaetnia)
       oTabla.FuafechaFallecimiento = IIf(IsNull(oRecordset!FuafechaFallecimiento), 0, oRecordset!FuafechaFallecimiento)
       oTabla.FuaUPS = IIf(IsNull(oRecordset!FuaUPS), "", oRecordset!FuaUPS)
       oTabla.FuaCodAutorizacion = IIf(IsNull(oRecordset!FuaCodAutorizacion), "", oRecordset!FuaCodAutorizacion)
       oTabla.FuaFechaCorteAdm = IIf(IsNull(oRecordset!FuaFechaCorteAdm), 0, oRecordset!FuaFechaCorteAdm)
       oTabla.FuaVersionFormato = IIf(IsNull(oRecordset!FuaVersionFormato), "", oRecordset!FuaVersionFormato)
       oTabla.FuaTipoAnexo2015 = IIf(IsNull(oRecordset!FuaTipoAnexo2015), 0, oRecordset!FuaTipoAnexo2015)
       oTabla.FuaCodOferFlexible = IIf(IsNull(oRecordset!FuaCodOferFlexible), "", oRecordset!FuaCodOferFlexible)
       
       SeleccionarPorId = True
       ms_MensajeError = ""
   Else
       oTabla.FuaHospitalizadoFalta = SIGHEntidades.FECHA_VACIA_DMY
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoSisFuaAtencion)
   oTabla.IdCuentaAtencion = 0
   oTabla.FuaDisa = ""
   oTabla.FuaLote = ""
   oTabla.FuaNumero = ""
   oTabla.EstablecimientoCodigoRenaes = ""
   oTabla.Reconsideracion = ""
   oTabla.ReconsideracionCodigoDisa = ""
   oTabla.ReconsideracionLote = ""
   oTabla.ReconsideracionNroFormato = ""
   oTabla.FuaComponente = ""
   oTabla.Situacion = ""
   oTabla.AfiliacionDisa = ""
   oTabla.AfiliacionTipoFormato = ""
   oTabla.AfiliacionNroFormato = ""
   oTabla.CodigoTipoFormato = ""
   oTabla.OrigenAseguradoInstitucion = ""
   oTabla.OrigenAseguradoCodigo = ""
   oTabla.Edad = 0
   oTabla.GrupoEtareo = ""
   oTabla.Genero = ""
   oTabla.FuaAtencion = 0
   oTabla.FuaCondicionMaterna = ""
   oTabla.FuaNrohistoria = ""
   oTabla.FuaConceptoPr = 0
   oTabla.FuaConceptoPrAutoriz = ""
   oTabla.FuaConceptoPrMonto = 0
   oTabla.FuaAtencionFecha = ""
   oTabla.FuaAtencionHora = ""
   oTabla.FuaReferidoOrigenCodigoRENAES = ""
   oTabla.FuaReferidoOrigenNreferencia = ""
   oTabla.FuaCodigoPrestacion = ""
   oTabla.FuaPersonalQatiende = 0
   oTabla.FuaAtencionLugar = ""
   oTabla.FuaDestino = ""
   oTabla.FuaHospitalizadoFingreso = ""
   oTabla.FuaHospitalizadoFalta = ""
   oTabla.FuaReferidoDestinoCodigoRENAES = ""
   oTabla.FuaReferidoDestinoNreferencia = ""
   oTabla.FuaMedicoDNI = ""
   oTabla.FuaMedico = ""
   oTabla.FuaMedicoTipo = ""
   oTabla.AfiliacionNroIntegrante = ""
   oTabla.codigo = ""
   oTabla.idSiasis = ""
   oTabla.FuaObservaciones = ""
   oTabla.CabDniUsuarioRegistra = ""
   oTabla.UltimaFechaAddMod = ""
   oTabla.CabEstado = ""
   oTabla.FuaFechaParto = ""
   oTabla.EstablecimientoDistrito = ""
   oTabla.Anio = ""
   oTabla.Mes = ""
   oTabla.CostoTotal = 0
   oTabla.Apaterno = ""
   oTabla.Amaterno = ""
   oTabla.Pnombre = ""
   oTabla.Onombre = ""
   oTabla.Fnacimiento = ""
   oTabla.Autogenerado = ""
   oTabla.DocumentoTipo = ""
   oTabla.DocumentoNumero = ""
   oTabla.EstablecimientoCategoria = ""
   oTabla.CostoServicio = 0
   oTabla.CostoMedicamento = 0
   oTabla.CostoProcedimiento = 0
   oTabla.CostoInsumo = 0
   oTabla.MedicoDocumentoTipo = ""
   oTabla.ate_grupoRiesgo = ""
   oTabla.CabCodigoPuntoDigitacion = 0
   oTabla.CabCodigoUDR = ""
   oTabla.CabNroEnvioAlSIS = ""
   oTabla.CabOrigenDelRegistro = ""
   oTabla.CabVersionAplicativo = ""
   oTabla.CabIdentificacionPaquete = 0
   oTabla.IdentificacionArfsis = 0
   oTabla.CabFechaFuaPrimeraVez = ""
   oTabla.PeriodoOrigen = ""
   oTabla.FuacolegioCodigo = ""
   oTabla.FuacolegioNivel = ""
   oTabla.FuacolegioGrado = ""
   oTabla.FuacolegioSeccion = ""
   oTabla.FuacolegioTurno = ""
   oTabla.Fuaetnia = ""
   oTabla.FuafechaFallecimiento = Now
   oTabla.FuaUPS = ""
   oTabla.FuaCodAutorizacion = ""
   oTabla.FuaFechaCorteAdm = Now
   oTabla.FuaVersionFormato = ""
   oTabla.FuaTipoAnexo2015 = 0
   oTabla.FuaCodOferFlexible = ""
End Sub


