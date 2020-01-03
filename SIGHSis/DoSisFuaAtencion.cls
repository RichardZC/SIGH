VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoSisFuaAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para estructura de la tabla SISFuaAtencion
'        Programado por: Barrantes D
'        Fecha: Enero 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdCuentaAtencion As Long
Dim ms_FuaDisa As String
Dim ms_FuaLote As String
Dim ms_FuaNumero As String
Dim ms_EstablecimientoCodigoRenaes As String
Dim ms_Reconsideracion As String
Dim ms_ReconsideracionCodigoDisa As String
Dim ms_ReconsideracionLote As String
Dim ms_ReconsideracionNroFormato As String
Dim ms_FuaComponente As String
Dim ms_Situacion As String
Dim ms_AfiliacionDisa As String
Dim ms_AfiliacionTipoFormato As String
Dim ms_AfiliacionNroFormato As String
Dim ms_CodigoTipoFormato As String
Dim ms_OrigenAseguradoInstitucion As String
Dim ms_OrigenAseguradoCodigo As String
Dim ml_Edad As Long
Dim ms_GrupoEtareo As String
Dim ms_Genero As String
Dim ml_FuaAtencion As Long
Dim ms_FuaCondicionMaterna As String
Dim ms_FuaNrohistoria As String
Dim ml_FuaConceptoPr As Long
Dim ms_FuaConceptoPrAutoriz As String
Dim mc_FuaConceptoPrMonto As Currency
Dim ms_FuaAtencionFecha As String
Dim ms_FuaAtencionHora As String
Dim ms_FuaReferidoOrigenCodigoRENAES As String
Dim ms_FuaReferidoOrigenNreferencia As String
Dim ms_FuaCodigoPrestacion As String
Dim ml_FuaPersonalQatiende As Long
Dim ms_FuaAtencionLugar As String
Dim ms_FuaDestino As String
Dim ms_FuaHospitalizadoFingreso As String
Dim ms_FuaHospitalizadoFalta As String
Dim ms_FuaReferidoDestinoCodigoRENAES As String
Dim ms_FuaReferidoDestinoNreferencia As String
Dim ms_FuaMedicoDNI As String
Dim ms_FuaMedico As String
Dim ms_FuaMedicoTipo As String
Dim ms_AfiliacionNroIntegrante As String
Dim ms_codigo As String
Dim ms_idSiasis As String
Dim ms_FuaObservaciones As String
Dim ms_CabDniUsuarioRegistra As String
Dim ms_UltimaFechaAddMod As String
Dim ms_CabEstado As String
Dim ms_FuaFechaParto As String
Dim ms_EstablecimientoDistrito As String
Dim ms_Anio As String
Dim ms_Mes As String
Dim mc_CostoTotal As Currency
Dim ms_Apaterno As String
Dim ms_Amaterno As String
Dim ms_Pnombre As String
Dim ms_Onombre As String
Dim ms_fnacimiento As String
Dim ms_Autogenerado As String
Dim ms_DocumentoTipo As String
Dim ms_DocumentoNumero As String
Dim ms_EstablecimientoCategoria As String
Dim mc_CostoServicio As Currency
Dim mc_CostoMedicamento As Currency
Dim mc_CostoProcedimiento As Currency
Dim mc_CostoInsumo As Currency
Dim ms_MedicoDocumentoTipo As String
Dim ms_ate_grupoRiesgo As String
Dim ml_CabCodigoPuntoDigitacion As Long
Dim ms_CabCodigoUDR As String
Dim ms_CabNroEnvioAlSIS As String
Dim ms_CabOrigenDelRegistro As String
Dim ms_CabVersionAplicativo As String
Dim ml_CabIdentificacionPaquete As Long
Dim ml_IdentificacionArfsis As Long
Dim ms_CabFechaFuaPrimeraVez As String
Dim ms_PeriodoOrigen As String
Dim ms_FuacolegioCodigo As String
Dim ms_FuacolegioNivel As String
Dim ms_FuacolegioGrado As String
Dim ms_FuacolegioSeccion As String
Dim ms_FuacolegioTurno As String
Dim ms_Fuaetnia As String
Dim mda_FuafechaFallecimiento As Date
Dim ms_FuaUPS As String
Dim ms_FuaCodAutorizacion As String
Dim mda_FuaFechaCorteAdm As Date
Dim ms_FuaVersionFormato As String
Dim ml_FuaTipoAnexo2015 As Long
Dim ms_FuaCodOferFlexible As String

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdCuentaAtencion(iValue As Long)
   ml_IdCuentaAtencion = iValue
End Property
Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
End Property
Property Let FuaDisa(sValue As String)
   ms_FuaDisa = sValue
End Property
Property Get FuaDisa() As String
   FuaDisa = ms_FuaDisa
End Property
Property Let FuaLote(sValue As String)
   ms_FuaLote = sValue
End Property
Property Get FuaLote() As String
   FuaLote = ms_FuaLote
End Property
Property Let FuaNumero(sValue As String)
   ms_FuaNumero = sValue
End Property
Property Get FuaNumero() As String
   FuaNumero = ms_FuaNumero
End Property
Property Let EstablecimientoCodigoRenaes(sValue As String)
   ms_EstablecimientoCodigoRenaes = sValue
End Property
Property Get EstablecimientoCodigoRenaes() As String
   EstablecimientoCodigoRenaes = ms_EstablecimientoCodigoRenaes
End Property
Property Let Reconsideracion(sValue As String)
   ms_Reconsideracion = sValue
End Property
Property Get Reconsideracion() As String
   Reconsideracion = ms_Reconsideracion
End Property
Property Let ReconsideracionCodigoDisa(sValue As String)
   ms_ReconsideracionCodigoDisa = sValue
End Property
Property Get ReconsideracionCodigoDisa() As String
   ReconsideracionCodigoDisa = ms_ReconsideracionCodigoDisa
End Property
Property Let ReconsideracionLote(sValue As String)
   ms_ReconsideracionLote = sValue
End Property
Property Get ReconsideracionLote() As String
   ReconsideracionLote = ms_ReconsideracionLote
End Property
Property Let ReconsideracionNroFormato(sValue As String)
   ms_ReconsideracionNroFormato = sValue
End Property
Property Get ReconsideracionNroFormato() As String
   ReconsideracionNroFormato = ms_ReconsideracionNroFormato
End Property
Property Let FuaComponente(sValue As String)
   ms_FuaComponente = sValue
End Property
Property Get FuaComponente() As String
   FuaComponente = ms_FuaComponente
End Property
Property Let Situacion(sValue As String)
   ms_Situacion = sValue
End Property
Property Get Situacion() As String
   Situacion = ms_Situacion
End Property
Property Let AfiliacionDisa(sValue As String)
   ms_AfiliacionDisa = sValue
End Property
Property Get AfiliacionDisa() As String
   AfiliacionDisa = ms_AfiliacionDisa
End Property
Property Let AfiliacionTipoFormato(sValue As String)
   ms_AfiliacionTipoFormato = sValue
End Property
Property Get AfiliacionTipoFormato() As String
   AfiliacionTipoFormato = ms_AfiliacionTipoFormato
End Property
Property Let AfiliacionNroFormato(sValue As String)
   ms_AfiliacionNroFormato = sValue
End Property
Property Get AfiliacionNroFormato() As String
   AfiliacionNroFormato = ms_AfiliacionNroFormato
End Property
Property Let CodigoTipoFormato(sValue As String)
   ms_CodigoTipoFormato = sValue
End Property
Property Get CodigoTipoFormato() As String
   CodigoTipoFormato = ms_CodigoTipoFormato
End Property
Property Let OrigenAseguradoInstitucion(sValue As String)
   ms_OrigenAseguradoInstitucion = sValue
End Property
Property Get OrigenAseguradoInstitucion() As String
   OrigenAseguradoInstitucion = ms_OrigenAseguradoInstitucion
End Property
Property Let OrigenAseguradoCodigo(sValue As String)
   ms_OrigenAseguradoCodigo = sValue
End Property
Property Get OrigenAseguradoCodigo() As String
   OrigenAseguradoCodigo = ms_OrigenAseguradoCodigo
End Property
Property Let Edad(iValue As Long)
   ml_Edad = iValue
End Property
Property Get Edad() As Long
   Edad = ml_Edad
End Property
Property Let GrupoEtareo(sValue As String)
   ms_GrupoEtareo = sValue
End Property
Property Get GrupoEtareo() As String
   GrupoEtareo = ms_GrupoEtareo
End Property
Property Let Genero(sValue As String)
   ms_Genero = sValue
End Property
Property Get Genero() As String
   Genero = ms_Genero
End Property
Property Let FuaAtencion(iValue As Long)
   ml_FuaAtencion = iValue
End Property
Property Get FuaAtencion() As Long
   FuaAtencion = ml_FuaAtencion
End Property
Property Let FuaCondicionMaterna(sValue As String)
   ms_FuaCondicionMaterna = sValue
End Property
Property Get FuaCondicionMaterna() As String
   FuaCondicionMaterna = ms_FuaCondicionMaterna
End Property
Property Let FuaNrohistoria(sValue As String)
   ms_FuaNrohistoria = sValue
End Property
Property Get FuaNrohistoria() As String
   FuaNrohistoria = ms_FuaNrohistoria
End Property
Property Let FuaConceptoPr(iValue As Long)
   ml_FuaConceptoPr = iValue
End Property
Property Get FuaConceptoPr() As Long
   FuaConceptoPr = ml_FuaConceptoPr
End Property
Property Let FuaConceptoPrAutoriz(sValue As String)
   ms_FuaConceptoPrAutoriz = sValue
End Property
Property Get FuaConceptoPrAutoriz() As String
   FuaConceptoPrAutoriz = ms_FuaConceptoPrAutoriz
End Property
Property Let FuaConceptoPrMonto(cValue As Currency)
   mc_FuaConceptoPrMonto = cValue
End Property
Property Get FuaConceptoPrMonto() As Currency
   FuaConceptoPrMonto = mc_FuaConceptoPrMonto
End Property
Property Let FuaAtencionFecha(sValue As String)
   ms_FuaAtencionFecha = sValue
End Property
Property Get FuaAtencionFecha() As String
   FuaAtencionFecha = ms_FuaAtencionFecha
End Property
Property Let FuaAtencionHora(sValue As String)
   ms_FuaAtencionHora = sValue
End Property
Property Get FuaAtencionHora() As String
   FuaAtencionHora = ms_FuaAtencionHora
End Property
Property Let FuaReferidoOrigenCodigoRENAES(sValue As String)
   ms_FuaReferidoOrigenCodigoRENAES = sValue
End Property
Property Get FuaReferidoOrigenCodigoRENAES() As String
   FuaReferidoOrigenCodigoRENAES = ms_FuaReferidoOrigenCodigoRENAES
End Property
Property Let FuaReferidoOrigenNreferencia(sValue As String)
   ms_FuaReferidoOrigenNreferencia = sValue
End Property
Property Get FuaReferidoOrigenNreferencia() As String
   FuaReferidoOrigenNreferencia = ms_FuaReferidoOrigenNreferencia
End Property
Property Let FuaCodigoPrestacion(sValue As String)
   ms_FuaCodigoPrestacion = sValue
End Property
Property Get FuaCodigoPrestacion() As String
   FuaCodigoPrestacion = ms_FuaCodigoPrestacion
End Property
Property Let FuaPersonalQatiende(iValue As Long)
   ml_FuaPersonalQatiende = iValue
End Property
Property Get FuaPersonalQatiende() As Long
   FuaPersonalQatiende = ml_FuaPersonalQatiende
End Property
Property Let FuaAtencionLugar(sValue As String)
   ms_FuaAtencionLugar = sValue
End Property
Property Get FuaAtencionLugar() As String
   FuaAtencionLugar = ms_FuaAtencionLugar
End Property
Property Let FuaDestino(sValue As String)
   ms_FuaDestino = sValue
End Property
Property Get FuaDestino() As String
   FuaDestino = ms_FuaDestino
End Property
Property Let FuaHospitalizadoFingreso(sValue As String)
   ms_FuaHospitalizadoFingreso = sValue
End Property
Property Get FuaHospitalizadoFingreso() As String
   FuaHospitalizadoFingreso = ms_FuaHospitalizadoFingreso
End Property
Property Let FuaHospitalizadoFalta(sValue As String)
   ms_FuaHospitalizadoFalta = sValue
End Property
Property Get FuaHospitalizadoFalta() As String
   FuaHospitalizadoFalta = ms_FuaHospitalizadoFalta
End Property
Property Let FuaReferidoDestinoCodigoRENAES(sValue As String)
   ms_FuaReferidoDestinoCodigoRENAES = sValue
End Property
Property Get FuaReferidoDestinoCodigoRENAES() As String
   FuaReferidoDestinoCodigoRENAES = ms_FuaReferidoDestinoCodigoRENAES
End Property
Property Let FuaReferidoDestinoNreferencia(sValue As String)
   ms_FuaReferidoDestinoNreferencia = sValue
End Property
Property Get FuaReferidoDestinoNreferencia() As String
   FuaReferidoDestinoNreferencia = ms_FuaReferidoDestinoNreferencia
End Property
Property Let FuaMedicoDNI(sValue As String)
   ms_FuaMedicoDNI = sValue
End Property
Property Get FuaMedicoDNI() As String
   FuaMedicoDNI = ms_FuaMedicoDNI
End Property
Property Let FuaMedico(sValue As String)
   ms_FuaMedico = sValue
End Property
Property Get FuaMedico() As String
   FuaMedico = ms_FuaMedico
End Property
Property Let FuaMedicoTipo(sValue As String)
   ms_FuaMedicoTipo = sValue
End Property
Property Get FuaMedicoTipo() As String
   FuaMedicoTipo = ms_FuaMedicoTipo
End Property
Property Let AfiliacionNroIntegrante(sValue As String)
   ms_AfiliacionNroIntegrante = sValue
End Property
Property Get AfiliacionNroIntegrante() As String
   AfiliacionNroIntegrante = ms_AfiliacionNroIntegrante
End Property
Property Let codigo(sValue As String)
   ms_codigo = sValue
End Property
Property Get codigo() As String
   codigo = ms_codigo
End Property
Property Let idSiasis(sValue As String)
   ms_idSiasis = sValue
End Property
Property Get idSiasis() As String
   idSiasis = ms_idSiasis
End Property
Property Let FuaObservaciones(sValue As String)
   ms_FuaObservaciones = sValue
End Property
Property Get FuaObservaciones() As String
   FuaObservaciones = ms_FuaObservaciones
End Property
Property Let CabDniUsuarioRegistra(sValue As String)
   ms_CabDniUsuarioRegistra = sValue
End Property
Property Get CabDniUsuarioRegistra() As String
   CabDniUsuarioRegistra = ms_CabDniUsuarioRegistra
End Property
Property Let UltimaFechaAddMod(sValue As String)
   ms_UltimaFechaAddMod = sValue
End Property
Property Get UltimaFechaAddMod() As String
   UltimaFechaAddMod = ms_UltimaFechaAddMod
End Property
Property Let CabEstado(sValue As String)
   ms_CabEstado = sValue
End Property
Property Get CabEstado() As String
   CabEstado = ms_CabEstado
End Property
Property Let FuaFechaParto(sValue As String)
   ms_FuaFechaParto = sValue
End Property
Property Get FuaFechaParto() As String
   FuaFechaParto = ms_FuaFechaParto
End Property
Property Let EstablecimientoDistrito(sValue As String)
   ms_EstablecimientoDistrito = sValue
End Property
Property Get EstablecimientoDistrito() As String
   EstablecimientoDistrito = ms_EstablecimientoDistrito
End Property
Property Let Anio(sValue As String)
   ms_Anio = sValue
End Property
Property Get Anio() As String
   Anio = ms_Anio
End Property
Property Let Mes(sValue As String)
   ms_Mes = sValue
End Property
Property Get Mes() As String
   Mes = ms_Mes
End Property
Property Let CostoTotal(cValue As Currency)
   mc_CostoTotal = cValue
End Property
Property Get CostoTotal() As Currency
   CostoTotal = mc_CostoTotal
End Property
Property Let Apaterno(sValue As String)
   ms_Apaterno = sValue
End Property
Property Get Apaterno() As String
   Apaterno = ms_Apaterno
End Property
Property Let Amaterno(sValue As String)
   ms_Amaterno = sValue
End Property
Property Get Amaterno() As String
   Amaterno = ms_Amaterno
End Property
Property Let Pnombre(sValue As String)
   ms_Pnombre = sValue
End Property
Property Get Pnombre() As String
   Pnombre = ms_Pnombre
End Property
Property Let Onombre(sValue As String)
   ms_Onombre = sValue
End Property
Property Get Onombre() As String
   Onombre = ms_Onombre
End Property
Property Let Fnacimiento(sValue As String)
   ms_fnacimiento = sValue
End Property
Property Get Fnacimiento() As String
   Fnacimiento = ms_fnacimiento
End Property
Property Let Autogenerado(sValue As String)
   ms_Autogenerado = sValue
End Property
Property Get Autogenerado() As String
   Autogenerado = ms_Autogenerado
End Property
Property Let DocumentoTipo(sValue As String)
   ms_DocumentoTipo = sValue
End Property
Property Get DocumentoTipo() As String
   DocumentoTipo = ms_DocumentoTipo
End Property
Property Let DocumentoNumero(sValue As String)
   ms_DocumentoNumero = sValue
End Property
Property Get DocumentoNumero() As String
   DocumentoNumero = ms_DocumentoNumero
End Property
Property Let EstablecimientoCategoria(sValue As String)
   ms_EstablecimientoCategoria = sValue
End Property
Property Get EstablecimientoCategoria() As String
   EstablecimientoCategoria = ms_EstablecimientoCategoria
End Property
Property Let CostoServicio(cValue As Currency)
   mc_CostoServicio = cValue
End Property
Property Get CostoServicio() As Currency
   CostoServicio = mc_CostoServicio
End Property
Property Let CostoMedicamento(cValue As Currency)
   mc_CostoMedicamento = cValue
End Property
Property Get CostoMedicamento() As Currency
   CostoMedicamento = mc_CostoMedicamento
End Property
Property Let CostoProcedimiento(cValue As Currency)
   mc_CostoProcedimiento = cValue
End Property
Property Get CostoProcedimiento() As Currency
   CostoProcedimiento = mc_CostoProcedimiento
End Property
Property Let CostoInsumo(cValue As Currency)
   mc_CostoInsumo = cValue
End Property
Property Get CostoInsumo() As Currency
   CostoInsumo = mc_CostoInsumo
End Property
Property Let MedicoDocumentoTipo(sValue As String)
   ms_MedicoDocumentoTipo = sValue
End Property
Property Get MedicoDocumentoTipo() As String
   MedicoDocumentoTipo = ms_MedicoDocumentoTipo
End Property
Property Let ate_grupoRiesgo(sValue As String)
   ms_ate_grupoRiesgo = sValue
End Property
Property Get ate_grupoRiesgo() As String
   ate_grupoRiesgo = ms_ate_grupoRiesgo
End Property
Property Let CabCodigoPuntoDigitacion(iValue As Long)
   ml_CabCodigoPuntoDigitacion = iValue
End Property
Property Get CabCodigoPuntoDigitacion() As Long
   CabCodigoPuntoDigitacion = ml_CabCodigoPuntoDigitacion
End Property
Property Let CabCodigoUDR(sValue As String)
   ms_CabCodigoUDR = sValue
End Property
Property Get CabCodigoUDR() As String
   CabCodigoUDR = ms_CabCodigoUDR
End Property
Property Let CabNroEnvioAlSIS(sValue As String)
   ms_CabNroEnvioAlSIS = sValue
End Property
Property Get CabNroEnvioAlSIS() As String
   CabNroEnvioAlSIS = ms_CabNroEnvioAlSIS
End Property
Property Let CabOrigenDelRegistro(sValue As String)
   ms_CabOrigenDelRegistro = sValue
End Property
Property Get CabOrigenDelRegistro() As String
   CabOrigenDelRegistro = ms_CabOrigenDelRegistro
End Property
Property Let CabVersionAplicativo(sValue As String)
   ms_CabVersionAplicativo = sValue
End Property
Property Get CabVersionAplicativo() As String
   CabVersionAplicativo = ms_CabVersionAplicativo
End Property
Property Let CabIdentificacionPaquete(iValue As Long)
   ml_CabIdentificacionPaquete = iValue
End Property
Property Get CabIdentificacionPaquete() As Long
   CabIdentificacionPaquete = ml_CabIdentificacionPaquete
End Property
Property Let IdentificacionArfsis(iValue As Long)
   ml_IdentificacionArfsis = iValue
End Property
Property Get IdentificacionArfsis() As Long
   IdentificacionArfsis = ml_IdentificacionArfsis
End Property
Property Let CabFechaFuaPrimeraVez(sValue As String)
   ms_CabFechaFuaPrimeraVez = sValue
End Property
Property Get CabFechaFuaPrimeraVez() As String
   CabFechaFuaPrimeraVez = ms_CabFechaFuaPrimeraVez
End Property
Property Let PeriodoOrigen(sValue As String)
   ms_PeriodoOrigen = sValue
End Property
Property Get PeriodoOrigen() As String
   PeriodoOrigen = ms_PeriodoOrigen
End Property
Property Let FuacolegioCodigo(sValue As String)
   ms_FuacolegioCodigo = sValue
End Property
Property Get FuacolegioCodigo() As String
   FuacolegioCodigo = ms_FuacolegioCodigo
End Property
Property Let FuacolegioNivel(sValue As String)
   ms_FuacolegioNivel = sValue
End Property
Property Get FuacolegioNivel() As String
   FuacolegioNivel = ms_FuacolegioNivel
End Property
Property Let FuacolegioGrado(sValue As String)
   ms_FuacolegioGrado = sValue
End Property
Property Get FuacolegioGrado() As String
   FuacolegioGrado = ms_FuacolegioGrado
End Property
Property Let FuacolegioSeccion(sValue As String)
   ms_FuacolegioSeccion = sValue
End Property
Property Get FuacolegioSeccion() As String
   FuacolegioSeccion = ms_FuacolegioSeccion
End Property
Property Let FuacolegioTurno(sValue As String)
   ms_FuacolegioTurno = sValue
End Property
Property Get FuacolegioTurno() As String
   FuacolegioTurno = ms_FuacolegioTurno
End Property
Property Let Fuaetnia(sValue As String)
   ms_Fuaetnia = sValue
End Property
Property Get Fuaetnia() As String
   Fuaetnia = ms_Fuaetnia
End Property
Property Let FuafechaFallecimiento(daValue As Date)
   mda_FuafechaFallecimiento = daValue
End Property
Property Get FuafechaFallecimiento() As Date
   FuafechaFallecimiento = mda_FuafechaFallecimiento
End Property
Property Let FuaUPS(sValue As String)
   ms_FuaUPS = sValue
End Property
Property Get FuaUPS() As String
   FuaUPS = ms_FuaUPS
End Property
Property Let FuaCodAutorizacion(sValue As String)
   ms_FuaCodAutorizacion = sValue
End Property
Property Get FuaCodAutorizacion() As String
   FuaCodAutorizacion = ms_FuaCodAutorizacion
End Property
Property Let FuaFechaCorteAdm(daValue As Date)
   mda_FuaFechaCorteAdm = daValue
End Property
Property Get FuaFechaCorteAdm() As Date
   FuaFechaCorteAdm = mda_FuaFechaCorteAdm
End Property
Property Let FuaVersionFormato(sValue As String)
   ms_FuaVersionFormato = sValue
End Property
Property Get FuaVersionFormato() As String
   FuaVersionFormato = ms_FuaVersionFormato
End Property
Property Let FuaTipoAnexo2015(iValue As Long)
   ml_FuaTipoAnexo2015 = iValue
End Property
Property Get FuaTipoAnexo2015() As Long
   FuaTipoAnexo2015 = ml_FuaTipoAnexo2015
End Property
Property Let FuaCodOferFlexible(iValue As String)
   ms_FuaCodOferFlexible = iValue
End Property
Property Get FuaCodOferFlexible() As String
   FuaCodOferFlexible = ms_FuaCodOferFlexible
End Property
