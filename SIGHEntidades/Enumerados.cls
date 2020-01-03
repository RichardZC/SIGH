VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Enumerados"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Procesos Generales de Tipo enumerados
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit

Enum sghOpciones
    sghAgregar = 1
    sghModificar = 2
    sghConsultar = 3
    sghEliminar = 4
    sghBuscar = 6
    sghImprimir = 7
End Enum

Enum sghTipoNumeracionDeNroHistoria
     sghHistoriaDefinitivaAutomatica = 1
     sghHistoriaDefinitivaManual = 2
     sghHistoriaDefinitivaReciclada = 3
     sghHistoriaTemporalCOnsultaExterna = 4
     sghHistoriaTemporalEmergencia = 5
     sghHistoriaTemporalAlojamiento = 6
     sghHistoriaTemporalServiciosIntermedios = 7
     sghSinHistoria = 9
End Enum

Enum sghTipoServicio
     sghConsultaExterna = 1
     sghEmergenciaConsultorios = 2
     sghHospitalizacion = 3
     sghEmergenciaObservacion = 4
End Enum

Enum sghBotonDetallePresionado
    sghAceptar = 10
    sghCancelar = 20
End Enum

Enum sghTipoFiltroPacientes
    sghFiltrarTodos = 10
    sghFiltrarConHistoriasTemporales = 20
    sghFiltrarConHistoriasDefinitivas = 30
End Enum

Enum sghTipoFiltroAdmision
    sghFiltrarConsultaExterna = 10
    sghFiltrarHospitalizacion = 20
    sghFiltrarConsultorioEmergencia = 30
    sghFiltrarObservacionEmergencia = 40
    sghFiltrarEmergencia = 50
End Enum

Enum sghEtapaPrestamoHistoriaClinica
    sghSolicitud = 1
    sghEnvio = 2
    sghDevolucion = 3
End Enum

Enum sghTipoBusquedaPrestamoHistoria
    sghTodasHistorias = 1
    sghHistoriaSolicitadas = 2
    sghHistoriaEnPrestamo = 3
    sghHistoriaDevueltas = 4
End Enum

Enum sghTipoVistaFormAtenciones
    sghVistaAdmision = 1
    sghVistaAtencion = 2
End Enum
Enum sghTiposDiagnostico
    sghAtencionConsultaExterna = 1
    sghHospitalizacionIngreso = 2
    sghHospitalizacionEgreso = 3
    sghHospitalizacionMortalidad = 4
    sghHospitalizacionNacimiento = 5
    sghHospitalizacionComplicaciones = 6
    sghInterconsultas = 7
End Enum
Enum sghTipoAccionEmergenciaYHospitalizacion
    sghAdmisionNormal = 1
    sghEnviarAObservacion = 2
    sghTrasladarAHospitalizacion = 3
    sghDarDeAlta = 4
    sghIngresarUnAlojamientoConjunto = 5
    sghTransferencias = 6
End Enum
Enum sghTiposReporteHospitalizacion
    sghReporteEgresosHospitalario = 1
    sghReporteIngresosHospitalario = 2
End Enum

Enum sghTipoDetalleComprobante
    sghDetalleComprobanteServicios = 1
    sghDetalleComprobanteInsumos = 2
End Enum

Enum sghEstadoFacturacion
    sghAtendido = 1
    sghPendientePago = 3
    sghPagadoYatendido = 4
    sghDevolver = 5
    sghDevuelto = 6
    sghAnulado = 9
    sghAutorizAutomática = 10
    sghDespachado = 11
    sghRegistrado = 12
    sghReembolsoParcial = 15
    sghConPreVenta = 16
End Enum

Enum sghTipoEstadoAtencion
    sghEstadoAtencionSolicitado = 1
    sghEstadoAtencionAtendido = 2
End Enum

Enum sghEstadoCuenta
    sghAbierto = 1
    sghPagado = 4
    sghCerrado = 5
    sghAnulado = 9
    sghConAltaMedica = 10
    sghPendientePagoSeguros = 11
    sghNoLlegaAlServicioHospitalizado = 12
    sghReembolsoParcial = 15
End Enum

Enum sghTipoFacturacionServicio
    sghFacturacionServicioPorEstancia = 1
    sghFacturacionServicioPorProcedimiento = 2
    sghFacturacionServicioTotal = 3
End Enum

Enum sghTipoFinanciamiento
    sghBase = 0
    sghPacienteNormal = 1
    sghSis = 2
    sghSOAT = 3
    sghConvenios = 4
    sghCreditoHospitalario = 5
    sghDefensaNacional = 6
    sghServicioSocial = 9
    sghCreditoPersonal = 10
End Enum

Enum sghFuenteFinanciamiento
    sghFFPaciente = 1
    sghFFSIS = 3
    sghFFSoat = 2
    sghFFParticularHospitalizado = 5
    sghFFFospoli = 6
    'sghFFSeguroPacifico = 2
    'sghFFSeguroRimac = 3
    'sghFFSeguroWieseAetna = 4
    'sghFFSeguroGenerali = 5
    'sghFFESSALUD = 6
    'sghFFFospoli = 7
    'sghFFSeguroLaPositiva = 10
    'sghFFPacienteParticular = 11
End Enum


Enum sghTipoEmpleado
    sghCajero = 1
    sghCuentaCorriente = 2
    sghSis = 3
    sghConvenio = 4
    sghAsistenta = 5
    sghSOAT = 6
    sghOtros = 7
End Enum

Enum sghTipoProducto
    sghbien = 1
    sghServicio = 2
    sghAmbos = 3
End Enum

Enum sghOpcionesPago
    sghNuevoPagoConHistoria = 1
    sghNuevoPagoSinHistoria = 2
    sghPagarOrdenExistente = 3
    sghPagarCuentaExistente = 4
    sghDevolucion = 5
    sghAnulacion = 6
    sghReimprimirComprobante = 7
    sghPagarOrdenExistenteF = 8
    sghPagarOrdenExistenteFS = 9
    sghPagarCuentaTotalFS = 10
    
    sghDevolucionINO = 11
    sghAnularDevolucionINO = 12
End Enum

Enum sghAreasLaboraEmpleado
    sghAlmacenFarmacia = 1
    sghImageneología = 2
    sghLaboratorio = 3
    sghSeguros = 4
    sghEspecialidadesCE = 5
    sghEspecialidadesHosp = 6
    sghEspecialidadesEmergObs = 7
    sghEspecialidadesEmergCons = 8
    sghAreaTramitaSeguros = 9
End Enum

Enum sghTipoOrden
    sghPorCodigo = 1
    sghPorDescripcion = 2
    sghPorIdProductoMasFecha = 3
    sghPorFechaYhora = 4
    sghPorIdProductoMasIdServiciopaciente = 5
    sghPorIdFuenteFinanciamientoIdTipoServicio = 6
    sghPorServicioNombre = 7
    sghPorDepartamentoEspecialidadServicioNombre = 8
End Enum

Enum sghEstadoTabla
    sghAnulado = 0
    sghRegistrado = 1
    sghCerrado = 2
End Enum

Enum sghTipoServicioOfrecidos
    sghSoloInsumos = 0
    sghSoloCPT = 1
    sghInsumosYcpt = 2
End Enum

Enum sghImpresion
    sghPantalla = 0
    sghImpresoraBoletaContinua = 1
    sghImpresoraBoletaPorBoleta = 2
    sghImpresora = 3
    sghExcel = 4
    sghImpresionFactura = 5
End Enum

Enum sghDatoDelEmpleado
    sghIniciales = 0
    sghUsuario = 1
    sghApellidosYnombres = 2
End Enum


Enum sghTipoServicioHospitalizacion
    sghSoloPacHospitalizados = 1
    sghSoloPacAlojados = 2
    sghTodos = 3
End Enum

Enum sghTipoEdades
    sghAño = 1
    sghMeses = 2
    sghDias = 3
    sghHoras = 4
End Enum

Enum sghTipoEstados
    sghFiltraSoloAnulados = 0
    sghFiltraSoloActivos = 1
    sghFiltraAnuladosYactivos = 2
End Enum

Enum sghOrdenDeServiciosHospital
    sghPorDescTipoServicio = 1
    sghPorDescServicio = 2
End Enum

Enum sghTipoFinanciamientoGeneraPago
    sghTodosLosQuePaganEnCaja = 1
    sghTodosLosQueTienenAlgunSeguro = 5
    sghSoloSeguroSIS = 2
    sghSoloSeguroSOAT = 3
    sghSoloSeguroConvenios = 4
End Enum

Enum sghComoSeTrabajaEnEstadoCuentaLosSeguros
    sghTrabajaNinguno = 0
    sghTrabajaParticular = 1
    sghTrabajaSeguroSIS = 2
    sghTrabajaSeguroSOAT = 3
    sghTrabajaSeguroConvenios = 4
    sghTrabajaServicioSocial = 9
End Enum



Enum sghPuntosCargaBasicos
    sghPtoCargaAdmisionEmergencia = 10
    sghPtoCargaAdmisionHospitalizacion = 9
    sghPtoCargaAdmisionCE = 6
    sghPtoCargaServicioHospitalizacion = 1
    sghPtoCargaCaja = 99
    sghPtoCargaRayosX = 21
    sghPtoCargaTomografia = 22
    sghPtoCargaEcogObstetrica = 23
    sghPtoCargaEcogGeneral = 20
    sghPtoCargaPatologiaClinica = 2
    sghPtoCargaAnatomiaPatologica1 = 3
    sghPtoCargaAnatomiaPatologica2 = 32
    sghPtoCargaBancoSangre1 = 11
    sghPtoCargaBancoSangre2 = 38
    sghPtoCargaFarmacia = 5
End Enum


Enum sghTipoPaquetes
    sghTipoPaqueteSoloServicio = 1
    sghTipoPaqueteSolofarmacia = 2
    sghTipoPaqueteServicioYfarmacia = 3
End Enum

Enum sghTipoPrecioFarmacia
    sghPrecioCompra = 1
    sghPrecioDistribucion = 2
    sghPrecioVentaContado = 3
    sghPrecioDonacion = 4
End Enum

Enum sghTipoConceptoImagen
    sghImgTCingreso = 1
    sghImgTCsalidaDeterioro = 2
    sghImgTCsalida = 3
End Enum

Enum sghTipoDx
    sghTipoDxNINGUNO = 0
    sghTipoDxDefinitivo = 1
    sghTipoDxPresuntivo = 2
End Enum

Enum sghTipoSalidaItemFarmacia
    sghSoloVenta = 1
    sghSoloEstrategico = 2
    sghVentaEstrategico = 3
    sghDonaciones = 4
End Enum

Enum sghBaseDatosExterna
    sghJamo = 273
    sghSis = 325
End Enum

'debb-24/03/2011
Enum sghFiltraCpt
    sghMuestraTodosCpt = 0
    sghCptSoloLaboratorio = 1
    sghCptSoloRayosX = 2
    sghCptSoloTomografia = 3
    sghCptSoloEcografiaObstetrica = 4
    sghCptSoloEcografiaGeneral = 5
    sghCptSoloLabPatologiaC = 102
    sghCptSoloLabAnatomiaP = 103
    sghCptSoloLabBancoS = 104
End Enum


Enum sghPerinatalModulos
    sighHasta28Dias = 1
    sighDesde29diasHasta1anio = 2
    sighDesde1Hasta4anios = 3
    sighDesde5Hasta9anios = 4
    sighDesde10Hasta11anios = 5
    sighDesde12Hasta17anios = 6
    sighDesde18anios = 7
End Enum

Enum sghPerinatalListas
    sighInmunizaciones = 1
    sighCptFrecuentes = 2
    sighMorbilidadDesarrollo = 3
    sighMorbilidadFrecuente = 4
End Enum

Enum sghDxDefinitivos
    sighDxCeDefinitivo = 102
    sighDxHospEmergPrincipal = 301
    sighDxHospEmergCausaFinal = 303
    sighDxHospEmergDefinitivo = 402
End Enum

Enum sghRecetaEstados
    sighRecetaAnulada = 0
    sighRecetaRegistrada = 1
    sighRecetaDespachada = 2
    sighRecetaConBoleta = 3
End Enum

Enum sghEstadosComprobante
    sighEstadosComprobantePagado = 4
    sighEstadosComprobanteDevuelto = 6
    sighEstadosComprobanteAnulado = 9
End Enum
Enum sghTipoConceptoFarmacia
    sghTipoConceptoSIS = 13
    sghTipoConceptoSOAT = 14
    sghTipoConceptoConvenios = 23
End Enum
Enum sghOpcionGalenHos
     sghRegistroCitaCE = 102
     sghRegistroAtencionCE = 103
     sghEstadoDeCuenta = 613
     sghFormatoFUA = 1345
     sghAdmisionEmergencia = 202
     sghAdmisionHospitalizacion = 302
     sghVentasFarmacia = 1307
     sghConsumoEnServicio = 601
     sghLaboratorioAP = 1321
     sghLaboratorioPC = 1312
     sghLaboratorioBS = 1322
     sghImagenEcogO = 1320
     sghImagenEcogG = 1317
     sghImagenRayosX = 1318
     sghImagenTomografia = 1319
     sghPacienteExternoConSeguro = 1339
     sghGestionGaja = 702
     sghPacientes = 101
     sghReembolsos = 1331
End Enum
Enum sghTipoCondicion
     sghTipoCondicionNuevo = 1
     sghTipoCondicionReingresante = 2
     sghTipoCondicionContinuador = 3
End Enum
Enum sghUltimaBusqueda
     sghEnBoleta = 1
     sghEnNroCuenta = 2
End Enum

'JVG 02-04-2012 - Adicion de Enumeracion para el Modulo HIS GalenHos
'Tipo de Actividades en HIS GalenHos
Enum sghHISTipoActividad
    Atencion = 1
    ActividadPreventivaPromocional = 2
    ActividadMasiva = 3
    ActividadConAnimales = 4
End Enum

Enum sghHISTipoEdades
    dias = 1
    meses = 2
    Años = 3
End Enum

Enum sghHISEstados
    Nuevo = 1
    Reingreso = 2
    Continuador = 3
End Enum

Enum sghCitaWebEstados
    CupoANULADO = 0
    CupoLlenadoEnCitaGalenHos = 1
    CupoDisponibleEnCitaWeb = 2
    CupoConfirmadoEnCitaWeb = 3
    CupoConfirmadoYconCitaEnGalenhos = 4
End Enum

Enum sghSIScodigo
    sghAfiliacionLPISpago = 1
    sghInscripcionLPIS = 2
    sghAfiliacionLPISgratis = 3
    sghAfiliacionLPISgratisAntiguo = 4
    sghAfiliacionAUXgratis = 7
    sghInscripcionAUX = 8
End Enum

Enum sghColores
    sghRojo = 1
    sghAzul = 2
    sghNegro = 3
    sghVerde = 4
    sghBlanco = 5
End Enum


Enum sghRecetasEstadosDetalle
    sghAnuladoPorMedico = 0
    sghRecetado = 1
    sghDespachado = 2
    sghEnDosisAlPaciente = 3
    sghConcluido = 4
End Enum

Enum sghVersionBD
    sighSql2000 = 0
    sighSql2008 = 1
    sighPosgreSql = 2
End Enum

Enum sghDefaultVentana
    sighApellidoPaterno = 0
    sighDNI = 1
    sighHistoria = 2
End Enum

Enum sghFuaTipo
    sghFuaTipoManual = 1
    sghFuaTipoAutomatico = 2
End Enum


Enum sghTipoProceso
    sghProcesaYgraba = 1
    sghSoloParaReporte = 2
End Enum

Enum sghCategoriaEstablecimiento
    sghOtros = 0
    sghHospital = 1
    sghCS = 2
    sghPS = 3
End Enum

'mgaray
Enum sighTriajeOrigen
    Triaje = 1
    ConsultaExterna = 2
    Emergencia = 3
    Hospitalizacion = 4
End Enum

Enum sighTriajeVariable
    Peso = 1
    Talla = 2
    PerimCefalico = 3
    PresArtSistolica = 4
    PresArtDiastolica = 5
    Temperatura = 6
    FrecCardiaca = 7
    FrecRespiratoria = 8
    Pulso = 9
End Enum

Enum sighTriajeEstadoPaciente
    NoRequerido = 0
    Despierto = 1
    Dormido = 2
End Enum

'atención integral
Enum sighGrupoEdad
    Nino = 1
    Adolescente = 2
    Joven = 3
    Adulto = 4
    AdultoMayor = 5
End Enum

Enum sighTipoDatoRespuesta
    Texto = 1
    Numerico = 2
    fecha = 3
End Enum

Enum sighItemPlanIntegral
    Inmunizacion = 1
    Crecimiento = 2
    Desarrollo = 3
    SuplementoNutricional = 4
    Tamizaje = 5
End Enum

Enum sghOrdenServicio
    sghPorEspecialidad = 0
    sghPorNombreServicio = 1
End Enum

Enum sghCitasWebEstados
    sghCWebAnulado = 0
    sghCWebCitaEnGalenhos = 1
    sghCWebCupoPCitaWeb = 2
    sghCWebCitaWebConfirmado = 3
    sghCWebCitaWebConfirmadoYcitado = 4
End Enum

 
Enum sghCitaDesde
    sghReprogramarFecha = 1
    sghReprogramarMedico = 2
    sghMantenimientoCita = 3
End Enum

'mgaray20141009
Enum sghIntegracionTipoSistema
    sghRisPacs = 1
End Enum

Enum sghIntegracionProveedorSistema
    sghCarestream = 1
End Enum

Enum sghSexo
    Masculino = 1
    Femenino = 2
End Enum

'DEBB2014a
Enum sghInventarioTipo
    sghManual = 1
    sghAutomatico = 2
End Enum

'mgaray201411e
Enum sghDesviacion
    sghDesviacionMenos2 = -2
    sghDesviacionMenos1 = -1
    sghNormal = 0
    sghDesviacion1 = 1
    sghDesviacion2 = 2
End Enum

Enum sghDesviacionColor
    sghDesviacionMenos3 = &H0&
    sghDesviacionMenos2 = &HFF&
    sghDesviacionMenos1 = &H80FF&
    sghNormal = &H8000&
    sghDesviacion1 = &H80FF&
    sghDesviacion2 = &HFF&
    sghDesviacion3 = &H0&
    sghValorTriaje = &H400000
End Enum
'mgaray201503
Enum sghIdesTipoFinanciador
    PersonaNatural = 1
    PersonaJuridica = 2
    Extranjero = 3
End Enum

Enum sghRENAESFuente
    SIS = 1
    SUNASA = 2
    GALENHOS = 3
    RENAESNORMA = 4
End Enum

Enum sghTipoEstablecimiento
    Hospital = 1
    CentroSalud = 2
    PuestoSalud = 3
End Enum

Enum sghEstadoPeticionHttp
    RespuestaOk = 200
    PagNoEncontrada = 404
    ErrorInternoServidor = 500
    ServidorNoEncontrado = 502
    TiempoEsperaAgotado = 504
    SinConexion = 12029
End Enum

'FRANK 24082015
Enum sghEstadoNotaCredito
    PorAprobar = 0
    Aprobado = 1
    anulado = 2
    Canjeado = 3
End Enum

Enum sghActividadesTipo
     TipoCPT = 1
     TipoLAB = 2
     TipoDX = 3
End Enum


Enum sghClinica
     sghVerificaLicenciaPorDia = 1
     sghVerificaSiEsClinicaYtieneLicencia = 2
     sghAsignaLicenciaAnualPorquePago = 3
     sghVerificaSiEsUnaClinica = 4
     sghVerificaSiTieneLicencia = 5
End Enum

Enum sghEstadosHistoria
     sghActiva = 1
     sghDepurada = 2
     sghDepuradaXeliminar = 3
End Enum


Enum sghTurnos
     sghTurnoAmbos = 0
     sghTurnoManana = 1
     sghTurnoTarde = 2
End Enum

Enum sghInterLabEstados
     sghIntAnulado = 0
     sghIntRegistradoEnGalenhos = 1
     sghIntLeidoPorSistemaLab = 2
     sghIntTieneResultadoLab = 3
     sghIntTieneResultadoLeidoEnGalenhos = 4
End Enum

Enum sghSiCitasEstados
     sghSiCitaAnulada = 0
     sghSiCitaActiva = 1
     sghSiCitaConTomaMuestra = 2
End Enum
Enum sghSiCuposProgamados
     sghSiProgramacionEnDataCombo = 1
     sghSiProgramacionEnListBox = 2
     sghSiDevuelveTotalCuposXdia = 3
End Enum

Enum sghAuditoria
     sghAudLimpiar = 1
     sghAudGrabaRegEdit = 2
     sghAudLeeRegedit = 3
End Enum

Enum sghCajaTipoComprobante
     sghCajaFactura = 2
     sghCajaBoleta = 3
End Enum

Enum sghCajaTipoOrden
     sghCajaServicioCpt = 1
     sghCajaFarmacia2 = 2
     sghCajaFarmacia3 = 3
End Enum

Enum sghTipoImpresion
     sghTIimpresora = 1
     sghTIexcel = 2
     sghTIpdf = 3
End Enum
