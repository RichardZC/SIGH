VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoAtencionDatosAdicionales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtencionDatosAdicionales
'        Programado por: Barrantes D
'        Fecha: Enero 2010
'
'------------------------------------------------------------------------------------

Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ms_idAtencion As Long
Dim ms_DireccionDomicilio As String
Dim ms_NombreAcompaniante As String
Dim ms_Observacion As String
Dim mda_ProximaCita As Date
Dim ml_NumeroDeHijos As Long
Dim ml_IdSiaSis As Long
Dim ms_FuaCodigoPrestacion As String

Dim ms_FuaCodigoPrestacion1 As String 'HRA 10/12/2020 Cambio Inicio

Dim ms_SisCodigo As String
Dim ml_IdTipoReferenciaDestino As Long
Dim ml_IdTipoReferenciaOrigen As Long
Dim ml_IdEstablecimientoDestino As Long
Dim ml_IdEstablecimientoOrigen As Long
Dim ml_IdEstablecimientoNoMinsaDestino As Long
Dim ml_IdEstablecimientoNoMinsaOrigen As Long
Dim mb_HuboInfeccionIntraHospitalaria As Boolean
Dim mb_TieneNecropsia As Boolean
Dim ml_IdMedicoRespNacimiento As Long
Dim mb_RecienNacido As Boolean
Dim ml_NroReferenciaOrigen As String
Dim ml_NroReferenciaDestino As String
Dim mb_SeImprimioFicha As Boolean
Dim ml_referenciaOservicio As String
Dim ml_referenciaOidDiagnostico As Long
Dim ml_referenciaDservicio As String
Dim ml_referenciaDfextension As Date
Dim ml_referenciaDftramite As Date
Dim ml_referenciaOdniMedico As String
Dim ml_acompanianteDNI As String
Dim ml_emergenciaCorrelativo As String
Dim ml_sisAfiliacion As String
Dim ml_idAtencionEmeg_CE As Long
Dim ml_ReferenciaMedicoOColeg As String   'FRANKLIN 2017
Dim ml_ReferenciaMedicoOIdcolegio As String   'FRANKLIN 2017
Dim ml_IdServicioDestino As Long
Property Let idServicioDestino(lValue As Long)
   ml_IdServicioDestino = lValue
End Property
Property Get idServicioDestino() As Long
   idServicioDestino = ml_IdServicioDestino
End Property

'FRANKLIN 2017
Property Let ReferenciaMedicoOIdcolegio(sValue As String)
   ml_ReferenciaMedicoOIdcolegio = sValue
End Property
Property Get ReferenciaMedicoOIdcolegio() As String
   ReferenciaMedicoOIdcolegio = ml_ReferenciaMedicoOIdcolegio
End Property
'FRANKLIN 2017
Property Let ReferenciaMedicoOColeg(sValue As String)
   ml_ReferenciaMedicoOColeg = sValue
End Property
Property Get ReferenciaMedicoOColeg() As String
   ReferenciaMedicoOColeg = ml_ReferenciaMedicoOColeg
End Property

Property Let idAtencionEmeg_CE(lValue As Long)
   ml_idAtencionEmeg_CE = lValue
End Property
Property Get idAtencionEmeg_CE() As Long
   idAtencionEmeg_CE = ml_idAtencionEmeg_CE
End Property

Property Let sisAfiliacion(sValue As String)
   ml_sisAfiliacion = sValue
End Property
Property Get sisAfiliacion() As String
   sisAfiliacion = ml_sisAfiliacion
End Property
Property Let referenciaOdniMedico(sValue As String)
   ml_referenciaOdniMedico = sValue
End Property
Property Get referenciaOdniMedico() As String
   referenciaOdniMedico = ml_referenciaOdniMedico
End Property
Property Let acompanianteDNI(sValue As String)
   ml_acompanianteDNI = sValue
End Property
Property Get acompanianteDNI() As String
   acompanianteDNI = ml_acompanianteDNI
End Property
Property Let emergenciaCorrelativo(sValue As String)
   ml_emergenciaCorrelativo = sValue
End Property
Property Get emergenciaCorrelativo() As String
   emergenciaCorrelativo = ml_emergenciaCorrelativo
End Property




Property Let referenciaOidDiagnostico(lValue As Long)
   ml_referenciaOidDiagnostico = lValue
End Property
Property Get referenciaOidDiagnostico() As Long
   referenciaOidDiagnostico = ml_referenciaOidDiagnostico
End Property

Property Let referenciaDftramite(sValue As Date)
   ml_referenciaDftramite = sValue
End Property
Property Get referenciaDftramite() As Date
   referenciaDftramite = ml_referenciaDftramite
End Property
Property Let referenciaDfextension(sValue As Date)
   ml_referenciaDfextension = sValue
End Property
Property Get referenciaDfextension() As Date
   referenciaDfextension = ml_referenciaDfextension
End Property
Property Let referenciaDservicio(sValue As String)
   ml_referenciaDservicio = sValue
End Property
Property Get referenciaDservicio() As String
   referenciaDservicio = ml_referenciaDservicio
End Property

Property Let referenciaOservicio(sValue As String)
   ml_referenciaOservicio = sValue
End Property
Property Get referenciaOservicio() As String
   referenciaOservicio = ml_referenciaOservicio
End Property

Property Let IdTipoReferenciaOrigen(lValue As Long)
   ml_IdTipoReferenciaOrigen = lValue
End Property
Property Get IdTipoReferenciaOrigen() As Long
   IdTipoReferenciaOrigen = ml_IdTipoReferenciaOrigen
End Property
Property Let IdTipoReferenciaDestino(lValue As Long)
   ml_IdTipoReferenciaDestino = lValue
End Property
Property Get IdTipoReferenciaDestino() As Long
   IdTipoReferenciaDestino = ml_IdTipoReferenciaDestino
End Property
Property Let IdEstablecimientoOrigen(lValue As Long)
   ml_IdEstablecimientoOrigen = lValue
End Property
Property Get IdEstablecimientoOrigen() As Long
   IdEstablecimientoOrigen = ml_IdEstablecimientoOrigen
End Property
Property Let IdEstablecimientoDestino(lValue As Long)
   ml_IdEstablecimientoDestino = lValue
End Property
Property Get IdEstablecimientoDestino() As Long
   IdEstablecimientoDestino = ml_IdEstablecimientoDestino
End Property
Property Let IdEstablecimientoNoMinsaOrigen(lValue As Long)
   ml_IdEstablecimientoNoMinsaOrigen = lValue
End Property
Property Get IdEstablecimientoNoMinsaOrigen() As Long
   IdEstablecimientoNoMinsaOrigen = ml_IdEstablecimientoNoMinsaOrigen
End Property

Property Let IdEstablecimientoNoMinsaDestino(lValue As Long)
   ml_IdEstablecimientoNoMinsaDestino = lValue
End Property
Property Get IdEstablecimientoNoMinsaDestino() As Long
   IdEstablecimientoNoMinsaDestino = ml_IdEstablecimientoNoMinsaDestino
End Property
Property Let HuboInfeccionIntraHospitalaria(bValue As Boolean)
   mb_HuboInfeccionIntraHospitalaria = bValue
End Property
Property Get HuboInfeccionIntraHospitalaria() As Boolean
   HuboInfeccionIntraHospitalaria = mb_HuboInfeccionIntraHospitalaria
End Property
Property Let TieneNecropsia(bValue As Boolean)
   mb_TieneNecropsia = bValue
End Property
Property Get TieneNecropsia() As Boolean
   TieneNecropsia = mb_TieneNecropsia
End Property
Property Let IdMedicoRespNacimiento(lValue As Long)
   ml_IdMedicoRespNacimiento = lValue
End Property
Property Get IdMedicoRespNacimiento() As Long
   IdMedicoRespNacimiento = ml_IdMedicoRespNacimiento
End Property
Property Let RecienNacido(bValue As Boolean)
   mb_RecienNacido = bValue
End Property
Property Get RecienNacido() As Boolean
   RecienNacido = mb_RecienNacido
End Property
Property Let NroReferenciaDestino(lValue As String)
   ml_NroReferenciaDestino = lValue
End Property
Property Get NroReferenciaDestino() As String
   NroReferenciaDestino = ml_NroReferenciaDestino
End Property
Property Let NroReferenciaOrigen(lValue As String)
   ml_NroReferenciaOrigen = lValue
End Property
Property Get NroReferenciaOrigen() As String
   NroReferenciaOrigen = ml_NroReferenciaOrigen
End Property



Property Let SisCodigo(sValue As String)
   ms_SisCodigo = sValue
End Property
Property Get SisCodigo() As String
   SisCodigo = ms_SisCodigo
End Property

Property Let idSiasis(lValue As Long)
   ml_IdSiaSis = lValue
End Property
Property Get idSiasis() As Long
   idSiasis = ml_IdSiaSis
End Property

Property Let NumeroDeHijos(lValue As Long)
   ml_NumeroDeHijos = lValue
End Property
Property Get NumeroDeHijos() As Long
   NumeroDeHijos = ml_NumeroDeHijos
End Property
Property Let ProximaCita(daValue As Date)
   mda_ProximaCita = daValue
End Property
Property Get ProximaCita() As Date
   ProximaCita = mda_ProximaCita
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idAtencion(sValue As Long)
   ms_idAtencion = sValue
End Property
Property Get idAtencion() As Long
   idAtencion = ms_idAtencion
End Property
Property Let DireccionDomicilio(sValue As String)
   ms_DireccionDomicilio = sValue
End Property
Property Get DireccionDomicilio() As String
   DireccionDomicilio = ms_DireccionDomicilio
End Property
Property Let NombreAcompaniante(sValue As String)
   ms_NombreAcompaniante = sValue
End Property
Property Get NombreAcompaniante() As String
   NombreAcompaniante = ms_NombreAcompaniante
End Property
Property Let Observacion(sValue As String)
   ms_Observacion = sValue
End Property
Property Get Observacion() As String
   Observacion = ms_Observacion
End Property

Property Let FuaCodigoPrestacion(sValue As String)
   ms_FuaCodigoPrestacion = sValue
End Property
Property Get FuaCodigoPrestacion() As String
   FuaCodigoPrestacion = ms_FuaCodigoPrestacion
End Property
Property Let SeImprimioFicha(sValue As Boolean)
   mb_SeImprimioFicha = sValue
End Property
Property Get SeImprimioFicha() As Boolean
   SeImprimioFicha = mb_SeImprimioFicha
End Property

'HRA 10/12/2020 Cambio Inicio
Property Let FuaCodigoPrestacion1(sValue As String)
   ms_FuaCodigoPrestacion1 = sValue
End Property
Property Get FuaCodigoPrestacion1() As String
   FuaCodigoPrestacion1 = ms_FuaCodigoPrestacion1
End Property
'HRA 10/12/2020 Cambio Fin
 

