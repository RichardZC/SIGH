VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoSunasaPacientesHistoricos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla SunasaPacientesHistoricos
'        Programado por: Barrantes D
'        Fecha: Julio 2011
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idSunasaPacienteHistorico As Long
Dim ml_IdPaciente As Long
Dim ms_CodigoIAFA As String
Dim ml_idPaisTitular As Long
Dim ml_idTipoDocumentoTitular As Long
Dim ms_NroDocumentoTitular As String
Dim ms_ApellidoCasada As String
Dim mb_ValidacionRegIdentidad As Boolean
Dim ms_NroCarnetIdentidad As String
Dim ml_EstadoDelSeguro As Long
Dim ml_IdAfiliacion As Long
Dim ms_ProductoYplan As String
Dim mda_FechaInicioAfiliacion As Date
Dim mda_FechaFinalAfiliacion As Date
Dim ml_idRegimen As Long
Dim ms_CodigoEstablecimientoIAFA As String
Dim ms_CodigoEstablecimientoRENAES As String
Dim ml_idParentesco As Long
Dim ms_RUCempleador As String
Dim ml_AnteriorIdTipoDocumentoAsegurado As Long
Dim ms_AnteriorNroDocumentoAsegurado As String
Dim ms_DNIusarioOperacion As String
Dim ml_idOperacion As Long
Dim mda_FechaEnvio As Date
Dim ms_SisSepelioParienteEncargado As String
Dim ms_SisSepelioDni As String
Dim mda_SisSepelioFnacimiento As Date
Dim ml_SisSepelioSexo As Long
Dim ms_SisNroAfiliacion As String
Dim mb_YaNoTieneSeguro As Boolean
Dim mb_NuevoSeguro As Boolean   'No se graba en la tabla

Property Let NuevoSeguro(bValue As Boolean)
   mb_NuevoSeguro = bValue
End Property
Property Get NuevoSeguro() As Boolean
   NuevoSeguro = mb_NuevoSeguro
End Property

Property Let YaNoTieneSeguro(bValue As Boolean)
   mb_YaNoTieneSeguro = bValue
End Property
Property Get YaNoTieneSeguro() As Boolean
   YaNoTieneSeguro = mb_YaNoTieneSeguro
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idSunasaPacienteHistorico(iValue As Long)
   ml_idSunasaPacienteHistorico = iValue
End Property
Property Get idSunasaPacienteHistorico() As Long
   idSunasaPacienteHistorico = ml_idSunasaPacienteHistorico
End Property
Property Let IdPaciente(iValue As Long)
   ml_IdPaciente = iValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let CodigoIAFA(sValue As String)
   ms_CodigoIAFA = sValue
End Property
Property Get CodigoIAFA() As String
   CodigoIAFA = ms_CodigoIAFA
End Property
Property Let idPaisTitular(iValue As Long)
   ml_idPaisTitular = iValue
End Property
Property Get idPaisTitular() As Long
   idPaisTitular = ml_idPaisTitular
End Property
Property Let idTipoDocumentoTitular(iValue As Long)
   ml_idTipoDocumentoTitular = iValue
End Property
Property Get idTipoDocumentoTitular() As Long
   idTipoDocumentoTitular = ml_idTipoDocumentoTitular
End Property
Property Let NroDocumentoTitular(sValue As String)
   ms_NroDocumentoTitular = sValue
End Property
Property Get NroDocumentoTitular() As String
   NroDocumentoTitular = ms_NroDocumentoTitular
End Property
Property Let ApellidoCasada(sValue As String)
   ms_ApellidoCasada = sValue
End Property
Property Get ApellidoCasada() As String
   ApellidoCasada = ms_ApellidoCasada
End Property
Property Let ValidacionRegIdentidad(bValue As Boolean)
   mb_ValidacionRegIdentidad = bValue
End Property
Property Get ValidacionRegIdentidad() As Boolean
   ValidacionRegIdentidad = mb_ValidacionRegIdentidad
End Property
Property Let NroCarnetIdentidad(sValue As String)
   ms_NroCarnetIdentidad = sValue
End Property
Property Get NroCarnetIdentidad() As String
   NroCarnetIdentidad = ms_NroCarnetIdentidad
End Property
Property Let EstadoDelSeguro(iValue As Long)
   ml_EstadoDelSeguro = iValue
End Property
Property Get EstadoDelSeguro() As Long
   EstadoDelSeguro = ml_EstadoDelSeguro
End Property
Property Let IdAfiliacion(iValue As Long)
   ml_IdAfiliacion = iValue
End Property
Property Get IdAfiliacion() As Long
   IdAfiliacion = ml_IdAfiliacion
End Property
Property Let ProductoYplan(sValue As String)
   ms_ProductoYplan = sValue
End Property
Property Get ProductoYplan() As String
   ProductoYplan = ms_ProductoYplan
End Property
Property Let FechaInicioAfiliacion(daValue As Date)
   mda_FechaInicioAfiliacion = daValue
End Property
Property Get FechaInicioAfiliacion() As Date
   FechaInicioAfiliacion = mda_FechaInicioAfiliacion
End Property
Property Let FechaFinalAfiliacion(daValue As Date)
   mda_FechaFinalAfiliacion = daValue
End Property
Property Get FechaFinalAfiliacion() As Date
   FechaFinalAfiliacion = mda_FechaFinalAfiliacion
End Property
Property Let idRegimen(iValue As Long)
   ml_idRegimen = iValue
End Property
Property Get idRegimen() As Long
   idRegimen = ml_idRegimen
End Property
Property Let CodigoEstablecimientoIAFA(sValue As String)
   ms_CodigoEstablecimientoIAFA = sValue
End Property
Property Get CodigoEstablecimientoIAFA() As String
   CodigoEstablecimientoIAFA = ms_CodigoEstablecimientoIAFA
End Property
Property Let CodigoEstablecimientoRENAES(sValue As String)
   ms_CodigoEstablecimientoRENAES = sValue
End Property
Property Get CodigoEstablecimientoRENAES() As String
   CodigoEstablecimientoRENAES = ms_CodigoEstablecimientoRENAES
End Property
Property Let idParentesco(iValue As Long)
   ml_idParentesco = iValue
End Property
Property Get idParentesco() As Long
   idParentesco = ml_idParentesco
End Property
Property Let RUCempleador(sValue As String)
   ms_RUCempleador = sValue
End Property
Property Get RUCempleador() As String
   RUCempleador = ms_RUCempleador
End Property
Property Let AnteriorIdTipoDocumentoAsegurado(iValue As Long)
   ml_AnteriorIdTipoDocumentoAsegurado = iValue
End Property
Property Get AnteriorIdTipoDocumentoAsegurado() As Long
   AnteriorIdTipoDocumentoAsegurado = ml_AnteriorIdTipoDocumentoAsegurado
End Property
Property Let AnteriorNroDocumentoAsegurado(sValue As String)
   ms_AnteriorNroDocumentoAsegurado = sValue
End Property
Property Get AnteriorNroDocumentoAsegurado() As String
   AnteriorNroDocumentoAsegurado = ms_AnteriorNroDocumentoAsegurado
End Property
Property Let DNIusarioOperacion(sValue As String)
   ms_DNIusarioOperacion = sValue
End Property
Property Get DNIusarioOperacion() As String
   DNIusarioOperacion = ms_DNIusarioOperacion
End Property
Property Let idOperacion(iValue As Long)
   ml_idOperacion = iValue
End Property
Property Get idOperacion() As Long
   idOperacion = ml_idOperacion
End Property
Property Let FechaEnvio(daValue As Date)
   mda_FechaEnvio = daValue
End Property
Property Get FechaEnvio() As Date
   FechaEnvio = mda_FechaEnvio
End Property
Property Let SisSepelioParienteEncargado(sValue As String)
   ms_SisSepelioParienteEncargado = sValue
End Property
Property Get SisSepelioParienteEncargado() As String
   SisSepelioParienteEncargado = ms_SisSepelioParienteEncargado
End Property
Property Let SisSepelioDni(sValue As String)
   ms_SisSepelioDni = sValue
End Property
Property Get SisSepelioDni() As String
   SisSepelioDni = ms_SisSepelioDni
End Property
Property Let SisSepelioFnacimiento(daValue As Date)
   mda_SisSepelioFnacimiento = daValue
End Property
Property Get SisSepelioFnacimiento() As Date
   SisSepelioFnacimiento = mda_SisSepelioFnacimiento
End Property
Property Let SisSepelioSexo(iValue As Long)
   ml_SisSepelioSexo = iValue
End Property
Property Get SisSepelioSexo() As Long
   SisSepelioSexo = ml_SisSepelioSexo
End Property
Property Let SisNroAfiliacion(sValue As String)
   ms_SisNroAfiliacion = sValue
End Property
Property Get SisNroAfiliacion() As String
   SisNroAfiliacion = ms_SisNroAfiliacion
End Property

