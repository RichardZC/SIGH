VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Pacientes
'        Programado por: Barrantes D.
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdPaisNacimiento As Long
Dim ms_ApellidoMaterno As String
Dim ms_DireccionDomicilio As String
Dim ms_Observacion As String
Dim ml_IdTipoNumeracion As Long
Dim ml_IdPaisProcedencia As Long
Dim ml_idPaciente As Long
Dim ms_ApellidoPaterno As String
Dim ms_PrimerNombre As String
Dim ms_SegundoNombre As String
Dim ms_TercerNombre As String
Dim mda_FechaNacimiento As Date
Dim ms_NroDocumento As String
Dim ms_Telefono As String
Dim ms_Autogenerado As String
Dim ml_idTipoSexo As Long
Dim ml_IdProcedencia As Long
Dim ml_IdGradoInstruccion As Long
Dim ml_IdEstadoCivil As Long
Dim ml_IdDocIdentidad As Long
Dim ml_IdTipoOcupacion As Long
Dim ml_IdCentroPobladoDomicilio As Long
Dim ms_NombrePadre As String
Dim ms_NombreMadre As String
Dim ml_IdPaisDomicilio As Long
Dim ml_NroHistoriaClinica As Long
Dim ml_IdCentroPobladoNacimiento As Long
Dim ml_IdCentroPobladoProcedencia As Long
Dim ml_IdDistritoProcedencia As Long
Dim ml_IdDistritoDomicilio As Long
Dim ml_IdDistritoNacimiento As Long
Dim ms_FichaFamiliar As String
Dim ms_IdEtnia As String
Dim ms_GrupoSanguineo As String
Dim ms_FactorRh As String
Dim mb_UsoWebReniec As Boolean

'<(Inicio) Añadido Por: WABG el: 22/10/2020-10:40:28 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
Dim mb_validacionReniec As Boolean
'</(Fin) Añadido Por: WABG el: 22/10/2020-10:40:28 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>

Dim ml_IdIdioma As Long
Dim ms_Email As String
Dim ms_madreDocumento As String
Dim ms_madreApellidoPaterno As String
Dim ms_madreApellidoMaterno As String
Dim ms_madrePrimerNombre As String
Dim ms_madreSegundoNombre As String
Dim ml_NroOrdenHijo As Long
Dim ml_madreTipoDocumento As Long
Dim ms_Sector As String
Dim ml_Sectorista As Long
Property Let Sectorista(lValue As Long)
   ml_Sectorista = lValue
End Property
Property Get Sectorista() As Long
   Sectorista = ml_Sectorista
End Property

Property Let Sector(sValue As String)
   ms_Sector = sValue
End Property
Property Get Sector() As String
   Sector = ms_Sector
End Property


Property Let madreTipoDocumento(lValue As Long)
   ml_madreTipoDocumento = lValue
End Property
Property Get madreTipoDocumento() As Long
   madreTipoDocumento = ml_madreTipoDocumento
End Property

Property Let NroOrdenHijo(lValue As Long)
   ml_NroOrdenHijo = lValue
End Property
Property Get NroOrdenHijo() As Long
   NroOrdenHijo = ml_NroOrdenHijo
End Property

Property Let madreSegundoNombre(sValue As String)
   ms_madreSegundoNombre = sValue
End Property
Property Get madreSegundoNombre() As String
   madreSegundoNombre = ms_madreSegundoNombre
End Property

Property Let madrePrimerNombre(sValue As String)
   ms_madrePrimerNombre = sValue
End Property
Property Get madrePrimerNombre() As String
   madrePrimerNombre = ms_madrePrimerNombre
End Property

Property Let madreApellidoMaterno(sValue As String)
   ms_madreApellidoMaterno = sValue
End Property
Property Get madreApellidoMaterno() As String
   madreApellidoMaterno = ms_madreApellidoMaterno
End Property

Property Let madreApellidoPaterno(sValue As String)
   ms_madreApellidoPaterno = sValue
End Property
Property Get madreApellidoPaterno() As String
   madreApellidoPaterno = ms_madreApellidoPaterno
End Property

Property Let madreDocumento(sValue As String)
   ms_madreDocumento = sValue
End Property
Property Get madreDocumento() As String
   madreDocumento = ms_madreDocumento
End Property



Property Let Email(sValue As String)
   ms_Email = sValue
End Property
Property Get Email() As String
   Email = ms_Email
End Property


Property Let IdIdioma(lValue As Long)
   ml_IdIdioma = lValue
End Property
Property Get IdIdioma() As Long
   IdIdioma = ml_IdIdioma
End Property

Property Let UsoWebReniec(bValue As Boolean)
   mb_UsoWebReniec = bValue
End Property
Property Get UsoWebReniec() As Boolean
   UsoWebReniec = mb_UsoWebReniec
End Property
'<(Inicio) Añadido Por: WABG el: 22/10/2020-10:42:56 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
Property Let validacionReniec(vValue As Boolean)
   mb_validacionReniec = vValue
End Property
Property Get validacionReniec() As Boolean
   validacionReniec = mb_validacionReniec
End Property
'</(Fin) Añadido Por: WABG el: 22/10/2020-10:42:56 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>

Property Let FactorRh(sValue As String)
   ms_FactorRh = sValue
End Property
Property Get FactorRh() As String
   FactorRh = ms_FactorRh
End Property

Property Let GrupoSanguineo(sValue As String)
   ms_GrupoSanguineo = sValue
End Property
Property Get GrupoSanguineo() As String
   GrupoSanguineo = ms_GrupoSanguineo
End Property

Property Let IdEtnia(sValue As String)
   ms_IdEtnia = sValue
End Property
Property Get IdEtnia() As String
   IdEtnia = ms_IdEtnia
End Property
Property Let FichaFamiliar(sValue As String)
   ms_FichaFamiliar = sValue
End Property
Property Get FichaFamiliar() As String
   FichaFamiliar = ms_FichaFamiliar
End Property

Property Let IdDistritoNacimiento(lValue As Long)
   ml_IdDistritoNacimiento = lValue
End Property
Property Get IdDistritoNacimiento() As Long
   IdDistritoNacimiento = ml_IdDistritoNacimiento
End Property
Property Let IdDistritoDomicilio(lValue As Long)
   ml_IdDistritoDomicilio = lValue
End Property
Property Get IdDistritoDomicilio() As Long
   IdDistritoDomicilio = ml_IdDistritoDomicilio
End Property
Property Let IdDistritoProcedencia(lValue As Long)
   ml_IdDistritoProcedencia = lValue
End Property
Property Get IdDistritoProcedencia() As Long
   IdDistritoProcedencia = ml_IdDistritoProcedencia
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdCentroPobladoProcedencia(lValue As Long)
   ml_IdCentroPobladoProcedencia = lValue
End Property
Property Get IdCentroPobladoProcedencia() As Long
   IdCentroPobladoProcedencia = ml_IdCentroPobladoProcedencia
End Property
Property Let IdCentroPobladoNacimiento(lValue As Long)
   ml_IdCentroPobladoNacimiento = lValue
End Property
Property Get IdCentroPobladoNacimiento() As Long
   IdCentroPobladoNacimiento = ml_IdCentroPobladoNacimiento
End Property
Property Let NroHistoriaClinica(lValue As Long)
   ml_NroHistoriaClinica = lValue
End Property
Property Get NroHistoriaClinica() As Long
   NroHistoriaClinica = ml_NroHistoriaClinica
End Property
Property Let IdPaisDomicilio(lValue As Long)
   ml_IdPaisDomicilio = lValue
End Property
Property Get IdPaisDomicilio() As Long
   IdPaisDomicilio = ml_IdPaisDomicilio
End Property
Property Let NombreMadre(sValue As String)
   ms_NombreMadre = sValue
End Property
Property Get NombreMadre() As String
   NombreMadre = ms_NombreMadre
End Property
Property Let NombrePadre(sValue As String)
   ms_NombrePadre = sValue
End Property
Property Get NombrePadre() As String
   NombrePadre = ms_NombrePadre
End Property
Property Let IdCentroPobladoDomicilio(lValue As Long)
   ml_IdCentroPobladoDomicilio = lValue
End Property
Property Get IdCentroPobladoDomicilio() As Long
   IdCentroPobladoDomicilio = ml_IdCentroPobladoDomicilio
End Property
Property Let IdTipoOcupacion(lValue As Long)
   ml_IdTipoOcupacion = lValue
End Property
Property Get IdTipoOcupacion() As Long
   IdTipoOcupacion = ml_IdTipoOcupacion
End Property
Property Let IdDocIdentidad(lValue As Long)
   ml_IdDocIdentidad = lValue
End Property
Property Get IdDocIdentidad() As Long
   IdDocIdentidad = ml_IdDocIdentidad
End Property
Property Let IdEstadoCivil(lValue As Long)
   ml_IdEstadoCivil = lValue
End Property
Property Get IdEstadoCivil() As Long
   IdEstadoCivil = ml_IdEstadoCivil
End Property
Property Let IdGradoInstruccion(lValue As Long)
   ml_IdGradoInstruccion = lValue
End Property
Property Get IdGradoInstruccion() As Long
   IdGradoInstruccion = ml_IdGradoInstruccion
End Property
Property Let IdProcedencia(lValue As Long)
   ml_IdProcedencia = lValue
End Property
Property Get IdProcedencia() As Long
   IdProcedencia = ml_IdProcedencia
End Property
Property Let idTipoSexo(lValue As Long)
   ml_idTipoSexo = lValue
End Property
Property Get idTipoSexo() As Long
   idTipoSexo = ml_idTipoSexo
End Property
Property Let Autogenerado(sValue As String)
   ms_Autogenerado = sValue
End Property
Property Get Autogenerado() As String
   Autogenerado = ms_Autogenerado
End Property
Property Let Telefono(sValue As String)
   ms_Telefono = sValue
End Property
Property Get Telefono() As String
   Telefono = ms_Telefono
End Property
Property Let NroDocumento(sValue As String)
   ms_NroDocumento = sValue
End Property
Property Get NroDocumento() As String
   NroDocumento = ms_NroDocumento
End Property
Property Let FechaNacimiento(daValue As Date)
   mda_FechaNacimiento = daValue
End Property
Property Get FechaNacimiento() As Date
   FechaNacimiento = mda_FechaNacimiento
End Property
Property Let TercerNombre(sValue As String)
   ms_TercerNombre = sValue
End Property
Property Get TercerNombre() As String
   TercerNombre = ms_TercerNombre
End Property
Property Let SegundoNombre(sValue As String)
   ms_SegundoNombre = sValue
End Property
Property Get SegundoNombre() As String
   SegundoNombre = ms_SegundoNombre
End Property
Property Let PrimerNombre(sValue As String)
   ms_PrimerNombre = sValue
End Property
Property Get PrimerNombre() As String
   PrimerNombre = ms_PrimerNombre
End Property
Property Let ApellidoPaterno(sValue As String)
   ms_ApellidoPaterno = sValue
End Property
Property Get ApellidoPaterno() As String
   ApellidoPaterno = ms_ApellidoPaterno
End Property
Property Let idPaciente(lValue As Long)
   ml_idPaciente = lValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_idPaciente
End Property
Property Let IdTipoNumeracion(lValue As Long)
   ml_IdTipoNumeracion = lValue
End Property
Property Get IdTipoNumeracion() As Long
   IdTipoNumeracion = ml_IdTipoNumeracion
End Property
Property Let Observacion(sValue As String)
   ms_Observacion = sValue
End Property
Property Get Observacion() As String
   Observacion = ms_Observacion
End Property
Property Let IdPaisProcedencia(lValue As Long)
   ml_IdPaisProcedencia = lValue
End Property
Property Get IdPaisProcedencia() As Long
   IdPaisProcedencia = ml_IdPaisProcedencia
End Property
Property Let IdPaisNacimiento(lValue As Long)
   ml_IdPaisNacimiento = lValue
End Property
Property Get IdPaisNacimiento() As Long
   IdPaisNacimiento = ml_IdPaisNacimiento
End Property
Property Let DireccionDomicilio(sValue As String)
   ms_DireccionDomicilio = sValue
End Property
Property Get DireccionDomicilio() As String
   DireccionDomicilio = ms_DireccionDomicilio
End Property
Property Let ApellidoMaterno(sValue As String)
   ms_ApellidoMaterno = sValue
End Property
Property Get ApellidoMaterno() As String
   ApellidoMaterno = ms_ApellidoMaterno
End Property


