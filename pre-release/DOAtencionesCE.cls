VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtencionesCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtencionesCE
'        Programado por: Barrantes D
'        Fecha: Marzo 2011
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdAtencion As Long
Dim ml_NroHistoriaClinica As Long
Dim ms_CitaDniMedicoJamo As String
Dim mda_CitaFecha As Date
Dim ms_CitaMedico As String
Dim ms_CitaServicioJamo As String
Dim ml_CitaIdServicio As Long
Dim ms_CitaMotivo As String
Dim ms_CitaExamenClinico As String
Dim ms_CitaDiagMed As String
Dim ms_CitaExClinicos As String
Dim ms_CitaTratamiento As String
Dim ms_CitaObservaciones As String
Dim mda_CitaFechaAtencion As Date
Dim ml_CitaIdUsuario As Long
Dim ms_TriajeEdad As String
Dim ms_TriajePresion As String
Dim ms_TriajeTalla As String
Dim ms_TriajeTemperatura As String
Dim ms_TriajePeso As String
Dim mda_TriajeFecha As Date
Dim ml_TriajeIdUsuario As Long
Dim ml_TriajePulso As Long
Dim ml_TriajeFrecRespiratoria As Long
Dim ms_CitaAntecedente As String
Dim ml_TriajePerimCefalico As Currency
Dim ml_TriajeFrecCardiaca As Long
Dim ml_TriajeOrigen As Long
Dim ms_TriajePerimAbdominal As String
Dim ms_TriajeSaturacionOxigeno As String

Property Let TriajeSaturacionOxigeno(sValue As String)
   ms_TriajeSaturacionOxigeno = sValue
End Property
Property Get TriajeSaturacionOxigeno() As String
   TriajeSaturacionOxigeno = ms_TriajeSaturacionOxigeno
End Property

Property Let TriajePerimAbdominal(sValue As String)
   ms_TriajePerimAbdominal = sValue
End Property
Property Get TriajePerimAbdominal() As String
   TriajePerimAbdominal = ms_TriajePerimAbdominal
End Property

Property Let TriajeFrecRespiratoria(lValue As Long)
   ml_TriajeFrecRespiratoria = lValue
End Property
Property Get TriajeFrecRespiratoria() As Long
   TriajeFrecRespiratoria = ml_TriajeFrecRespiratoria
End Property
Property Let TriajePulso(lValue As Long)
   ml_TriajePulso = lValue
End Property
Property Get TriajePulso() As Long
   TriajePulso = ml_TriajePulso
End Property

Property Let TriajePerimCefalico(lValue As Currency)
   ml_TriajePerimCefalico = lValue
End Property
Property Get TriajePerimCefalico() As Currency
   TriajePerimCefalico = ml_TriajePerimCefalico
End Property

Property Let TriajeFrecCardiaca(lValue As Long)
   ml_TriajeFrecCardiaca = lValue
End Property
Property Get TriajeFrecCardiaca() As Long
   TriajeFrecCardiaca = ml_TriajeFrecCardiaca
End Property

Property Let TriajeOrigen(lValue As Long)
   ml_TriajeOrigen = lValue
End Property
Property Get TriajeOrigen() As Long
   TriajeOrigen = ml_TriajeOrigen
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdAtencion(iValue As Long)
   ml_IdAtencion = iValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property
Property Let NroHistoriaClinica(iValue As Long)
   ml_NroHistoriaClinica = iValue
End Property
Property Get NroHistoriaClinica() As Long
   NroHistoriaClinica = ml_NroHistoriaClinica
End Property
Property Let CitaDniMedicoJamo(sValue As String)
   ms_CitaDniMedicoJamo = sValue
End Property
Property Get CitaDniMedicoJamo() As String
   CitaDniMedicoJamo = ms_CitaDniMedicoJamo
End Property
Property Let CitaFecha(daValue As Date)
   mda_CitaFecha = daValue
End Property
Property Get CitaFecha() As Date
   CitaFecha = mda_CitaFecha
End Property
Property Let CitaMedico(sValue As String)
   ms_CitaMedico = sValue
End Property
Property Get CitaMedico() As String
   CitaMedico = ms_CitaMedico
End Property
Property Let CitaServicioJamo(sValue As String)
   ms_CitaServicioJamo = sValue
End Property
Property Get CitaServicioJamo() As String
   CitaServicioJamo = ms_CitaServicioJamo
End Property
Property Let CitaIdServicio(iValue As Long)
   ml_CitaIdServicio = iValue
End Property
Property Get CitaIdServicio() As Long
   CitaIdServicio = ml_CitaIdServicio
End Property
Property Let CitaMotivo(sValue As String)
   ms_CitaMotivo = sValue
End Property
Property Get CitaMotivo() As String
   CitaMotivo = ms_CitaMotivo
End Property
Property Let CitaExamenClinico(sValue As String)
   ms_CitaExamenClinico = sValue
End Property
Property Get CitaExamenClinico() As String
   CitaExamenClinico = ms_CitaExamenClinico
End Property
Property Let CitaDiagMed(sValue As String)
   ms_CitaDiagMed = sValue
End Property
Property Get CitaDiagMed() As String
   CitaDiagMed = ms_CitaDiagMed
End Property
Property Let CitaExClinicos(sValue As String)
   ms_CitaExClinicos = sValue
End Property
Property Get CitaExClinicos() As String
   CitaExClinicos = ms_CitaExClinicos
End Property
Property Let CitaTratamiento(sValue As String)
   ms_CitaTratamiento = sValue
End Property
Property Get CitaTratamiento() As String
   CitaTratamiento = ms_CitaTratamiento
End Property
Property Let CitaObservaciones(sValue As String)
   ms_CitaObservaciones = sValue
End Property
Property Get CitaObservaciones() As String
   CitaObservaciones = ms_CitaObservaciones
End Property
Property Let CitaFechaAtencion(daValue As Date)
   mda_CitaFechaAtencion = daValue
End Property
Property Get CitaFechaAtencion() As Date
   CitaFechaAtencion = mda_CitaFechaAtencion
End Property
Property Let CitaIdUsuario(iValue As Long)
   ml_CitaIdUsuario = iValue
End Property
Property Get CitaIdUsuario() As Long
   CitaIdUsuario = ml_CitaIdUsuario
End Property
Property Let TriajeEdad(sValue As String)
   ms_TriajeEdad = sValue
End Property
Property Get TriajeEdad() As String
   TriajeEdad = ms_TriajeEdad
End Property
Property Let TriajePresion(sValue As String)
   ms_TriajePresion = sValue
End Property
Property Get TriajePresion() As String
   TriajePresion = ms_TriajePresion
End Property
Property Let TriajeTalla(sValue As String)
   ms_TriajeTalla = sValue
End Property
Property Get TriajeTalla() As String
   TriajeTalla = ms_TriajeTalla
End Property
Property Let TriajeTemperatura(sValue As String)
   ms_TriajeTemperatura = sValue
End Property
Property Get TriajeTemperatura() As String
   TriajeTemperatura = ms_TriajeTemperatura
End Property
Property Let TriajePeso(sValue As String)
   ms_TriajePeso = sValue
End Property
Property Get TriajePeso() As String
   TriajePeso = ms_TriajePeso
End Property
Property Let TriajeFecha(daValue As Date)
   mda_TriajeFecha = daValue
End Property
Property Get TriajeFecha() As Date
   TriajeFecha = mda_TriajeFecha
End Property
Property Let TriajeIdUsuario(iValue As Long)
   ml_TriajeIdUsuario = iValue
End Property
Property Get TriajeIdUsuario() As Long
   TriajeIdUsuario = ml_TriajeIdUsuario
End Property
Property Let CitaAntecedente(sValue As String)
   ms_CitaAntecedente = sValue
End Property
Property Get CitaAntecedente() As String
   CitaAntecedente = ms_CitaAntecedente
End Property


 

