VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOHistoriaClinica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla HistoriasClinicas
'        Programado por: Barrantes D
'        Fecha: Octubre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdTipoNumeracionAnterior As Long
Dim ml_NroHistoriaClinicaAnterior As Long
Dim ml_IdTipoNumeracion As Long
Dim ml_NroHistoriaClinica As Long
Dim mda_FechaCreacion As Date
Dim mda_FechaPasoAPasivo As Date
Dim ml_IdTipoHistoria As Long
Dim ml_IdEstadoHistoria As Long
Dim ml_idPaciente As Long
Dim mda_FechaUltimoMovimiento As Date
Property Get FechaUltimoMovimiento() As Date
   FechaUltimoMovimiento = mda_FechaUltimoMovimiento
End Property
Property Let FechaUltimoMovimiento(daValue As Date)
   mda_FechaUltimoMovimiento = daValue
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let idPaciente(lValue As Long)
   ml_idPaciente = lValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_idPaciente
End Property
Property Let IdEstadoHistoria(lValue As Long)
   ml_IdEstadoHistoria = lValue
End Property
Property Get IdEstadoHistoria() As Long
   IdEstadoHistoria = ml_IdEstadoHistoria
End Property
Property Let IdTipoHistoria(lValue As Long)
   ml_IdTipoHistoria = lValue
End Property
Property Get IdTipoHistoria() As Long
   IdTipoHistoria = ml_IdTipoHistoria
End Property
Property Let FechaPasoAPasivo(daValue As Date)
   mda_FechaPasoAPasivo = daValue
End Property
Property Get FechaPasoAPasivo() As Date
   FechaPasoAPasivo = mda_FechaPasoAPasivo
End Property
Property Let FechaCreacion(daValue As Date)
   mda_FechaCreacion = daValue
End Property
Property Get FechaCreacion() As Date
   FechaCreacion = mda_FechaCreacion
End Property
Property Let IdTipoNumeracionAnterior(lValue As Long)
   ml_IdTipoNumeracionAnterior = lValue
End Property
Property Get IdTipoNumeracionAnterior() As Long
   IdTipoNumeracionAnterior = ml_IdTipoNumeracionAnterior
End Property
Property Let NroHistoriaClinicaAnterior(lValue As Long)
   ml_NroHistoriaClinicaAnterior = lValue
End Property
Property Get NroHistoriaClinicaAnterior() As Long
   NroHistoriaClinicaAnterior = ml_NroHistoriaClinicaAnterior
End Property
Property Let IdTipoNumeracion(lValue As Long)
   ml_IdTipoNumeracion = lValue
End Property
Property Get IdTipoNumeracion() As Long
   IdTipoNumeracion = ml_IdTipoNumeracion
End Property
Property Let NroHistoriaClinica(lValue As Long)
   ml_NroHistoriaClinica = lValue
End Property
Property Get NroHistoriaClinica() As Long
   NroHistoriaClinica = ml_NroHistoriaClinica
End Property


