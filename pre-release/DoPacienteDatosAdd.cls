VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoPacienteDatosAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla PacientesDatosAdicionales
'        Programado por: Barrantes D.
'        Fecha: Mayo 2012
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdPaciente As Long
Dim ms_antecedentes As String
Dim ms_antecedAlergico As String
Dim ms_antecedObstetrico As String
Dim ms_antecedQuirurgico As String
Dim ms_antecedFamiliar As String
Dim ms_antecedPatologico As String
Dim mb_fNacimientoCalculada As Boolean

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdPaciente(iValue As Long)
   ml_IdPaciente = iValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let antecedentes(sValue As String)
   ms_antecedentes = sValue
End Property
Property Get antecedentes() As String
   antecedentes = ms_antecedentes
End Property


Property Let antecedAlergico(sValue As String)
   ms_antecedAlergico = sValue
End Property
Property Get antecedAlergico() As String
   antecedAlergico = ms_antecedAlergico
End Property
Property Let antecedObstetrico(sValue As String)
   ms_antecedObstetrico = sValue
End Property
Property Get antecedObstetrico() As String
   antecedObstetrico = ms_antecedObstetrico
End Property
Property Let antecedQuirurgico(sValue As String)
   ms_antecedQuirurgico = sValue
End Property
Property Get antecedQuirurgico() As String
   antecedQuirurgico = ms_antecedQuirurgico
End Property
Property Let antecedFamiliar(sValue As String)
   ms_antecedFamiliar = sValue
End Property
Property Get antecedFamiliar() As String
   antecedFamiliar = ms_antecedFamiliar
End Property
Property Let antecedPatologico(sValue As String)
   ms_antecedPatologico = sValue
End Property
Property Get antecedPatologico() As String
   antecedPatologico = ms_antecedPatologico
End Property

Property Let fNacimientoCalculada(sValue As Boolean)
   mb_fNacimientoCalculada = sValue
End Property
Property Get fNacimientoCalculada() As Boolean
   fNacimientoCalculada = mb_fNacimientoCalculada
End Property

