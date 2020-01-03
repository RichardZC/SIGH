VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoInteroperaPacs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: InteroperaPacs
'        Autor: Daniel Enrique Barrantes
'        Fecha: 23/10/2018 05:52:12 a.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdOrden As Long
Dim ms_cptCodigo As String
Dim ms_cpt As String
Dim ml_PacienteId As Long
Dim ml_PacienteNroHistoria As Long
Dim ml_PacienteIdTipoDocumento As Long
Dim ms_PacienteDocumento As String
Dim ms_Paciente As String
Dim ms_PacienteSexo As String
Dim mda_PacienteFNacimiento As Date
Dim ms_medicoDNI As String
Dim ms_medico As String
Dim mda_fechaHoraMuestra As Date
Dim ms_EquipoCodigo As String
Dim ms_EquipoRuta As String
Dim ms_Resultado As String
Dim mda_ResultadoFechaHora As Date
Dim ms_ResultadoDNI As String
Dim ml_idEstado As Long
Dim ms_PuntoCarga As String
Dim ms_PacPaterno As String
Dim ms_PacMaterno As String
Dim ms_PacPrimerNombre As String
Dim ms_PacSegundoNombre As String
Dim ms_MedPaterno As String
Dim ms_MedMaterno As String
Dim ms_MedNombres As String
Dim ms_FteFinanciamiento As String
Dim ms_TipoServicio As String
Dim ms_PACS As String
Dim ms_LLAVE As String

Property Let PACS(sValue As String)
   ms_PACS = sValue
End Property
Property Get PACS() As String
   PACS = ms_PACS
End Property
Property Let LLAVE(sValue As String)
   ms_LLAVE = sValue
End Property
Property Get LLAVE() As String
   LLAVE = ms_LLAVE
End Property


Property Let PacPaterno(sValue As String)
   ms_PacPaterno = sValue
End Property
Property Get PacPaterno() As String
   PacPaterno = ms_PacPaterno
End Property
Property Let PacMaterno(sValue As String)
   ms_PacMaterno = sValue
End Property
Property Get PacMaterno() As String
   PacMaterno = ms_PacMaterno
End Property
Property Let PacPrimerNombre(sValue As String)
   ms_PacPrimerNombre = sValue
End Property
Property Get PacPrimerNombre() As String
   PacPrimerNombre = ms_PacPrimerNombre
End Property
Property Let PacSegundoNombre(sValue As String)
   ms_PacSegundoNombre = sValue
End Property
Property Get PacSegundoNombre() As String
   PacSegundoNombre = ms_PacSegundoNombre
End Property
Property Let MedPaterno(sValue As String)
   ms_MedPaterno = sValue
End Property
Property Get MedPaterno() As String
   MedPaterno = ms_MedPaterno
End Property
Property Let MedMaterno(sValue As String)
   ms_MedMaterno = sValue
End Property
Property Get MedMaterno() As String
   MedMaterno = ms_MedMaterno
End Property
Property Let MedNombres(sValue As String)
   ms_MedNombres = sValue
End Property
Property Get MedNombres() As String
   MedNombres = ms_MedNombres
End Property
Property Let FteFinanciamiento(sValue As String)
   ms_FteFinanciamiento = sValue
End Property
Property Get FteFinanciamiento() As String
   FteFinanciamiento = ms_FteFinanciamiento
End Property
Property Let TipoServicio(sValue As String)
   ms_TipoServicio = sValue
End Property
Property Get TipoServicio() As String
   TipoServicio = ms_TipoServicio
End Property



Property Let PuntoCarga(sValue As String)
   ms_PuntoCarga = sValue
End Property
Property Get PuntoCarga() As String
   PuntoCarga = ms_PuntoCarga
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdOrden(iValue As Long)
   ml_IdOrden = iValue
End Property
Property Get IdOrden() As Long
   IdOrden = ml_IdOrden
End Property
Property Let cptCodigo(sValue As String)
   ms_cptCodigo = sValue
End Property
Property Get cptCodigo() As String
   cptCodigo = ms_cptCodigo
End Property
Property Let cpt(sValue As String)
   ms_cpt = sValue
End Property
Property Get cpt() As String
   cpt = ms_cpt
End Property
Property Let PacienteId(iValue As Long)
   ml_PacienteId = iValue
End Property
Property Get PacienteId() As Long
   PacienteId = ml_PacienteId
End Property
Property Let PacienteNroHistoria(iValue As Long)
   ml_PacienteNroHistoria = iValue
End Property
Property Get PacienteNroHistoria() As Long
   PacienteNroHistoria = ml_PacienteNroHistoria
End Property
Property Let PacienteIdTipoDocumento(iValue As Long)
   ml_PacienteIdTipoDocumento = iValue
End Property
Property Get PacienteIdTipoDocumento() As Long
   PacienteIdTipoDocumento = ml_PacienteIdTipoDocumento
End Property
Property Let PacienteDocumento(sValue As String)
   ms_PacienteDocumento = sValue
End Property
Property Get PacienteDocumento() As String
   PacienteDocumento = ms_PacienteDocumento
End Property
Property Let Paciente(sValue As String)
   ms_Paciente = sValue
End Property
Property Get Paciente() As String
   Paciente = ms_Paciente
End Property
Property Let PacienteSexo(sValue As String)
   ms_PacienteSexo = sValue
End Property
Property Get PacienteSexo() As String
   PacienteSexo = ms_PacienteSexo
End Property
Property Let PacienteFNacimiento(daValue As Date)
   mda_PacienteFNacimiento = daValue
End Property
Property Get PacienteFNacimiento() As Date
   PacienteFNacimiento = mda_PacienteFNacimiento
End Property
Property Let medicoDNI(sValue As String)
   ms_medicoDNI = sValue
End Property
Property Get medicoDNI() As String
   medicoDNI = ms_medicoDNI
End Property
Property Let medico(sValue As String)
   ms_medico = sValue
End Property
Property Get medico() As String
   medico = ms_medico
End Property
Property Let fechaHoraMuestra(daValue As Date)
   mda_fechaHoraMuestra = daValue
End Property
Property Get fechaHoraMuestra() As Date
   fechaHoraMuestra = mda_fechaHoraMuestra
End Property
Property Let EquipoCodigo(sValue As String)
   ms_EquipoCodigo = sValue
End Property
Property Get EquipoCodigo() As String
   EquipoCodigo = ms_EquipoCodigo
End Property
Property Let EquipoRuta(sValue As String)
   ms_EquipoRuta = sValue
End Property
Property Get EquipoRuta() As String
   EquipoRuta = ms_EquipoRuta
End Property
Property Let Resultado(sValue As String)
   ms_Resultado = sValue
End Property
Property Get Resultado() As String
   Resultado = ms_Resultado
End Property
Property Let ResultadoFechaHora(daValue As Date)
   mda_ResultadoFechaHora = daValue
End Property
Property Get ResultadoFechaHora() As Date
   ResultadoFechaHora = mda_ResultadoFechaHora
End Property
Property Let ResultadoDNI(sValue As String)
   ms_ResultadoDNI = sValue
End Property
Property Get ResultadoDNI() As String
   ResultadoDNI = ms_ResultadoDNI
End Property
Property Let idEstado(iValue As Long)
   ml_idEstado = iValue
End Property
Property Get idEstado() As Long
   idEstado = ml_idEstado
End Property

'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: InteroperaPacs
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: InteroperaPacs
'        Autor:   DAniel Barrantes B.
'        Fecha: 23/10/2018 05:52:12 a.m.
'------------------------------------------------------------------------------------
 

