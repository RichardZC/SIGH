VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOInterLabCabecera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: InterLabCabecera
'        Autor: Daniel Enrique Barrantes
'        Fecha: 12/05/2018 09:32:39 a.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_identificador As Long
Dim ms_num_orden As String
Dim ms_Historia As String
Dim ms_PacienteEmail As String
Dim ms_PacienteFiliacion As String
Dim ml_PacienteId As Long
Dim ml_PacienteTipoDocumento As Long
Dim ms_PacienteDocumento As String
Dim ms_paciente As String
Dim ms_sexo As String
Dim mda_fnacimiento As Date
Dim ms_cama As String
Dim ms_dxCodigo As String
Dim ms_dx As String
Dim ms_medicoDNI As String
Dim ms_medico As String
Dim ms_servicioCodigo As String
Dim mda_fechaHoraMuestra As Date
Dim ml_idEstado As Long
Dim ms_medicoColegiatura As String
Property Let medicoColegiatura(sValue As String)
   ms_medicoColegiatura = sValue
End Property
Property Get medicoColegiatura() As String
   medicoColegiatura = ms_medicoColegiatura
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let identificador(iValue As Long)
   ml_identificador = iValue
End Property
Property Get identificador() As Long
   identificador = ml_identificador
End Property
Property Let num_orden(sValue As String)
   ms_num_orden = sValue
End Property
Property Get num_orden() As String
   num_orden = ms_num_orden
End Property
Property Let Historia(sValue As String)
   ms_Historia = sValue
End Property
Property Get Historia() As String
   Historia = ms_Historia
End Property
Property Let PacienteEmail(sValue As String)
   ms_PacienteEmail = sValue
End Property
Property Get PacienteEmail() As String
   PacienteEmail = ms_PacienteEmail
End Property
Property Let PacienteFiliacion(sValue As String)
   ms_PacienteFiliacion = sValue
End Property
Property Get PacienteFiliacion() As String
   PacienteFiliacion = ms_PacienteFiliacion
End Property
Property Let PacienteId(iValue As Long)
   ml_PacienteId = iValue
End Property
Property Get PacienteId() As Long
   PacienteId = ml_PacienteId
End Property
Property Let PacienteTipoDocumento(iValue As Long)
   ml_PacienteTipoDocumento = iValue
End Property
Property Get PacienteTipoDocumento() As Long
   PacienteTipoDocumento = ml_PacienteTipoDocumento
End Property
Property Let PacienteDocumento(sValue As String)
   ms_PacienteDocumento = sValue
End Property
Property Get PacienteDocumento() As String
   PacienteDocumento = ms_PacienteDocumento
End Property
Property Let paciente(sValue As String)
   ms_paciente = sValue
End Property
Property Get paciente() As String
   paciente = ms_paciente
End Property
Property Let sexo(sValue As String)
   ms_sexo = sValue
End Property
Property Get sexo() As String
   sexo = ms_sexo
End Property
Property Let fnacimiento(daValue As Date)
   mda_fnacimiento = daValue
End Property
Property Get fnacimiento() As Date
   fnacimiento = mda_fnacimiento
End Property
Property Let cama(sValue As String)
   ms_cama = sValue
End Property
Property Get cama() As String
   cama = ms_cama
End Property
Property Let dxCodigo(sValue As String)
   ms_dxCodigo = sValue
End Property
Property Get dxCodigo() As String
   dxCodigo = ms_dxCodigo
End Property
Property Let dx(sValue As String)
   ms_dx = sValue
End Property
Property Get dx() As String
   dx = ms_dx
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
Property Let servicioCodigo(sValue As String)
   ms_servicioCodigo = sValue
End Property
Property Get servicioCodigo() As String
   servicioCodigo = ms_servicioCodigo
End Property
Property Let fechaHoraMuestra(daValue As Date)
   mda_fechaHoraMuestra = daValue
End Property
Property Get fechaHoraMuestra() As Date
   fechaHoraMuestra = mda_fechaHoraMuestra
End Property
Property Let idEstado(iValue As Long)
   ml_idEstado = iValue
End Property
Property Get idEstado() As Long
   idEstado = ml_idEstado
End Property

