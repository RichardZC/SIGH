VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCitasWebCupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla CitasWebCupos
'        Programado por: Benavides M
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idWeb As Long
Dim mda_FECHA As Date
Dim ml_IdServicio As Long
Dim ml_idMedico As Long
Dim ms_HoraInicio As String
Dim ms_HoraFinal As String
Dim ml_idEstadoCitaWeb As Long
Dim ml_idCitaBloqueada As Long
Dim ms_DNI As String
Dim ms_ApellidoPaterno As String
Dim ms_ApellidoMaterno As String
Dim ms_PrimerNombre As String
Dim ms_SegundoNombre As String
Dim ml_IdTipoSexo As Long
Dim mda_FechaNacimiento As Date
Dim ml_Ubigeo As Long
Dim mda_FechaConfirmacion As Date
Dim ms_HoraConfirmacion As String
Dim ml_idFuenteFinanciamiento As Long
Dim ml_idTurno As Long
Dim ml_idPaciente As Long
Dim ms_Email As String
Dim ms_Telefono As String

Dim ml_refIdEess As Long
Dim ms_refNumero As String
Dim ms_refUPS As String
Dim ms_refidDiagnostico As Long
Dim ms_refMedicoCOLEGIATURA As String
Property Let refIdEess(lValue As Long)
   ml_refIdEess = lValue
End Property
Property Get refIdEess() As Long
   refIdEess = ml_refIdEess
End Property
Property Let refNumero(sValue As String)
   ms_refNumero = sValue
End Property
Property Get refNumero() As String
   refNumero = ms_refNumero
End Property
Property Let refUPS(sValue As String)
   ms_refUPS = sValue
End Property
Property Get refUPS() As String
   refUPS = ms_refUPS
End Property
Property Let refidDiagnostico(sValue As String)
   ms_refidDiagnostico = sValue
End Property
Property Get refidDiagnostico() As String
   refidDiagnostico = ms_refidDiagnostico
End Property
Property Let refMedicoCOLEGIATURA(sValue As String)
   ms_refMedicoCOLEGIATURA = sValue
End Property
Property Get refMedicoCOLEGIATURA() As String
   refMedicoCOLEGIATURA = ms_refMedicoCOLEGIATURA
End Property



Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idWeb(iValue As Long)
   ml_idWeb = iValue
End Property
Property Get idWeb() As Long
   idWeb = ml_idWeb
End Property
Property Let FECHA(daValue As Date)
   mda_FECHA = daValue
End Property
Property Get FECHA() As Date
   FECHA = mda_FECHA
End Property
Property Let IdServicio(iValue As Long)
   ml_IdServicio = iValue
End Property
Property Get IdServicio() As Long
   IdServicio = ml_IdServicio
End Property
Property Let idMedico(iValue As Long)
   ml_idMedico = iValue
End Property
Property Get idMedico() As Long
   idMedico = ml_idMedico
End Property
Property Let HoraInicio(sValue As String)
   ms_HoraInicio = sValue
End Property
Property Get HoraInicio() As String
   HoraInicio = ms_HoraInicio
End Property
Property Let HoraFinal(sValue As String)
   ms_HoraFinal = sValue
End Property
Property Get HoraFinal() As String
   HoraFinal = ms_HoraFinal
End Property
Property Let idEstadoCitaWeb(iValue As Long)
   ml_idEstadoCitaWeb = iValue
End Property
Property Get idEstadoCitaWeb() As Long
   idEstadoCitaWeb = ml_idEstadoCitaWeb
End Property
Property Let idCitaBloqueada(iValue As Long)
   ml_idCitaBloqueada = iValue
End Property
Property Get idCitaBloqueada() As Long
   idCitaBloqueada = ml_idCitaBloqueada
End Property
Property Let DNI(sValue As String)
   ms_DNI = sValue
End Property
Property Get DNI() As String
   DNI = ms_DNI
End Property
Property Let ApellidoPaterno(sValue As String)
   ms_ApellidoPaterno = sValue
End Property
Property Get ApellidoPaterno() As String
   ApellidoPaterno = ms_ApellidoPaterno
End Property
Property Let ApellidoMaterno(sValue As String)
   ms_ApellidoMaterno = sValue
End Property
Property Get ApellidoMaterno() As String
   ApellidoMaterno = ms_ApellidoMaterno
End Property
Property Let PrimerNombre(sValue As String)
   ms_PrimerNombre = sValue
End Property
Property Get PrimerNombre() As String
   PrimerNombre = ms_PrimerNombre
End Property
Property Let SegundoNombre(sValue As String)
   ms_SegundoNombre = sValue
End Property
Property Get SegundoNombre() As String
   SegundoNombre = ms_SegundoNombre
End Property
Property Let IdTipoSexo(iValue As Long)
   ml_IdTipoSexo = iValue
End Property
Property Get IdTipoSexo() As Long
   IdTipoSexo = ml_IdTipoSexo
End Property
Property Let FechaNacimiento(daValue As Date)
   mda_FechaNacimiento = daValue
End Property
Property Get FechaNacimiento() As Date
   FechaNacimiento = mda_FechaNacimiento
End Property
Property Let Ubigeo(iValue As Long)
   ml_Ubigeo = iValue
End Property
Property Get Ubigeo() As Long
   Ubigeo = ml_Ubigeo
End Property
Property Let FechaConfirmacion(daValue As Date)
   mda_FechaConfirmacion = daValue
End Property
Property Get FechaConfirmacion() As Date
   FechaConfirmacion = mda_FechaConfirmacion
End Property
Property Let HoraConfirmacion(sValue As String)
   ms_HoraConfirmacion = sValue
End Property
Property Get HoraConfirmacion() As String
   HoraConfirmacion = ms_HoraConfirmacion
End Property
Property Let idFuenteFinanciamiento(iValue As Long)
   ml_idFuenteFinanciamiento = iValue
End Property
Property Get idFuenteFinanciamiento() As Long
   idFuenteFinanciamiento = ml_idFuenteFinanciamiento
End Property
Property Let idTurno(iValue As Long)
   ml_idTurno = iValue
End Property
Property Get idTurno() As Long
   idTurno = ml_idTurno
End Property
Property Let idPaciente(iValue As Long)
   ml_idPaciente = iValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_idPaciente
End Property
Property Let Email(sValue As String)
   ms_Email = sValue
End Property
Property Get Email() As String
   Email = ms_Email
End Property
Property Let Telefono(sValue As String)
   ms_Telefono = sValue
End Property
Property Get Telefono() As String
   Telefono = ms_Telefono
End Property


