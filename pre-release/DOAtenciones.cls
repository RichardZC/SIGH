VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Atencion
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------

Option Explicit

Dim ml_Auditoria As Long
Dim ms_PisoDomicilio As String
Dim ms_HoraIngreso As String
Dim mda_FechaIngreso As Date
Dim ml_IdTipoServicio As Long
Dim ml_idPaciente As Long
Dim ml_idAtencion As Long
Dim ml_IdTipoCondicionALEstab As Long
Dim mda_FechaEgresoAdministrativo As Date

Dim ml_IdCamaEgreso As Long
Dim ml_IdCamaIngreso As Long
Dim ml_IdServicioEgreso As Long
Dim ml_IdTipoAlta As Long
Dim ml_IdCondicionAlta As Long
Dim ml_IdTipoEdad As Long

Dim ml_IdOrigenAtencion As Long

Dim ml_IdDestinoAtencion As Long
Dim ms_HoraEgresoAdministrativo As String
Dim ml_IdTipoCondicionAlServicio As Long
Dim ms_HoraEgreso As String
Dim mda_FechaEgreso As Date
Dim ml_IdMedicoEgreso As Long
Dim ml_Edad As Long
Dim ml_IdEspecialidadMedico As Long
Dim ml_IdMedicoIngreso As Long
Dim ml_IdServicioIngreso As Long
Dim ml_IdTipoGravedad As Long
Dim ml_idCuentaAtencion As Long

Dim ml_IdComprobantePago As Long
'*********debb*******************
Dim ml_IdFormaPago As Long
Dim ml_idFuenteFinanciamiento As Long
Dim ml_IdEstadoAtencion As Long
Dim mb_EsPacienteExterno As Boolean
Dim ml_idSunasaPacienteHistorico As Long
Dim ms_horaInicioAtencion As String
Dim ms_EpsPorcentaje As Double
Property Let EpsPorcentaje(sValue As Double)
   ms_EpsPorcentaje = sValue
End Property
Property Get EpsPorcentaje() As Double
   EpsPorcentaje = ms_EpsPorcentaje
End Property

Property Let HoraInicioAtencion(sValue As String)
   ms_horaInicioAtencion = sValue
End Property
Property Get HoraInicioAtencion() As String
   HoraInicioAtencion = ms_horaInicioAtencion
End Property



Property Let idSunasaPacienteHistorico(lValue As Long)
   ml_idSunasaPacienteHistorico = lValue
End Property
Property Get idSunasaPacienteHistorico() As Long
   idSunasaPacienteHistorico = ml_idSunasaPacienteHistorico
End Property


Property Let EsPacienteExterno(bValue As Boolean)
   mb_EsPacienteExterno = bValue
End Property
Property Get EsPacienteExterno() As Boolean
   EsPacienteExterno = mb_EsPacienteExterno
End Property

Property Let IdEstadoAtencion(lValue As Long)
   ml_IdEstadoAtencion = lValue
End Property
Property Get IdEstadoAtencion() As Long
   IdEstadoAtencion = ml_IdEstadoAtencion
End Property
Property Let idFuenteFinanciamiento(lValue As Long)
   ml_idFuenteFinanciamiento = lValue
End Property
Property Get idFuenteFinanciamiento() As Long
   idFuenteFinanciamiento = ml_idFuenteFinanciamiento
End Property

Property Let IdFormaPago(lValue As Long)
   ml_IdFormaPago = lValue
End Property
Property Get IdFormaPago() As Long
   IdFormaPago = ml_IdFormaPago
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let PisoDomicilio(sValue As String)
   ms_PisoDomicilio = sValue
End Property
Property Get PisoDomicilio() As String
   PisoDomicilio = ms_PisoDomicilio
End Property
Property Let HoraIngreso(sValue As String)
   ms_HoraIngreso = sValue
End Property
Property Get HoraIngreso() As String
   HoraIngreso = ms_HoraIngreso
End Property
Property Let FechaIngreso(daValue As Date)
   mda_FechaIngreso = daValue
End Property
Property Get FechaIngreso() As Date
   FechaIngreso = mda_FechaIngreso
End Property
Property Let IdTipoServicio(lValue As Long)
   ml_IdTipoServicio = lValue
End Property
Property Get IdTipoServicio() As Long
   IdTipoServicio = ml_IdTipoServicio
End Property
Property Let idPaciente(lValue As Long)
   ml_idPaciente = lValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_idPaciente
End Property
Property Let idAtencion(lValue As Long)
   ml_idAtencion = lValue
End Property
Property Get idAtencion() As Long
   idAtencion = ml_idAtencion
End Property
Property Let IdTipoCondicionALEstab(lValue As Long)
   ml_IdTipoCondicionALEstab = lValue
End Property
Property Get IdTipoCondicionALEstab() As Long
   IdTipoCondicionALEstab = ml_IdTipoCondicionALEstab
End Property
Property Let FechaEgresoAdministrativo(daValue As Date)
   mda_FechaEgresoAdministrativo = daValue
End Property
Property Get FechaEgresoAdministrativo() As Date
   FechaEgresoAdministrativo = mda_FechaEgresoAdministrativo
End Property

Property Let IdCamaEgreso(lValue As Long)
   ml_IdCamaEgreso = lValue
End Property
Property Get IdCamaEgreso() As Long
   IdCamaEgreso = ml_IdCamaEgreso
End Property
Property Let IdCamaIngreso(lValue As Long)
   ml_IdCamaIngreso = lValue
End Property
Property Get IdCamaIngreso() As Long
   IdCamaIngreso = ml_IdCamaIngreso
End Property
Property Let IdServicioEgreso(lValue As Long)
   ml_IdServicioEgreso = lValue
End Property
Property Get IdServicioEgreso() As Long
   IdServicioEgreso = ml_IdServicioEgreso
End Property
Property Let IdTipoAlta(lValue As Long)
   ml_IdTipoAlta = lValue
End Property
Property Get IdTipoAlta() As Long
   IdTipoAlta = ml_IdTipoAlta
End Property
Property Let IdCondicionAlta(lValue As Long)
   ml_IdCondicionAlta = lValue
End Property
Property Get IdCondicionAlta() As Long
   IdCondicionAlta = ml_IdCondicionAlta
End Property
Property Let IdTipoEdad(lValue As Long)
   ml_IdTipoEdad = lValue
End Property
Property Get IdTipoEdad() As Long
   IdTipoEdad = ml_IdTipoEdad
End Property

Property Let IdOrigenAtencion(lValue As Long)
   ml_IdOrigenAtencion = lValue
End Property
Property Get IdOrigenAtencion() As Long
   IdOrigenAtencion = ml_IdOrigenAtencion
End Property

Property Let IdDestinoAtencion(lValue As Long)
   ml_IdDestinoAtencion = lValue
End Property
Property Get IdDestinoAtencion() As Long
   IdDestinoAtencion = ml_IdDestinoAtencion
End Property
Property Let HoraEgresoAdministrativo(sValue As String)
   ms_HoraEgresoAdministrativo = sValue
End Property
Property Get HoraEgresoAdministrativo() As String
   HoraEgresoAdministrativo = ms_HoraEgresoAdministrativo
End Property
Property Let IdTipoCondicionAlServicio(lValue As Long)
   ml_IdTipoCondicionAlServicio = lValue
End Property
Property Get IdTipoCondicionAlServicio() As Long
   IdTipoCondicionAlServicio = ml_IdTipoCondicionAlServicio
End Property
Property Let HoraEgreso(sValue As String)
   ms_HoraEgreso = sValue
End Property
Property Get HoraEgreso() As String
   HoraEgreso = ms_HoraEgreso
End Property
Property Let FechaEgreso(daValue As Date)
   mda_FechaEgreso = daValue
End Property
Property Get FechaEgreso() As Date
   FechaEgreso = mda_FechaEgreso
End Property
Property Let IdMedicoEgreso(lValue As Long)
   ml_IdMedicoEgreso = lValue
End Property
Property Get IdMedicoEgreso() As Long
   IdMedicoEgreso = ml_IdMedicoEgreso
End Property


Property Let Edad(lValue As Long)
   ml_Edad = lValue
End Property
Property Get Edad() As Long
   Edad = ml_Edad
End Property
Property Let IdEspecialidadMedico(lValue As Long)
   ml_IdEspecialidadMedico = lValue
End Property
Property Get IdEspecialidadMedico() As Long
   IdEspecialidadMedico = ml_IdEspecialidadMedico
End Property
Property Let IdMedicoIngreso(lValue As Long)
   ml_IdMedicoIngreso = lValue
End Property
Property Get IdMedicoIngreso() As Long
   IdMedicoIngreso = ml_IdMedicoIngreso
End Property
Property Let IdServicioIngreso(lValue As Long)
   ml_IdServicioIngreso = lValue
End Property
Property Get IdServicioIngreso() As Long
   IdServicioIngreso = ml_IdServicioIngreso
End Property
Property Let IdTipoGravedad(lValue As Long)
   ml_IdTipoGravedad = lValue
End Property
Property Get IdTipoGravedad() As Long
   IdTipoGravedad = ml_IdTipoGravedad
End Property
Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property


