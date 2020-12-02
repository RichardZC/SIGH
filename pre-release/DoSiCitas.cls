VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoSiCitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: SiCitas
'        Autor: Daniel Enrique Barrantes
'        Fecha: 13/11/2018 09:18:59 a.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idCitaSI As Long
Dim ml_idPuntoCarga As Long
Dim mda_Fecha As Date
Dim ms_HoraInicio As String
Dim ms_HoraFinal As String
Dim ml_idCuentaAtencion As Long
Dim ml_idComprobantePago As Long
Dim ms_Paciente As String
Dim ml_idTipoSexo As Long
Dim mda_FechaNacimiento As Date
Dim ml_idPaciente As Long
Dim ml_idMedico As Long
Dim ml_idMovimiento As Long
Dim ml_idUsuario As Long
Dim mda_FechaCreacion As Date
Dim ml_idEstado As Long
Dim ml_IdReceta As Long
Dim ml_IdSala As Long
Dim ms_llaveTicket As String
Dim ml_Cupo As Long
Dim ms_Telefono As String
Dim ms_Direccion As String
Dim ml_IdProgramacion As Long
Dim ml_IdResponsable As Long

Property Let IdResponsable(lValue As Long)
   ml_IdResponsable = lValue
End Property
Property Get IdResponsable() As Long
   IdResponsable = ml_IdResponsable
End Property



Property Let idProgramacion(lValue As Long)
   ml_IdProgramacion = lValue
End Property
Property Get idProgramacion() As Long
   idProgramacion = ml_IdProgramacion
End Property

Property Let Telefono(sValue As String)
   ms_Telefono = sValue
End Property
Property Get Telefono() As String
   Telefono = ms_Telefono
End Property
Property Let Direccion(sValue As String)
   ms_Direccion = sValue
End Property
Property Get Direccion() As String
   Direccion = ms_Direccion
End Property

Property Let Cupo(lValue As Long)
   ml_Cupo = lValue
End Property
Property Get Cupo() As Long
   Cupo = ml_Cupo
End Property
Property Let llaveTicket(sValue As String)
   ms_llaveTicket = sValue
End Property
Property Get llaveTicket() As String
   llaveTicket = ms_llaveTicket
End Property


Property Let IdSala(lValue As Long)
   ml_IdSala = lValue
End Property
Property Get IdSala() As Long
   IdSala = ml_IdSala
End Property


Property Let IdReceta(lValue As Long)
   ml_IdReceta = lValue
End Property
Property Get IdReceta() As Long
   IdReceta = ml_IdReceta
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idCitaSI(iValue As Long)
   ml_idCitaSI = iValue
End Property
Property Get idCitaSI() As Long
   idCitaSI = ml_idCitaSI
End Property
Property Let idPuntoCarga(iValue As Long)
   ml_idPuntoCarga = iValue
End Property
Property Get idPuntoCarga() As Long
   idPuntoCarga = ml_idPuntoCarga
End Property
Property Let Fecha(daValue As Date)
   mda_Fecha = daValue
End Property
Property Get Fecha() As Date
   Fecha = mda_Fecha
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
Property Let idCuentaAtencion(iValue As Long)
   ml_idCuentaAtencion = iValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let idComprobantePago(iValue As Long)
   ml_idComprobantePago = iValue
End Property
Property Get idComprobantePago() As Long
   idComprobantePago = ml_idComprobantePago
End Property
Property Let Paciente(sValue As String)
   ms_Paciente = sValue
End Property
Property Get Paciente() As String
   Paciente = ms_Paciente
End Property
Property Let idTipoSexo(iValue As Long)
   ml_idTipoSexo = iValue
End Property
Property Get idTipoSexo() As Long
   idTipoSexo = ml_idTipoSexo
End Property
Property Let FechaNacimiento(daValue As Date)
   mda_FechaNacimiento = daValue
End Property
Property Get FechaNacimiento() As Date
   FechaNacimiento = mda_FechaNacimiento
End Property
Property Let idPaciente(iValue As Long)
   ml_idPaciente = iValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_idPaciente
End Property
Property Let idMedico(iValue As Long)
   ml_idMedico = iValue
End Property
Property Get idMedico() As Long
   idMedico = ml_idMedico
End Property
Property Let idMovimiento(iValue As Long)
   ml_idMovimiento = iValue
End Property
Property Get idMovimiento() As Long
   idMovimiento = ml_idMovimiento
End Property
Property Let idUsuario(iValue As Long)
   ml_idUsuario = iValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let FechaCreacion(daValue As Date)
   mda_FechaCreacion = daValue
End Property
Property Get FechaCreacion() As Date
   FechaCreacion = mda_FechaCreacion
End Property
Property Let idEstado(iValue As Long)
   ml_idEstado = iValue
End Property
Property Get idEstado() As Long
   idEstado = ml_idEstado
End Property

'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: SiCitas
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: SiCitas
'        Autor:   DAniel Barrantes B.
'        Fecha: 13/11/2018 09:19:00 a.m.
'------------------------------------------------------------------------------------
 

