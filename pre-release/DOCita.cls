VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Citas
'        Programado por: Castro W
'        Fecha: Enero 2006
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ms_HoraSolicitud As String
Dim mda_FechaSolicitud As Date
Dim ml_IdProducto As Long
Dim ml_IdProgramacion As Long
Dim ml_IdServicio As Long
Dim ms_HoraFin As String
Dim ms_HoraInicio As String
Dim ml_IdCita As Long
Dim mda_fecha As Date
Dim ml_IdEstadoCita As Long
Dim ml_idMedico As Long
Dim ml_IdEspecialidad As Long
Dim ml_IdAtencion As Long
Dim ml_IdPaciente As Long
Dim mb_EsCitaAdicional As Boolean

Property Let EsCitaAdicional(bValue As Boolean)
   mb_EsCitaAdicional = bValue
End Property
Property Get EsCitaAdicional() As Boolean
   EsCitaAdicional = mb_EsCitaAdicional
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let IdAtencion(lValue As Long)
   ml_IdAtencion = lValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property
Property Let IdEspecialidad(lValue As Long)
   ml_IdEspecialidad = lValue
End Property
Property Get IdEspecialidad() As Long
   IdEspecialidad = ml_IdEspecialidad
End Property
Property Let idMedico(lValue As Long)
   ml_idMedico = lValue
End Property
Property Get idMedico() As Long
   idMedico = ml_idMedico
End Property
Property Let IdEstadoCita(lValue As Long)
   ml_IdEstadoCita = lValue
End Property
Property Get IdEstadoCita() As Long
   IdEstadoCita = ml_IdEstadoCita
End Property
Property Let fecha(daValue As Date)
   mda_fecha = daValue
End Property
Property Get fecha() As Date
   fecha = mda_fecha
End Property
Property Let IdCita(lValue As Long)
   ml_IdCita = lValue
End Property
Property Get IdCita() As Long
   IdCita = ml_IdCita
End Property
Property Let HoraSolicitud(sValue As String)
   ms_HoraSolicitud = sValue
End Property
Property Get HoraSolicitud() As String
   HoraSolicitud = ms_HoraSolicitud
End Property
Property Let FechaSolicitud(daValue As Date)
   mda_FechaSolicitud = daValue
End Property
Property Get FechaSolicitud() As Date
   FechaSolicitud = mda_FechaSolicitud
End Property
Property Let IdProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let IdProgramacion(lValue As Long)
   ml_IdProgramacion = lValue
End Property
Property Get IdProgramacion() As Long
   IdProgramacion = ml_IdProgramacion
End Property
Property Let IdServicio(lValue As Long)
   ml_IdServicio = lValue
End Property
Property Get IdServicio() As Long
   IdServicio = ml_IdServicio
End Property
Property Let HoraFin(sValue As String)
   ms_HoraFin = sValue
End Property
Property Get HoraFin() As String
   HoraFin = ms_HoraFin
End Property
Property Let HoraInicio(sValue As String)
   ms_HoraInicio = sValue
End Property
Property Get HoraInicio() As String
   HoraInicio = ms_HoraInicio
End Property


