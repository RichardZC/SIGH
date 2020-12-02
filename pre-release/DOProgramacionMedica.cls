VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOProgramacionMedica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla ProgramacionMedica
'        Programado por: Castro W
'        Fecha: Diciembre 2004
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdEspecialidad As Long
Dim ml_idTurno As Long
Dim ml_Color As Long
Dim ml_IdServicio As Long
Dim ml_IdProgramacion As Long
Dim ml_idMedico As Long
Dim ml_IdDepartamento As Long
Dim ml_IdTipoServicio As Long
Dim mda_fecha As Date
Dim ms_HoraInicio As String
Dim ms_HoraFin As String
Dim ms_Descripcion As String
Dim ml_IdTipoProgramacion As Long
Dim mda_FechaReg As Date
Dim ml_TiempoPromedioAtencion As Long
Property Let TiempoPromedioAtencion(lValue As Long)
   ml_TiempoPromedioAtencion = lValue
End Property
Property Get TiempoPromedioAtencion() As Long
   TiempoPromedioAtencion = ml_TiempoPromedioAtencion
End Property

Property Let FechaReg(daValue As Date)
   mda_FechaReg = daValue
End Property
Property Get FechaReg() As Date
   FechaReg = mda_FechaReg
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdTipoProgramacion(lValue As Long)
   ml_IdTipoProgramacion = lValue
End Property
Property Get IdTipoProgramacion() As Long
   IdTipoProgramacion = ml_IdTipoProgramacion
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
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
Property Let fecha(daValue As Date)
   mda_fecha = daValue
End Property
Property Get fecha() As Date
   fecha = mda_fecha
End Property
Property Let IdTipoServicio(lValue As Long)
   ml_IdTipoServicio = lValue
End Property
Property Get IdTipoServicio() As Long
   IdTipoServicio = ml_IdTipoServicio
End Property
Property Let IdDepartamento(lValue As Long)
   ml_IdDepartamento = lValue
End Property
Property Get IdDepartamento() As Long
   IdDepartamento = ml_IdDepartamento
End Property
Property Let idMedico(lValue As Long)
   ml_idMedico = lValue
End Property
Property Get idMedico() As Long
   idMedico = ml_idMedico
End Property
Property Let IdProgramacion(lValue As Long)
   ml_IdProgramacion = lValue
End Property
Property Get IdProgramacion() As Long
   IdProgramacion = ml_IdProgramacion
End Property
Property Let IdEspecialidad(lValue As Long)
   ml_IdEspecialidad = lValue
End Property
Property Get IdEspecialidad() As Long
   IdEspecialidad = ml_IdEspecialidad
End Property
Property Let idTurno(lValue As Long)
   ml_idTurno = lValue
End Property
Property Get idTurno() As Long
   idTurno = ml_idTurno
End Property
Property Let Color(lValue As Long)
   ml_Color = lValue
End Property
Property Get Color() As Long
   Color = ml_Color
End Property
Property Let IdServicio(lValue As Long)
   ml_IdServicio = lValue
End Property
Property Get IdServicio() As Long
   IdServicio = ml_IdServicio
End Property


