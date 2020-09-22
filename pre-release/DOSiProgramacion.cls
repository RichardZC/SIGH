VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOSiProgramacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: SiProgramacion
'        Autor: Daniel Enrique Barrantes
'        Fecha: 03/01/2019 05:33:02 a.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdProgramacion As Long
Dim ml_IdResponsable As Long
Dim mda_Fecha As Date
Dim ms_HoraInicio As String
Dim ms_HoraFin As String
Dim ml_IdTurno As Long
Dim ml_IdSala As Long
Dim mda_FechaReg As Date
Dim ml_TiempoPromedioAtencion As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdProgramacion(iValue As Long)
   ml_IdProgramacion = iValue
End Property
Property Get IdProgramacion() As Long
   IdProgramacion = ml_IdProgramacion
End Property
Property Let IdResponsable(iValue As Long)
   ml_IdResponsable = iValue
End Property
Property Get IdResponsable() As Long
   IdResponsable = ml_IdResponsable
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
Property Let HoraFin(sValue As String)
   ms_HoraFin = sValue
End Property
Property Get HoraFin() As String
   HoraFin = ms_HoraFin
End Property
Property Let IdTurno(iValue As Long)
   ml_IdTurno = iValue
End Property
Property Get IdTurno() As Long
   IdTurno = ml_IdTurno
End Property
Property Let IdSala(iValue As Long)
   ml_IdSala = iValue
End Property
Property Get IdSala() As Long
   IdSala = ml_IdSala
End Property
Property Let FechaReg(daValue As Date)
   mda_FechaReg = daValue
End Property
Property Get FechaReg() As Date
   FechaReg = mda_FechaReg
End Property
Property Let TiempoPromedioAtencion(iValue As Long)
   ml_TiempoPromedioAtencion = iValue
End Property
Property Get TiempoPromedioAtencion() As Long
   TiempoPromedioAtencion = ml_TiempoPromedioAtencion
End Property

'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: SiProgramacion
'        Autor:   DAniel Barrantes B.
'        Fecha: 03/01/2019 05:33:02 a.m.
'------------------------------------------------------------------------------------
 

