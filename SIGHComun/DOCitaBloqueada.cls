VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCitaBloqueada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla CitasBloqueadas
'        Programado por: Castro W
'        Fecha: Enero 2006
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_idCitaBloqueada As Long
Dim ms_HoraBloqueo As String
Dim mda_FechaBloqueo As Date
Dim ml_idMedico As Long
Dim ms_HoraFin As String
Dim ms_HoraInicio As String
Dim mda_fecha As Date
Dim ml_idUsuario As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let idCitaBloqueada(lValue As Long)
   ml_idCitaBloqueada = lValue
End Property
Property Get idCitaBloqueada() As Long
   idCitaBloqueada = ml_idCitaBloqueada
End Property
Property Let HoraBloqueo(sValue As String)
   ms_HoraBloqueo = sValue
End Property
Property Get HoraBloqueo() As String
   HoraBloqueo = ms_HoraBloqueo
End Property
Property Let FechaBloqueo(daValue As Date)
   mda_FechaBloqueo = daValue
End Property
Property Get FechaBloqueo() As Date
   FechaBloqueo = mda_FechaBloqueo
End Property
Property Let idMedico(lValue As Long)
   ml_idMedico = lValue
End Property
Property Get idMedico() As Long
   idMedico = ml_idMedico
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
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property


