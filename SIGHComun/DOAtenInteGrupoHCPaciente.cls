VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtenInteGrupoHCPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtenInteGrupoHCPaciente
'        Programado por: Garay M
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------

Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdPaciente As Long
Dim ml_IdGrupoHCPaciente As Long
Dim ml_IdAtenInteGrupo As Long
Dim ml_IdPregunta As Long
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
Property Let IdGrupoHCPaciente(iValue As Long)
   ml_IdGrupoHCPaciente = iValue
End Property
Property Get IdGrupoHCPaciente() As Long
   IdGrupoHCPaciente = ml_IdGrupoHCPaciente
End Property
Property Let IdAtenInteGrupo(iValue As Long)
   ml_IdAtenInteGrupo = iValue
End Property
Property Get IdAtenInteGrupo() As Long
   IdAtenInteGrupo = ml_IdAtenInteGrupo
End Property
Property Let IdPregunta(iValue As Long)
   ml_IdPregunta = iValue
End Property
Property Get IdPregunta() As Long
   IdPregunta = ml_IdPregunta
End Property
