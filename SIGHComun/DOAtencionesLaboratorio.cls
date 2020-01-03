VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtencionesLaboratorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: AtencionesLabSinDxSinCpt
'        Autor: Daniel Enrique Barrantes
'        Fecha: 01/06/2018 10:17:15 a.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idAtencion As Long
Dim ml_Grupo As Long
Dim ml_SubGrupo As Long
Dim ms_lab As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idAtencion(iValue As Long)
   ml_idAtencion = iValue
End Property
Property Get idAtencion() As Long
   idAtencion = ml_idAtencion
End Property
Property Let Grupo(iValue As Long)
   ml_Grupo = iValue
End Property
Property Get Grupo() As Long
   Grupo = ml_Grupo
End Property
Property Let SubGrupo(iValue As Long)
   ml_SubGrupo = iValue
End Property
Property Get SubGrupo() As Long
   SubGrupo = ml_SubGrupo
End Property
Property Let lab(sValue As String)
   ms_lab = sValue
End Property
Property Get lab() As String
   lab = ms_lab
End Property
