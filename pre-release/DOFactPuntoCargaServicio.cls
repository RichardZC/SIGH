VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOFactPuntoCargaServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactPuntosCargaServicio
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdServicioSubGrupo As Long
Dim ml_idPuntoCarga As Long
Dim ml_IdPuntoCargaServicio As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdServicioSubGrupo(lValue As Long)
   ml_IdServicioSubGrupo = lValue
End Property
Property Get IdServicioSubGrupo() As Long
   IdServicioSubGrupo = ml_IdServicioSubGrupo
End Property
Property Let idPuntoCarga(lValue As Long)
   ml_idPuntoCarga = lValue
End Property
Property Get idPuntoCarga() As Long
   idPuntoCarga = ml_idPuntoCarga
End Property
Property Let IdPuntoCargaServicio(lValue As Long)
   ml_IdPuntoCargaServicio = lValue
End Property
Property Get IdPuntoCargaServicio() As Long
   IdPuntoCargaServicio = ml_IdPuntoCargaServicio
End Property


