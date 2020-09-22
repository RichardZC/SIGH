VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOFactPuntoCargaBienesInsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactPuntosCargaBienesInsumos
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdSubGrupoFarmacologico As Long
Dim ml_idPuntoCarga As Long
Dim ml_IdPuntoCargaBienInsumo As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdSubGrupoFarmacologico(lValue As Long)
   ml_IdSubGrupoFarmacologico = lValue
End Property
Property Get IdSubGrupoFarmacologico() As Long
   IdSubGrupoFarmacologico = ml_IdSubGrupoFarmacologico
End Property
Property Let idPuntoCarga(lValue As Long)
   ml_idPuntoCarga = lValue
End Property
Property Get idPuntoCarga() As Long
   idPuntoCarga = ml_idPuntoCarga
End Property
Property Let IdPuntoCargaBienInsumo(lValue As Long)
   ml_IdPuntoCargaBienInsumo = lValue
End Property
Property Get IdPuntoCargaBienInsumo() As Long
   IdPuntoCargaBienInsumo = ml_IdPuntoCargaBienInsumo
End Property


