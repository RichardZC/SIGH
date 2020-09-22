VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOInsumosSubGrupoFarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactInsumosSubGrupoFarmacologico
'        Programado por: Barrantes D
'        Fecha: Abril 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdGrupoFarmacologico As Long
Dim ms_Descripcion As String
Dim ml_IdSubGrupoFarmacologico As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdGrupoFarmacologico(lValue As Long)
   ml_IdGrupoFarmacologico = lValue
End Property
Property Get IdGrupoFarmacologico() As Long
   IdGrupoFarmacologico = ml_IdGrupoFarmacologico
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let IdSubGrupoFarmacologico(lValue As Long)
   ml_IdSubGrupoFarmacologico = lValue
End Property
Property Get IdSubGrupoFarmacologico() As Long
   IdSubGrupoFarmacologico = ml_IdSubGrupoFarmacologico
End Property


