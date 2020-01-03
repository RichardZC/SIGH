VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOTiposTarifaCPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla TiposTarifaCpt
'        Programado por: Castro W
'        Fecha: Julio 2004
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idTipoTarifa As Long
Dim ml_idProductoCpt As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idTipoTarifa(iValue As Long)
   ml_idTipoTarifa = iValue
End Property
Property Get idTipoTarifa() As Long
   idTipoTarifa = ml_idTipoTarifa
End Property
Property Let idProductoCpt(iValue As Long)
   ml_idProductoCpt = iValue
End Property
Property Get idProductoCpt() As Long
   idProductoCpt = ml_idProductoCpt
End Property
