VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtenIntePlantillaItemPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtenIntePlantillaItemPlan
'        Programado por: Garay M
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdPlantillaItemPlan As Long
Dim mi_IdAtenInteGrupo As Integer
Dim ml_IdAtenInteItemPlan As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdPlantillaItemPlan(iValue As Long)
   ml_IdPlantillaItemPlan = iValue
End Property
Property Get IdPlantillaItemPlan() As Long
   IdPlantillaItemPlan = ml_IdPlantillaItemPlan
End Property
Property Let IdAtenInteGrupo(iValue As Integer)
   mi_IdAtenInteGrupo = iValue
End Property
Property Get IdAtenInteGrupo() As Integer
   IdAtenInteGrupo = mi_IdAtenInteGrupo
End Property
Property Let IdAtenInteItemPlan(iValue As Long)
   ml_IdAtenInteItemPlan = iValue
End Property
Property Get IdAtenInteItemPlan() As Long
   IdAtenInteItemPlan = ml_IdAtenInteItemPlan
End Property
