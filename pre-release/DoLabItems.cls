VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoLabItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla LabItems
'        Programado por: Bonilla A
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------

Option Explicit


Dim ml_IdUsuarioAuditoria As Long
Dim ml_idItem As Long
Dim ms_Item As String
Dim ml_idProductoCpt As Long

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idItem(iValue As Long)
   ml_idItem = iValue
End Property
Property Get idItem() As Long
   idItem = ml_idItem
End Property
Property Let Item(sValue As String)
   ms_Item = sValue
End Property
Property Get Item() As String
   Item = ms_Item
End Property
Property Let idProductoCpt(iValue As Long)
   ml_idProductoCpt = iValue
End Property
Property Get idProductoCpt() As Long
   idProductoCpt = ml_idProductoCpt
End Property
