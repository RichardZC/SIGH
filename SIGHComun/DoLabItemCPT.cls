VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoLabItemsCPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla LabItemsCpt
'        Programado por: Bonilla A
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idProductoCpt As Long
Dim ml_ordenXresultado As Long
Dim ml_idGrupo As Long
Dim ml_idItemGrupo As Long
Dim ml_idItem As Long
Dim ms_ValorSiEsCombo As String
Dim ms_ValorReferencial As String
Dim ms_Metodo As String
Dim mb_SoloNumero As Boolean
Dim mb_SoloTexto As Boolean
Dim mb_SoloCombo As Boolean
Dim mb_SoloCheck As Boolean

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idProductoCpt(iValue As Long)
   ml_idProductoCpt = iValue
End Property
Property Get idProductoCpt() As Long
   idProductoCpt = ml_idProductoCpt
End Property
Property Let ordenXresultado(iValue As Long)
   ml_ordenXresultado = iValue
End Property
Property Get ordenXresultado() As Long
   ordenXresultado = ml_ordenXresultado
End Property
Property Let idGrupo(iValue As Long)
   ml_idGrupo = iValue
End Property
Property Get idGrupo() As Long
   idGrupo = ml_idGrupo
End Property
Property Let idItemGrupo(iValue As Long)
   ml_idItemGrupo = iValue
End Property
Property Get idItemGrupo() As Long
   idItemGrupo = ml_idItemGrupo
End Property
Property Let idItem(iValue As Long)
   ml_idItem = iValue
End Property
Property Get idItem() As Long
   idItem = ml_idItem
End Property
Property Let ValorSiEsCombo(sValue As String)
   ms_ValorSiEsCombo = sValue
End Property
Property Get ValorSiEsCombo() As String
   ValorSiEsCombo = ms_ValorSiEsCombo
End Property
Property Let ValorReferencial(sValue As String)
   ms_ValorReferencial = sValue
End Property
Property Get ValorReferencial() As String
   ValorReferencial = ms_ValorReferencial
End Property
Property Let Metodo(sValue As String)
   ms_Metodo = sValue
End Property
Property Get Metodo() As String
   Metodo = ms_Metodo
End Property
Property Let SoloNumero(bValue As Boolean)
   mb_SoloNumero = bValue
End Property
Property Get SoloNumero() As Boolean
   SoloNumero = mb_SoloNumero
End Property
Property Let SoloTexto(bValue As Boolean)
   mb_SoloTexto = bValue
End Property
Property Get SoloTexto() As Boolean
   SoloTexto = mb_SoloTexto
End Property
Property Let SoloCombo(bValue As Boolean)
   mb_SoloCombo = bValue
End Property
Property Get SoloCombo() As Boolean
   SoloCombo = mb_SoloCombo
End Property
Property Let SoloCheck(bValue As Boolean)
   mb_SoloCheck = bValue
End Property
Property Get SoloCheck() As Boolean
   SoloCheck = mb_SoloCheck
End Property


 

