VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOTiposImpuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla TiposImpuesto
'        Programado por: Castro W
'        Fecha: Julio 2004
'
'------------------------------------------------------------------------------------

Option Explicit

Dim ml_Auditoria As Long
Dim mc_Valor As Currency
Dim ms_CodImpuesto As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Valor(cValue As Currency)
   mc_Valor = cValue
End Property
Property Get Valor() As Currency
   Valor = mc_Valor
End Property
Property Let CodImpuesto(sValue As String)
   ms_CodImpuesto = sValue
End Property
Property Get CodImpuesto() As String
   CodImpuesto = ms_CodImpuesto
End Property

