VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOTiposTarifa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla TiposTarifa
'        Programado por: Castro W
'        Fecha: Julio 2004
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idTipoTarifa As Long
Dim ms_Codigo As String
Dim ms_TipoTarifa As String
Dim ml_EsFarmacia As Boolean


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
Property Let Codigo(sValue As String)
   ms_Codigo = sValue
End Property
Property Get Codigo() As String
   Codigo = ms_Codigo
End Property
Property Let TipoTarifa(sValue As String)
   ms_TipoTarifa = sValue
End Property
Property Get TipoTarifa() As String
   TipoTarifa = ms_TipoTarifa
End Property

Property Let EsFarmacia(lValue As Boolean)
   ml_EsFarmacia = lValue
End Property
Property Get EsFarmacia() As Boolean
   EsFarmacia = ml_EsFarmacia
End Property




