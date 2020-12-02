VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOEstablecimientoNoMinsa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla EstablecimientosNoMinsa
'        Programado por: Castro W
'        Fecha: Abril 2006
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdEstablecimientoNoMinsa As Long
Dim ms_Nombre As String
Dim ml_IdTipoSubsector As Long
Dim ml_IdDistrito As Long
Dim ms_Codigo As String
Property Let Codigo(sValue As String)
   ms_Codigo = sValue
End Property
Property Get Codigo() As String
   Codigo = ms_Codigo
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdDistrito(lValue As Long)
   ml_IdDistrito = lValue
End Property
Property Get IdDistrito() As Long
   IdDistrito = ml_IdDistrito
End Property
Property Let IdTipoSubsector(lValue As Long)
   ml_IdTipoSubsector = lValue
End Property
Property Get IdTipoSubsector() As Long
   IdTipoSubsector = ml_IdTipoSubsector
End Property
Property Let Nombre(sValue As String)
   ms_Nombre = sValue
End Property
Property Get Nombre() As String
   Nombre = ms_Nombre
End Property
Property Let IdEstablecimientoNoMinsa(lValue As Long)
   ml_IdEstablecimientoNoMinsa = lValue
End Property
Property Get IdEstablecimientoNoMinsa() As Long
   IdEstablecimientoNoMinsa = ml_IdEstablecimientoNoMinsa
End Property


