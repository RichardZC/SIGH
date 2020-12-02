VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtenInteValorTalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtenInteValorTalla
'        Programado por: Garay M
'        Fecha: Noviembre 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdValorTalla As Long
Dim ml_IdTipoSexo As Long
Dim ml_EdadMeses As Long
Dim ml_NroDesviacion As Long
Dim mc_ValorTalla As Currency
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdValorTalla(iValue As Long)
   ml_IdValorTalla = iValue
End Property
Property Get IdValorTalla() As Long
   IdValorTalla = ml_IdValorTalla
End Property
Property Let IdTipoSexo(iValue As Long)
   ml_IdTipoSexo = iValue
End Property
Property Get IdTipoSexo() As Long
   IdTipoSexo = ml_IdTipoSexo
End Property
Property Let EdadMeses(iValue As Long)
   ml_EdadMeses = iValue
End Property
Property Get EdadMeses() As Long
   EdadMeses = ml_EdadMeses
End Property
Property Let NroDesviacion(iValue As Long)
   ml_NroDesviacion = iValue
End Property
Property Get NroDesviacion() As Long
   NroDesviacion = ml_NroDesviacion
End Property
Property Let ValorTalla(cValue As Currency)
   mc_ValorTalla = cValue
End Property
Property Get ValorTalla() As Currency
   ValorTalla = mc_ValorTalla
End Property
