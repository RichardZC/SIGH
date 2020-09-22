VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoAtencionHospCenso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: AtenHospCenso
'        Autor: Franklin Cachay Velasquez
'        Fecha: 27/02/2015 03:24:07
'        Auspiciado por: Ministerio de Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdRangoCensoHosp As Long
Dim mc_RangoInicial As Currency
Dim mc_RangoFinal As Currency
Dim ml_RGBRojo As Long
Dim ml_RGBVerde As Long
Dim ml_RGBAzul As Long

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdRangoCensoHosp(iValue As Long)
   ml_IdRangoCensoHosp = iValue
End Property
Property Get IdRangoCensoHosp() As Long
   IdRangoCensoHosp = ml_IdRangoCensoHosp
End Property
Property Let RangoInicial(cValue As Currency)
   mc_RangoInicial = cValue
End Property
Property Get RangoInicial() As Currency
   RangoInicial = mc_RangoInicial
End Property
Property Let RangoFinal(cValue As Currency)
   mc_RangoFinal = cValue
End Property
Property Get RangoFinal() As Currency
   RangoFinal = mc_RangoFinal
End Property
Property Let RGBRojo(iValue As Long)
   ml_RGBRojo = iValue
End Property
Property Get RGBRojo() As Long
   RGBRojo = ml_RGBRojo
End Property
Property Let RGBVerde(iValue As Long)
   ml_RGBVerde = iValue
End Property
Property Get RGBVerde() As Long
   RGBVerde = ml_RGBVerde
End Property
Property Let RGBAzul(iValue As Long)
   ml_RGBAzul = iValue
End Property
Property Get RGBAzul() As Long
   RGBAzul = ml_RGBAzul
End Property
