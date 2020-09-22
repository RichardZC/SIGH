VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtenInteValorPeso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: AtenInteValorPeso
'        Autor: Daniel Enrique Barrantes
'        Fecha: 11/11/2014 07:12:29
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdValorPeso As Long
Dim ml_IdTipoSexo As Long
Dim ml_EdadMeses As Long
Dim ml_NroDesviacion As Long
Dim mc_ValorPeso As Currency
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdValorPeso(iValue As Long)
   ml_IdValorPeso = iValue
End Property
Property Get IdValorPeso() As Long
   IdValorPeso = ml_IdValorPeso
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
Property Let ValorPeso(cValue As Currency)
   mc_ValorPeso = cValue
End Property
Property Get ValorPeso() As Currency
   ValorPeso = mc_ValorPeso
End Property
