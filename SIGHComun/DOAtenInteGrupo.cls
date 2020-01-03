VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtenInteGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtenInteGrupo
'        Programado por: Garay M
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_IdUsuarioAuditoria As Long
Dim mi_IdAtenInteGrupo As Long
Dim ms_AtencionIntegralGrupo As String
Dim ml_DesdeAnio As Long
Dim mi_DesdeMes As Integer
Dim mi_DesdeDia As Integer
Dim ml_HastaAnio As Long
Dim mi_HastaMes As Integer
Dim mi_HastaDia As Integer

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let IdAtenInteGrupo(iValue As Long)
   mi_IdAtenInteGrupo = iValue
End Property
Property Get IdAtenInteGrupo() As Long
   IdAtenInteGrupo = mi_IdAtenInteGrupo
End Property
Property Let AtencionIntegralGrupo(sValue As String)
   ms_AtencionIntegralGrupo = sValue
End Property
Property Get AtencionIntegralGrupo() As String
   AtencionIntegralGrupo = ms_AtencionIntegralGrupo
End Property
Property Let DesdeAnio(iValue As Long)
   ml_DesdeAnio = iValue
End Property
Property Get DesdeAnio() As Long
   DesdeAnio = ml_DesdeAnio
End Property
Property Let DesdeMes(iValue As Long)
   mi_DesdeMes = iValue
End Property
Property Get DesdeMes() As Long
   DesdeMes = mi_DesdeMes
End Property
Property Let DesdeDia(iValue As Long)
   mi_DesdeDia = iValue
End Property
Property Get DesdeDia() As Long
   DesdeDia = mi_DesdeDia
End Property
Property Let HastaAnio(iValue As Long)
   ml_HastaAnio = iValue
End Property
Property Get HastaAnio() As Long
   HastaAnio = ml_HastaAnio
End Property
Property Let HastaMes(iValue As Long)
   mi_HastaMes = iValue
End Property
Property Get HastaMes() As Long
   HastaMes = mi_HastaMes
End Property
Property Let HastaDia(iValue As Long)
   mi_HastaDia = iValue
End Property
Property Get HastaDia() As Long
   HastaDia = mi_HastaDia
End Property
