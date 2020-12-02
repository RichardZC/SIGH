VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOHis_FactCatalogoServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla HIS_FACTCATALOGOSERVICIOS
'        Programado por: Cachay F
'        Fecha: Octubre 2013
'
'------------------------------------------------------------------------------------
Option Explicit


Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdDiagCpt As Long
Dim ms_CodigoDiagCpt As String
Dim ms_DescripcionDiagCpt As String
Dim ms_EsCpt As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdDiagCpt(iValue As Long)
   ml_IdDiagCpt = iValue
End Property
Property Get IdDiagCpt() As Long
   IdDiagCpt = ml_IdDiagCpt
End Property
Property Let CodigoDiagCpt(sValue As String)
   ms_CodigoDiagCpt = sValue
End Property
Property Get CodigoDiagCpt() As String
   CodigoDiagCpt = ms_CodigoDiagCpt
End Property
Property Let DescripcionDiagCpt(sValue As String)
   ms_DescripcionDiagCpt = sValue
End Property
Property Get DescripcionDiagCpt() As String
   DescripcionDiagCpt = ms_DescripcionDiagCpt
End Property
Property Let EsCpt(sValue As String)
   ms_EsCpt = sValue
End Property
Property Get EsCpt() As String
   EsCpt = ms_EsCpt
End Property
