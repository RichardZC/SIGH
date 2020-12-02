VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOHIS_DetalleDx_Verifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla HIS_DetalleDx_Verifica
'        Programado por: Cachay F
'        Fecha: Octubre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdHisDetalle As Long
Dim ml_IdCIE As Long
Dim ml_IdSubClasificacionDX As Long
Dim ms_CodLAB As String

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdHisDetalle(iValue As Long)
   ml_IdHisDetalle = iValue
End Property
Property Get IdHisDetalle() As Long
   IdHisDetalle = ml_IdHisDetalle
End Property
Property Let IdCIE(iValue As Long)
   ml_IdCIE = iValue
End Property
Property Get IdCIE() As Long
   IdCIE = ml_IdCIE
End Property
Property Let IdSubClasificacionDX(iValue As Long)
   ml_IdSubClasificacionDX = iValue
End Property
Property Get IdSubClasificacionDX() As Long
   IdSubClasificacionDX = ml_IdSubClasificacionDX
End Property
Property Let CodLAB(sValue As String)
   ms_CodLAB = sValue
End Property
Property Get CodLAB() As String
   CodLAB = ms_CodLAB
End Property

