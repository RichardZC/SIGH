VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOPeriodoTiempo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla PeriodoTiempo
'        Programado por: Barrantes D
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim mi_IdPeriodoTiempo As Integer
Dim ms_PeriodoTiempo As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdPeriodoTiempo(iValue As Integer)
   mi_IdPeriodoTiempo = iValue
End Property
Property Get IdPeriodoTiempo() As Integer
   IdPeriodoTiempo = mi_IdPeriodoTiempo
End Property
Property Let PeriodoTiempo(sValue As String)
   ms_PeriodoTiempo = sValue
End Property
Property Get PeriodoTiempo() As String
   PeriodoTiempo = ms_PeriodoTiempo
End Property
