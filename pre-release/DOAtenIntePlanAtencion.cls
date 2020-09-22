VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtenIntePlanAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtenIntePlanAtencion
'        Programado por: Garay M
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdPlanAtencion As Long
Dim mi_IdAtenInteGrupo As Integer
Dim mi_IdPeriodoTiempo As Integer
Dim ml_EdadAnio As Long
Dim mi_EdadMes As Integer
Dim mi_EdadDia As Integer
Dim ms_Descripcion As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdPlanAtencion(iValue As Long)
   ml_IdPlanAtencion = iValue
End Property
Property Get IdPlanAtencion() As Long
   IdPlanAtencion = ml_IdPlanAtencion
End Property
Property Let IdAtenInteGrupo(iValue As Integer)
   mi_IdAtenInteGrupo = iValue
End Property
Property Get IdAtenInteGrupo() As Integer
   IdAtenInteGrupo = mi_IdAtenInteGrupo
End Property
Property Let IdPeriodoTiempo(iValue As Integer)
   mi_IdPeriodoTiempo = iValue
End Property
Property Get IdPeriodoTiempo() As Integer
   IdPeriodoTiempo = mi_IdPeriodoTiempo
End Property
Property Let EdadAnio(iValue As Long)
   ml_EdadAnio = iValue
End Property
Property Get EdadAnio() As Long
   EdadAnio = ml_EdadAnio
End Property
Property Let EdadMes(iValue As Integer)
   mi_EdadMes = iValue
End Property
Property Get EdadMes() As Integer
   EdadMes = mi_EdadMes
End Property
Property Let EdadDia(iValue As Integer)
   mi_EdadDia = iValue
End Property
Property Get EdadDia() As Integer
   EdadDia = mi_EdadDia
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
