VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtenIntePlanDesPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtenIntePlanDesarrolloPaciente
'        Programado por: Garay M
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdPlanDesarrolloPaciente As Long
Dim ml_IdPlanIntegralPaciente As Long
Dim ml_Evaluacion As Long
Dim ml_IdPlanAtencion As Long
Dim ml_IdAtenInteItemPlan As Long
Dim mda_FechaProgramada As Date
Dim mda_FechaEjecucion As Date
Dim ml_NumeroSesion As Long
Dim ml_IdAtencion As Long
Dim ml_IdEstablecimiento As Long
'propiedades no persistentes
'mgaray20141012
Dim ms_Descripcion As String
Dim ml_EdadAnio As Integer
Dim ml_EdadMes As Integer
Dim ml_EdadDia As Integer
Dim ms_EvaluacionDesc As String
Dim ms_Establecimiento As String

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdPlanDesarrolloPaciente(iValue As Long)
   ml_IdPlanDesarrolloPaciente = iValue
End Property
Property Get IdPlanDesarrolloPaciente() As Long
   IdPlanDesarrolloPaciente = ml_IdPlanDesarrolloPaciente
End Property
Property Let IdPlanIntegralPaciente(lValue As Long)
   ml_IdPlanIntegralPaciente = lValue
End Property
Property Get IdPlanIntegralPaciente() As Long
   IdPlanIntegralPaciente = ml_IdPlanIntegralPaciente
End Property
Property Let Evaluacion(iValue As Long)
   ml_Evaluacion = iValue
End Property
Property Get Evaluacion() As Long
   Evaluacion = ml_Evaluacion
End Property
Property Let IdPlanAtencion(iValue As Long)
   ml_IdPlanAtencion = iValue
End Property
Property Get IdPlanAtencion() As Long
   IdPlanAtencion = ml_IdPlanAtencion
End Property
Property Let IdAtenInteItemPlan(iValue As Long)
   ml_IdAtenInteItemPlan = iValue
End Property
Property Get IdAtenInteItemPlan() As Long
   IdAtenInteItemPlan = ml_IdAtenInteItemPlan
End Property
Property Let FechaProgramada(daValue As Date)
   mda_FechaProgramada = daValue
End Property
Property Get FechaProgramada() As Date
   FechaProgramada = mda_FechaProgramada
End Property
Property Let FechaEjecucion(daValue As Date)
   mda_FechaEjecucion = daValue
End Property
Property Get FechaEjecucion() As Date
   FechaEjecucion = mda_FechaEjecucion
End Property
Property Let NumeroSesion(iValue As Long)
   ml_NumeroSesion = iValue
End Property
Property Get NumeroSesion() As Long
   NumeroSesion = ml_NumeroSesion
End Property
Property Let IdAtencion(iValue As Long)
   ml_IdAtencion = iValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property
Property Let IdEstablecimiento(iValue As Long)
   ml_IdEstablecimiento = iValue
End Property
Property Get IdEstablecimiento() As Long
   IdEstablecimiento = ml_IdEstablecimiento
End Property
'mgaray20141012
Property Let EdadAnio(iValue As Long)
   ml_EdadAnio = iValue
End Property
Property Get EdadAnio() As Long
   EdadAnio = ml_EdadAnio
End Property

Property Let EdadMes(iValue As Long)
   ml_EdadMes = iValue
End Property
Property Get EdadMes() As Long
   EdadMes = ml_EdadMes
End Property

Property Let EdadDia(iValue As Long)
   ml_EdadDia = iValue
End Property
Property Get EdadDia() As Long
   EdadDia = ml_EdadDia
End Property

Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property

Property Let EvaluacionDesc(sValue As String)
   ms_EvaluacionDesc = sValue
End Property
Property Get EvaluacionDesc() As String
   EvaluacionDesc = ms_EvaluacionDesc
End Property

Property Let Establecimiento(sValue As String)
   ms_Establecimiento = sValue
End Property
Property Get Establecimiento() As String
   Establecimiento = ms_Establecimiento
End Property

