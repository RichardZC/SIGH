VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOhis_historicoAten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla his_historicoAtenciones
'        Programado por: Barrantes D
'        Fecha: Octubre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdPaciente As Long
Dim mda_fecha As Date
Dim ms_diagnost As String
Dim ms_cpt As String
Dim ms_ups As String

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdPaciente(iValue As Long)
   ml_IdPaciente = iValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let fecha(daValue As Date)
   mda_fecha = daValue
End Property
Property Get fecha() As Date
   fecha = mda_fecha
End Property
Property Let diagnost(sValue As String)
   ms_diagnost = sValue
End Property
Property Get diagnost() As String
   diagnost = ms_diagnost
End Property
Property Let cpt(sValue As String)
   ms_cpt = sValue
End Property
Property Get cpt() As String
   cpt = ms_cpt
End Property
Property Let ups(sValue As String)
   ms_ups = sValue
End Property
Property Get ups() As String
   ups = ms_ups
End Property


