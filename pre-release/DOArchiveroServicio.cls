VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOArchiveroServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla ArchiveroServicio
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------

Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdServicio As Long
Dim ml_IdEmpleado As Long
Dim ml_IdArchivero As Long
Dim ml_EsConsultorioAsignado As Boolean
Property Let EsConsultorioAsignado(lValue As Boolean)
   ml_EsConsultorioAsignado = lValue
End Property
Property Get EsConsultorioAsignado() As Boolean
   EsConsultorioAsignado = ml_EsConsultorioAsignado
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdServicio(lValue As Long)
   ml_IdServicio = lValue
End Property
Property Get IdServicio() As Long
   IdServicio = ml_IdServicio
End Property
Property Let IdEmpleado(lValue As Long)
   ml_IdEmpleado = lValue
End Property
Property Get IdEmpleado() As Long
   IdEmpleado = ml_IdEmpleado
End Property
Property Let IdArchivero(lValue As Long)
   ml_IdArchivero = lValue
End Property
Property Get IdArchivero() As Long
   IdArchivero = ml_IdArchivero
End Property


