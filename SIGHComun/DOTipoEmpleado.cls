VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOTipoEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla TiposEmpleado
'        Programado por: Castro W
'        Fecha: Julio 2004
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdTipoEmpleado As Long
Dim ms_Descripcion As String
Dim mb_EsColegiatura As Boolean

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let IdTipoEmpleado(lValue As Long)
   ml_IdTipoEmpleado = lValue
End Property
Property Get IdTipoEmpleado() As Long
   IdTipoEmpleado = ml_IdTipoEmpleado
End Property

Property Let IdEsColegiatura(sValue As Boolean)
   mb_EsColegiatura = sValue
End Property

Property Get IdEsColegiatura() As Boolean
   IdEsColegiatura = mb_EsColegiatura
End Property

