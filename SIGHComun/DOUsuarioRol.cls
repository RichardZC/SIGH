VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOUsuarioRol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla usuariosRoles
'        Programado por: Castro W
'        Fecha: Julio 2004
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdEmpleado As Long
Dim ml_IdRol As Long
Dim ml_IdUsuarioRol As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdEmpleado(lValue As Long)
   ml_IdEmpleado = lValue
End Property
Property Get IdEmpleado() As Long
   IdEmpleado = ml_IdEmpleado
End Property
Property Let IdRol(lValue As Long)
   ml_IdRol = lValue
End Property
Property Get IdRol() As Long
   IdRol = ml_IdRol
End Property
Property Let IdUsuarioRol(lValue As Long)
   ml_IdUsuarioRol = lValue
End Property
Property Get IdUsuarioRol() As Long
   IdUsuarioRol = ml_IdUsuarioRol
End Property


