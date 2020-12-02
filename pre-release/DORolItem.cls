VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DORolItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla RolesItems
'        Programado por: Castro W
'        Fecha: Marzo 2006
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim mb_Consultar As Boolean
Dim mb_Eliminar As Boolean
Dim mb_Modificar As Boolean
Dim mb_Agregar As Boolean
Dim ml_IdRol As Long
Dim ml_IdListItem As Long
Dim ml_IdRolItem As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Consultar(bValue As Boolean)
   mb_Consultar = bValue
End Property
Property Get Consultar() As Boolean
   Consultar = mb_Consultar
End Property
Property Let Eliminar(bValue As Boolean)
   mb_Eliminar = bValue
End Property
Property Get Eliminar() As Boolean
   Eliminar = mb_Eliminar
End Property
Property Let Modificar(bValue As Boolean)
   mb_Modificar = bValue
End Property
Property Get Modificar() As Boolean
   Modificar = mb_Modificar
End Property
Property Let Agregar(bValue As Boolean)
   mb_Agregar = bValue
End Property
Property Get Agregar() As Boolean
   Agregar = mb_Agregar
End Property
Property Let IdRol(lValue As Long)
   ml_IdRol = lValue
End Property
Property Get IdRol() As Long
   IdRol = ml_IdRol
End Property
Property Let IdListItem(lValue As Long)
   ml_IdListItem = lValue
End Property
Property Get IdListItem() As Long
   IdListItem = ml_IdListItem
End Property
Property Let IdRolItem(lValue As Long)
   ml_IdRolItem = lValue
End Property
Property Get IdRolItem() As Long
   IdRolItem = ml_IdRolItem
End Property


