VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCamasMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla CamasMovimientos
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdMovimiento As Long
Dim ml_IdCama As Long
Dim ml_IdServicio As Long
Dim ml_IdFechaIngreso As Date
Dim ml_IdFechaSalida As Date
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdMovimiento(lValue As Long)
   ml_IdMovimiento = lValue
End Property
Property Get IdMovimiento() As Long
   IdMovimiento = ml_IdMovimiento
End Property
Property Let IdCama(lValue As Long)
   ml_IdCama = lValue
End Property
Property Get IdCama() As Long
   IdCama = ml_IdCama
End Property
Property Let IdServicio(lValue As Long)
   ml_IdServicio = lValue
End Property
Property Get IdServicio() As Long
   IdServicio = ml_IdServicio
End Property
Property Let IdFechaIngreso(lValue As Date)
   ml_IdFechaIngreso = lValue
End Property
Property Get IdFechaIngreso() As Date
   IdFechaIngreso = ml_IdFechaIngreso
End Property
Property Let IdFechaSalida(lValue As Date)
   ml_IdFechaSalida = lValue
End Property
Property Get IdFechaSalida() As Date
   IdFechaSalida = ml_IdFechaSalida
End Property


