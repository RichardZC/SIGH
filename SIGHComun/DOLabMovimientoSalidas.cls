VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOLabMovimientoSalidas"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla LabMovimientoSalidas
'        Programado por: Bonilla A
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdMovimiento As Long
Dim ml_IdResponsable As Long
Dim ml_IdMotivoSalida As Long

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property

Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let IdMovimiento(iValue As Long)
   ml_IdMovimiento = iValue
End Property

Property Get IdMovimiento() As Long
   IdMovimiento = ml_IdMovimiento
End Property

Property Let IdResponsable(iValue As Long)
   ml_IdResponsable = iValue
End Property

Property Get IdResponsable() As Long
   IdResponsable = ml_IdResponsable
End Property

Property Let IdMotivoSalida(iValue As Long)
   ml_IdMotivoSalida = iValue
End Property

Property Get IdMotivoSalida() As Long
   IdMotivoSalida = ml_IdMotivoSalida
End Property

