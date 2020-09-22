VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoRecetaCabecera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla RecetaCabecera
'        Programado por: Barrantes D
'        Fecha: Marzo 2012
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdReceta As Long
Dim ml_IdPuntoCarga As Long
Dim mda_FechaReceta As Date
Dim ml_IdCuentaAtencion As Long
Dim ml_idServicioReceta As Long
Dim ml_idEstado As Long
Dim ml_idComprobantePago As Long
Dim ml_idMedicoReceta As Long
Dim mda_FechaVigencia As Date

Property Let FechaVigencia(daValue As Date)
   mda_FechaVigencia = daValue
End Property
Property Get FechaVigencia() As Date
   FechaVigencia = mda_FechaVigencia
End Property

Property Let idMedicoReceta(lValue As Long)
   ml_idMedicoReceta = lValue
End Property
Property Get idMedicoReceta() As Long
   idMedicoReceta = ml_idMedicoReceta
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdReceta(iValue As Long)
   ml_IdReceta = iValue
End Property
Property Get IdReceta() As Long
   IdReceta = ml_IdReceta
End Property
Property Let IdPuntoCarga(iValue As Long)
   ml_IdPuntoCarga = iValue
End Property
Property Get IdPuntoCarga() As Long
   IdPuntoCarga = ml_IdPuntoCarga
End Property
Property Let FechaReceta(daValue As Date)
   mda_FechaReceta = daValue
End Property
Property Get FechaReceta() As Date
   FechaReceta = mda_FechaReceta
End Property
Property Let IdCuentaAtencion(iValue As Long)
   ml_IdCuentaAtencion = iValue
End Property
Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
End Property
Property Let idServicioReceta(iValue As Long)
   ml_idServicioReceta = iValue
End Property
Property Get idServicioReceta() As Long
   idServicioReceta = ml_idServicioReceta
End Property
Property Let idEstado(iValue As Long)
   ml_idEstado = iValue
End Property
Property Get idEstado() As Long
   idEstado = ml_idEstado
End Property
Property Let idComprobantePago(iValue As Long)
   ml_idComprobantePago = iValue
End Property
Property Get idComprobantePago() As Long
   idComprobantePago = ml_idComprobantePago
End Property



