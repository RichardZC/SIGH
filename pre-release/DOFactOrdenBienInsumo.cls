VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOFactOrdenBienInsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactOrdenesBienesInsumo
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim mda_FechaModificacion As Date
Dim mda_FechaCreacion As Date
Dim ml_idUsuarioModifica As Long
Dim ml_IdUsuarioCrea As Long
Dim ml_IdAtencion As Long
Dim mda_FechaOrden As Date
Dim ml_idOrden As Long
Dim ml_IdEstadoOrden As Long
Dim ml_idComprobantePago As Long
Dim ml_idPuntoCarga As Long
Dim ml_IdPaciente As Long
Dim ml_IdFormaPago As Long
Dim ml_idFarmacia As Long
Property Let idFarmacia(lValue As Long)
   ml_idFarmacia = lValue
End Property
Property Get idFarmacia() As Long
   idFarmacia = ml_idFarmacia
End Property
Property Let IdFormaPago(lValue As Long)
   ml_IdFormaPago = lValue
End Property
Property Get IdFormaPago() As Long
   IdFormaPago = ml_IdFormaPago
End Property
Property Let IdPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let FechaModificacion(daValue As Date)
   mda_FechaModificacion = daValue
End Property
Property Get FechaModificacion() As Date
   FechaModificacion = mda_FechaModificacion
End Property
Property Let FechaCreacion(daValue As Date)
   mda_FechaCreacion = daValue
End Property
Property Get FechaCreacion() As Date
   FechaCreacion = mda_FechaCreacion
End Property
Property Let idUsuarioModifica(lValue As Long)
   ml_idUsuarioModifica = lValue
End Property
Property Get idUsuarioModifica() As Long
   idUsuarioModifica = ml_idUsuarioModifica
End Property
Property Let IdUsuarioCrea(lValue As Long)
   ml_IdUsuarioCrea = lValue
End Property
Property Get IdUsuarioCrea() As Long
   IdUsuarioCrea = ml_IdUsuarioCrea
End Property
Property Let IdAtencion(lValue As Long)
   ml_IdAtencion = lValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property
Property Let FechaOrden(daValue As Date)
   mda_FechaOrden = daValue
End Property
Property Get FechaOrden() As Date
   FechaOrden = mda_FechaOrden
End Property
Property Let idOrden(lValue As Long)
   ml_idOrden = lValue
End Property
Property Get idOrden() As Long
   idOrden = ml_idOrden
End Property
Property Let IdEstadoOrden(lValue As Long)
   ml_IdEstadoOrden = lValue
End Property
Property Get IdEstadoOrden() As Long
   IdEstadoOrden = ml_IdEstadoOrden
End Property
Property Let idComprobantePago(lValue As Long)
   ml_idComprobantePago = lValue
End Property
Property Get idComprobantePago() As Long
   idComprobantePago = ml_idComprobantePago
End Property
Property Let idPuntoCarga(lValue As Long)
   ml_idPuntoCarga = lValue
End Property
Property Get idPuntoCarga() As Long
   idPuntoCarga = ml_idPuntoCarga
End Property


