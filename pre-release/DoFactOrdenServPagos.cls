VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFactOrdenServPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactOrdenServicioPagos
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idOrdenPago As Long
Dim ml_idComprobantePago As Long
Dim ml_idOrden As Long
Dim mda_FechaCreacion As Date
Dim ml_idUsuario As Long
Dim ml_IdEstadoFacturacion As Long
Dim mc_ImporteExonerado As Currency
Dim ml_idUsuarioExonera As Long

Property Let idUsuarioExonera(lValue As Long)
   ml_idUsuarioExonera = lValue
End Property
Property Get idUsuarioExonera() As Long
   idUsuarioExonera = ml_idUsuarioExonera
End Property

Property Let ImporteExonerado(cValue As Currency)
   mc_ImporteExonerado = cValue
End Property
Property Get ImporteExonerado() As Currency
   ImporteExonerado = mc_ImporteExonerado
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let idOrdenPago(iValue As Long)
   ml_idOrdenPago = iValue
End Property
Property Get idOrdenPago() As Long
   idOrdenPago = ml_idOrdenPago
End Property
Property Let idComprobantePago(iValue As Long)
   ml_idComprobantePago = iValue
End Property
Property Get idComprobantePago() As Long
   idComprobantePago = ml_idComprobantePago
End Property
Property Let idOrden(iValue As Long)
   ml_idOrden = iValue
End Property
Property Get idOrden() As Long
   idOrden = ml_idOrden
End Property

Property Let FechaCreacion(daValue As Date)
   mda_FechaCreacion = daValue
End Property
Property Get FechaCreacion() As Date
   FechaCreacion = mda_FechaCreacion
End Property
Property Let idUsuario(iValue As Long)
   ml_idUsuario = iValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let IdEstadoFacturacion(iValue As Long)
   ml_IdEstadoFacturacion = iValue
End Property
Property Get IdEstadoFacturacion() As Long
   IdEstadoFacturacion = ml_IdEstadoFacturacion
End Property



