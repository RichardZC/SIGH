VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOFacturacionPaquetes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FacturacionPaquetes
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idComprobantePago As Long
Dim ml_idOrdenPago As Long
Dim ml_IdProducto As Long
Dim ml_idFactPaquete As Long
Dim ml_idPuntoCarga As Long
Dim ml_idEspecialidadServicio As Long
Dim ml_AtencionId As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idComprobantePago(iValue As Long)
   ml_idComprobantePago = iValue
End Property
Property Get idComprobantePago() As Long
   idComprobantePago = ml_idComprobantePago
End Property
Property Let idOrdenPago(iValue As Long)
   ml_idOrdenPago = iValue
End Property
Property Get idOrdenPago() As Long
   idOrdenPago = ml_idOrdenPago
End Property
Property Let IdProducto(iValue As Long)
   ml_IdProducto = iValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let idFactPaquete(iValue As Long)
   ml_idFactPaquete = iValue
End Property
Property Get idFactPaquete() As Long
   idFactPaquete = ml_idFactPaquete
End Property
Property Let idPuntoCarga(iValue As Long)
   ml_idPuntoCarga = iValue
End Property
Property Get idPuntoCarga() As Long
   idPuntoCarga = ml_idPuntoCarga
End Property
Property Let idEspecialidadServicio(iValue As Long)
   ml_idEspecialidadServicio = iValue
End Property
Property Get idEspecialidadServicio() As Long
   idEspecialidadServicio = ml_idEspecialidadServicio
End Property
Property Let AtencionId(iValue As Long)
   ml_AtencionId = iValue
End Property
Property Get AtencionId() As Long
   AtencionId = ml_AtencionId
End Property


