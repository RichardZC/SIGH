VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoRecetaDetalleItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla RecetaDetalleItem
'        Programado por: Barrantes D
'        Fecha: Marzo 2012
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdReceta As Long
Dim ml_idItem As Long
Dim ms_DocumentoDespacho As String
Dim ml_CantidadDespachada As Long
Dim ml_IdComprobantePago As Long
Dim ml_idEstado As Long
Property Let idEstado(iValue As Long)
   ml_idEstado = iValue
End Property
Property Get idEstado() As Long
   idEstado = ml_idEstado
End Property
Property Let IdComprobantePago(lValue As Long)
   ml_IdComprobantePago = lValue
End Property
Property Get IdComprobantePago() As Long
   IdComprobantePago = ml_IdComprobantePago
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
Property Let idItem(iValue As Long)
   ml_idItem = iValue
End Property
Property Get idItem() As Long
   idItem = ml_idItem
End Property
Property Let DocumentoDespacho(sValue As String)
   ms_DocumentoDespacho = sValue
End Property
Property Get DocumentoDespacho() As String
   DocumentoDespacho = ms_DocumentoDespacho
End Property
Property Let CantidadDespachada(iValue As Long)
   ml_CantidadDespachada = iValue
End Property
Property Get CantidadDespachada() As Long
   CantidadDespachada = ml_CantidadDespachada
End Property



