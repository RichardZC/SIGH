VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOFarmaciaRecetasDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOFarmaciaRecetasDetalle
'        Autor: William Castro Grijalva
'        Fecha: 07/08/2005 08:37:49
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_Cantidad As Long
Dim ml_IdFacturacionBienes As Long
Dim ml_IdRecetaDetalle As Long
Dim ml_IdProducto As Long
Dim ml_IdReceta As Long
Dim ms_EstadoRegistro As String

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Cantidad(lValue As Long)
   ml_Cantidad = lValue
End Property
Property Get Cantidad() As Long
   Cantidad = ml_Cantidad
End Property
Property Let IdFacturacionBienes(lValue As Long)
   ml_IdFacturacionBienes = lValue
End Property
Property Get IdFacturacionBienes() As Long
   IdFacturacionBienes = ml_IdFacturacionBienes
End Property
Property Let IdRecetaDetalle(lValue As Long)
   ml_IdRecetaDetalle = lValue
End Property
Property Get IdRecetaDetalle() As Long
   IdRecetaDetalle = ml_IdRecetaDetalle
End Property
Property Let IdProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let IdReceta(lValue As Long)
   ml_IdReceta = lValue
End Property
Property Get IdReceta() As Long
   IdReceta = ml_IdReceta
End Property

Property Let EstadoRegistro(sValue As String)
   ms_EstadoRegistro = sValue
End Property
Property Get EstadoRegistro() As String
   EstadoRegistro = ms_EstadoRegistro
End Property

