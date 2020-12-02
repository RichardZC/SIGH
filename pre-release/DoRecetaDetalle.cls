VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoRecetaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla RecetaDetalle
'        Programado por: Barrantes D
'        Fecha: Marzo 2012
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdReceta As Long
Dim ml_idItem As Long
Dim ml_CantidadPedida As Long
Dim mc_precio As Currency
Dim mc_Total As Currency
Dim ml_SaldoEnRegistroReceta As Long
Dim ml_SaldoEnDespachoReceta As Long
Dim ml_CantidadDespachada As Long
Dim ml_idDosisRecetada As Long
Dim ml_idEstadoDetalle As Long
Dim ml_MotivoAnulacionMedico As String
Dim ml_Observaciones As String
Dim ml_IdViaAdministracion As Long 'Actualizado 26092014
Dim ml_Dx As String
Dim ml_idPuntoCarga As Long

Property Let idPuntoCarga(lValue As Long)
   ml_idPuntoCarga = lValue
End Property
Property Get idPuntoCarga() As Long
   idPuntoCarga = ml_idPuntoCarga
End Property

Property Let dx(lValue As String)
   ml_Dx = lValue
End Property
Property Get dx() As String
   dx = ml_Dx
End Property

Property Let Observaciones(lValue As String)
   ml_Observaciones = lValue
End Property
Property Get Observaciones() As String
   Observaciones = ml_Observaciones
End Property
Property Let MotivoAnulacionMedico(lValue As String)
   ml_MotivoAnulacionMedico = lValue
End Property
Property Get MotivoAnulacionMedico() As String
   MotivoAnulacionMedico = ml_MotivoAnulacionMedico
End Property

Property Let idDosisRecetada(lValue As Long)
   ml_idDosisRecetada = lValue
End Property
Property Get idDosisRecetada() As Long
   idDosisRecetada = ml_idDosisRecetada
End Property
Property Let idEstadoDetalle(lValue As Long)
   ml_idEstadoDetalle = lValue
End Property
Property Get idEstadoDetalle() As Long
   idEstadoDetalle = ml_idEstadoDetalle
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
Property Let CantidadPedida(iValue As Long)
   ml_CantidadPedida = iValue
End Property
Property Get CantidadPedida() As Long
   CantidadPedida = ml_CantidadPedida
End Property
Property Let precio(cValue As Currency)
   mc_precio = cValue
End Property
Property Get precio() As Currency
   precio = mc_precio
End Property
Property Let TOTAL(cValue As Currency)
   mc_Total = cValue
End Property
Property Get TOTAL() As Currency
   TOTAL = mc_Total
End Property
Property Let SaldoEnRegistroReceta(iValue As Long)
   ml_SaldoEnRegistroReceta = iValue
End Property
Property Get SaldoEnRegistroReceta() As Long
   SaldoEnRegistroReceta = ml_SaldoEnRegistroReceta
End Property
Property Let SaldoEnDespachoReceta(iValue As Long)
   ml_SaldoEnDespachoReceta = iValue
End Property
Property Get SaldoEnDespachoReceta() As Long
   SaldoEnDespachoReceta = ml_SaldoEnDespachoReceta
End Property
Property Let CantidadDespachada(iValue As Long)
   ml_CantidadDespachada = iValue
End Property
Property Get CantidadDespachada() As Long
   CantidadDespachada = ml_CantidadDespachada
End Property

'Actualizado 26092014
Property Let IdViaAdministracion(lValue As Long)
   ml_IdViaAdministracion = lValue
End Property
Property Get IdViaAdministracion() As Long
   IdViaAdministracion = ml_IdViaAdministracion
End Property

