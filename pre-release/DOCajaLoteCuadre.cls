VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCajaLoteCuadre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'MZD Ini 02/06/2005 [Todo el archivo]

Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOCajaLoteCuadre
'        Autor: William Castro Grijalva
'        Fecha: 03/06/2005 10:24:49
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_IdLote As Long
Dim mc_Diferencia As Currency
Dim mc_CuadreUsuario As Currency
Dim mc_Calculado As Currency
Dim ml_IdLoteCuadre As Long
Dim ml_IdTipoFormaPago As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdLote(lValue As Long)
   ml_IdLote = lValue
End Property
Property Get IdLote() As Long
   IdLote = ml_IdLote
End Property
Property Let Diferencia(cValue As Currency)
   mc_Diferencia = cValue
End Property
Property Get Diferencia() As Currency
   Diferencia = mc_Diferencia
End Property
Property Let CuadreUsuario(cValue As Currency)
   mc_CuadreUsuario = cValue
End Property
Property Get CuadreUsuario() As Currency
   CuadreUsuario = mc_CuadreUsuario
End Property
Property Let Calculado(cValue As Currency)
   mc_Calculado = cValue
End Property
Property Get Calculado() As Currency
   Calculado = mc_Calculado
End Property
Property Let IdLoteCuadre(lValue As Long)
   ml_IdLoteCuadre = lValue
End Property
Property Get IdLoteCuadre() As Long
   IdLoteCuadre = ml_IdLoteCuadre
End Property
Property Let IdTipoFormaPago(lValue As Long)
   ml_IdTipoFormaPago = lValue
End Property
Property Get IdTipoFormaPago() As Long
   IdTipoFormaPago = ml_IdTipoFormaPago
End Property

