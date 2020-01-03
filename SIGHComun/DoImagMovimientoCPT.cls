VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoImagMovimientoCPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla ImagMovimientoCPT
'        Programado por: Barrantes D
'        Fecha: Abril 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idMovimiento As Long
Dim ml_idProductoCPT As Long
Dim ml_Cantidad As Long
Dim mc_Precio As Currency
Dim mc_Importe As Currency
Dim ml_ResultadoAutomatico As Long
Property Let ResultadoAutomatico(lValue As Long)
   ml_ResultadoAutomatico = lValue
End Property
Property Get ResultadoAutomatico() As Long
   ResultadoAutomatico = ml_ResultadoAutomatico
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let idMovimiento(iValue As Long)
   ml_idMovimiento = iValue
End Property
Property Get idMovimiento() As Long
   idMovimiento = ml_idMovimiento
End Property
Property Let idProductoCPT(iValue As Long)
   ml_idProductoCPT = iValue
End Property
Property Get idProductoCPT() As Long
   idProductoCPT = ml_idProductoCPT
End Property
Property Let Cantidad(iValue As Long)
   ml_Cantidad = iValue
End Property
Property Get Cantidad() As Long
   Cantidad = ml_Cantidad
End Property
Property Let PRECIO(cValue As Currency)
   mc_Precio = cValue
End Property
Property Get PRECIO() As Currency
   PRECIO = mc_Precio
End Property
Property Let Importe(cValue As Currency)
   mc_Importe = cValue
End Property
Property Get Importe() As Currency
   Importe = mc_Importe
End Property



