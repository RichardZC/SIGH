VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOLabMovimientoCPT"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla LabMovimientoCPT
'        Programado por: Bonilla A
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdMovimiento As Long
Dim ml_idProductoCpt As Long
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
Property Let IdMovimiento(iValue As Long)
   ml_IdMovimiento = iValue
End Property

Property Get IdMovimiento() As Long
   IdMovimiento = ml_IdMovimiento
End Property

Property Let idProductoCpt(iValue As Long)
   ml_idProductoCpt = iValue
End Property

Property Get idProductoCpt() As Long
   idProductoCpt = ml_idProductoCpt
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


