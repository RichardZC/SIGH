VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCajaDevoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: CajaDevoluciones
'        Autor: Daniel Enrique Barrantes
'        Fecha: 13/01/2015 12:38:14
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idDevolucion As Long
Dim ml_idComprobantePago As Long
Dim mc_montoDevuelto As Currency
Dim mc_montoTotal As Currency
Dim mda_fechaDevolucion As Date
Dim ml_motivo As String
Dim ml_idUsuario As Long


Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idDevolucion(ivalue As Long)
   ml_idDevolucion = ivalue
End Property
Property Get idDevolucion() As Long
   idDevolucion = ml_idDevolucion
End Property
Property Let idComprobantePago(ivalue As Long)
   ml_idComprobantePago = ivalue
End Property
Property Get idComprobantePago() As Long
   idComprobantePago = ml_idComprobantePago
End Property
Property Let montoDevuelto(cValue As Currency)
   mc_montoDevuelto = cValue
End Property
Property Get montoDevuelto() As Currency
   montoDevuelto = mc_montoDevuelto
End Property
Property Let montoTotal(cValue As Currency)
   mc_montoTotal = cValue
End Property
Property Get montoTotal() As Currency
   montoTotal = mc_montoTotal
End Property
Property Let fechaDevolucion(daValue As Date)
   mda_fechaDevolucion = daValue
End Property
Property Get fechaDevolucion() As Date
   fechaDevolucion = mda_fechaDevolucion
End Property
Property Let idUsuario(ivalue As Long)
   ml_idUsuario = ivalue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property

Property Let mMotivo(moValue As String)
   ml_motivo = moValue
End Property

Property Get mMotivo() As String
   mMotivo = ml_motivo
End Property
