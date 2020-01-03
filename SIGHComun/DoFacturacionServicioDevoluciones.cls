VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFacturacionServicioDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FacturacionServicioDevoluciones
'        Programado por: Barrantes D
'        Fecha: Abril 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idOrden As Long
Dim ml_IdProducto As Long
Dim ml_CantidadAdevolver As Long
Dim ml_idComprobantePago As Long
Dim ml_idEstadoDevolucion As Long
Dim mda_FechaAutoriza As Date
Dim ml_IdUsuarioAutoriza As Long

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let idOrden(iValue As Long)
   ml_idOrden = iValue
End Property
Property Get idOrden() As Long
   idOrden = ml_idOrden
End Property
Property Let IdProducto(iValue As Long)
   ml_IdProducto = iValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let CantidadAdevolver(iValue As Long)
   ml_CantidadAdevolver = iValue
End Property
Property Get CantidadAdevolver() As Long
   CantidadAdevolver = ml_CantidadAdevolver
End Property
Property Let idComprobantePago(iValue As Long)
   ml_idComprobantePago = iValue
End Property
Property Get idComprobantePago() As Long
   idComprobantePago = ml_idComprobantePago
End Property
Property Let idEstadoDevolucion(iValue As Long)
   ml_idEstadoDevolucion = iValue
End Property
Property Get idEstadoDevolucion() As Long
   idEstadoDevolucion = ml_idEstadoDevolucion
End Property
Property Let FechaAutoriza(daValue As Date)
   mda_FechaAutoriza = daValue
End Property
Property Get FechaAutoriza() As Date
   FechaAutoriza = mda_FechaAutoriza
End Property
Property Let IdUsuarioAutoriza(iValue As Long)
   ml_IdUsuarioAutoriza = iValue
End Property
Property Get IdUsuarioAutoriza() As Long
   IdUsuarioAutoriza = ml_IdUsuarioAutoriza
End Property

