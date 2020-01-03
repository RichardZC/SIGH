VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFacturacionBienesDevol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FacturacionBienesDevoluciones
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ms_movNumero As String
Dim ms_movTipo As String
Dim ml_IdProducto As Long
Dim ml_CantidadAdevolver As Long
Dim ml_idComprobantePago As Long
Dim ml_idEstadoDevolucion As Long
Dim mda_FechaAutoriza As Date
Dim ml_IdUsuarioAutoriza As Long
Dim ms_MovNumeroE As String
Dim ms_MovTipoE As String

Property Let MovNumeroE(sValue As String)
   ms_MovNumeroE = sValue
End Property
Property Get MovNumeroE() As String
   MovNumeroE = ms_MovNumeroE
End Property
Property Let MovTipoE(sValue As String)
   ms_MovTipoE = sValue
End Property
Property Get MovTipoE() As String
   MovTipoE = ms_MovTipoE
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let movNumero(sValue As String)
   ms_movNumero = sValue
End Property
Property Get movNumero() As String
   movNumero = ms_movNumero
End Property
Property Let movTipo(sValue As String)
   ms_movTipo = sValue
End Property
Property Get movTipo() As String
   movTipo = ms_movTipo
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

