VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFacturacionServicioFinanciamientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: FacturacionServicioFinanciamientos
'        Autor: Daniel Enrique Barrantes
'        Fecha: 05/03/2009 10:56:52 p.m.
'        Auspiciado por: Hospital Regional Ayacucho
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idOrden As Long
Dim ml_IdProducto As Long
Dim ml_IdTipoFinanciamiento As Long
Dim ml_CantidadFinanciada As Long
Dim mc_PrecioFinanciado As Currency
Dim mc_TotalFinanciado As Currency
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
Property Let IdTipoFinanciamiento(iValue As Long)
   ml_IdTipoFinanciamiento = iValue
End Property
Property Get IdTipoFinanciamiento() As Long
   IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property
Property Let CantidadFinanciada(iValue As Long)
   ml_CantidadFinanciada = iValue
End Property
Property Get CantidadFinanciada() As Long
   CantidadFinanciada = ml_CantidadFinanciada
End Property
Property Let PrecioFinanciado(cValue As Currency)
   mc_PrecioFinanciado = cValue
End Property
Property Get PrecioFinanciado() As Currency
   PrecioFinanciado = mc_PrecioFinanciado
End Property
Property Let TotalFinanciado(cValue As Currency)
   mc_TotalFinanciado = cValue
End Property
Property Get TotalFinanciado() As Currency
   TotalFinanciado = mc_TotalFinanciado
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


